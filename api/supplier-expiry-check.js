// api/supplier-expiry-check.js
// Runs daily via Vercel Cron — sends alert email ONLY for documents expiring
// within the thresholds set in mob_supplier_alert_settings (default: 90 / 15 days).
// Required env vars in Vercel Dashboard:
//   RESEND_API_KEY        — from resend.com
//   SUPABASE_URL          — https://oaerqjrkdpuhiproppaz.supabase.co
//   SUPABASE_SERVICE_KEY  — service_role key (not anon)
//   CRON_SECRET           — any random string you set

export default async function handler(req, res) {
  const authHeader = req.headers['authorization'];
  if (authHeader !== `Bearer ${process.env.CRON_SECRET}`) {
    return res.status(401).json({ error: 'Unauthorized' });
  }

  const SUPA_URL   = process.env.SUPABASE_URL || 'https://oaerqjrkdpuhiproppaz.supabase.co';
  const SUPA_KEY   = process.env.SUPABASE_SERVICE_KEY;
  const RESEND_KEY = process.env.RESEND_API_KEY;

  if (!SUPA_KEY || !RESEND_KEY) {
    return res.status(500).json({ error: 'Missing env vars: SUPABASE_SERVICE_KEY or RESEND_API_KEY' });
  }

  const settingsRes = await fetch(`${SUPA_URL}/rest/v1/mob_supplier_alert_settings?limit=1`, {
    headers: { apikey: SUPA_KEY, Authorization: `Bearer ${SUPA_KEY}` }
  });
  const settingsArr = await settingsRes.json();
  const settings    = settingsArr?.[0] || {};

  const alertEmail   = settings.alert_email || 'admin@satcoarabiaengg.com';
  const warnDays     = settings.warn_days  ?? 90;
  const criticalDays = settings.alert_days ?? 15;
  const enabled      = settings.enabled    ?? true;

  if (!enabled) {
    return res.status(200).json({ message: 'Alerts disabled in settings — skipping.' });
  }

  const suppRes   = await fetch(`${SUPA_URL}/rest/v1/mob_suppliers?active=eq.true&select=*`, {
    headers: { apikey: SUPA_KEY, Authorization: `Bearer ${SUPA_KEY}` }
  });
  const suppliers = await suppRes.json();

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  function daysLeft(dateStr) {
    if (!dateStr) return null;
    const d = new Date(dateStr);
    d.setHours(0, 0, 0, 0);
    return Math.ceil((d - today) / 86400000);
  }

  function expiryStatus(days) {
    if (days === null)        return null;
    if (days > warnDays)      return null;
    if (days < 0)             return { level: 'EXPIRED',  color: '#dc2626', emoji: '🔴' };
    if (days <= criticalDays) return { level: 'CRITICAL', color: '#dc2626', emoji: '🔴' };
    return                           { level: 'WARNING',  color: '#d97706', emoji: '🟡' };
  }

  const issues = [];

  for (const s of suppliers) {
    const docsToCheck = [
      { label: 'Trade / Commercial License',    dateStr: s.license_expiry, ref: s.license_number || '—' },
      { label: 'Workmen Comp. Insurance (WCI)', dateStr: s.wci_expiry,     ref: s.wci_policy_no  || '—' },
    ];
    for (const od of (Array.isArray(s.other_docs) ? s.other_docs : [])) {
      if (od.expiry) docsToCheck.push({ label: od.doc_name || 'Other Document', dateStr: od.expiry, ref: '' });
    }
    for (const { label, dateStr, ref } of docsToCheck) {
      const days   = daysLeft(dateStr);
      const status = expiryStatus(days);
      if (!status) continue;
      issues.push({ supplier: s.name, supplier_code: s.supplier_code || '—', doc: label, ref, days, status, expiry: dateStr });
    }
  }

  if (issues.length === 0) {
    return res.status(200).json({ message: `No supplier documents expiring within ${warnDays} days. No email sent.`, checked: suppliers.length });
  }

  const expired  = issues.filter(i => i.days < 0);
  const critical = issues.filter(i => i.days >= 0 && i.days <= criticalDays);
  const warning  = issues.filter(i => i.days > criticalDays);
  const todayStr = today.toLocaleDateString('en-GB', { day: '2-digit', month: 'long', year: 'numeric' });
  const uniqueSuppliers = [...new Set(issues.map(i => i.supplier))].length;

  function tableRows(list) {
    return list.map(i => `<tr>
      <td style="padding:8px 12px;border-bottom:1px solid #e2e8f0;font-weight:700;">${i.status.emoji} ${i.supplier}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #e2e8f0;font-family:monospace;font-size:12px;">${i.supplier_code}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #e2e8f0;">${i.doc}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #e2e8f0;font-family:monospace;font-size:12px;">${i.ref}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #e2e8f0;font-family:monospace;">${i.expiry || '—'}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #e2e8f0;font-weight:800;color:${i.status.color};">
        ${i.days < 0 ? `EXPIRED ${Math.abs(i.days)}d ago` : `${i.days} day${i.days !== 1 ? 's' : ''} left`}
      </td></tr>`).join('');
  }

  function section(title, list, headBg, headColor) {
    if (!list.length) return '';
    return `<h3 style="font-size:13px;font-weight:800;color:${headColor};text-transform:uppercase;letter-spacing:0.5px;margin:0 0 10px;">${title}</h3>
    <table style="width:100%;border-collapse:collapse;margin-bottom:24px;font-size:13px;">
      <thead><tr style="background:${headBg};">
        <th style="padding:8px 12px;text-align:left;font-size:11px;color:${headColor};">Supplier</th>
        <th style="padding:8px 12px;text-align:left;font-size:11px;color:${headColor};">Code</th>
        <th style="padding:8px 12px;text-align:left;font-size:11px;color:${headColor};">Document</th>
        <th style="padding:8px 12px;text-align:left;font-size:11px;color:${headColor};">Reference #</th>
        <th style="padding:8px 12px;text-align:left;font-size:11px;color:${headColor};">Expiry Date</th>
        <th style="padding:8px 12px;text-align:left;font-size:11px;color:${headColor};">Status</th>
      </tr></thead>
      <tbody>${tableRows(list)}</tbody>
    </table>`;
  }

  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8"/></head>
<body style="font-family:'Segoe UI',sans-serif;background:#f1f5f9;padding:24px;color:#0f172a;margin:0;">
<div style="max-width:760px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,0.08);">
  <div style="background:#0f172a;padding:20px 28px;">
    <div style="font-size:11px;font-weight:800;color:#b8972a;letter-spacing:2px;">SATCO ARABIA</div>
    <div style="font-size:18px;font-weight:700;color:#fff;margin-top:3px;">Supplier Document Expiry Alert</div>
    <div style="font-size:12px;color:#94a3b8;margin-top:3px;">${todayStr}</div>
  </div>
  <div style="padding:14px 28px;background:${expired.length ? '#fef2f2' : '#fffbeb'};border-bottom:1px solid ${expired.length ? '#fecaca' : '#fde68a'};">
    <div style="font-weight:800;font-size:14px;color:${expired.length ? '#991b1b' : '#92400e'};">
      ${[expired.length ? `🔴 ${expired.length} EXPIRED` : '', critical.length ? `🔴 ${critical.length} Critical (≤${criticalDays} days)` : '', warning.length ? `🟡 ${warning.length} Warning (≤${warnDays} days)` : ''].filter(Boolean).join(' · ')}
    </div>
    <div style="font-size:12px;color:#64748b;margin-top:4px;">
      ${issues.length} document${issues.length !== 1 ? 's' : ''} across ${uniqueSuppliers} supplier${uniqueSuppliers !== 1 ? 's' : ''} require attention.
      Alert window: within <strong>${warnDays} days</strong> (warning) · within <strong>${criticalDays} days</strong> (critical).
    </div>
  </div>
  <div style="padding:20px 28px;">
    ${section('🔴 Expired — Immediate Action Required', expired, '#fee2e2', '#991b1b')}
    ${section(`🔴 Critical — Expiring within ${criticalDays} days`, critical, '#fee2e2', '#991b1b')}
    ${section(`🟡 Warning — Expiring within ${warnDays} days`, warning, '#fef3c7', '#92400e')}
    <div style="margin-top:8px;padding:12px 16px;background:#f8fafc;border-radius:8px;font-size:12px;color:#64748b;">
      Update supplier documents at <a href="https://satco-mobilisation.vercel.app" style="color:#0369a1;font-weight:700;">satco-mobilisation.vercel.app</a> → 🏭 Suppliers tab.
    </div>
  </div>
  <div style="padding:12px 28px;background:#f8fafc;border-top:1px solid #e2e8f0;font-size:11px;color:#94a3b8;">
    SATCO Arabia General Contracting LLC — SPC · Automated daily alert from Mobilisation Portal
  </div>
</div></body></html>`;

  const subject = expired.length
    ? `🚨 SATCO Supplier Alert — ${expired.length} EXPIRED document${expired.length > 1 ? 's' : ''}`
    : `⚠️ SATCO Supplier Alert — ${issues.length} document${issues.length > 1 ? 's' : ''} expiring soon`;

  const emailRes = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: { Authorization: `Bearer ${RESEND_KEY}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      from:    'SATCO Alerts <alerts@satcoarabiaengg.com>',
      to:      alertEmail.split(',').map(e => e.trim()).filter(Boolean),
      subject,
      html,
    }),
  });

  const emailResult = await emailRes.json();
  if (!emailRes.ok) return res.status(500).json({ error: 'Email send failed', detail: emailResult });

  return res.status(200).json({
    message: 'Alert email sent', to: alertEmail,
    issues: issues.length, expired: expired.length, critical: critical.length, warning: warning.length,
    email_id: emailResult.id,
  });
}
