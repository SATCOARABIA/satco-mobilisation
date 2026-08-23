// api/supplier-expiry-check.js
// Checks BOTH supplier documents AND candidate documents for expiry
// Sends daily alert email via Resend
// Required Vercel env vars: RESEND_API_KEY, SUPABASE_SERVICE_KEY, SUPABASE_URL, CRON_SECRET

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  if (req.method === 'OPTIONS') return res.status(200).end();

  // Auth: allow same-origin (no token) OR valid CRON_SECRET bearer
  const authHeader = req.headers['authorization'] || '';
  const cronSecret = process.env.CRON_SECRET;
  if (cronSecret && authHeader.startsWith('Bearer ') && authHeader !== `Bearer ${cronSecret}`) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  if (cronSecret && !authHeader.startsWith('Bearer ')) {
    const origin = req.headers['origin'] || req.headers['referer'] || '';
    if (origin && !origin.includes('satco-mobilisation') && !origin.includes('localhost')) {
      return res.status(401).json({ error: 'Unauthorized' });
    }
  }

  const SUPA_URL = process.env.SUPABASE_URL || 'https://oaerqjrkdpuhiproppaz.supabase.co';
  const SUPA_KEY = process.env.SUPABASE_SERVICE_KEY;
  const RESEND_KEY = process.env.RESEND_API_KEY;

  if (!SUPA_KEY) return res.status(500).json({ error: 'Missing env var: SUPABASE_SERVICE_KEY — add it in Vercel Dashboard → Settings → Environment Variables' });
  if (!RESEND_KEY) return res.status(500).json({ error: 'Missing env var: RESEND_API_KEY — add it in Vercel Dashboard → Settings → Environment Variables' });

  const headers = { apikey: SUPA_KEY, Authorization: `Bearer ${SUPA_KEY}` };

  // ── Load settings ──────────────────────────────────────────────────────────
  const settingsRes = await fetch(`${SUPA_URL}/rest/v1/mob_supplier_alert_settings?limit=1`, { headers });
  const settingsArr = await settingsRes.json();
  const settings = settingsArr?.[0] || {};
  const alertEmail   = settings.alert_email || 'admin@satcoarabiaengg.com';
  const warnDays     = settings.warn_days  ?? 30;
  const criticalDays = settings.alert_days ?? 15;
  const enabled      = settings.enabled    ?? true;

  if (!enabled) return res.status(200).json({ message: 'Alerts disabled in settings.' });

  const today = new Date(); today.setHours(0,0,0,0);

  function daysLeft(dateStr) {
    if (!dateStr) return null;
    const d = new Date(dateStr); d.setHours(0,0,0,0);
    return Math.ceil((d - today) / 86400000);
  }
  function status(days) {
    if (days === null || days > warnDays) return null;
    if (days < 0)             return { level:'EXPIRED',  color:'#dc2626', emoji:'🔴' };
    if (days <= criticalDays) return { level:'CRITICAL', color:'#dc2626', emoji:'🔴' };
    return                           { level:'WARNING',  color:'#d97706', emoji:'🟡' };
  }

  const issues = { suppliers: [], candidates: [] };

  // ── SUPPLIER documents ─────────────────────────────────────────────────────
  const suppRes = await fetch(`${SUPA_URL}/rest/v1/mob_suppliers?select=*`, { headers });
  const suppliers = await suppRes.json();

  for (const s of (Array.isArray(suppliers) ? suppliers : [])) {
    const docs = [
      { label:'Trade / Commercial License',    date: s.license_expiry, ref: s.license_number||'—' },
      { label:'Workmen Comp. Insurance (WCI)', date: s.wci_expiry,     ref: s.wci_policy_no||'—' },
    ];
    for (const od of (Array.isArray(s.other_docs) ? s.other_docs : [])) {
      if (od.expiry) docs.push({ label: od.doc_name||'Other Document', date: od.expiry, ref:'—' });
    }
    for (const { label, date, ref } of docs) {
      const days = daysLeft(date); const st = status(days);
      if (!st) continue;
      issues.suppliers.push({ name: s.name, code: s.supplier_code||'—', doc: label, ref, days, status: st, expiry: date });
    }
  }

  // ── CANDIDATE documents ────────────────────────────────────────────────────
  const candRes = await fetch(
    `${SUPA_URL}/rest/v1/mob_candidates?active=eq.true&select=full_name,designation,supplier_name,passport_no,passport_expiry,emirates_id,eid_expiry,wp_no,wp_expiry,security_pass_no,cicpa_status,cicpa_location,adnoc_medical_status,adnoc_medical_date`,
    { headers }
  );
  const candidates = await candRes.json();

  for (const c of (Array.isArray(candidates) ? candidates : [])) {
    const name = c.full_name || 'Unknown';
    const title = c.designation || '—';
    const supp  = c.supplier_name || 'SATCO Own';
    const docs = [
      { label:'Passport',      date: c.passport_expiry, ref: c.passport_no||'—' },
      { label:'Emirates ID',   date: c.eid_expiry,      ref: c.emirates_id||'—' },
      { label:'Work Permit',   date: c.wp_expiry,       ref: c.wp_no||'—' },
    ];
    // CICPA — flag if confiscated or no pass
    if (c.cicpa_status === 'confiscated') {
      issues.candidates.push({ name, title, supp, doc:'CICPA Gate Pass', ref: c.security_pass_no||'—', days:-1, status:{level:'EXPIRED',color:'#dc2626',emoji:'🔴'}, expiry:'Confiscated', note:'CICPA Confiscated — reapply immediately' });
    } else if (c.cicpa_status === 'active' && c.security_pass_no) {
      // no expiry field on candidates — skip date check for now
    }
    for (const { label, date, ref } of docs) {
      const days = daysLeft(date); const st = status(days);
      if (!st) continue;
      issues.candidates.push({ name, title, supp, doc: label, ref, days, status: st, expiry: date });
    }
  }

  const totalIssues = issues.suppliers.length + issues.candidates.length;
  if (totalIssues === 0) {
    return res.status(200).json({ message: `All documents are valid — nothing expiring within ${warnDays} days.`, checked_suppliers: suppliers.length, checked_candidates: candidates.length });
  }

  // ── Build email ────────────────────────────────────────────────────────────
  const todayStr = today.toLocaleDateString('en-GB', { day:'2-digit', month:'long', year:'numeric' });

  function tableRows(list, type) {
    return list.map(i => {
      const label = type === 'supplier'
        ? `<strong>${i.name}</strong><br/><span style="font-size:11px;color:#64748b;">${i.code}</span>`
        : `<strong>${i.name}</strong><br/><span style="font-size:11px;color:#64748b;">${i.title} · ${i.supp}</span>`;
      const daysStr = i.days < 0 ? `EXPIRED ${Math.abs(i.days)}d ago` : `${i.days} day${i.days!==1?'s':''} left`;
      return `<tr>
        <td style="padding:8px 12px;border-bottom:1px solid #e2e8f0;">${i.status.emoji} ${label}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #e2e8f0;">${i.doc}${i.note?`<br/><span style="color:#dc2626;font-size:11px;">${i.note}</span>`:''}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #e2e8f0;font-family:monospace;font-size:12px;">${i.ref}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #e2e8f0;font-family:monospace;">${i.expiry||'—'}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #e2e8f0;font-weight:800;color:${i.status.color};">${daysStr}</td>
      </tr>`;
    }).join('');
  }

  function section(title, list, type, headBg, headColor) {
    if (!list.length) return '';
    return `
      <h3 style="font-size:13px;font-weight:800;color:${headColor};text-transform:uppercase;margin:20px 0 8px;">${title}</h3>
      <table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:20px;">
        <thead><tr style="background:${headBg};">
          <th style="padding:8px 12px;text-align:left;font-size:11px;color:${headColor};">${type==='supplier'?'Supplier':'Candidate'}</th>
          <th style="padding:8px 12px;text-align:left;font-size:11px;color:${headColor};">Document</th>
          <th style="padding:8px 12px;text-align:left;font-size:11px;color:${headColor};">Reference</th>
          <th style="padding:8px 12px;text-align:left;font-size:11px;color:${headColor};">Expiry Date</th>
          <th style="padding:8px 12px;text-align:left;font-size:11px;color:${headColor};">Status</th>
        </tr></thead>
        <tbody>${tableRows(list, type)}</tbody>
      </table>`;
  }

  const suppExpired  = issues.suppliers.filter(i=>i.days<0);
  const suppCritical = issues.suppliers.filter(i=>i.days>=0&&i.days<=criticalDays);
  const suppWarning  = issues.suppliers.filter(i=>i.days>criticalDays);
  const candExpired  = issues.candidates.filter(i=>i.days<0);
  const candCritical = issues.candidates.filter(i=>i.days>=0&&i.days<=criticalDays);
  const candWarning  = issues.candidates.filter(i=>i.days>criticalDays);

  const totalExpired = suppExpired.length + candExpired.length;

  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8"/></head>
<body style="font-family:'Segoe UI',sans-serif;background:#f1f5f9;padding:24px;color:#0f172a;margin:0;">
<div style="max-width:800px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,0.08);">
  <div style="background:#0f172a;padding:20px 28px;">
    <div style="font-size:11px;font-weight:800;color:#b8972a;letter-spacing:2px;">SATCO ARABIA</div>
    <div style="font-size:18px;font-weight:700;color:#fff;margin-top:3px;">Document Expiry Alert — Daily Report</div>
    <div style="font-size:12px;color:#94a3b8;margin-top:3px;">${todayStr}</div>
  </div>

  <div style="padding:14px 28px;background:${totalExpired?'#fef2f2':'#fffbeb'};border-bottom:1px solid ${totalExpired?'#fecaca':'#fde68a'};">
    <div style="font-weight:800;font-size:15px;color:${totalExpired?'#991b1b':'#92400e'};">
      ${totalIssues} document${totalIssues!==1?'s':''} require attention
    </div>
    <div style="display:flex;gap:16px;margin-top:6px;flex-wrap:wrap;">
      ${issues.suppliers.length?`<span style="font-size:12px;color:#64748b;">🏭 <strong>${issues.suppliers.length}</strong> Supplier docs</span>`:''}
      ${issues.candidates.length?`<span style="font-size:12px;color:#64748b;">👤 <strong>${issues.candidates.length}</strong> Candidate docs</span>`:''}
      ${totalExpired?`<span style="font-size:12px;font-weight:800;color:#dc2626;">🔴 ${totalExpired} EXPIRED</span>`:''}
    </div>
  </div>

  <div style="padding:20px 28px;">
    ${issues.suppliers.length ? `<div style="font-size:14px;font-weight:800;color:#0f172a;padding:10px 14px;background:#f8fafc;border-radius:8px;border-left:4px solid #b8972a;margin-bottom:4px;">🏭 SUPPLIER DOCUMENTS</div>` : ''}
    ${section('🔴 Expired — Immediate Action', suppExpired, 'supplier', '#fee2e2', '#991b1b')}
    ${section(`🔴 Critical — ≤${criticalDays} days`, suppCritical, 'supplier', '#fee2e2', '#991b1b')}
    ${section(`🟡 Warning — ≤${warnDays} days`, suppWarning, 'supplier', '#fef3c7', '#92400e')}

    ${issues.candidates.length ? `<div style="font-size:14px;font-weight:800;color:#0f172a;padding:10px 14px;background:#f8fafc;border-radius:8px;border-left:4px solid #2563eb;margin-bottom:4px;">👤 CANDIDATE DOCUMENTS</div>` : ''}
    ${section('🔴 Expired — Immediate Action', candExpired, 'candidate', '#fee2e2', '#991b1b')}
    ${section(`🔴 Critical — ≤${criticalDays} days`, candCritical, 'candidate', '#fee2e2', '#991b1b')}
    ${section(`🟡 Warning — ≤${warnDays} days`, candWarning, 'candidate', '#fef3c7', '#92400e')}

    <div style="margin-top:12px;padding:12px 16px;background:#f8fafc;border-radius:8px;font-size:12px;color:#64748b;">
      Checked: <strong>${suppliers.length}</strong> suppliers · <strong>${candidates.length}</strong> candidates · Alert window: ${warnDays} days warning / ${criticalDays} days critical<br/>
      Update documents at <a href="https://satco-mobilisation.vercel.app" style="color:#0369a1;font-weight:700;">satco-mobilisation.vercel.app</a>
    </div>
  </div>
  <div style="padding:12px 28px;background:#f8fafc;border-top:1px solid #e2e8f0;font-size:11px;color:#94a3b8;">
    SATCO Arabia General Contracting LLC — SPC · Automated daily alert · 06:00 UAE time
  </div>
</div></body></html>`;

  const subject = totalExpired
    ? `🚨 SATCO Alert — ${totalExpired} EXPIRED document${totalExpired>1?'s':''} (${issues.suppliers.length} supplier + ${issues.candidates.length} candidate)`
    : `⚠️ SATCO Alert — ${totalIssues} document${totalIssues>1?'s':''} expiring (${issues.suppliers.length} supplier + ${issues.candidates.length} candidate)`;

  const emailRes = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: { Authorization: `Bearer ${RESEND_KEY}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      from: 'SATCO Alerts <alerts@satcoarabiaengg.com>',
      to: alertEmail.split(',').map(e=>e.trim()).filter(Boolean),
      subject, html,
    }),
  });

  const emailResult = await emailRes.json();
  if (!emailRes.ok) return res.status(500).json({ error: 'Email send failed', detail: emailResult });

  return res.status(200).json({
    message: `Alert sent to ${alertEmail}`,
    supplier_issues: issues.suppliers.length,
    candidate_issues: issues.candidates.length,
    total: totalIssues,
    email_id: emailResult.id,
  });
}
