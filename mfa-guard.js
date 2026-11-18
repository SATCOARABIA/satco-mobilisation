// SATCO two-step login (authenticator app) + auto logout.
// Load right after supabase-js:  <script src="./mfa-guard.js" data-url="https://<ref>.supabase.co" data-key="<anon key>"></script>
// After the normal email/password login, this asks for the 6-digit code from Google / Microsoft
// Authenticator. First time: it shows a QR code to scan. The databases refuse to return any data
// until the code has been entered (database rule "aal2"), so a stolen password alone shows nothing.
// Also signs the user out after 30 minutes without any activity.
(function () {
  var me = document.currentScript;
  var URL_ = me.getAttribute('data-url'), KEY = me.getAttribute('data-key');
  var IDLE_MIN = 30;
  var client = null, busy = false, box = null, lastActive = Date.now();

  function css() {
    if (document.getElementById('mfa-css')) return;
    var s = document.createElement('style'); s.id = 'mfa-css';
    s.textContent = '#mfa-ov{position:fixed;inset:0;z-index:2147483647;background:#0f2447;display:flex;align-items:center;justify-content:center;font-family:Segoe UI,Arial,sans-serif}' +
      '#mfa-ov .c{background:#16305c;border:1px solid #c9a227;border-radius:12px;padding:28px 30px;width:340px;max-width:92vw;color:#fff;text-align:center;box-shadow:0 10px 40px rgba(0,0,0,.4)}' +
      '#mfa-ov h2{margin:0 0 4px;font-size:19px;color:#e6c35c}#mfa-ov p{font-size:13px;line-height:1.45;color:#d6deeb;margin:8px 0}' +
      '#mfa-ov input{width:170px;font-size:26px;letter-spacing:8px;text-align:center;padding:8px;border-radius:8px;border:1px solid #c9a227;background:#0f2447;color:#fff;margin:10px 0}' +
      '#mfa-ov button{background:#c9a227;color:#0f2447;border:0;border-radius:8px;padding:10px 22px;font-weight:700;font-size:14px;cursor:pointer}' +
      '#mfa-ov .l{display:block;margin-top:14px;color:#9fb3d1;font-size:12px;cursor:pointer;text-decoration:underline;background:none;padding:0;font-weight:400}' +
      '#mfa-ov .e{color:#ff9b9b;font-size:13px;min-height:18px}#mfa-ov img{background:#fff;border-radius:8px;padding:8px;width:180px;height:180px}' +
      '#mfa-ov code{display:block;font-size:11px;word-break:break-all;color:#e6c35c;margin-top:6px}';
    document.head.appendChild(s);
  }
  function show(html) {
    css();
    if (!box) { box = document.createElement('div'); box.id = 'mfa-ov'; document.body.appendChild(box); }
    box.innerHTML = '<div class="c">' + html + '<button class="l" id="mfa-out">Sign out</button></div>';
    document.getElementById('mfa-out').onclick = function () { client.auth.signOut().finally(function () { location.reload(); }); };
  }
  function hide() { if (box) { box.remove(); box = null; } }

  function codeForm(title, intro, extra, onCode) {
    show('<h2>' + title + '</h2>' + (extra || '') + '<p>' + intro + '</p>' +
      '<input id="mfa-code" inputmode="numeric" autocomplete="one-time-code" maxlength="6" placeholder="000000">' +
      '<div class="e" id="mfa-err"></div><button id="mfa-go">Verify</button>');
    var inp = document.getElementById('mfa-code'), btn = document.getElementById('mfa-go'), err = document.getElementById('mfa-err');
    setTimeout(function () { inp.focus(); }, 50);
    function go() {
      var code = inp.value.replace(/\D/g, '');
      if (code.length !== 6) { err.textContent = 'Enter the 6-digit code'; return; }
      btn.disabled = true; btn.textContent = 'Checking…'; err.textContent = '';
      onCode(code).then(function () { location.reload(); })
        .catch(function (e) { btn.disabled = false; btn.textContent = 'Verify'; err.textContent = (e && e.message) || 'Wrong code — try again'; inp.select(); });
    }
    btn.onclick = go; inp.onkeydown = function (e) { if (e.key === 'Enter') go(); };
  }
  function qrSrc(q) { q = String(q || ''); return q.indexOf('data:') === 0 ? q : 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(q); }
  function unwrap(r) { if (r.error) throw r.error; return r.data; }

  async function check() {
    if (busy || !client) return;
    busy = true;
    try {
      var s = unwrap(await client.auth.getSession()).session;
      if (!s) { hide(); return; }
      var aal = unwrap(await client.auth.mfa.getAuthenticatorAssuranceLevel());
      if (aal.currentLevel === 'aal2') { hide(); return; }
      if (box) return; // form already showing
      var f = unwrap(await client.auth.mfa.listFactors());
      var verified = (f.totp || []).filter(function (x) { return x.status === 'verified'; });
      if (verified.length) {
        var id = verified[0].id;
        codeForm('Enter your security code', 'Open <b>Google Authenticator</b> (or Microsoft Authenticator) on your phone and type the 6-digit code for <b>SATCO</b>.', '',
          function (code) { return client.auth.mfa.challengeAndVerify({ factorId: id, code: code }).then(unwrap); });
      } else {
        // first time: remove half-finished setups, then show QR
        for (var i = 0; i < (f.all || []).length; i++) if (f.all[i].status !== 'verified') await client.auth.mfa.unenroll({ factorId: f.all[i].id });
        var en = unwrap(await client.auth.mfa.enroll({ factorType: 'totp', friendlyName: 'SATCO ' + location.hostname.split('.')[0] + ' ' + Date.now() }));
        codeForm('Set up two-step login',
          '1. Install <b>Google Authenticator</b> on your phone.<br>2. Tap <b>+</b> → <b>Scan a QR code</b> and scan the code above.<br>3. Type the 6-digit code it shows.',
          '<img alt="QR" src="' + qrSrc(en.totp.qr_code) + '"><code>Key: ' + en.totp.secret + '</code>',
          function (code) { return client.auth.mfa.challengeAndVerify({ factorId: en.id, code: code }).then(unwrap); });
      }
    } catch (e) { console.warn('mfa-guard', e); }
    finally { busy = false; }
  }

  // auto logout after inactivity
  ['mousemove', 'keydown', 'click', 'scroll', 'touchstart'].forEach(function (ev) { window.addEventListener(ev, function () { lastActive = Date.now(); }, { passive: true }); });
  setInterval(function () {
    if (!client || Date.now() - lastActive < IDLE_MIN * 60000) return;
    client.auth.getSession().then(function (r) {
      if (r.data && r.data.session) client.auth.signOut().finally(function () { location.reload(); });
    });
    lastActive = Date.now();
  }, 60000);

  function start() {
    if (!window.supabase || !window.supabase.createClient) return setTimeout(start, 200);
    client = window.supabase.createClient(URL_, KEY, { auth: { autoRefreshToken: false, persistSession: true, detectSessionInUrl: false } });
    setTimeout(check, 800);
    setInterval(check, 15000);
    window.addEventListener('storage', function (e) { if (e.key && e.key.indexOf('-auth-token') > 0) setTimeout(check, 300); });
    document.addEventListener('visibilitychange', function () { if (!document.hidden) check(); });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start); else start();
})();
