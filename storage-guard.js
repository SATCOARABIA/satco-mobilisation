// SATCO storage guard (loaded before the app scripts).
//
// The HR document buckets (hr-documents, cv-uploads) are PRIVATE: a document can only be opened by a
// signed-in portal user. Older records still hold "…/storage/v1/object/public/<bucket>/<path>" links,
// so this script transparently turns any such link into a short-lived signed link at the moment it is
// used — in <a href>, <img src>, <iframe>/<embed>/<object>, window.open(...) and fetch(...).
// It also adds the signed-in session to this portal's own /api/claude and /api/telegram calls so those
// endpoints can refuse anyone who is not logged in.
(function () {
  var SUPA = 'https://oaerqjrkdpuhiproppaz.supabase.co';
  var ANON = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9hZXJxanJrZHB1aGlwcm9wcGF6Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk5NTQ0NjksImV4cCI6MjA5NTUzMDQ2OX0.qBtb3OV1aFGX8e1QUg19qZmOwIIjipF6IZwBOLXY3YI';
  var BUCKETS = ['hr-documents', 'cv-uploads'];
  var RE = new RegExp('^' + SUPA.replace(/[.]/g, '\\.') + '/storage/v1/object/(?:public|authenticated)/(' + BUCKETS.join('|') + ')/([^?#]+)');
  var cache = {}; // url -> {signed, exp}
  var realFetch = window.fetch.bind(window);

  function token() {
    try {
      for (var i = 0; i < localStorage.length; i++) {
        var k = localStorage.key(i);
        if (/^sb-oaerqjrkdpuhiproppaz-auth-token/.test(k)) {
          var v = JSON.parse(localStorage.getItem(k));
          var s = v && (v.access_token ? v : (v.currentSession || v.session));
          if (s && s.access_token) return s.access_token;
        }
      }
    } catch (e) {}
    return null;
  }
  function match(url) { return typeof url === 'string' ? RE.exec(url) : null; }

  async function sign(url) {
    var m = match(url); if (!m) return url;
    var c = cache[url]; if (c && c.exp > Date.now()) return c.signed;
    var t = token(); if (!t) return url;
    var path = decodeURIComponent(m[2]);
    var r = await realFetch(SUPA + '/storage/v1/object/sign/' + m[1] + '/' + path.split('/').map(encodeURIComponent).join('/'), {
      method: 'POST', headers: { apikey: ANON, Authorization: 'Bearer ' + t, 'Content-Type': 'application/json' },
      body: JSON.stringify({ expiresIn: 600 })
    });
    if (!r.ok) return url;
    var j = await r.json();
    var signed = SUPA + '/storage/v1' + (j.signedURL || j.signedUrl);
    cache[url] = { signed: signed, exp: Date.now() + 540000 };
    return signed;
  }
  window.satcoSignDocUrl = sign;

  // fetch(): public document URL -> signed URL; own /api/claude, /api/telegram -> add session
  window.fetch = function (input, init) {
    var url = typeof input === 'string' ? input : (input && input.url);
    if (match(url)) return sign(url).then(function (s) { return realFetch(s, init); });
    try {
      var u = new URL(url, location.href);
      if (u.origin === location.origin && (u.pathname === '/api/claude' || u.pathname === '/api/telegram')) {
        var t = token();
        if (t) {
          init = Object.assign({}, init || {});
          var h = new Headers(init.headers || (typeof input !== 'string' && input.headers) || {});
          if (!h.has('Authorization')) h.set('Authorization', 'Bearer ' + t);
          init.headers = h;
        }
      }
    } catch (e) {}
    return realFetch(input, init);
  };

  // window.open(): open the tab immediately (keeps popup blockers happy), then point it at the signed link
  var realOpen = window.open.bind(window);
  window.open = function (url, target, features) {
    if (!match(url)) return realOpen.apply(window, arguments);
    var w = realOpen('about:blank', target || '_blank', features);
    sign(url).then(function (s) { if (w) w.location.href = s; else location.href = s; });
    return w;
  };

  // links: sign on click
  document.addEventListener('click', function (e) {
    var a = e.target && e.target.closest && e.target.closest('a[href]');
    if (!a || !match(a.href)) return;
    e.preventDefault(); e.stopPropagation();
    var newTab = a.target === '_blank' || e.ctrlKey || e.metaKey;
    var w = newTab ? realOpen('about:blank', '_blank') : null;
    sign(a.href).then(function (s) {
      if (a.hasAttribute('download')) { var x = document.createElement('a'); x.href = s; x.download = a.getAttribute('download') || ''; x.target = '_blank'; document.body.appendChild(x); x.click(); x.remove(); if (w) w.close(); }
      else if (w) w.location.href = s; else location.href = s;
    });
  }, true);

  // images / embedded viewers: swap src for a signed link
  var ATTRS = { IMG: 'src', IFRAME: 'src', EMBED: 'src', SOURCE: 'src', OBJECT: 'data' };
  function fix(el) {
    var attr = ATTRS[el.tagName]; if (!attr) return;
    var v = el.getAttribute(attr); if (!match(v)) return;
    sign(v).then(function (s) { if (s !== v && el.getAttribute(attr) === v) el.setAttribute(attr, s); });
  }
  function scan(root) {
    if (root.nodeType !== 1) return;
    fix(root);
    if (root.querySelectorAll) root.querySelectorAll('img,iframe,embed,source,object').forEach(fix);
  }
  new MutationObserver(function (list) {
    list.forEach(function (m) {
      if (m.type === 'attributes') fix(m.target);
      else m.addedNodes.forEach(scan);
    });
  }).observe(document.documentElement, { subtree: true, childList: true, attributes: true, attributeFilter: ['src', 'data'] });
})();
