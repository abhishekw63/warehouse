/* ============================================================================
 * shell-nav.js — App-shell partial navigation for the B2B app.
 *
 * Keeps the SIDEBAR + HEADER DOM 100% intact across page changes: instead of a
 * full document reload (which destroys and rebuilds the shell → the "blink"),
 * in-app link clicks are intercepted, the next page is fetched, and ONLY
 * #MainContent is swapped in. The shell nodes are never re-rendered.
 *
 * Progressive enhancement + bulletproof fallback: ANY uncertainty (cross-origin,
 * downloads, fetch error, missing #MainContent, exception) falls back to a normal
 * full navigation, so navigation can never end up broken. Kill-switch: remove the
 * <script> include in base_b2b, or add data-no-boost to a link/<html>.
 * ========================================================================== */
(function () {
  var d = document, w = window;
  if (d.documentElement.hasAttribute('data-no-boost')) return;
  var shell = d.querySelector('.b2b-app');
  var main = d.getElementById('MainContent');
  if (!shell || !main || !w.history || !w.history.pushState || !w.DOMParser || !w.fetch) return;

  function fullNav(url) { w.location.href = url; }

  // Which anchors do we take over? Same-origin, in-app, plain left-clicks only.
  function candidate(a) {
    if (!a || a.defaultPrevented) return null;
    if (a.target && a.target !== '_self') return null;
    if (a.hasAttribute('download') || a.hasAttribute('data-no-boost')) return null;
    if (a.hasAttribute('data-download-modal') || a.hasAttribute('data-download-bg')) return null;
    var href = a.getAttribute('href') || '';
    if (!href || href.charAt(0) === '#') return null;
    var low = href.toLowerCase();
    if (low.indexOf('javascript:') === 0 || low.indexOf('mailto:') === 0 || low.indexOf('tel:') === 0) return null;
    var url; try { url = new URL(a.href, location.href); } catch (e) { return null; }
    if (url.origin !== location.origin) return null;
    return url.href;
  }

  // Re-execute <script> tags found inside freshly-swapped content (innerHTML does
  // NOT run them). External libs already present are skipped so they don't reload.
  function reExecInline(root) {
    root.querySelectorAll('script').forEach(function (old) {
      var src = old.getAttribute('src');
      if (src && d.querySelector('script[src="' + src + '"]')) return;   // already loaded
      var s = d.createElement('script');
      [].forEach.call(old.attributes, function (at) { s.setAttribute(at.name, at.value); });
      if (!src) s.textContent = old.textContent;
      old.parentNode.replaceChild(s, old);
    });
  }

  // Page-specific CSS (extra_css links in <head>): add the incoming page's, drop
  // the previous page's. Shared sheets (b2b.css/enhance.css) stay untouched.
  function syncPageCss(doc) {
    var have = {};
    d.querySelectorAll('head link[rel="stylesheet"]').forEach(function (l) { have[l.getAttribute('href')] = 1; });
    var want = {};
    doc.querySelectorAll('head link[rel="stylesheet"]').forEach(function (l) {
      var h = l.getAttribute('href'); if (!h) return; want[h] = 1;
      if (!have[h]) { var nl = d.createElement('link'); nl.rel = 'stylesheet'; nl.href = h; nl.setAttribute('data-nav-css', '1'); d.head.appendChild(nl); }
    });
    d.querySelectorAll('head link[rel="stylesheet"][data-nav-css="1"]').forEach(function (l) {
      if (!want[l.getAttribute('href')]) l.remove();
    });
  }

  // Page-specific external JS (extra_js, marked data-page-js): remove the previous
  // page's and (re-)add the incoming page's so it re-runs even on a return visit.
  function syncPageJs(doc) {
    d.querySelectorAll('script[data-page-js]').forEach(function (s) { s.remove(); });
    doc.querySelectorAll('script[data-page-js]').forEach(function (s) {
      var ns = d.createElement('script');
      [].forEach.call(s.attributes, function (at) { ns.setAttribute(at.name, at.value); });
      if (!s.getAttribute('src')) ns.textContent = s.textContent;
      d.body.appendChild(ns);
    });
  }

  // Mirror the server's own active-nav decision: read which links the FETCHED page
  // rendered as .on and apply that to the live (never-rebuilt) sidebar. Accurate,
  // no prefix-guessing. Falls back to no-op if the fetched sidebar isn't found.
  function setActive(doc) {
    var want = {};
    doc.querySelectorAll('.b2b-side .sn, .b2b-side .side-group').forEach(function (a) {
      var key = a.getAttribute('href') || a.id;
      if (key && (a.classList.contains('on') || a.getAttribute('aria-expanded') === 'true')) want[key] = 1;
    });
    d.querySelectorAll('.b2b-side .sn').forEach(function (a) {
      var h = a.getAttribute('href'); if (h) a.classList.toggle('on', !!want[h]);
    });
    // Collapsible groups (Record Verify, Admin, …): the persistent sidebar isn't
    // re-rendered on partial nav, so the server-side auto-open for the ACTIVE
    // section was never re-applied — navigating to /record-verify left its group
    // collapsed (its two sub-links hidden). Re-sync each group's open state from
    // the freshly-fetched sidebar: force-open the active section's group; leave
    // the rest as the user left them (never fight a manual toggle).
    doc.querySelectorAll('.b2b-side .side-subnav[id]').forEach(function (nsub) {
      if (!nsub.classList.contains('open')) return;      // only the active section
      var cur = d.getElementById(nsub.id); if (!cur) return;
      cur.classList.add('open');
      var btn = d.querySelector('.b2b-side [aria-controls="' + nsub.id + '"]');
      if (btn) btn.setAttribute('aria-expanded', 'true');
    });
  }

  var token = 0, bar = d.getElementById('navProgress');
  function barStart() { if (bar) { bar.className = 'np on creep'; bar.style.width = '0'; requestAnimationFrame(function () { bar.style.width = '90%'; }); } }
  function barDone() { if (!bar) return; bar.className = 'np on finish'; bar.style.width = '100%'; setTimeout(function () { bar.style.opacity = '0'; setTimeout(function () { bar.className = 'np'; bar.style.width = '0'; bar.style.opacity = ''; }, 320); }, 180); }

  // Fetch a shell page's HTML (with the same redirect/ok guard the nav uses).
  function fetchText(url) {
    return fetch(url, { credentials: 'same-origin', headers: { 'X-Requested-With': 'fetch' } })
      .then(function (r) { if (!r.ok || (r.redirected && new URL(r.url).pathname !== new URL(url).pathname)) throw 0; return r.text(); });
  }

  // ── Hover/press PREFETCH ── the big perceived-speed win: warm the next page in
  // the background while the pointer is still on the link, so the click swaps an
  // already-fetched document instantly. Entries self-expire so nothing goes stale.
  var pf = {};                                     // url -> { p: Promise<html> }
  function prefetch(url) {
    if (pf[url]) return;
    var rec = { p: fetchText(url).catch(function () { if (pf[url] === rec) delete pf[url]; throw 0; }) };
    pf[url] = rec;
    w.setTimeout(function () { if (pf[url] === rec) delete pf[url]; }, 10000);
  }
  function grab(url) {                             // reuse a warm prefetch if we have one
    var rec = pf[url]; if (rec) { delete pf[url]; return rec.p; }
    return fetchText(url);
  }

  function navigate(url, push) {
    var my = ++token;
    barStart();
    grab(url)
      .then(function (html) {
        if (my !== token) return;                       // a newer nav superseded us
        var doc = new DOMParser().parseFromString(html, 'text/html');
        var nm = doc.getElementById('MainContent');
        if (!nm) { fullNav(url); return; }              // not a shell page → full nav
        try {
          syncPageCss(doc);
          d.title = doc.title || d.title;
          var bc = d.querySelector('.header__breadcrumb'), nbc = doc.querySelector('.header__breadcrumb');
          if (bc && nbc) bc.innerHTML = nbc.innerHTML;
          // Swap ONLY the middle — the sidebar + header DOM are never touched.
          main.innerHTML = nm.innerHTML;
          reExecInline(main);
          syncPageJs(doc);
          setActive(doc);
          if (w.B2B) {
            B2B.enhanceTables && B2B.enhanceTables(main);
            B2B.countUp && B2B.countUp(main);
            B2B.revealBars && B2B.revealBars(main);
            B2B.applyViewOnly && B2B.applyViewOnly(main);
            B2B.bindLoading && B2B.bindLoading(main);
          }
          // subtle fade-in of the new content only (shell stays put)
          main.classList.add('nav-in');
          setTimeout(function () { main.classList.remove('nav-in'); }, 280);
          if (push) w.history.pushState({ shellNav: 1 }, '', url);
          w.scrollTo(0, 0);
          barDone();
        } catch (e) { fullNav(url); }
      })
      .catch(function () { fullNav(url); });
  }

  // Prefetch on hover (after a tiny delay so a quick pass-over doesn't fire) and on
  // press — whichever lands first warms the page before the click resolves.
  var hoverT;
  d.addEventListener('mouseover', function (e) {
    var a = e.target.closest && e.target.closest('a[href]'); if (!a) return;
    // Scope hover-prefetch to real navigation (sidebar), not arbitrary in-content links.
    if (!(a.classList.contains('sn') || (a.closest && a.closest('.b2b-side')))) return;
    var url = candidate(a); if (!url || url === location.href) return;
    clearTimeout(hoverT); hoverT = w.setTimeout(function () { prefetch(url); }, 60);
  });
  d.addEventListener('mouseout', function () { clearTimeout(hoverT); });
  d.addEventListener('mousedown', function (e) {
    if (e.button !== 0) return;
    var a = e.target.closest && e.target.closest('a[href]'); if (!a) return;
    var url = candidate(a); if (url && url !== location.href) prefetch(url);
  });

  d.addEventListener('click', function (e) {
    if (e.button !== 0 || e.metaKey || e.ctrlKey || e.shiftKey || e.altKey) return;
    var a = e.target.closest('a[href]');
    if (!a) return;
    var url = candidate(a);
    if (!url) return;
    e.preventDefault();
    if (url === location.href) return;
    navigate(url, true);
  });
  w.addEventListener('popstate', function () { navigate(location.href, false); });
})();
