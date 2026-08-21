/* online_b2b/upload.html — page script (separated). margins via #upload-cfg JSON. */
var CFG = JSON.parse(document.getElementById("upload-cfg").textContent);
(function () {
  // Auto-fill Margin % to the selected marketplace's default (Flipkart 77, Blink 70).
  var MARGINS = CFG.margins;
  var mpSel = document.getElementById('id_marketplace');
  var marginIn = document.getElementById('id_margin_pct');
  if (mpSel && marginIn) {
    var applyMargin = function () {
      var def = MARGINS[mpSel.value];
      if (def !== undefined && def !== null) marginIn.value = def;
    };
    applyMargin();                         // set correct default on first load
    mpSel.addEventListener('change', applyMargin);
  }

  // Per-marketplace file hint — show the picked marketplace's needs + a link to
  // its full template. Shown for EVERY marketplace (generic fallback if no note).
  var HINTS = {}, TPLS = [], FORMATS = {};
  try { HINTS = JSON.parse(document.getElementById('mp-hints').textContent || '{}'); } catch (e) {}
  try { TPLS = JSON.parse(document.getElementById('mp-templates').textContent || '[]'); } catch (e) {}
  try { FORMATS = JSON.parse(document.getElementById('mp-formats').textContent || '{}'); } catch (e) {}
  var hintBox = document.getElementById('mp-hint');
  var fmtBadge = document.getElementById('mp-fmt-badge');
  // "expects XLSX / PDF" badge beside the picker — same file_type the profile shows.
  function applyBadge() {
    if (!fmtBadge || !mpSel) return;
    // Collapse the accepted extensions to the primary type(s) — a quick
    // "expects XLSX / PDF" cue that matches the profile card's badge.
    var r = (FORMATS[mpSel.value] || '').toLowerCase();
    var hasPdf = r.indexOf('pdf') !== -1, hasXls = r.indexOf('xls') !== -1, hasCsv = r.indexOf('csv') !== -1;
    var label = (hasXls && hasPdf) ? 'XLSX / PDF'
              : hasPdf ? 'PDF' : hasXls ? 'XLSX' : hasCsv ? 'CSV'
              : r.replace(/\./g, '').toUpperCase().trim();
    if (label) { fmtBadge.textContent = label; fmtBadge.hidden = false; }
    else { fmtBadge.hidden = true; }
  }
  if (mpSel && hintBox) {
    var applyHint = function () {
      var mp = mpSel.value;
      var h = HINTS[mp] || 'the marketplace’s PO file(s) — see the full template for the exact columns.';
      var link = (TPLS.indexOf(mp) !== -1)
        ? ' <a href="/b2b/rules/template/' + encodeURIComponent(mp) + '/" style="color:#1f4fb3;font-weight:700;text-decoration:underline;">Full detail →</a>'
        : '';
      hintBox.innerHTML = '📎 <b>' + mp + ' needs:</b> ' + h + link;
      hintBox.style.display = 'block';
      applyBadge();
    };
    applyHint();
    mpSel.addEventListener('change', applyHint);
  } else {
    applyBadge();
    if (mpSel) mpSel.addEventListener('change', applyBadge);
  }

  var dz = document.getElementById('dz');
  if (!dz) return;
  var input = dz.querySelector('input[type=file]');
  var main = document.getElementById('dz-main');
  var form = document.getElementById('up-form'), btn = document.getElementById('up-submit');
  // Process stays disabled until a file is chosen AND the "import" feedback finishes,
  // so the operator clearly sees the click registered + the file loading.
  btn.disabled = true;
  input.addEventListener('change', function () {
    if (!input.files.length) {
      main.textContent = 'Click to choose PO file(s)';
      dz.classList.remove('dz-loading', 'dz-ready'); btn.disabled = true; return;
    }
    var name = input.files.length === 1 ? input.files[0].name
                                        : input.files.length + ' files selected';
    dz.classList.remove('dz-ready'); dz.classList.add('dz-loading');
    main.textContent = 'Importing ' + name + '…';
    btn.disabled = true; btn.textContent = 'Importing…';
    // brief read/import feedback (min visible time so it's seen even for small files)
    setTimeout(function () {
      dz.classList.remove('dz-loading'); dz.classList.add('dz-ready');
      main.textContent = '✓ ' + name;
      btn.disabled = false; btn.textContent = 'Process';
    }, 850);
  });
  var overlay = document.getElementById('proc-overlay');
  // Portal the overlay to <body> so its position:fixed centres on the VIEWPORT,
  // not within #MainContent (padding-left:252px for the sidebar + the page-in
  // transform were pulling the card off-centre / low). Mirrors the shared
  // body-level #b2b-load overlay pattern.
  if (overlay && overlay.parentNode !== document.body) document.body.appendChild(overlay);
  var fill = document.getElementById('proc-fill'), pctEl = document.getElementById('proc-pct'),
      stageEl = document.getElementById('proc-stage'), timeEl = document.getElementById('proc-time');
  var STAGES = [
    { p: 0,  t: 'Loading master & Ship-To mapping…' },
    { p: 42, t: 'Reading & processing PO file…' },
    { p: 68, t: 'Validating against master…' },
    { p: 86, t: 'Building preview & workbook…' },
    { p: 95, t: 'Finalizing…' }
  ];
  function clock(ms) { var s = Math.floor(ms / 1000); return Math.floor(s / 60) + ':' + ('0' + (s % 60)).slice(-2); }

  var card = overlay.querySelector('.proc-card');
  var titleEl = overlay.querySelector('.proc-title');
  var spin = overlay.querySelector('.proc-spin');

  form.addEventListener('submit', function (e) {
    if (!input.files.length) return;          // let native "required" handle empty
    e.preventDefault();                       // AJAX: drive the overlay ourselves
    btn.disabled = true; btn.textContent = 'Processing…';
    overlay.hidden = false;
    // Perceived staged progress while the engine runs; the elapsed timer is
    // real, the % is a smooth asymptotic estimate that holds near 96% until the
    // server responds — then we snap to 100% and show a definitive ✓ Imported.
    var start = Date.now(), EST = 14000, pct = 0, done = false;
    var tick = setInterval(function () {
      if (done) return;
      var el = Date.now() - start;
      var target = 96 * (1 - Math.exp(-el / (EST * 0.55)));
      pct = Math.max(pct, target);
      fill.style.width = pct.toFixed(0) + '%';
      pctEl.textContent = pct.toFixed(0) + '%';
      var st = STAGES[0];
      for (var i = 0; i < STAGES.length; i++) { if (pct >= STAGES[i].p) st = STAGES[i]; }
      stageEl.textContent = st.t;
      timeEl.textContent = 'elapsed ' + clock(el);
    }, 150);

    function fail(msg) {
      done = true; clearInterval(tick);
      card.classList.add('proc-err');
      if (spin) spin.style.display = 'none';
      titleEl.textContent = '⚠ Import failed';
      // Uniform, always-informative reason. The backend returns the real cause;
      // if it's blank or the bare generic, show a consistent guidance line so the
      // operator NEVER sees an empty/"Starting…" state on failure.
      var reason = (msg == null ? '' : String(msg)).trim();
      if (!reason || /^import failed\.?$/i.test(reason)) {
        reason = "Couldn't process the file. Check it's the right marketplace's PO "
               + 'in the expected format (open “Full detail” for the exact columns) '
               + 'and that no sheet or column was renamed/removed, then retry.';
      }
      stageEl.textContent = reason;
      fill.style.width = '100%'; if (fill.classList) fill.classList.add('proc-fill-err');
      pctEl.textContent = '';
      timeEl.innerHTML = '<button type="button" id="proc-close" class="proc-close">Close</button>';
      var c = document.getElementById('proc-close');
      if (c) c.addEventListener('click', function () {
        overlay.hidden = true; card.classList.remove('proc-err');
        if (spin) spin.style.display = ''; titleEl.textContent = 'Processing PO';
        btn.disabled = false; btn.textContent = 'Process';
      });
    }

    function doFetch() {
      fetch(form.action || window.location.href, {
        method: 'POST', body: new FormData(form),
        headers: { 'X-Requested-With': 'XMLHttpRequest' }, credentials: 'same-origin'
      }).then(function (r) { return r.json(); }).then(function (j) {
        if (!j.ok) { fail(j.error); return; }
        done = true; clearInterval(tick);
        // snap to 100% + a clear completion the operator can't miss
        card.classList.add('proc-done');
        if (spin) spin.style.display = 'none';
        fill.style.width = '100%'; pctEl.textContent = '100%';
        var el = Date.now() - start;
        titleEl.innerHTML = '✓ Imported';
        stageEl.textContent = j.pos + ' PO(s) · ' + j.lines + ' line(s)' +
          (j.affected ? ' · ' + j.affected + ' to review' : ' · all clean') +
          (j.warnings ? ' · ' + j.warnings + ' warning(s)' : '');
        timeEl.textContent = 'done in ' + clock(el) + ' — opening review…';
        setTimeout(function () { window.location = j.review_url; }, 1100);
      }).catch(function (err) {
        var name = (err && err.name) ? err.name : 'NetworkError';
        fail('Could not read/send the file(s) — is one still OPEN in Excel? ' +
             'Close them, then retry. [' + name + ']');
      });
    }

    // Pre-flight: try to read 1 byte of each selected file. A file OPEN/locked in
    // Excel rejects here — so we can name the exact offender(s) instead of failing
    // the whole batch with a vague error (matters most for many-file uploads).
    Promise.all(Array.prototype.map.call(input.files, function (f) {
      return f.slice(0, 1).arrayBuffer().then(function () { return null; },
                                              function () { return f.name; });
    })).then(function (bad) {
      bad = (bad || []).filter(Boolean);
      if (bad.length) {
        fail('Close these file(s) in Excel, then retry (' + bad.length + '): ' +
             bad.join(', '));
      } else {
        doFetch();
      }
    });
  });
})();

// Live marketplace profile panel — swaps to the selected MP's profile via AJAX
// on the marketplace <select> change (no page reload, no extra click). Smooth
// cross-fade + stale-response guard. Vanilla; reuses the shared _mp_profile
// partial rendered by /b2b/mp-profile/<mp>/.
(function () {
  var sel = document.getElementById('id_marketplace');
  var panel = document.getElementById('mp-panel');
  if (!sel || !panel) return;
  var urlTpl = panel.getAttribute('data-url-tpl') || '';
  var reqId = 0;
  var FADE = 200;   // keep in sync with .mp-panel transition duration

  // Column-trace micro-interaction — delegated so it survives innerHTML swaps.
  var hot = [];
  function clearHot() { hot.forEach(function (el) { el.classList.remove('col-hot'); }); hot = []; }
  panel.addEventListener('pointerover', function (e) {
    var cell = e.target && e.target.closest ? e.target.closest('[data-col]') : null;
    if (!cell || !panel.contains(cell)) return;
    var col = cell.getAttribute('data-col');
    if (hot.length && hot[0].getAttribute('data-col') === col) return;
    clearHot();
    hot = Array.prototype.slice.call(panel.querySelectorAll('[data-col="' + col + '"]'));
    hot.forEach(function (el) { el.classList.add('col-hot'); });
  });
  panel.addEventListener('pointerleave', clearHot);

  function delay(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  sel.addEventListener('change', function () {
    if (!urlTpl) return;
    var mp = sel.value;
    var myId = ++reqId;
    var url = urlTpl.replace('__MP__', encodeURIComponent(mp));
    clearHot();
    panel.classList.add('is-swapping');          // fade the current profile out
    var fetchP = fetch(url, {
      headers: { 'X-Requested-With': 'XMLHttpRequest' }, credentials: 'same-origin'
    }).then(function (r) { return r.ok ? r.text() : Promise.reject(r.status); });
    // Cross-fade: swap only once BOTH the fade-out and the response are ready.
    Promise.all([fetchP, delay(FADE)]).then(function (res) {
      if (myId !== reqId) return;                 // a newer selection superseded us
      panel.innerHTML = res[0];
      requestAnimationFrame(function () { panel.classList.remove('is-swapping'); });
    }).catch(function () {
      if (myId !== reqId) return;
      panel.innerHTML = '<div class="mp-empty">Could not load this marketplace’s ' +
        'profile — please try again.</div>';
      panel.classList.remove('is-swapping');
    });
  });
})();
