/* Record Verification (/b2b/record-verify) — page script.
   Separated out of the template. Row selection reuses the shared B2B.checkAll();
   the AJAX confirm (record only the ticked POs, no page refresh) is page-specific. */
(function () {
  var rows = [].slice.call(document.querySelectorAll('.rv-chk'));
  var selc = document.querySelector('.rv-selc');
  var confirmBtn = document.querySelector('.rv-confirm form button');

  // Shared select-all + live count; onChange also dims rows that won't be recorded.
  B2B.checkAll({
    items: rows,
    master: '#rvAll',
    onChange: function (n) {
      if (selc) selc.textContent = n;
      if (confirmBtn) confirmBtn.disabled = (n === 0);
      rows.forEach(function (c) {
        var tr = c.closest('tr'); if (tr) tr.classList.toggle('rv-unpicked', !c.checked);
      });
    }
  });

  // ── AJAX confirm — records ONLY the ticked POs, no page refresh ────────────
  var form = document.querySelector('.rv-confirm form');
  if (!form) return;
  form.addEventListener('submit', function (e) {
    e.preventDefault();
    var btn = form.querySelector('button'), bar = form.closest('.rv-confirm');
    var picked = rows.filter(function (c) { return c.checked; });
    if (!picked.length) {
      if (window.B2B && B2B.toast) B2B.toast('Tick at least one PO to record.', { type: 'error' });
      return;
    }
    var orig = btn.innerHTML;
    btn.disabled = true; btn.innerHTML = '<span class="rv-spin"></span>Recording…';
    var fd = new FormData(form);
    picked.forEach(function (c) { fd.append('push_po', c.value); });
    fetch(form.action, { method: 'POST', body: fd,
      headers: { 'X-Requested-With': 'XMLHttpRequest' }, credentials: 'same-origin' })
      .then(function (r) { return r.json(); })
      .then(function (j) {
        if (j && j.ok) {
          bar.classList.add('done');
          bar.innerHTML = '<span class="rv-c-text">✓ <b>Verification recorded</b> — ' + j.confirmed + ' PO(s) logged.</span>';
          if (window.B2B && B2B.toast) B2B.toast(j.message, { type: 'success', title: 'Recorded' });
          if (window.confetti) { try { confetti({ particleCount: 70, spread: 68, origin: { y: .55 } }); } catch (e) {} }
        } else {
          btn.disabled = false; btn.innerHTML = orig;
          if (window.B2B && B2B.toast) B2B.toast((j && j.error) || 'Could not record.', { type: 'error' });
        }
      })
      .catch(function () {
        btn.disabled = false; btn.innerHTML = orig;
        if (window.B2B && B2B.toast) B2B.toast('Network error — nothing recorded.', { type: 'error' });
      });
  });
})();

/* ── Tabs — same behaviour as the review page (show one pane, hide the rest) ── */
(function () {
  var tabs = [].slice.call(document.querySelectorAll('.rv-tabs .tab'));
  if (!tabs.length) return;
  tabs.forEach(function (t) {
    t.addEventListener('click', function () {
      tabs.forEach(function (x) { x.classList.remove('on'); });
      t.classList.add('on');
      var name = t.getAttribute('data-tab');
      document.querySelectorAll('.tabpane').forEach(function (p) {
        p.style.display = (p.getAttribute('data-pane') === name) ? '' : 'none';
      });
    });
  });
})();

/* ── Per-PO drill-down: click an Orders/Externals row (not its checkbox) to open
      its line items and pinpoint the mismatch. Delegated → works after re-render.
      Bound ONCE on document: under the persistent shell-nav this page script re-runs
      on every re-visit, and removing the old <script> tag does NOT detach a listener
      it already added to document. Without this guard a second listener accumulates
      and each click opens-then-closes the row (net nothing → "row won't open"). ── */
(function () {
  if (window.__rvDrillBound) return;
  window.__rvDrillBound = true;
  document.addEventListener('click', function (e) {
    if (!e.target.closest) return;
    if (e.target.closest('.rv-chk-col')) return;         // let the checkbox toggle
    var row = e.target.closest('.rv-orow');
    if (!row) return;
    var detail = row.nextElementSibling;
    if (!detail || !detail.classList.contains('rv-detail')) return;
    var opening = detail.hasAttribute('hidden');
    if (opening) { detail.removeAttribute('hidden'); row.classList.add('rv-open'); }
    else { detail.setAttribute('hidden', ''); row.classList.remove('rv-open'); }
    var caret = row.querySelector('.rv-caret');
    if (caret) caret.textContent = opening ? '▾' : '▸';
  });
})();

/* ── Clear history — wipe the checked-PO log (AJAX) on the history page ── */
(function () {
  var btn = document.querySelector('.rv-clearlog');
  if (!btn) return;
  btn.addEventListener('click', function () {
    var url = btn.getAttribute('data-url'); if (!url) return;
    if (!window.confirm('Clear the entire verification history? Recorded orders are NOT affected.')) return;
    btn.disabled = true; btn.innerHTML = 'Clearing…';
    B2B.postForm(url, {}).then(function (j) {
      if (j && j.ok) {
        if (window.B2B && B2B.toast) B2B.toast(j.message || 'Cleared.', { type: 'info' });
        window.location.href = (j && j.redirect) || '/b2b/record-verify/log/';
      } else {
        btn.disabled = false; btn.innerHTML = '🗑 Clear history';
        if (window.B2B && B2B.toast) B2B.toast((j && j.error) || 'Could not clear.', { type: 'error' });
      }
    }).catch(function () {
      btn.disabled = false; btn.innerHTML = '🗑 Clear history';
      if (window.B2B && B2B.toast) B2B.toast('Network error.', { type: 'error' });
    });
  });
})();

/* ── Discard — delete this check (AJAX), then go to a clean page ── */
(function () {
  var btn = document.querySelector('.rv-discard');
  if (!btn) return;
  btn.addEventListener('click', function () {
    var url = btn.getAttribute('data-url'); if (!url) return;
    var pending = parseInt(btn.getAttribute('data-pending') || '0', 10);
    var q = pending > 0
      ? 'Discard this ENTIRE check? You still have ' + pending + ' import order(s) pending — they will be discarded too. Nothing is recorded.'
      : 'Discard this entire check? It will be removed — nothing is recorded.';
    if (!window.confirm(q)) return;
    btn.disabled = true; btn.innerHTML = 'Discarding…';
    B2B.postForm(url, {}).then(function (j) {
      if (j && j.ok) {
        if (window.B2B && B2B.toast) B2B.toast(j.message || 'Discarded.', { type: 'info' });
        window.location.href = (j && j.redirect) || '/b2b/record-verify/';
      } else {
        btn.disabled = false; btn.innerHTML = '🗑 Discard';
        if (window.B2B && B2B.toast) B2B.toast((j && j.error) || 'Could not discard.', { type: 'error' });
      }
    }).catch(function () {
      btn.disabled = false; btn.innerHTML = '🗑 Discard';
      if (window.B2B && B2B.toast) B2B.toast('Network error.', { type: 'error' });
    });
  });
})();

/* ── Save for Review Later — park the run (AJAX), nothing recorded ── */
(function () {
  var btn = document.querySelector('.rv-savelater');
  if (!btn) return;
  btn.addEventListener('click', function () {
    var url = btn.getAttribute('data-url'); if (!url) return;
    var orig = btn.innerHTML; btn.disabled = true; btn.innerHTML = 'Saving…';
    B2B.postForm(url, {}).then(function (j) {
      if (j && j.ok) {
        btn.innerHTML = '🕒 Saved for later';
        if (window.B2B && B2B.toast) B2B.toast(j.message, { type: 'success', title: 'Saved for later' });
      } else {
        btn.disabled = false; btn.innerHTML = orig;
        if (window.B2B && B2B.toast) B2B.toast((j && j.error) || 'Could not save.', { type: 'error' });
      }
    }).catch(function () {
      btn.disabled = false; btn.innerHTML = orig;
      if (window.B2B && B2B.toast) B2B.toast('Network error — not saved.', { type: 'error' });
    });
  });
})();

/* ── Capture external orders → tracker (separate action) — classify any unknown
      posting group (Segment → Marketplace → MT child), then record into the DB. ── */
(function () {
  var panel = document.getElementById('rvCapture');
  if (!panel) return;
  var btn = document.getElementById('rvCapBtn'), msg = document.getElementById('rvCapMsg');
  var classify = document.getElementById('rvCapClassify');
  var taxEl = document.getElementById('rvTaxonomy');
  var TAX = taxEl ? JSON.parse(taxEl.textContent) : { online: [], offline: [], mt_children: [] };
  var overrides = {}, needsN = 0;
  var selAll = document.getElementById('rvCapAll'),
      selN = document.getElementById('rvCapSel'),
      btnN = document.getElementById('rvCapBtnN');
  function chks() { return [].slice.call(panel.querySelectorAll('.rv-cap-chk')); }
  function selectedPos() { return chks().filter(function (c) { return c.checked; }).map(function (c) { return c.value; }); }

  function opts(list) {
    return '<option value="">—</option>' + list.map(function (o) {
      return '<option value="' + o.value + '">' + o.label + '</option>';
    }).join('');
  }
  function csrf() {
    var el = document.querySelector('[name=csrfmiddlewaretoken]');
    if (el) return el.value;
    var m = document.cookie.match(/csrftoken=([^;]+)/); return m ? m[1] : '';
  }
  function collect() {
    overrides = {};
    if (classify) [].slice.call(classify.querySelectorAll('.rv-cap-cl-row')).forEach(function (row) {
      var seg = row.querySelector('.rv-cap-seg').value, mpEl = row.querySelector('.rv-cap-mp'),
          chEl = row.querySelector('.rv-cap-child');
      var mp = mpEl.value; if (!seg || !mp) return;
      var isMT = (mp === 'MT'), childV = chEl.hidden ? '' : chEl.value; if (isMT && !childV) return;
      var mpText = mpEl.options[mpEl.selectedIndex] ? mpEl.options[mpEl.selectedIndex].text : mp;
      // MT child: the option VALUE is the canonical db_label the tracker keys on
      // (e.g. 'Health & Glow'), NOT its display text ('H&G').
      overrides[row.getAttribute('data-key')] = { posting_group: row.getAttribute('data-pg'),
        segment: seg, marketplace: mp, marketplace_label: isMT ? chEl.value : mpText };
    });
    previewChannels();
    update();
  }
  // Reflect each classified group's chosen channel in the order table's CHANNEL cell
  // (preview the change before pushing); restore the original label when cleared.
  function previewChannels() {
    chks().forEach(function (c) {
      var row = c.closest('tr'); if (!row) return;
      var cell = row.querySelector('.rv-cap-ch'); if (!cell) return;
      var ov = overrides[c.getAttribute('data-classkey')];
      if (ov && ov.marketplace_label) { cell.textContent = ov.marketplace_label; cell.classList.add('rv-cap-ch-new'); }
      else { cell.textContent = cell.getAttribute('data-orig') || cell.textContent; cell.classList.remove('rv-cap-ch-new'); }
    });
  }
  function classifyKeys() {
    return classify ? [].slice.call(classify.querySelectorAll('.rv-cap-cl-row'))
      .map(function (r) { return r.getAttribute('data-key'); }) : [];
  }
  function update() {
    var all = chks();
    // Reflect include/exclude on the table like the review page: ticked = INCLUDED
    // (green), unticked = EXCLUDED (red + struck) — the user sees what will push.
    all.forEach(function (c) {
      var tr = c.closest('tr'); if (!tr) return;
      tr.classList.toggle('rv-cap-excl', !c.checked);
      tr.classList.toggle('rv-cap-inc', c.checked);
    });
    var selChks = all.filter(function (c) { return c.checked; });
    var sel = selChks.length;
    // Only unknown posting groups that STILL have a selected order must be classified;
    // a group whose orders are ALL unticked (excluded) never blocks the push.
    // NB: '' IS a real key — the '(blank)' posting group — so it must NOT be skipped
    // (an empty key falsily skipped would let blank-group orders bypass the gate).
    var selKeys = {};
    selChks.forEach(function (c) { selKeys[c.getAttribute('data-classkey') || ''] = true; });
    var keys = classifyKeys();
    // Safety net: a selected order with a blank key but NO '(blank)' classify row means
    // its group is unmappable (e.g. a pre-class_key saved-later run) — conservatively
    // require EVERY unknown group be placed, so nothing slips through unclassified.
    var orphanBlank = selKeys[''] && keys.indexOf('') === -1;
    var remaining = keys.filter(function (k) { return !overrides[k] && (selKeys[k] || orphanBlank); }).length;
    if (selN) selN.textContent = sel;
    if (btnN) btnN.textContent = sel;
    if (remaining > 0) { btn.disabled = true; if (msg) msg.textContent = 'Place the ' + remaining + " remaining posting group(s) first (excluded ones don't need it)."; return; }
    if (!sel) { btn.disabled = true; if (msg) msg.textContent = 'Tick at least one order to push.'; return; }
    btn.disabled = false; if (msg) msg.textContent = '';
  }
  if (classify) {
    var rowsC = [].slice.call(classify.querySelectorAll('.rv-cap-cl-row'));
    needsN = rowsC.length;
    rowsC.forEach(function (row) {
      var seg = row.querySelector('.rv-cap-seg'), mp = row.querySelector('.rv-cap-mp'),
          child = row.querySelector('.rv-cap-child');
      seg.addEventListener('change', function () {
        var mps = seg.value === 'Offline' ? TAX.offline : (seg.value === 'OnlineB2B' ? TAX.online : []);
        mp.innerHTML = opts(mps); mp.disabled = !mps.length;
        child.hidden = true; child.innerHTML = '<option value="">—</option>'; collect();
      });
      mp.addEventListener('change', function () {
        var mps = seg.value === 'Offline' ? TAX.offline : TAX.online, isMT = false;
        mps.forEach(function (o) { if (o.value === mp.value && o.mt) isMT = true; });
        if (isMT) { child.hidden = false; child.innerHTML = opts(TAX.mt_children); }
        else { child.hidden = true; child.innerHTML = '<option value="">—</option>'; }
        collect();
      });
      child.addEventListener('change', collect);
    });
  }
  // selectable order list — select-all + per-row ticks drive the live count.
  if (selAll) selAll.addEventListener('change', function () {
    chks().forEach(function (c) { c.checked = selAll.checked; }); update();
  });
  chks().forEach(function (c) {
    c.addEventListener('change', function () {
      if (selAll) selAll.checked = chks().every(function (x) { return x.checked; });
      update();
    });
  });
  // ↺ Reset — re-select every order + clear all classifications + restore channels.
  var resetBtn = document.getElementById('rvCapReset');
  if (resetBtn) resetBtn.addEventListener('click', function () {
    chks().forEach(function (c) { c.checked = true; });
    if (selAll) selAll.checked = true;
    if (classify) [].slice.call(classify.querySelectorAll('.rv-cap-cl-row')).forEach(function (row) {
      var seg = row.querySelector('.rv-cap-seg'), mp = row.querySelector('.rv-cap-mp'), child = row.querySelector('.rv-cap-child');
      seg.value = ''; mp.innerHTML = '<option value="">—</option>'; mp.disabled = true;
      child.hidden = true; child.innerHTML = '<option value="">—</option>';
    });
    collect();
    if (window.B2B && B2B.toast) B2B.toast('Reset — all selected, classifications cleared.', { type: 'info' });
  });
  update();

  btn.addEventListener('click', function () {
    var url = btn.getAttribute('data-url'); if (!url) return;
    var picked = selectedPos();
    if (!picked.length) return;
    var orig = btn.innerHTML; btn.disabled = true; btn.innerHTML = 'Pushing…';
    fetch(url, { method: 'POST', credentials: 'same-origin',
      headers: { 'X-Requested-With': 'XMLHttpRequest', 'X-CSRFToken': csrf(), 'Content-Type': 'application/json' },
      body: JSON.stringify({ overrides: overrides, only_pos: picked }) })
      .then(function (r) { return r.json(); }).then(function (j) {
        if (j && j.ok) {
          panel.classList.add('done');
          panel.innerHTML = '<span>✓ <b>Captured ' + j.imported + ' order(s) · ' + j.lines +
            ' line(s)</b> into the tracker (' + j.skipped + ' already present).</span> ' +
            '<a class="btn-ghost" href="' + j.redirect + '">View in Tracker →</a>';
          if (window.B2B && B2B.toast) B2B.toast('Captured ' + j.imported + ' order(s) into the tracker.', { type: 'success', title: 'Captured' });
          if (window.confetti) { try { confetti({ particleCount: 80, spread: 70, origin: { y: .5 } }); } catch (e) {} }
        } else {
          btn.disabled = false; btn.innerHTML = orig;
          if (window.B2B && B2B.toast) B2B.toast((j && j.error) || 'Capture failed.', { type: 'error' });
        }
      }).catch(function () {
        btn.disabled = false; btn.innerHTML = orig;
        if (window.B2B && B2B.toast) B2B.toast('Network error during capture.', { type: 'error' });
      });
  });
})();

/* ── Upload dropzone — show picked filenames + a drag-over state ── */
(function () {
  var drop = document.getElementById('rvDrop');
  if (!drop) return;
  var input = drop.querySelector('.rv-drop-input');
  var out = document.getElementById('rvDropFiles');
  if (input) input.addEventListener('change', function () {
    var names = [].map.call(input.files, function (f) { return f.name; });
    if (out) { out.hidden = !names.length; out.textContent = names.length ? '✓ ' + names.join('   ·   ') : ''; }
  });
  ['dragenter', 'dragover'].forEach(function (ev) {
    drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add('rv-drop-over'); });
  });
  ['dragleave', 'drop'].forEach(function (ev) {
    drop.addEventListener(ev, function () { drop.classList.remove('rv-drop-over'); });
  });
})();

/* ── Import / Verify top-level tabs — switch panes over the SAME uploaded
      Headers+Lines (initial state is set server-side, so no flash on load). ── */
(function () {
  var tabs = [].slice.call(document.querySelectorAll('.rv-toptab'));
  if (!tabs.length) return;
  var panes = [].slice.call(document.querySelectorAll('[data-toppane]'));
  function show(name) {
    tabs.forEach(function (t) {
      var on = t.getAttribute('data-toptab') === name;
      t.classList.toggle('on', on);
      t.setAttribute('aria-selected', on ? 'true' : 'false');
    });
    panes.forEach(function (p) { p.hidden = (p.getAttribute('data-toppane') !== name); });
  }
  tabs.forEach(function (t) {
    t.addEventListener('click', function () { show(t.getAttribute('data-toptab')); });
  });
})();

/* ── Per-tab discard — drop ONLY the import OR the verification (AJAX), leaving the
      other tab intact. Bound once on document so it survives shell-nav re-runs. ── */
(function () {
  if (window.__rvPartDiscardBound) return;
  window.__rvPartDiscardBound = true;
  document.addEventListener('click', function (e) {
    var btn = e.target.closest && e.target.closest('.rv-discard-part');
    if (!btn) return;
    var url = btn.getAttribute('data-url'), part = btn.getAttribute('data-part') || '';
    if (!url) return;
    var label = part === 'import' ? 'import (new orders)' : 'verification';
    if (!window.confirm('Discard the ' + label + ' only? The other tab stays. Nothing is recorded.')) return;
    var orig = btn.innerHTML; btn.disabled = true; btn.innerHTML = 'Discarding…';
    B2B.postForm(url, { part: part }).then(function (j) {
      if (j && j.ok) {
        if (window.B2B && B2B.toast) B2B.toast(j.message || 'Discarded.', { type: 'info' });
        window.location.href = (j && j.redirect) || window.location.href;
      } else {
        btn.disabled = false; btn.innerHTML = orig;
        if (window.B2B && B2B.toast) B2B.toast((j && j.error) || 'Could not discard.', { type: 'error' });
      }
    }).catch(function () {
      btn.disabled = false; btn.innerHTML = orig;
      if (window.B2B && B2B.toast) B2B.toast('Network error.', { type: 'error' });
    });
  });
})();
