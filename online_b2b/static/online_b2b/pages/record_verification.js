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
      its line items and pinpoint the mismatch. Delegated → works after re-render. ── */
(function () {
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

/* ── Discard — delete this check (AJAX), then go to a clean page ── */
(function () {
  var btn = document.querySelector('.rv-discard');
  if (!btn) return;
  btn.addEventListener('click', function () {
    var url = btn.getAttribute('data-url'); if (!url) return;
    if (!window.confirm('Discard this check? It will be removed — nothing is recorded.')) return;
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
