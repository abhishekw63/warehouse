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
