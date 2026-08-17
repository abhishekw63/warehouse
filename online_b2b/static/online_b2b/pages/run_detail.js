/* online_b2b/run_detail.html — page script (separated from template). */
(function () {
  var open = document.getElementById('rd-delbtn');
  var modal = document.getElementById('rd-delmodal');
  var cancel = document.getElementById('rd-delcancel');
  var field = document.getElementById('rd-delconfirm');
  var go = document.getElementById('rd-delgo');
  if (!open || !modal) return;
  function show(v) { modal.hidden = !v; if (v) { field.value = ''; go.disabled = true; field.focus(); } }
  open.addEventListener('click', function () { show(true); });
  cancel.addEventListener('click', function () { show(false); });
  modal.addEventListener('click', function (e) { if (e.target === modal) show(false); });
  document.addEventListener('keydown', function (e) { if (e.key === 'Escape' && !modal.hidden) show(false); });
  field.addEventListener('input', function () {
    go.disabled = field.value.trim().toUpperCase() !== 'DELETE';
  });
})();
