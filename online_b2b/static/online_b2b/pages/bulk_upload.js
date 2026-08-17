/* online_b2b/bulk_upload.html — page script (separated from template). */
(function () {
  var dz = document.getElementById('dz'); if (!dz) return;
  var input = dz.querySelector('input[type=file]'), main = document.getElementById('dz-main');
  input.addEventListener('change', function () { main.textContent = input.files.length ? input.files[0].name : 'Click to choose the ERP export'; });
  document.getElementById('up-form').addEventListener('submit', function () {
    if (input.files.length) { var b = document.getElementById('up-submit'); b.disabled = true; b.textContent = 'Reading…'; }
  });
})();
