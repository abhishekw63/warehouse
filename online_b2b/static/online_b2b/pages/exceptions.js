/* online_b2b/exceptions.html — page script (separated from template). */
(function () {
  var xc = document.querySelector('.xc');
  var csrf = B2B.csrf();
  function post(url, body) {
    return fetch(url, { method: 'POST', credentials: 'same-origin',
      headers: { 'X-CSRFToken': csrf, 'X-Requested-With': 'XMLHttpRequest',
                 'Content-Type': 'application/x-www-form-urlencoded' },
      body: body }).then(function (r) { return r.json(); });
  }
  // Lens toggle (fade handled by CSS animation on .xview)
  var lensBtns = xc.querySelectorAll('.xlens button'), views = xc.querySelectorAll('.xview');
  lensBtns.forEach(function (b) { b.addEventListener('click', function () {
    lensBtns.forEach(function (x) { x.classList.toggle('on', x === b); });
    views.forEach(function (p) { p.hidden = p.dataset.view !== b.dataset.lens; });
  }); });
  // Collapse / expand a marketplace card by clicking its header (chevron rotates).
  // Ignore clicks on the action buttons so edit/delete still work.
  xc.querySelectorAll('.mpc-h').forEach(function (h) {
    h.addEventListener('click', function (ev) {
      if (ev.target.closest('.ibtn')) return;
      h.parentElement.classList.toggle('collapsed');
    });
  });
  // Add
  var form = document.getElementById('ex-form'), addBtn = document.getElementById('ex-add-btn');
  form.addEventListener('submit', function () {
    var fd = new FormData(form);
    if (!fd.get('marketplace') || !fd.get('source_code')) { alert('Marketplace and SKU are required.'); return; }
    addBtn.disabled = true; addBtn.textContent = 'Adding…';
    post(xc.dataset.addUrl, new URLSearchParams(fd).toString())
      .then(function (j) { if (j.ok) location.reload();
        else { addBtn.disabled = false; addBtn.textContent = 'Add exception'; alert(j.error || 'Could not add.'); } })
      .catch(function () { addBtn.disabled = false; addBtn.textContent = 'Add exception'; alert('Network error.'); });
  });
  // Delete + Edit (delegated)
  xc.addEventListener('click', function (ev) {
    var del = ev.target.closest('.ibtn.del');
    if (del && !del.disabled) {
      if (!confirm('Delete this exception?')) return;
      del.disabled = true;
      post('/b2b/exceptions/' + del.dataset.id + '/delete/', '')
        .then(function (j) { if (j.ok) location.reload(); else { del.disabled = false; alert(j.error || 'Could not delete.'); } })
        .catch(function () { del.disabled = false; alert('Network error.'); });
      return;
    }
    var edit = ev.target.closest('.ibtn.edit');
    if (edit) { openEdit(edit.closest('tr')); }
  });
  function esc(s) { var d = document.createElement('div'); d.textContent = s || ''; return d.innerHTML.replace(/"/g, '&quot;'); }
  function openEdit(tr) {
    if (tr.nextSibling && tr.nextSibling.classList && tr.nextSibling.classList.contains('edrow')) return;
    var d = tr.dataset, cols = tr.children.length;
    var er = document.createElement('tr'); er.className = 'edrow';
    er.innerHTML = '<td colspan="' + cols + '"><form class="edform">' +
      '<div><label>Override MRP</label><input name="override_mrp" value="' + esc(d.mrp) + '"></div>' +
      '<div><label>Override margin %</label><input name="override_margin" value="' + esc(d.margin) + '"></div>' +
      '<div><label>Maps to (EAN)</label><input name="maps_to" value="' + esc(d.maps) + '"></div>' +
      '<div style="grid-column:span 2;"><label>Note</label><input name="note" value="' + esc(d.note) + '"></div>' +
      '<div style="display:flex;gap:8px;"><button type="submit" class="save">Save</button>' +
      '<button type="button" class="cancel">Cancel</button></div></form></td>';
    tr.after(er);
    var f = er.querySelector('.edform');
    f.querySelector('.cancel').addEventListener('click', function () { er.remove(); });
    f.addEventListener('submit', function (e) { e.preventDefault();
      var body = new URLSearchParams(new FormData(f)).toString();
      post('/b2b/exceptions/' + d.id + '/update/', body)
        .then(function (j) { if (j.ok) location.reload(); else alert(j.error || 'Could not save.'); })
        .catch(function () { alert('Network error.'); });
    });
  }
})();
