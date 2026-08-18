/* online_b2b/online_b2b/ship_to.html — page script (separated). Server values via #ship_to-cfg JSON. */
var CFG = JSON.parse(document.getElementById("ship_to-cfg").textContent);
(function () {
  var csrf = document.getElementById('stm-csrf').value;
  var hdrs = { 'X-CSRFToken': csrf, 'X-Requested-With': 'XMLHttpRequest', 'Content-Type': 'application/x-www-form-urlencoded' };
  var tbody = document.getElementById('stm-tbody');
  var msWrap = document.getElementById('stm-party-ms');
  var msBtn = document.getElementById('stm-party-btn');
  var msPop = document.getElementById('stm-party-pop');
  var msLbl = document.getElementById('stm-party-lbl');
  var msList = document.getElementById('stm-party-list');
  function selectedParties() {
    return Array.prototype.slice.call(msList.querySelectorAll('input:checked')).map(function (c) { return c.value; });
  }
  var input = document.getElementById('stmq');
  var shown = document.getElementById('stm-shown');
  var spin = document.getElementById('stmspin');
  var searchUrl = CFG["b2b_ship_to_search"];
  var addUrl = CFG["b2b_ship_to_add"];
  var exportBase = CFG["b2b_ship_to_export"];
  var exportLink = document.getElementById('stm-export');
  var FIELDS = ['party', 'del_location', 'cust_no', 'ship_to', 'city', 'postcode'];

  function syncExport() {
    var p = new URLSearchParams();
    var _sp = selectedParties(); if (_sp.length) p.set('party', _sp.join(','));
    if (input.value.trim()) p.set('q', input.value.trim());
    var qs = p.toString();
    if (exportLink) exportLink.href = exportBase + (qs ? '?' + qs : '');
  }

  function reload() {
    var p = new URLSearchParams();
    var _sp = selectedParties(); if (_sp.length) p.set('party', _sp.join(','));
    if (input.value.trim()) p.set('q', input.value.trim());
    syncExport();
    if (spin) spin.hidden = false;
    fetch(searchUrl + '?' + p.toString(), { headers: { 'X-Requested-With': 'XMLHttpRequest' } })
      .then(function (r) { return r.text(); })
      .then(function (html) { tbody.innerHTML = html; if (spin) spin.hidden = true; bind(); });
  }
  var t = null;
  input.addEventListener('input', function () { clearTimeout(t); t = setTimeout(reload, 180); });
  // ── multi-select party filter ──────────────────────────────────────────────
  function updatePartyLabel() {
    var sel = selectedParties();
    msLbl.textContent = sel.length === 0 ? 'All parties'
      : (sel.length === 1 ? sel[0] : sel.length + ' parties');
  }
  function openMs(open) {
    msPop.hidden = !open;
    msWrap.classList.toggle('open', open);
    msBtn.setAttribute('aria-expanded', open ? 'true' : 'false');
    if (open) {
      var f = document.getElementById('stm-party-filter');
      if (f) { f.value = ''; msList.querySelectorAll('.ms-item').forEach(function (it) { it.classList.remove('hide'); }); f.focus(); }
    }
  }
  msBtn.addEventListener('click', function (e) { e.stopPropagation(); openMs(msPop.hidden); });
  msList.addEventListener('change', function () { updatePartyLabel(); reload(); });
  document.getElementById('stm-party-all').addEventListener('click', function () {
    msList.querySelectorAll('input[type=checkbox]').forEach(function (c) { c.checked = true; });
    updatePartyLabel(); reload();
  });
  document.getElementById('stm-party-none').addEventListener('click', function () {
    msList.querySelectorAll('input[type=checkbox]').forEach(function (c) { c.checked = false; });
    updatePartyLabel(); reload();
  });
  var msFilter = document.getElementById('stm-party-filter');
  if (msFilter) msFilter.addEventListener('input', function () {
    var q = this.value.trim().toLowerCase();
    msList.querySelectorAll('.ms-item').forEach(function (it) {
      it.classList.toggle('hide', !!q && it.textContent.trim().toLowerCase().indexOf(q) < 0);
    });
  });
  document.addEventListener('click', function (e) { if (!msWrap.contains(e.target)) openMs(false); });
  updatePartyLabel();
  syncExport();   // reflect any pre-set party/search on first load

  // add panel
  var addBtn = document.getElementById('stm-addbtn'), panel = document.getElementById('stm-addpanel');
  addBtn.addEventListener('click', function () { panel.classList.toggle('open'); });
  document.getElementById('stm-addcancel').addEventListener('click', function () { panel.classList.remove('open'); });
  document.getElementById('stm-addsave').addEventListener('click', function () {
    var body = new URLSearchParams();
    panel.querySelectorAll('input[name]').forEach(function (i) { body.set(i.name, i.value); });
    var msg = document.getElementById('stm-addmsg');
    fetch(addUrl, { method: 'POST', headers: hdrs, body: body.toString() })
      .then(function (r) { return r.json(); })
      .then(function (j) {
        if (!j.ok) { msg.textContent = j.error || 'failed'; msg.className = 'stm-msg err'; return; }
        msg.textContent = '✓ added'; msg.className = 'stm-msg ok';
        panel.querySelectorAll('input[name]').forEach(function (i) { i.value = ''; });
        panel.classList.remove('open'); reload();
      });
  });

  // ── personalization: add / remove a custom column ──────────────────────────
  var fieldBtn = document.getElementById('stm-fieldbtn');
  var fieldPanel = document.getElementById('stm-fieldpanel');
  var fieldName = document.getElementById('stm-fieldname');
  var fieldMsg = document.getElementById('stm-fieldmsg');
  var fieldAddUrl = CFG["b2b_ship_to_field_add"];
  var fieldDelUrl = CFG["b2b_ship_to_field_delete"];
  if (fieldBtn) fieldBtn.addEventListener('click', function () {
    fieldPanel.classList.toggle('open');
    if (fieldPanel.classList.contains('open')) fieldName.focus();
  });
  document.getElementById('stm-fieldcancel').addEventListener('click', function () {
    fieldPanel.classList.remove('open'); fieldMsg.textContent = '';
  });
  function saveField() {
    var label = (fieldName.value || '').trim();
    if (!label) { fieldMsg.textContent = 'Enter a field name.'; fieldMsg.className = 'stm-msg err'; return; }
    var body = new URLSearchParams(); body.set('label', label);
    fetch(fieldAddUrl, { method: 'POST', headers: hdrs, body: body.toString() })
      .then(function (r) { return r.json(); })
      .then(function (j) {
        if (!j.ok) { fieldMsg.textContent = j.error || 'failed'; fieldMsg.className = 'stm-msg err'; return; }
        fieldMsg.textContent = '✓ added — reloading…'; fieldMsg.className = 'stm-msg ok';
        location.reload();
      });
  }
  document.getElementById('stm-fieldsave').addEventListener('click', saveField);
  fieldName.addEventListener('keydown', function (e) { if (e.key === 'Enter') { e.preventDefault(); saveField(); } });
  document.querySelectorAll('.cf-del').forEach(function (b) {   // thead ✕, static
    b.addEventListener('click', function () {
      var name = b.getAttribute('data-cf');
      if (!confirm('Remove the "' + name + '" column? Saved values are kept and restored if you re-add it.')) return;
      var body = new URLSearchParams(); body.set('name', name);
      fetch(fieldDelUrl, { method: 'POST', headers: hdrs, body: body.toString() })
        .then(function (r) { return r.json(); })
        .then(function (j) { if (j.ok) location.reload(); });
    });
  });

  function bind() {
    tbody.querySelectorAll('.stm-del').forEach(function (b) {
      b.addEventListener('click', function () {
        var tr = b.closest('tr'); if (!confirm('Delete this mapping row?')) return;
        fetch('/b2b/ship-to/' + tr.getAttribute('data-id') + '/delete/', { method: 'POST', headers: hdrs })
          .then(function (r) { return r.json(); })
          .then(function (j) { if (j.ok) { tr.remove(); } });
      });
    });
    tbody.querySelectorAll('.stm-edit').forEach(function (b) {
      b.addEventListener('click', function () { editRow(b.closest('tr')); });
    });
  }
  function editRow(tr) {
    if (tr.classList.contains('editing')) return;
    tr.classList.add('editing');
    FIELDS.forEach(function (f) {
      var td = tr.querySelector('.' + f); if (!td) return;
      var v = td.getAttribute('data-v') || td.textContent.trim();
      if (v === '—') v = '';
      td.innerHTML = '<input class="stm-in" data-f="' + f + '" value="' + v.replace(/"/g, '&quot;') + '">';
    });
    tr.querySelectorAll('td.cf').forEach(function (td) {   // custom columns
      var name = td.getAttribute('data-cf');
      var v = td.getAttribute('data-v') || td.textContent.trim();
      if (v === '—') v = '';
      td.innerHTML = '<input class="stm-in" data-f="cf_' + name + '" value="' + v.replace(/"/g, '&quot;') + '">';
    });
    var act = tr.querySelector('.stm-actions');
    act.innerHTML = '<button type="button" class="stm-save">save</button> <button type="button" class="stm-cancel">✕</button>';
    act.querySelector('.stm-cancel').addEventListener('click', reload);
    act.querySelector('.stm-save').addEventListener('click', function () {
      var body = new URLSearchParams();
      tr.querySelectorAll('.stm-in').forEach(function (i) { body.set(i.getAttribute('data-f'), i.value); });
      fetch('/b2b/ship-to/' + tr.getAttribute('data-id') + '/edit/', { method: 'POST', headers: hdrs, body: body.toString() })
        .then(function (r) { return r.json(); })
        .then(function (j) { if (j.ok) reload(); else alert(j.error || 'failed'); });
    });
  }
  bind();
})();
