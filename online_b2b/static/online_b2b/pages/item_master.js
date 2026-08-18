/* online_b2b/online_b2b/item_master.html — page script (separated). Server values via #item_master-cfg JSON. */
var CFG = JSON.parse(document.getElementById("item_master-cfg").textContent);
(function () {
  // ── Add-item panel toggle ──
  var addBtn = document.getElementById('im-addbtn');
  var panel  = document.getElementById('im-addpanel');
  var cancel = document.getElementById('im-addcancel');
  if (addBtn && panel) {
    addBtn.addEventListener('click', function () {
      panel.classList.toggle('open');
      if (panel.classList.contains('open')) {
        var f = panel.querySelector('input[name=item_no]'); if (f) f.focus();
      }
    });
  }
  if (cancel && panel) cancel.addEventListener('click', function () { panel.classList.remove('open'); });

  // ── Live (as-you-type) search ──
  var input = document.getElementById('imq');
  var tbody = document.getElementById('im-tbody');
  var shown = document.getElementById('im-shown');
  var spin  = document.getElementById('imspin');
  if (!input || !tbody) return;
  var URL = CFG["b2b_item_master_search"];
  var EXPORT_URL = CFG["b2b_item_master_export"];
  var exportBtn = document.getElementById('im-export');
  var timer = null, ctl = null;

  // keep the Export link pointed at the CURRENT search term (all rows if blank)
  function syncExport() {
    if (!exportBtn) return;
    var q = input.value.trim();
    exportBtn.href = EXPORT_URL + (q ? ('?q=' + encodeURIComponent(q)) : '');
  }

  function esc(s) { var d = document.createElement('div'); d.textContent = (s == null ? '' : s); return d.innerHTML; }
  function render(data) {
    if (!data.rows.length) {
      tbody.innerHTML = '<tr><td colspan="7" class="muted">No items match.</td></tr>';
    } else {
      tbody.innerHTML = data.rows.map(function (r) {
        var win = r.mrp_start ? (esc(r.mrp_start) + ' → ' + esc(r.mrp_end)) : '—';
        var mrp = (r.mrp == null) ? '—' : r.mrp;
        return '<tr><td class="mono">' + esc(r.item_no) + '</td>' +
          '<td class="mono">' + (esc(r.ean) || '—') + '</td>' +
          '<td class="desc" title="' + esc(r.description) + '">' + esc(r.description) + '</td>' +
          '<td class="r">' + mrp + '</td>' +
          '<td>' + (esc(r.gst_code) || '—') + '</td>' +
          '<td class="mono">' + (esc(r.hsn) || '—') + '</td>' +
          '<td class="muted">' + win + '</td></tr>';
      }).join('');
    }
    shown.textContent = 'Showing ' + data.shown + ' of ' + data.total +
      (data.q ? (' matching “' + data.q + '”') : '');
  }
  function go() {
    var q = input.value.trim();
    if (spin) spin.hidden = false;
    if (ctl) ctl.abort();
    ctl = (window.AbortController ? new AbortController() : null);
    fetch(URL + '?q=' + encodeURIComponent(q), {
      headers: { 'X-Requested-With': 'XMLHttpRequest' },
      signal: ctl ? ctl.signal : undefined
    }).then(function (r) { return r.json(); })
      .then(function (data) { render(data); if (spin) spin.hidden = true; })
      .catch(function () { if (spin) spin.hidden = true; });
  }
  input.addEventListener('input', function () { clearTimeout(timer); timer = setTimeout(go, 160); syncExport(); });
})();
