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

  // keep the Export link pointed at the CURRENT search term + margin (so the
  // downloaded LR Unit Price / CP match what's on screen).
  function syncExport() {
    if (!exportBtn) return;
    var q = input.value.trim(), p = [];
    if (q) p.push('q=' + encodeURIComponent(q));
    p.push('mult=' + curMult());
    exportBtn.href = EXPORT_URL + '?' + p.join('&');
  }

  function esc(s) { var d = document.createElement('div'); d.textContent = (s == null ? '' : s); return d.innerHTML; }

  // ── LR Unit Price + CP (dynamic margin) ──
  // LR = MRP × margin% ; CP = LR ÷ (1 + GST for the item's GST group). The margin
  // input drives both; default 60%. Rows carry data-mrp + data-gst so we recompute
  // live with zero server round-trips (and re-apply after an AJAX search render).
  function curMult() {
    var el = document.getElementById('im-mult');
    var v = el ? parseFloat(el.value) : 60;
    return (isFinite(v) && v > 0) ? v : 60;
  }
  function fmt2(n) { return (n == null || isNaN(n)) ? '—' : (Math.round(n * 100) / 100).toFixed(2); }
  function recompute() {
    var m = curMult();
    tbody.querySelectorAll('.im-lr').forEach(function (cell) {
      var mrp = parseFloat(cell.getAttribute('data-mrp')) || 0;
      var gst = parseFloat(cell.getAttribute('data-gst')) || 0;
      var lr = mrp * (m / 100);
      cell.textContent = fmt2(lr);
      var cp = cell.nextElementSibling;
      if (cp && cp.classList.contains('im-cp')) cp.textContent = fmt2(lr / (1 + gst));
    });
    syncExport();
  }
  function render(data) {
    if (!data.rows.length) {
      tbody.innerHTML = '<tr><td colspan="9" class="muted">No items match.</td></tr>';
    } else {
      var m = curMult();
      tbody.innerHTML = data.rows.map(function (r) {
        var win = r.mrp_start ? (esc(r.mrp_start) + ' → ' + esc(r.mrp_end)) : '—';
        var mrp = (r.mrp == null) ? '—' : r.mrp;
        var mv = parseFloat(r.mrp) || 0, gr = parseFloat(r.gst_rate) || 0, lr = mv * (m / 100);
        return '<tr><td class="mono">' + esc(r.item_no) + '</td>' +
          '<td class="mono">' + (esc(r.ean) || '—') + '</td>' +
          '<td class="desc" title="' + esc(r.description) + '">' + esc(r.description) + '</td>' +
          '<td class="r">' + mrp + '</td>' +
          '<td class="r im-lr" data-mrp="' + mv + '" data-gst="' + gr + '">' + fmt2(lr) + '</td>' +
          '<td class="r im-cp">' + fmt2(lr / (1 + gr)) + '</td>' +
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

  // Margin multiplier → recompute LR Unit Price + CP live (default 60%).
  var multEl = document.getElementById('im-mult');
  if (multEl) multEl.addEventListener('input', recompute);
  syncExport();   // seed the export link with the current margin
})();
