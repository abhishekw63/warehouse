/* online_b2b/online_b2b/sku_summary.html — page script (separated). Server values via #sku_summary-cfg JSON. */
var CFG = JSON.parse(document.getElementById("sku_summary-cfg").textContent);
(function () {
  var form = document.getElementById('sku-filters');
  if (!form) return;
  var tbody = document.getElementById('sku-tbody');
  var spin = document.getElementById('sku-spin');
  var shown = document.getElementById('sku-shown');
  var wrap = document.getElementById('sku-wrap');
  var base = CFG["b2b_sku_summary"];
  var linesUrl = CFG["b2b_sku_summary_lines"];
  var timer = null, ctrl = null;

  function load() {
    var p = new URLSearchParams();
    if (form.marketplace.value) p.set('marketplace', form.marketplace.value);
    if (form.from.value) p.set('from', form.from.value);
    if (form.to.value) p.set('to', form.to.value);
    if (form.q.value.trim()) p.set('q', form.q.value.trim());
    if (form.issues.checked) p.set('issues', '1');
    if (ctrl) ctrl.abort(); ctrl = new AbortController();
    if (spin) spin.hidden = false;
    fetch(base + '?' + p.toString(), { headers: { 'X-Requested-With': 'XMLHttpRequest' }, signal: ctrl.signal })
      .then(function (r) { return r.text(); })
      .then(function (html) { tbody.innerHTML = html; if (spin) spin.hidden = true; bindRows(); syncShown(); })
      .catch(function (e) { if (e.name !== 'AbortError' && spin) spin.hidden = true; });
  }
  function syncShown() {
    var n = tbody.querySelectorAll('tr.sku-row').length;
    if (shown) shown.firstChild ? shown.firstChild.nodeValue = 'Showing ' + n + ' SKUs' : 0;
  }
  form.marketplace.addEventListener('change', load);
  form.from.addEventListener('change', load);
  form.to.addEventListener('change', load);
  form.issues.addEventListener('change', load);
  form.q.addEventListener('input', function () { clearTimeout(timer); timer = setTimeout(load, 300); });

  // Qty ↔ Lines toggle (pure CSS via a body-ish class)
  document.querySelectorAll('.sk-tg').forEach(function (b) {
    b.addEventListener('click', function () {
      document.querySelectorAll('.sk-tg').forEach(function (x) { x.classList.remove('on'); });
      b.classList.add('on');
      wrap.classList.toggle('show-lines', b.getAttribute('data-mode') === 'lines');
    });
  });

  // Drill-down: click a SKU row → load its PO lines under it
  function bindRows() {
    tbody.querySelectorAll('tr.sku-row').forEach(function (tr) {
      tr.addEventListener('click', function () {
        var nxt = tr.nextElementSibling;
        if (nxt && nxt.classList.contains('sku-drill')) { nxt.remove(); tr.classList.remove('open'); return; }
        tr.classList.add('open');
        var det = document.createElement('tr'); det.className = 'sku-drill';
        var td = document.createElement('td'); td.colSpan = 12; td.innerHTML = '<div class="muted" style="padding:8px;">loading…</div>';
        det.appendChild(td); tr.parentNode.insertBefore(det, tr.nextElementSibling);
        var p = new URLSearchParams({ item_no: tr.getAttribute('data-item'), ean: tr.getAttribute('data-ean') });
        fetch(linesUrl + '?' + p.toString(), { headers: { 'X-Requested-With': 'XMLHttpRequest' } })
          .then(function (r) { return r.text(); })
          .then(function (html) { td.innerHTML = html; });
      });
    });
  }
  bindRows();
})();
