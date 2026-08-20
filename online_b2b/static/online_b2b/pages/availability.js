/* online_b2b/online_b2b/availability.html — page script (separated). Server values via #availability-cfg JSON. */
var CFG = JSON.parse(document.getElementById("availability-cfg").textContent);
(function () {
  var csrf = B2B.csrf();
  var ta = document.getElementById('av-orders'), whSel = document.getElementById('av-wh');
  var btn = document.getElementById('av-check'), statusEl = document.getElementById('av-status');
  var results = document.getElementById('av-results');
  var checkUrl = CFG["b2b_availability_check"];

  // Persist the last input so a refresh doesn't lose it (restored + auto-checked
  // on load below). Local to this browser; nothing sent until Check.
  var LS_KEY = 'b2b_availability_last';
  function saveState() {
    try { localStorage.setItem(LS_KEY, JSON.stringify({ orders: ta.value, warehouse: whSel.value })); } catch (e) {}
  }

  function nf(n) { return (n === null || n === undefined) ? '—' : Number(n).toLocaleString('en-IN'); }
  function money(v) { return v ? '₹' + Number(v).toLocaleString('en-IN', { maximumFractionDigits: 0 }) : '—'; }
  function pctCell(p, has) { return '<td class="num" style="color:' + pctColor(p) + ';font-weight:700;">' + (has === false ? '—' : p + '%') + '</td>'; }
  function pctColor(p) { return p >= 95 ? '#16a34a' : (p >= 70 ? '#f59e0b' : 'var(--red)'); }
  function stCls(s) { return 'st ' + (s || '').replace(/\s+/g, ''); }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function availCell(v) { return '<td class="num' + (v < 0 ? ' av-neg' : '') + '">' + nf(v) + '</td>'; }

  function fillBar(p) {
    return '<div class="av-fill"><div class="av-bar"><span style="width:' + Math.min(100, p) + '%;background:' + pctColor(p) + ';"></span></div>' +
           '<span class="av-fpct" style="color:' + pctColor(p) + ';">' + p + '%</span></div>';
  }

  function render(data) {
    // not found
    var nf_ = document.getElementById('av-nf');
    nf_.innerHTML = data.not_found && data.not_found.length
      ? '<div class="av-nf">⚠️ ' + data.not_found.length + ' order(s) not found in the system: <b>' + data.not_found.map(esc).join(', ') + '</b></div>' : '';
    // KPIs — qty + value
    var s = data.summary;
    var kpis = [
      ['Orders', nf(s.orders)], ['SKUs', nf(s.skus)],
      ['Ordered qty', nf(s.ord_qty)], ['Fillable qty', nf(s.fillable_qty)],
      ['Fill rate (qty)', s.fill_pct + '%']];
    if (s.has_value) kpis.push(
      ['Ordered ₹', money(s.ord_value)], ['Fillable ₹', money(s.fillable_value)],
      ['Fill rate (val)', s.fill_val_pct + '%']);
    kpis.push(['Fully coverable', nf(s.fully) + '/' + nf(s.orders)]);
    document.getElementById('av-kpis').innerHTML = kpis.map(function (k) {
      return '<div class="av-kpi"><div class="n">' + k[1] + '</div><div class="l">' + k[0] + '</div></div>';
    }).join('');
    // inventory as of
    var asof = document.getElementById('av-asof');
    var whmap = s.wh_stock_as_of || {};
    var keys = Object.keys(whmap);
    asof.innerHTML = keys.length
      ? '🕒 Inventory as of ' + keys.map(function (k) { return '<b>' + esc(k) + '</b> · ' + esc(whmap[k]); }).join(' &nbsp;|&nbsp; ')
      : '';
    var orders = data.orders || [];
    // ── Tab 1 · PO-wise fill rate (one row per order — qty AND value) ──
    var poRows = orders.map(function (o) {
      return '<tr><td class="av-po2">' + esc(o.po) + '</td><td>' + esc(o.marketplace) + '</td>' +
        '<td>' + esc(o.wh_short) + (o.overridden ? ' <span class="av-ov">ovr</span>' : '') + '</td>' +
        '<td class="num">' + nf(o.skus) + '</td>' +
        '<td class="num">' + nf(o.ord_qty) + '</td><td class="num">' + nf(o.fillable_qty) + '</td>' +
        '<td class="num">' + nf(o.short_qty) + '</td>' + pctCell(o.fill_pct) +
        '<td class="num">' + money(o.ord_value) + '</td><td class="num">' + money(o.fillable_value) + '</td>' +
        pctCell(o.fill_val_pct, o.has_value) +
        '<td><span class="' + (o.fully ? 'st OK' : 'st SHORT') + '">' + (o.fully ? 'FULL' : 'PARTIAL') + '</span></td></tr>';
    }).join('');
    document.getElementById('pane-powise').innerHTML =
      '<div class="av-lines"><table><thead><tr><th>Order No</th><th>MP</th><th>WH</th><th class="num">SKUs</th>' +
      '<th class="num">Ord Qty</th><th class="num">Fill Qty</th><th class="num">Short</th><th class="num">Fill% Qty</th>' +
      '<th class="num">Ord ₹</th><th class="num">Fill ₹</th><th class="num">Fill% Val</th><th>Status</th></tr></thead><tbody>' +
      (poRows || '<tr><td colspan="12" class="av-empty">No recognised orders.</td></tr>') + '</tbody></table></div>';

    // ── Tab 2 · PO line items (expandable cards, qty + value) ──
    document.getElementById('pane-lines').innerHTML = orders.map(function (o) {
      var whtag = '<span class="av-whtag' + (o.overridden ? ' ov' : '') + '">' + esc(o.wh_short) +
        (o.overridden ? ' (override · auto ' + esc(o.wh_auto_short) + ')' : '') + '</span>';
      var rows = o.lines.map(function (l) {
        return '<tr><td>' + esc(l.item_no) + '</td><td>' + esc(l.ean) + '</td>' +
          '<td class="av-desc" title="' + esc(l.description) + '">' + esc(l.description) + '</td>' +
          '<td class="num">' + nf(l.ordered) + '</td>' + availCell(l.available) +
          '<td class="num">' + nf(l.fillable) + '</td><td class="num">' + nf(l.short) + '</td>' +
          '<td class="num">' + money(l.ordered_value) + '</td><td class="num">' + money(l.fillable_value) + '</td>' +
          '<td><span class="' + stCls(l.status) + '">' + esc(l.status) + '</span></td></tr>';
      }).join('');
      return '<div class="av-order"><div class="av-ohead" data-toggle>' +
        '<span class="av-cx">&#9656;</span>' +
        '<span class="av-po">' + esc(o.po) + '</span><span class="av-mp">' + esc(o.marketplace) + '</span>' + whtag +
        '<span class="av-ospacer"></span>' +
        '<div class="av-stats">' +
        '<span class="av-stat"><i>Ord</i><b>' + nf(o.ord_qty) + '</b></span>' +
        '<span class="av-stat"><i>Fill</i><b class="pos">' + nf(o.fillable_qty) + '</b></span>' +
        '<span class="av-stat"><i>Short</i><b' + (o.short_qty > 0 ? ' class="neg"' : '') + '>' + nf(o.short_qty) + '</b></span>' +
        '<span class="av-stat"><i>SKU</i><b>' + nf(o.skus) + '</b></span>' +
        (o.has_value ? '<span class="av-stat"><i>Fill ₹</i><b>' + money(o.fillable_value) + '</b></span>' : '') +
        '</div>' + fillBar(o.fill_pct) + '</div>' +
        '<div class="av-lines" style="display:none"><table><thead><tr><th>Item</th><th>EAN</th><th>Description</th>' +
        '<th class="num">Ordered</th><th class="num">Available</th><th class="num">Fillable</th><th class="num">Short</th>' +
        '<th class="num">Ord ₹</th><th class="num">Fill ₹</th><th>Status</th></tr></thead>' +
        '<tbody>' + rows + '</tbody></table></div></div>';
    }).join('') || '<div class="av-empty">No recognised orders.</div>';

    // ── Tab 3 · By SKU (aggregated, qty + value) ──
    var skuRows = (data.skus || []).map(function (k) {
      return '<tr class="sku-row" data-wh="' + esc(k.wh) + '" data-item="' + esc(k.item_no) + '" title="Click to see which bins hold this item">' +
        '<td><span class="sku-cx">&#9656;</span> ' + esc(k.item_no) + '</td><td>' + esc(k.ean) + '</td>' +
        '<td class="av-desc" title="' + esc(k.description) + '">' + esc(k.description) + '</td>' +
        '<td>' + esc(k.wh_short) + '</td><td class="num">' + nf(k.pos) + '</td>' +
        '<td class="num">' + nf(k.ordered) + '</td>' + availCell(k.available) +
        '<td class="num">' + nf(k.fillable) + '</td><td class="num">' + nf(k.short) + '</td>' + pctCell(k.fill_pct) +
        '<td class="num">' + money(k.ordered_value) + '</td><td class="num">' + money(k.fillable_value) + '</td>' +
        pctCell(k.fill_val_pct, (k.ordered_value || 0) > 0) +
        '<td><span class="' + stCls(k.status) + '">' + esc(k.status) + '</span></td></tr>';
    }).join('');
    document.getElementById('pane-sku').innerHTML =
      '<div class="av-lines"><table><thead><tr><th>Item</th><th>EAN</th><th>Description</th><th>WH</th>' +
      '<th class="num">POs</th><th class="num">Ordered</th><th class="num">Available</th><th class="num">Fillable</th>' +
      '<th class="num">Short</th><th class="num">Fill% Qty</th><th class="num">Ord ₹</th><th class="num">Fill ₹</th>' +
      '<th class="num">Fill% Val</th><th>Status</th></tr></thead><tbody>' +
      (skuRows || '<tr><td colspan="14" class="av-empty">No SKUs.</td></tr>') + '</tbody></table></div>';
    results.hidden = false;
    // collapse/expand line tables — collapsed by DEFAULT (hidden in markup).
    document.querySelectorAll('#pane-lines [data-toggle]').forEach(function (h) {
      h.addEventListener('click', function () {
        var t = h.nextElementSibling, show = t.style.display === 'none';
        t.style.display = show ? '' : 'none';
        h.parentNode.classList.toggle('open', show);
      });
    });
  }

  // ── Best-warehouse comparison (overall + PO-wise + SKU-wise, all 3 WHs) ──
  function scCell(p, best) {
    return '<td class="num sc-cell' + (best ? ' sc-best' : '') + '" style="color:' + pctColor(p) +
      ';font-weight:700;">' + p + '%' + (best ? ' <span class="sc-tick">✓</span>' : '') + '</td>';
  }
  function renderScenarios(sc) {
    var pane = document.getElementById('pane-scenario');
    if (!sc || !sc.ok) {
      pane.innerHTML = '<div class="av-empty">' + esc((sc && sc.error) || 'Comparison unavailable.') + '</div>';
      return;
    }
    var whs = sc.warehouses || [];
    var ov = sc.overall || [];
    var best = ov.filter(function (o) { return o.best; })[0] || ov[0] || {};
    // recommendation banner
    var banner = '<div class="sc-rec"><span class="sc-rec-ic">🏆</span><div class="sc-rec-txt">' +
      'Ship this batch from <b>' + esc(best.wh_short) + '</b> — fills <b>' + best.fill_pct + '%</b> of ordered qty' +
      (best.fill_val_pct ? ' <span class="sc-rec-sub">(' + best.fill_val_pct + '% by value)</span>' : '') +
      ', the highest of the three warehouses.' +
      '</div><span class="sc-rec-tot">' + nf(sc.total_qty) + ' units · ' + nf(sc.n_skus) + ' SKUs · ' + nf(sc.n_orders) + ' orders</span></div>';
    // overall cards — one per WH
    var cards = ov.map(function (o) {
      return '<div class="sc-card' + (o.best ? ' sc-card-best' : '') + '">' +
        (o.best ? '<span class="sc-badge">BEST</span>' : '') +
        '<div class="sc-card-wh">' + esc(o.wh_short) + '</div>' +
        '<div class="sc-card-main" style="color:' + pctColor(o.fill_pct) + '">' + o.fill_pct + '%</div>' +
        '<div class="sc-card-sub">qty fill · ' + o.fill_val_pct + '% by value</div>' +
        '<div class="sc-card-row"><span>Fillable</span><b class="pos">' + nf(o.fillable_qty) + '</b></div>' +
        '<div class="sc-card-row"><span>Short</span><b' + (o.short_qty > 0 ? ' class="neg"' : '') + '>' + nf(o.short_qty) + '</b></div>' +
        '<div class="sc-card-row"><span>OOS SKUs</span><b' + (o.oos_skus > 0 ? ' class="neg"' : '') + '>' + nf(o.oos_skus) + ' / ' + nf(o.skus) + '</b></div>' +
        (o.as_of ? '<div class="sc-card-asof">🕒 ' + esc(o.as_of) + '</div>' : '') +
        '</div>';
    }).join('');
    // PO-wise matrix — fill% of each PO in each WH. Each WH cell is a live control
    // for editors: click it to SHIP that PO from that warehouse (persisted).
    var canWrite = CFG["can_write"];
    var whHead = whs.map(function (w) { return '<th class="num">' + esc(w) + '</th>'; }).join('');
    var poRows = (sc.po_wise || []).map(function (p) {
      var cells = whs.map(function (w) {
        var dd = p.by_wh[w] || {}, isBest = p.best_wh === w, isCur = p.cur_wh === w;
        var shiftable = canWrite && !isCur;
        var cls = 'num sc-cell' + (isBest ? ' sc-best' : '') + (isCur ? ' sc-cur' : '') + (shiftable ? ' sc-shiftable' : '');
        var attrs = shiftable ? ' data-shift="' + esc(w) + '" data-po="' + esc(p.po) + '" data-cur="' + esc(p.cur_code) + '" title="Ship ' + esc(p.po) + ' from ' + esc(w) + ' → ' + (dd.fill_pct || 0) + '% fill"' : '';
        return '<td class="' + cls + '"' + attrs + ' style="color:' + pctColor(dd.fill_pct || 0) + ';font-weight:700;">' +
          (dd.fill_pct || 0) + '%' + (isCur ? ' <span class="sc-nowtag">now</span>' : '') + (isBest ? ' <span class="sc-tick">✓</span>' : '') + '</td>';
      }).join('');
      var curCell = '<td class="sc-curcell"><b>' + esc(p.cur_wh) + '</b>' +
        (p.shifted ? '<span class="sc-shifttag" title="shifted from ' + esc(p.orig_wh) + (p.shifted_by ? ' by ' + esc(p.shifted_by) : '') + (p.shifted_at ? ' · ' + esc(p.shifted_at) : '') + '">was ' + esc(p.orig_wh) + '</span>' : '') + '</td>';
      var act = '<td class="sc-actcell">';
      if (canWrite) {
        if (p.can_improve) act += '<button type="button" class="sc-shiftbtn" data-shift="' + esc(p.best_wh) + '" data-po="' + esc(p.po) + '" data-cur="' + esc(p.cur_code) + '">Shift → ' + esc(p.best_wh) + '</button>';
        else if (!p.shifted) act += '<span class="sc-optimal" title="already on its best-fill warehouse">✓ optimal</span>';
        if (p.shifted) act += ' <button type="button" class="sc-revertbtn" data-revert="1" data-po="' + esc(p.po) + '" title="Revert to auto-mapped warehouse">↩ Revert</button>';
      }
      act += '</td>';
      return '<tr><td class="av-po2">' + esc(p.po) + '</td><td>' + esc(p.mp) + '</td>' +
        '<td class="num">' + nf(p.ord_qty) + '</td>' + curCell + cells +
        '<td><span class="sc-bestpill">' + esc(p.best_wh) + '</span></td>' + act + '</tr>';
    }).join('');
    var poTable = '<h3 class="sc-h">PO-wise — fill rate if shipped from each warehouse' +
      (canWrite ? ' <span class="sc-hhint">· click any % (or “Shift”) to reassign an order</span>' : '') + '</h3>' +
      '<div class="av-lines"><table><thead><tr><th>Order No</th><th>MP</th><th class="num">Ord Qty</th><th>Current WH</th>' +
      whHead + '<th>Best</th><th></th></tr></thead><tbody>' +
      (poRows || '<tr><td colspan="' + (6 + whs.length) + '" class="av-empty">No orders.</td></tr>') +
      '</tbody></table></div>';
    // SKU-wise matrix — available / fillable in each WH
    var skuRows = (sc.sku_wise || []).map(function (k) {
      var cells = whs.map(function (w) {
        var d = k.by_wh[w] || {}, isBest = k.best_wh === w;
        var cls = 'num sc-cell' + (isBest ? ' sc-best' : '') + (d.oos ? ' sc-oos' : '');
        return '<td class="' + cls + '">' + (d.oos ? '<span class="sc-oostag">OOS</span>' : nf(d.fillable) + ' <i class="sc-av">/ ' + nf(d.available) + '</i>') + '</td>';
      }).join('');
      return '<tr><td class="mono">' + esc(k.item_no) + '</td>' +
        '<td class="av-desc" title="' + esc(k.description) + '">' + esc(k.description) + '</td>' +
        '<td class="num">' + nf(k.ordered) + '</td><td class="num">' + nf(k.pos) + '</td>' + cells +
        '<td><span class="sc-bestpill">' + esc(k.best_wh) + '</span></td></tr>';
    }).join('');
    var skuTable = '<h3 class="sc-h">SKU-wise — fillable <i class="sc-av">/ available</i> in each warehouse</h3>' +
      '<div class="av-lines"><table><thead><tr><th>Item</th><th>Description</th><th class="num">Ordered</th><th class="num">POs</th>' +
      whHead + '<th>Best</th></tr></thead><tbody>' +
      (skuRows || '<tr><td colspan="' + (5 + whs.length) + '" class="av-empty">No SKUs.</td></tr>') +
      '</tbody></table></div>';
    pane.innerHTML = banner + '<div class="sc-cards">' + cards + '</div>' + poTable + skuTable;
  }

  var scenUrl = CFG["b2b_availability_scenarios"];
  function runScenarios(orders) {
    var pane = document.getElementById('pane-scenario');
    pane.innerHTML = '<div class="av-empty">Comparing warehouses…</div>';
    var body = new URLSearchParams(); body.set('orders', orders);
    B2B.postForm(scenUrl, body)
      .then(renderScenarios)
      .catch(function () { pane.innerHTML = '<div class="av-empty">Comparison failed — retry.</div>'; });
  }

  // Shift a PO to another warehouse (persisted) — or revert to auto. Editor-only
  // (the button/cells only render for can_write; the server re-checks the role).
  var shiftUrl = CFG["b2b_availability_shift"];
  function doShift(po, wh, cur, revert) {
    if (revert) {
      if (!confirm('Revert ' + po + ' to its auto-mapped warehouse?')) return;
    } else if (!confirm('Ship order ' + po + ' from ' + wh + '?\nThis updates its fulfilment warehouse across availability and the inventory fill-rate.')) {
      return;
    }
    var body = new URLSearchParams(); body.set('po', po);
    if (revert) { body.set('action', 'revert'); }
    else { body.set('warehouse', wh); if (cur) body.set('orig_warehouse', cur); }
    B2B.postForm(shiftUrl, body)
      .then(function (j) {
        if (!j.ok) { if (window.B2B && B2B.toast) B2B.toast(j.error || 'Shift failed.', { type: 'error' }); return; }
        if (window.B2B && B2B.toast) B2B.toast(revert ? (po + ' reverted to auto WH') : (po + ' now ships from ' + (j.wh_short || wh)), { type: 'success', title: 'Warehouse updated' });
        run();   // refresh every tab so fill-rate + checks reflect the new WH
      })
      .catch(function () { if (window.B2B && B2B.toast) B2B.toast('Network error — retry.', { type: 'error' }); });
  }
  document.getElementById('pane-scenario').addEventListener('click', function (e) {
    var s = e.target.closest('[data-shift]'), r = e.target.closest('[data-revert]');
    if (s) doShift(s.getAttribute('data-po'), s.getAttribute('data-shift'), s.getAttribute('data-cur'), false);
    else if (r) doShift(r.getAttribute('data-po'), null, null, true);
  });

  function run() {
    var orders = ta.value.trim();
    if (!orders) { statusEl.textContent = 'Paste at least one order number.'; return; }
    saveState();
    btn.disabled = true; statusEl.textContent = 'Checking…';
    var body = new URLSearchParams(); body.set('orders', orders); body.set('warehouse', whSel.value);
    B2B.postForm(checkUrl, body)
      .then(function (j) {
        btn.disabled = false;
        if (!j.ok) { statusEl.textContent = j.error || 'Failed.'; return; }
        statusEl.textContent = '✓ ' + j.summary.orders + ' order(s) checked';
        render(j);
        runScenarios(orders);   // best-warehouse comparison (independent of WH override)
      })
      .catch(function () { btn.disabled = false; statusEl.textContent = 'Network error — retry.'; });
  }
  btn.addEventListener('click', run);
  document.querySelectorAll('.av-tab').forEach(function (t) {
    t.addEventListener('click', function () {
      document.querySelectorAll('.av-tab').forEach(function (x) { x.classList.remove('on'); });
      document.querySelectorAll('.av-pane').forEach(function (x) { x.classList.remove('on'); });
      t.classList.add('on'); document.getElementById('pane-' + t.dataset.pane).classList.add('on');
    });
  });

  // Export → styled multi-sheet .xlsx (POST current orders + WH → blob download).
  var exportBtn = document.getElementById('av-export');
  exportBtn.addEventListener('click', function () {
    var orders = ta.value.trim();
    if (!orders) { statusEl.textContent = 'Nothing to export — run a check first.'; return; }
    exportBtn.disabled = true;
    if (window.B2B && B2B.toast) B2B.toast('Building the availability workbook…', { type: 'info', title: 'Downloading', timeout: 6000 });
    var body = new URLSearchParams(); body.set('orders', orders); body.set('warehouse', whSel.value);
    fetch(CFG["b2b_availability_export"], { method: 'POST', headers: { 'X-CSRFToken': csrf, 'X-Requested-With': 'XMLHttpRequest', 'Content-Type': 'application/x-www-form-urlencoded' }, body: body.toString() })
      .then(function (r) { return r.blob(); })
      .then(function (blob) {
        var url = URL.createObjectURL(blob), a = document.createElement('a');
        a.href = url; a.download = 'availability.xlsx'; document.body.appendChild(a); a.click();
        setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 1500);
        exportBtn.disabled = false;
      })
      .catch(function () { exportBtn.disabled = false; if (window.B2B && B2B.toast) B2B.toast('Export failed — retry.', { type: 'error' }); });
  });

  // Click an SKU row → lazy-load which bins hold that item (INCLUDED vs EXCLUDED).
  var binsUrl = CFG["b2b_availability_bins"];
  var binCache = {};
  document.getElementById('pane-sku').addEventListener('click', function (e) {
    var tr = e.target.closest('.sku-row'); if (!tr) return;
    var cx = tr.querySelector('.sku-cx');
    var nx = tr.nextElementSibling;
    if (nx && nx.classList.contains('sku-bins-row')) { nx.remove(); cx.innerHTML = '&#9656;'; return; }
    var wh = tr.getAttribute('data-wh'), item = tr.getAttribute('data-item');
    cx.innerHTML = '&#9662;';
    var row = document.createElement('tr'); row.className = 'sku-bins-row';
    var td = document.createElement('td'); td.colSpan = 14;
    td.innerHTML = '<div class="sku-bins-load">Loading bins…</div>';
    row.appendChild(td); tr.parentNode.insertBefore(row, tr.nextSibling);
    function paint(bins) {
      if (!bins.length) { td.innerHTML = '<div class="sku-bins-empty">No sellable-bin stock for this item.</div>'; return; }
      var total = bins.reduce(function (s, b) { return s + (Number(b.qty) || 0); }, 0);
      td.innerHTML = '<div class="sku-bins"><table><thead><tr><th>Sellable Bin</th><th>Zone</th><th class="num">Qty</th></tr></thead><tbody>' +
        bins.map(function (b) {
          return '<tr><td>' + esc(b.bin) + '</td><td>' + esc(b.zone) + '</td>' +
            '<td class="num">' + nf(b.qty) + '</td></tr>';
        }).join('') +
        '<tr class="sku-bins-tot"><td>Available</td><td></td><td class="num">' + nf(total) + '</td></tr>' +
        '</tbody></table></div>';
    }
    var key = wh + '|' + item;
    if (binCache[key]) { paint(binCache[key]); return; }
    fetch(binsUrl + '?wh=' + encodeURIComponent(wh) + '&item=' + encodeURIComponent(item), { headers: { 'X-Requested-With': 'XMLHttpRequest' } })
      .then(function (r) { return r.json(); })
      .then(function (j) { var b = (j.ok && j.bins) || []; binCache[key] = b; paint(b); })
      .catch(function () { td.innerHTML = '<div class="sku-bins-empty">Could not load bins.</div>'; });
  });

  // Persist input on edit, and restore + auto-run on load so a refresh keeps
  // your list and results.
  ta.addEventListener('input', saveState);
  whSel.addEventListener('change', saveState);
  (function restore() {
    var s;
    try { s = JSON.parse(localStorage.getItem(LS_KEY) || '{}'); } catch (e) { s = {}; }
    if (s && s.orders && s.orders.trim()) {
      ta.value = s.orders;
      if (s.warehouse) whSel.value = s.warehouse;
      run();   // re-check → results reappear without re-pasting
    }
  })();
})();
