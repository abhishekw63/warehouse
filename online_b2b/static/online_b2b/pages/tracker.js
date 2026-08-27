/* online_b2b/tracker.html — page script (separated from template). */
(function () {
  var body = document.getElementById('trk-body');
  var loader = document.getElementById('trkLoader');
  var filter = document.getElementById('trkFilter');
  var clearBtn = document.getElementById('trkClear');
  var exportBtn = document.getElementById('trkExport');
  var base = location.pathname;

  // ── top loading bar ─────────────────────────────────────────────────
  // A slim bar slides across the top while async work (filter swap · billing ·
  // today KPIs) is in flight — but only if it runs past ~0.7s, so quick ops never
  // flash it. Body-appended so #MainContent's view-transition can't trap it.
  Array.prototype.forEach.call(document.querySelectorAll('body > #trkTopbar'), function (el) { el.remove(); });
  var topbar = document.getElementById('trkTopbar');
  if (topbar) document.body.appendChild(topbar);
  var _pending = 0, _progTimer = null;
  function progStart() {
    _pending++;
    if (!_progTimer && topbar) _progTimer = setTimeout(function () { topbar.hidden = false; }, 700);
  }
  function progDone() {
    _pending = Math.max(0, _pending - 1);
    if (_pending === 0) { if (_progTimer) { clearTimeout(_progTimer); _progTimer = null; } if (topbar) topbar.hidden = true; }
  }

  // last Today-KPI payload (drives the facility drawer) + which metric it's open on
  var todayData = null, facMetric = null;

  var FIELDS = ['segment', 'marketplace', 'warehouse', 'q', 'uploaded_from', 'uploaded_to'];
  function params() {
    var p = new URLSearchParams();
    FIELDS.forEach(function (n) {
      var el = filter.querySelector('[name="' + n + '"]');
      if (el && el.value) p.set(n, el.value);
    });
    return p;
  }

  function loadBody() {
    var p = params();
    var lt = setTimeout(function () { if (loader) loader.hidden = false; }, 180);  // loader only if slow
    body.classList.add('fade'); progStart();
    fetch(base + '?partial=1&' + p.toString(), { headers: { 'X-Requested-With': 'fetch' } })
      .then(function (r) { return r.text(); })
      .then(function (html) {
        clearTimeout(lt); if (loader) loader.hidden = true;
        body.innerHTML = html; body.classList.remove('fade');
        initTable(); syncUI(p); loadBilling(); loadTodayKPIs(); renderPills();
        try { document.dispatchEvent(new CustomEvent('trk:filterchange')); } catch (e) { }
        history.replaceState(null, '', p.toString() ? base + '?' + p.toString() : base);
        progDone();
      })
      .catch(function () { clearTimeout(lt); if (loader) loader.hidden = true; body.classList.remove('fade'); progDone(); });
  }

  function syncUI(p) {
    var any = FIELDS.some(function (n) { return p.get(n); });
    if (clearBtn) clearBtn.hidden = !any;
    if (exportBtn) exportBtn.href = base + 'export/?' + p.toString();
  }

  // filter events (no reload — loadBody does an AJAX partial swap with a loader)
  filter.querySelectorAll('select').forEach(function (s) { s.addEventListener('change', loadBody); });
  filter.querySelectorAll('input[type=date]').forEach(function (d) { d.addEventListener('change', loadBody); });
  filter.addEventListener('submit', function (e) { e.preventDefault(); loadBody(); });

  // ── single date-range control (one button → popover: presets + from/to) ──
  var drBtn = document.getElementById('trkDRBtn'), drPop = document.getElementById('trkDRPop'),
      drLabel = document.getElementById('trkDRLabel'), drApply = document.getElementById('trkDRApply'),
      drWrap = document.getElementById('trkDR'),
      dFrom = filter.querySelector('[name=uploaded_from]'), dTo = filter.querySelector('[name=uploaded_to]');
  var MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  function ymd(d) { return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
  function drFmt(s) { if (!s) return ''; var p = s.split('-'); return p[2] + ' ' + MON[(+p[1]) - 1]; }
  function drLabelUpdate() {
    var a = dFrom && dFrom.value, b = dTo && dTo.value;
    if (drLabel) drLabel.textContent = (a || b) ? ((drFmt(a) || '…') + ' – ' + (drFmt(b) || '…')) : 'Any dates';
    if (drWrap) drWrap.classList.toggle('on', !!(a || b));
  }
  function drOpen(open) {
    if (!drPop) return;
    drPop.hidden = !open;
    if (drBtn) drBtn.setAttribute('aria-expanded', open ? 'true' : 'false');
    if (drWrap) drWrap.classList.toggle('pop', open);
  }
  if (drBtn) drBtn.addEventListener('click', function (e) { e.stopPropagation(); drOpen(drPop.hidden); });
  document.addEventListener('click', function (e) { if (drWrap && !drWrap.contains(e.target)) drOpen(false); });
  document.addEventListener('keydown', function (e) { if (e.key === 'Escape') drOpen(false); });
  if (drPop) drPop.addEventListener('click', function (e) {
    var p = e.target.closest && e.target.closest('[data-preset]');
    if (!p) return;
    var k = p.getAttribute('data-preset'), now = new Date(), f = '', t = '', d2;
    if (k === 'today') { f = t = ymd(now); }
    else if (k === '7') { d2 = new Date(now); d2.setDate(now.getDate() - 6); f = ymd(d2); t = ymd(now); }
    else if (k === '30') { d2 = new Date(now); d2.setDate(now.getDate() - 29); f = ymd(d2); t = ymd(now); }
    else if (k === 'month') { f = ymd(new Date(now.getFullYear(), now.getMonth(), 1)); t = ymd(now); }
    if (dFrom) dFrom.value = f; if (dTo) dTo.value = t;   // 'clear' → both ''
    drLabelUpdate(); drOpen(false); loadBody();
  });
  if (drApply) drApply.addEventListener('click', function () { drLabelUpdate(); drOpen(false); loadBody(); });
  [dFrom, dTo].forEach(function (el) { if (el) el.addEventListener('change', drLabelUpdate); });
  drLabelUpdate();

  // ── multi-order search (paste a list → filter to exactly those orders) ──
  // A single-line <input> strips newlines on paste, so a popover textarea holds the
  // list; on apply we stuff the comma-joined list into q (survives the input's value
  // sanitisation) and the normal filter + export path carries it. Backend flips to an
  // exact-match po/external_doc lookup when it sees 2+ separated tokens.
  (function () {
    var mBtn = document.getElementById('trkMultiBtn'), mPop = document.getElementById('trkMultiPop'),
        mWrap = document.getElementById('trkMulti'), mTA = document.getElementById('trkMultiTA'),
        mCnt = document.getElementById('trkMultiCnt'), mApply = document.getElementById('trkMultiApply'),
        mClear = document.getElementById('trkMultiClear'), qEl = filter.querySelector('[name=q]');
    if (!mBtn || !mTA || !qEl) return;
    function toks(s) {
      var out = [], seen = {};
      String(s || '').split(/[\n\r,;|\t]+/).forEach(function (t) {
        t = t.trim(); var k = t.toLowerCase();
        if (t && !seen[k]) { seen[k] = 1; out.push(t); }
      });
      return out;
    }
    function isMulti(s) { return toks(s).length >= 2; }
    function count() {
      var n = toks(mTA.value).length;
      if (mCnt) mCnt.textContent = n + ' order' + (n === 1 ? '' : 's');
      if (mApply) mApply.disabled = n === 0;
    }
    function open(o) {
      if (!mPop) return;
      mPop.hidden = !o; mBtn.setAttribute('aria-expanded', o ? 'true' : 'false');
      if (mWrap) mWrap.classList.toggle('on', o);
      if (o) { count(); setTimeout(function () { mTA.focus(); }, 20); }
    }
    if (isMulti(qEl.value)) mTA.value = toks(qEl.value).join('\n');   // seed from an existing multi q
    mBtn.addEventListener('click', function (e) { e.stopPropagation(); open(mPop.hidden); });
    mTA.addEventListener('input', count);
    document.addEventListener('click', function (e) { if (mWrap && !mWrap.contains(e.target)) open(false); });
    document.addEventListener('keydown', function (e) { if (e.key === 'Escape') open(false); });
    if (mApply) mApply.addEventListener('click', function () {
      qEl.value = toks(mTA.value).join(', '); open(false); loadBody();
    });
    if (mClear) mClear.addEventListener('click', function () {
      mTA.value = ''; count();
      if (qEl.value) { qEl.value = ''; loadBody(); }
      open(false);
    });
    // paste an Excel column straight into the q box → auto-switch to multi
    qEl.addEventListener('paste', function (e) {
      var cd = e.clipboardData || window.clipboardData, txt = cd && cd.getData('text');
      if (txt && isMulti(txt)) {
        e.preventDefault();
        var t = toks(txt); qEl.value = t.join(', '); mTA.value = t.join('\n'); count(); loadBody();
      }
    });
    count();
  })();

  // Facility chips live INSIDE the re-rendered body → delegate on the persistent
  // container: a click sets the warehouse filter and reloads (one filter path).
  body.addEventListener('click', function (e) {
    var chip = e.target.closest && e.target.closest('.trk-fac');
    if (!chip) return;
    var whSel = filter.querySelector('[name="warehouse"]');
    if (whSel) whSel.value = chip.getAttribute('data-fac') || '';
    loadBody();
  });
  var qEl = filter.querySelector('[name=q]'), qT;
  if (qEl) qEl.addEventListener('input', function () { clearTimeout(qT); qT = setTimeout(loadBody, 400); });
  function resetSeg() {
    var si = filter.querySelector('input[name=segment]'); if (si) si.value = '';
    filter.querySelectorAll('.trk-seg-b').forEach(function (b) { b.classList.toggle('on', !b.getAttribute('data-seg')); });
  }
  if (clearBtn) clearBtn.addEventListener('click', function (e) {
    e.preventDefault();
    filter.querySelectorAll('select').forEach(function (s) { s.value = ''; });
    filter.querySelectorAll('input[type=date]').forEach(function (d) { d.value = ''; });
    if (qEl) qEl.value = '';
    resetSeg();
    if (typeof drLabelUpdate === 'function') drLabelUpdate();
    loadBody();
  });

  // segmented Dept control → drives the hidden segment input, then reloads
  var segInput = filter.querySelector('input[name=segment]');
  filter.querySelectorAll('.trk-seg-b').forEach(function (btn) {
    btn.addEventListener('click', function () {
      if (segInput) segInput.value = btn.getAttribute('data-seg') || '';
      filter.querySelectorAll('.trk-seg-b').forEach(function (b) { b.classList.remove('on'); });
      btn.classList.add('on');
      loadBody();
    });
  });

  // density toggle (compact ↔ comfortable), remembered per browser
  var DKEY = 'trk_density', densityBtn = document.getElementById('trkDensity');
  function applyDensity() {
    var dense = false; try { dense = localStorage.getItem(DKEY) === '1'; } catch (e) {}
    document.body.classList.toggle('trk-dense', dense);
    if (densityBtn) { densityBtn.classList.toggle('on', dense); densityBtn.title = dense ? 'Comfortable rows' : 'Compact rows'; }
  }
  if (densityBtn) densityBtn.addEventListener('click', function () {
    var dense = !document.body.classList.contains('trk-dense');
    try { localStorage.setItem(DKEY, dense ? '1' : '0'); } catch (e) {}
    applyDensity();
  });
  applyDensity();

  // active-filter pills — dismissible chips reflecting the applied filters
  var pillsBox = document.getElementById('trkPills');
  var PILLDEFS = [{ n: 'segment', l: 'Dept' }, { n: 'marketplace', l: 'Marketplace' },
    { n: 'warehouse', l: 'Facility' }, { n: 'q', l: 'Search' },
    { n: 'uploaded_from', l: 'From' }, { n: 'uploaded_to', l: 'To' }];
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function renderPills() {
    if (!pillsBox) return;
    var html = '';
    PILLDEFS.forEach(function (d) {
      var el = filter.querySelector('[name=' + d.n + ']'), v = el ? (el.value || '').trim() : '';
      if (!v) return;
      var disp = v;
      if (d.n === 'q') {                       // a pasted multi-order list → show a count, not the raw string
        var parts = v.split(/[\n\r,;|\t]+/).map(function (x) { return x.trim(); }).filter(Boolean);
        if (parts.length >= 2) disp = parts.length + ' orders';
      }
      html += '<span class="trk-pill" data-clear="' + d.n + '"><span class="tp-l">' + d.l + '</span>' + esc(disp) + '<button type="button" aria-label="remove ' + d.l + ' filter">×</button></span>';
    });
    pillsBox.innerHTML = html;
    pillsBox.classList.toggle('has', !!html);
  }
  if (pillsBox) pillsBox.addEventListener('click', function (e) {
    var p = e.target.closest('.trk-pill'); if (!p) return;
    var name = p.getAttribute('data-clear'), el = filter.querySelector('[name=' + name + ']');
    if (el) el.value = '';
    if (name === 'q') { var ta = document.getElementById('trkMultiTA'); if (ta) { ta.value = ''; ta.dispatchEvent(new Event('input')); } }
    if (name === 'segment') resetSeg();
    loadBody();
  });

  // ── async "Est. Billing" column ────────────────────────────────────
  // Billable-from-current-stock ₹ per PO is a ~2s inventory pass, so it's fetched
  // AFTER the table paints (never blocks render) and filled into placeholder cells.
  // Re-runs after every filter swap for the freshly-shown rows.
  function csrfToken() {
    var el = document.querySelector('[name=csrfmiddlewaretoken]');
    if (el) return el.value;
    var m = document.cookie.match(/csrftoken=([^;]+)/);
    return m ? m[1] : '';
  }
  function fillPctClass(p) { return p >= 95 ? 'ok' : (p >= 60 ? 'mid' : 'low'); }
  function inrShortJs(v) {                    // mirrors the |inr_short filter for JS-built KPIs
    v = Number(v) || 0;
    if (v >= 1e7) return '₹' + (v / 1e7).toFixed(2).replace(/\.?0+$/, '') + ' Cr';
    if (v >= 1e5) return '₹' + (v / 1e5).toFixed(2).replace(/\.?0+$/, '') + ' L';
    return '₹' + Math.round(v).toLocaleString('en-IN');
  }
  function compactJs(v) {                      // mirrors the |compact filter (1.2k / 3.4M)
    v = Number(v) || 0;
    if (v >= 1e6) return (v / 1e6).toFixed(1) + 'M';
    if (v >= 1e3) return (v / 1e3).toFixed(1) + 'k';
    return String(Math.round(v));
  }
  function fillBreakdown(bill) {               // Full→D365→Excluded columns (async)
    body.querySelectorAll('.trk-bd[data-bd-po]').forEach(function (c) {
      var b = bill[c.getAttribute('data-bd-po')], kind = c.getAttribute('data-bd');
      if (!b) { c.innerHTML = '<span class="muted">—</span>'; return; }
      if (kind === 'uplqty') { c.setAttribute('data-v', b.upl_qty); c.innerHTML = compactJs(b.upl_qty); }
      else if (kind === 'uplval') { c.setAttribute('data-v', b.upl_value); c.innerHTML = inrShortJs(b.upl_value); }
      else if (kind === 'exclqty') { c.setAttribute('data-v', b.excl_qty); c.innerHTML = b.excl_qty > 0 ? '<span class="trk-excl-v">−' + compactJs(b.excl_qty) + '</span>' : '<span class="muted">0</span>'; }
      else if (kind === 'exclval') { c.setAttribute('data-v', b.excl_value); c.innerHTML = b.excl_value > 0 ? '<span class="trk-excl-v">−' + inrShortJs(b.excl_value) + '</span>' : '<span class="muted">—</span>'; }
    });
  }
  function loadBilling() {
    var cells = body.querySelectorAll('.trk-bill[data-bill-po]');
    if (!cells.length) return;
    var pos = [], seen = {};
    cells.forEach(function (c) { var p = c.getAttribute('data-bill-po'); if (p && !seen[p]) { seen[p] = 1; pos.push(p); } });
    progStart();
    fetch(appRoot + 'tracker/billing/', {
      method: 'POST', headers: { 'Content-Type': 'application/json', 'X-CSRFToken': csrfToken(), 'X-Requested-With': 'fetch' },
      body: JSON.stringify({ pos: pos })
    })
      .then(function (r) { return r.json(); })
      .then(function (d) {
        if (!d || !d.ok) throw 0;
        var bill = d.bill || {};
        cells.forEach(function (c) {
          var b = bill[c.getAttribute('data-bill-po')];
          if (b && !b.no_stock) {
            c.setAttribute('data-v', b.est);
            var cls = fillPctClass(b.fill), pf = Math.round(b.fill);
            c.innerHTML = '<div class="trk-bill-cell"><div class="trk-bill-top">' +
              '<span class="trk-bill-v" title="₹' + Math.round(b.est).toLocaleString('en-IN') +
              ' billable of ₹' + Math.round(b.ord).toLocaleString('en-IN') + ' · short ₹' +
              Math.round(b.short).toLocaleString('en-IN') + ' · from ' + b.wh + ' stock">' + b.est_fmt + '</span>' +
              '<span class="trk-bill-p ' + cls + '">' + pf + '%</span></div>' +
              '<div class="trk-bill-meter"><i class="' + cls + '" style="width:' + Math.min(100, pf) + '%"></i></div></div>';
          } else if (b && b.no_stock) {
            c.innerHTML = '<span class="muted" title="No current inventory snapshot for ' + b.wh + '">— <span class="trk-bill-na">no stock</span></span>';
          } else {
            c.innerHTML = '<span class="muted" title="No recorded lines to bill">—</span>';
          }
        });
        fillBreakdown(bill);            // Full → D365 → Excluded qty/value columns
        var asof = document.getElementById('trkBillAsof'), foot = document.getElementById('trkBillFoot');
        if (asof) asof.textContent = d.as_of_short ? 'as of ' + d.as_of_short : '';
        if (foot && d.as_of) foot.textContent = ' (as of ' + d.as_of + ')';
        progDone();
      })
      .catch(function () {
        body.querySelectorAll('.trk-bill[data-bill-po] .trk-bill-wait').forEach(function (w) {
          w.parentElement.innerHTML = '<span class="muted">—</span>';
        });
        progDone();
      });
  }

  // ── today's fulfilment KPI strip ────────────────────────────────────
  // Scoped to orders UPLOADED TODAY (client local date), honoring the current
  // filters — NOT all POs, NOT the shown window. Server does the full-day count /
  // value / billable / at-risk in one call; async so it never blocks render.
  function localToday() {
    var d = new Date();
    return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
  }
  function loadTodayKPIs() {
    var box = document.querySelector('.trk-kpis');
    if (!box) return;
    var p = new URLSearchParams();
    ['segment', 'marketplace', 'warehouse', 'q'].forEach(function (n) {
      var el = filter.querySelector('[name="' + n + '"]');
      if (el && el.value) p.set(n, el.value);
    });
    p.set('d', localToday());
    var $ = function (id) { return document.getElementById(id); };
    progStart();
    fetch(appRoot + 'tracker/today/?' + p.toString(), { headers: { 'X-Requested-With': 'fetch' } })
      .then(function (r) { return r.json(); })
      .then(function (d) {
        if (!d || !d.ok) throw 0;
        if ($('kpiOrdersN')) $('kpiOrdersN').textContent = (d.count || 0).toLocaleString('en-IN');
        if ($('kpiValN')) $('kpiValN').textContent = d.value_fmt || '₹0';
        if ($('kpiBillVal')) $('kpiBillVal').textContent = d.billable_fmt || '₹0';
        if ($('kpiBillSub')) $('kpiBillSub').textContent = (d.avg_fill || 0) + '% avg fill' + (d.as_of_short ? ' · as of ' + d.as_of_short : '');
        var rn = $('kpiRiskN');
        if (rn) { rn.textContent = d.risk || 0; rn.parentElement.parentElement.classList.toggle('has-risk', (d.risk || 0) > 0); }
        if ($('kpiRiskSub')) $('kpiRiskSub').textContent = d.risk ? (d.short_fmt + ' short · fill < 60%') : 'all today ≥ 60% fill';
        // dept-wise split (Online B2B vs Offline): a proportional stacked bar + labels
        var segs = d.segments || [];
        document.querySelectorAll('.trk-kpis .k-split').forEach(function (el) {
          var kpi = el.getAttribute('data-kpi');
          var raw = function (s) { return Number(kpi === 'orders' ? s.count : kpi === 'value' ? s.value : kpi === 'bill' ? s.billable : s.risk) || 0; };
          var disp = function (s) { return kpi === 'orders' ? s.count : kpi === 'value' ? s.value_fmt : kpi === 'bill' ? s.billable_fmt : s.risk; };
          var total = segs.reduce(function (a, s) { return a + raw(s); }, 0);
          if (total <= 0) { el.innerHTML = ''; return; }
          var bar = segs.map(function (s) {
            return '<i class="' + (s.code === 'Offline' ? 'off' : 'on') + '" style="width:' + (raw(s) / total * 100).toFixed(1) + '%"></i>';
          }).join('');
          var leg = segs.map(function (s) {
            return '<span class="ks ' + (s.code === 'Offline' ? 'off' : 'on') + '"><i></i>' + (s.code === 'Offline' ? 'Off' : 'B2B') + ' <b>' + disp(s) + '</b></span>';
          }).join('');
          el.innerHTML = '<div class="k-bar">' + bar + '</div><div class="k-barleg">' + leg + '</div>';
        });
        todayData = d;
        if (facMetric) { renderFacDrawer(facMetric); positionFacDrawer(facMetric); }   // keep an open drawer in sync
        progDone();
      })
      .catch(function () {
        ['kpiOrdersN', 'kpiValN', 'kpiBillVal', 'kpiRiskN'].forEach(function (id) { if ($(id)) $(id).textContent = '—'; });
        progDone();
      });
  }

  // ── facility (AHD/BLR/North) breakdown drawer ─────────────────────────────
  // A KPI card is clickable; clicking opens a drawer right below the strip with
  // that card's metric split across the 3 facilities (proportional bar + value).
  // Click the same card (or ✕) to close; clicking another card swaps the metric.
  var FAC_LABEL = { orders: 'Orders', value: 'Value', bill: 'Billable', risk: 'At-risk' };
  function facRaw(f, m) { return Number(m === 'orders' ? f.count : m === 'value' ? f.value : m === 'bill' ? f.billable : f.risk) || 0; }
  function facDisp(f, m) { return m === 'orders' ? (f.count || 0).toLocaleString('en-IN') : m === 'value' ? f.value_fmt : m === 'bill' ? f.billable_fmt : f.risk; }
  function renderFacDrawer(m) {
    var body = document.getElementById('tfdBody'), metEl = document.getElementById('tfdMetric');
    if (!body) return;
    if (metEl) metEl.textContent = FAC_LABEL[m] || m;
    var facs = (todayData && todayData.facilities) || [];
    var total = facs.reduce(function (a, f) { return a + facRaw(f, m); }, 0);
    if (!facs.length) { body.innerHTML = '<div class="tfd-empty">No orders recorded today.</div>'; return; }
    body.innerHTML = facs.map(function (f) {
      var val = facRaw(f, m), pct = total > 0 ? (val / total * 100) : 0;
      var cls = 'fc-' + String(f.code || '').toLowerCase();
      return '<div class="tfd-row">' +
        '<span class="tfd-fac ' + cls + '"><i></i>' + (f.label || f.code) + '</span>' +
        '<div class="tfd-track"><span class="tfd-fill ' + cls + '" style="width:' + pct.toFixed(1) + '%"></span></div>' +
        '<span class="tfd-val"><b>' + facDisp(f, m) + '</b><em>' + (f.count || 0) + ' PO' + ((f.count || 0) === 1 ? '' : 's') + ' · ' + Math.round(pct) + '%</em></span>' +
        '</div>';
    }).join('');
  }
  // Anchor the drawer UNDER the clicked card, matching that card's width + left
  // offset (not full-width). If the cards have wrapped/stacked (narrow screen),
  // fall back to full width so it stays readable.
  function positionFacDrawer(m) {
    var dr = document.getElementById('trkFacDrawer');
    var card = document.querySelector('.trk-kpi[data-metric="' + m + '"]');
    var strip = document.querySelector('.trk-kpis');
    if (!dr || !card || !strip) return;
    var s = strip.getBoundingClientRect(), c = card.getBoundingClientRect();
    if (c.width >= s.width - 4) { dr.style.marginLeft = '0px'; dr.style.width = ''; }  // stacked → full width
    else { dr.style.marginLeft = (c.left - s.left) + 'px'; dr.style.width = c.width + 'px'; }
  }
  function openFacDrawer(m) {
    var dr = document.getElementById('trkFacDrawer');
    if (!dr) return;
    facMetric = m;
    renderFacDrawer(m);
    dr.hidden = false;
    positionFacDrawer(m);        // size + place it under the clicked card
    // reflect state on the cards (active outline + aria)
    document.querySelectorAll('.trk-kpi[data-metric]').forEach(function (c) {
      var on = c.getAttribute('data-metric') === m;
      c.classList.toggle('kpi-active', on);
      c.setAttribute('aria-expanded', on ? 'true' : 'false');
    });
    requestAnimationFrame(function () { dr.classList.add('open'); });
  }
  function closeFacDrawer() {
    var dr = document.getElementById('trkFacDrawer');
    facMetric = null;
    document.querySelectorAll('.trk-kpi[data-metric]').forEach(function (c) { c.classList.remove('kpi-active'); c.setAttribute('aria-expanded', 'false'); });
    if (!dr) return;
    dr.classList.remove('open');
    clearTimeout(dr._t);
    dr._t = setTimeout(function () { if (!facMetric) dr.hidden = true; }, 260);
  }
  function toggleFacDrawer(m) { if (facMetric === m) closeFacDrawer(); else openFacDrawer(m); }
  function wireFacDrawer() {
    document.querySelectorAll('.trk-kpi[data-metric]').forEach(function (c) {
      var m = c.getAttribute('data-metric');
      c.addEventListener('click', function () { toggleFacDrawer(m); });
      c.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleFacDrawer(m); } });
    });
    var x = document.getElementById('tfdX');
    if (x) x.addEventListener('click', closeFacDrawer);
    // re-anchor under its card when the layout reflows (cards are fluid width)
    var _rt;
    window.addEventListener('resize', function () {
      if (!facMetric) return;
      clearTimeout(_rt); _rt = setTimeout(function () { if (facMetric) positionFacDrawer(facMetric); }, 80);
    });
  }

  // table init (re-run after every AJAX swap): sort + quick-filter
  function initTable() {
    var table = body.querySelector('table.sortable');
    if (table) table.querySelectorAll('th.srt').forEach(function (th) {
      th.addEventListener('click', function () {
        var k = th.getAttribute('data-k'), asc = th.getAttribute('data-dir') !== 'asc';
        table.querySelectorAll('th.srt').forEach(function (x) { x.removeAttribute('data-dir'); x.classList.remove('on'); });
        th.setAttribute('data-dir', asc ? 'asc' : 'desc'); th.classList.add('on');
        var tb = table.tBodies[0];
        Array.prototype.slice.call(tb.rows).sort(function (a, b) {
          var ae = a.querySelector('[data-k="' + k + '"]'), be = b.querySelector('[data-k="' + k + '"]');
          var av = ae ? (ae.getAttribute('data-v') || ae.textContent).trim() : '';
          var bv = be ? (be.getAttribute('data-v') || be.textContent).trim() : '';
          var na = parseFloat(av), nb = parseFloat(bv);
          if (!isNaN(na) && !isNaN(nb)) return asc ? na - nb : nb - na;
          return asc ? av.localeCompare(bv) : bv.localeCompare(av);
        }).forEach(function (r) { tb.appendChild(r); });
      });
    });
  }

  // quick client-side filter over loaded rows (re-queries fresh each keypress)
  var quick = document.getElementById('trkQuick'), qkT;
  if (quick) quick.addEventListener('input', function () {
    clearTimeout(qkT);
    qkT = setTimeout(function () {            // debounce so fast typing never janks
      var v = quick.value.trim().toLowerCase();
      body.querySelectorAll('.trk-tbl tbody tr').forEach(function (tr) {
        tr.style.display = (!v || tr.textContent.toLowerCase().indexOf(v) >= 0) ? '' : 'none';
      });
    }, 140);
  });

  // ── per-PO drill drawer ────────────────────────────────────────────
  // Click a PO → open its detail in place (header + Full→Excluded→Final +
  // every line) — no full-run navigation. The PO cell is re-rendered on every
  // filter swap, so the click is delegated on the persistent #trk-body.
  // #MainContent carries a view-transition-name, which makes it a containing block
  // for position:fixed descendants (Chrome) — that trapped the drawer below the
  // shell header, hiding its top bar (resize/close). Adopt the drawer + overlay
  // onto <body> so they're viewport-anchored. Remove any orphan first: shell-nav
  // swaps #MainContent, so a prior visit's copy can linger on <body>.
  Array.prototype.forEach.call(
    document.querySelectorAll('body > #odDrawer, body > #odOverlay'),
    function (el) { el.remove(); });
  var drawer = document.getElementById('odDrawer'),
      overlay = document.getElementById('odOverlay'),
      panel = document.getElementById('odPanel'),
      odClose = document.getElementById('odClose'),
      odResize = document.getElementById('odResize'),
      appRoot = base.replace(/tracker\/?$/, ''),   // /…/tracker/ → /…/
      odLast = null;
  if (drawer) document.body.appendChild(drawer);
  if (overlay) document.body.appendChild(overlay);

  function openDrawer(po) {
    if (!drawer || !po) return;
    odLast = document.activeElement;
    panel.innerHTML = '<div class="od-loading"><span class="trk-spin"></span> Loading…</div>';
    drawer.hidden = false; overlay.hidden = false;
    requestAnimationFrame(function () { drawer.classList.add('open'); overlay.classList.add('open'); });
    drawer.setAttribute('aria-hidden', 'false');
    document.body.classList.add('od-lock');
    if (odClose) odClose.focus();
    // keep '/' raw (PO ids like SO/GTM/8982 → matched by the <path:po> route),
    // encode every other special char.
    var poPath = encodeURIComponent(po).replace(/%2F/g, '/');
    fetch(appRoot + 'order/' + poPath + '/', { headers: { 'X-Requested-With': 'fetch' } })
      .then(function (r) { return r.text(); })
      .then(function (html) {
        panel.innerHTML = html; panel.scrollTop = 0;
        // dynamic fill-meter width comes from data-w (kept out of the template so
        // styling stays in the css/js layer, not inline markup).
        panel.querySelectorAll('.od-fmeter i[data-w]').forEach(function (el) { el.style.width = el.getAttribute('data-w') + '%'; });
      })
      .catch(function () { panel.innerHTML = '<div class="od-empty">Could not load PO ' + po + '.</div>'; });
  }
  function closeDrawer() {
    if (!drawer || drawer.hidden) return;
    drawer.classList.remove('open'); overlay.classList.remove('open');
    drawer.setAttribute('aria-hidden', 'true');
    document.body.classList.remove('od-lock');
    setTimeout(function () { drawer.hidden = true; overlay.hidden = true; }, 260);
    if (odLast && odLast.focus) odLast.focus();
  }
  body.addEventListener('click', function (e) {
    var cell = e.target.closest && e.target.closest('.trk-po');
    if (cell) openDrawer(cell.getAttribute('data-po'));
  });
  body.addEventListener('keydown', function (e) {
    var t = e.target;
    if ((e.key === 'Enter' || e.key === ' ') && t.classList && t.classList.contains('trk-po')) {
      e.preventDefault(); openDrawer(t.getAttribute('data-po'));
    }
  });
  if (overlay) overlay.addEventListener('click', closeDrawer);
  if (odClose) odClose.addEventListener('click', closeDrawer);
  if (odResize) odResize.addEventListener('click', function () {   // '‹' widen ↔ '›' restore
    var wide = drawer.classList.toggle('od-wide');
    odResize.innerHTML = wide ? '›' : '‹';
    odResize.title = wide ? 'Restore width' : 'Widen panel';
  });
  document.addEventListener('keydown', function (e) { if (e.key === 'Escape') closeDrawer(); });

  // full-screen view — expands the whole tracker to fill the screen (native API)
  var fullBtn = document.getElementById('trkFull'), trkWrap = document.querySelector('.trk-wrap');
  if (fullBtn && trkWrap) {
    fullBtn.addEventListener('click', function () {
      if (!document.fullscreenElement) {
        (trkWrap.requestFullscreen || trkWrap.webkitRequestFullscreen || function () {}).call(trkWrap);
      } else {
        (document.exitFullscreen || document.webkitExitFullscreen || function () {}).call(document);
      }
    });
    document.addEventListener('fullscreenchange', function () {
      var on = document.fullscreenElement === trkWrap;
      fullBtn.textContent = on ? '⛶ Exit full screen' : '⛶ Full screen';
      fullBtn.classList.toggle('on', on);
      // the drawer + overlay live on <body>; move them into the fullscreen element
      // while FS is active so the per-PO drawer still shows, then back on exit.
      var host = on ? trkWrap : document.body;
      if (drawer && drawer.parentNode !== host) host.appendChild(drawer);
      if (overlay && overlay.parentNode !== host) host.appendChild(overlay);
    });
  }

  // collapse / expand the orders table (remembered) — toggled from BOTH the hero
  // button and the header bar sitting right above the table (always reachable).
  var tblBtn = document.getElementById('trkTableToggle'),
      tblHead = document.getElementById('trkTableHead');
  // The orders table starts COLLAPSED on every load (per request). The body class is
  // the single source of truth; toggled from the hero button or the header bar.
  function applyTableCollapse() {
    var on = !document.body.classList.contains('trk-tbl-open');   // collapsed = the default (no .trk-tbl-open)
    if (tblBtn) { tblBtn.classList.toggle('on', on); tblBtn.innerHTML = on ? '⊞ Table' : '⊟ Table'; tblBtn.title = on ? 'Show the orders table' : 'Collapse the orders table'; }
    if (tblHead) {
      tblHead.setAttribute('aria-expanded', on ? 'false' : 'true');
      var hint = tblHead.querySelector('.trk-thead-hint');
      if (hint) hint.textContent = on ? 'click to expand' : 'click to collapse';
    }
  }
  function toggleTable() {
    document.body.classList.toggle('trk-tbl-open');
    applyTableCollapse();
  }
  if (tblBtn) tblBtn.addEventListener('click', toggleTable);
  if (tblHead) {
    tblHead.addEventListener('click', toggleTable);
    tblHead.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleTable(); } });
  }
  applyTableCollapse();   // default is collapsed (CSS: body:not(.trk-tbl-open)) — no class to add, no paint-then-yank flash

  // add-PO panel toggle
  var addBtn = document.getElementById('trkAddBtn'), addForm = document.getElementById('trkAdd'),
      addCancel = document.getElementById('trkAddCancel');
  if (addBtn && addForm) addBtn.addEventListener('click', function () {
    addForm.hidden = !addForm.hidden;
    if (!addForm.hidden) { var f = addForm.querySelector('[name=po]'); if (f) f.focus(); }
  });
  if (addCancel) addCancel.addEventListener('click', function () { addForm.hidden = true; });

  initTable();
  loadBilling();          // fill the Est. Billing column after first paint
  loadTodayKPIs();        // fill the Today KPI strip
  wireFacDrawer();        // KPI cards → facility (AHD/BLR/North) drawer
  renderPills();
})();
