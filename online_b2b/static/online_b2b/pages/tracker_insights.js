/* Tracker → Insights charts (SELF-CONTAINED, REMOVABLE).
   Renders 4 ECharts from b2b_tracker_insights, honoring the current tracker
   filters. Collapsible (remembered) + browser-cached (instant re-open, then a
   background refresh). Does NOT touch the tracker table/KPIs/drawer — it only
   reads the shared filter bar and listens for a 'trk:filterchange' hint. */
(function () {
  var panel = document.getElementById('tiPanel');
  if (!panel) return;
  var toggle = document.getElementById('tiToggle');
  var bodyEl = document.getElementById('tiBody');
  var emptyEl = document.getElementById('tiEmpty');
  var filter = document.getElementById('trkFilter');
  var base = location.pathname;
  var appRoot = base.replace(/tracker\/?$/, '');
  var url = appRoot + 'tracker/insights/';
  var OPEN_KEY = 'trk_insights_open', CACHE_KEY = 'trk_insights_c_';
  var charts = {}, metric = 'count', lastData = null;

  // ── filters (read the shared bar; date is the client's local day) ──────
  function localToday() {
    var d = new Date();
    return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
  }
  function params() {
    var p = new URLSearchParams();
    ['segment', 'marketplace', 'warehouse', 'q', 'uploaded_from', 'uploaded_to'].forEach(function (n) {
      var el = filter && filter.querySelector('[name="' + n + '"]');
      if (el && el.value) p.set(n, el.value);
    });
    p.set('d', localToday());
    return p;
  }
  function sig() { return params().toString(); }

  // ── browser cache (instant re-open) ────────────────────────────────────
  function cacheGet(s) { try { var r = localStorage.getItem(CACHE_KEY + s); return r ? JSON.parse(r) : null; } catch (e) { return null; } }
  function cacheSet(s, d) { try { localStorage.setItem(CACHE_KEY + s, JSON.stringify(d)); } catch (e) { } }

  function themeColors() {
    var cs = getComputedStyle(document.body);
    var v = function (n, fb) { return (cs.getPropertyValue(n) || '').trim() || fb; };
    return {
      accent: v('--accent', '#4f46e5'), text: v('--text', '#0f172a'),
      text2: v('--text-2', '#64748b'), border: v('--border', '#e6e8f0'),
      surface: v('--surface', '#ffffff'), off: '#11998e',
      green: '#10b981', red: v('--red', '#e11d48'), amber: '#f59e0b'
    };
  }
  function chart(id) {
    if (!window.echarts) return null;
    var el = document.getElementById(id);
    if (!el) return null;
    if (!charts[id]) charts[id] = window.echarts.init(el);
    return charts[id];
  }
  var BASE = { grid: { left: 6, right: 10, top: 12, bottom: 4, containLabel: true }, animationDuration: 480 };
  function inrCr(v) {                       // 23.53 Cr · 2.00 Lakh · 1,200 Rs
    v = Number(v) || 0;
    if (v >= 1e7) return (v / 1e7).toFixed(2) + ' Cr';
    if (v >= 1e5) return (v / 1e5).toFixed(2) + ' Lakh';
    return Math.round(v).toLocaleString('en-IN') + ' Rs';
  }

  function render(d) {
    lastData = d;
    if (!window.echarts) return;
    var c = themeColors();
    var hasAny = (d.marketplaces && d.marketplaces.length) || (d.facilities && d.facilities.length);
    if (emptyEl) emptyEl.hidden = hasAny;
    // 1) daily trend — stacked area by dept
    var tr = chart('tiTrend');
    if (tr) {
      var dl = (d.daily && d.daily.labels) || [], ser = (d.daily && d.daily.series) || {};
      var pick = function (code) { return ((ser[code] || {})[metric]) || []; };
      var fmtLbl = dl.map(function (s) { return s.slice(5); });   // MM-DD
      tr.setOption({
        grid: { left: 6, right: 12, top: 14, bottom: 4, containLabel: true },
        tooltip: {
          trigger: 'axis', valueFormatter: (metric === 'value' ? function (v) { return inrCr(v); } : null)
        }, animationDuration: 480,
        legend: { data: ['Online B2B', 'Offline'], top: 0, right: 0, itemWidth: 10, itemHeight: 10, textStyle: { color: c.text2, fontSize: 10 } },
        xAxis: { type: 'category', data: fmtLbl, axisLine: { lineStyle: { color: c.border } }, axisLabel: { color: c.text2, fontSize: 9.5, interval: 4 }, axisTick: { show: false } },
        yAxis: { type: 'value', splitLine: { lineStyle: { color: c.border, opacity: .5 } }, axisLabel: { color: c.text2, fontSize: 9.5, formatter: (metric === 'value' ? function (v) { return inrCr(v); } : '{value}') } },
        series: [
          { name: 'Online B2B', type: 'line', stack: 't', smooth: true, symbol: 'none', areaStyle: { opacity: .28 }, lineStyle: { width: 2 }, itemStyle: { color: c.accent }, data: pick('OnlineB2B') },
          { name: 'Offline', type: 'line', stack: 't', smooth: true, symbol: 'none', areaStyle: { opacity: .28 }, lineStyle: { width: 2 }, itemStyle: { color: c.off }, data: pick('Offline') }
        ]
      }, true);
    }
    // 2) marketplaces — horizontal bar (top 8, orders)
    var mk = chart('tiMkt');
    if (mk) {
      var mm = (d.marketplaces || []).slice(0, 8).reverse();
      mk.setOption({
        grid: { left: 6, right: 30, top: 6, bottom: 4, containLabel: true }, animationDuration: 480,
        tooltip: { trigger: 'item', formatter: function (p) { var x = mm[p.dataIndex]; return x.name + ' · ' + (x.dept === 'Offline' ? 'Offline' : 'Online B2B') + ': <b>' + x.count + '</b>'; } },
        xAxis: { type: 'value', splitLine: { show: false }, axisLabel: { show: false }, axisLine: { show: false } },
        yAxis: { type: 'category', data: mm.map(function (x) { return x.name; }), axisLabel: { color: c.text2, fontSize: 10 }, axisLine: { show: false }, axisTick: { show: false } },
        series: [{ type: 'bar', barWidth: '62%', label: { show: true, position: 'right', color: c.text2, fontSize: 10, formatter: '{c}' },
          data: mm.map(function (x) { return { value: x.count, itemStyle: { color: x.dept === 'Offline' ? c.off : c.accent, borderRadius: [0, 4, 4, 0] } }; }) }]
      }, true);
    }
    // 3) facility load — donut (AHD/BLR/North)
    var fc = chart('tiFac');
    if (fc) {
      var fpal = { AHD: c.accent, BLR: c.off, North: c.amber };
      fc.setOption({
        animationDuration: 480, tooltip: { trigger: 'item', formatter: '{b}: {c} ({d}%)' },
        legend: { bottom: 0, left: 'center', itemWidth: 9, itemHeight: 9, textStyle: { color: c.text2, fontSize: 10 } },
        series: [{ type: 'pie', radius: ['42%', '68%'], center: ['50%', '44%'], avoidLabelOverlap: true, label: { show: false }, data: (d.facilities || []).map(function (x) { return { name: x.code, value: x.count, itemStyle: { color: fpal[x.code] || c.muted } }; }) }]
      }, true);
    }
    // 4) fill-rate today — donut (billable vs short)
    var fl = chart('tiFill');
    if (fl) {
      var b = (d.fill && d.fill.billable) || 0, sh = (d.fill && d.fill.short) || 0;
      fl.setOption({
        animationDuration: 480,
        tooltip: { trigger: 'item', formatter: function (p) { return p.name + ': ' + inrCr(p.value) + ' (' + p.percent + '%)'; } },
        legend: { bottom: 0, left: 'center', itemWidth: 9, itemHeight: 9, textStyle: { color: c.text2, fontSize: 10 } },
        series: [{ type: 'pie', radius: ['48%', '70%'], center: ['50%', '44%'], label: { show: false }, data: [{ name: 'Billable', value: b, itemStyle: { color: c.green } }, { name: 'Short', value: sh, itemStyle: { color: c.red } }] }]
      }, true);
    }
    // 5) arrival pattern — marketplace × weekday heatmap (predict arrivals)
    var ar = chart('tiArrival');
    if (ar) {
      var A = d.arrival || { markets: [], dow: [], data: [], max: 0 };
      ar.setOption({
        animationDuration: 480,
        tooltip: { position: 'top', formatter: function (p) { return A.markets[p.value[1]] + ' · ' + A.dow[p.value[0]] + ': <b>' + p.value[2] + '%</b> chance of an order'; } },
        grid: { left: 6, right: 12, top: 6, bottom: 8, containLabel: true },
        xAxis: { type: 'category', data: A.dow, splitArea: { show: true }, axisLabel: { color: c.text2, fontSize: 10 }, axisLine: { show: false }, axisTick: { show: false } },
        yAxis: { type: 'category', data: A.markets, splitArea: { show: true }, axisLabel: { color: c.text2, fontSize: 10 }, axisLine: { show: false }, axisTick: { show: false } },
        visualMap: { min: 0, max: 100, calculable: false, show: false, inRange: { color: [c.surface, c.accent] } },
        series: [{ type: 'heatmap', data: A.data, label: { show: true, formatter: function (p) { return p.value[2] ? p.value[2] + '%' : ''; }, fontSize: 9.5, fontWeight: 600, color: '#1e293b' }, itemStyle: { borderColor: c.surface, borderWidth: 1.5 }, emphasis: { itemStyle: { borderColor: c.text } } }]
      }, true);
    }
    resize();
  }
  function resize() { Object.keys(charts).forEach(function (k) { if (charts[k]) charts[k].resize(); }); }

  var inflight = false;
  function load(force) {
    if (bodyEl.hidden) return;                 // only when open
    var s = sig(), cached = cacheGet(s);
    if (cached && !force) render(cached);      // instant paint from cache
    if (inflight) return;
    inflight = true;
    fetch(url + '?' + s, { headers: { 'X-Requested-With': 'fetch' } })
      .then(function (r) { return r.json(); })
      .then(function (d) { inflight = false; if (d && d.ok) { cacheSet(s, d); render(d); } })
      .catch(function () { inflight = false; });
  }

  // ── open / collapse (remembered) ───────────────────────────────────────
  function setOpen(open) {
    bodyEl.hidden = !open;
    toggle.setAttribute('aria-expanded', open ? 'true' : 'false');
    panel.classList.toggle('open', open);
    try { localStorage.setItem(OPEN_KEY, open ? '1' : '0'); } catch (e) { }
    if (open) load(false);
  }
  toggle.addEventListener('click', function () { setOpen(bodyEl.hidden); });

  // trend metric toggle (orders / value)
  var mseg = document.getElementById('tiTrendMetric');
  if (mseg) mseg.addEventListener('click', function (e) {
    var b = e.target.closest && e.target.closest('button[data-m]');
    if (!b) return;
    metric = b.getAttribute('data-m');
    mseg.querySelectorAll('button').forEach(function (x) { x.classList.toggle('on', x === b); });
    if (lastData) render(lastData);
  });

  // ── expand any chart into a big modal (zoom on cartesian charts) ───────
  var modal = document.getElementById('tiModal'), modalOv = document.getElementById('tiModalOv'),
      modalX = document.getElementById('tiModalX'), modalT = document.getElementById('tiModalT'),
      modalChartEl = document.getElementById('tiModalChart'), modalChart = null;
  Array.prototype.forEach.call(document.querySelectorAll('body > #tiModal, body > #tiModalOv'), function (el) { el.remove(); });
  if (modal) document.body.appendChild(modal);
  if (modalOv) document.body.appendChild(modalOv);
  function openModal(id, title) {
    if (!charts[id] || !window.echarts || !modal) return;
    if (modalT) modalT.textContent = title || 'Chart';
    modal.hidden = false; modalOv.hidden = false;
    if (!modalChart) modalChart = window.echarts.init(modalChartEl);
    var opt = charts[id].getOption();
    if (opt.xAxis && opt.xAxis.length) {
      // scroll/drag zoom + a visible slider; widen the grid so the slider sits
      // below the axis instead of over the labels.
      opt.dataZoom = [{ type: 'inside' }, { type: 'slider', bottom: 14, height: 22 }];
      (Array.isArray(opt.grid) ? opt.grid : [opt.grid || {}]).forEach(function (g) { g.bottom = 64; });
      if (opt.legend) (Array.isArray(opt.legend) ? opt.legend : [opt.legend]).forEach(function (l) { l.top = 4; });
    }
    modalChart.setOption(opt, true);
    setTimeout(function () { modalChart.resize(); }, 30);
  }
  function closeModal() { if (modal) { modal.hidden = true; modalOv.hidden = true; } }
  document.addEventListener('click', function (e) {
    var b = e.target.closest && e.target.closest('.ti-exp');
    if (b) { e.preventDefault(); openModal(b.getAttribute('data-exp'), b.getAttribute('data-title')); }
  });
  if (modalX) modalX.addEventListener('click', closeModal);
  if (modalOv) modalOv.addEventListener('click', closeModal);
  document.addEventListener('keydown', function (e) { if (e.key === 'Escape') closeModal(); });

  // follow filter changes (a hint from tracker.js; harmless if absent) + resize
  document.addEventListener('trk:filterchange', function () { load(true); });
  window.addEventListener('resize', function () { resize(); if (modalChart && modal && !modal.hidden) modalChart.resize(); });

  // restore open state
  var wasOpen = false;
  try { wasOpen = localStorage.getItem(OPEN_KEY) === '1'; } catch (e) { }
  if (wasOpen) setOpen(true);
})();
