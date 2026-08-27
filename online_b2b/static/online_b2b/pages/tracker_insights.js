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
  var charts = {}, metric = 'count', ivMetric = 'qty', lastData = null;

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
      // day granularity → MM-DD; a single selected day comes back HOURLY → format 9a/12p
      var gran = (d.daily && d.daily.gran) || 'day';
      var hourFmt = function (h) { h = +h; var ap = h < 12 ? 'a' : 'p'; var hh = h % 12; if (!hh) hh = 12; return hh + ap; };
      var fmtLbl = dl.map(function (s) { return gran === 'hour' ? hourFmt(s) : s.slice(5); });
      tr.setOption({
        grid: { left: 6, right: 12, top: 14, bottom: 4, containLabel: true },
        tooltip: {
          trigger: 'axis', valueFormatter: (metric === 'value' ? function (v) { return inrCr(v); } : null)
        }, animationDuration: 480,
        legend: { data: ['Online B2B', 'Offline'], top: 0, right: 0, itemWidth: 10, itemHeight: 10, textStyle: { color: c.text2, fontSize: 10 } },
        xAxis: { type: 'category', data: fmtLbl, axisLine: { lineStyle: { color: c.border } }, axisLabel: { color: c.text2, fontSize: 9.5, interval: 4 }, axisTick: { show: false } },
        yAxis: { type: 'value', splitLine: { lineStyle: { color: c.border, opacity: .5 } }, axisLabel: { color: c.text2, fontSize: 9.5, formatter: (metric === 'value' ? function (v) { return inrCr(v); } : '{value}') } },
        series: [
          // Straight segments (no spline). Daily intake legitimately swings 0→peak
          // (weekend zeros), and a smoothed line MUST overshoot/wiggle through those
          // points — that overshoot is what read as a "distorted" waveform, worst
          // when stretched wide in the expand modal. Faithful polyline + clip:true.
          { name: 'Online B2B', type: 'line', stack: 't', clip: true, symbol: 'none', areaStyle: { opacity: .26 }, lineStyle: { width: 2 }, itemStyle: { color: c.accent }, data: pick('OnlineB2B') },
          { name: 'Offline', type: 'line', stack: 't', clip: true, symbol: 'none', areaStyle: { opacity: .26 }, lineStyle: { width: 2 }, itemStyle: { color: c.off }, data: pick('Offline') }
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
    // 6) order timeline — an ACTIVITY timeline: a straight spine with a GLOWING
    //    (rippling) dot for EACH marketplace's arrival, at its exact clock minute
    //    (9:00 AM, 9:14 AM…). Marketplaces are NEVER merged — GT Mass and GT Select
    //    each get their own colored dot even in the same hour. Labels alternate
    //    above/below the spine so close arrivals stay readable. Qty/Value sizes dots.
    var iv = chart('tiIntraday');
    if (iv) {
      var I = d.intraday || { markets: [], points: [] };
      var ivVal = (ivMetric === 'value');
      var minFmt = function (m) {
        m = Math.round(+m); var mi = ((m % 60) + 60) % 60; var h = ((Math.floor(m / 60) % 24) + 24) % 24;
        var ap = h < 12 ? 'AM' : 'PM'; var hh = h % 12; if (!hh) hh = 12;
        return hh + ':' + ('0' + mi).slice(-2) + ' ' + ap;
      };
      // one dot per (marketplace, minute) — each MP its own color, sorted along the day
      var pts = (I.points || []).slice().sort(function (a, b) { return (a.min - b.min) || (a.mi - b.mi); });
      var palette = [c.accent, '#22c55e', '#f59e0b', '#ec4899', '#06b6d4', '#a855f7', '#ef4444', '#14b8a6', '#eab308', '#6366f1'];
      var mColor = {};
      (I.markets || []).forEach(function (m, i) { mColor[m] = palette[i % palette.length]; });
      var maxM = 1;
      pts.forEach(function (p) { var m = ivVal ? p.value : p.qty; if (m > maxM) maxM = m; });
      var mins = pts.map(function (p) { return p.min; });
      var lo = mins.length ? Math.max(0, Math.floor((Math.min.apply(null, mins) - 30) / 60) * 60) : 480;
      var hi = mins.length ? Math.min(1440, Math.ceil((Math.max.apply(null, mins) + 30) / 60) * 60) : 1080;
      if (hi - lo < 120) { hi = Math.min(1440, lo + 120); lo = Math.max(0, hi - 120); }  // min span, clamped within the day (no phantom post-midnight ticks)
      iv.setOption({
        animationDuration: 480,
        grid: { left: 6, right: 16, top: 40, bottom: 36, containLabel: true },
        tooltip: {
          trigger: 'item', formatter: function (p) {
            var a = p.data.a;
            return '<b>' + a.mp + '</b> · ' + minFmt(a.min) + '<br/>' +
              a.orders + ' order' + (a.orders === 1 ? '' : 's') + ' · ' +
              a.qty.toLocaleString('en-IN') + ' qty · <b>' + inrCr(a.value) + '</b>';
          }
        },
        dataZoom: [{ type: 'inside', filterMode: 'none', xAxisIndex: 0 }],
        xAxis: {
          type: 'value', min: lo, max: hi, interval: 60, axisTick: { show: false },
          axisLine: { lineStyle: { color: c.border } },
          axisLabel: { color: c.text2, fontSize: 9.5, formatter: minFmt }, splitLine: { show: false }
        },
        yAxis: { type: 'value', min: -1, max: 1, show: false },
        series: [
          { type: 'line', z: 1, silent: true, showSymbol: false,
            lineStyle: { color: c.border, width: 2 }, data: [[lo, 0], [hi, 0]] },
          { type: 'effectScatter', z: 2, showEffectOn: 'render',
            rippleEffect: { scale: 2.4, brushType: 'stroke' },
            symbolSize: function (v) { return 11 + Math.sqrt((v[2] || 0) / maxM) * 18; },
            data: pts.map(function (p, i) {
              var col = mColor[p.mp] || c.accent;
              var val = ivVal ? ('₹' + inrCr(p.value)) : (p.qty.toLocaleString('en-IN') + ' qty');
              return {
                value: [p.min, 0, ivVal ? p.value : p.qty], a: p,
                itemStyle: { color: col, shadowBlur: 14, shadowColor: col },
                label: {
                  show: true, position: (i % 2 === 0) ? 'top' : 'bottom', distance: 12,
                  formatter: '{m|' + p.mp + '}  {t|' + minFmt(p.min) + '}\n{v|' + val + '}',
                  rich: {
                    m: { fontSize: 10.5, fontWeight: 800, color: col },
                    t: { fontSize: 9, color: c.text2 },
                    v: { fontSize: 11.5, fontWeight: 800, color: c.text, align: 'center', padding: [2, 0, 0, 0] }
                  }
                }
              };
            }) }
        ]
      }, true);
    }
    bodyEl.classList.remove('ti-loading');   // charts painted → drop the skeleton
    resize();
    requestAnimationFrame(resize);           // container may have JUST laid out
    updateDateRange(d);
  }
  function resize() { Object.keys(charts).forEach(function (k) { if (charts[k]) charts[k].resize(); }); }

  // Show WHICH date window the charts cover (the filter range, else the trend's
  // own last-30-days span).
  function updateDateRange(d) {
    var el = document.getElementById('tiDateRange');
    if (!el) return;
    var f = params(), from = f.get('uploaded_from'), to = f.get('uploaded_to');
    var lbls = (d && d.daily && d.daily.labels) || [];
    var txt;
    if (from || to) txt = (from || '…') + '  →  ' + (to || '…');
    else if (lbls.length) txt = lbls[0] + '  →  ' + lbls[lbls.length - 1] + ' · last 30d';
    else txt = 'last 30 days';
    el.textContent = '📅 ' + txt;
    // keep the trend card's own note honest (a single day comes back hourly)
    var note = document.getElementById('tiTrendNote');
    if (note) {
      var single = from && to && from === to;
      note.textContent = (single ? 'hourly · ' + from : (from || to ? 'selected range' : 'last 30 days')) + ' · dept-split';
    }
  }

  var inflight = false;
  function load(force) {
    if (!panel.classList.contains('open')) return;   // only when open
    var s = sig(), cached = cacheGet(s);
    if (cached && !force) render(cached);      // instant paint from cache
    else if (!lastData) bodyEl.classList.add('ti-loading');   // nothing yet → skeleton (no blank white)
    if (inflight) return;
    inflight = true;
    fetch(url + '?' + s, { headers: { 'X-Requested-With': 'fetch' } })
      .then(function (r) { return r.json(); })
      .then(function (d) { inflight = false; bodyEl.classList.remove('ti-loading');
        if (d && d.ok) { cacheSet(s, d); render(d); } })
      .catch(function () { inflight = false; bodyEl.classList.remove('ti-loading'); });
  }

  // ── open / collapse (remembered) — SMOOTH accordion, not a hard snap ─────
  var REDUCE = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  function setOpen(open, instant) {
    panel.classList.toggle('open', open);
    toggle.setAttribute('aria-expanded', open ? 'true' : 'false');
    try { localStorage.setItem(OPEN_KEY, open ? '1' : '0'); } catch (e) { }
    clearTimeout(bodyEl._t);
    if (open) {
      bodyEl.hidden = false;
      load(false);                                   // charts fill the fixed-height cells
      if (instant || REDUCE) { bodyEl.style.maxHeight = 'none'; bodyEl.style.opacity = '1'; resize(); return; }
      bodyEl.style.maxHeight = '0px'; bodyEl.style.opacity = '0';
      void bodyEl.offsetHeight;                       // commit the 0 state, then grow to real height
      bodyEl.style.maxHeight = bodyEl.scrollHeight + 'px'; bodyEl.style.opacity = '1';
      bodyEl._t = setTimeout(function () {            // release the cap so filter changes can resize freely
        if (panel.classList.contains('open')) { bodyEl.style.maxHeight = 'none'; resize(); }
      }, 360);
    } else {
      if (instant || REDUCE) { bodyEl.hidden = true; bodyEl.style.maxHeight = ''; bodyEl.style.opacity = ''; return; }
      bodyEl.style.maxHeight = bodyEl.scrollHeight + 'px'; bodyEl.style.opacity = '1';
      void bodyEl.offsetHeight;
      bodyEl.style.maxHeight = '0px'; bodyEl.style.opacity = '0';
      bodyEl._t = setTimeout(function () {
        if (!panel.classList.contains('open')) { bodyEl.hidden = true; bodyEl.style.maxHeight = ''; bodyEl.style.opacity = ''; }
      }, 360);
    }
  }
  toggle.addEventListener('click', function () { setOpen(!panel.classList.contains('open')); });

  // trend metric toggle (orders / value)
  var mseg = document.getElementById('tiTrendMetric');
  if (mseg) mseg.addEventListener('click', function (e) {
    var b = e.target.closest && e.target.closest('button[data-m]');
    if (!b) return;
    metric = b.getAttribute('data-m');
    mseg.querySelectorAll('button').forEach(function (x) { x.classList.toggle('on', x === b); });
    if (lastData) render(lastData);
  });

  // order-timeline metric toggle (qty / value) — same pattern as the trend toggle
  var iseg = document.getElementById('tiIntradayMetric');
  if (iseg) iseg.addEventListener('click', function (e) {
    var b = e.target.closest && e.target.closest('button[data-m]');
    if (!b) return;
    ivMetric = b.getAttribute('data-m');
    iseg.querySelectorAll('button').forEach(function (x) { x.classList.toggle('on', x === b); });
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
    // Power-BI-style interactivity in the big modal: mouse-wheel ZOOM + drag to PAN
    // via 'inside' dataZoom (no visible scrollbar — the slider read as clutter). Both
    // axes for a scatter (category y + value x); x-only for the time-series bars/lines.
    var ya = opt.yAxis && opt.yAxis[0];
    var zooms = [{ type: 'inside', xAxisIndex: 0, filterMode: 'none' }];
    if (ya && ya.type === 'category') zooms.push({ type: 'inside', yAxisIndex: 0, filterMode: 'none' });
    opt.dataZoom = zooms;
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

  // ECharts measures a chart's container ONCE at init; a 0-size div (first paint,
  // or a just-expanded panel) makes it render blank until the next resize — the
  // "shows only on the 2nd load" bug. Observe the body so every chart snaps to its
  // real size the instant it gets one. Debounced.
  if (window.ResizeObserver) {
    var _roT;
    new ResizeObserver(function () { clearTimeout(_roT); _roT = setTimeout(resize, 60); }).observe(bodyEl);
  }

  // Insights starts COLLAPSED on every load (per request) — the panel's initial
  // markup is already closed, so we simply DON'T auto-open it. Deferring chart init
  // to an explicit user-open also structurally kills the blank-on-first-load race:
  // echarts no longer measures a 0-size container mid nav-transition (the "only
  // shows on refresh" bug). Charts init the moment the panel is opened, layout settled.
})();
