/* online_b2b/analytics.html — page script (separated from template). */
(function () {
  var AN_URL = location.pathname;
  var activeTab = 'daily';

  function url(params) {
    var p = new URLSearchParams();
    Object.keys(params).forEach(function (k) {
      if (params[k] !== undefined && params[k] !== null) p.set(k, params[k]);
    });
    return AN_URL + '?' + p.toString();
  }

  // keep the address bar in sync so a refresh restores tab + both filters
  function syncURL() {
    var p = new URLSearchParams();
    p.set('tab', activeTab);
    var d = document.querySelector('#pane-daily .an-daily');
    if (d) {
      if (d.dataset.start && d.dataset.end) { p.set('start', d.dataset.start); p.set('end', d.dataset.end); }
      else { p.set('days', d.dataset.days || 30); }
    }
    // SKU + Fulfilment share the same param names — reflect only the active one
    var s = null;
    if (activeTab === 'sku') s = document.querySelector('#pane-sku .an-sku');
    else if (activeTab === 'fulfil') s = document.querySelector('#pane-fulfil .an-fulfil');
    else if (activeTab === 'exc') s = document.querySelector('#pane-exc .an-exc');
    else if (activeTab === 'geo') s = document.querySelector('#pane-geo .an-geo');
    else if (activeTab === 'otif') {
      var ot = document.querySelector('#pane-otif .an-otif');
      if (ot) { if (ot.dataset.mp) p.set('sku_mp', ot.dataset.mp); if (ot.dataset.horizon && ot.dataset.horizon !== '0') p.set('horizon', ot.dataset.horizon); }
    }
    else if (activeTab === 'dos') {
      var ds = document.querySelector('#pane-dos .an-dos');
      if (ds) { if (ds.dataset.mp) p.set('sku_mp', ds.dataset.mp); p.set('days', ds.dataset.days || 30); }
    }
    if (s) {
      if (s.dataset.mp) p.set('sku_mp', s.dataset.mp);
      if (s.dataset.from) p.set('sku_from', s.dataset.from);
      if (s.dataset.to) p.set('sku_to', s.dataset.to);
    }
    if (activeTab === 'geo' && s && s.dataset.seg) p.set('geo_seg', s.dataset.seg);   // segment filter
    history.replaceState(null, '', AN_URL + '?' + p.toString());
  }

  // ── animated multi-step loader (for the slower tabs) ──
  // Cycles through `steps` so a multi-second fetch feels like real progress.
  // Returns the interval id — clear it once the response lands.
  function showLoader(pane, steps, sub) {
    var i = 0;
    pane.innerHTML =
      '<div class="an-loader"><div class="an-spin"></div>' +
      '<div class="an-lmsg"></div>' +
      '<div class="an-lsub">' + (sub || '') + '</div></div>';
    var el = pane.querySelector('.an-lmsg');
    if (el) el.textContent = steps[0];
    return setInterval(function () {
      i = (i + 1) % steps.length;
      if (!el) return;
      el.style.opacity = '0';
      setTimeout(function () { el.textContent = steps[i]; el.style.opacity = '1'; }, 160);
    }, 1200);
  }

  // ── generic click-to-sort for a table.sortable inside `root` ──
  function wireSort(root) {
    root.querySelectorAll('table.sortable').forEach(function (table) {
      table.querySelectorAll('th.srt').forEach(function (th) {
        th.addEventListener('click', function () {
          var k = th.getAttribute('data-k');
          var asc = th.getAttribute('data-dir') !== 'asc';
          table.querySelectorAll('th.srt').forEach(function (x) { x.removeAttribute('data-dir'); x.classList.remove('on'); });
          th.setAttribute('data-dir', asc ? 'asc' : 'desc'); th.classList.add('on');
          var tb = table.tBodies[0];
          Array.prototype.slice.call(tb.rows).sort(function (a, b) {
            var ae = a.querySelector('[data-k="' + k + '"]'), be = b.querySelector('[data-k="' + k + '"]');
            var av = parseFloat(ae ? ae.getAttribute('data-v') : 0) || 0;
            var bv = parseFloat(be ? be.getAttribute('data-v') : 0) || 0;
            return asc ? av - bv : bv - av;
          }).forEach(function (r) { tb.appendChild(r); });
        });
      });
    });
  }

  // ══ DAILY tab ══
  function buildDailyChart() {
    var el = document.getElementById('daily-data');
    if (!el) return;
    var data; try { data = JSON.parse(el.textContent); } catch (e) { return; }
    function inr(v) {
      v = Number(v) || 0; var s = v < 0 ? '-' : ''; v = Math.abs(v);
      if (v >= 1e7) return s + '₹' + (v / 1e7).toFixed(2) + ' Cr';
      if (v >= 1e5) return s + '₹' + (v / 1e5).toFixed(2) + ' L';
      if (v >= 1000) return s + '₹' + Math.round(v).toLocaleString('en-IN');
      return s + '₹' + Math.round(v);
    }
    function num(v) {
      v = Number(v) || 0;
      if (v >= 1e5) return (v / 1e5).toFixed(1) + 'L';
      if (v >= 1000) return (v / 1000).toFixed(1) + 'k';
      return Math.round(v);
    }
    var COLORS = { 'Online B2B': 'var(--accent)', 'Offline': '#11998e', 'Other': '#9aa1b2' };
    var metric = 'value';
    var chart = null;
    function fmt(v) { return metric === 'value' ? inr(v) : num(v); }
    function series() { return data.segments.map(function (s) { return { name: s, data: (data[metric] || {})[s] || [] }; }); }
    var box = document.getElementById('daily-chart');
    if (!window.ApexCharts) { box.textContent = 'Charts library unavailable.'; return; }
    if (!data.segments || !data.segments.length) { box.innerHTML = '<div class="chart-empty">No orders in this period.</div>'; return; }
    box.innerHTML = '';
    var cfg = {
      chart: { type: 'bar', height: 380, stacked: true, fontFamily: 'Inter, sans-serif', toolbar: { show: false }, animations: { enabled: true, speed: 600, dynamicAnimation: { enabled: true, speed: 450 } } },
      series: series(),
      colors: data.segments.map(function (s) { return COLORS[s] || '#9aa1b2'; }),
      plotOptions: { bar: { columnWidth: '62%', borderRadius: 3 } },
      dataLabels: { enabled: false },
      xaxis: { categories: data.labels, axisBorder: { show: false }, axisTicks: { show: false }, labels: { style: { colors: '#9aa1b2', fontSize: '10px' } } },
      yaxis: { labels: { style: { colors: '#9aa1b2', fontSize: '10px' }, formatter: function (v) { return fmt(v); } } },
      legend: { position: 'top', horizontalAlign: 'right', fontSize: '12px', labels: { colors: '#9aa1b2' } },
      grid: { borderColor: '#eef1f5', strokeDashArray: 4 },
      tooltip: { y: { formatter: function (v) { return fmt(v); } } }
    };
    if (data.focus) {
      cfg.fill = { opacity: 0.32 };
      cfg.annotations = { xaxis: [{ x: data.focus, borderColor: 'var(--accent)', strokeDashArray: 0,
        label: { text: '● ' + data.focus, orientation: 'horizontal', position: 'top',
          style: { background: 'var(--accent)', color: '#fff', fontSize: '10px', fontWeight: 700, padding: { left: 7, right: 7, top: 3, bottom: 3 } } } }] };
    }
    chart = new ApexCharts(box, cfg);
    chart.render();
    document.querySelectorAll('.ctg').forEach(function (b) {
      b.addEventListener('click', function () {
        document.querySelectorAll('.ctg').forEach(function (x) { x.classList.toggle('on', x === b); });
        metric = b.getAttribute('data-metric');
        chart.updateSeries(series(), true);
      });
    });
  }

  function loadDaily(params) {
    var pane = document.getElementById('pane-daily');
    pane.classList.add('an-busy');
    fetch(url(Object.assign({ partial: 'daily' }, params)), { headers: { 'X-Requested-With': 'fetch' } })
      .then(function (r) { return r.text(); })
      .then(function (html) { pane.innerHTML = html; pane.classList.remove('an-busy'); wireDaily(); syncURL(); })
      .catch(function () { pane.classList.remove('an-busy'); });
  }

  // ── Breakdown views: Tree · Sunburst · Treemap · metric switch · growth colour ──
  var BD = { view: 'tree', metric: 'value', color: 'seg' };
  function bdColors() {                              // ECharts can't read CSS vars → resolve --accent to a real hex
    var acc = (getComputedStyle(document.body).getPropertyValue('--accent') || '').trim() || '#4f46e5';
    return [acc, '#11998e', '#f7971e', '#cb2d3e', '#2193b0', '#7b4397', '#16a34a', '#db2777'];
  }
  function bdInr(v) {
    v = Number(v) || 0; var s = v < 0 ? '-' : ''; v = Math.abs(v);
    if (v >= 1e7) return s + '₹' + (v / 1e7).toFixed(2) + ' Cr';
    if (v >= 1e5) return s + '₹' + (v / 1e5).toFixed(2) + ' L';
    if (v >= 1000) return s + '₹' + Math.round(v).toLocaleString('en-IN');
    return s + '₹' + Math.round(v);
  }
  function bdNum(v) {
    v = Number(v) || 0;
    if (v >= 1e5) return (v / 1e5).toFixed(2) + ' L';
    if (v >= 1000) return (v / 1000).toFixed(1) + 'k';
    return Math.round(v).toLocaleString('en-IN');
  }
  function bdFmt(v) { return BD.metric === 'value' ? bdInr(v) : bdNum(v); }
  function growthColor(g) {                          // green ↑ · red ↓ · grey new
    if (g === null || g === undefined) return '#c7ccd6';
    if (g >= 25) return '#059669'; if (g >= 5) return '#34d399';
    if (g <= -25) return '#dc2626'; if (g <= -5) return '#f87171';
    return '#cbd5e1';
  }
  function bdData() {
    var el = document.getElementById('bd-data'); if (!el) return null;
    try { return JSON.parse(el.textContent); } catch (e) { return null; }
  }
  function toE(nodes) {                               // raw → ECharts data for current metric + colour
    return nodes.map(function (n) {
      var out = { name: n.name };
      if (n.children) { out.children = toE(n.children); }
      else {
        out.value = (n.m && n.m[BD.metric]) || 0;
        out.mp = n.mp; out.growth = n.growth;
        if (BD.color === 'growth') out.itemStyle = { color: growthColor(n.growth) };
      }
      return out;
    });
  }
  function bdTooltip(p) {
    var g = p.data && p.data.growth;
    var gt = (g === null || g === undefined) ? '' :
      '<br/><span style="color:' + (g >= 0 ? '#059669' : '#dc2626') + '">' + (g >= 0 ? '▲ ' : '▼ ') + Math.abs(g) + '% vs prev</span>';
    return '<b>' + p.name + '</b><br/>' + bdFmt(p.value) + gt;
  }
  var BD_TOOLBOX = { feature: { saveAsImage: { title: 'Save PNG', name: 'breakdown', pixelRatio: 2 } }, right: 8, top: 2 };
  function bdChart(box) {                             // fresh ECharts instance for a container
    if (box._chart) { try { box._chart.dispose(); } catch (e) { } }
    box._chart = echarts.init(box); return box._chart;
  }
  function bdReady(box) {
    if (!box) return null;
    var raw = bdData();
    if (!window.echarts) { box.innerHTML = '<div class="chart-empty">Charts library unavailable.</div>'; return null; }
    if (!raw || !raw.length) { box.innerHTML = '<div class="chart-empty">No orders in this period.</div>'; return null; }
    return raw;
  }
  function bdClick(chart, box) {
    var url = box.getAttribute('data-orders-url') || '';
    chart.off('click');
    chart.on('click', function (p) {
      if (p.data && p.data.mp && url) window.location.href = url + '?marketplace=' + encodeURIComponent(p.data.mp);
    });
  }
  function buildSunburst() {
    var box = document.getElementById('an-sunburst'); var raw = bdReady(box); if (!raw) return;
    var chart = bdChart(box);
    chart.setOption({
      tooltip: { formatter: bdTooltip }, toolbox: BD_TOOLBOX,
      color: BD.color === 'growth' ? ['#94a3b8'] : bdColors(),
      series: [{
        type: 'sunburst', data: toE(raw), radius: ['12%', '98%'],
        emphasis: { focus: 'ancestor' },
        label: { rotate: 'radial', minAngle: 8, fontSize: 11, color: '#fff' },
        levels: [
          {},
          { r0: '12%', r: '44%', itemStyle: { borderWidth: 2 }, label: { fontWeight: 700, fontSize: 12 } },
          { r0: '44%', r: '73%', label: { fontSize: 11 } },
          { r0: '73%', r: '98%', label: { fontSize: 10 }, itemStyle: { borderWidth: 1 } }
        ]
      }]
    }, true);
    bdClick(chart, box);
  }
  function buildTreemap() {
    var box = document.getElementById('an-treemap'); var raw = bdReady(box); if (!raw) return;
    var chart = bdChart(box);
    var shade = BD.color === 'growth' ? [] : [{ colorSaturation: [.35, .55] }, { colorSaturation: [.3, .5] }];
    chart.setOption({
      tooltip: { formatter: bdTooltip }, toolbox: BD_TOOLBOX,
      color: BD.color === 'growth' ? ['#94a3b8'] : bdColors(),
      series: [{
        type: 'treemap', data: toE(raw), roam: false, nodeClick: 'zoomToNode', drillDownIcon: '▸',
        top: 30, left: 2, right: 2, bottom: 2,   // fill the container (override ECharts' default 80% centered → no gap)
        breadcrumb: {
          show: true, top: 4, left: 'center', height: 22, emptyItemWidth: 26,
          itemStyle: {
            color: '#eef1f5', borderColor: '#e3e6ee', borderWidth: 1, gapWidth: 3,
            textStyle: { color: '#5b6478', fontSize: 11 }
          },
          emphasis: { itemStyle: { color: '#e2e8f0' } }
        },
        itemStyle: { borderColor: '#fff', gapWidth: 2, borderWidth: 2 },
        levels: [{ itemStyle: { gapWidth: 4, borderWidth: 0 } }].concat(shade),
        label: { show: true, formatter: '{b}', fontSize: 12 },
        upperLabel: { show: true, height: 24, fontSize: 12, fontWeight: 600, color: '#fff' }
      }]
    }, true);
    bdClick(chart, box);
  }
  function bdRenderActive() { if (BD.view === 'sun') buildSunburst(); else if (BD.view === 'map') buildTreemap(); }
  function wireBreakdown() {
    var views = [].slice.call(document.querySelectorAll('.an-bd-btn')); if (!views.length) return;
    BD.view = 'tree'; BD.metric = 'value'; BD.color = 'seg';          // reset each partial load
    var tree = document.querySelector('.tree[data-bd-pane="tree"]');
    var sun = document.getElementById('an-sunburst'), map = document.getElementById('an-treemap');
    var mBox = document.querySelector('.an-bd-metric'), cBox = document.querySelector('.an-bd-color');
    var reset = document.getElementById('an-sun-reset');
    function paint() {
      if (tree) tree.hidden = BD.view !== 'tree';
      if (sun) sun.hidden = BD.view !== 'sun';
      if (map) map.hidden = BD.view !== 'map';
      var chartView = BD.view !== 'tree';
      if (mBox) mBox.hidden = !chartView;
      if (cBox) cBox.hidden = !chartView;
      if (reset) reset.hidden = BD.view !== 'sun';
      bdRenderActive();
    }
    views.forEach(function (b) {
      b.addEventListener('click', function () {
        views.forEach(function (x) { x.classList.toggle('on', x === b); });
        BD.view = b.getAttribute('data-bd'); paint();
      });
    });
    function segToggle(sel, key) {
      var bs = [].slice.call(document.querySelectorAll(sel));
      bs.forEach(function (b) {
        b.addEventListener('click', function () {
          bs.forEach(function (x) { x.classList.toggle('on', x === b); });
          BD[key] = b.getAttribute('data-' + key); bdRenderActive();
        });
      });
    }
    segToggle('.an-bd-metric .an-sbtn', 'metric');
    segToggle('.an-bd-color .an-sbtn', 'color');
    if (reset) reset.addEventListener('click', function () { buildSunburst(); });  // rebuild = zoom back to root
  }
  if (!window.__bdResizeBound) {                     // resize the visible breakdown chart (bound once)
    window.__bdResizeBound = true;
    window.addEventListener('resize', function () {
      ['an-sunburst', 'an-treemap'].forEach(function (id) {
        var b = document.getElementById(id);
        if (b && b._chart && !b.hidden) { try { b._chart.resize(); } catch (e) { } }
      });
    });
  }

  function wireDaily() {
    buildDailyChart();
    wireBreakdown();
    var root = document.querySelector('#pane-daily .an-daily');
    if (!root) return;
    var days = root.dataset.days || 30;
    // quick presets — a preset always clears any active range
    root.querySelectorAll('[data-daily-days]').forEach(function (a) {
      a.addEventListener('click', function (e) { e.preventDefault(); loadDaily({ days: a.getAttribute('data-daily-days') }); });
    });
    // custom from–to range
    var s = root.querySelector('[data-daily-start]'), en = root.querySelector('[data-daily-end]');
    var apply = root.querySelector('[data-daily-apply]');
    function doApply() {
      if (s.value && en.value) loadDaily({ start: s.value, end: en.value });
      else (s.value ? en : s).focus();               // need both ends
    }
    if (apply) apply.addEventListener('click', doApply);
    // auto-apply once the second date is filled
    if (s) s.addEventListener('change', function () { if (s.value && en.value) doApply(); });
    if (en) en.addEventListener('change', function () { if (s.value && en.value) doApply(); });
    var cl = root.querySelector('[data-daily-clear]');
    if (cl) cl.addEventListener('click', function (e) { e.preventDefault(); loadDaily({ days: days }); });
  }

  // ══ SKU tab ══
  function loadSku(params) {
    var pane = document.getElementById('pane-sku');
    var lid = showLoader(pane, [
      'Reading uploaded PO lines…',
      'Rolling up demand per SKU…',
      'Summing qty & value…',
      'Sorting the SKU list…'
    ], 'Aggregating every SKU across the selected uploads.');
    fetch(url(Object.assign({ partial: 'sku' }, params)), { headers: { 'X-Requested-With': 'fetch' } })
      .then(function (r) { return r.text(); })
      .then(function (html) { clearInterval(lid); pane.innerHTML = html; pane.dataset.loaded = '1'; wireSku(); syncURL(); })
      .catch(function () { clearInterval(lid); pane.innerHTML = '<div class="an-loading">Could not load SKU demand.</div>'; });
  }

  function wireSku() {
    var root = document.querySelector('#pane-sku .an-sku');
    if (!root) return;
    wireSort(root);
    var form = root.querySelector('[data-sku-filter]');
    function apply() {
      loadSku({
        sku_mp: form.querySelector('[name=sku_mp]').value,
        sku_from: form.querySelector('[name=sku_from]').value,
        sku_to: form.querySelector('[name=sku_to]').value
      });
    }
    if (form) form.addEventListener('submit', function (e) { e.preventDefault(); apply(); });
    root.querySelectorAll('[data-sku-preset]').forEach(function (b) {
      b.addEventListener('click', function (e) {
        e.preventDefault();
        var today = new Date().toISOString().slice(0, 10);
        if (b.getAttribute('data-sku-preset') === 'today') loadSku({ sku_mp: form.querySelector('[name=sku_mp]').value, sku_from: today, sku_to: today });
        else loadSku({ sku_mp: form.querySelector('[name=sku_mp]').value, sku_from: '', sku_to: '' });
      });
    });
  }

  // ══ TRENDS tab ══
  function loadTrends(params) {
    var pane = document.getElementById('pane-trends');
    pane.classList.add('an-busy');
    fetch(url(Object.assign({ partial: 'trends' }, params)), { headers: { 'X-Requested-With': 'fetch' } })
      .then(function (r) { return r.text(); })
      .then(function (html) { pane.innerHTML = html; pane.dataset.loaded = '1'; pane.classList.remove('an-busy'); wireTrends(); syncURL(); })
      .catch(function () { pane.classList.remove('an-busy'); pane.innerHTML = '<div class="an-loading">Could not load trends.</div>'; });
  }

  function wireTrends() {
    var root = document.querySelector('#pane-trends .an-trends');
    if (!root) return;
    root.querySelectorAll('[data-trends-days]').forEach(function (a) {
      a.addEventListener('click', function (e) { e.preventDefault(); loadTrends({ days: a.getAttribute('data-trends-days') }); });
    });
  }

  // ══ FULFILMENT RISK tab ══
  function loadFulfil(params) {
    var pane = document.getElementById('pane-fulfil');
    var lid = showLoader(pane, [
      'Reading the period’s demand…',
      'Finding the latest run for each PO…',
      'Resolving each order’s warehouse…',
      'Checking current stock in each warehouse…',
      'Netting demand against available stock…',
      'Ranking at-risk SKUs by value…'
    ], 'Crunching every PO line against live inventory — this can take a few seconds.');
    fetch(url(Object.assign({ partial: 'fulfil' }, params)), { headers: { 'X-Requested-With': 'fetch' } })
      .then(function (r) { return r.text(); })
      .then(function (html) { clearInterval(lid); pane.innerHTML = html; pane.dataset.loaded = '1'; wireFulfil(); syncURL(); })
      .catch(function () { clearInterval(lid); pane.innerHTML = '<div class="an-loading">Could not load fulfilment risk.</div>'; });
  }

  function wireFulfil() {
    var root = document.querySelector('#pane-fulfil .an-fulfil');
    if (!root) return;
    wireSort(root);
    var form = root.querySelector('[data-fulfil-filter]');
    function apply() {
      loadFulfil({
        sku_mp: form.querySelector('[name=sku_mp]').value,
        sku_from: form.querySelector('[name=sku_from]').value,
        sku_to: form.querySelector('[name=sku_to]').value
      });
    }
    if (form) form.addEventListener('submit', function (e) { e.preventDefault(); apply(); });
    root.querySelectorAll('[data-fulfil-preset]').forEach(function (b) {
      b.addEventListener('click', function (e) {
        e.preventDefault();
        var today = new Date().toISOString().slice(0, 10);
        var mp = form.querySelector('[name=sku_mp]').value;
        if (b.getAttribute('data-fulfil-preset') === 'today') loadFulfil({ sku_mp: mp, sku_from: today, sku_to: today });
        else loadFulfil({ sku_mp: mp, sku_from: '', sku_to: '' });
      });
    });
  }

  // ══ EXCEPTIONS & QUALITY tab ══
  function loadExc(params) {
    var pane = document.getElementById('pane-exc');
    var lid = showLoader(pane, [
      'Reading uploaded PO lines…',
      'Flagging price mismatches…',
      'Grouping exceptions by marketplace…',
      'Comparing clean rate vs the previous window…'
    ], 'Scanning line-level status across the selected uploads.');
    fetch(url(Object.assign({ partial: 'exc' }, params)), { headers: { 'X-Requested-With': 'fetch' } })
      .then(function (r) { return r.text(); })
      .then(function (html) { clearInterval(lid); pane.innerHTML = html; pane.dataset.loaded = '1'; wireExc(); syncURL(); })
      .catch(function () { clearInterval(lid); pane.innerHTML = '<div class="an-loading">Could not load exceptions.</div>'; });
  }

  function wireExc() {
    var root = document.querySelector('#pane-exc .an-exc');
    if (!root) return;
    wireSort(root);
    var form = root.querySelector('[data-exc-filter]');
    function apply() {
      loadExc({
        sku_mp: form.querySelector('[name=sku_mp]').value,
        sku_from: form.querySelector('[name=sku_from]').value,
        sku_to: form.querySelector('[name=sku_to]').value
      });
    }
    if (form) form.addEventListener('submit', function (e) { e.preventDefault(); apply(); });
    root.querySelectorAll('[data-exc-preset]').forEach(function (b) {
      b.addEventListener('click', function (e) {
        e.preventDefault();
        var today = new Date().toISOString().slice(0, 10);
        var mp = form.querySelector('[name=sku_mp]').value;
        if (b.getAttribute('data-exc-preset') === 'today') loadExc({ sku_mp: mp, sku_from: today, sku_to: today });
        else loadExc({ sku_mp: mp, sku_from: '', sku_to: '' });
      });
    });
  }

  // ══ GEOGRAPHY & CONCENTRATION tab ══
  function loadGeo(params) {
    var pane = document.getElementById('pane-geo');
    var lid = showLoader(pane, [
      'Reading order value by location…',
      'Resolving locations to state & city…',
      'Ranking states and cities…',
      'Building the SKU concentration (Pareto) curve…'
    ], 'Mapping demand geographically and computing the ABC split.');
    fetch(url(Object.assign({ partial: 'geo' }, params)), { headers: { 'X-Requested-With': 'fetch' } })
      .then(function (r) { return r.text(); })
      .then(function (html) { clearInterval(lid); pane.innerHTML = html; pane.dataset.loaded = '1'; wireGeo(); syncURL(); })
      .catch(function () { clearInterval(lid); pane.innerHTML = '<div class="an-loading">Could not load geography.</div>'; });
  }

  function wireGeo() {
    var root = document.querySelector('#pane-geo .an-geo');
    if (!root) return;
    wireSort(root);
    var form = root.querySelector('[data-geo-filter]');
    var mpx = root.querySelector('[data-mpx]');
    function boxes() { return mpx ? [].slice.call(mpx.querySelectorAll('input[type=checkbox]')) : []; }
    function selMps() { return boxes().filter(function (b) { return b.checked; }).map(function (b) { return b.value; }).join('|'); }
    function seg() { return form.querySelector('[name=geo_seg]').value; }

    // marketplace multi-select dropdown (checkbox panel)
    if (mpx) {
      var btn = mpx.querySelector('[data-mpx-btn]'), panel = mpx.querySelector('[data-mpx-panel]'), label = mpx.querySelector('[data-mpx-label]');
      function refresh() { var n = boxes().filter(function (b) { return b.checked; }).length; label.textContent = n ? (n + ' marketplace' + (n > 1 ? 's' : '')) : 'All marketplaces'; }
      btn.addEventListener('click', function (e) { e.preventDefault(); panel.hidden = !panel.hidden; });
      document.addEventListener('click', function (e) { if (!mpx.contains(e.target)) panel.hidden = true; });
      mpx.querySelector('[data-mpx-all]').addEventListener('click', function (e) { e.preventDefault(); boxes().forEach(function (b) { b.checked = true; }); refresh(); });
      mpx.querySelector('[data-mpx-none]').addEventListener('click', function (e) { e.preventDefault(); boxes().forEach(function (b) { b.checked = false; }); refresh(); });
      mpx.addEventListener('change', refresh);
    }

    function apply() {
      loadGeo({ geo_seg: seg(), sku_mp: selMps(),
        sku_from: form.querySelector('[name=sku_from]').value,
        sku_to: form.querySelector('[name=sku_to]').value });
    }
    if (form) form.addEventListener('submit', function (e) { e.preventDefault(); apply(); });
    form.querySelector('[name=geo_seg]').addEventListener('change', apply);   // segment auto-applies
    root.querySelectorAll('[data-geo-preset]').forEach(function (b) {
      b.addEventListener('click', function (e) {
        e.preventDefault();
        var today = new Date().toISOString().slice(0, 10);
        var base = { geo_seg: seg(), sku_mp: selMps() };
        if (b.getAttribute('data-geo-preset') === 'today') loadGeo(Object.assign(base, { sku_from: today, sku_to: today }));
        else loadGeo(Object.assign(base, { sku_from: '', sku_to: '' }));
      });
    });
  }

  // ══ OTIF · READINESS tab ══
  function loadOtif(params) {
    var pane = document.getElementById('pane-otif');
    var lid = showLoader(pane, [
      'Reading open orders (still due)…',
      'Resolving each order’s warehouse…',
      'Checking current stock coverage…',
      'Scoring readiness & due-date urgency…'
    ], 'Projected OTIF — open orders vs current stock. Not actual delivery.');
    fetch(url(Object.assign({ partial: 'otif' }, params)), { headers: { 'X-Requested-With': 'fetch' } })
      .then(function (r) { return r.text(); })
      .then(function (html) { clearInterval(lid); pane.innerHTML = html; pane.dataset.loaded = '1'; wireOtif(); syncURL(); })
      .catch(function () { clearInterval(lid); pane.innerHTML = '<div class="an-loading">Could not load readiness.</div>'; });
  }

  function wireOtif() {
    var root = document.querySelector('#pane-otif .an-otif');
    if (!root) return;
    wireSort(root);
    var mp = root.querySelector('[name=sku_mp]');
    if (mp) mp.addEventListener('change', function () { loadOtif({ sku_mp: mp.value, horizon: root.dataset.horizon || 0 }); });
    root.querySelectorAll('[data-otif-horizon]').forEach(function (a) {
      a.addEventListener('click', function (e) { e.preventDefault(); loadOtif({ sku_mp: mp ? mp.value : '', horizon: a.getAttribute('data-otif-horizon') }); });
    });
  }

  // ══ INVENTORY · DAYS-OF-SUPPLY tab ══
  function loadDos(params) {
    var pane = document.getElementById('pane-dos');
    var lid = showLoader(pane, [
      'Reading current stock on-hand…',
      'Computing average daily demand…',
      'Calculating days of supply…',
      'Bucketing stockout risk & overstock…'
    ], 'Stock cover per SKU — on-hand vs demand rate.');
    fetch(url(Object.assign({ partial: 'dos' }, params)), { headers: { 'X-Requested-With': 'fetch' } })
      .then(function (r) { return r.text(); })
      .then(function (html) { clearInterval(lid); pane.innerHTML = html; pane.dataset.loaded = '1'; wireDos(); syncURL(); })
      .catch(function () { clearInterval(lid); pane.innerHTML = '<div class="an-loading">Could not load days of supply.</div>'; });
  }

  function wireDos() {
    var root = document.querySelector('#pane-dos .an-dos');
    if (!root) return;
    wireSort(root);
    var mp = root.querySelector('[name=sku_mp]');
    if (mp) mp.addEventListener('change', function () { loadDos({ sku_mp: mp.value, days: root.dataset.days || 30 }); });
    root.querySelectorAll('[data-dos-days]').forEach(function (a) {
      a.addEventListener('click', function (e) { e.preventDefault(); loadDos({ sku_mp: mp ? mp.value : '', days: a.getAttribute('data-dos-days') }); });
    });
  }

  // ══ tabs ══
  function showTab(name) {
    activeTab = name;
    document.querySelectorAll('.an-tab').forEach(function (t) { t.classList.toggle('on', t.dataset.tab === name); });
    document.querySelectorAll('.an-pane').forEach(function (p) { p.hidden = (p.id !== 'pane-' + name); });
    var q = new URLSearchParams(location.search);
    if (name === 'sku' && document.getElementById('pane-sku').dataset.loaded === '0') {
      loadSku({ sku_mp: q.get('sku_mp') || '', sku_from: q.get('sku_from') || undefined, sku_to: q.get('sku_to') || undefined });
    } else if (name === 'trends' && document.getElementById('pane-trends').dataset.loaded === '0') {
      loadTrends({ days: q.get('days') || 30 });
    } else if (name === 'fulfil' && document.getElementById('pane-fulfil').dataset.loaded === '0') {
      loadFulfil({ sku_mp: q.get('sku_mp') || '', sku_from: q.get('sku_from') || undefined, sku_to: q.get('sku_to') || undefined });
    } else if (name === 'exc' && document.getElementById('pane-exc').dataset.loaded === '0') {
      loadExc({ sku_mp: q.get('sku_mp') || '', sku_from: q.get('sku_from') || undefined, sku_to: q.get('sku_to') || undefined });
    } else if (name === 'geo' && document.getElementById('pane-geo').dataset.loaded === '0') {
      loadGeo({ sku_mp: q.get('sku_mp') || '', geo_seg: q.get('geo_seg') || '', sku_from: q.get('sku_from') || undefined, sku_to: q.get('sku_to') || undefined });
    } else if (name === 'otif' && document.getElementById('pane-otif').dataset.loaded === '0') {
      loadOtif({ sku_mp: q.get('sku_mp') || '', horizon: q.get('horizon') || 0 });
    } else if (name === 'dos' && document.getElementById('pane-dos').dataset.loaded === '0') {
      loadDos({ sku_mp: q.get('sku_mp') || '', days: q.get('days') || 30 });
    } else {
      syncURL();
    }
  }
  document.querySelectorAll('.an-tab').forEach(function (t) {
    t.addEventListener('click', function () { showTab(t.dataset.tab); });
  });

  // ── boot ──
  function boot() {
    wireDaily();                                  // daily was server-rendered
    var q = new URLSearchParams(location.search);
    var t = q.get('tab');
    if (['sku', 'trends', 'fulfil', 'exc', 'geo', 'otif', 'dos'].indexOf(t) >= 0) showTab(t);  // deep-link
    else syncURL();
  }
  // ApexCharts is a separate <script> that loads ASYNC under the persistent
  // shell-nav — wait for it (up to ~6s) before booting so charts don't render into
  // a "library unavailable" note on a timing race. Non-chart wiring can wait too;
  // 6s is a safety ceiling, the lib is normally ready in well under a second.
  var booted = false;
  function bootWhenReady() {
    if (booted) return;
    var waited = 0;
    (function wait() {
      if (booted) return;
      if (window.ApexCharts || waited >= 6000) { booted = true; boot(); return; }
      waited += 100; setTimeout(wait, 100);
    })();
  }
  if (document.readyState !== 'loading') bootWhenReady();
  else document.addEventListener('DOMContentLoaded', bootWhenReady);
})();
