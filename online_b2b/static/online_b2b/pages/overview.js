/* online_b2b/overview.html — page script (separated from template).
   Charts on ECharts (was ApexCharts) so the app ships ONE chart library. */
(function () {
  var el = document.getElementById('b2b-charts');
  if (!el) return;
  var data;
  try { data = JSON.parse(el.textContent); } catch (e) { return; }

  function inr(v) {
    v = Number(v) || 0; var s = v < 0 ? '-' : ''; v = Math.abs(v);
    if (v >= 1e7) return s + '₹' + (v / 1e7).toFixed(2) + ' Cr';
    if (v >= 1e5) return s + '₹' + (v / 1e5).toFixed(2) + ' L';
    if (v >= 1000) return s + '₹' + Math.round(v).toLocaleString('en-IN');
    return s + '₹' + Math.round(v);
  }
  // ECharts draws on canvas → CSS vars don't resolve there; read them to real hex.
  function cssVar(name, fb) {
    try { return getComputedStyle(document.documentElement).getPropertyValue(name).trim() || fb; }
    catch (e) { return fb; }
  }
  var ACCENT = cssVar('--accent', '#4f46e5');
  var GREY = '#9aa1b2';
  var COLORS = [ACCENT, cssVar('--accent-2', '#22c1c3'), '#11998e', '#f7971e', '#cb2d3e',
    '#2193b0', '#7b4397', '#16a34a', '#db2777', '#9aa1b2'];
  function has(a) { return Array.isArray(a) && a.some(function (x) { return Number(x) > 0; }); }
  function note(id, t) { var n = document.getElementById(id); if (n) n.innerHTML = '<div class="chart-empty">' + t + '</div>'; }

  var areaChart = null, donut = null;

  function build() {
    if (!window.echarts) { note('b2b-area', 'Charts library failed to load.'); note('b2b-donut', ''); return; }
    var E = window.echarts;

    // ── Area: 30-day trend (value / POs toggle) ──
    function areaOption(isVal) {
      return {
        grid: { left: 6, right: 14, top: 14, bottom: 4, containLabel: true },
        xAxis: {
          type: 'category', data: data.trend.labels, boundaryGap: false,
          axisLine: { show: false }, axisTick: { show: false },
          axisLabel: { color: GREY, fontSize: 10,
            interval: Math.max(0, Math.ceil(data.trend.labels.length / 6) - 1) }
        },
        yAxis: {
          type: 'value', splitLine: { lineStyle: { color: '#eef1f5', type: 'dashed' } },
          axisLabel: { color: GREY, fontSize: 10,
            formatter: function (v) { return isVal ? inr(v) : v; } }
        },
        tooltip: {
          trigger: 'axis',
          formatter: function (ps) {
            var p = ps[0];
            return p.axisValue + '<br/>' + p.marker +
              (isVal ? inr(p.value) : (p.value + ' POs'));
          }
        },
        series: [{
          name: isVal ? 'Value' : 'POs', type: 'line', smooth: true, symbol: 'none',
          lineStyle: { width: 2.5, color: ACCENT }, itemStyle: { color: ACCENT },
          areaStyle: { color: ACCENT, opacity: 0.14 },
          data: isVal ? data.trend.value : data.trend.orders
        }]
      };
    }
    try {
      if (!has(data.trend.value) && !has(data.trend.orders)) {
        note('b2b-area', 'No orders in the last 30 days.');
      } else {
        var ae = document.getElementById('b2b-area'); ae.innerHTML = '';
        areaChart = E.init(ae);
        areaChart.setOption(areaOption(true));
      }
    } catch (e) { console.error(e); note('b2b-area', 'Chart error.'); }

    if (areaChart) {
      document.querySelectorAll('.ctg').forEach(function (b) {
        b.addEventListener('click', function () {
          document.querySelectorAll('.ctg').forEach(function (x) { x.classList.toggle('on', x === b); });
          var isVal = b.getAttribute('data-metric') === 'value';
          areaChart.setOption(areaOption(isVal));
          document.getElementById('ch-metric').textContent = isVal ? 'Value' : 'POs';
        });
      });
    }

    // ── Donut: marketplace mix ──
    try {
      if (!has(data.mix.value)) {
        note('b2b-donut', 'No marketplace data.');
      } else {
        var de = document.getElementById('b2b-donut'); de.innerHTML = '';
        donut = E.init(de);
        donut.setOption({
          tooltip: {
            trigger: 'item',
            formatter: function (p) { return p.name + ': ' + inr(p.value) + ' (' + p.percent + '%)'; }
          },
          legend: { bottom: 0, type: 'scroll', textStyle: { color: GREY, fontSize: 11 } },
          series: [{
            type: 'pie', radius: ['58%', '78%'], center: ['50%', '42%'],
            avoidLabelOverlap: true, label: { show: false },
            itemStyle: { borderColor: '#fff', borderWidth: 2 },
            data: data.mix.labels.map(function (lbl, i) {
              return { name: lbl, value: data.mix.value[i], itemStyle: { color: COLORS[i % COLORS.length] } };
            })
          }]
        });
      }
    } catch (e) { console.error(e); note('b2b-donut', 'Chart error.'); }

    document.querySelectorAll('[data-countup]').forEach(function (n) {
      var target = parseFloat(n.getAttribute('data-countup')) || 0, t0 = null;
      function tick(ts) {
        if (!t0) t0 = ts;
        var p = Math.min((ts - t0) / 900, 1);
        n.textContent = Math.round(target * (0.5 - Math.cos(p * Math.PI) / 2)).toLocaleString('en-IN');
        if (p < 1) requestAnimationFrame(tick);
      }
      requestAnimationFrame(tick);
    });
  }

  window.addEventListener('resize', function () {
    if (areaChart) areaChart.resize();
    if (donut) donut.resize();
  });

  // ECharts loads as a separate <script> and (under the shell-nav) may be injected
  // ASYNC — so wait for window.echarts (up to ~6s) before building; only after the
  // timeout does build() surface the real "failed to load" note.
  var built = false;
  function start() {
    if (built) return;
    var waited = 0;
    (function waitLib() {
      if (built) return;
      if (window.echarts || waited >= 6000) {
        built = true;
        try { build(); } catch (e) { console.error(e); }
        return;
      }
      waited += 100;
      setTimeout(waitLib, 100);
    })();
  }
  if (document.readyState !== 'loading') start();
  else document.addEventListener('DOMContentLoaded', start);
})();
