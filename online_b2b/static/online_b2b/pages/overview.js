/* online_b2b/overview.html — page script (separated from template). */
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
  var ACCENT = 'var(--accent)', GREY = '#9aa1b2';
  var COLORS = ['var(--accent)', 'var(--accent-2)', '#11998e', '#f7971e', '#cb2d3e', '#2193b0', '#7b4397', '#16a34a', '#db2777', '#9aa1b2'];
  function has(a) { return Array.isArray(a) && a.some(function (x) { return Number(x) > 0; }); }
  function note(id, t) { var n = document.getElementById(id); if (n) n.innerHTML = '<div class="chart-empty">' + t + '</div>'; }

  function build() {
    if (!window.ApexCharts) { note('b2b-area', 'Charts library failed to load.'); note('b2b-donut', ''); return; }

    var area = null;
    try {
      if (!has(data.trend.value) && !has(data.trend.orders)) {
        note('b2b-area', 'No orders in the last 30 days.');
      } else {
        document.getElementById('b2b-area').innerHTML = '';
        area = new ApexCharts(document.getElementById('b2b-area'), {
          chart: { type: 'area', height: 260, fontFamily: 'Inter, sans-serif', toolbar: { show: false } },
          series: [{ name: 'Value', data: data.trend.value }],
          colors: [ACCENT],
          stroke: { curve: 'smooth', width: 2.5 },
          fill: { type: 'gradient', gradient: { shadeIntensity: 1, opacityFrom: 0.4, opacityTo: 0.03, stops: [0, 95] } },
          dataLabels: { enabled: false },
          markers: { size: 0, hover: { size: 5 } },
          xaxis: { categories: data.trend.labels, tickAmount: 6, axisBorder: { show: false }, axisTicks: { show: false }, labels: { style: { colors: GREY, fontSize: '10px' } } },
          yaxis: { labels: { style: { colors: GREY, fontSize: '10px' }, formatter: function (v) { return inr(v); } } },
          grid: { borderColor: '#eef1f5', strokeDashArray: 4 },
          tooltip: { y: { formatter: function (v) { return inr(v); } } }
        });
        area.render();
      }
    } catch (e) { console.error(e); note('b2b-area', 'Chart error.'); }

    if (area) {
      document.querySelectorAll('.ctg').forEach(function (b) {
        b.addEventListener('click', function () {
          document.querySelectorAll('.ctg').forEach(function (x) { x.classList.toggle('on', x === b); });
          var isVal = b.getAttribute('data-metric') === 'value';
          area.updateSeries([{ name: isVal ? 'Value' : 'POs', data: isVal ? data.trend.value : data.trend.orders }]);
          area.updateOptions({ tooltip: { y: { formatter: function (v) { return isVal ? inr(v) : (v + ' POs'); } } } });
          document.getElementById('ch-metric').textContent = isVal ? 'Value' : 'POs';
        });
      });
    }

    try {
      if (!has(data.mix.value)) {
        note('b2b-donut', 'No marketplace data.');
      } else {
        document.getElementById('b2b-donut').innerHTML = '';
        new ApexCharts(document.getElementById('b2b-donut'), {
          chart: { type: 'donut', height: 260, fontFamily: 'Inter, sans-serif' },
          series: data.mix.value,
          labels: data.mix.labels,
          colors: COLORS,
          stroke: { width: 2, colors: ['#fff'] },
          dataLabels: { enabled: false },
          legend: { position: 'bottom', fontSize: '11px', labels: { colors: GREY } },
          tooltip: { y: { formatter: function (v) { return inr(v); } } },
          plotOptions: { pie: { donut: { size: '70%' } } }
        }).render();
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

  // ApexCharts loads as a separate <script>. Under the persistent shell-nav the
  // library is (re)injected and loads ASYNC, so this page script can run before it
  // finishes — checking window.ApexCharts once then would wrongly report "failed to
  // load". Wait for the library to appear (up to ~6s) before building; only after
  // the timeout does build() surface the real failure note.
  var built = false;
  function start() {
    if (built) return;
    var waited = 0;
    (function waitLib() {
      if (built) return;
      if (window.ApexCharts || waited >= 6000) {
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
