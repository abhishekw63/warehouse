/* online_b2b/ui_lab.html — page script (separated from template). */
// Tiny ApexCharts demo (front-end only) — mirrors the house chart style.
  window.addEventListener('DOMContentLoaded', function () {
    if (!window.ApexCharts) return;
    var dark = document.documentElement.getAttribute('data-theme') === 'dark';
    new ApexCharts(document.querySelector('#lab-chart'), {
      chart: { type: 'area', height: 200, toolbar: { show: false }, animations: { easing: 'easeinout', speed: 700 } },
      series: [{ name: 'Demo intake', data: [31, 40, 28, 51, 42, 62, 58] }],
      xaxis: { categories: ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'] },
      colors: ['var(--accent)'], dataLabels: { enabled: false }, stroke: { curve: 'smooth', width: 2 },
      fill: { type: 'gradient', gradient: { opacityFrom: .4, opacityTo: .05 } },
      grid: { borderColor: dark ? '#272c3a' : '#eef1f5' },
      theme: { mode: dark ? 'dark' : 'light' },
      tooltip: { theme: dark ? 'dark' : 'light' }
    }).render();
  });
