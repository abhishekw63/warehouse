/* online_b2b/sku_demand.html — page script (separated from template). */
(function () {
  document.querySelectorAll('table.sortable').forEach(function (table) {
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
})();
