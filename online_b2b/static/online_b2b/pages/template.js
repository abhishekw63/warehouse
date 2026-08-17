/* online_b2b/template.html — page script (separated from template). */
(function () {
    var table = document.querySelector('.tpl-table');
    if (!table) return;
    var hot = [];
    function clear() { hot.forEach(function (el) { el.classList.remove('col-hot'); }); hot = []; }
    table.addEventListener('pointerover', function (e) {
      var cell = e.target.closest('[data-col]');
      if (!cell) { return; }
      var col = cell.getAttribute('data-col');
      if (hot.length && hot[0].getAttribute('data-col') === col) { return; }
      clear();
      hot = Array.prototype.slice.call(table.querySelectorAll('[data-col="' + col + '"]'));
      hot.forEach(function (el) { el.classList.add('col-hot'); });
    });
    table.addEventListener('pointerleave', clear);
  })();
