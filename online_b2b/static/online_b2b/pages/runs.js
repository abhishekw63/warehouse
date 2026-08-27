/* online_b2b/runs.html — page script: instant client-side filter over loaded runs. */
(function () {
  var input = document.getElementById('runsFilter');
  var table = document.getElementById('runsTable');
  var count = document.getElementById('runsCount');
  var foot = document.getElementById('runsFoot');
  if (!input || !table) return;
  var rows = [].slice.call(table.tBodies[0] ? table.tBodies[0].rows : []);
  var total = rows.length;

  function apply() {
    var q = input.value.trim().toLowerCase();
    var shown = 0;
    rows.forEach(function (tr) {
      var key = (tr.getAttribute('data-key') || tr.textContent || '').toLowerCase();
      var hit = !q || key.indexOf(q) !== -1;
      tr.hidden = !hit;
      if (hit) shown++;
    });
    if (count) count.textContent = (q ? shown + ' / ' + total : total) + ' run' + (total === 1 ? '' : 's');
    if (foot) foot.hidden = shown !== 0;
  }
  var t;
  input.addEventListener('input', function () { clearTimeout(t); t = setTimeout(apply, 80); });
})();
