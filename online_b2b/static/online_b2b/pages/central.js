/* online_b2b/central.html — page script (separated from template). */
// Hub range switch — fetch the windowed KPI block and swap it in, no page reload.
// Vanilla, no libraries. Request-id guard drops stale responses; history.replaceState
// keeps the range on refresh; subtitle + active chip update in lockstep.
(function () {
  var bar = document.getElementById('hubRange');
  var slot = document.getElementById('hubWindowed');
  if (!bar || !slot) return;
  var LABELS = {today:'today','7d':'last 7 days','30d':'last 30 days',mtd:'this month',all:'all-time'};
  var reqId = 0;

  function setActive(range) {
    bar.querySelectorAll('.hr-chip').forEach(function (c) {
      var on = c.getAttribute('data-range') === range;
      c.classList.toggle('on', on);
      c.setAttribute('aria-selected', on ? 'true' : 'false');
    });
    var lab = document.getElementById('hubRangeLabel');
    if (lab) lab.textContent = LABELS[range] || range;
  }

  function load(range) {
    var mine = ++reqId;
    slot.classList.add('is-swapping');
    fetch('?range=' + encodeURIComponent(range) + '&partial=1',
          {credentials:'same-origin', headers:{'X-Requested-With':'XMLHttpRequest'}})
      .then(function (r) { if (!r.ok) throw new Error('HTTP ' + r.status); return r.text(); })
      .then(function (html) {
        if (mine !== reqId) return;                 // a newer click won — ignore
        slot.innerHTML = html;
        slot.classList.remove('is-swapping');
        setActive(range);
        var u = new URL(window.location.href);
        u.searchParams.set('range', range);
        u.searchParams.delete('partial');
        history.replaceState(null, '', u);
      })
      .catch(function () {
        if (mine === reqId) slot.classList.remove('is-swapping');
      });
  }

  bar.addEventListener('click', function (ev) {
    var chip = ev.target.closest('.hr-chip');
    if (!chip || chip.classList.contains('on')) return;
    load(chip.getAttribute('data-range'));
  });
})();
