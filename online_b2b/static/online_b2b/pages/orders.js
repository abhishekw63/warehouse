/* online_b2b/orders.html — page script (separated from template).
   Server-rendered values come from the #orders-cfg JSON block. */
(function () {
  var form = document.getElementById('b2b-filters');
  if (!form) return;
  var results = document.getElementById('b2b-results');
  var spin = document.getElementById('b2b-spin');
  var CFG = JSON.parse(document.getElementById('orders-cfg').textContent);
  var base = CFG.base;
  var moreUrl = CFG.more;
  var state = { sort: CFG.sort, dir: CFG.dir };
  var timer = null, ctrl = null;

  var SKELETON = (function () {
    var rows = '';
    for (var i = 0; i < 10; i++) rows += '<div class="sk-row"><span class="sk sk-1"></span><span class="sk sk-2"></span><span class="sk sk-3"></span></div>';
    return '<div class="sk-panel">' + rows + '</div>';
  })();

  function params(extra) {
    var p = new URLSearchParams();
    var seg = form.segment.value, mp = form.marketplace.value, days = form.days.value,
        wh = form.warehouse.value, ot = form.order_type.value, df = form.date_from.value,
        dt = form.date_to.value, q = form.q.value.trim();
    if (seg) p.set('segment', seg);
    if (mp) p.set('marketplace', mp);
    if (days && days !== '0') p.set('days', days);
    if (wh) p.set('warehouse', wh);
    if (ot) p.set('order_type', ot);
    if (df) p.set('date_from', df);
    if (dt) p.set('date_to', dt);
    if (q) p.set('q', q);
    if (state.sort && state.sort !== 'date') p.set('sort', state.sort);
    if (state.dir && state.dir !== 'desc') p.set('dir', state.dir);
    if (extra) for (var k in extra) p.set(k, extra[k]);
    return p;
  }

  function load(push) {
    var p = params(); var qs = p.toString();
    if (ctrl) ctrl.abort();
    ctrl = new AbortController();
    results.innerHTML = SKELETON; spin.hidden = false;
    p.set('partial', '1');
    fetch(base + '?' + p.toString(), { headers: { 'X-Requested-With': 'XMLHttpRequest' }, signal: ctrl.signal })
      .then(function (r) { return r.text(); })
      .then(function (html) {
        results.innerHTML = html; spin.hidden = true; bind();
        if (push) history.pushState({}, '', qs ? '?' + qs : base);
      })
      .catch(function (e) { if (e.name !== 'AbortError') { spin.hidden = true; } });
  }

  function bind() {
    results.querySelectorAll('th.srt').forEach(function (th) {
      th.addEventListener('click', function () {
        var s = th.getAttribute('data-sort');
        if (state.sort === s) { state.dir = (state.dir === 'asc') ? 'desc' : 'asc'; }
        else { state.sort = s; state.dir = 'desc'; }
        load(true);
      });
    });
    var more = document.getElementById('b2b-more');
    if (more) more.addEventListener('click', function () {
      var off = parseInt(more.getAttribute('data-offset'), 10);
      var total = parseInt(more.getAttribute('data-total'), 10);
      var lim = parseInt(more.getAttribute('data-limit'), 10);
      more.disabled = true; more.textContent = 'Loading…';
      fetch(moreUrl + '?' + params({ offset: off }).toString(),
            { headers: { 'X-Requested-With': 'XMLHttpRequest' } })
        .then(function (r) { return r.text(); })
        .then(function (html) {
          document.getElementById('b2b-orders-body').insertAdjacentHTML('beforeend', html);
          var got = (html.match(/<tr/g) || []).length;
          var now = off + got;
          if (now >= total || got < lim) { more.remove(); }
          else { more.disabled = false; more.setAttribute('data-offset', now);
                 more.textContent = 'Load more (' + now + ' of ' + total + ')'; }
        });
    });
  }

  // Segment change reloads the page so the Marketplace dropdown repopulates
  // for that segment (Online vs Offline have different marketplaces).
  form.segment.addEventListener('change', function () {
    form.marketplace.value = '';
    location.search = params().toString();
  });
  form.marketplace.addEventListener('change', function () { load(true); });
  form.days.addEventListener('change', function () { load(true); });
  form.warehouse.addEventListener('change', function () { load(true); });
  form.order_type.addEventListener('change', function () { load(true); });
  form.date_from.addEventListener('change', function () { load(true); });
  form.date_to.addEventListener('change', function () { load(true); });
  form.q.addEventListener('input', function () { clearTimeout(timer); timer = setTimeout(function () { load(true); }, 350); });
  form.addEventListener('submit', function (e) { e.preventDefault(); load(true); });
  document.getElementById('b2b-reset').addEventListener('click', function (e) {
    e.preventDefault();
    form.segment.value = ''; form.marketplace.value = ''; form.days.value = '0';
    form.warehouse.value = ''; form.order_type.value = ''; form.date_from.value = '';
    form.date_to.value = ''; form.q.value = '';
    state.sort = 'date'; state.dir = 'desc'; load(true);
  });
  window.addEventListener('popstate', function () {
    var p = new URLSearchParams(location.search);
    form.marketplace.value = p.get('marketplace') || ''; form.days.value = p.get('days') || '0';
    form.warehouse.value = p.get('warehouse') || ''; form.order_type.value = p.get('order_type') || '';
    form.date_from.value = p.get('date_from') || ''; form.date_to.value = p.get('date_to') || '';
    form.q.value = p.get('q') || ''; state.sort = p.get('sort') || 'date'; state.dir = p.get('dir') || 'desc';
    load(false);
  });

  bind();
})();
