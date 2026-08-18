/* online_b2b/online_b2b/lines.html — page script (separated). Server values via #lines-cfg JSON. */
var CFG = JSON.parse(document.getElementById("lines-cfg").textContent);
(function () {
  var form = document.getElementById('ln-filters');
  if (!form) return;
  var results = document.getElementById('ln-results');
  var spin = document.getElementById('ln-spin');
  var base = CFG["b2b_lines"];
  var moreUrl = CFG["b2b_lines_more"];
  var timer = null, ctrl = null;

  function params(extra) {
    var p = new URLSearchParams();
    if (form.marketplace.value) p.set('marketplace', form.marketplace.value);
    if (form.status.value) p.set('status', form.status.value);
    if (form.po.value.trim()) p.set('po', form.po.value.trim());
    if (form.q.value.trim()) p.set('q', form.q.value.trim());
    if (extra) for (var k in extra) p.set(k, extra[k]);
    return p;
  }
  function load(push) {
    var p = params(); var qs = p.toString();
    if (ctrl) ctrl.abort(); ctrl = new AbortController();
    results.classList.add('loading'); spin.hidden = false;
    p.set('partial', '1');
    fetch(base + '?' + p.toString(), { headers: { 'X-Requested-With': 'XMLHttpRequest' }, signal: ctrl.signal })
      .then(function (r) { return r.text(); })
      .then(function (html) { results.innerHTML = html; results.classList.remove('loading'); spin.hidden = true; bindMore();
        if (push) history.pushState({}, '', qs ? '?' + qs : base); })
      .catch(function (e) { if (e.name !== 'AbortError') { results.classList.remove('loading'); spin.hidden = true; } });
  }
  function bindMore() {
    var more = document.getElementById('lines-more');
    if (!more) return;
    more.addEventListener('click', function () {
      var off = parseInt(more.getAttribute('data-offset'), 10);
      var total = parseInt(more.getAttribute('data-total'), 10);
      var lim = parseInt(more.getAttribute('data-limit'), 10);
      more.disabled = true; more.textContent = 'Loading…';
      fetch(moreUrl + '?' + params({ offset: off }).toString(),
            { headers: { 'X-Requested-With': 'XMLHttpRequest' } })
        .then(function (r) { return r.text(); })
        .then(function (html) {
          document.getElementById('lines-body').insertAdjacentHTML('beforeend', html);
          var got = (html.match(/<tr/g) || []).length, now = off + got;
          if (now >= total || got < lim) { more.remove(); }
          else { more.disabled = false; more.setAttribute('data-offset', now);
                 more.textContent = 'Load more (' + now + ' of ' + total + ')'; }
        });
    });
  }
  form.marketplace.addEventListener('change', function () { load(true); });
  form.status.addEventListener('change', function () { load(true); });
  form.po.addEventListener('input', function () { clearTimeout(timer); timer = setTimeout(function () { load(true); }, 350); });
  form.q.addEventListener('input', function () { clearTimeout(timer); timer = setTimeout(function () { load(true); }, 350); });
  document.getElementById('ln-reset').addEventListener('click', function (e) {
    e.preventDefault(); form.marketplace.value = ''; form.status.value = ''; form.po.value = ''; form.q.value = ''; load(true);
  });
  bindMore();
})();
