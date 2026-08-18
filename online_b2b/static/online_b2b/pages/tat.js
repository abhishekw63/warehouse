/* online_b2b/tat.html — page script (separated from template).
   Server-rendered values (URLs) come from the #tat-cfg JSON block; POST + CSRF go
   through B2B.postForm. No Django tags here so it lives as a static file. */
(function () {
  var form = document.getElementById('tat-filters'); if (!form) return;
  var results = document.getElementById('tat-results');
  var spin = document.getElementById('tat-spin');
  var cfgEl = document.getElementById('tat-cfg'); if (!cfgEl) return;
  var CFG = JSON.parse(cfgEl.textContent);
  var base = CFG.base, saveUrl = CFG.save;
  var exportLink = document.getElementById('tat-export'), exportBase = CFG.export;
  var timer = null, ctrl = null;

  function params() {
    var p = new URLSearchParams();
    p.set('status', form.status.value);
    if (form.segment.value) p.set('segment', form.segment.value);
    if (form.run.value) p.set('run', form.run.value);
    if (form.q.value.trim()) p.set('q', form.q.value.trim());
    if (form.date_from.value) p.set('date_from', form.date_from.value);
    if (form.date_to.value) p.set('date_to', form.date_to.value);
    return p;
  }
  function updateExport() { if (exportLink) exportLink.href = exportBase + '?' + params().toString(); }
  function load() {
    var p = params(); if (ctrl) ctrl.abort(); ctrl = new AbortController();
    results.classList.add('loading'); spin.hidden = false; p.set('partial', '1');
    fetch(base + '?' + p.toString(), { headers: { 'X-Requested-With': 'XMLHttpRequest' }, signal: ctrl.signal })
      .then(function (r) { return r.text(); })
      .then(function (html) { results.innerHTML = html; results.classList.remove('loading'); spin.hidden = true; bind(); updateExport(); })
      .catch(function (e) { if (e.name !== 'AbortError') { results.classList.remove('loading'); spin.hidden = true; } });
  }
  function save(order) {
    var sel = results.querySelector('.tat-reason[data-order="' + order + '"]');
    var note = results.querySelector('.tat-note[data-order="' + order + '"]');
    B2B.postForm(saveUrl, {
      order_id: order,
      reason_code: sel ? sel.value : '',
      note: note ? note.value : ''
    })
      .then(function (j) {
        if (!j.ok) return;
        // if filtering by pending and a reason was just set, refresh so it leaves the view
        if (form.status.value === 'pending' && sel && sel.value) { load(); return; }
        if (sel) { var t = document.createElement('span'); t.className = 'tat-saved'; t.textContent = '✓'; sel.parentNode.appendChild(t); setTimeout(function () { t.remove(); }, 1200); }
      });
  }
  function bind() {
    results.querySelectorAll('.tat-reason').forEach(function (s) { s.addEventListener('change', function () { save(s.getAttribute('data-order')); }); });
    results.querySelectorAll('.tat-note').forEach(function (i) { i.addEventListener('change', function () { save(i.getAttribute('data-order')); }); });
    results.querySelectorAll('.ik[data-filter]').forEach(function (card) {
      card.addEventListener('click', function () { var f = card.getAttribute('data-filter').split(':'); if (f[0] === 'status') form.status.value = f[1]; load(); });
    });
  }
  form.status.addEventListener('change', load);
  form.segment.addEventListener('change', load);
  form.run.addEventListener('change', load);
  form.q.addEventListener('input', B2B.debounce(load, 300));
  form.date_from.addEventListener('change', load);
  form.date_to.addEventListener('change', load);
  document.getElementById('tat-today').addEventListener('click', function () {
    var t = new Date(); var iso = t.getFullYear() + '-' + String(t.getMonth() + 1).padStart(2, '0') + '-' + String(t.getDate()).padStart(2, '0');
    form.date_from.value = iso; form.date_to.value = iso; load();
  });
  document.getElementById('tat-reset').addEventListener('click', function () {
    form.status.value = 'pending'; form.segment.value = ''; form.run.value = ''; form.q.value = ''; form.date_from.value = ''; form.date_to.value = ''; load();
  });
  updateExport(); bind();
})();
