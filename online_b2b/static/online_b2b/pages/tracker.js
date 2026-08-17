/* online_b2b/tracker.html — page script (separated from template). */
(function () {
  var body = document.getElementById('trk-body');
  var loader = document.getElementById('trkLoader');
  var filter = document.getElementById('trkFilter');
  var clearBtn = document.getElementById('trkClear');
  var exportBtn = document.getElementById('trkExport');
  var base = location.pathname;

  function params() {
    var p = new URLSearchParams();
    ['segment', 'marketplace', 'warehouse', 'q'].forEach(function (n) {
      var el = filter.querySelector('[name="' + n + '"]');
      if (el && el.value) p.set(n, el.value);
    });
    return p;
  }

  function loadBody() {
    var p = params();
    var lt = setTimeout(function () { if (loader) loader.hidden = false; }, 180);  // loader only if slow
    body.classList.add('fade');
    fetch(base + '?partial=1&' + p.toString(), { headers: { 'X-Requested-With': 'fetch' } })
      .then(function (r) { return r.text(); })
      .then(function (html) {
        clearTimeout(lt); if (loader) loader.hidden = true;
        body.innerHTML = html; body.classList.remove('fade');
        initTable(); syncUI(p);
        history.replaceState(null, '', p.toString() ? base + '?' + p.toString() : base);
      })
      .catch(function () { clearTimeout(lt); if (loader) loader.hidden = true; body.classList.remove('fade'); });
  }

  function syncUI(p) {
    var any = ['segment', 'marketplace', 'warehouse', 'q'].some(function (n) { return p.get(n); });
    if (clearBtn) clearBtn.hidden = !any;
    if (exportBtn) exportBtn.href = base + 'export/?' + p.toString();
  }

  // filter events (no reload)
  filter.querySelectorAll('select').forEach(function (s) { s.addEventListener('change', loadBody); });
  filter.addEventListener('submit', function (e) { e.preventDefault(); loadBody(); });

  // Facility chips live INSIDE the re-rendered body → delegate on the persistent
  // container: a click sets the warehouse filter and reloads (one filter path).
  body.addEventListener('click', function (e) {
    var chip = e.target.closest && e.target.closest('.trk-fac');
    if (!chip) return;
    var whSel = filter.querySelector('[name="warehouse"]');
    if (whSel) whSel.value = chip.getAttribute('data-fac') || '';
    loadBody();
  });
  var qEl = filter.querySelector('[name=q]'), qT;
  if (qEl) qEl.addEventListener('input', function () { clearTimeout(qT); qT = setTimeout(loadBody, 400); });
  if (clearBtn) clearBtn.addEventListener('click', function (e) {
    e.preventDefault();
    filter.querySelectorAll('select').forEach(function (s) { s.value = ''; });
    if (qEl) qEl.value = '';
    loadBody();
  });

  // table init (re-run after every AJAX swap): sort + quick-filter
  function initTable() {
    var table = body.querySelector('table.sortable');
    if (table) table.querySelectorAll('th.srt').forEach(function (th) {
      th.addEventListener('click', function () {
        var k = th.getAttribute('data-k'), asc = th.getAttribute('data-dir') !== 'asc';
        table.querySelectorAll('th.srt').forEach(function (x) { x.removeAttribute('data-dir'); x.classList.remove('on'); });
        th.setAttribute('data-dir', asc ? 'asc' : 'desc'); th.classList.add('on');
        var tb = table.tBodies[0];
        Array.prototype.slice.call(tb.rows).sort(function (a, b) {
          var ae = a.querySelector('[data-k="' + k + '"]'), be = b.querySelector('[data-k="' + k + '"]');
          var av = ae ? (ae.getAttribute('data-v') || ae.textContent).trim() : '';
          var bv = be ? (be.getAttribute('data-v') || be.textContent).trim() : '';
          var na = parseFloat(av), nb = parseFloat(bv);
          if (!isNaN(na) && !isNaN(nb)) return asc ? na - nb : nb - na;
          return asc ? av.localeCompare(bv) : bv.localeCompare(av);
        }).forEach(function (r) { tb.appendChild(r); });
      });
    });
  }

  // quick client-side filter over loaded rows (re-queries fresh each keypress)
  var quick = document.getElementById('trkQuick'), qkT;
  if (quick) quick.addEventListener('input', function () {
    clearTimeout(qkT);
    qkT = setTimeout(function () {            // debounce so fast typing never janks
      var v = quick.value.trim().toLowerCase();
      body.querySelectorAll('.trk-tbl tbody tr').forEach(function (tr) {
        tr.style.display = (!v || tr.textContent.toLowerCase().indexOf(v) >= 0) ? '' : 'none';
      });
    }, 140);
  });

  // add-PO panel toggle
  var addBtn = document.getElementById('trkAddBtn'), addForm = document.getElementById('trkAdd'),
      addCancel = document.getElementById('trkAddCancel');
  if (addBtn && addForm) addBtn.addEventListener('click', function () {
    addForm.hidden = !addForm.hidden;
    if (!addForm.hidden) { var f = addForm.querySelector('[name=po]'); if (f) f.focus(); }
  });
  if (addCancel) addCancel.addEventListener('click', function () { addForm.hidden = true; });

  initTable();
})();
