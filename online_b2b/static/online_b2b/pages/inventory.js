/* online_b2b/inventory.html — page script (separated from template).
   Server values (URLs + new_total) come from the #inventory-cfg JSON block. */
var CFG = JSON.parse(document.getElementById("inventory-cfg").textContent);

(function () {
  var input = document.getElementById('ivq');
  var table = document.querySelector('.iv-stock');
  if (!input || !table || !table.tBodies[0]) return;
  var tb = table.tBodies[0];
  var rows = [].slice.call(tb.rows);
  var meta = document.querySelector('.iv-stockmeta');
  var metaHTML = meta ? meta.innerHTML : '';
  var clearBtn = document.querySelector('.iv-clear');
  var reduce = window.matchMedia && matchMedia('(prefers-reduced-motion: reduce)').matches;
  // Precompute each row's searchable text (item + description + EAN only).
  rows.forEach(function (r) {
    var c = r.cells;
    r._s = ((c[0] ? c[0].textContent : '') + ' ' + (c[1] ? c[1].textContent : '') +
            ' ' + (c[2] ? c[2].textContent : '')).toLowerCase();
  });
  var noRow = null;
  function apply() {
    var raw = input.value.trim().toLowerCase();
    // Split on space / comma / semicolon / pipe / tab / newline so a PASTED list
    // of item nos or EANs filters to ALL of them at once (OR match).
    var tokens = raw ? raw.split(/[\s,;|]+/).filter(Boolean) : [];
    var shown = 0;
    for (var i = 0; i < rows.length; i++) {
      var s = rows[i]._s, m = !tokens.length;
      for (var k = 0; !m && k < tokens.length; k++) if (s.indexOf(tokens[k]) !== -1) m = true;
      rows[i].hidden = !m; if (m) shown++;
    }
    if (meta) meta.innerHTML = tokens.length
      ? '<b>' + shown.toLocaleString('en-IN') + '</b> of <b>' + rows.length.toLocaleString('en-IN') + '</b> SKUs match'
        + (tokens.length > 1 ? ' · ' + tokens.length + ' terms' : '')
      : metaHTML;
    if (clearBtn) clearBtn.hidden = !tokens.length;
    if (tokens.length && shown === 0) {
      if (!noRow) {
        noRow = tb.insertRow(); noRow.className = 'iv-nomatch';
        var td = noRow.insertCell(); td.className = 'iv-none';
        td.colSpan = rows[0] ? rows[0].cells.length : 6;
      }
      noRow.hidden = false;
      noRow.cells[0].textContent = 'No items match “' + input.value.trim() + '”.';
    } else if (noRow) { noRow.hidden = true; }
  }
  var t;
  function run() {
    if (reduce) { apply(); return; }
    tb.style.transition = 'opacity .13s ease';
    tb.style.opacity = '0.3';
    window.setTimeout(function () { apply(); tb.style.opacity = '1'; }, 110);
  }
  input.addEventListener('input', function () { clearTimeout(t); t = window.setTimeout(run, 120); });
  if (clearBtn) clearBtn.addEventListener('click', function (e) {
    e.preventDefault(); input.value = ''; input.focus(); run();
  });
  if (input.value.trim()) apply();   // honour an initial ?q= without a reload
})();

/* Count-up the warehouse "sellable units" heroes on load (rAF, eased). */
(function () {
  var reduce = window.matchMedia && matchMedia('(prefers-reduced-motion: reduce)').matches;
  var nums = [].slice.call(document.querySelectorAll('.iv-num[data-to]'));
  if (!nums.length) return;
  function fmtIN(v) {                       // Indian digit grouping
    var s = String(Math.round(v)), n = s.length;
    if (n <= 3) return s;
    var head = s.slice(0, n - 3), tail = s.slice(n - 3);
    return head.replace(/\B(?=(\d{2})+(?!\d))/g, ',') + ',' + tail;
  }
  nums.forEach(function (el) {
    var to = parseFloat(el.getAttribute('data-to')) || 0;
    if (reduce || to <= 0) { el.textContent = fmtIN(to); return; }
    var dur = 900, start = null;
    function step(ts) {
      if (start === null) start = ts;
      var p = Math.min(1, (ts - start) / dur);
      var e = 1 - Math.pow(1 - p, 3);        // easeOutCubic
      el.textContent = fmtIN(to * e);
      if (p < 1) requestAnimationFrame(step);
      else el.textContent = fmtIN(to);
    }
    requestAnimationFrame(step);
  });
})();

/* Bin-coverage <details>: re-trigger the reveal animation on EVERY open (a CSS
   animation on a persistent [open] selector only fires once — reflow resets it). */
(function () {
  var reduce = window.matchMedia && matchMedia('(prefers-reduced-motion: reduce)').matches;
  if (reduce) return;
  document.querySelectorAll('details.iv-cov').forEach(function (d) {
    var body = d.querySelector('.iv-cov-body');
    if (!body) return;
    d.addEventListener('toggle', function () {
      if (!d.open) return;
      body.style.animation = 'none';
      void body.offsetWidth;      // force reflow so the animation can restart
      body.style.animation = '';
    });
  });
})();

/* ── Bin coverage: click-to-toggle Include/Exclude (durable) + Lock & apply ── */
(function () {
  var d = document;
  var tok = d.querySelector('[name=csrfmiddlewaretoken]');
  var csrf = tok ? tok.value : '';
  function toast(msg, type) {
    if (window.B2B && B2B.toast) { B2B.toast(msg, {type: type || 'ok'}); }
  }
  function whSel(wh) { return '[data-wh="' + String(wh).replace(/(["\\])/g, '\\$1') + '"]'; }

  function recount(cols) {
    ['include', 'exclude'].forEach(function (col) {
      var box = cols.querySelector('.iv-cov-col[data-col="' + col + '"]');
      if (!box) return;
      var n = box.querySelectorAll('.iv-cov-bin').length;
      var cnt = box.querySelector('.cnt');
      if (cnt) cnt.textContent = n;
    });
  }

  function moveBin(btn, wh, next) {
    var cols = d.querySelector('.iv-cov-cols' + whSel(wh));
    if (!cols) return;
    var dest = cols.querySelector('.iv-cov-col[data-col="' + next + '"] .iv-cov-list');
    if (!dest) return;
    btn.classList.add('mv-out');
    setTimeout(function () {
      btn.classList.remove('mv-out', 'isnew');
      btn.setAttribute('data-dec', next);
      btn.title = next === 'include' ? 'Click to Exclude' : 'Click to Include';
      var mv = btn.querySelector('.mv');
      if (mv) {
        mv.className = 'mv ' + (next === 'include' ? 'out' : 'in');
        mv.textContent = next === 'include' ? 'Exclude →' : '← Include';
      }
      var nd = btn.querySelector('.newdot');
      if (nd) nd.remove();
      var none = dest.querySelector('.iv-cov-none');
      if (none) none.remove();
      dest.insertBefore(btn, dest.firstChild);
      btn.classList.add('mv-in');
      setTimeout(function () { btn.classList.remove('mv-in'); }, 340);
      recount(cols);
      var dirty = d.querySelector('.iv-cov-dirty' + whSel(wh));
      var apply = d.querySelector('.iv-cov-apply' + whSel(wh));
      if (dirty) dirty.hidden = false;
      if (apply) apply.hidden = false;
    }, 200);
  }

  // Toggle a single bin (optimistic; server persists the durable exact-bin rule).
  d.addEventListener('click', function (e) {
    var btn = e.target.closest ? e.target.closest('.iv-cov-bin') : null;
    if (!btn || btn.classList.contains('saving')) return;
    var wh = btn.getAttribute('data-wh');
    var cur = btn.getAttribute('data-dec');
    var next = cur === 'include' ? 'exclude' : 'include';
    btn.classList.add('saving');
    var body = new FormData();
    body.append('bin_code', btn.getAttribute('data-bin'));
    body.append('warehouse', wh);
    body.append('decision', next);
    body.append('csrfmiddlewaretoken', csrf);
    fetch(CFG.bin_set, {
      method: 'POST', credentials: 'same-origin',
      headers: {'X-Requested-With': 'XMLHttpRequest'}, body: body
    }).then(function (r) { return r.json(); }).then(function (j) {
      btn.classList.remove('saving');
      if (!j || !j.ok) { toast((j && j.error) || 'Could not update bin.', 'error'); return; }
      moveBin(btn, wh, next);
    }).catch(function () {
      btn.classList.remove('saving');
      toast('Network error — bin not changed.', 'error');
    });
  });

  // Lock & apply → reclassify the snapshot so available stock reflects the new split.
  d.addEventListener('click', function (e) {
    var b = e.target.closest ? e.target.closest('.iv-cov-apply') : null;
    if (!b || b.disabled) return;
    var wh = b.getAttribute('data-wh');
    b.disabled = true;
    var label = b.innerHTML;
    b.textContent = 'Applying…';
    var body = new FormData();
    body.append('warehouse', wh);
    body.append('csrfmiddlewaretoken', csrf);
    fetch(CFG.apply, {
      method: 'POST', credentials: 'same-origin',
      headers: {'X-Requested-With': 'XMLHttpRequest'}, body: body
    }).then(function (r) { return r.json(); }).then(function (j) {
      if (j && j.ok) {
        toast('Locked & applied — available stock recalculated.', 'ok');
        setTimeout(function () { location.reload(); }, 750);
      } else {
        b.disabled = false; b.innerHTML = label;
        toast((j && j.error) || 'Apply failed.', 'error');
      }
    }).catch(function () {
      b.disabled = false; b.innerHTML = label;
      toast('Network error — nothing applied.', 'error');
    });
  });

  // Nudge the user when unrecognised bins landed in Excluded.
  var newTotal = CFG.new_total;
  if (newTotal > 0) {
    setTimeout(function () {
      toast(newTotal + ' unrecognised bin' + (newTotal === 1 ? '' : 's') +
            ' sit in Excluded — open “Bin coverage” and click any to Include if it holds sellable stock.', 'info');
    }, 800);
  }
})();
