/* online_b2b/inventory_bins.html — Manage bins (per-bin Include/Exclude).
   Server values (URLs + wh + new_total) come from the #inventory-bins-cfg block.
   Each toggle POSTs to bin_set (durable per-bin exact rule); "Lock & apply"
   reclassifies the snapshot so available stock reflects the choices. */
var CFG = JSON.parse(document.getElementById("inventory-bins-cfg").textContent);

(function () {
  var d = document;
  var B = window.B2B || {};
  var csrf = B.csrf ? B.csrf() : '';
  function toast(msg, type) { if (B.toast) B.toast(msg, { type: type || 'ok' }); }

  var counts = {
    inc: d.querySelector('.c-inc'), exc: d.querySelector('.c-exc'), nw: d.querySelector('.c-new'),
  };
  function setCount(el, delta) {
    if (!el) return;
    el.textContent = Math.max(0, (parseInt(el.textContent, 10) || 0) + delta);
  }
  function markDirty() {
    var dirty = d.querySelector('.ivb-dirty'), apply = d.querySelector('.ivb-apply');
    if (dirty) dirty.hidden = false;
    if (apply) apply.hidden = false;
  }

  // ── toggle a single bin ─────────────────────────────────────────────────────
  d.addEventListener('click', function (e) {
    var opt = e.target.closest ? e.target.closest('.ivb-opt') : null;
    if (!opt) return;
    var row = opt.closest('.ivb-row');
    if (!row || row.classList.contains('saving')) return;
    var next = opt.getAttribute('data-set');               // 'include' | 'exclude'
    var cur = row.getAttribute('data-dec');                // 'include' | 'exclude' | 'new'
    if (cur === next) return;                              // already there → no-op

    row.classList.add('saving');
    var body = new FormData();
    body.append('bin_code', row.getAttribute('data-bin'));
    body.append('warehouse', CFG.wh);
    body.append('decision', next);
    body.append('csrfmiddlewaretoken', csrf);
    B.postForm(CFG.bin_set, body).then(function (j) {
      row.classList.remove('saving');
      if (!j || !j.ok) { toast((j && j.error) || 'Could not update bin.', 'error'); return; }
      // active button
      row.querySelectorAll('.ivb-opt').forEach(function (b) {
        b.classList.toggle('on', b.getAttribute('data-set') === next);
      });
      // summary deltas (was cur → now next); 'new' counts as an excluded-side flag
      if (cur === 'include') setCount(counts.inc, -1); else setCount(counts.exc, -1);
      if (next === 'include') setCount(counts.inc, +1); else setCount(counts.exc, +1);
      if (cur === 'new') {                                 // leaving the new state
        setCount(counts.nw, -1);
        var nd = row.querySelector('.newdot'); if (nd) nd.remove();
        row.classList.remove('isnew');
        var ns = d.querySelector('.ivb-sum .new');
        if (ns && counts.nw && (parseInt(counts.nw.textContent, 10) || 0) === 0) ns.hidden = true;
      }
      row.setAttribute('data-dec', next);
      markDirty();
    }).catch(function () {
      row.classList.remove('saving');
      toast('Network error — bin not changed.', 'error');
    });
  });

  // ── Lock & apply → reclassify the snapshot ──────────────────────────────────
  d.addEventListener('click', function (e) {
    var b = e.target.closest ? e.target.closest('.ivb-apply') : null;
    if (!b || b.disabled) return;
    b.disabled = true; var label = b.innerHTML; b.textContent = 'Applying…';
    var body = new FormData();
    body.append('warehouse', CFG.wh);
    body.append('csrfmiddlewaretoken', csrf);
    B.postForm(CFG.apply, body).then(function (j) {
      if (j && j.ok) {
        toast('Locked & applied — available stock recalculated.', 'ok');
        setTimeout(function () { location.reload(); }, 700);
      } else {
        b.disabled = false; b.innerHTML = label;
        toast((j && j.error) || 'Apply failed.', 'error');
      }
    }).catch(function () {
      b.disabled = false; b.innerHTML = label;
      toast('Network error — nothing applied.', 'error');
    });
  });

  // ── client-side filter ──────────────────────────────────────────────────────
  var search = d.querySelector('.ivb-search');
  if (search) {
    search.addEventListener('input', function () {
      var q = this.value.trim().toLowerCase();
      d.querySelectorAll('.ivb-row').forEach(function (r) {
        r.hidden = q && (r.getAttribute('data-search') || '').indexOf(q) === -1;
      });
    });
  }

  // ── nudge on unclassified new bins ──────────────────────────────────────────
  if (CFG.new_total > 0) {
    setTimeout(function () {
      toast(CFG.new_total + ' new bin' + (CFG.new_total === 1 ? '' : 's') +
            ' need a decision — they sit in Exclude until you Include them.', 'info');
    }, 700);
  }
})();
