/* GT Select · D365 Import — single-page flow (upload → inline preview → classify
   unknown posting groups → confirm). Talks to the CBV endpoints (JSON). Kept out
   of the template per our asset-separation rule. */
(function () {
  function $(id) { return document.getElementById(id); }
  var cfgEl = $('gtsConfig');
  if (!cfgEl) return;
  var CFG = JSON.parse(cfgEl.textContent);
  var taxEl = $('gtsTaxonomy');
  var TAX = taxEl ? JSON.parse(taxEl.textContent)
                  : { segments: [], online: [], offline: [], mt_children: [] };
  var form = $('gtsForm');
  var STATE = { token: null, needs: [], overrides: {}, nNew: 0 };

  function csrf() {
    var el = form.querySelector('[name=csrfmiddlewaretoken]');
    return el ? el.value : '';
  }
  function inr(v) {
    v = Number(v) || 0;
    if (v >= 1e7) return (v / 1e7).toFixed(2) + ' Cr';
    if (v >= 1e5) return (v / 1e5).toFixed(2) + ' Lakh';
    return Math.round(v).toLocaleString('en-IN') + ' Rs';
  }
  function compact(v) {
    v = Number(v) || 0;
    if (v >= 1e5) return (v / 1e5).toFixed(1) + 'L';
    if (v >= 1e3) return (v / 1e3).toFixed(1) + 'k';
    return String(v);
  }
  function esc(s) { var d = document.createElement('div'); d.textContent = s == null ? '' : s; return d.innerHTML; }
  function show(el, on) { if (el) el.hidden = !on; }
  function fail(msg) { var e = $('gtsError'); e.textContent = msg; show(e, true); }

  // ── upload → preview ──────────────────────────────────────────────────
  form.addEventListener('submit', function (e) {
    e.preventDefault();
    show($('gtsError'), false); show($('gtsPreview'), false); show($('gtsDone'), false);
    show($('gtsLoading'), true);
    $('gtsPreviewBtn').disabled = true;
    fetch(CFG.uploadUrl, {
      method: 'POST', body: new FormData(form), credentials: 'same-origin',
      headers: { 'X-Requested-With': 'fetch', 'X-CSRFToken': csrf() }
    }).then(function (r) { return r.json(); }).then(function (j) {
      show($('gtsLoading'), false); $('gtsPreviewBtn').disabled = false;
      if (!j.ok) { fail(j.error || 'Could not read the files.'); return; }
      STATE.token = j.token; STATE.needs = j.needs_class || []; STATE.overrides = {};
      STATE.nNew = (j.summary && j.summary.new) || 0;
      render(j);
    }).catch(function () {
      show($('gtsLoading'), false); $('gtsPreviewBtn').disabled = false; fail('Network error.');
    });
  });

  // ── render preview ────────────────────────────────────────────────────
  function render(j) {
    var s = j.summary;
    $('gtsSrc').textContent = (j.meta ? (j.meta.headers_name + ' + ' + j.meta.lines_name + ' · ') : '') +
      s.total + ' order(s) · ' + s.empty + ' empty shell(s) skipped';
    $('gtsKNew').textContent = s.new;
    $('gtsKLines').textContent = s.new_lines;
    $('gtsKQty').textContent = compact(s.qty);
    $('gtsKVal').textContent = inr(s.value);
    $('gtsKDup').textContent = s.dup;

    var wb = $('gtsWarn');
    if (j.warnings && j.warnings.length) {
      wb.innerHTML = '<strong>' + j.warnings.length + ' note(s)</strong><ul>' +
        j.warnings.map(function (w) { return '<li>' + esc(w) + '</li>'; }).join('') + '</ul>';
      show(wb, true);
    } else show(wb, false);

    var ct = $('gtsChanTbl').querySelector('tbody');
    ct.innerHTML = (j.channels || []).map(function (c) {
      return '<tr class="' + (c.new ? '' : 'rnim') + '"><td>' + esc(c.segment) + '</td><td><b>' +
        esc(c.marketplace) + '</b></td><td class="r">' + (c.new ? '<span class="st ok">' + c.new + '</span>' : '0') +
        '</td><td class="r muted">' + c.dup + '</td><td class="r">' + c.new_qty + '</td><td class="r">' +
        inr(c.new_value) + '</td></tr>';
    }).join('');
    $('gtsChanCount').textContent = (j.channels || []).length;

    var ot = $('gtsOrdersTbl').querySelector('tbody');
    ot.innerHTML = (j.new_orders || []).map(function (h) {
      return '<tr><td class="mono">' + esc(h.so_no) + '</td><td class="mono">' + esc(h.external_doc || '—') +
        '</td><td><span class="st ok">' + esc(h.marketplace) + '</span></td><td class="desc" title="' +
        esc(h.ship_name) + '">' + esc(h.ship_name || '—') + '</td><td class="mono">' + esc(h.warehouse) +
        '</td><td class="muted">' + esc(h.po_date || '—') + '</td><td class="r">' + h.line_count +
        '</td><td class="r">' + h.qty + '</td><td class="r">' + inr(h.order_value) + '</td></tr>';
    }).join('');
    $('gtsNewCount').textContent = s.new;

    renderClassify();
    updateConfirm();
    show($('gtsPreview'), true);
    $('gtsPreview').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  // ── unknown posting-group classification ──────────────────────────────
  function optionsHTML(list) {
    return '<option value="">—</option>' + list.map(function (o) {
      return '<option value="' + esc(o.value) + '">' + esc(o.label) + '</option>';
    }).join('');
  }
  function renderClassify() {
    var box = $('gtsClassify'), list = $('gtsClList');
    if (!STATE.needs.length) { show(box, false); return; }
    show(box, true);
    $('gtsClCount').textContent = STATE.needs.length;
    list.innerHTML = STATE.needs.map(function (u) {
      return '<div class="gts-cl-row" data-key="' + esc(u.key) + '" data-pg="' + esc(u.posting_group) + '">' +
        '<div class="gts-cl-pg"><b>' + esc(u.posting_group) + '</b><span>' + u.count + ' order(s) · ' +
        u.qty + ' qty · ' + inr(u.value) + '</span></div>' +
        '<select class="gts-cl-seg">' + optionsHTML(TAX.segments) + '</select>' +
        '<select class="gts-cl-mp" disabled><option value="">—</option></select>' +
        '<select class="gts-cl-child" hidden><option value="">—</option></select></div>';
    }).join('');
    Array.prototype.forEach.call(list.querySelectorAll('.gts-cl-row'), function (row) {
      var segEl = row.querySelector('.gts-cl-seg'), mpEl = row.querySelector('.gts-cl-mp'),
          chEl = row.querySelector('.gts-cl-child');
      segEl.addEventListener('change', function () {
        var mps = segEl.value === 'Offline' ? TAX.offline : (segEl.value === 'OnlineB2B' ? TAX.online : []);
        mpEl.innerHTML = optionsHTML(mps); mpEl.disabled = !mps.length;
        chEl.hidden = true; chEl.innerHTML = '<option value="">—</option>';
        collect();
      });
      mpEl.addEventListener('change', function () {
        var mps = segEl.value === 'Offline' ? TAX.offline : TAX.online, isMT = false;
        mps.forEach(function (o) { if (o.value === mpEl.value && o.mt) isMT = true; });
        if (isMT) { chEl.hidden = false; chEl.innerHTML = optionsHTML(TAX.mt_children); }
        else { chEl.hidden = true; chEl.innerHTML = '<option value="">—</option>'; }
        collect();
      });
      chEl.addEventListener('change', collect);
    });
  }
  function collect() {
    var ov = {};
    Array.prototype.forEach.call($('gtsClList').querySelectorAll('.gts-cl-row'), function (row) {
      var key = row.getAttribute('data-key'), pg = row.getAttribute('data-pg');
      var segEl = row.querySelector('.gts-cl-seg'), mpEl = row.querySelector('.gts-cl-mp'),
          chEl = row.querySelector('.gts-cl-child');
      var seg = segEl.value, mp = mpEl.value;
      if (!seg || !mp) return;
      var isMT = (mp === 'MT'), childV = chEl.hidden ? '' : chEl.value;
      if (isMT && !childV) return;                       // MT needs a child
      var mpText = mpEl.options[mpEl.selectedIndex] ? mpEl.options[mpEl.selectedIndex].text : mp;
      var chText = chEl.hidden || !chEl.options[chEl.selectedIndex] ? '' : chEl.options[chEl.selectedIndex].text;
      ov[key] = { posting_group: pg, segment: seg, marketplace: mp,
                  marketplace_label: isMT ? chText : mpText };
    });
    STATE.overrides = ov;
    updateConfirm();
  }

  // ── confirm ───────────────────────────────────────────────────────────
  function updateConfirm() {
    var btn = $('gtsConfirm'), txt = $('gtsConfirmText');
    if (!STATE.nNew) {
      txt.innerHTML = '<span class="cb-warn">Nothing new</span> — all orders are already captured.';
      btn.disabled = true; return;
    }
    var remaining = STATE.needs.length - Object.keys(STATE.overrides).length;
    if (remaining > 0) {
      txt.innerHTML = 'Place the <b>' + remaining + '</b> remaining unknown posting group(s) to enable import.';
      btn.disabled = true; return;
    }
    txt.innerHTML = 'Import <b>' + STATE.nNew + ' new order(s)</b> — each recorded under its own channel.';
    btn.disabled = false;
  }

  var confirmBtn = $('gtsConfirm');
  if (confirmBtn) confirmBtn.addEventListener('click', function () {
    if (!STATE.token) return;
    confirmBtn.disabled = true; confirmBtn.textContent = 'Importing…';
    fetch(CFG.gtBase + STATE.token + '/confirm/', {
      method: 'POST', credentials: 'same-origin',
      headers: { 'X-Requested-With': 'fetch', 'X-CSRFToken': csrf(), 'Content-Type': 'application/json' },
      body: JSON.stringify({ overrides: STATE.overrides })
    }).then(function (r) { return r.json(); }).then(function (j) {
      if (!j.ok) { confirmBtn.disabled = false; confirmBtn.textContent = '✓ Confirm & import'; fail(j.error || 'Import failed.'); return; }
      show($('gtsPreview'), false); show($('gtsForm'), false);
      var done = $('gtsDone');
      done.innerHTML = '<div class="gts-done-card"><div class="gd-ic">✓</div>' +
        '<h2>Imported ' + j.imported + ' order(s) · ' + j.lines + ' line(s)</h2>' +
        '<p>' + j.skipped + ' already-captured order(s) skipped. They\'re in the tracker now.</p>' +
        '<div class="gts-done-act"><a class="btn-primary" href="' + j.redirect + '">View in Tracker →</a>' +
        '<a class="btn-ghost" href="' + CFG.gtBase + '">Import another</a></div></div>';
      show(done, true);
      done.scrollIntoView({ behavior: 'smooth', block: 'center' });
      if (window.B2B && window.B2B.celebrate) window.B2B.celebrate();
    }).catch(function () {
      confirmBtn.disabled = false; confirmBtn.textContent = '✓ Confirm & import'; fail('Network error during import.');
    });
  });

  var discardBtn = $('gtsDiscard');
  if (discardBtn) discardBtn.addEventListener('click', function () {
    show($('gtsPreview'), false);
    STATE = { token: null, needs: [], overrides: {}, nNew: 0 };
    form.reset();
    form.scrollIntoView({ behavior: 'smooth', block: 'start' });
  });
})();
