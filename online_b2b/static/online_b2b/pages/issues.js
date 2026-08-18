/* online_b2b/online_b2b/issues.html — page script (separated). Server values via #issues-cfg JSON. */
var CFG = JSON.parse(document.getElementById("issues-cfg").textContent);
(function () {
  var form = document.getElementById('iss-filters');
  if (!form) return;
  var results = document.getElementById('iss-results');
  var spin = document.getElementById('iss-spin');
  var base = CFG["b2b_issues"];
  var saveUrl = CFG["b2b_issues_save"];
  var bulkUrl = CFG["b2b_issues_save_bulk"];
  var fixUrl  = CFG["b2b_issues_fix_ean"];
  var csrf = document.getElementById('iss-csrf').value;
  var hdrs = { 'X-CSRFToken': csrf, 'X-Requested-With': 'XMLHttpRequest',
               'Content-Type': 'application/x-www-form-urlencoded' };
  var timer = null, ctrl = null;

  var exportLink = document.getElementById('iss-export');
  var exportBase = CFG["b2b_issues_export"];
  function filterParams() {
    var p = new URLSearchParams();
    p.set('resolution', form.resolution.value);
    if (form.status.value) p.set('status', form.status.value);
    if (form.q.value.trim()) p.set('q', form.q.value.trim());
    if (form.date_from.value) p.set('date_from', form.date_from.value);
    if (form.date_to.value) p.set('date_to', form.date_to.value);
    return p;
  }
  function updateExport() { if (exportLink) exportLink.href = exportBase + '?' + filterParams().toString(); }

  // Lift the KPI card row OUT of the results container to ABOVE the bulk-set
  // bar (it renders inside the partial so it stays live on every filtered load).
  function relocateKpis() {
    var bulk = document.querySelector('.bulk-bar');
    if (!bulk || !bulk.parentNode) return;
    var old = document.getElementById('iss-kpis-top');
    if (old) old.remove();
    var kpis = results.querySelector('.iss-kpis');
    if (kpis) { kpis.id = 'iss-kpis-top'; bulk.parentNode.insertBefore(kpis, bulk); }
  }

  function load() {
    var p = filterParams();
    if (ctrl) ctrl.abort(); ctrl = new AbortController();
    results.classList.add('loading'); spin.hidden = false;
    p.set('partial', '1');
    fetch(base + '?' + p.toString(), { headers: { 'X-Requested-With': 'XMLHttpRequest' }, signal: ctrl.signal })
      .then(function (r) { return r.text(); })
      .then(function (html) { results.innerHTML = html; relocateKpis(); results.classList.remove('loading'); spin.hidden = true; bindActions(); updateExport(); })
      .catch(function (e) { if (e.name !== 'AbortError') { results.classList.remove('loading'); spin.hidden = true; } });
  }

  function saveRow(line, reload) {
    var sel = results.querySelector('.act-sel[data-line="' + line + '"]');
    var rem = results.querySelector('.act-rem[data-line="' + line + '"]');
    var tick = results.querySelector('.saved[data-line="' + line + '"]');
    var body = new URLSearchParams();
    body.set('line_id', line);
    body.set('action', sel ? sel.value : '');
    body.set('remark', rem ? rem.value : '');
    fetch(saveUrl, { method: 'POST', body: body.toString(), headers: hdrs })
      .then(function (r) { return r.json(); })
      .then(function (j) {
        if (!j.ok) return;
        if (reload) { load(); }                 // action changed → resolution may change → refresh
        else if (tick) { tick.hidden = false; setTimeout(function () { tick.hidden = true; }, 1500); }
      });
  }
  function fixEan(line, btn) {
    var box = results.querySelector('.iss-eanfix[data-line="' + line + '"]');
    var inp = results.querySelector('.ef-in[data-line="' + line + '"]');
    var msg = results.querySelector('.ef-msg[data-line="' + line + '"]');
    var val = (inp ? inp.value : '').trim();
    if (!val) { if (inp) inp.focus(); return; }
    if (val === btn.getAttribute('data-ean')) {
      if (msg) { msg.textContent = 'same as the wrong EAN'; msg.className = 'ef-msg err'; }
      return;
    }
    btn.disabled = true; if (inp) inp.disabled = true;
    if (msg) { msg.textContent = 'resolving…'; msg.className = 'ef-msg'; }
    var body = new URLSearchParams();
    body.set('line_id', line); body.set('correct_ean', val);
    fetch(fixUrl, { method: 'POST', body: body.toString(), headers: hdrs })
      .then(function (r) { return r.json(); })
      .then(function (j) {
        if (!j.ok) {
          btn.disabled = false; if (inp) inp.disabled = false;
          if (msg) { msg.textContent = j.error || 'failed'; msg.className = 'ef-msg err'; }
          return;
        }
        if (msg) {
          msg.textContent = '✓ ' + j.status + ' · item ' + j.item_no;
          msg.className = 'ef-msg ok';
        }
        if (box) box.classList.add('done');
        setTimeout(load, 750);    // refresh → row leaves Pending, lands in Resolved/audit
      })
      .catch(function () {
        btn.disabled = false; if (inp) inp.disabled = false;
        if (msg) { msg.textContent = 'network error'; msg.className = 'ef-msg err'; }
      });
  }
  // NOT_IN_MASTER that isn't a real SKU (e.g. a virtual combo already dropped
  // from the PO) → resolve it as EXCLUDE, so it leaves Unresolved.
  function dropNim(line, btn) {
    var msg = results.querySelector('.ef-msg[data-line="' + line + '"]');
    btn.disabled = true;
    if (msg) { msg.textContent = 'excluding…'; msg.className = 'ef-msg'; }
    var body = new URLSearchParams();
    body.set('line_id', line);
    body.set('action', 'EXCLUDE');
    body.set('remark', 'not a real SKU — dropped from PO');
    fetch(saveUrl, { method: 'POST', body: body.toString(), headers: hdrs })
      .then(function (r) { return r.json(); })
      .then(function (j) {
        if (!j.ok) { btn.disabled = false; if (msg) { msg.textContent = j.error || 'failed'; msg.className = 'ef-msg err'; } return; }
        if (msg) { msg.textContent = '✓ excluded'; msg.className = 'ef-msg ok'; }
        setTimeout(load, 700);   // → leaves Unresolved, lands in Resolved
      })
      .catch(function () { btn.disabled = false; if (msg) { msg.textContent = 'network error'; msg.className = 'ef-msg err'; } });
  }
  function bindActions() {
    results.querySelectorAll('.ef-go').forEach(function (b) {
      b.addEventListener('click', function () { fixEan(b.getAttribute('data-line'), b); });
    });
    results.querySelectorAll('.ef-drop').forEach(function (b) {
      b.addEventListener('click', function () { dropNim(b.getAttribute('data-line'), b); });
    });
    results.querySelectorAll('.ef-in').forEach(function (inp) {
      inp.addEventListener('keydown', function (e) {
        if (e.key === 'Enter') {
          e.preventDefault();
          var b = results.querySelector('.ef-go[data-line="' + inp.getAttribute('data-line') + '"]');
          if (b) fixEan(inp.getAttribute('data-line'), b);
        }
      });
    });
    results.querySelectorAll('.act-sel').forEach(function (s) {
      s.addEventListener('change', function () { saveRow(s.getAttribute('data-line'), true); });
    });
    results.querySelectorAll('.act-rem').forEach(function (inp) {
      inp.addEventListener('change', function () { saveRow(inp.getAttribute('data-line'), false); });
    });
    // Click a count card to filter (incl. RESOLVED → view resolved lines).
    results.querySelectorAll('.ik[data-filter]').forEach(function (card) {
      card.addEventListener('click', function () {
        var f = card.getAttribute('data-filter').split(':');
        if (f[0] === 'status') form.status.value = f[1] || '';
        else if (f[0] === 'resolution') { form.resolution.value = f[1]; form.status.value = ''; }
        load();
      });
    });
  }

  form.resolution.addEventListener('change', load);
  form.status.addEventListener('change', load);
  form.q.addEventListener('input', function () { clearTimeout(timer); timer = setTimeout(load, 300); });
  form.date_from.addEventListener('change', load);
  form.date_to.addEventListener('change', load);
  document.getElementById('iss-today').addEventListener('click', function () {
    var t = new Date();
    var iso = t.getFullYear() + '-' + String(t.getMonth() + 1).padStart(2, '0') + '-' + String(t.getDate()).padStart(2, '0');
    form.date_from.value = iso; form.date_to.value = iso; load();
  });
  document.getElementById('iss-reset').addEventListener('click', function () {
    form.resolution.value = 'pending'; form.status.value = '';
    form.q.value = ''; form.date_from.value = ''; form.date_to.value = '';
    load();
  });
  updateExport();
  relocateKpis();          // lift the KPI row above the bulk bar on first paint

  document.getElementById('bulk-apply').addEventListener('click', function () {
    var act = document.getElementById('bulk-action').value;
    var rem = document.getElementById('bulk-remark').value;
    var sels = results.querySelectorAll('.act-sel[data-line]');
    var msg = document.getElementById('bulk-msg');
    if (!sels.length || (!act && !rem)) return;
    var body = new URLSearchParams();
    body.set('action', act); body.set('remark', rem);
    sels.forEach(function (s) { body.append('line_ids', s.getAttribute('data-line')); });
    fetch(bulkUrl, { method: 'POST', body: body.toString(), headers: hdrs })
      .then(function (r) { return r.json(); })
      .then(function (j) {
        if (!j.ok) return;
        msg.textContent = '✓ applied to ' + j.updated; msg.hidden = false;
        setTimeout(function () { msg.hidden = true; load(); }, 700);
      });
  });

  bindActions();
})();

/* Email modal — preview then send the filtered issue lines. Self-contained;
   reads filters straight off the Issues form so it always matches the view.
   Generic enough to reuse for future email features (point data-*-url at them). */
(function () {
  var modal = document.getElementById('email-modal');
  var btn = document.getElementById('iss-email');
  var form = document.getElementById('iss-filters');
  if (!modal || !btn || !form) return;
  var csrfEl = document.getElementById('iss-csrf');
  var csrf = csrfEl ? csrfEl.value : '';
  var $ = function (id) { return document.getElementById(id); };

  var toIn = $('em-to-in'), ccIn = $('em-cc-in'), noteIn = $('em-note');
  var lastCount = 0, prefilled = false, reTimer = null;
  var EMAIL_RE = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;

  function params() {
    var p = new URLSearchParams();
    if (form.resolution) p.set('resolution', form.resolution.value);
    if (form.status && form.status.value) p.set('status', form.status.value);
    if (form.q && form.q.value.trim()) p.set('q', form.q.value.trim());
    if (form.date_from && form.date_from.value) p.set('date_from', form.date_from.value);
    if (form.date_to && form.date_to.value) p.set('date_to', form.date_to.value);
    return p;
  }
  // Split a recipients field into trimmed tokens (comma / semicolon / newline).
  function splitEmails(v) {
    return (v || '').split(/[,;\n]+/).map(function (s) { return s.trim(); }).filter(Boolean);
  }
  function invalidEmails(v) {
    return splitEmails(v).filter(function (e) { return !EMAIL_RE.test(e); });
  }
  // Once the operator has typed recipients/note, send them so the render + send
  // both reflect the edits. Before the first prefill we send nothing (defaults).
  function extraParams(p) {
    if (prefilled) {
      p.set('to', toIn.value);
      p.set('cc', ccIn.value);
    }
    p.set('note', noteIn.value);
    return p;
  }
  function setStatus(html, cls) {
    var s = $('em-status'); s.innerHTML = html || '';
    s.className = 'em-status' + (cls ? ' ' + cls : '');
  }
  function close() { modal.hidden = true; setStatus(''); }

  function loadPreview(isInitial) {
    setStatus('Loading preview…');
    $('em-send').disabled = true;
    fetch(modal.getAttribute('data-preview-url') + '?' + extraParams(params()).toString(),
          { headers: { 'X-Requested-With': 'XMLHttpRequest' } })
      .then(function (r) { return r.json(); })
      .then(function (j) {
        if (!j.ok) { setStatus(j.error || 'Could not build preview.', 'err'); return; }
        // Prefill the To/Cc boxes with the config defaults on first open only —
        // never clobber what the operator has since typed.
        if (isInitial && !prefilled) {
          toIn.value = (j.to || []).join(', ');
          ccIn.value = (j.cc || []).join(', ');
          prefilled = true;
        }
        $('em-subject').textContent = j.subject || '—';
        $('em-frame').srcdoc = j.html || '';
        lastCount = j.count || 0;
        refreshSendState();
      })
      .catch(function () { setStatus('Could not load preview — please retry.', 'err'); });
  }

  // Decide whether Send is allowed + what the status line says.
  function refreshSendState() {
    var badTo = invalidEmails(toIn.value), badCc = invalidEmails(ccIn.value);
    toIn.classList.toggle('bad', badTo.length > 0);
    ccIn.classList.toggle('bad', badCc.length > 0);
    var hasTo = splitEmails(toIn.value).length > 0;
    if (!lastCount) { setStatus('No issue lines in the current filter — nothing to send.', 'err'); $('em-send').disabled = true; return; }
    if (badTo.length || badCc.length) { setStatus('Fix invalid email(s): ' + badTo.concat(badCc).join(', '), 'err'); $('em-send').disabled = true; return; }
    if (!hasTo) { setStatus('Add at least one "To" recipient.', 'err'); $('em-send').disabled = true; return; }
    setStatus(lastCount + ' line(s) will be emailed to ' + splitEmails(toIn.value).length + ' recipient(s).');
    $('em-send').disabled = false;
  }
  function scheduleRepreview() {
    clearTimeout(reTimer);
    reTimer = setTimeout(function () { loadPreview(false); }, 500);
  }

  btn.addEventListener('click', function () {
    modal.hidden = false; prefilled = false; lastCount = 0;
    $('em-frame').srcdoc = ''; noteIn.value = '';
    loadPreview(true);
  });

  // Live validation as recipients change; note re-renders the preview (debounced)
  // so the "Note from sender" block reflects the text before sending.
  toIn.addEventListener('input', refreshSendState);
  ccIn.addEventListener('input', refreshSendState);
  noteIn.addEventListener('input', scheduleRepreview);

  $('em-send').addEventListener('click', function () {
    if (invalidEmails(toIn.value).length || invalidEmails(ccIn.value).length || !splitEmails(toIn.value).length) {
      refreshSendState(); return;
    }
    $('em-send').disabled = true;
    setStatus('<span class="em-spin"></span>Sending…');
    var body = new URLSearchParams();
    body.set('to', toIn.value); body.set('cc', ccIn.value); body.set('note', noteIn.value);
    fetch(modal.getAttribute('data-send-url') + '?' + params().toString(), {
      method: 'POST', body: body.toString(),
      headers: { 'X-CSRFToken': csrf, 'X-Requested-With': 'XMLHttpRequest',
                 'Content-Type': 'application/x-www-form-urlencoded' }
    }).then(function (r) { return r.json(); }).then(function (j) {
      if (j.ok) { setStatus('✓ Sent ' + (j.count || '') + ' line(s).', 'ok'); setTimeout(close, 1500); }
      else { setStatus(j.error || 'Send failed.', 'err'); $('em-send').disabled = false; }
    }).catch(function () { setStatus('Network error — please retry.', 'err'); $('em-send').disabled = false; });
  });

  $('em-close').addEventListener('click', close);
  $('em-cancel').addEventListener('click', close);
  modal.addEventListener('click', function (e) { if (e.target === modal) close(); });
})();
