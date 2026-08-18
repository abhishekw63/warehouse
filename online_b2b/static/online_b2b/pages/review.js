/* online_b2b/review.html — page script (separated). Server URLs via #review-cfg JSON. */
var CFG = JSON.parse(document.getElementById("review-cfg").textContent);
(function () {
  var tabs = document.querySelectorAll('.tab');
  tabs.forEach(function (t) {
    t.addEventListener('click', function () {
      tabs.forEach(function (x) { x.classList.toggle('on', x === t); });
      var name = t.getAttribute('data-tab');
      document.querySelectorAll('.tabpane').forEach(function (p) {
        p.style.display = (p.getAttribute('data-pane') === name) ? '' : 'none';
      });
    });
  });

  // Applied CP: "Include (our CP)" (OVERRIDE) auto-fills OUR CP; the field stays
  // read-only — the operator doesn't price it, it IS our CP. Other actions clear it.
  document.querySelectorAll('#confirm-form tbody .act-sel').forEach(function (sel) {
    if (sel.disabled) return;                       // locked → leave as recorded
    var row = sel.closest('tr');                    // skip the bulk-bar select (not in a row)
    if (!row) return;
    var cp = row.querySelector('.act-ocp');
    function decorate() {                            // row-wide colour, visible at any scroll
      row.classList.remove('dec-included', 'dec-excluded');
      if (sel.value === 'INCLUDE' || sel.value === 'OVERRIDE') row.classList.add('dec-included');
      else if (sel.value === 'EXCLUDE') row.classList.add('dec-excluded');
    }
    function sync() {
      if (cp) {
        if (sel.value === 'OVERRIDE') cp.value = cp.getAttribute('data-ourcp') || '';        // Include (our CP)
        else if (sel.value === 'INCLUDE') cp.value = cp.getAttribute('data-vendorcp') || ''; // Include (their CP)
        else cp.value = '';                         // Exclude / undecided → nothing applied
        cp.setAttribute('readonly', 'readonly');    // read-only display of the applied CP
      }
      decorate();
    }
    sel.addEventListener('change', sync);
    decorate();                                     // paint the already-saved state on load
  });

  // ── Per-line decision auto-save (decide each line, then lock) ──────────
  var saveUrl = CFG.save;
  var csrfEl = document.querySelector('#confirm-form [name=csrfmiddlewaretoken]');
  var csrf = csrfEl ? csrfEl.value : '';

  // ✉ Email issue — send the flagged CP-issue lines to the ecom stakeholders.
  var emailBtn = document.getElementById('emailIssueBtn');
  if (emailBtn) emailBtn.addEventListener('click', function () {
    if (emailBtn.disabled) return;
    var orig = emailBtn.innerHTML;
    emailBtn.disabled = true; emailBtn.innerHTML = '✉ Sending…';
    fetch(emailBtn.dataset.url, { method: 'POST', credentials: 'same-origin',
      headers: { 'X-CSRFToken': csrf, 'X-Requested-With': 'XMLHttpRequest' } })
      .then(function (r) { return r.json(); })
      .then(function (j) {
        emailBtn.disabled = false;
        if (j.ok) {
          emailBtn.innerHTML = '✓ Emailed';
          var to = (j.to || []).join(', ');
          if (window.B2B && B2B.toast) B2B.toast('Issue email sent to ' + to + ' (' + j.lines + ' line(s)).', 'ok');
          else alert('Issue email sent to ' + to + ' (' + j.lines + ' flagged line(s)).');
          setTimeout(function () { emailBtn.innerHTML = orig; }, 4000);
        } else {
          emailBtn.innerHTML = orig;
          if (window.B2B && B2B.toast) B2B.toast(j.error || 'Could not send.', 'err');
          else alert(j.error || 'Could not send the email.');
        }
      })
      .catch(function () {
        emailBtn.disabled = false; emailBtn.innerHTML = orig;
        alert('Network error — email not sent.');
      });
  });

  function saveDecision(tr) {
    var keyEl = tr.querySelector('input[name=aff_key]');
    var sel = tr.querySelector('.act-sel');
    if (!keyEl || !sel || sel.disabled) return;
    var ocp = tr.querySelector('.act-ocp'), rem = tr.querySelector('.act-rem');
    var tick = tr.querySelector('.dec-saved');
    var body = new URLSearchParams();
    body.set('key', keyEl.value);
    body.set('action', sel.value || '');
    body.set('override_cp', ocp ? ocp.value : '');
    body.set('remark', rem ? rem.value : '');
    fetch(saveUrl, { method: 'POST', credentials: 'same-origin',
      headers: { 'X-CSRFToken': csrf, 'X-Requested-With': 'XMLHttpRequest',
                 'Content-Type': 'application/x-www-form-urlencoded' },
      body: body.toString() })
      .then(function (r) { return r.json(); })
      .then(function (j) {
        if (j.ok && tick) { tick.hidden = false; tick.classList.add('flash');
          setTimeout(function () { tick.classList.remove('flash'); }, 700); }
      }).catch(function () {});
  }
  document.querySelectorAll('#confirm-form tbody tr').forEach(function (tr) {
    if (!tr.querySelector('input[name=aff_key]')) return;
    ['.act-sel', '.act-ocp', '.act-rem'].forEach(function (s) {
      var el = tr.querySelector(s);
      if (el && !el.disabled) el.addEventListener('change', function () { saveDecision(tr); });
    });
  });

  // ── Bulk decisions — tick rows, apply one action to all of them ────────
  // (e.g. 10 NOT_IN_MASTER freebie lines → select 5, Exclude; select 5, fix EAN)
  var affChks = function () {
    return Array.prototype.slice.call(document.querySelectorAll('.aff-chk'));
  };
  var checkedRows = function () {
    return affChks().filter(function (c) { return c.checked; })
      .map(function (c) { return c.closest('tr'); });
  };
  var affCount = document.getElementById('affCount');
  function refreshCount() {
    var n = affChks().filter(function (c) { return c.checked; }).length;
    if (affCount) affCount.textContent = n;
    var all = document.getElementById('affAll');
    var boxes = affChks();
    if (all) all.checked = boxes.length > 0 && n === boxes.length;
  }
  var affAll = document.getElementById('affAll');
  if (affAll) affAll.addEventListener('change', function () {
    affChks().forEach(function (c) { c.checked = affAll.checked; });
    refreshCount();
  });
  affChks().forEach(function (c) { c.addEventListener('change', refreshCount); });

  function setSelVal(sel, val) {   // only set if that option exists on the row
    var ok = Array.prototype.some.call(sel.options, function (o) { return o.value === val; });
    if (ok) { sel.value = val; return true; }
    return false;
  }
  var bulkApply = document.getElementById('bulkApplyAct');
  if (bulkApply) bulkApply.addEventListener('click', function () {
    var act = document.getElementById('bulkAction').value;
    var rows = checkedRows();
    if (!act || !rows.length) { return; }
    var applied = 0, skipped = 0;
    rows.forEach(function (tr) {
      var sel = tr.querySelector('[name=aff_action]');
      if (!sel || sel.disabled) return;
      if (!setSelVal(sel, act)) { skipped++; return; }   // e.g. INCLUDE on a NOT_IN_MASTER row
      var rowOcp = tr.querySelector('.act-ocp');
      if (rowOcp) {
        // Applied CP: Include (our CP) → OUR CP · Include (their CP) → THEIR CP ·
        // Exclude → cleared. Read-only display of what will actually be applied.
        rowOcp.value = (act === 'OVERRIDE') ? (rowOcp.getAttribute('data-ourcp') || '')
                     : (act === 'INCLUDE') ? (rowOcp.getAttribute('data-vendorcp') || '')
                     : '';
        rowOcp.setAttribute('readonly', 'readonly');
      }
      tr.classList.remove('dec-included', 'dec-excluded');   // row-wide decision colour
      if (act === 'INCLUDE' || act === 'OVERRIDE') tr.classList.add('dec-included');
      else if (act === 'EXCLUDE') tr.classList.add('dec-excluded');
      saveDecision(tr);
      applied++;
    });
    var msg = document.getElementById('affBulkMsg');
    if (msg) { msg.textContent = '✓ ' + applied + ' set' + (skipped ? ' · ' + skipped + ' skipped (n/a)' : ''); msg.hidden = false;
      setTimeout(function () { msg.hidden = true; }, 2200); }
  });
  var bulkFill = document.getElementById('bulkFillEan');
  if (bulkFill) bulkFill.addEventListener('click', function () {
    var ean = document.getElementById('bulkEan').value.trim();
    if (!ean) { return; }
    var filled = 0;
    checkedRows().forEach(function (tr) {
      var nim = tr.querySelector('.nim-input');
      if (nim && !nim.readOnly) { nim.value = ean; filled++; }
    });
    var msg = document.getElementById('affBulkMsg');
    if (msg) { msg.textContent = '✓ EAN filled into ' + filled + ' row(s) — now “Apply & re-validate”'; msg.hidden = false;
      setTimeout(function () { msg.hidden = true; }, 3000); }
  });

  // ── NOTHING progresses until you click a button ───────────────────────
  // 1) Block Enter on EVERY field (both forms) so typing/choosing never
  //    submits. Enter just blurs → which saves that line's decision.
  document.querySelectorAll('.b2b-wrap input, .b2b-wrap select, .b2b-wrap textarea').forEach(function (el) {
    el.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') { e.preventDefault(); el.blur(); }
    });
  });
  // 2) Belt-and-suspenders: block IMPLICIT submits (Enter with no button).
  //    SubmitEvent.submitter is null only for implicit submission; a real
  //    button click (Lock / Discard / Generate, incl. ones injected after an
  //    in-place lock) sets it, so they always work.
  var _cf = document.getElementById('confirm-form');
  if (_cf) {
    _cf.addEventListener('submit', function (e) {
      // Lock is AJAX-ONLY (click → fetch). Block every native submit of this
      // form EXCEPT the intentional formaction buttons (Discard / Generate):
      //   • implicit submit (Enter, submitter == null) → blocked
      //   • the Lock button itself → blocked (so it can never post /confirm/
      //     natively and "suddenly" lock without the progress overlay)
      var b = e.submitter;
      if (b == null || b.id === 'lockBtn' || !b.hasAttribute('formaction')) {
        e.preventDefault();
      }
    });
  }

  // ── "Apply EAN fixes & re-validate" — show progress (full reload re-runs the
  //    engine, so without feedback it feels stuck) ──
  var applyBtn = document.querySelector('.aff-eanbar button[type=submit]');
  if (applyBtn) {
    var _applying = false;
    applyBtn.addEventListener('click', function (e) {
      if (_applying) { e.preventDefault(); return; }   // ignore double-clicks
      _applying = true;
      applyBtn.classList.add('btn-loading');
      applyBtn.innerHTML = '<span class="btn-spin"></span> Re-validating…';
      // dim + blur + lock the whole page until the reload completes
      var ov = document.getElementById('revalOverlay');
      if (ov) {
        ov.classList.add('show'); document.body.classList.add('lo-open');
        var el = document.getElementById('revalElapsed'), t0 = Date.now();
        setInterval(function () {
          var s = Math.floor((Date.now() - t0) / 1000);
          if (el) el.textContent = 'elapsed ' + Math.floor(s / 60) + ':' + ('0' + (s % 60)).slice(-2);
        }, 250);
      }
    });
  }

  // ── Lock & Record via AJAX — progress bar, no page reload ──────────────
  var form    = document.getElementById('confirm-form');
  var lockBtn = document.getElementById('lockBtn');
  var ov      = document.getElementById('lockOverlay');
  var fill    = document.getElementById('loFill');
  var errBox  = document.getElementById('loErr');
  var steps   = ov ? ov.querySelectorAll('.lo-steps li') : [];
  var subBox  = document.getElementById('loSub');
  var noteBox = document.getElementById('loNote');
  var timers  = [];
  // Total is known up-front (from the preview) — show it so the operator sees
  // scale + an elapsed timer, instead of a blank bar that reads as "hung".
  var LO_POS   = parseInt((ov && ov.getAttribute('data-total')) || '0', 10) || 0;
  var LO_LINES = parseInt((ov && ov.getAttribute('data-lines')) || '0', 10) || 0;
  var _animTimer = null, _elapTimer = null, _startT = 0, _curFill = 0;

  function setFill(p) { _curFill = p; if (fill) fill.style.width = p + '%'; }
  function clearTimers() {
    timers.forEach(clearTimeout); timers = [];
    if (_animTimer) { clearInterval(_animTimer); _animTimer = null; }
    if (_elapTimer) { clearInterval(_elapTimer); _elapTimer = null; }
  }
  function _fmtEl(ms) { var s = Math.floor(ms / 1000); return s < 60 ? (s + 's') : (Math.floor(s / 60) + 'm ' + (s % 60) + 's'); }
  function _label(verb) {
    return LO_POS ? (verb + ' ' + LO_POS + ' PO' + (LO_POS > 1 ? 's' : '') +
      (LO_LINES ? (' · ' + LO_LINES + ' lines') : '')) : (verb + '…');
  }
  function _renderSub() {
    if (subBox) subBox.textContent = _label('Recording') + '  ·  ' + _fmtEl(Date.now() - _startT);
    if (noteBox && (Date.now() - _startT) > 8000) noteBox.hidden = false;   // reassure on long runs
  }

  function startProgress() {
    ov.classList.remove('done', 'error');
    errBox.textContent = '';
    if (noteBox) noteBox.hidden = true;
    steps.forEach(function (s) { s.classList.remove('active', 'done'); });
    ov.classList.add('show');
    document.body.classList.add('lo-open');
    _startT = Date.now();
    setFill(8);
    if (steps[0]) steps[0].classList.add('active');
    // stamp the real line count into the "Writing line items" step
    var lnEl = ov.querySelector('.lo-lines');
    if (lnEl && LO_LINES) lnEl.textContent = 'Writing ' + LO_LINES.toLocaleString('en-IN') + ' line items';
    _renderSub();
    _elapTimer = setInterval(_renderSub, 1000);
    // Asymptotic bar: eases toward a ceiling it NEVER reaches until the real
    // server response lands (finishProgress). Always visibly moving, never lies
    // about being done. The visual steps advance as the fill crosses thresholds,
    // so they track real elapsed time — no fixed fake schedule.
    _animTimer = setInterval(function () {
      var ceil = 93, next = _curFill + (ceil - _curFill) * 0.045;
      setFill(next > ceil ? ceil : next);
      // advance the visual steps as the fill crosses evenly-spaced thresholds —
      // works for any number of steps (currently 6).
      var _n = steps.length, _ceil = 93;
      for (var _i = 1; _i < _n; _i++) {
        if (_curFill > (_i / _n) * _ceil && steps[_i] && !steps[_i].classList.contains('active')) {
          steps[_i - 1].classList.add('done'); steps[_i].classList.add('active');
        }
      }
    }, 220);
  }

  function finishProgress() {
    clearTimers();
    steps.forEach(function (s) { s.classList.remove('active'); s.classList.add('done'); });
    setFill(100);
    if (subBox) subBox.textContent = _label('Recorded') + '  ·  ' + _fmtEl(Date.now() - _startT);
    if (noteBox) noteBox.hidden = true;
    ov.classList.add('done');
  }

  function failProgress(msg) {
    clearTimers();
    if (noteBox) noteBox.hidden = true;
    var text = msg || 'Lock failed — nothing was recorded. Please try again.';
    ov.classList.add('error');
    errBox.textContent = text;
    // Persistent red toast too — the overlay error auto-dismisses in 2.4s, but
    // the toast lingers so a rolled-back/atomic-failed lock can't be missed.
    if (window.B2B && B2B.toast) B2B.toast(text, {
      type: 'error', title: 'Not recorded — safe to retry', timeout: 10000 });
    timers.push(setTimeout(function () {
      ov.classList.remove('show', 'error');
      document.body.classList.remove('lo-open');
      if (lockBtn) lockBtn.disabled = false;
    }, 2400));
  }

  function morphLocked(j) {
    var bar = document.getElementById('confirmBar');
    if (bar) {
      bar.classList.add('locked');
      var txt = bar.querySelector('.cb-text');
      var act = bar.querySelector('.cb-actions');
      if (txt) txt.innerHTML = '🔒 <b>Decisions locked &amp; recorded</b> (run #' + j.run_id +
        '). Generate the ERP <b>D365 dump</b>, or download the <b>Completed</b> SO Workbook (accepted lines only — Excludes dropped, Overrides repriced).';
      if (act) {
        // Swap the download button to the COMPLETED route. The pre-lock button
        // pointed at the Review/full workbook; it must NOT persist after an
        // in-place lock (that bug served the full file post-lock).
        var completedDl = '<a href="' + COMPLETED_DL_URL + '" id="dlWorkbook" class="btn-ghost" ' +
          'data-download-bg="SO Workbook is downloading in the background — check your browser Downloads. You can keep working." title="Completed — accepted lines pushed to ' +
          'D365 (Excludes dropped, Overrides repriced)">⬇ Download SO Workbook (Completed)</a>';
        act.innerHTML = completedDl +
          ' <a href="' + j.run_url + '" class="btn-ghost">View run →</a>' +
          ' <button class="btn-primary" type="submit" formaction="' + j.d365_url +
          '" formnovalidate onclick="var b=this;b.innerHTML=\'<span class=&quot;btn-spin&quot;></span> Generating…\';setTimeout(function(){b.innerHTML=\'⬇ Generate D365 dump\';},4500);">⬇ Generate D365 dump</button>';
      }
    }
    document.querySelectorAll('.act-sel').forEach(function (s) { s.disabled = true; });
    document.querySelectorAll('.act-ocp,.act-rem,.nim-input').forEach(function (x) { x.setAttribute('readonly', 'readonly'); });
    // Fully match the server's locked render — an in-place lock must leave NOTHING
    // editable. The confirm-bar + action selects above were handled; also strip the
    // draft banner, the bulk "apply to selected" bar, the EAN-fix bar, and every
    // row checkbox — otherwise the page still LOOKS editable after recording.
    var _db = document.getElementById('draftBanner'); if (_db) _db.remove();
    document.querySelectorAll('#affBulk, .aff-eanbar').forEach(function (x) { x.remove(); });
    document.querySelectorAll('#affAll, .aff-chk').forEach(function (x) { x.remove(); });
  }

  var HUB_URL = CFG.hub;
  var UPLOAD_URL = CFG.upload;
  var COMPLETED_DL_URL = CFG.completed_dl;

  // Kick off the Completed SO Workbook via a browser-NATIVE download (hidden
  // <a download> handed to the browser's own download manager) instead of a
  // page-owned blob fetch. The old blob fetch was cancelled the instant the user
  // closed the popup or navigated away; a native download keeps running in the
  // browser independently — so the operator can immediately move to other pages.
  function autoDownload(url, label) {
    if (window.B2B && B2B.bgDownload) B2B.bgDownload(url);
    else {   // ultra-safe fallback if the shared helper isn't present
      var t = document.createElement('a');
      t.href = url; t.setAttribute('download', ''); t.style.display = 'none';
      document.body.appendChild(t); t.click();
      setTimeout(function () { t.remove(); }, 2000);
    }
    if (window.B2B && B2B.toast) B2B.toast(
      'SO Workbook is downloading in the background — check your browser Downloads. You can keep working.',
      { type: 'info', title: 'SO Workbook', timeout: 9000 });
  }

  function flourish(j) {
    // No blocking modal — the page has ALREADY morphed to its locked state
    // (morphLocked): the run banner + Download SO Workbook (Completed) / View run
    // / Generate D365 dump buttons are right there. So just: kick off the auto
    // download, celebrate, and confirm with a non-blocking toast. Zero clicks.
    autoDownload(COMPLETED_DL_URL, 'Building the Completed SO Workbook');
    if (window.B2B && B2B.celebrate) B2B.celebrate();
    if (window.B2B && B2B.toast) B2B.toast(
      'Run #' + j.run_id + ' · ' + j.pos + ' PO(s) · ' + j.lines +
        ' line(s) pushed to D365. Download buttons are ready below.',
      { type: 'ok', title: 'Locked & recorded ✓', timeout: 7000 });
  }

  // The actual lock+record (progress bar → AJAX). Called once all guards pass.
  function doLock() {
    lockBtn.disabled = true;
    startProgress();
    var t0 = Date.now();
    fetch(form.action, {
      method: 'POST',
      headers: { 'X-Requested-With': 'XMLHttpRequest' },
      body: new FormData(form),
      credentials: 'same-origin'
    }).then(function (r) {
      return r.json().then(function (j) { return { ok: r.ok, j: j }; });
    }).then(function (res) {
      var j = res.j || {};
      if (!res.ok || !j.ok) { failProgress(j.error); return; }
      if (j.redirect) { finishProgress(); setTimeout(function () { location.href = j.redirect; }, 500); return; }
      var wait = Math.max(0, 1150 - (Date.now() - t0));  // min on-screen time
      setTimeout(function () {
        finishProgress();
        setTimeout(function () {
          ov.classList.remove('show');
          document.body.classList.remove('lo-open');
          morphLocked(j);
          flourish(j);
        }, 580);
      }, wait);
    }).catch(function () {
      failProgress('Network error — nothing was locked. Try again.');
    });
  }

  // ── CP / undecided lock-guard ─────────────────────────────────────────
  // Any affected line with NO decision (Include/Override/Exclude) blocks the
  // lock: shake the button, flash the pending row(s), and ASK — Save for Review
  // Later, go decide, or Record anyway. Never silently records undecided lines.
  function undecidedSelects() {
    return Array.prototype.slice.call(
      document.querySelectorAll('#confirm-form select[name=aff_action]'))
      .filter(function (s) { return !s.disabled && (s.value || '').trim() === ''; });
  }
  function cpGuard(undec) {
    lockBtn.classList.remove('cp-shake'); void lockBtn.offsetWidth;
    lockBtn.classList.add('cp-shake');
    var affTab = document.querySelector('.tab[data-tab="affected"]');
    if (affTab) affTab.click();
    undec.forEach(function (s) {
      var tr = s.closest('tr');
      if (tr) { tr.classList.remove('cp-flash'); void tr.offsetWidth; tr.classList.add('cp-flash'); }
    });
    var first = undec[0].closest('tr');
    if (first) first.scrollIntoView({ behavior: 'smooth', block: 'center' });
    showCpModal(undec.length);
  }
  function showCpModal(n) {
    if (document.querySelector('.cpg-ov')) return;
    var ov2 = document.createElement('div');
    ov2.className = 'cpg-ov';
    ov2.innerHTML =
      '<div class="cpg-card"><h3>⚠ ' + n + ' line(s) still need a decision</h3>' +
      '<p>These lines have no Include / Override / Exclude set — e.g. an unresolved CP mismatch. What do you want to do?</p>' +
      '<div class="cpg-actions">' +
        '<button class="cpg-btn cpg-later" data-a="later">🕒 Save for Review Later' +
          '<small>Park the whole run — decide after the team corrects the price. No re-upload.</small></button>' +
        '<button class="cpg-btn cpg-cancel" data-a="decide">Go decide them now' +
          '<small>Jump to the flashing line(s) on the Affected tab.</small></button>' +
        '<button class="cpg-btn cpg-anyway" data-a="anyway">Lock &amp; Record anyway' +
          '<small>Include the ' + n + ' undecided line(s) as-is (still flagged in the DB).</small></button>' +
      '</div></div>';
    document.body.appendChild(ov2);
    ov2.addEventListener('click', function (e) {
      var btn = e.target.closest && e.target.closest('[data-a]');
      var a = btn ? btn.getAttribute('data-a') : null;
      if (e.target === ov2 || a === 'decide') { ov2.remove(); return; }
      if (a === 'later') { document.getElementById('save-later-form').submit(); return; }
      if (a === 'anyway') { ov2.remove(); lockBtn._forced = true; lockBtn.click(); }
    });
  }

  if (lockBtn && form && window.fetch) {
    lockBtn.addEventListener('click', function (e) {
      e.preventDefault();
      if (lockBtn.disabled) return;
      // Guard: a Correct EAN typed but NOT applied would be silently lost on
      // lock (the line would lock as NOT_IN_MASTER). Make them apply it first.
      var typed = Array.prototype.slice.call(document.querySelectorAll('.nim-input'))
        .filter(function (i) { return (i.value || '').trim(); });
      if (typed.length && !window.confirm(
          'You typed a Correct EAN but did NOT click "Apply EAN fixes & re-validate".\n\n' +
          'If you lock now it will NOT be saved — the line stays NOT_IN_MASTER.\n\n' +
          'Lock anyway? (Cancel to go back and apply the fix first.)')) {
        return;
      }
      // Guard: undecided affected line(s) → shake + flash + ask (unless the
      // operator already chose "Record anyway"). The server enforces the SAME
      // rule; the hidden flag tells it this was the deliberate "anyway" path.
      var forced = !!lockBtn._forced;
      var inclEl = document.getElementById('inclUndecided');
      if (inclEl) inclEl.value = forced ? '1' : '';   // explicit escape only
      if (!forced) {
        var undec = undecidedSelects();
        if (undec.length) { cpGuard(undec); return; }
      }
      lockBtn._forced = false;
      doLock();
    });
  }
})();
