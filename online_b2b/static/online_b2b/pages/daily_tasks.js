/* online_b2b/daily_tasks.html — page script (separated). Server values via #daily_tasks-cfg. */
var CFG = JSON.parse(document.getElementById("daily_tasks-cfg").textContent);
(function () {
  var day = document.getElementById('dt-day').value;
  var toggleUrl = CFG.toggle;
  // Calendar: jump to the picked date (query-param nav, no future dates).
  var cal = document.getElementById('dt-cal');
  if (cal) cal.addEventListener('change', function () {
    if (cal.value) window.location.search = '?day=' + cal.value;
  });
  var STEPS = CFG.steps;
  var TOTAL = CFG.total;

  function recompute() {
    var done = 0, nopo = 0, held = 0;
    document.querySelectorAll('tr[data-channel]').forEach(function (tr) {
      if (tr.classList.contains('dt-hold')) { held++; return; }   // parked = handled
      if (tr.classList.contains('dt-nopo')) { nopo++; return; }
      var boxes = tr.querySelectorAll('.dt-cell input[type=checkbox]');
      var d = 0; boxes.forEach(function (b) { if (b.checked) d++; });
      var pct = boxes.length ? Math.round(d * 100 / boxes.length) : 0;
      var bar = tr.querySelector('.dt-pbar span'); if (bar) bar.style.width = pct + '%';
      var full = d === boxes.length;
      var dot = tr.querySelector('.dt-dot');
      if (dot) dot.className = 'dt-dot' + (full ? ' ok' : (d ? ' part' : ''));
      if (full) { done++; tr.classList.add('dt-done'); } else tr.classList.remove('dt-done');
    });
    var handled = done + nopo + held;
    document.getElementById('dt-ovfill').style.width = Math.round(handled * 100 / TOTAL) + '%';
    document.getElementById('dt-ovpct').textContent = Math.round(handled * 100 / TOTAL) + '%';
    document.getElementById('dt-ovdone').textContent = done;
    document.getElementById('dt-ovnopo').textContent = nopo;
    document.getElementById('dt-ovpend').textContent = TOTAL - handled;
    recomputeParents();
  }

  // Roll each parent (e.g. MT Select) up from its children: mean pct on the bar,
  // "X/N done" count. A child is "done" when its row carries dt-done; "handled"
  // when done OR no-PO OR on-hold (mirrors the server rollup).
  function recomputeParents() {
    document.querySelectorAll('tr.dt-parent').forEach(function (ptr) {
      var pkey = ptr.dataset.parent;
      var kids = document.querySelectorAll('tr[data-parent-of="' + pkey + '"]');
      if (!kids.length) return;
      var doneK = 0, handledK = 0, sumPct = 0;
      kids.forEach(function (tr) {
        var d = tr.classList.contains('dt-done');
        var nopo = tr.classList.contains('dt-nopo');
        var hold = tr.classList.contains('dt-hold');
        if (d) doneK++;
        if (d || nopo || hold) handledK++;
        var pct;
        if (nopo || hold) { pct = 100; }
        else {
          var boxes = tr.querySelectorAll('.dt-cell input[type=checkbox]');
          var c = 0; boxes.forEach(function (b) { if (b.checked) c++; });
          pct = boxes.length ? Math.round(c * 100 / boxes.length) : 0;
        }
        sumPct += pct;
      });
      var mean = Math.round(sumPct / kids.length);
      var bar = ptr.querySelector('.dt-parent-bar'); if (bar) bar.style.width = mean + '%';
      var pc = ptr.querySelector('.dt-pc-done'); if (pc) pc.textContent = doneK;
    });
  }

  function post(cb, onOk, extra) {
    var want = cb.checked; cb.disabled = true;
    var body = new URLSearchParams({ day: day, channel: cb.dataset.channel,
      step: cb.dataset.step, checked: want ? '1' : '0' });
    if (extra) { Object.keys(extra).forEach(function (k) { body.set(k, extra[k]); }); }
    B2B.postForm(toggleUrl, body)
      .then(function (j) {
        cb.disabled = false;
        if (!j.ok) { cb.checked = !want; alert(j.error || 'Could not save.'); return; }
        if (onOk) onOk(j); recompute();
      })
      .catch(function () { cb.disabled = false; cb.checked = !want; alert('Network error — please retry.'); });
  }

  // work-step checkboxes (all cells; auto/disabled ones never fire change)
  document.querySelectorAll('.dt-cell input[type=checkbox]').forEach(function (cb) {
    cb.addEventListener('change', function () {
      post(cb, function (j) {
        var t = document.querySelector('.dt-time[data-cell="' + cb.dataset.channel + ':' + cb.dataset.step + '"]');
        if (t) t.textContent = j.checked ? (j.at + (j.by ? ' · ' + j.by : '')) : '';
      });
    });
  });

  // "No PO today" toggles — mark handled without ticking the steps
  document.querySelectorAll('.dt-nopo-cb').forEach(function (cb) {
    cb.addEventListener('change', function () {
      var on = cb.checked;
      post(cb, function () {
        var tr = cb.closest('tr');
        tr.classList.toggle('dt-nopo', on);
        tr.querySelectorAll('.dt-cell input[type=checkbox]').forEach(function (b) {
          var isAuto = b.closest('.dt-box').classList.contains('auto');
          b.disabled = on || isAuto;
        });
        var prog = tr.querySelector('.dt-prog');
        if (prog) prog.innerHTML = on ? '<span class="dt-nopo-lbl">— no PO</span>'
                                      : '<div class="dt-pbar"><span style="width:0%"></span></div>';
      });
    });
  });

  // "Hold" toggles — park the channel (e.g. an unresolved CP issue). Update the
  // badge, greyed steps, progress cell + counts IN PLACE (no page refresh).
  document.querySelectorAll('.dt-hold-cb').forEach(function (cb) {
    cb.addEventListener('change', function () {
      var on = cb.checked;
      post(cb, function (j) {
        var tr = cb.closest('tr');
        tr.classList.toggle('dt-hold', on);
        if (on) tr.classList.remove('dt-done');
        // ON HOLD badge in the channel cell (before the "no PO today" label)
        var cell = tr.querySelector('.dt-ch');
        var badge = cell.querySelector('.dt-hold-badge');
        if (on && !badge) {
          badge = document.createElement('span');
          badge.className = 'dt-hold-badge';
          badge.textContent = '⏸ ON HOLD';
          cell.insertBefore(badge, cell.querySelector('.dt-nopo-t'));
        } else if (!on && badge) { badge.remove(); }
        if (badge) { badge.title = 'On hold' + (j && j.at ? ' since ' + j.at : ''); }
        // progress cell: hold label + inline reason field ↔ progress bar
        var prog = tr.querySelector('td.dt-prog');
        if (prog) {
          if (on) {
            prog.innerHTML = '<span class="dt-hold-lbl">⏸ on hold' + (j && j.at ? ' · ' + j.at : '') + '</span>'
              + '<input type="text" class="dt-hold-rin" data-channel="' + cb.dataset.channel
              + '" maxlength="500" value="" placeholder="＋ add reason…"'
              + ' title="Type a reason for the hold, then press Enter">';
            var rin = prog.querySelector('.dt-hold-rin');
            if (rin) rin.focus();          // ready to type immediately, no dialog
          } else {
            prog.innerHTML = '<div class="dt-pbar"><span style="width:0%"></span></div>';
          }
        }
        recompute();
        // recompute skips held rows, so set this row's dot explicitly on hold
        if (on) { var dot = tr.querySelector('.dt-dot'); if (dot) dot.className = 'dt-dot hold'; }
      });
    });
  });

  // Inline hold-reason field — save on Enter or blur (like the My Tasks adder).
  var holdReasonUrl = CFG.hold_reason;
  function saveHoldReason(rin) {
    var val = rin.value.trim();
    if (val === rin.dataset.saved) return;          // unchanged → skip
    rin.dataset.saved = val;
    B2B.postForm(holdReasonUrl, new URLSearchParams({ day: day, channel: rin.dataset.channel, remark: val }))
      .then(function (j) {
        if (!j.ok) { alert(j.error || 'Could not save the reason.'); return; }
        rin.classList.toggle('saved', !!val);
        var badge = rin.closest('tr').querySelector('.dt-hold-badge');
        if (badge) badge.title = 'On hold' + (val ? ' — ' + val : '');
      })
      .catch(function () { alert('Network error — please retry.'); });
  }
  document.addEventListener('keydown', function (ev) {
    if (ev.key === 'Enter' && ev.target.classList.contains('dt-hold-rin')) {
      ev.preventDefault(); ev.target.blur();
    }
  });
  document.addEventListener('blur', function (ev) {
    if (ev.target.classList && ev.target.classList.contains('dt-hold-rin')) saveHoldReason(ev.target);
  }, true);

  // Expandable parent (MT Select): click the header row to slide its children
  // open/closed. Default collapsed. Clicks on the child checkboxes never bubble
  // here (children are separate rows), so only the header toggles.
  function toggleParent(ptr) {
    var open = ptr.getAttribute('aria-expanded') === 'true';
    var pkey = ptr.dataset.parent;
    var kids = document.querySelectorAll('tr[data-parent-of="' + pkey + '"]');
    ptr.setAttribute('aria-expanded', open ? 'false' : 'true');
    kids.forEach(function (tr) {
      if (open) {
        tr.style.display = 'none';
      } else {
        tr.style.display = '';
        // brief fade/slide-in
        tr.style.opacity = '0';
        requestAnimationFrame(function () {
          tr.style.transition = 'opacity .2s ease';
          tr.style.opacity = '1';
        });
      }
    });
  }
  document.querySelectorAll('tr.dt-parent').forEach(function (ptr) {
    ptr.addEventListener('click', function () { toggleParent(ptr); });
  });

  // Seed inline hold-reason fields with their server value so blur won't re-POST
  // an unchanged reason, and show the saved (non-editing) style when filled.
  document.querySelectorAll('.dt-hold-rin').forEach(function (rin) {
    rin.dataset.saved = rin.value.trim();
    rin.classList.toggle('saved', !!rin.value.trim());
  });

  recomputeParents();   // seed the parent bars/counts from server-rendered state

  // ── My Tasks (personal ad-hoc list) ────────────────────────────────────────
  var adAddUrl = CFG.ad_add;
  var adTogUrl = CFG.ad_tog;
  var adDelUrl = CFG.ad_del;

  function adPost(url, data) {
    return B2B.postForm(url, new URLSearchParams(data));
  }
  function esc(s) { var d = document.createElement('div'); d.textContent = s; return d.innerHTML; }
  function adRefreshEmpty() {
    var openList = document.getElementById('dt-ad-open');
    var empty = openList.querySelector('.dt-ad-empty');
    var any = openList.querySelectorAll('.dt-ad-row').length > 0;
    if (empty) empty.style.display = any ? 'none' : '';
  }
  function adBind(row) {
    var cb = row.querySelector('.dt-ad-cb');
    var del = row.querySelector('.dt-ad-del');
    if (cb) cb.addEventListener('change', function () {
      var done = cb.checked; cb.disabled = true;
      adPost(adTogUrl, { id: row.dataset.id, done: done ? '1' : '0' }).then(function (j) {
        cb.disabled = false;
        if (!j.ok) { cb.checked = !done; alert('Could not save.'); return; }
        row.remove();
        var target = done ? document.getElementById('dt-ad-done') : document.getElementById('dt-ad-open');
        row.classList.toggle('done', done);
        row.classList.remove('od');
        // strip due/overdue meta styling; show the done/added stamp
        var meta = row.querySelector('.dt-ad-meta');
        if (done && meta) meta.innerHTML = '<span class="dt-ad-age">done ' + (j.done_at || 'just now') + '</span>';
        if (done) { document.querySelector('.dt-ad-done-wrap').open = true; }
        target.insertBefore(row, done ? target.firstChild : target.querySelector('.dt-ad-empty'));
        var dn = document.getElementById('dt-ad-donen');
        if (dn) dn.textContent = document.querySelectorAll('#dt-ad-done .dt-ad-row').length;
        adRefreshEmpty();
      }).catch(function () { cb.disabled = false; cb.checked = !done; alert('Network error.'); });
    });
    if (del) del.addEventListener('click', function () {
      if (!confirm('Delete this task?')) return;
      adPost(adDelUrl, { id: row.dataset.id }).then(function (j) {
        if (j.ok) { row.remove(); adRefreshEmpty();
          var dn = document.getElementById('dt-ad-donen');
          if (dn) dn.textContent = document.querySelectorAll('#dt-ad-done .dt-ad-row').length; }
      });
    });
  }
  document.querySelectorAll('#dt-adhoc .dt-ad-row').forEach(adBind);

  var addForm = document.getElementById('dt-ad-add');
  if (addForm) addForm.addEventListener('submit', function (e) {
    e.preventDefault();
    var titleEl = document.getElementById('dt-ad-title');
    var dueEl = document.getElementById('dt-ad-due');
    var title = titleEl.value.trim();
    if (!title) { titleEl.focus(); return; }
    var btn = addForm.querySelector('.dt-ad-btn'); btn.disabled = true;
    adPost(adAddUrl, { title: title, due: dueEl.value || '' }).then(function (j) {
      btn.disabled = false;
      if (!j.ok) { alert(j.error || 'Could not add.'); return; }
      var due = dueEl.value || '';
      var overdue = due && due < day;
      var row = document.createElement('div');
      row.className = 'dt-ad-row' + (overdue ? ' od' : '');
      row.dataset.id = j.id;
      row.innerHTML =
        '<label class="dt-ad-chk"><input type="checkbox" class="dt-ad-cb"><span class="dt-ad-mark"></span></label>' +
        '<div class="dt-ad-body"><div class="dt-ad-t">' + esc(title) + '</div>' +
        '<div class="dt-ad-meta">' +
          (due ? '<span class="dt-ad-due-b' + (overdue ? ' od' : '') + '">📅 ' + due + (overdue ? ' · overdue' : '') + '</span>' : '') +
          '<span class="dt-ad-age">added just now</span></div></div>' +
        '<button type="button" class="dt-ad-del" title="Delete">✕</button>';
      var openList = document.getElementById('dt-ad-open');
      openList.insertBefore(row, openList.querySelector('.dt-ad-empty'));
      adBind(row);
      adRefreshEmpty();
      titleEl.value = ''; dueEl.value = ''; titleEl.focus();
    }).catch(function () { btn.disabled = false; alert('Network error — please retry.'); });
  });

  // ── Email daily activity to senior (preview → send) ─────────────────────────
  (function () {
    var modal = document.getElementById('email-modal');
    var openBtn = document.getElementById('dt-email');
    if (!modal || !openBtn) return;
    var $ = function (id) { return document.getElementById(id); };
    var toIn = $('em-to-in'), ccIn = $('em-cc-in'), noteIn = $('em-note');
    var lastCount = 0, prefilled = false, reTimer = null;
    var EMAIL_RE = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;

    function dayParams() { var p = new URLSearchParams(); if (day) p.set('day', day); return p; }
    function splitEmails(v) { return (v || '').split(/[,;\n]+/).map(function (s) { return s.trim(); }).filter(Boolean); }
    function invalidEmails(v) { return splitEmails(v).filter(function (e) { return !EMAIL_RE.test(e); }); }
    function extraParams(p) { if (prefilled) { p.set('to', toIn.value); p.set('cc', ccIn.value); } p.set('note', noteIn.value); return p; }
    function setStatus(html, cls) { var s = $('em-status'); s.innerHTML = html || ''; s.className = 'em-status' + (cls ? ' ' + cls : ''); }
    function close() { modal.hidden = true; setStatus(''); }

    function loadPreview(isInitial) {
      setStatus('Loading preview…'); $('em-send').disabled = true;
      fetch(modal.getAttribute('data-preview-url') + '?' + extraParams(dayParams()).toString(),
            { headers: { 'X-Requested-With': 'XMLHttpRequest' } })
        .then(function (r) { return r.json(); })
        .then(function (j) {
          if (!j.ok) { setStatus(j.error || 'Could not build preview.', 'err'); return; }
          if (isInitial && !prefilled) {
            toIn.value = (j.to || []).join(', '); ccIn.value = (j.cc || []).join(', '); prefilled = true;
          }
          $('em-subject').textContent = j.subject || '—';
          $('em-frame').srcdoc = j.html || '';
          lastCount = j.count || 0; refreshSendState();
        })
        .catch(function () { setStatus('Could not load preview — please retry.', 'err'); });
    }
    function refreshSendState() {
      var badTo = invalidEmails(toIn.value), badCc = invalidEmails(ccIn.value);
      toIn.classList.toggle('bad', badTo.length > 0); ccIn.classList.toggle('bad', badCc.length > 0);
      var hasTo = splitEmails(toIn.value).length > 0;
      if (badTo.length || badCc.length) { setStatus('Fix invalid email(s): ' + badTo.concat(badCc).join(', '), 'err'); $('em-send').disabled = true; return; }
      if (!hasTo) { setStatus('Add your senior’s "To" email.', 'err'); $('em-send').disabled = true; return; }
      setStatus(lastCount + ' active channel(s) will be emailed to ' + splitEmails(toIn.value).length + ' recipient(s).');
      $('em-send').disabled = false;
    }
    function scheduleRepreview() { clearTimeout(reTimer); reTimer = setTimeout(function () { loadPreview(false); }, 500); }

    openBtn.addEventListener('click', function () {
      modal.hidden = false; prefilled = false; lastCount = 0;
      $('em-frame').srcdoc = ''; noteIn.value = ''; loadPreview(true);
    });
    $('em-close').addEventListener('click', close);
    $('em-cancel').addEventListener('click', close);
    modal.addEventListener('click', function (e) { if (e.target === modal) close(); });
    toIn.addEventListener('input', refreshSendState);
    ccIn.addEventListener('input', refreshSendState);
    noteIn.addEventListener('input', scheduleRepreview);

    $('em-send').addEventListener('click', function () {
      if (invalidEmails(toIn.value).length || invalidEmails(ccIn.value).length || !splitEmails(toIn.value).length) { refreshSendState(); return; }
      $('em-send').disabled = true; setStatus('<span class="em-spin"></span>Sending…');
      var body = new URLSearchParams();
      body.set('to', toIn.value); body.set('cc', ccIn.value); body.set('note', noteIn.value);
      B2B.postForm(modal.getAttribute('data-send-url') + '?' + dayParams().toString(), body)
        .then(function (j) {
        if (j.ok) { setStatus('✓ Sent to your senior.', 'ok'); setTimeout(close, 1500); }
        else { setStatus(j.error || 'Send failed.', 'err'); $('em-send').disabled = false; }
      }).catch(function () { setStatus('Network error — please retry.', 'err'); $('em-send').disabled = false; });
    });
  })();
})();
