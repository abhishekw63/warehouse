/* online_b2b/email.html — page script (separated from template).
   Server values (day/seg + URLs) come from the #email-cfg JSON block. */
var CFG = JSON.parse(document.getElementById("email-cfg").textContent);

// Drill-down (MP → PO → SKU). MP→PO is client-side (POs already in DOM, hidden).
// PO→SKU is LAZY: SKUs are fetched on click and injected — never pre-rendered — so
// the board stays light and a MP click can NEVER over-expand into SKUs.
var EM_DAY = CFG.day;
var EM_SEG = CFG.seg;
var EM_SKUS_URL = CFG.skus_url;

function emSetHidden(gid, hidden) {
  var rows = document.querySelectorAll('tr[data-child="' + gid + '"]');
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    row.hidden = hidden;
    if (!hidden) {
      row.classList.remove('em-anim'); void 0; row.classList.add('em-anim');
    } else {
      row.classList.remove('em-anim');
      var childGid = row.getAttribute('data-gid');   // this row is itself a parent?
      if (childGid) {                                 // collapsing → close its subtree
        emSetHidden(childGid, true);
        var a = document.getElementById('arw-' + childGid);
        if (a) a.classList.remove('open');
      }
    }
  }
}
// MP toggle → show/hide its PO rows (no SKUs in the DOM yet, so it can't over-expand).
function emToggle(gid) {
  var rows = document.querySelectorAll('tr[data-child="' + gid + '"]');
  if (!rows.length) return;
  var willOpen = rows[0].hidden;
  emSetHidden(gid, !willOpen);
  var arw = document.getElementById('arw-' + gid);
  if (arw) arw.classList.toggle('open', willOpen);
}
// PO toggle → LAZY-load its SKUs on first open, then plain show/hide thereafter.
function emTogglePO(row) {
  var pgid = row.getAttribute('data-pgid');
  var arw = document.getElementById('arw-' + pgid);
  var existing = document.querySelectorAll('tr[data-child="' + pgid + '"]');
  if (existing.length) {                              // already loaded → toggle
    var willOpen = existing[0].hidden;
    emSetHidden(pgid, !willOpen);
    if (arw) arw.classList.toggle('open', willOpen);
    return;
  }
  if (row.dataset.loading) return;                   // fetch in flight
  row.dataset.loading = '1';
  if (arw) { arw.classList.remove('open'); arw.classList.add('em-spin'); }
  var loader = document.createElement('tr');
  loader.setAttribute('data-loader', pgid); loader.className = 'em-sku-loader';
  loader.innerHTML = '<td colspan="7" class="em-loading"><span class="em-dot"></span>'
    + '<span class="em-dot"></span><span class="em-dot"></span> loading SKUs…</td>';
  row.parentNode.insertBefore(loader, row.nextSibling);
  fetch(EM_SKUS_URL + '?po=' + encodeURIComponent(row.getAttribute('data-po'))
        + '&day=' + encodeURIComponent(EM_DAY) + '&seg=' + encodeURIComponent(EM_SEG)
        + '&pgid=' + encodeURIComponent(pgid),
        { headers: { 'X-Requested-With': 'XMLHttpRequest' }, credentials: 'same-origin' })
    .then(function (r) { return r.text(); })
    .then(function (html) {
      loader.remove(); row.insertAdjacentHTML('afterend', html);
      row.dataset.loading = '';
      if (arw) { arw.classList.remove('em-spin'); arw.classList.add('open'); }
    })
    .catch(function () {
      loader.remove(); row.dataset.loading = '';
      if (arw) arw.classList.remove('em-spin');
      window.B2B && B2B.toast('Could not load SKUs — please retry.', 'error');
    });
}
// "No PO today" collapse — reveal/hide the not-received MPs for a segment.
function emNtd(key, btn) {
  var rows = document.querySelectorAll('tr.em-ntd-' + key);
  if (!rows.length) return;
  var open = rows[0].hidden;                          // currently hidden → reveal
  for (var i = 0; i < rows.length; i++) {
    rows[i].hidden = !open;
    if (open) { rows[i].classList.remove('em-anim'); rows[i].classList.add('em-anim'); }
  }
  var arw = btn.querySelector('.em-arrow');
  if (arw) arw.classList.toggle('open', open);
  btn.querySelector('.em-ntd-txt').textContent =
    (open ? 'Hide' : 'Show') + ' ' + rows.length + ' with no PO today';
}
// Single delegated click handler — closest() guarantees a PO click resolves to the
// PO row (emTogglePO), NEVER its MP; an MP click resolves to the MP (emToggle). This
// removes the inline-onclick + bubbling ambiguity that made a PO click collapse the MP.
// Bound ONCE: under the persistent shell-nav this script re-runs on re-visit and a
// removed <script> tag does NOT detach a listener already on document — a second
// toggle handler would open-then-close each row (net nothing).
if (!window.__emRowToggleBound) {
  window.__emRowToggleBound = true;
  document.addEventListener('click', function (e) {
    var t = e.target;
    if (!t || !t.closest) return;
    var po = t.closest('tr.em-porow');
    if (po) { emTogglePO(po); return; }
    var mp = t.closest('tr.em-mprow');
    if (mp) { emToggle(mp.getAttribute('data-gid')); return; }
  });
}
// Label each segment's not-today toggle with its live count (re-run after AJAX swap).
function emInitNtd() {
  document.querySelectorAll('.em-ntd-btn').forEach(function (btn) {
    var key = btn.getAttribute('data-seg');
    var n = document.querySelectorAll('tr.em-ntd-' + key).length;
    var t = btn.querySelector('.em-ntd-txt');
    if (!n) { btn.style.display = 'none'; }
    else { btn.style.display = ''; if (t) t.textContent = 'Show ' + n + ' with no PO today'; }
  });
}
// AJAX segment/day switch — swap only #em-dynamic (no full reload).
function emScope(seg, day) {
  var dyn = document.getElementById('em-dynamic');
  if (!dyn) return;
  dyn.classList.add('em-loading-dim');
  var url = CFG.email + '?seg=' + encodeURIComponent(seg) + '&day=' + encodeURIComponent(day);
  fetch(url, { headers: { 'X-Requested-With': 'XMLHttpRequest' }, credentials: 'same-origin' })
    .then(function (r) { return r.text(); })
    .then(function (txt) {
      var doc = new DOMParser().parseFromString(txt, 'text/html');
      var nd = doc.getElementById('em-dynamic');
      if (nd) dyn.innerHTML = nd.innerHTML;
      var ns = doc.getElementById('em-subject'), cs = document.getElementById('em-subject');
      if (ns && cs) cs.value = ns.value;                 // subject follows the day
      document.querySelectorAll('.em-tab').forEach(function (t) {
        t.classList.toggle('on', t.getAttribute('data-seg') === seg);
      });
      EM_DAY = day;                                       // lazy SKU fetch uses the new day
      var ex = document.getElementById('em-export');      // keep the Excel export on the shown board
      if (ex && CFG.export) ex.href = CFG.export + '?seg=' + encodeURIComponent(seg) + '&day=' + encodeURIComponent(day);
      try { history.replaceState(null, '', url); } catch (e) {}
      dyn.classList.remove('em-loading-dim');
      emInitNtd();
    })
    .catch(function () {
      dyn.classList.remove('em-loading-dim');
      window.B2B && B2B.toast('Could not refresh — please retry.', 'error');
    });
}
// Runs immediately (script is deferred / injected after DOM is ready), so it works
// on a normal load AND under app-shell partial navigation (where DOMContentLoaded
// has already fired and would never call back).
(function emInit() {
  emInitNtd();
  document.querySelectorAll('.em-tab').forEach(function (tab) {
    tab.addEventListener('click', function () {
      var di = document.getElementById('em-day');
      emScope(tab.getAttribute('data-seg'), di ? di.value : EM_DAY);
    });
  });
  var day = document.getElementById('em-day');
  if (day) day.addEventListener('change', function () {
    var on = document.querySelector('.em-tab.on');
    emScope(on ? on.getAttribute('data-seg') : 'online', day.value);
  });
})();

(function () {
  var $ = function (id) { return document.getElementById(id); };
  var modal = $('em-modal'), frame = $('em-frame');
  var reviewBtn = $('em-review'), sendBtn = $('em-send');
  if (!reviewBtn || !sendBtn) return;   // chooser landing → no compose/send UI, skip.
  var sending = false, previewed = false;

  function fields() {
    var day = document.querySelector('.em-dayform input[name=day]');
    var seg = document.querySelector('.em-dayform select[name=seg]');
    return {
      day: day ? day.value : '',
      seg: seg ? seg.value : '',
      subject: ($('em-subject').value || '').trim(),
      to: $('em-to').value || '',
      cc: $('em-cc').value || '',
      note: $('em-note').value || ''
    };
  }
  function qs(f) {
    return 'day=' + encodeURIComponent(f.day) +
      '&seg=' + encodeURIComponent(f.seg) +
      '&subject=' + encodeURIComponent(f.subject) +
      '&to=' + encodeURIComponent(f.to) +
      '&cc=' + encodeURIComponent(f.cc) +
      '&note=' + encodeURIComponent(f.note);
  }
  function setMsg(text, kind) {
    var m = $('em-sendmsg'); m.textContent = text || '';
    m.className = 'em-sendmsg' + (kind ? ' ' + kind : '');
  }

  // Review = fetch the EXACT email (preview, no send) → show it in the modal.
  reviewBtn.addEventListener('click', function () {
    reviewBtn.disabled = true; reviewBtn.textContent = 'Preparing…';
    fetch(CFG.preview + '?' + qs(fields()), {
      headers: { 'X-Requested-With': 'XMLHttpRequest' }, credentials: 'same-origin'
    }).then(function (r) { return r.json(); }).then(function (j) {
      reviewBtn.disabled = false; reviewBtn.textContent = 'Review & Send…';
      if (!j.ok) { window.B2B && B2B.toast('Could not build the preview.', 'error'); return; }
      $('em-msubject').textContent = j.subject || '(no subject)';
      $('em-mrecips').textContent = 'To: ' + ((j.to || []).join(', ') || '—') +
        ((j.cc && j.cc.length) ? '   ·   Cc: ' + j.cc.join(', ') : '');
      frame.srcdoc = j.html || '';
      previewed = true; setMsg(''); sendBtn.disabled = false;
      modal.hidden = false;
    }).catch(function () {
      reviewBtn.disabled = false; reviewBtn.textContent = 'Review & Send…';
      window.B2B && B2B.toast('Preview failed — please retry.', 'error');
    });
  });

  function close() { modal.hidden = true; }
  $('em-close').addEventListener('click', close);
  $('em-cancel').addEventListener('click', close);
  modal.addEventListener('click', function (e) { if (e.target === modal) close(); });

  // Send = deliberate, only after review; guarded against double-send.
  sendBtn.addEventListener('click', function () {
    if (sending || !previewed) return;
    sending = true; sendBtn.disabled = true; sendBtn.textContent = 'Sending…';
    setMsg('Contacting the mail server…');
    var f = fields();
    var body = 'subject=' + encodeURIComponent(f.subject) +
      '&to=' + encodeURIComponent(f.to) + '&cc=' + encodeURIComponent(f.cc) +
      '&note=' + encodeURIComponent(f.note);
    B2B.postForm(CFG.send + '?day=' + encodeURIComponent(f.day) +
          '&seg=' + encodeURIComponent(f.seg), body)
      .then(function (j) {
      sending = false;
      if (j.ok) {
        setMsg('✓ Sent.', 'ok');
        window.B2B && B2B.toast('Summary email sent.', 'success');
        sendBtn.textContent = 'Sent ✓';           // stays disabled → no double-send
        setTimeout(close, 1200);
      } else {
        sendBtn.disabled = false; sendBtn.textContent = 'Send now';
        setMsg(j.error || 'Send failed.', 'err');
        window.B2B && B2B.toast(j.error || 'Send failed.', 'error');
      }
    }).catch(function () {
      sending = false; sendBtn.disabled = false; sendBtn.textContent = 'Send now';
      setMsg('Network error — please retry.', 'err');
    });
  });
})();
