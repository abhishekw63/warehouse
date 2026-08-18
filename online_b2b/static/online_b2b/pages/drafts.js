/* online_b2b/drafts.html — page script (separated from template). */
(function () {
  // Reopen & re-validate is a full-page nav to a slow server revalidation — show
  // an overlay immediately so the wait is narrated (nothing renders until the
  // server responds otherwise).
  var ov = document.getElementById('ro-overlay');
  var msg = document.getElementById('ro-msg');
  var STEPS = ['Reopening the parked run…', 'Re-validating against the current master…',
    'Checking orders & mapping…', 'Re-pricing the flagged lines…', 'Almost there…'];
  document.querySelectorAll('a.ropen[href*="revalidate"]').forEach(function (a) {
    a.addEventListener('click', function () {
      if (!ov) return;                       // don't block navigation if missing
      ov.classList.add('on');
      var i = 0;
      // Advance through the phases ONCE, then HOLD on the last one ("Almost
      // there…") until the server responds and the page navigates. Never loop
      // back to the start — restarting from "Reopening…" reads as if the whole
      // thing began again, which is misleading; the work only moves forward.
      var timer = setInterval(function () {
        if (i >= STEPS.length - 1) { clearInterval(timer); return; }  // hold on last step
        i++;
        if (!msg) return;
        msg.style.opacity = '0';
        setTimeout(function () { msg.textContent = STEPS[i]; msg.style.opacity = '1'; }, 160);
      }, 1300);
    });
  });

  document.querySelectorAll('.dl-email').forEach(function (btn) {
    btn.addEventListener('click', function () {
      if (btn.disabled) return;
      var orig = btn.innerHTML; btn.disabled = true; btn.innerHTML = '✉ Sending…';
      B2B.postForm(btn.dataset.url)
        .then(function (j) {
          btn.disabled = false;
          if (j.ok) {
            btn.innerHTML = '✓ Emailed';
            var msg = 'Issue email sent to ' + (j.to || []).join(', ') + ' (' + j.lines + ' line(s)).';
            if (window.B2B && B2B.toast) B2B.toast(msg, 'ok'); else alert(msg);
            setTimeout(function () { btn.innerHTML = orig; }, 4000);
          } else {
            btn.innerHTML = orig;
            if (window.B2B && B2B.toast) B2B.toast(j.error || 'Could not send.', 'err');
            else alert(j.error || 'Could not send the email.');
          }
        })
        .catch(function () { btn.disabled = false; btn.innerHTML = orig; alert('Network error — email not sent.'); });
    });
  });
})();
