/* online_b2b/eka_data.html — page script (separated from template). */
// Save a row edit WITHOUT reloading — POST via fetch, toast + brief highlight.
  // FormData(form) picks up the row's inputs even though they're associated via
  // the HTML5 form= attribute (they're part of the form's entry list).
  document.querySelectorAll('form[id^="ekaf"]').forEach(function (f) {
    f.addEventListener('submit', function (e) {
      e.preventDefault();
      fetch(f.action, {
        method: 'POST',
        headers: { 'X-Requested-With': 'XMLHttpRequest' },
        body: new FormData(f),
        credentials: 'same-origin'
      }).then(function (r) {
        return r.json().then(function (j) { return { ok: r.ok, j: j }; });
      }).then(function (res) {
        var anchor = document.querySelector('[form="' + f.id + '"]');
        var tr = anchor && anchor.closest('tr');
        if (res.ok && res.j.ok) {
          if (window.B2B && B2B.toast) B2B.toast('Saved.', { type: 'ok', title: 'EKA store updated', timeout: 2200 });
          if (tr) { tr.style.transition = 'background .35s'; tr.style.background = 'color-mix(in srgb,#16a34a 16%,transparent)'; setTimeout(function () { tr.style.background = ''; }, 800); }
        } else {
          if (window.B2B && B2B.toast) B2B.toast((res.j && res.j.error) || 'Save failed.', { type: 'error', title: 'Not saved' });
          if (tr) { tr.style.transition = 'background .35s'; tr.style.background = 'color-mix(in srgb,#e5484d 16%,transparent)'; setTimeout(function () { tr.style.background = ''; }, 900); }
        }
      }).catch(function () {
        if (window.B2B && B2B.toast) B2B.toast('Network error — not saved.', { type: 'error', title: 'Not saved' });
      });
    });
  });
