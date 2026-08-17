/* online_b2b/setup.html — page script (separated from template). */
// AJAX the Setup actions — NEVER a full reload, and never collapse the TiDB form.
// On a state change we update ONLY the live-status line + the two cards' active
// state from the server-returned status. Every failure is caught → error toast.
(function () {
  function toast(msg, ok) {
    // Prefer the shared skeleton toast; otherwise a self-contained slide-in
    // toast — NEVER a native alert().
    if (window.B2B && typeof B2B.toast === 'function') {
      B2B.toast(msg, { type: ok ? 'success' : 'error', title: 'Database', timeout: 7000 });
      return;
    }
    var host = document.getElementById('su-toasts');
    if (!host) {
      host = document.createElement('div');
      host.id = 'su-toasts';
      host.style.cssText = 'position:fixed;top:74px;right:20px;z-index:99999;'
        + 'display:flex;flex-direction:column;gap:10px;max-width:380px';
      document.body.appendChild(host);
    }
    var t = document.createElement('div');
    t.textContent = msg;
    t.style.cssText = 'background:' + (ok ? '#047857' : '#b91c1c') + ';color:#fff;'
      + 'padding:12px 15px;border-radius:11px;font-size:.85rem;line-height:1.4;'
      + 'box-shadow:0 14px 36px -12px rgba(0,0,0,.55);opacity:0;transform:translateY(-8px);'
      + 'transition:opacity .2s ease,transform .2s ease';
    host.appendChild(t);
    requestAnimationFrame(function () { t.style.opacity = '1'; t.style.transform = 'none'; });
    setTimeout(function () {
      t.style.opacity = '0'; t.style.transform = 'translateY(-8px)';
      setTimeout(function () { if (t.parentNode) t.remove(); }, 260);
    }, 6500);
  }
  function el(id) { return document.getElementById(id); }

  // Update the "currently connected" line + each card's active state in place.
  function applyStatus(st) {
    if (!st) return;
    var line = el('su-connline');
    if (line) line.innerHTML = '<b>' + (st.host || '—') + ':' + (st.port || '—') + '</b> / '
      + (st.database || '—') + (st.active ? ' · profile “' + st.active + '”' : '');
    var tls = el('su-tls');
    if (tls) { tls.textContent = st.tls ? 'TLS ON' : 'NO TLS'; tls.className = 'tls' + (st.tls ? '' : ' off'); }
    [['local', 'Local'], ['tidb', 'Server']].forEach(function (p) {
      var n = p[0], isact = (st.active === n);
      var card = el('card-' + n), badge = el('badge-' + n), btn = el('switchbtn-' + n);
      if (card) card.classList.toggle('on', isact);
      if (badge) { badge.textContent = isact ? '● Active' : p[1]; badge.classList.toggle('active', isact); }
      if (btn) { btn.disabled = isact; btn.textContent = isact ? 'In use' : ('Switch to ' + (n === 'local' ? 'Local' : 'TiDB')); }
    });
    // Last-backup line on the backup card.
    var lb = el('su-lastbackup');
    if (lb && st.last_backup) {
      var b = st.last_backup;
      lb.style.color = '#16a34a';
      lb.textContent = '✓ Last backup: ' + b.at + ' · ' + (b.rows || 0) + ' rows, '
        + (b.tables || 0) + ' tables, ' + (b.views || 0) + ' views (' + (b.elapsed || 0) + 's)';
    }
  }

  document.addEventListener('submit', function (ev) {
    var form = ev.target.closest && ev.target.closest('form');
    if (!form || !form.closest('.setup-wrap')) return;
    ev.preventDefault();
    var actEl = form.querySelector('[name=action]');
    var action = actEl ? actEl.value : 'switch';
    // Destructive backup → confirm before running (default already prevented above,
    // so returning here cancels cleanly with no request).
    if (action === 'backup_local' && !window.confirm(
        'Backup TiDB → local MySQL?\n\nThis OVERWRITES your local MySQL with ALL TiDB '
        + 'data (schema + rows + views). TiDB is not changed. Continue?')) return;
    var btn = form.querySelector('button[type=submit]');
    if (btn) {
      btn.disabled = true; btn._label = btn.textContent;
      btn.textContent = (action === 'backup_local') ? 'Backing up… (up to a minute)' : 'Working…';
    }
    function restore() { if (btn) { btn.disabled = false; btn.textContent = btn._label; } }

    // Always POST to the trailing-slash URL — Django APPEND_SLASH does NOT
    // redirect POSTs, so a missing slash would 404. Belt-and-suspenders.
    var url = form.getAttribute('action') || window.location.pathname;
    if (url.indexOf('?') === -1 && url.charAt(url.length - 1) !== '/') url += '/';
    fetch(url, { method: 'POST', body: new FormData(form),
        headers: { 'X-Requested-With': 'XMLHttpRequest' } })
      .then(function (r) {
        if (!r.ok) throw new Error('server returned ' + r.status);
        return r.json();
      })
      .then(function (d) {
        toast(d.message, d.ok);
        // Switch / save changed the active target → reflect it live (no reload,
        // no closing the TiDB form). Test never changes state.
        if (d.ok && action !== 'test' && d.status) applyStatus(d.status);
        restore();
      })
      .catch(function (e) { toast('Action failed: ' + e.message, false); restore(); });
  });
})();
