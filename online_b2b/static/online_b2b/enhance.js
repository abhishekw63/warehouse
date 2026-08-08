/* =========================================================================
 * enhance.js — one global, dependency-free UI skeleton for the B2B app.
 *
 * Design: a SINGLE namespaced module (window.B2B) exposing reusable
 * primitives — theme, toast, palette, tables. Each is implemented ONCE here
 * and reused across every page; pages never re-implement toasts/sort/etc.
 * No framework, no build step — safe to load on every server-rendered page.
 *
 *   B2B.toast(msg, {type,title,timeout})   -> toast notification
 *   B2B.theme.toggle() / .set('dark'|'light')
 *   B2B.palette.open()                      -> ⌘K command palette
 *   B2B.enhanceTables(root)                 -> sticky + sortable + tabular-nums
 *
 * Progressive enhancement only: if this file fails to load, the app still
 * works exactly as before.
 * ======================================================================= */
(function (w, d) {
  "use strict";
  if (w.B2B) return;                         // idempotent — never double-init
  var B2B = (w.B2B = {});
  var $ = function (s, r) { return (r || d).querySelector(s); };
  var el = function (tag, cls, html) {
    var n = d.createElement(tag);
    if (cls) n.className = cls;
    if (html != null) n.innerHTML = html;
    return n;
  };

  /* --- shared inline icons (defined once, referenced by key) ------------- */
  var ICON = {
    sun:   '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/></svg>',
    moon:  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8z"/></svg>',
    search:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="11" cy="11" r="7"/><path d="M21 21l-4.3-4.3"/></svg>',
    ok:    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6L9 17l-5-5"/></svg>',
    warn:  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 9v4M12 17h.01M10.3 3.9L1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/></svg>',
    error: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M15 9l-6 6M9 9l6 6"/></svg>',
    info:  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 11v5M12 8h.01"/></svg>'
  };

  /* ===================================================================== *
   * 1. THEME  — light/dark, persisted, respects OS preference             *
   * ===================================================================== */
  var THEME_KEY = "b2b-theme";
  B2B.theme = {
    get: function () {
      return d.documentElement.getAttribute("data-theme") || "light";
    },
    set: function (mode, animate) {
      if (animate !== false) {
        d.documentElement.classList.add("enh-theming");
        w.setTimeout(function () { d.documentElement.classList.remove("enh-theming"); }, 320);
      }
      d.documentElement.setAttribute("data-theme", mode);
      try { localStorage.setItem(THEME_KEY, mode); } catch (e) {}
      var btn = $("#enh-theme-btn");
      if (btn) btn.innerHTML = mode === "dark" ? ICON.sun : ICON.moon;
    },
    toggle: function () { this.set(this.get() === "dark" ? "light" : "dark"); }
  };
  // Apply saved/OS theme immediately (before paint where possible)
  (function () {
    var saved;
    try { saved = localStorage.getItem(THEME_KEY); } catch (e) {}
    if (!saved && w.matchMedia && w.matchMedia("(prefers-color-scheme: dark)").matches) saved = "dark";
    if (saved) d.documentElement.setAttribute("data-theme", saved);
  })();

  /* ===================================================================== *
   * 2. TOAST  — single reusable notifier (replaces ad-hoc alerts)         *
   * ===================================================================== */
  var toastHost;
  B2B.toast = function (msg, opts) {
    opts = opts || {};
    var type = opts.type || "info";
    if (!toastHost) { toastHost = el("div", "enh-toasts"); d.body.appendChild(toastHost); }
    var t = el("div", "enh-toast " + type);
    var icon = el("span", "enh-ticon", ICON[type] || ICON.info);
    var tx = el("div", "enh-tx");
    if (opts.title) tx.appendChild(el("p", "enh-tt", opts.title));
    tx.appendChild(el("div", "enh-tc", msg));
    var close = el("button", "enh-tclose", "&times;");
    t.appendChild(icon); t.appendChild(tx); t.appendChild(close);
    toastHost.appendChild(t);
    var timer, dismiss = function () {
      if (t.classList.contains("enh-out")) return;
      t.classList.add("enh-out");
      w.setTimeout(function () { t.remove(); }, 260);
      w.clearTimeout(timer);
    };
    close.addEventListener("click", dismiss);
    if (opts.timeout !== 0) timer = w.setTimeout(dismiss, opts.timeout || 4200);
    return { dismiss: dismiss };
  };
  // Upgrade Django's server-rendered messages into toasts (one code path).
  function adoptDjangoMessages() {
    var map = { success: "ok", error: "error", warning: "warn", info: "info", debug: "info" };
    d.querySelectorAll("#toast-container .toast").forEach(function (node) {
      var cls = (node.className.match(/toast-(\w+)/) || [])[1];
      var span = node.querySelector("span");
      B2B.toast((span ? span.textContent : node.textContent).trim(),
                { type: map[cls] || "info", timeout: 5500 });
      node.remove();
    });
  }

  /* ---- ONE toast system, app-wide -------------------------------------- *
   * Route every legacy notifier into B2B.toast so no page can render a
   * different popup: (a) native alert() calls scattered through templates,
   * (b) the old core/js/script.js createToast(). confirm()/prompt() stay
   * native — they're blocking and callers depend on the return value.      */
  B2B._nativeAlert = w.alert && w.alert.bind ? w.alert.bind(w) : null;
  function inferType(msg) {
    var s = String(msg || "").toLowerCase();
    if (/error|failed|could ?n'?t|cannot|can'?t|network|invalid|required|⚠|denied|no .*found/.test(s)) return "error";
    if (/success|saved|sent|added|done|ready|complete/.test(s)) return "success";
    return "info";
  }
  // Any alert(...) anywhere becomes the single toast (no native browser box).
  w.alert = function (msg) { B2B.toast(String(msg == null ? "" : msg), { type: inferType(msg) }); };
  // Legacy programmatic toast helper → same single system (type names remapped).
  var _legacyMap = { success: "success", error: "error", warning: "warn", warn: "warn", info: "info", danger: "error" };
  w.createToast = function (message, type) {
    return B2B.toast(String(message == null ? "" : message), { type: _legacyMap[type] || inferType(message) });
  };

  /* ===================================================================== *
   * 3. COMMAND PALETTE (⌘K / Ctrl-K) — jump anywhere; index from sidebar  *
   * ===================================================================== */
  var pal = { root: null, input: null, list: null, items: [], sel: 0 };
  function buildIndex() {
    var out = [];
    // Navigation: harvested live from the sidebar so it never goes stale.
    d.querySelectorAll(".b2b-side a[href], nav.b2b-side a[href], .side-links a[href]").forEach(function (a) {
      var label = (a.getAttribute("title") || a.textContent || "").trim();
      if (label && a.getAttribute("href") && a.getAttribute("href") !== "#")
        out.push({ label: label, href: a.getAttribute("href"), sec: "Navigate", icon: "→" });
    });
    // Actions: primary buttons/links on the current page.
    d.querySelectorAll("a.btn-primary, .b2b-head a.btn, .actions a").forEach(function (a) {
      var label = (a.textContent || "").trim().replace(/\s+/g, " ");
      if (label && a.href) out.push({ label: label, href: a.href, sec: "Action on this page", icon: "•" });
    });
    // De-duplicate by label+href (standard once, not per-render churn).
    var seen = {};
    return out.filter(function (o) {
      var k = o.label + "|" + o.href;
      if (seen[k]) return false; seen[k] = 1; return true;
    });
  }
  function ensurePalette() {
    if (pal.root) return;
    pal.root = el("div", "enh-cmdk");
    pal.root.innerHTML =
      '<div class="enh-cmdk-panel" role="dialog" aria-label="Command palette">' +
        '<div class="enh-cmdk-in">' + ICON.search +
          '<input type="text" placeholder="Search pages & actions…" aria-label="Search" autocomplete="off">' +
          '<span class="enh-kbd">ESC</span></div>' +
        '<div class="enh-cmdk-list"></div>' +
      '</div>';
    d.body.appendChild(pal.root);
    pal.input = $(".enh-cmdk-in input", pal.root);
    pal.list = $(".enh-cmdk-list", pal.root);
    pal.root.addEventListener("mousedown", function (e) { if (e.target === pal.root) B2B.palette.close(); });
    pal.input.addEventListener("input", render);
    pal.input.addEventListener("keydown", onKey);
  }
  function render() {
    var q = pal.input.value.trim().toLowerCase();
    var all = pal.all || (pal.all = buildIndex());
    var hits = q ? all.filter(function (o) { return o.label.toLowerCase().indexOf(q) > -1; }) : all;
    pal.items = hits; pal.sel = 0;
    if (!hits.length) { pal.list.innerHTML = '<div class="enh-cmdk-empty">No matches</div>'; return; }
    var sec = "", html = "";
    hits.forEach(function (o, i) {
      if (o.sec !== sec) { sec = o.sec; html += '<div class="enh-cmdk-sec">' + sec + "</div>"; }
      html += '<a class="enh-cmdk-item' + (i === 0 ? " sel" : "") + '" href="' + o.href + '" data-i="' + i + '">' +
                '<span class="enh-ci">' + o.icon + "</span>" +
                '<span class="enh-cl">' + o.label + "</span></a>";
    });
    pal.list.innerHTML = html;
    pal.list.querySelectorAll(".enh-cmdk-item").forEach(function (node) {
      node.addEventListener("mousemove", function () { select(+node.dataset.i); });
    });
  }
  function select(i) {
    if (!pal.items.length) return;
    pal.sel = (i + pal.items.length) % pal.items.length;
    pal.list.querySelectorAll(".enh-cmdk-item").forEach(function (n) {
      n.classList.toggle("sel", +n.dataset.i === pal.sel);
    });
    var cur = pal.list.querySelector(".enh-cmdk-item.sel");
    if (cur) cur.scrollIntoView({ block: "nearest" });
  }
  function onKey(e) {
    if (e.key === "ArrowDown") { e.preventDefault(); select(pal.sel + 1); }
    else if (e.key === "ArrowUp") { e.preventDefault(); select(pal.sel - 1); }
    else if (e.key === "Enter") { e.preventDefault(); var o = pal.items[pal.sel]; if (o) w.location.href = o.href; }
    else if (e.key === "Escape") { B2B.palette.close(); }
  }
  B2B.palette = {
    open: function () {
      ensurePalette();
      pal.all = null;                        // rebuild index each open (page may have changed)
      pal.input.value = ""; render();
      pal.root.classList.add("on");
      pal.input.focus();
    },
    close: function () { if (pal.root) pal.root.classList.remove("on"); }
  };

  /* ===================================================================== *
   * 4. TABLES — sticky header, click-to-sort, tabular-nums (opt-in)       *
   * ===================================================================== */
  var NUM_RE = /^[₹$\s]*-?[\d,]+(\.\d+)?%?\s*$/;
  function numVal(s) { var n = parseFloat((s || "").replace(/[^0-9.\-]/g, "")); return isNaN(n) ? null : n; }
  function sortBy(table, col, dir) {
    var tb = table.tBodies[0]; if (!tb) return;
    var rows = [].slice.call(tb.rows);
    rows.sort(function (a, b) {
      var x = (a.cells[col] ? a.cells[col].textContent : "").trim();
      var y = (b.cells[col] ? b.cells[col].textContent : "").trim();
      // Empty / em-dash cells always sink to the bottom, both directions.
      var bx = (x === "" || x === "—"), by = (y === "" || y === "—");
      if (bx && !by) return 1;
      if (!bx && by) return -1;
      var nx = numVal(x), ny = numVal(y), r;
      if (nx !== null && ny !== null) r = nx - ny;
      else r = x.localeCompare(y, undefined, { numeric: true });
      return dir === "descending" ? -r : r;
    });
    // Smooth reorder — brief fade while the DOM re-sequences (no page reload,
    // no reflow jank). Honours reduced-motion. Reorder is instant either way.
    var reduce = window.matchMedia && matchMedia("(prefers-reduced-motion: reduce)").matches;
    if (reduce) { rows.forEach(function (row) { tb.appendChild(row); }); return; }
    tb.style.transition = "opacity .14s ease";
    tb.style.opacity = "0.25";
    window.setTimeout(function () {
      rows.forEach(function (row) { tb.appendChild(row); });
      tb.style.opacity = "1";
    }, 130);
  }
  B2B.enhanceTables = function (root) {
    (root || d).querySelectorAll("table").forEach(function (table) {
      // Only enhance genuine data tables: a header row + a real body.
      var head = table.tHead || (table.rows[0] && table.rows[0].parentNode.tagName === "THEAD" ? table.rows[0].parentNode : null);
      if (!head || !table.tBodies[0] || table.tBodies[0].rows.length < 2) return;
      if (table.classList.contains("enh-table") || table.hasAttribute("data-enh-skip")) return;
      table.classList.add("enh-table");
      // Right-align + tabular-nums for numeric columns (detected by body cells).
      var ths = head.rows[head.rows.length - 1].cells;
      [].forEach.call(ths, function (th, ci) {
        var body = table.tBodies[0].rows, numeric = 0, seen = 0;
        for (var r = 0; r < body.length && seen < 8; r++) {
          var c = body[r].cells[ci]; if (!c) continue; seen++;
          if (NUM_RE.test(c.textContent.trim())) numeric++;
        }
        var isNum = seen && numeric / seen >= 0.7;
        if (isNum) for (var r2 = 0; r2 < body.length; r2++) if (body[r2].cells[ci]) body[r2].cells[ci].classList.add("enh-num");
        if (th.hasAttribute("data-nosort")) { th.classList.add("enh-nosort"); return; }
        // Neutral up/down glyph shows the column IS sortable; it snaps to a solid
        // ▲ (asc) / ▼ (desc) once clicked.
        th.insertAdjacentHTML("beforeend", '<span class="enh-arrow">⇅</span>');
        th.addEventListener("click", function () {
          var cur = th.getAttribute("aria-sort");
          var dir = cur === "ascending" ? "descending" : "ascending";
          [].forEach.call(ths, function (o) { o.removeAttribute("aria-sort"); var a = o.querySelector(".enh-arrow"); if (a) a.textContent = "⇅"; });
          th.setAttribute("aria-sort", dir);
          th.querySelector(".enh-arrow").textContent = dir === "ascending" ? "▲" : "▼";
          sortBy(table, ci, dir);
        });
      });
    });
  };

  /* ===================================================================== *
   * 5. HEADER CONTROLS + GLOBAL WIRING                                    *
   * ===================================================================== */
  function mountControls() {
    var host = $(".header__right");
    if (!host || $("#enh-theme-btn")) return;
    var cmd = el("button", "enh-ctl");
    cmd.type = "button"; cmd.id = "enh-cmd-btn"; cmd.title = "Search (Ctrl-K)";
    cmd.innerHTML = ICON.search + '<span class="enh-kbd">Ctrl K</span>';
    cmd.addEventListener("click", function () { B2B.palette.open(); });
    var theme = el("button", "enh-ctl");
    theme.type = "button"; theme.id = "enh-theme-btn"; theme.title = "Toggle theme";
    theme.innerHTML = B2B.theme.get() === "dark" ? ICON.sun : ICON.moon;
    theme.addEventListener("click", function () { B2B.theme.toggle(); });
    host.insertBefore(theme, host.firstChild);
    host.insertBefore(cmd, host.firstChild);
  }

  function init() {
    mountControls();
    adoptDjangoMessages();
    B2B.enhanceTables(d);
  }

  // Global shortcut: ⌘K / Ctrl-K opens the palette anywhere.
  d.addEventListener("keydown", function (e) {
    if ((e.metaKey || e.ctrlKey) && (e.key === "k" || e.key === "K")) {
      e.preventDefault(); B2B.palette.open();
    }
  });

  if (d.readyState === "loading") d.addEventListener("DOMContentLoaded", init);
  else init();
})(window, document);
