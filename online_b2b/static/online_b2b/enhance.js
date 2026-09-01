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
  // Idempotent against double-loading THIS script, but must NOT bail just because
  // a partial window.B2B already exists — base_b2b.html's body_end inline script
  // (load overlay + bgDownload) runs during parse and creates window.B2B BEFORE
  // this deferred script executes. So EXTEND that object (keeping .load/.bgDownload)
  // rather than replacing it, and guard on our own init marker.
  if (w.B2B && w.B2B._enhanced) return;
  var B2B = (w.B2B = w.B2B || {});
  B2B._enhanced = true;
  // Mark the document as enhanced so CSS can hide the raw Django `.toast`
  // (server-rendered in #toast-container) — enhance is the SINGLE visible toast
  // system and adoptDjangoMessages() re-renders those messages as enh-toasts.
  // The raw ones stay in the DOM only as a no-JS fallback (shown when this class
  // is absent), so a message is never lost, but two toasts never show at once.
  try { d.documentElement.classList.add("enh-on"); } catch (e) {}
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
    // Day/night removed — the app is light-only. set()/toggle() are kept as
    // light-locked no-ops so any lingering caller can never force dark.
    set: function () { d.documentElement.setAttribute("data-theme", "light"); },
    toggle: function () { this.set("light"); }
  };
  // Force light and drop any previously-stored 'dark' choice so it can't linger.
  (function () {
    try { localStorage.removeItem(THEME_KEY); } catch (e) {}
    d.documentElement.setAttribute("data-theme", "light");
  })();

  /* ===================================================================== *
   * 2. TOAST  — single reusable notifier (replaces ad-hoc alerts)         *
   * ===================================================================== */
  var toastHost;
  B2B.toast = function (msg, opts) {
    // Tolerant signature: a STRING 2nd arg is taken as the type (many call sites do
    // B2B.toast(msg, 'ok'/'error')). Alias legacy/synonym names to the 4 real toast
    // types (ok · error · warn · info) so styling is always correct — 'success'→'ok',
    // 'err'/'danger'→'error', etc. Existing {type:'ok'|'error'|…} still pass through.
    if (typeof opts === "string") opts = { type: opts };
    opts = opts || {};
    var _TYPE = { success: "ok", ok: "ok", err: "error", error: "error", danger: "error",
                  warning: "warn", warn: "warn", info: "info" };
    var type = _TYPE[opts.type] || opts.type || "info";
    if (!toastHost) { toastHost = el("div", "enh-toasts"); d.body.appendChild(toastHost); }
    var t = el("div", "enh-toast " + type);
    var badge = el("span", "enh-tbadge");            // circular icon badge
    badge.appendChild(el("span", "enh-ticon", ICON[type] || ICON.info));
    var tx = el("div", "enh-tx");
    if (opts.title) tx.appendChild(el("p", "enh-tt", opts.title));
    tx.appendChild(el("div", "enh-tc", msg));
    var close = el("button", "enh-tclose", "&times;");
    t.appendChild(badge); t.appendChild(tx); t.appendChild(close);
    // Countdown progress bar — depletes over the toast's lifetime (skipped for
    // sticky toasts, timeout === 0). Duration is set inline so it always matches.
    var life = (opts.timeout === 0) ? 0 : (opts.timeout || 4000);
    if (life) {
      var bar = el("div", "enh-tbar");
      bar.style.animationDuration = life + "ms";
      t.appendChild(bar);
    }
    toastHost.appendChild(t);
    var timer, dismiss = function () {
      if (t.classList.contains("enh-out")) return;
      t.classList.add("enh-out");
      w.setTimeout(function () { t.remove(); }, 460);   // matches the .45s enh-out fade
      w.clearTimeout(timer);
    };
    close.addEventListener("click", dismiss);
    // Auto-dismiss after ~4s (smooth fade). Pass opts.timeout to override,
    // or opts.timeout === 0 to keep a toast until dismissed.
    if (opts.timeout !== 0) timer = w.setTimeout(dismiss, opts.timeout || 4000);
    return { dismiss: dismiss };
  };
  // Upgrade Django's server-rendered messages into toasts (one code path).
  function adoptDjangoMessages() {
    var map = { success: "ok", error: "error", warning: "warn", info: "info", debug: "info" };
    d.querySelectorAll("#toast-container .toast").forEach(function (node) {
      var cls = (node.className.match(/toast-(\w+)/) || [])[1];
      var span = node.querySelector("span");
      B2B.toast((span ? span.textContent : node.textContent).trim(),
                { type: map[cls] || "info", timeout: 4000 });
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
    if (/success|saved|sent|added|done|ready|complete/.test(s)) return "ok";
    return "info";
  }
  // Any alert(...) anywhere becomes the single toast (no native browser box).
  w.alert = function (msg) { B2B.toast(String(msg == null ? "" : msg), { type: inferType(msg) }); };
  // Legacy programmatic toast helper → same single system. Canonical toast types
  // are ok/warn/error/info (matches ICON + CSS), so map success→ok, warning→warn.
  var _legacyMap = { success: "ok", ok: "ok", error: "error", danger: "error",
                     warning: "warn", warn: "warn", info: "info" };
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

  /* ---- Reusable AJAX helpers (CSRF + fetch) ------------------------------- *
   * One CSRF getter (cookie → hidden input → meta) and two POST helpers so pages
   * stop hand-rolling the fetch envelope. B2B.postForm sends form-encoded (accepts
   * a plain object, a query string, or FormData); B2B.postJSON sends JSON. Both add
   * the CSRF + XHR headers and resolve to the parsed JSON.                        */
  B2B.csrf = function () {
    var m = d.cookie.match(/csrftoken=([^;]+)/);
    if (m) return m[1];
    var inp = d.querySelector("input[name=csrfmiddlewaretoken]");
    if (inp) return inp.value;
    var meta = d.querySelector("meta[name=csrf-token]");
    return meta ? meta.getAttribute("content") : "";
  };
  B2B.postForm = function (url, body) {
    var headers = { "X-CSRFToken": B2B.csrf(), "X-Requested-With": "XMLHttpRequest" };
    var b = body;
    if (body instanceof FormData) {
      // leave b as the FormData — the browser sets the multipart Content-Type.
    } else if (body instanceof URLSearchParams) {
      b = body.toString();
      headers["Content-Type"] = "application/x-www-form-urlencoded";
    } else if (body && typeof body === "object") {
      var p = new URLSearchParams();
      for (var k in body) if (Object.prototype.hasOwnProperty.call(body, k)) p.set(k, body[k]);
      b = p.toString();
      headers["Content-Type"] = "application/x-www-form-urlencoded";
    } else if (typeof body === "string") {
      headers["Content-Type"] = "application/x-www-form-urlencoded";
    }
    return fetch(url, { method: "POST", credentials: "same-origin", headers: headers, body: b })
      .then(function (r) { return r.json(); });
  };
  B2B.postJSON = function (url, obj) {
    return fetch(url, {
      method: "POST", credentials: "same-origin",
      headers: { "X-CSRFToken": B2B.csrf(), "X-Requested-With": "XMLHttpRequest",
                 "Content-Type": "application/json" },
      body: JSON.stringify(obj || {})
    }).then(function (r) { return r.json(); });
  };
  // Trailing debounce — returns a wrapped fn that only runs ms after the last call
  // (search inputs, resize handlers, etc.). Preserves this/args.
  B2B.debounce = function (fn, ms) {
    var t;
    return function () {
      var self = this, args = arguments;
      clearTimeout(t);
      t = setTimeout(function () { fn.apply(self, args); }, ms || 250);
    };
  };

  /* ---- Reusable "select-all + live count" for tick tables ----------------- *
   * Review-page style row selection: a master checkbox drives N item checkboxes
   * (with the indeterminate state), and onChange(count,total) fires on every
   * change (initial included). Shared so record-verify / review / ship-to don't
   * each re-implement it. Returns { items, checked(), sync() }.
   *   B2B.checkAll({ items:'.rv-chk', master:'#rvAll', onChange(n,total){…} })  */
  B2B.checkAll = function (opts) {
    opts = opts || {};
    var items = typeof opts.items === "string"
      ? [].slice.call((opts.root || d).querySelectorAll(opts.items))
      : [].slice.call(opts.items || []);
    var master = typeof opts.master === "string" ? $(opts.master) : opts.master;
    function checked() { return items.filter(function (c) { return c.checked; }); }
    function sync() {
      var n = checked().length;
      if (master) { master.checked = n > 0 && n === items.length; master.indeterminate = n > 0 && n < items.length; }
      if (opts.onChange) opts.onChange(n, items.length);
    }
    items.forEach(function (c) { c.addEventListener("change", sync); });
    if (master) master.addEventListener("change", function () {
      items.forEach(function (c) { c.checked = master.checked; }); sync();
    });
    sync();
    return { items: items, checked: checked, sync: sync };
  };

  /* ===================================================================== *
   * 5. HEADER CONTROLS + GLOBAL WIRING                                    *
   * ===================================================================== */
  function mountControls() {
    var host = $(".header__right");
    if (!host || $("#enh-cmd-btn")) return;
    var cmd = el("button", "enh-ctl");
    cmd.type = "button"; cmd.id = "enh-cmd-btn"; cmd.title = "Search (Ctrl-K)";
    cmd.innerHTML = ICON.search + '<span class="enh-kbd">Ctrl K</span>';
    cmd.addEventListener("click", function () { B2B.palette.open(); });
    // Day/night toggle removed — app is light-only.
    host.insertBefore(cmd, host.firstChild);
  }

  /* ===================================================================== *
   * 4b. COUNT-UP — dashboard KPI numbers tick 0 → value on load           *
   * ===================================================================== */
  // Animates the leading number in each target while preserving its prefix/suffix
  // and formatting (₹, Cr, M, %, commas). Reduced-motion → shows the final value
  // instantly. Idempotent per element. Exposed so an AJAX refresh can re-run it.
  // KPI/stat number elements across the app that should tick up on reveal.
  var COUNTUP_SEL = [
    ".hub-kpis .hk .n", "[data-countup]",           // Hub
    ".kpis .card .n",                                  // Review KPI cards
    ".iss-kpis .ik .n",                                // Issues
    ".fr-n", ".tr-n", ".zone-val", ".xstat b", ".dos-pill b",  // Analytics family
    ".av-kpi .n", ".iv-kpi .k-val"                    // Availability / Inventory cockpit
  ].join(", ");
  B2B.countUp = function (root) {
    var reduce = w.matchMedia && w.matchMedia("(prefers-reduced-motion: reduce)").matches;
    var els = (root || d).querySelectorAll(COUNTUP_SEL);
    [].forEach.call(els, function (node) {
      if (node._counted) return;
      node._counted = true;
      var raw = node.textContent.trim();
      var m = raw.match(/-?[\d,]*\.?\d+/);          // first number in the string
      if (!m) return;
      var target = parseFloat(m[0].replace(/,/g, ""));
      if (isNaN(target) || reduce) return;          // leave text as-is
      var prefix = raw.slice(0, m.index);
      var suffix = raw.slice(m.index + m[0].length);
      var decimals = (m[0].split(".")[1] || "").length;
      var hasComma = m[0].indexOf(",") !== -1;
      function fmt(v) {
        var s = decimals ? v.toFixed(decimals) : String(Math.round(v));
        if (hasComma) { try { s = Number(s).toLocaleString("en-IN"); } catch (e) {} }
        return prefix + s + suffix;
      }
      var dur = 900, t0 = null;
      function step(ts) {
        if (t0 === null) t0 = ts;
        var p = Math.min(1, (ts - t0) / dur);
        var eased = 1 - Math.pow(1 - p, 3);         // easeOutCubic
        node.textContent = fmt(target * eased);
        if (p < 1) w.requestAnimationFrame(step);
        else node.textContent = raw;                // restore exact original
      }
      node.textContent = fmt(0);
      w.requestAnimationFrame(step);
    });
  };

  /* ===================================================================== *
   * 4c. BAR FILLS — every progress/fill bar springs 0 → target when it     *
   * scrolls into view (Motion One). Reads the inline width % the server    *
   * rendered. Reduced-motion / no-Motion → left at full width (unchanged). *
   * ===================================================================== */
  // Curated fill-bar selectors across analytics, tasks, TAT, availability.
  var BAR_SEL = [
    ".fr-fill", ".zone-bar > span", ".ex-tbar > span", ".tbar > span",
    ".tat-prog__track > span", ".av-bar > span", ".dt-parent-bar > span",
    "#dt-ovfill"
  ].join(", ");
  // Native IntersectionObserver + a CSS width-transition — replaces the 63 KB
  // Motion One lib that used to load on EVERY page for just this one effect.
  // Same graceful fallbacks: reduced-motion OR no IntersectionObserver → bars are
  // left at full width (we never zero them), so nothing ever hides.
  B2B.revealBars = function (root) {
    var reduce = w.matchMedia && w.matchMedia("(prefers-reduced-motion: reduce)").matches;
    if (reduce || !("IntersectionObserver" in w)) return;  // graceful: bars stay filled
    var io = new IntersectionObserver(function (entries, obs) {
      entries.forEach(function (en) {
        if (!en.isIntersecting) return;
        var el = en.target;
        obs.unobserve(el);
        w.requestAnimationFrame(function () { el.style.width = el._barTarget; });
      });
    }, { threshold: 0.15 });
    (root || d).querySelectorAll(BAR_SEL).forEach(function (el) {
      if (el._barred) return;
      var target = (el.style && el.style.width) || "";
      if (target.indexOf("%") < 0) return;                 // only inline-% bars
      el._barred = true;
      el._barTarget = target;
      el.style.transition = "width .9s cubic-bezier(.2,.8,.2,1)";
      el.style.width = "0%";
      io.observe(el);
    });
  };

  /* ===================================================================== *
   * 4d. CELEBRATE — a confetti burst for big wins (Lock & Record, Backup). *
   * Reduced-motion / no-confetti → silent no-op. Call B2B.celebrate() or   *
   * add [data-celebrate] to an element that appears on success.            *
   * ===================================================================== */
  B2B.celebrate = function (opts) {
    var reduce = w.matchMedia && w.matchMedia("(prefers-reduced-motion: reduce)").matches;
    if (reduce) return;
    var burst = function () {
      if (typeof w.confetti !== "function") return;   // load failed → silent no-op
      var o = opts || {};
      var end = Date.now() + (o.ms || 900);
      var colors = ["#4f46e5", "#10b981", "#f59e0b", "#ec4899", "#06b6d4"];
      (function frame() {
        w.confetti({ particleCount: 4, angle: 60, spread: 60, origin: { x: 0 }, colors: colors });
        w.confetti({ particleCount: 4, angle: 120, spread: 60, origin: { x: 1 }, colors: colors });
        if (Date.now() < end) w.requestAnimationFrame(frame);
      })();
    };
    if (typeof w.confetti === "function") { burst(); return; }
    // LAZY-LOAD the vendored confetti lib on first celebration, then fire (and
    // queue any bursts requested while it's still loading). URL comes from the
    // enhance.js <script data-confetti-src> tag — no global lib load on every page.
    if (B2B._confettiQ) { B2B._confettiQ.push(burst); return; }
    B2B._confettiQ = [burst];
    var tag = d.querySelector("script[data-confetti-src]");
    var src = tag && tag.getAttribute("data-confetti-src");
    if (!src) { B2B._confettiQ = null; return; }
    var s = d.createElement("script");
    s.src = src; s.defer = true;
    s.onload = function () {
      (B2B._confettiQ || []).forEach(function (f) { try { f(); } catch (e) {} });
      B2B._confettiQ = null;
    };
    s.onerror = function () { B2B._confettiQ = null; };   // offline / blocked → no-op
    d.head.appendChild(s);
  };

  /* ── View-only (RBAC) ────────────────────────────────────────────────── *
   * When the user is a Viewer (body[data-can-write="0"]), disable every write   *
   * control so writes read as "disabled with tooltip" instead of click→error.   *
   * The SAFE regex MIRRORS the server allowlist in core/access.py (export /      *
   * download / search / pagination / preview / data / recon-run / availability  *
   * / d365 / auth) so client + server never disagree. The server middleware is  *
   * the real boundary; this is the UX layer. Idempotent + re-runs on new DOM.   */
  var _VO_SAFE = /\/(export|download|search|more|preview|data|run|check|bins|d365|login|logout|signup|password-change|profile|po-skus)\/?$/i;
  B2B.applyViewOnly = function (root) {
    root = root || d;
    if (d.body.getAttribute("data-can-write") !== "0") return;
    var mark = function (b) {
      if (b.getAttribute("data-vo")) return;
      b.setAttribute("data-vo", "1");
      // Look disabled (greyed + not-allowed) but stay CLICKABLE, so the capture
      // click-guard can fire the one view-only toast on every write button. A
      // real `disabled` swallows the click and gives no feedback.
      b.setAttribute("aria-disabled", "true");
      b.title = "View-only access — changes are disabled";
      b.style.cursor = "not-allowed";
      b.style.opacity = "0.55";
    };
    root.querySelectorAll("form").forEach(function (f) {
      if ((f.getAttribute("method") || "get").toLowerCase() !== "post") return;
      if (f.getAttribute("data-read") !== null) return;
      var action = f.getAttribute("action") || w.location.pathname;
      if (_VO_SAFE.test(action)) return;
      if (!f.getAttribute("data-vo")) {
        f.setAttribute("data-vo", "1");
        f.addEventListener("submit", function (ev) {   // hard stop, even if a control slipped through
          ev.preventDefault(); ev.stopPropagation();
          if (B2B.toast) B2B.toast("View-only access — you can't make changes. Ask an admin for Editor access.",
            { type: "error", title: "Read-only" });
        }, true);
      }
      // submit controls INSIDE the form + any associated by the HTML form= attr
      // (Discard / Save-for-Review-Later sit OUTSIDE their <form>).
      var btns = [].slice.call(f.querySelectorAll("button, input[type=submit], input[type=image]"));
      if (f.id) btns = btns.concat([].slice.call(d.querySelectorAll('[form="' + f.id + '"]')));
      btns.forEach(function (b) {
        var t = (b.getAttribute("type") || "submit").toLowerCase();
        if (t === "button" || t === "reset") return;   // non-submit buttons untouched
        mark(b);
      });
    });
    root.querySelectorAll("[data-write]").forEach(mark);   // explicitly-tagged AJAX write buttons
    root.querySelectorAll("[data-url]").forEach(function (b) {   // AJAX buttons/links to a write endpoint
      var u = b.getAttribute("data-url");
      if (u && !_VO_SAFE.test(u)) mark(b);
    });
  };

  /* Capture-phase click guard: the belt to applyViewOnly's suspenders. Runs
   * BEFORE any page handler (undecided-lock guard, email AJAX, form submit), so
   * clicking ANY write control gives the ONE view-only toast — never the page's
   * own intermediate UI. Catches: submit buttons (incl. external form= /
   * formaction), data-url AJAX buttons, and [data-write]. Editors: no-op. */
  function _voClickGuard(e) {
    if (d.body.getAttribute("data-can-write") !== "0") return;
    var el = e.target.closest && e.target.closest(
      "button, input[type=submit], input[type=image], a[data-url], [data-write]");
    if (!el) return;
    var block = false;
    if (el.hasAttribute("data-write")) block = true;
    if (!block) {
      var u = el.getAttribute("data-url");
      if (u && !_VO_SAFE.test(u)) block = true;                 // AJAX write button/link
    }
    if (!block && el.matches("button, input[type=submit], input[type=image]")) {
      var t = (el.getAttribute("type") || "submit").toLowerCase();
      if (t !== "button" && t !== "reset") {
        var f = el.form || (el.closest && el.closest("form"));
        var act = el.getAttribute("formaction") || (f && f.getAttribute("action")) || w.location.pathname;
        var method = (el.getAttribute("formmethod") || (f && f.getAttribute("method")) || "get").toLowerCase();
        if (f && method === "post" && f.getAttribute("data-read") === null && !_VO_SAFE.test(act)) block = true;
      }
    }
    if (block) {
      e.preventDefault(); e.stopImmediatePropagation();
      if (B2B.toast) B2B.toast("View-only access — you can't make changes. Ask an admin for Editor access.",
        { type: "error", title: "Read-only" });
    }
  }

  function init() {
    mountControls();
    adoptDjangoMessages();
    B2B.enhanceTables(d);
    B2B.countUp(d);
    B2B.revealBars(d);
    B2B.applyViewOnly(d);
    d.addEventListener("click", _voClickGuard, true);   // capture: block writes for Viewers

    // AJAX-loaded content (htmx swaps) gets the same treatment — tables, count-up
    // and bar-fills re-run on the swapped-in fragment so nothing lands "flat".
    d.body.addEventListener("htmx:afterSwap", function (e) {
      var t = (e && e.target) || d;
      B2B.enhanceTables(t); B2B.countUp(t); B2B.revealBars(t);
    });

    // Step 4: fade the newly-shown tab panel on switch. Delegated + runs AFTER
    // each page's own tab handler (setTimeout 0), so it works everywhere without
    // touching per-page JS. Re-triggers the animation by reflowing the class.
    d.addEventListener("click", function (e) {
      var tab = e.target.closest && e.target.closest(".tabs .tab[data-tab]");
      if (!tab) return;
      var name = tab.getAttribute("data-tab");
      w.setTimeout(function () {
        var pane = d.querySelector('.tabpane[data-pane="' + name + '"]');
        if (pane) { pane.classList.remove("enh-panefade"); void pane.offsetWidth; pane.classList.add("enh-panefade"); }
      }, 0);
    });

    // Universal re-animate: watch the whole document for NEW content (added by
    // fetch, htmx, or hand-rolled JS render — Analytics tabs, Availability
    // results, the Hub range switch, etc.) and re-run the enhancers on it. Cheap:
    // debounced, and each enhancer is idempotent (per-element guards skip
    // already-processed nodes), so nothing double-animates.
    if (w.MutationObserver) {
      var deb;
      var obs = new MutationObserver(function (muts) {
        var added = false;
        for (var i = 0; i < muts.length; i++) {
          if (muts[i].addedNodes && muts[i].addedNodes.length) { added = true; break; }
        }
        if (!added) return;
        w.clearTimeout(deb);
        deb = w.setTimeout(function () {
          B2B.enhanceTables(d); B2B.countUp(d); B2B.revealBars(d); B2B.applyViewOnly(d);
          // Celebrate a big win the moment its success node appears — the Lock &
          // Record "done" card, or anything tagged [data-celebrate]. Once each.
          var win = d.querySelector(".ld-card:not([data-celebrated]), [data-celebrate]:not([data-celebrated])");
          if (win) { win.setAttribute("data-celebrated", "1"); B2B.celebrate(); }
        }, 120);
      });
      obs.observe(d.body, { childList: true, subtree: true });
    }
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

/* ── Frontend perf beacon ─────────────────────────────────────────────────
   Reports the browser's Navigation Timing to the Audit Log on SLOW loads, so the
   log separates FRONTEND (network + browser render) from the server code/db split
   it already records. Fire-and-forget via sendBeacon; only fires above a threshold
   to keep the volume sane. Self-contained — delete this block to disable. */
(function () {
  if (!window.performance || !performance.getEntriesByType || !navigator.sendBeacon) return;
  var THRESHOLD = 700;   // ms — only report loads slower than this
  function report() {
    try {
      var nav = performance.getEntriesByType("navigation")[0];
      if (!nav || !nav.loadEventEnd) return;
      var total = Math.round(nav.duration || (nav.loadEventEnd - nav.startTime));
      if (!total || total < THRESHOLD) return;
      var span = function (a, b) { var v = Math.round(a - b); return v > 0 ? v : 0; };
      var payload = JSON.stringify({
        path: location.pathname,
        total: total,
        ttfb: span(nav.responseStart, nav.requestStart),   // network + server round-trip
        dl: span(nav.responseEnd, nav.responseStart),       // download
        dom: span(nav.domInteractive, nav.responseEnd),     // parse
        render: span(nav.loadEventEnd, nav.domInteractive)  // scripts + paint
      });
      navigator.sendBeacon("/perf/nav/", new Blob([payload], { type: "application/json" }));
    } catch (e) { /* telemetry must never break the page */ }
  }
  // Delay after load so loadEventEnd is finalised (it's 0 during the load event).
  if (document.readyState === "complete") setTimeout(report, 300);
  else window.addEventListener("load", function () { setTimeout(report, 300); });
})();

/* ── Global submit-once guard ─────────────────────────────────────────────── *
 * A double-click / double-tap on a POST form fires the POST twice; for the many
 * token-consuming confirms (Lock & Record, Item-Master rebuild, record-verify,
 * MT/GT confirm, ship-to upload …) the 2nd hit lands on an already-consumed
 * token → a 404 ("Upload not found or expired"). This blocks the 2nd submit and
 * shows a busy label on the submit control — app-wide, in ONE place.
 *
 * Safe by construction:
 *  • Bubble phase + defaultPrevented check → AJAX forms that preventDefault()
 *    (review lock, exceptions, email…) are LEFT to their own page JS, untouched.
 *  • Only POST forms (GET filters/search are exempt); opt out with data-no-guard.
 *  • Buttons are disabled on the NEXT tick, so the click's own formaction /
 *    name-value are still captured into the request first.
 *  • Native submits navigate → the page unloads (fresh buttons); no re-enable
 *    needed, nothing sticks. */
(function () {
  if (window.__b2bSubmitGuard) return;                    // attach ONCE (script re-runs on shell-nav)
  window.__b2bSubmitGuard = true;
  document.addEventListener("submit", function (e) {
    if (e.defaultPrevented) return;                       // AJAX/managed form → skip
    var f = e.target;
    if (!f || f.tagName !== "FORM" || f.hasAttribute("data-no-guard")) return;
    if ((f.getAttribute("method") || "get").toLowerCase() !== "post") return;
    if (f.getAttribute("data-guarded") === "1") { e.preventDefault(); return; }
    f.setAttribute("data-guarded", "1");
    var btns = f.querySelectorAll(
      "button[type=submit], input[type=submit], button:not([type])");
    setTimeout(function () {                               // request already captured
      btns.forEach(function (b) {
        if (b.disabled) return;
        b.disabled = true; b.setAttribute("aria-busy", "true");
        var t = (b.textContent || "").trim();
        if (t) {
          b.setAttribute("data-lbl", b.innerHTML);
          b.textContent = b.getAttribute("data-busy")
            || (t.replace(/[.…]+$/, "") + "…");
        }
      });
    }, 0);
  }, false);
})();
