/* ============================================================================
 * B2B date-range picker — dual-month, smooth, self-contained. No dependencies.
 *
 * Usage:
 *   <div data-drp
 *        data-from-name="sku_from" data-to-name="sku_to"
 *        data-from="2026-06-27" data-to="2026-07-27"
 *        data-max-days="0"            (optional cap; 0 = unlimited)
 *        data-submit="1"              (optional: submit the form on Apply)
 *        data-presets="1"></div>      (optional: show quick presets)
 *   The component injects two hidden inputs (from/to) so it posts in any form.
 * ==========================================================================*/
(function () {
  "use strict";
  var MON = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  var DOW = ["Su","Mo","Tu","We","Th","Fr","Sa"];

  function pad(n){ return (n < 10 ? "0" : "") + n; }
  function iso(d){ return d ? d.getFullYear() + "-" + pad(d.getMonth()+1) + "-" + pad(d.getDate()) : ""; }
  function parse(s){ if(!s) return null; var p = String(s).slice(0,10).split("-"); if(p.length!==3) return null;
    var d = new Date(+p[0], +p[1]-1, +p[2]); return isNaN(d) ? null : d; }
  function nice(d){ return d ? d.getDate() + " " + MON[d.getMonth()].slice(0,3) + " " + d.getFullYear() : "…"; }
  function same(a,b){ return a && b && a.getFullYear()===b.getFullYear() && a.getMonth()===b.getMonth() && a.getDate()===b.getDate(); }
  function dayIdx(d){ return Math.floor(d.getTime()/86400000); }
  function addMonths(d,n){ return new Date(d.getFullYear(), d.getMonth()+n, 1); }
  function el(tag, cls){ var e = document.createElement(tag); if(cls) e.className = cls; return e; }

  function DRP(root){
    var fromName = root.getAttribute("data-from-name") || "from";
    var toName   = root.getAttribute("data-to-name")   || "to";
    var maxDays  = parseInt(root.getAttribute("data-max-days") || "0", 10) || 0;
    var doSubmit = root.getAttribute("data-submit") === "1";
    var wantPresets = root.getAttribute("data-presets") === "1";

    var start = parse(root.getAttribute("data-from"));
    var end   = parse(root.getAttribute("data-to"));
    var today = new Date(); today.setHours(0,0,0,0);
    var view  = addMonths(start || today, 0); view.setDate(1);
    var hoverDay = null, open = false;
    var cells = [];   // {el, d} for every rendered day — repainted without DOM rebuild

    // hidden inputs (so the picker posts inside any form)
    var hf = el("input"); hf.type="hidden"; hf.name=fromName; hf.value = start ? iso(start) : "";
    var ht = el("input"); ht.type="hidden"; ht.name=toName;   ht.value = end   ? iso(end)   : "";
    root.appendChild(hf); root.appendChild(ht);

    // trigger
    var trigger = el("button","drp-trigger"); trigger.type="button";
    trigger.innerHTML = '<span class="drp-cal-ic">🗓</span><span class="drp-label"></span><span class="drp-chev">▾</span>';
    var label = trigger.querySelector(".drp-label");
    root.appendChild(trigger);

    // popover (visibility handled by the .drp-open class → smooth open + close)
    var pop = el("div","drp-pop");
    var cal = el("div","drp-cal"); pop.appendChild(cal);
    var presetsBar = null;
    if (wantPresets){
      presetsBar = el("div","drp-presets");
      [["7d",7],["30d",30],["90d",90],["This month","M"],["Last month","LM"]].forEach(function(p){
        var b = el("button","drp-preset"); b.type="button"; b.textContent = p[0];
        b.addEventListener("click", function(){ preset(p[1]); }); presetsBar.appendChild(b);
      });
      pop.appendChild(presetsBar);
    }
    var foot = el("div","drp-foot");
    var rlbl = el("div","drp-range-lbl");
    var btns = el("div","drp-foot-btns");
    var cancel = el("button","drp-btn ghost"); cancel.type="button"; cancel.textContent="Cancel";
    var apply  = el("button","drp-btn apply"); apply.type="button"; apply.textContent="Apply";
    btns.appendChild(cancel); btns.appendChild(apply);
    foot.appendChild(rlbl); foot.appendChild(btns); pop.appendChild(foot);
    root.appendChild(pop);

    function setLabel(){
      if (start && end){ label.textContent = nice(start) + "  –  " + nice(end); label.classList.remove("placeholder"); }
      else if (start){ label.textContent = nice(start) + "  –  …"; label.classList.remove("placeholder"); }
      else { label.textContent = "Select date range"; label.classList.add("placeholder"); }
    }
    function setFoot(){
      rlbl.innerHTML = "Range: <b>" + nice(start) + "</b> – <b>" + nice(end) + "</b>";
      apply.disabled = !(start && end);
    }

    function monthEl(base){
      var wrap = el("div","drp-month");
      var head = el("div","drp-mhead");
      var lab = el("div","drp-mlabel"); lab.textContent = MON[base.getMonth()] + " " + base.getFullYear();
      head.appendChild(lab); wrap.appendChild(head);
      var dow = el("div","drp-dow"); DOW.forEach(function(d){ var s=el("span"); s.textContent=d; dow.appendChild(s); });
      wrap.appendChild(dow);
      var grid = el("div","drp-grid");
      var first = new Date(base.getFullYear(), base.getMonth(), 1);
      var lead = first.getDay();
      var dim = new Date(base.getFullYear(), base.getMonth()+1, 0).getDate();
      for (var i=0;i<lead;i++){ var m=el("button","drp-day drp-mut"); m.type="button"; m.disabled=true; grid.appendChild(m); }
      for (var day=1; day<=dim; day++){
        (function(day){
          var d = new Date(base.getFullYear(), base.getMonth(), day);
          var b = el("button","drp-day"); b.type="button"; b.textContent = day;
          if (same(d,today)) b.classList.add("drp-today");
          if (d > today){ b.disabled = true; }        // no future dates
          else {
            // paint()-only on hover so the button is never rebuilt mid-click
            b.addEventListener("click", function(){ pick(d); });
            b.addEventListener("mouseenter", function(){ if(start && !end){ hoverDay=d; paint(); } });
          }
          cells.push({ el:b, d:d });
          grid.appendChild(b);
        })(day);
      }
      wrap.appendChild(grid);
      return wrap;
    }

    // Repaint range classes on the EXISTING day buttons (no DOM rebuild) — this
    // is what runs on hover + pick, so a click is never swallowed by a rebuild.
    function paint(){
      for (var i=0;i<cells.length;i++){
        var d = cells[i].d, cl = cells[i].el.classList;
        cl.remove("drp-start","drp-end","drp-inrange","drp-prev");
        if (start && same(d,start)) cl.add("drp-start");
        if (end && same(d,end)) cl.add("drp-end");
        if (start && end && d>start && d<end) cl.add("drp-inrange");
        if (start && !end && hoverDay){
          var lo = start<hoverDay?start:hoverDay, hi = start<hoverDay?hoverDay:start;
          if (d>lo && d<hi) cl.add("drp-prev");
        }
      }
      setFoot();
    }

    function render(){
      cal.innerHTML = ""; cells = [];
      var left = monthEl(view);
      // nav lives on the left month header
      var lh = left.querySelector(".drp-mhead");
      var nav = el("div","drp-nav");
      var pj = navBtn("«", function(){ view = addMonths(view,-12); render(); });
      var pm = navBtn("‹", function(){ view = addMonths(view,-1); render(); });
      nav.appendChild(pj); nav.appendChild(pm); lh.insertBefore(nav, lh.firstChild);
      cal.appendChild(left);
      var right = monthEl(addMonths(view,1));
      var rh = right.querySelector(".drp-mhead");
      var nav2 = el("div","drp-nav");
      var nm = navBtn("›", function(){ view = addMonths(view,1); render(); });
      var nj = navBtn("»", function(){ view = addMonths(view,12); render(); });
      nav2.appendChild(nm); nav2.appendChild(nj); rh.appendChild(nav2);
      cal.appendChild(right);
      paint();
    }
    function navBtn(txt, fn){ var b=el("button","drp-navbtn"); b.type="button"; b.textContent=txt; b.addEventListener("click",fn); return b; }

    function pick(d){
      if (!start || (start && end)){ start = d; end = null; }   // begin new range
      else {                                                     // pick the 2nd endpoint
        if (d < start){ end = start; start = d; } else { end = d; }
        if (maxDays && (dayIdx(end)-dayIdx(start)) > maxDays-1){  // enforce cap
          end = new Date(start.getTime() + (maxDays-1)*86400000);
        }
      }
      hoverDay = null; setLabel(); paint();
    }
    function preset(kind){
      var s, e = new Date(today);
      if (kind === "M"){ s = new Date(today.getFullYear(), today.getMonth(), 1); }
      else if (kind === "LM"){ s = new Date(today.getFullYear(), today.getMonth()-1, 1); e = new Date(today.getFullYear(), today.getMonth(), 0); }
      else { s = new Date(today.getTime() - (kind-1)*86400000); }
      start = s; end = e; view = addMonths(start,0); view.setDate(1);
      hoverDay=null; setLabel(); render();
    }

    function openPop(){ open=true; view = addMonths(start || end || today, 0); view.setDate(1); render();
      pop.classList.add("drp-open"); trigger.classList.add("open"); document.addEventListener("mousedown", outside); }
    function closePop(){ open=false; pop.classList.remove("drp-open"); trigger.classList.remove("open");
      document.removeEventListener("mousedown", outside); }
    function outside(e){ if(!root.contains(e.target)) closePop(); }

    trigger.addEventListener("click", function(){ open ? closePop() : openPop(); });
    cancel.addEventListener("click", closePop);
    apply.addEventListener("click", function(){
      hf.value = start ? iso(start) : ""; ht.value = end ? iso(end) : "";
      setLabel(); closePop();
      root.dispatchEvent(new CustomEvent("drp:apply", { bubbles:true, detail:{ from:hf.value, to:ht.value } }));
      if (doSubmit){ var f = root.closest("form"); if (f) f.submit(); }
    });
    document.addEventListener("keydown", function(e){ if(open && e.key==="Escape") closePop(); });

    setLabel();
  }

  function init(scope){ (scope||document).querySelectorAll("[data-drp]").forEach(function(n){
    if(!n.__drp){ n.__drp = true; try{ new DRP(n); }catch(err){ /* never break the page */ } } }); }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", function(){ init(); });
  else init();
  window.B2BDateRange = { init: init };
})();
