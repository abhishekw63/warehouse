/* online_b2b/rules.html — page script (separated from template). */
(function () {
  var tabs = document.querySelectorAll('.seg-tab');
  var reduce = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  function stagger(cards) {
    cards.forEach(function (c, i) {
      c.classList.remove('in');
      if (reduce) { c.classList.add('in'); return; }
      // clear then re-add with a small per-card delay for a cascading reveal
      setTimeout(function () { c.classList.add('in'); }, 60 + i * 70);
    });
  }

  // Scroll-reveal the visible pane's cards as they enter the viewport.
  if (!reduce && 'IntersectionObserver' in window) {
    document.documentElement.classList.add('js-reveal');
    var io = new IntersectionObserver(function (entries) {
      entries.forEach(function (e) {
        if (e.isIntersecting) { e.target.classList.add('in'); io.unobserve(e.target); }
      });
    }, { threshold: 0.12, rootMargin: '0px 0px -6% 0px' });
    document.querySelectorAll('.seg-pane.on .rule-card.reveal').forEach(function (c) { io.observe(c); });
  }

  tabs.forEach(function (t) {
    t.addEventListener('click', function () {
      tabs.forEach(function (x) { x.classList.toggle('on', x === t); });
      var seg = t.getAttribute('data-seg');
      document.querySelectorAll('.seg-pane').forEach(function (p) {
        var show = p.getAttribute('data-pane') === seg;
        p.style.display = show ? '' : 'none';
        p.classList.toggle('on', show);
        if (show) {
          // re-trigger the pane fade + cascade its cards back in
          p.style.animation = 'none'; void p.offsetWidth; p.style.animation = '';
          stagger(p.querySelectorAll('.rule-card.reveal'));
        }
      });
    });
  });
})();
