/* CMF Tool — 樣式的圖和設定卡片（標註分頁 js/callout.js、尺寸分頁 js/dimension.js 共用）
 *
 * 圖上的部位是 <g data-part="名稱">。點一下打開設定卡片，裡面是那個部位的設定（卡片裡的 .pane[data-pane="名稱"]）；
 * 卡片開著時，點（或按住拖曳）別的部位就換過去。卡片的 ✕、Esc、換分頁、收起區塊時關掉。
 * 卡片放在圖的下面（不會蓋住圖；面板不會自己變寬）。
 * 圖上選到的部位框起來，卡片的小三角指著它。鍵盤：Tab 到部位上按 Enter，游標移到卡片的第一個欄位；Esc 關掉後回到部位。
 * 第一次點過圖之後，圖上的提示（.fig-hint）就不再出現。
 *
 *   var fc = App.figCard({ svg, card, title, names, key, hintKey, first })
 *     names：{ 部位: 卡片標題 } 或 function (部位) → 標題；key、hintKey：localStorage 記住選的部位、看過提示了；
 *     first：沒有記錄（或記錄的部位不在圖上）時選的部位
 *   fc.marked(部位)：圖上要不要框起來；fc.hint()：還要不要畫提示；fc.update()：圖重畫之後呼叫
 */
(function () {
  'use strict';

  var App = window.CMFApp;

  App.figCard = function (o) {
    var part = null, hint = true, open = false, byKey = false;
    try { part = localStorage.getItem(o.key); } catch (e) {}
    try { hint = localStorage.getItem(o.hintKey) !== '1'; } catch (e) {}

    function title(p) { return typeof o.names === 'function' ? o.names(p) : o.names[p]; }
    function parts() { return o.svg.querySelectorAll('[data-part]'); }
    function partEl(p) { return o.svg.querySelector('[data-part="' + p + '"]'); }

    // 卡片的標題和內容換成選到的部位
    function show() {
      o.title.textContent = title(part);
      Array.prototype.forEach.call(o.card.querySelectorAll('.pane'), function (el) { el.hidden = el.getAttribute('data-pane') !== part; });
    }

    function mark() {
      Array.prototype.forEach.call(parts(), function (g) { g.classList.toggle('sel', open && g.getAttribute('data-part') === part); });
    }

    // 卡片裡第一個看得到、可以用的欄位
    function focusFirst() {
      var f = Array.prototype.filter.call(o.card.querySelectorAll('.pane[data-pane="' + part + '"] input, .pane[data-pane="' + part + '"] select'),
        function (el) { return el.type !== 'hidden' && !el.disabled && el.offsetParent !== null; })[0];
      if (f) f.focus();
    }

    function setPart(p) {
      part = p;
      try { localStorage.setItem(o.key, p); } catch (e) {}
      show();
      mark();
    }

    // 打開（或換到）部位 p 的卡片；focus = 用鍵盤開的
    function openPart(p, focus) {
      setPart(p);
      byKey = !!focus;
      if (!open) {
        open = true;
        mark();
        o.card.hidden = false;
      }
      place();
      if (focus) focusFirst();
    }

    // refocus：用鍵盤開的，關掉後回到圖上的部位
    function close(refocus) {
      if (!open) return;
      open = false;
      if (o.card.contains(document.activeElement)) document.activeElement.blur();
      o.card.hidden = true;
      mark();
      if (refocus && byKey) { var g = partEl(part); if (g) g.focus(); }
    }

    // 卡片的小三角：對準選到的部位的框
    function place() {
      var g = partEl(part);
      if (!open || o.card.hidden || !g || o.card.offsetParent === null) return;   // 分頁沒開、區塊收起來
      var gr = (g.querySelector('.hl') || g).getBoundingClientRect(), cr = o.card.getBoundingClientRect();
      o.card.style.setProperty('--caret', Math.max(14, Math.min(cr.width - 14, (gr.left + gr.right) / 2 - cr.left)) + 'px');
    }

    function dropHint() {
      if (!hint) return;
      hint = false;
      try { localStorage.setItem(o.hintKey, '1'); } catch (e) {}
      var h = o.svg.querySelector('.fig-hint');
      if (h) h.parentNode.removeChild(h);
    }

    function partOf(e) { return e.target.closest && e.target.closest('[data-part]'); }

    o.svg.addEventListener('click', function (e) {
      var g = partOf(e);
      if (!g) return;
      openPart(g.getAttribute('data-part'), false);
      dropHint();
    });
    // 卡片開著時，按下去就換：拖曳改大小時，卡片裡的欄位跟著變
    o.svg.addEventListener('pointerdown', function (e) {
      var g = open && e.button === 0 && partOf(e);
      if (g && g.getAttribute('data-part') !== part) { setPart(g.getAttribute('data-part')); place(); }
    });
    o.svg.addEventListener('keydown', function (e) {
      var g = (e.key === 'Enter' || e.key === ' ') && partOf(e);
      if (!g) return;
      e.preventDefault();
      openPart(g.getAttribute('data-part'), true);
      dropHint();
    });
    var x = o.card.querySelector('.pop-close');
    if (x) x.addEventListener('click', function () { close(true); });
    // Esc 關掉（正在取色時 Esc 是取消取色；拖曳中的 Esc 由 js/numfield.js 處理）
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape' && open && !document.querySelector('.pick.active')) close(true);
    });
    var details = o.svg.closest('details');
    if (details) details.addEventListener('toggle', function () { if (!details.open) close(false); });
    App.on('tab', function () { close(false); });
    window.addEventListener('resize', place);

    return {
      place: place,
      marked: function (p) { return open && part === p; },
      hint: function () { return hint; },
      update: function () {
        var has = function (p) { return !!p && !!partEl(p); };
        if (!has(part)) {
          if (has(o.first)) setPart(o.first);
          else if (parts().length) setPart(parts()[0].getAttribute('data-part'));
        } else show();
        mark();
        place();
      }
    };
  };
})();
