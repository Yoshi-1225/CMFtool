/* CMF Tool — 樣式的圖和設定卡片（標註分頁 js/callout.js、尺寸分頁 js/dimension.js 共用）
 *
 * 圖上的部位是 <g data-part="名稱">；點一下（或按住拖曳），卡片換成那個部位的設定（卡片裡的 .pane[data-pane="名稱"]）。
 * 卡片固定在圖的下面，面板夠寬時在圖的右邊（css 的 .fig-wrap），不會蓋住圖。
 * 圖上選到的部位框起來，卡片的小三角指著它。鍵盤：Tab 到部位上按 Enter，游標移到卡片的第一個欄位。
 * 第一次點過圖之後，圖上的提示（.fig-hint）就不再出現。
 *
 *   var fc = App.figCard({ svg, card, title, names, key, hintKey, first })
 *     names：{ 部位: 卡片標題 } 或 function (部位) → 標題；key、hintKey：localStorage 記住選的部位、看過提示了；
 *     first：一開始選的部位（沒有記錄、記錄的部位不在圖上時）
 *   fc.part()：選到的部位；fc.hint()：還要不要畫提示；fc.update()：圖重畫之後呼叫
 */
(function () {
  'use strict';

  var App = window.CMFApp;

  App.figCard = function (o) {
    var part = null, hint = true;
    try { part = localStorage.getItem(o.key); } catch (e) {}
    try { hint = localStorage.getItem(o.hintKey) !== '1'; } catch (e) {}

    function title(p) { return typeof o.names === 'function' ? o.names(p) : o.names[p]; }

    function parts() { return o.svg.querySelectorAll('[data-part]'); }

    // 卡片的標題和內容換成選到的部位
    function show() {
      o.title.textContent = title(part);
      Array.prototype.forEach.call(o.card.querySelectorAll('.pane'), function (el) { el.hidden = el.getAttribute('data-pane') !== part; });
      mark();
    }

    function select(p, focus) {
      part = p;
      try { localStorage.setItem(o.key, p); } catch (e) {}
      show();
      if (focus) {
        // 卡片裡第一個看得到、可以用的欄位
        var f = Array.prototype.filter.call(o.card.querySelectorAll('.pane[data-pane="' + p + '"] input, .pane[data-pane="' + p + '"] select'),
          function (el) { return el.type !== 'hidden' && !el.disabled && el.offsetParent !== null; })[0];
        if (f) f.focus();
      }
    }

    function mark() {
      Array.prototype.forEach.call(parts(), function (g) { g.classList.toggle('sel', g.getAttribute('data-part') === part); });
      placeCaret();
    }

    // 小三角對準選到的部位（它的框）。卡片在圖的右邊時：小三角在左邊，卡片往下移到部位旁邊（不超出圖的高度）
    function placeCaret() {
      var g = o.svg.querySelector('[data-part="' + part + '"]');
      if (!g || o.card.offsetParent === null) return;   // 分頁沒開、區塊收起來
      o.card.style.marginTop = '';
      var gr = (g.querySelector('.hl') || g).getBoundingClientRect(), cr = o.card.getBoundingClientRect(), fr = o.svg.getBoundingClientRect();
      var side = cr.left >= fr.right - 1, v;
      o.card.classList.toggle('side', side);
      if (side) {
        var mid = (gr.top + gr.bottom) / 2, off = Math.max(0, Math.min(fr.bottom - cr.top - cr.height, mid - cr.top - cr.height / 2));
        o.card.style.marginTop = off + 'px';
        v = mid - cr.top - off;
      } else v = (gr.left + gr.right) / 2 - cr.left;
      o.card.style.setProperty('--caret', Math.max(14, Math.min((side ? cr.height : cr.width) - 14, v)) + 'px');
    }

    function dropHint() {
      if (!hint) return;
      hint = false;
      try { localStorage.setItem(o.hintKey, '1'); } catch (e) {}
      var h = o.svg.querySelector('.fig-hint');
      if (h) h.parentNode.removeChild(h);
    }

    function partOf(e) { return e.target.closest && e.target.closest('[data-part]'); }

    // 按下去就換：拖曳改大小時，卡片裡的欄位跟著變
    o.svg.addEventListener('pointerdown', function (e) {
      var g = e.button === 0 && partOf(e);
      if (!g) return;
      if (g.getAttribute('data-part') !== part) select(g.getAttribute('data-part'));
      dropHint();
    });
    o.svg.addEventListener('keydown', function (e) {
      var g = (e.key === 'Enter' || e.key === ' ') && partOf(e);
      if (!g) return;
      e.preventDefault();
      select(g.getAttribute('data-part'), true);
      dropHint();
    });

    var details = o.svg.closest('details');
    if (details) details.addEventListener('toggle', placeCaret);
    App.on('tab', function () { setTimeout(placeCaret, 0); });
    window.addEventListener('resize', placeCaret);

    return {
      part: function () { return part; },
      hint: function () { return hint; },
      update: function () {
        var has = function (p) { return !!p && !!o.svg.querySelector('[data-part="' + p + '"]'); };
        if (has(part)) show();
        else if (has(o.first)) select(o.first);
        else if (parts().length) select(parts()[0].getAttribute('data-part'));
      }
    };
  };
})();
