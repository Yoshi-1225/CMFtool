/* CMF Tool — 樣式的圖和設定卡片（標註分頁 js/callout.js、尺寸分頁 js/dimension.js 共用）
 *
 * 圖上的部位是 <g data-part="名稱">。點一下打開設定卡片，裡面是那個部位的設定（卡片裡的 .pane[data-pane="名稱"]）；
 * 卡片開著時，點（或按住拖曳）別的部位就換過去。卡片的 ✕、Esc、換分頁、收起區塊時關掉。
 * 卡片放在哪裡（都不會蓋住圖）：
 *   浮出：面板往右變寬（CEP 的 resizeContent），卡片出現在面板原本的右邊外面、對著那個部位；關掉時面板縮回原本的寬度。
 *     浮動的面板才會變寬；停靠的面板變不了寬，就改放在圖的下面（之後使用者自己改了面板大小，下次再試）。
 *   面板本來就夠寬（480 px 以上）：在圖的右邊，移到部位旁邊。其他：在圖的下面。
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
  var cep = window.__adobe_cep__;
  var FLY_W = 252, FLY_GAP = 8, WIDE = 480;    // 浮出的卡片寬、跟面板原本右邊的距離；面板寬度到 WIDE 時卡片在圖的右邊

  /* ---------- 浮出：面板往右變寬 ---------- */
  // ok：null = 還沒試過，true = 面板變得了寬，false = 變不了（停靠）；on：浮出中的卡片
  // k：resizeContent 的數值 × k = 面板實際的大小（Mac 的 Illustrator 2022、2023 會乘上介面縮放）
  var fly = { ok: null, on: null, w0: 0, added: 0, k: 1, ours: 0 };
  var root = document.documentElement;

  function canFly() { return !!(cep && cep.resizeContent) && fly.ok !== false && !fly.on && window.innerWidth < WIDE; }

  function hostResize(w, h) {
    fly.ours = Date.now();
    try { cep.resizeContent(Math.round(w / fly.k), Math.round(h / fly.k)); } catch (e) {}
  }

  // 等面板的大小改好（停了 80 ms 沒再變，最多等 450 ms）
  function settle(cb) {
    var t = null, done = false;
    var max = setTimeout(fin, 450);
    function fin() { if (done) return; done = true; clearTimeout(t); clearTimeout(max); window.removeEventListener('resize', on); cb(); }
    function on() { clearTimeout(t); t = setTimeout(fin, 80); }
    window.addEventListener('resize', on);
  }

  // 內容維持原本的寬度（右邊空出來給卡片）
  function hold(w) {
    root.classList.toggle('fly', !!w);
    document.body.style.width = w ? w + 'px' : '';
  }

  // 改面板大小一次做一件（例如剛關掉、面板還在縮回去時又點了部位）：job(finish)，做完呼叫 finish
  var queue = [];
  function enqueue(job) {
    queue.push(job);
    if (queue.length === 1) job(next);
  }
  function next() {
    queue.shift();
    if (queue.length) queue[0](next);
  }
  // 沒有在改大小時才做 cb
  function whenIdle(cb) { enqueue(function (finish) { finish(); cb(); }); }

  // 面板往右加寬：done(true) 變寬了、done(false) 沒變（恢復原狀）
  function widen(done) {
    enqueue(function (finish) {
      var w0 = window.innerWidth, h0 = window.innerHeight, want = w0 + FLY_W + FLY_GAP * 2;
      hold(w0);
      hostResize(want, h0);
      settle(function () {
        // 高度也變了 = 數值被乘上縮放：換算後再要一次
        if (h0 > 0 && Math.abs(window.innerHeight - h0) > 2) {
          fly.k *= window.innerHeight / h0;
          hostResize(want, h0);
          settle(check);
        } else check();
      });
      function check() {
        var added = window.innerWidth - w0, ok = added >= 200;
        if (ok) {
          fly.ok = true;
          fly.w0 = w0;
          fly.added = added;
        } else {
          fly.ok = false;
          if (Math.abs(added) > 2 || Math.abs(window.innerHeight - h0) > 2) hostResize(w0, h0);
          hold(0);
        }
        finish();
        done(ok);
      }
    });
  }

  // 縮回原本的寬度
  function unwiden() {
    enqueue(function (finish) {
      var added = fly.added;
      fly.added = 0;
      hostResize(window.innerWidth - added, window.innerHeight);
      settle(function () {
        if (!fly.on) hold(0);
        finish();
      });
    });
  }

  // 使用者自己改了面板大小：不再浮出（卡片照一般的位置放），之後可以再試
  window.addEventListener('resize', function () {
    if (Date.now() - fly.ours < 700) return;
    fly.ok = null;
    if (!fly.on) return;
    var c = fly.on;
    fly.on = null;
    fly.added = 0;
    hold(0);
    c.place();
  });
  // 重新載入、關掉面板時縮回去
  window.addEventListener('beforeunload', function () { if (fly.on) hostResize(window.innerWidth - fly.added, window.innerHeight); });

  /* ---------- 圖和卡片 ---------- */
  App.figCard = function (o) {
    var part = null, hint = true, open = false, byKey = false, ticket = 0, inst;
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
      if (open) { place(); if (focus) focusFirst(); return; }
      open = true;
      mark();
      var t = ++ticket;
      whenIdle(function () {   // 面板還在縮回去：等它好了再決定放哪裡
        if (t !== ticket || !open) return;
        if (!canFly()) { reveal(focus); return; }
        widen(function (ok) {
          if (t !== ticket || !open) { if (ok) unwiden(); return; }   // 等的時候關掉了
          if (ok) fly.on = inst;
          reveal(focus);
        });
      });
    }

    function reveal(focus) {
      o.card.hidden = false;
      place();
      if (focus) focusFirst();
    }

    // refocus：用鍵盤開的，關掉後回到圖上的部位
    function close(refocus) {
      if (!open) return;
      open = false;
      ticket++;
      if (o.card.contains(document.activeElement)) document.activeElement.blur();
      o.card.hidden = true;
      o.card.classList.remove('flying');
      o.card.style.left = o.card.style.top = o.card.style.width = '';
      if (fly.on === inst) { fly.on = null; unwiden(); }
      mark();
      if (refocus && byKey) { var g = partEl(part); if (g) g.focus(); }
    }

    // 卡片的位置和小三角（對準選到的部位的框）
    function place() {
      var g = partEl(part);
      if (!open || o.card.hidden || !g || o.card.offsetParent === null) return;   // 分頁沒開、區塊收起來
      var gr = (g.querySelector('.hl') || g).getBoundingClientRect(), mid = (gr.top + gr.bottom) / 2, v;
      o.card.style.marginTop = '';
      if (fly.on === inst) {
        // 浮出：面板原本的右邊外面，上下對著部位，不超出面板（分頁列、狀態列）
        o.card.classList.add('flying', 'side');
        o.card.style.left = (fly.w0 + FLY_GAP) + 'px';
        o.card.style.width = Math.min(FLY_W, fly.added - FLY_GAP * 2) + 'px';
        var h = o.card.offsetHeight, top0 = document.querySelector('.tabs').getBoundingClientRect().bottom + 4;
        var bottom = document.getElementById('status').getBoundingClientRect().top - 4;
        var top = Math.max(top0, Math.min(bottom - h, mid - h / 2));
        o.card.style.top = (top + window.pageYOffset) + 'px';
        v = mid - top;
        o.card.style.setProperty('--caret', Math.max(14, Math.min(h - 14, v)) + 'px');
        return;
      }
      o.card.classList.remove('flying');
      o.card.style.left = o.card.style.top = o.card.style.width = '';
      var cr = o.card.getBoundingClientRect(), fr = o.svg.getBoundingClientRect();
      var side = cr.left >= fr.right - 1;
      o.card.classList.toggle('side', side);
      if (side) {
        // 圖的右邊：卡片往下移到部位旁邊（不超出圖的高度）
        var off = Math.max(0, Math.min(fr.bottom - cr.top - cr.height, mid - cr.top - cr.height / 2));
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
    window.addEventListener('resize', function () { if (Date.now() - fly.ours >= 700) place(); });

    inst = {
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
    return inst;
  };
})();
