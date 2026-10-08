/* CMF Tool — 數字欄位：中文輸入法開著也能直接打數字
 *
 * 注音等中文輸入法開著時，數字鍵會變成ㄅㄉˇˋ…進入組字，數字欄位收不到數字。
 * Chromium 在不能編輯的欄位上不啟用輸入法，所以數字欄位取得焦點時設成唯讀，
 * 按鍵就會以原本的字元進來；按下的當下暫時解除唯讀，讓瀏覽器照常插入、刪除、移動游標、
 * 複製貼上，處理完馬上鎖回去。唯讀時瀏覽器不畫游標，所以自己畫一條。
 *
 * 套用的欄位：
 *   <input type="number">：轉成文字欄位並加上 data-num（數字欄位讀不到游標位置），↑↓ 照 step 加減，
 *     Shift 一次 10 格；只接受數字、小數點、正負號。其他程式用 App.isNum(el) 判斷。
 *   有 inputmode="decimal" 或 data-ime="off" 的文字欄位（可以帶單位，例如 150mm、1:2）。
 * 之後才加入的欄位（例如編號表）也會自動套用。全形數字會換成半形。
 */
(function () {
  'use strict';

  var App = window.CMFApp;
  var LOCK = 'data-ime-lock';   // 這個欄位現在由這裡鎖成唯讀
  var NUM_KEYS = /^[0-9.,+\-eE]$/;
  var FULL = /[０-９．－＋：，]/g;   // ０-９ ． － ＋ ： ，

  App.isNum = function (el) { return !!el && el.tagName === 'INPUT' && el.hasAttribute('data-num'); };

  function applies(el) {
    return !!el && el.tagName === 'INPUT' && !el.disabled &&
      (el.hasAttribute('data-num') || el.getAttribute('inputmode') === 'decimal' || el.getAttribute('data-ime') === 'off');
  }

  function isLocked(el) { return !!el && el.nodeType === 1 && el.hasAttribute(LOCK); }

  // type="number" → 文字欄位（step、min、max 屬性留著，↑↓ 用）
  function convert(el) {
    if (el.type !== 'number') return;
    var v = el.value;
    el.type = 'text';
    el.value = v;
    el.setAttribute('data-num', '');
    el.setAttribute('inputmode', 'decimal');
    el.setAttribute('spellcheck', 'false');
    el.setAttribute('autocomplete', 'off');
  }

  Array.prototype.forEach.call(document.querySelectorAll('input[type=number]'), convert);
  new MutationObserver(function (list) {
    list.forEach(function (m) {
      Array.prototype.forEach.call(m.addedNodes, function (n) {
        if (n.nodeType !== 1) return;
        if (n.tagName === 'INPUT') convert(n);
        else Array.prototype.forEach.call(n.querySelectorAll('input[type=number]'), convert);
      });
    });
  }).observe(document.body, { childList: true, subtree: true });

  /* ---------- 自己畫的游標 ---------- */
  var caret = document.createElement('div');
  caret.className = 'ime-caret';
  caret.hidden = true;
  document.body.appendChild(caret);
  var measure = document.createElement('canvas').getContext('2d');

  function placeCaret() {
    var el = document.activeElement;
    if (!isLocked(el) || !el.readOnly || el.selectionStart !== el.selectionEnd) { caret.hidden = true; return; }
    var cs = getComputedStyle(el), r = el.getBoundingClientRect();
    measure.font = cs.fontStyle + ' ' + cs.fontWeight + ' ' + cs.fontSize + ' ' + cs.fontFamily;
    var padL = parseFloat(cs.paddingLeft) + parseFloat(cs.borderLeftWidth);
    var padR = parseFloat(cs.paddingRight) + parseFloat(cs.borderRightWidth);
    var inner = r.width - padL - padR, text = el.value;
    var full = measure.measureText(text).width, before = measure.measureText(text.slice(0, el.selectionStart)).width;
    var off = 0;
    if (full < inner) off = cs.textAlign === 'center' ? (inner - full) / 2 : /right|end/.test(cs.textAlign) ? inner - full : 0;
    var x = Math.round(r.left + padL + off + before - el.scrollLeft);
    var h = Math.round(parseFloat(cs.fontSize) * 1.3);
    caret.style.left = x + 'px';
    caret.style.top = Math.round(r.top + (r.height - h) / 2) + 'px';
    caret.style.height = h + 'px';
    caret.hidden = x < r.left + padL - 1 || x > r.right - padR + 1;
    caret.classList.remove('blink');
    void caret.offsetWidth;   // 重新開始閃爍，打字時游標保持顯示
    caret.classList.add('blink');
  }

  function placeSoon() { setTimeout(placeCaret, 0); }

  /* ---------- 按鍵不要傳給 Illustrator ---------- */
  // CEP：焦點不在可以打字的欄位時，按鍵也會傳給 Illustrator（數字鍵台會被當成方向鍵移動物件、
  // Backspace 會刪掉選取的物件、字母會切換工具）。鎖住的欄位是唯讀的，所以編輯期間跟 CEP 登記
  // 這些按鍵由面板自己處理（CSInterface.registerKeyEventsInterest），離開欄位就取消。
  var cep = window.__adobe_cep__;
  var KEY_INTEREST = (function () {
    var mac = /^Mac/.test(navigator.platform), list = [], i;
    function add(codes, mods) {
      codes.forEach(function (c) {
        (mods || [{}]).forEach(function (m) {
          var o = { keyCode: c };
          for (var k in m) o[k] = m[k];
          list.push(o);
        });
      });
    }
    var plainShift = [{}, { shiftKey: true }], cmd = mac ? 'metaKey' : 'ctrlKey', edit = {}, editShift = {};
    edit[cmd] = true; editShift[cmd] = true; editShift.shiftKey = true;
    if (mac) {
      // Mac 虛擬鍵碼：字母、數字列、數字鍵台、符號、空白、刪除、方向、Home/End、Return、Tab、Esc
      var keys = [];
      for (i = 0; i <= 50; i++) if (i !== 10) keys.push(i);
      add(keys.concat([51, 53, 65, 67, 69, 75, 76, 78, 81, 82, 83, 84, 85, 86, 87, 88, 89, 91, 92, 115, 117, 119, 123, 124, 125, 126]), plainShift);
      add([0, 6, 7, 8, 9, 51, 117, 123, 124], [edit]);         // ⌘A ⌘Z ⌘X ⌘C ⌘V、刪字、到行首行尾
      add([6, 123, 124], [editShift]);                          // ⌘⇧Z、選到行首行尾
    } else {
      // Windows 虛擬鍵碼：Backspace Tab Enter Esc 空白 End Home 方向 Delete、數字、字母、數字鍵台、符號
      var win = [8, 9, 13, 27, 32, 35, 36, 37, 38, 39, 40, 46];
      for (i = 48; i <= 57; i++) win.push(i);
      for (i = 65; i <= 90; i++) win.push(i);
      for (i = 96; i <= 111; i++) win.push(i);
      for (i = 186; i <= 192; i++) win.push(i);
      for (i = 219; i <= 222; i++) win.push(i);
      add(win, plainShift);
      add([65, 67, 86, 88, 89, 90, 8, 46, 35, 36, 37, 39], [edit]);   // Ctrl+A C V X Y Z、刪字、跳字、行首行尾
      add([90, 35, 36, 37, 39], [editShift]);
    }
    return JSON.stringify(list);
  })();

  var claimed = false;
  function claimKeys(on) {
    if (on === claimed || !cep || typeof cep.registerKeyEventsInterest !== 'function') return;
    try { cep.registerKeyEventsInterest(on ? KEY_INTEREST : ''); claimed = on; } catch (e) {}
  }

  /* ---------- 鎖定 / 暫時解鎖 ---------- */
  var pending = null;   // 暫時解鎖、等這次按鍵處理完的欄位

  function relock() {
    var el = pending;
    pending = null;
    if (el && isLocked(el) && document.activeElement === el) el.readOnly = true;
    placeCaret();
  }

  document.addEventListener('focusin', function (e) {
    var el = e.target;
    if (!applies(el) || (el.readOnly && !isLocked(el))) return;   // 原本就唯讀的欄位不動
    el.setAttribute(LOCK, '');
    el.readOnly = true;
    claimKeys(true);
    placeSoon();
  });

  document.addEventListener('focusout', function (e) {
    var el = e.target;
    if (!isLocked(el)) return;
    el.removeAttribute(LOCK);
    el.readOnly = false;
    if (pending === el) pending = null;
    caret.hidden = true;
    claimKeys(false);
  });

  function step(el, dir, times) {
    var s = parseFloat(el.getAttribute('step')) || 1, v = parseFloat(el.value);
    var min = parseFloat(el.getAttribute('min')), max = parseFloat(el.getAttribute('max'));
    v = isFinite(v) ? v + dir * s * times : (isFinite(min) ? min : 0);
    if (isFinite(min) && v < min) v = min;
    if (isFinite(max) && v > max) v = max;
    el.value = String(Math.round(v * 1e6) / 1e6);
    el.setSelectionRange(el.value.length, el.value.length);
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
    placeCaret();
  }

  // 先於欄位自己的處理（捕獲階段）
  document.addEventListener('keydown', function (e) {
    var el = e.target;
    if (!isLocked(el)) return;
    var plain = !e.ctrlKey && !e.metaKey && !e.altKey;
    if (App.isNum(el) && plain && (e.key === 'ArrowUp' || e.key === 'ArrowDown')) {
      e.preventDefault();
      step(el, e.key === 'ArrowUp' ? 1 : -1, e.shiftKey ? 10 : 1);
      return;
    }
    if (e.key === 'Tab' || e.key === 'Escape') return;
    var printable = e.key.length === 1 && plain;
    if (printable && App.isNum(el) && !NUM_KEYS.test(e.key)) { e.preventDefault(); return; }
    el.readOnly = false;
    caret.hidden = true;
    pending = el;
    // 字元在 keydown 之後才插入：等 input 事件（或放開按鍵）再鎖；其他按鍵的預設動作在 keydown 就做完了
    if (!printable) setTimeout(relock, 0);
  }, true);

  document.addEventListener('input', function (e) {
    var el = e.target;
    if (!isLocked(el)) return;
    if (!e.isComposing && FULL.test(el.value)) {
      var pos = el.selectionStart;
      el.value = el.value.replace(FULL, function (c) {
        var code = c.charCodeAt(0);
        return code >= 0xFF10 && code <= 0xFF19 ? String(code - 0xFF10) : { '．': '.', '－': '-', '＋': '+', '：': ':', '，': ',' }[c];
      });
      el.setSelectionRange(pos, pos);
    }
    if (pending === el) relock(); else placeCaret();
  }, true);

  document.addEventListener('keyup', function (e) {
    if (!isLocked(e.target)) return;
    if (pending === e.target) relock(); else placeCaret();
  }, true);

  ['mouseup', 'select'].forEach(function (name) {
    document.addEventListener(name, function (e) { if (isLocked(e.target)) placeSoon(); }, true);
  });
  document.addEventListener('selectionchange', function () { if (isLocked(document.activeElement)) placeCaret(); });
  window.addEventListener('scroll', function () { if (!caret.hidden) placeCaret(); }, true);
  window.addEventListener('resize', function () { if (!caret.hidden) placeCaret(); });
})();
