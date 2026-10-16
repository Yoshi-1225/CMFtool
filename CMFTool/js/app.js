/* CMF Tool — 面板共用：分頁、配色、狀態列，以及標註和 Excel 兩邊互相通知 */
(function () {
  'use strict';

  var cep = window.__adobe_cep__;
  var $ = function (id) { return document.getElementById(id); };
  var handlers = {};

  var App = window.CMFApp = {
    cep: cep,
    excel: null,     // js/excel.js 提供：cmfInfo()、numbersChanged()
    on: function (name, fn) { (handlers[name] = handlers[name] || []).push(fn); },
    emit: function (name, data) {
      (handlers[name] || []).forEach(function (fn) { try { fn(data); } catch (e) { console.error(e); } });
    },
    // 正在面板的欄位裡打字。Illustrator 執行面板的腳本時會把鍵盤焦點搶回文件，
    // 打字中途呼叫的話，後面的字就打不進欄位（例如只打得進兩位數的第一位）
    typing: function () {
      var el = document.activeElement;
      return !!el && document.hasFocus() &&
        (el.tagName === 'TEXTAREA' || (el.tagName === 'INPUT' && /^(text|number|search)$/.test(el.type)));
    },
    // 滑鼠移回面板、面板取得焦點時自動更新用：點進欄位準備打字時就不呼叫 Illustrator
    whenIdle: function (fn) {
      return function () { setTimeout(function () { if (!App.typing()) fn(); }, 60); };
    },
    setStatus: function (msg, isError) {
      var s = $('status');
      s.textContent = msg || '就緒';
      s.classList.toggle('error', !!isError);
    },
    // 吸取樣式：把標註、尺寸記錄的值放進面板的欄位（不觸發欄位的事件）。
    // 值不能用（記錄裡沒有、選單裡沒有這個選項、不是數字或顏色）時不改、回傳 false
    setField: function (el, v) {
      if (!el || v === undefined || v === null) return false;
      if (el.type === 'checkbox') { el.checked = !!v; return true; }
      if (el.tagName === 'SELECT') {
        v = String(v);
        for (var i = 0; i < el.options.length; i++) if (el.options[i].value === v) { el.value = v; return true; }
        return false;
      }
      if (el.type === 'color') {
        var m = /^#?([0-9a-f]{6}|[0-9a-f]{3})$/i.exec(String(v));
        if (!m) return false;
        var h = m[1].toLowerCase();
        el.value = '#' + (h.length === 3 ? h.replace(/./g, '$&$&') : h);
        return true;
      }
      if (App.isNum && App.isNum(el)) {
        var n = Number(v);
        if (typeof v === 'boolean' || v === '' || !isFinite(n)) return false;
        el.value = String(Math.round(n * 1000) / 1000);
        return true;
      }
      el.value = String(v);
      return true;
    },
    // 擴充功能資料夾（Windows：C:/Users/...，Mac：/Users/...）
    extPath: function () {
      var p = decodeURI(cep.getSystemPath('extension'));
      return /^file:\/\/\/[A-Za-z]:/.test(p) ? p.replace('file:///', '') : p.replace('file://', '');
    },
    showTab: function (name) {
      Array.prototype.forEach.call(document.querySelectorAll('.tab'), function (t) {
        t.classList.toggle('is-active', t.getAttribute('data-tab') === name);
      });
      Array.prototype.forEach.call(document.querySelectorAll('.page'), function (p) {
        p.hidden = p.id !== 'page-' + name;
      });
      try { localStorage.setItem('cmftool:tab', name); } catch (e) {}
      window.scrollTo(0, 0);
      App.emit('tab', name);
    }
  };

  /* ---------- 分頁 ---------- */
  Array.prototype.forEach.call(document.querySelectorAll('.tab'), function (t) {
    t.addEventListener('click', function () { App.showTab(t.getAttribute('data-tab')); });
  });
  try {
    var last = localStorage.getItem('cmftool:tab');
    if (last && $('page-' + last)) App.showTab(last);
  } catch (e) {}

  /* ---------- 重新載入外掛：換了新版的檔案後，不用重開 Illustrator ---------- */
  // Illustrator 只在載入擴充功能時執行 manifest 指定的 jsx/host.jsx，重新整理面板不會再執行。
  // 所以面板每次載入都自己再執行一次，重新載入時 Illustrator 端也換成新的（jsx/excel.jsx 本來就由 excel.js 載入）
  var RELOADED = 'cmftool:reloaded';
  if (cep) {
    var hostFile = JSON.stringify(App.extPath() + '/jsx/host.jsx').replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
    cep.evalScript('(function(){try{$.evalFile(' + hostFile + ');return "ok";}catch(e){return "error:"+e.message+(e.line?"（第 "+e.line+" 行）":"");}})()', function (r) {
      var reloaded = false;
      try { reloaded = Date.now() - Number(localStorage.getItem(RELOADED)) < 15000; localStorage.removeItem(RELOADED); } catch (e) {}
      if (r !== 'ok') App.setStatus('外掛的 jsx/host.jsx 載入失敗：' + String(r).replace(/^error:/, ''), true);
      else if (reloaded) App.setStatus('已重新載入外掛');
    });
  }

  // Node 載入過的外掛檔案（js/table.js、lib/xlsx）從快取拿掉，重新載入後才會讀到新版
  function reload() {
    try {
      var req = typeof cep_node !== 'undefined' && cep_node.require ? cep_node.require : typeof require === 'function' ? require : null;
      var root = cep ? App.extPath().replace(/\\/g, '/').toLowerCase() + '/' : null;
      if (req && req.cache && root) {
        Object.keys(req.cache).forEach(function (k) {
          if (k.replace(/\\/g, '/').toLowerCase().indexOf(root) === 0) delete req.cache[k];
        });
      }
    } catch (e) {}
    try { localStorage.setItem(RELOADED, String(Date.now())); } catch (e) {}
    location.reload();
  }
  $('btnReload').addEventListener('click', reload);

  /* ---------- 「重設大小」：依目前工作區的大小 ---------- */
  // 預設的字級、線寬⋯是給 A4（長邊 297 mm = 841.89 pt）的工作區用的。工作區較長的一邊是它的幾倍，預設值就乘幾倍。
  // cb(k, 說明)：k = 倍率（讀不到工作區時 1），說明 =「依工作區「01」420 × 297 mm，× 1.41」
  var A4_LONG = 841.89;
  App.artboardFactor = function (cb) {
    if (!cep) { cb(1, ''); return; }
    cep.evalScript('CMF.artboardSize()', function (raw) {
      var r = null;
      try { r = JSON.parse(raw); } catch (e) { r = null; }
      if (!r || !r.ok || !(Math.max(r.w, r.h) > 0)) { cb(1, ''); return; }
      var k = Math.max(0.05, Math.min(50, Math.max(r.w, r.h) / A4_LONG));
      var f = function (v) { return String(Math.round(v * 10) / 10); };
      cb(k, '依工作區「' + r.name + '」' + f(r.wmm) + ' × ' + f(r.hmm) + ' mm，× ' + (Math.round(k * 100) / 100));
    });
  };
  // 重設用的數值：預設 × k，留 3 位有效數字
  App.scaledDefault = function (el, k) {
    var v = (parseFloat(el.defaultValue) || 0) * k;
    if (!v) return '0';
    var p = Math.pow(10, Math.max(0, 2 - Math.floor(Math.log(Math.abs(v)) / Math.LN10)));
    return String(Math.round(v * p) / p);
  };

  /* ---------- CMF 圖層只放標註和尺寸 ---------- */
  // 每 1.5 秒請 host.jsx 的 guardLayers 檢查：CMF Callouts、CMF Dimensions 圖層裡別的東西搬到其他圖層，
  // 沒選取東西時作用中的圖層切回一般圖層。正在新增標註、兩點距離、取色（App.busy 裡有 true）或在面板打字時不檢查
  App.busy = {};
  (function () {
    var box = $('guardLayers'), KEY = 'cmftool:guardLayers', running = false;
    if (!box) return;
    try { var v = localStorage.getItem(KEY); if (v !== null) box.checked = v === '1'; } catch (e) {}
    function busy() { for (var k in App.busy) if (App.busy[k]) return true; return false; }
    function guard() {
      if (!cep || !box.checked || running || busy() || App.typing()) return;
      running = true;
      cep.evalScript('CMF.guardLayers()', function (raw) {
        running = false;
        var r;
        try { r = JSON.parse(raw); } catch (e) { return; }
        if (r && r.moved > 0) App.setStatus('已把 ' + r.moved + ' 個物件搬到「' + r.layer + '」（CMF 圖層只放標註和尺寸）');
      });
    }
    box.addEventListener('change', function () {
      try { localStorage.setItem(KEY, box.checked ? '1' : '0'); } catch (e) {}
      guard();
    });
    setInterval(guard, 1500);
  })();

  /* ---------- 收合區塊：記住開合狀態 ---------- */
  Array.prototype.forEach.call(document.querySelectorAll('details[data-key]'), function (d) {
    var key = 'cmftool:open:' + d.getAttribute('data-key');
    try {
      var v = localStorage.getItem(key);
      if (v !== null) d.open = v === '1';
    } catch (e) {}
    d.addEventListener('toggle', function () {
      try { localStorage.setItem(key, d.open ? '1' : '0'); } catch (e) {}
    });
  });

  /* ---------- 說明泡泡：滑鼠停在有 title 的東西上 0.45 秒，顯示圓角的說明（取代瀏覽器的小框） ----------
   * title 搬到 data-tip（之後程式再設 title 也會照新的搬），「名稱：說明」的名稱寫成粗體 */
  (function () {
    var tip = document.createElement('div'), cur = null, timer = null;
    tip.className = 'tip';
    tip.hidden = true;
    tip.setAttribute('role', 'tooltip');
    document.body.appendChild(tip);
    function textOf(el) {
      if (el.hasAttribute('title')) { el.setAttribute('data-tip', el.getAttribute('title')); el.removeAttribute('title'); }
      return el.getAttribute('data-tip') || '';
    }
    function hide() { clearTimeout(timer); tip.hidden = true; }
    function show(el, t) {
      if (!document.body.contains(el)) return;
      var i = t.indexOf('\uFF1A');
      tip.innerHTML = '';
      if (i > 0 && i <= 14) {
        var b = document.createElement('b'), sp = document.createElement('span');
        b.textContent = t.slice(0, i);
        sp.textContent = t.slice(i + 1);
        tip.appendChild(b);
        tip.appendChild(sp);
      } else tip.textContent = t;
      tip.hidden = false;
      var r = el.getBoundingClientRect(), w = tip.offsetWidth, h = tip.offsetHeight, vw = document.documentElement.clientWidth;
      var cx = (r.left + r.right) / 2, left = Math.max(6, Math.min(vw - w - 6, cx - w / 2)), below = r.top - h - 8 < 4;
      tip.classList.toggle('below', below);
      tip.style.left = left + 'px';
      tip.style.top = (below ? r.bottom + 8 : r.top - h - 8) + 'px';
      tip.style.setProperty('--tip-x', Math.max(10, Math.min(w - 10, cx - left)) + 'px');
    }
    document.addEventListener('mouseover', function (e) {
      var el = e.target.closest ? e.target.closest('[title], [data-tip]') : null;
      if (el === cur) return;
      hide();
      cur = el;
      var t = el ? textOf(el) : '';
      if (t) timer = setTimeout(function () { show(el, t); }, 450);
    });
    document.addEventListener('mousedown', function () { hide(); cur = null; }, true);
    document.addEventListener('keydown', hide, true);
    window.addEventListener('scroll', hide, true);
    window.addEventListener('blur', hide);
  })();

  /* ---------- 配色跟隨 Illustrator 介面亮度 ---------- */
  function applyTheme() {
    if (!cep) return;
    try {
      var c = JSON.parse(cep.getHostEnvironment()).appSkinInfo.panelBackgroundColor.color;
      var r = Math.round(c.red), g = Math.round(c.green), b = Math.round(c.blue);
      var light = (0.299 * r + 0.587 * g + 0.114 * b) / 255 > 0.55;
      var shift = function (d) {
        var f = function (v) { return Math.max(0, Math.min(255, v + d)); };
        return 'rgb(' + f(r) + ',' + f(g) + ',' + f(b) + ')';
      };
      var s = document.documentElement.style;
      document.documentElement.classList.toggle('light', light);
      s.setProperty('--bg', shift(0));
      s.setProperty('--field', light ? shift(28) : shift(-12));
      s.setProperty('--line', light ? shift(-40) : shift(20));
      s.setProperty('--line-soft', light ? shift(-20) : shift(9));
      s.setProperty('--pop', light ? shift(14) : shift(11));
      s.setProperty('--raise', light ? shift(40) : shift(24));
    } catch (e) {}
    App.emit('theme');
  }
  App.applyTheme = applyTheme;

  if (cep) {
    try { cep.addEventListener('com.adobe.csxs.events.ThemeColorChanged', applyTheme); } catch (e) {}
  }
  applyTheme();
})();
