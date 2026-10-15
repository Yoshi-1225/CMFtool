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
    } catch (e) {}
    App.emit('theme');
  }
  App.applyTheme = applyTheme;

  if (cep) {
    try { cep.addEventListener('com.adobe.csxs.events.ThemeColorChanged', applyTheme); } catch (e) {}
  }
  applyTheme();
})();
