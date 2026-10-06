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
    setStatus: function (msg, isError) {
      var s = $('status');
      s.textContent = msg || '就緒';
      s.classList.toggle('error', !!isError);
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
    } catch (e) {}
    App.emit('theme');
  }
  App.applyTheme = applyTheme;

  if (cep) {
    try { cep.addEventListener('com.adobe.csxs.events.ThemeColorChanged', applyTheme); } catch (e) {}
  }
  applyTheme();
})();
