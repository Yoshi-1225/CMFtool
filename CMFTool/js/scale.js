/* CMF Tool — 縮放分頁：選一個物件、輸入它的目標寬度或高度，整份文件（含工作區）一起等比縮放 */
(function () {
  'use strict';

  var App = window.CMFApp;
  var cep = window.__adobe_cep__;
  var $ = function (id) { return document.getElementById(id); };
  var SCOPE_KEY = 'cmftool:scaleScope';
  var FIELDS = { w: 'scaleW', h: 'scaleH', pct: 'scalePct' };
  // 輸入時可以加單位：150mm、15 cm、6in；沒寫單位 = 文件的單位
  var UNIT_PT = { pt: 1, px: 1, pc: 12, 'in': 72, '"': 72, mm: 72 / 25.4, cm: 72 / 2.54, m: 72 / 0.0254, q: 72 / 25.4 / 4 };

  var info = null;   // CMF.scaleInfo()：{ doc, unit, unitPt, preview, sel, w, h, ab, abName, artboards }
  var side = null;   // 使用者輸入的欄位：'w'、'h' 或 'pct'，另外兩個跟著算
  var busy = false;

  function call(fn, obj, cb) {
    if (!cep) { App.setStatus('請在 Illustrator 中開啟此面板', true); return; }
    var arg = obj ? JSON.stringify(JSON.stringify(obj)).replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029') : '';
    cep.evalScript('CMF.' + fn + '(' + arg + ')', function (raw) {
      var r;
      try { r = JSON.parse(raw); } catch (e) { r = { ok: false, msg: '腳本錯誤：' + raw }; }
      cb(r);
    });
  }

  function fmt(v) { return String(Math.round(v * 1000) / 1000); }

  // 回傳 pt
  function parseLen(text) {
    var m = /^\s*(\d+(?:\.\d*)?|\.\d+)\s*(mm|cm|m|in|"|pt|px|pc|q)?\s*$/i.exec(text);
    if (!m) return NaN;
    return parseFloat(m[1]) * (m[2] ? UNIT_PT[m[2].toLowerCase()] : info.unitPt);
  }

  function parsePct(text) {
    var m = /^\s*(\d+(?:\.\d*)?|\.\d+)\s*%?\s*$/.exec(text);
    return m ? parseFloat(m[1]) : NaN;
  }

  function hasSel() { return !!(info && info.doc && info.sel > 0 && info.w >= 0); }

  // 目前輸入的縮放倍率；沒有或無效 = NaN
  function factor() {
    if (!side || !info || !info.doc) return NaN;
    var text = $(FIELDS[side]).value;
    if (side === 'pct') { var p = parsePct(text); return p > 0 ? p / 100 : NaN; }
    if (!hasSel()) return NaN;
    var t = parseLen(text), cur = side === 'w' ? info.w : info.h;
    return t > 0 && cur > 0 ? t / cur : NaN;
  }

  function scope() {
    var el = document.querySelector('input[name=scaleScope]:checked');
    return el ? el.value : 'artboard';
  }

  function render() {
    var doc = !!(info && info.doc), sel = hasSel(), unit = doc ? info.unit : '';
    $('scaleSel').textContent = !info ? '' : !doc ? '沒有開啟的文件' : sel ? '已選取 ' + info.sel + ' 個物件' : '請選取物件';
    $('scaleSel').title = sel && info.preview ? '寬高跟「變形」面板一樣，含線寬（有勾「使用預視邊界」）' : '寬高跟「變形」面板一樣';
    $('scaleNowW').textContent = sel ? fmt(info.w / info.unitPt) + ' ' + unit : '—';
    $('scaleNowH').textContent = sel ? fmt(info.h / info.unitPt) + ' ' + unit : '—';
    Array.prototype.forEach.call(document.querySelectorAll('#page-scale .unit[data-len]'), function (el) { el.textContent = unit; });
    $('scaleW').disabled = $('scaleH').disabled = !sel;
    $('scalePct').disabled = !doc;
    $('scaleW').placeholder = sel ? fmt(info.w / info.unitPt) : '';
    $('scaleH').placeholder = sel ? fmt(info.h / info.unitPt) : '';
    $('scalePct').placeholder = doc ? '100' : '';

    // 另外兩個欄位跟著輸入的那個算
    var s = factor();
    Object.keys(FIELDS).forEach(function (k) {
      if (k === side) return;
      var v = '';
      if (s > 0) {
        if (k === 'pct') v = fmt(s * 100);
        else if (sel) v = fmt((k === 'w' ? info.w : info.h) * s / info.unitPt);
      }
      $(FIELDS[k]).value = v;
    });

    var where = doc && info.abName !== undefined ? '：' + info.abName : '';
    $('scaleAbLabel').textContent = (sel || !doc ? '物件所在的工作區' : '目前的工作區') + where;
    $('scaleAllLabel').textContent = '全部工作區' + (doc && info.artboards > 1 ? '（' + info.artboards + ' 個）' : '');
    $('btnScale').disabled = busy || !(s > 0) || Math.abs(s - 1) < 1e-6;
    $('btnScale').textContent = busy ? '縮放中…' : s > 0 && Math.abs(s - 1) >= 1e-6 ? '縮放 ' + fmt(s * 100) + '%' : '縮放';
  }

  function clearFields() {
    side = null;
    Object.keys(FIELDS).forEach(function (k) { $(FIELDS[k]).value = ''; });
  }

  // 只在縮放分頁開著時讀取，不影響其他分頁
  var loading = false, again = false;
  function refresh() {
    if (!cep || $('page-scale').hidden || busy) return;
    if (loading) { again = true; return; }
    loading = true;
    call('scaleInfo', null, function (r) {
      loading = false;
      info = r && r.ok ? r : null;
      render();
      if (again) { again = false; refresh(); }
    });
  }

  function doScale() {
    var s = factor();
    if (busy || !(s > 0) || Math.abs(s - 1) < 1e-6) return;
    var text = $(FIELDS[side]).value;
    var o = { side: side, value: side === 'pct' ? parsePct(text) : parseLen(text), scope: scope() };
    busy = true;
    render();
    App.setStatus('縮放中…');
    call('scaleDoc', o, function (r) {
      busy = false;
      // 全部工作區：標註、尺寸分頁的樣式也跟著縮放，新增的標註、尺寸和「同步全部」才會跟縮放後的一樣大
      if (r.ok && r.factor > 0 && o.scope === 'all') {
        App.emit('scaled', r.factor);
        r.msg += '；標註和尺寸的樣式也一起縮放';
      }
      App.setStatus(r.msg, !r.ok);
      if (r.ok) clearFields();
      render();
      refresh();
    });
  }

  /* ---------- 事件 ---------- */
  Object.keys(FIELDS).forEach(function (k) {
    var el = $(FIELDS[k]);
    el.addEventListener('input', function () {
      side = el.value.trim() ? k : null;
      render();
    });
    el.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') { e.preventDefault(); doScale(); }
      if (e.key === 'Escape') { clearFields(); render(); el.blur(); }
    });
  });

  Array.prototype.forEach.call(document.querySelectorAll('input[name=scaleScope]'), function (r) {
    r.addEventListener('change', function () {
      try { localStorage.setItem(SCOPE_KEY, scope()); } catch (e) {}
    });
  });
  try {
    var saved = localStorage.getItem(SCOPE_KEY);
    var radio = saved && document.querySelector('input[name=scaleScope][value="' + saved + '"]');
    if (radio) radio.checked = true;
  } catch (e) {}

  $('btnScale').addEventListener('click', doScale);
  $('btnScaleRefresh').addEventListener('click', refresh);

  /* ---------- 交換兩個物件的位置和大小 ---------- */
  var RATIO_KEY = 'cmftool:swapRatio';
  try { $('swapRatio').checked = localStorage.getItem(RATIO_KEY) !== '0'; } catch (e) {}
  $('swapRatio').addEventListener('change', function () {
    try { localStorage.setItem(RATIO_KEY, $('swapRatio').checked ? '1' : '0'); } catch (e) {}
  });
  $('btnSwap').addEventListener('click', function () {
    if (busy) return;
    call('swapItems', { keepRatio: $('swapRatio').checked }, function (r) {
      App.setStatus(r.msg, !r.ok);
      if (r.ok) refresh();
    });
  });

  // 在 Illustrator 裡換了選取之後，滑鼠移回面板時更新
  App.on('tab', function (name) { if (name === 'scale') refresh(); });
  var autoRefresh = App.whenIdle(refresh);
  document.documentElement.addEventListener('mouseenter', autoRefresh);
  window.addEventListener('focus', autoRefresh);
  if (cep) {
    try { cep.addEventListener('documentAfterActivate', refresh); } catch (e) {}
  }

  render();
  refresh();
})();
