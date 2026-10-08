/* CMF Tool — 尺寸分頁：選取物件後標寬度、高度，圓的直徑 Ø，圓弧和圓角的半徑 R */
(function () {
  'use strict';

  var App = window.CMFApp;
  var cep = window.__adobe_cep__;
  var $ = function (id) { return document.getElementById(id); };
  var KEY = 'cmftool:dim';
  // 樣式欄位：記在每個尺寸上，「同步全部」會套用。id 去掉 dim 就是樣式的名稱（dimFontSize → fontSize）
  var STYLE = ['dimFontSize', 'dimTextColor', 'dimFontName', 'dimTextAlign', 'dimTextPos', 'dimTextGap', 'dimBreakGap', 'dimTextBg', 'dimTextBgColor',
    'dimLineWidth', 'dimLineColor', 'dimEndStyle', 'dimEndSize', 'dimArrowPos', 'dimExtLine', 'dimExtGap', 'dimExtOver',
    'dimUnit', 'dimDecimals', 'dimShowUnit', 'dimRatio', 'dimDiaSym', 'dimCount', 'dimPrefix', 'dimSuffix',
    'dimRadMode', 'dimLeader', 'dimShelf', 'dimCenterMark', 'dimAuto', 'dimAutoBase'];
  // 要標的類型（可以多選）和寬高的位置（上下左右，可以多選）
  var TYPES = ['dimTypeSel', 'dimTypeEach', 'dimTypeDia', 'dimTypeRad'];
  var POS = ['dimPosTop', 'dimPosBottom', 'dimPosLeft', 'dimPosRight'];
  var FIELDS = TYPES.concat(POS, ['dimOffset', 'dimAngle', 'dimVisible', 'dimAutoApply'], STYLE);
  var UNIT_PT = { mm: 72 / 25.4, cm: 72 / 2.54, 'in': 72, pt: 1, px: 1 };
  var DIA_SYMS = { slash: '\u00D8', phi: '\u03C6', sign: '\u2300' };
  var DIA = DIA_SYMS.slash;

  var info = null;   // CMF.dimInfo()：{ doc, sel, w, h, docUnit, docUnitPt }
  var busy = false;

  function call(fn, obj, cb) {
    if (!cep) { App.setStatus('請在 Illustrator 中開啟此面板', true); return; }
    var arg = obj ? JSON.stringify(JSON.stringify(obj)).replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029') : '';
    cep.evalScript('CMF.' + fn + '(' + arg + ')', function (raw) {
      var r;
      try { r = JSON.parse(raw); } catch (e) { r = { ok: false, msg: '腳本錯誤：' + raw }; }
      if (cb) cb(r);
    });
  }

  /* ---------- 設定 ---------- */
  function val(id) {
    var el = $(id);
    if (el.type === 'checkbox') return el.checked;
    if (App.isNum(el)) return parseFloat(el.value);
    return el.value;
  }

  // 數字欄位空白或不合理時用預設值
  function num(id, def, min) {
    var v = val(id);
    return isFinite(v) && v >= (min || 0) ? v : def;
  }

  function readStyle() {
    var o = {};
    STYLE.forEach(function (id) { o[id.charAt(3).toLowerCase() + id.slice(4)] = val(id); });
    o.fontSize = num('dimFontSize', 8, 0.1);
    o.lineWidth = num('dimLineWidth', 0.5, 0.01);
    o.endSize = num('dimEndSize', 4);
    o.textGap = num('dimTextGap', 2);
    o.breakGap = num('dimBreakGap', 2);
    o.extGap = num('dimExtGap', 2);
    o.extOver = num('dimExtOver', 2);
    o.leader = num('dimLeader', 12);
    // 依物件大小自動調整：樣式是給多大 (pt) 的物件用的，0 = 不調整（host.jsx 的 autoFactor）
    o.autoBase = o.auto ? num('dimAutoBase', 100, 0.001) * UNIT_PT.mm : 0;
    delete o.auto;
    return o;
  }

  function save() {
    var o = {};
    FIELDS.forEach(function (id) { o[id] = val(id); });
    try { localStorage.setItem(KEY, JSON.stringify(o)); } catch (e) {}
  }

  function load() {
    var o = null;
    try { o = JSON.parse(localStorage.getItem(KEY) || 'null'); } catch (e) {}
    if (!o) return;
    // 舊版是「寬在上／下」「高在右／左」兩個選單和「每個物件分別標寬高」
    if (o.dimPosTop === undefined && o.dimWSide !== undefined) {
      o.dimPosTop = o.dimWSide !== 'bottom';
      o.dimPosBottom = o.dimWSide === 'bottom';
      o.dimPosRight = o.dimHSide !== 'left';
      o.dimPosLeft = o.dimHSide === 'left';
      o.dimTypeEach = !!o.dimEach;
      o.dimTypeSel = !o.dimEach;
    }
    FIELDS.forEach(function (id) {
      if (o[id] === undefined || o[id] === null) return;
      var el = $(id);
      if (el.type === 'checkbox') el.checked = !!o[id]; else el.value = o[id];
    });
  }

  /* ---------- 數字 ---------- */
  // 跟 host.jsx 一樣：小數最多 dec 位，去掉尾數的 0
  function fmt(v, dec) {
    var s = v.toFixed(dec);
    if (s.indexOf('.') >= 0) s = s.replace(/0+$/, '').replace(/\.$/, '');
    return s === '-0' ? '0' : s;
  }

  // 圖面比例 1:2 → 2；只寫一個數字 n = 1:n；格式不對 = NaN
  function ratioOf(text) {
    var m = /^\s*(\d*\.?\d+)\s*[:\uFF1A\/]\s*(\d*\.?\d+)\s*$/.exec(text);
    if (m) { var a = parseFloat(m[1]), b = parseFloat(m[2]); return a > 0 && b > 0 ? b / a : NaN; }
    m = /^\s*(\d*\.?\d+)\s*$/.exec(text);
    return m && parseFloat(m[1]) > 0 ? parseFloat(m[1]) : NaN;
  }

  function unitOf() {
    var u = $('dimUnit').value;
    if (u === 'doc') return info && info.doc ? [info.docUnit, info.docUnitPt] : ['', 1];
    return [u, UNIT_PT[u]];
  }

  function decimals() { return parseInt($('dimDecimals').value, 10) || 0; }

  function diaSym() { return DIA_SYMS[$('dimDiaSym').value] || DIA; }

  /* ---------- 選取的物件 ---------- */
  function renderInfo() {
    var el = $('dimSel'), text = '';
    if (info && !info.doc) text = '沒有開啟的文件';
    else if (info && !info.sel) text = '請選取物件';
    else if (info) {
      var u = unitOf(), r = ratioOf($('dimRatio').value) || 1, dec = decimals();
      text = info.sel + ' 個物件 · ' + fmt(info.w * r / u[1], dec) + ' × ' + fmt(info.h * r / u[1], dec) + ' ' + u[0];
    }
    el.textContent = text;
    el.title = info && info.sel ? '選取的物件合起來的寬 × 高（依下面的單位和比例）' : '';
    renderAuto();
  }

  // 依物件大小自動調整：目前選取的物件會用幾倍的樣式（跟 host.jsx 的 autoFactor 一樣）
  var AUTO_MIN = 0.5, AUTO_MAX = 50;
  function renderAuto() {
    var on = $('dimAuto').checked, hint = $('dimAutoHint');
    $('dimAutoRow').hidden = hint.hidden = !on;
    if (!on) return;
    var base = num('dimAutoBase', 100, 0.001) * UNIT_PT.mm;
    if (!(info && info.sel)) { hint.textContent = '依量的物件大小，等比調整「樣式」的字級、線寬、箭頭和距離'; return; }
    var z = Math.max(info.w, info.h), k = Math.max(AUTO_MIN, Math.min(AUTO_MAX, z / base));
    hint.textContent = '選取的物件較長邊 ' + fmt(z / UNIT_PT.mm, 1) + ' mm → 樣式 × ' + fmt(k, 2) +
      '（字級 ' + fmt(num('dimFontSize', 8, 0.1) * k, 1) + ' pt）' + (k === AUTO_MIN ? '，已是最小' : '');
  }

  // 只在尺寸分頁開著時讀取
  var loading = false, again = false;
  function refresh() {
    if (!cep || $('page-dim').hidden) return;
    if (loading) { again = true; return; }
    loading = true;
    call('dimInfo', { visible: $('dimVisible').checked }, function (r) {
      loading = false;
      info = r && r.ok ? r : null;
      renderInfo();
      if (again) { again = false; refresh(); }
    });
  }

  /* ---------- 字體選單（字體清單跟標註分頁共用） ---------- */
  var fonts = {}, fontIndex = {}; // fontIndex：PostScript 名稱 → [家族, 字重]

  function esc(s) { return String(s).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;'); }

  function setFonts(list) {
    fonts = {}; fontIndex = {};
    (list || []).forEach(function (f) {
      (fonts[f[0]] = fonts[f[0]] || []).push([f[1], f[2]]);
      fontIndex[f[2]] = [f[0], f[1]];
    });
    var html = '<option value="">預設字體</option>';
    Object.keys(fonts).sort(function (a, b) { return a.localeCompare(b); }).forEach(function (name) {
      html += '<option value="' + esc(name) + '">' + esc(name) + '</option>';
    });
    $('dimFontFamily').innerHTML = html;
    var hit = fontIndex[$('dimFontName').value];
    $('dimFontFamily').value = hit ? hit[0] : '';
    fillStyles(hit ? hit[1] : null);
  }

  function fillStyles(selected) {
    var st = $('dimFontStyle'), styles = fonts[$('dimFontFamily').value] || [], pick = null;
    st.disabled = !styles.length;
    st.innerHTML = styles.map(function (s) { return '<option value="' + esc(s[1]) + '">' + esc(s[0]) + '</option>'; }).join('');
    if (!styles.length) return;
    styles.forEach(function (s) { if (s[0] === selected) pick = s[1]; });
    // 換家族時，優先選一般字重
    if (!pick) styles.forEach(function (s) { if (!pick && /^(regular|roman|book|normal|w3|medium)$/i.test(s[0])) pick = s[1]; });
    st.value = pick || styles[0][1];
  }

  function setFontName(name) {
    $('dimFontName').value = name;
    changed();
    scheduleApply();
  }

  /* ---------- 標註 ---------- */
  function setBusy(on) {
    busy = on;
    $('btnDimAdd').disabled = on;
  }

  // 選好的類型 × 位置 → 要標的項目；寬高：上下 = 寬度、左右 = 高度
  function jobs() {
    var out = [], sides = [['dimPosTop', 'w', 'top'], ['dimPosBottom', 'w', 'bottom'], ['dimPosLeft', 'h', 'left'], ['dimPosRight', 'h', 'right']];
    [['dimTypeSel', false], ['dimTypeEach', true]].forEach(function (t) {
      if (!$(t[0]).checked) return;
      sides.forEach(function (sd) { if ($(sd[0]).checked) out.push({ kind: sd[1], side: sd[2], each: t[1] }); });
    });
    if ($('dimTypeDia').checked) out.push({ kind: 'dia' });
    if ($('dimTypeRad').checked) out.push({ kind: 'rad' });
    return out;
  }

  function addAll() {
    if (busy) return;
    if (!(ratioOf($('dimRatio').value) > 0)) { App.setStatus('比例請寫成 1:2 這樣的格式', true); return; }
    var list = jobs(), wh = $('dimTypeSel').checked || $('dimTypeEach').checked;
    if (!list.length) {
      App.setStatus(wh ? '請選寬高要標在哪一邊（上下左右）' : '請先選要標的類型：寬高、直徑或半徑', true);
      return;
    }
    setBusy(true);
    App.setStatus('標註中…');
    call('dimAdd', {
      jobs: list,
      offset: num('dimOffset', 12),
      angle: num('dimAngle', 45, -1e9),
      visible: $('dimVisible').checked,
      style: JSON.stringify(readStyle())
    }, function (r) {
      setBusy(false);
      var msg = r.msg;
      // 依物件大小自動調整：樣式和基準改成剛建立的尺寸實際用的數值（兩個一起乘，下一個物件照樣算得對）
      if (r.ok && $('dimAuto').checked && r.k > 0 && Math.abs(r.k - 1) > 1e-6) {
        scaleSize(r.k);
        msg += '；樣式改成這個尺寸的數值（× ' + fmt(r.k, 2) + '）';
      }
      if (r.ok && wh && !POS.some(function (id) { return $(id).checked; })) msg += '（寬高沒有選位置，沒有標）';
      App.setStatus(msg, !r.ok);
    });
    if (!cep) setBusy(false);
  }

  function restyle(fn) {
    if (!(ratioOf($('dimRatio').value) > 0)) { App.setStatus('比例請寫成 1:2 這樣的格式', true); return; }
    call(fn, readStyle(), function (r) { App.setStatus(r.msg, !r.ok); });
  }

  // 「修改時自動套用到選取的尺寸」：停手一下才套用；沒有選取尺寸時不提示
  var applyTimer = null;
  function scheduleApply() {
    if (!$('dimAutoApply').checked || !cep) return;
    clearTimeout(applyTimer);
    applyTimer = setTimeout(function () {
      if (!(ratioOf($('dimRatio').value) > 0)) return;
      call('dimRestyle', readStyle(), function (r) { if (r.ok) App.setStatus(r.msg); });
    }, 600);
  }

  /* ---------- 事件 ---------- */
  function changed() {
    $('dimRatio').classList.toggle('invalid', !(ratioOf($('dimRatio').value) > 0));
    $('dimTextBgColor').disabled = $('dimTextBg').value !== 'box';
    $('dimBreakGap').disabled = $('dimTextPos').value !== 'middle';
    $('dimPad').classList.toggle('off', !$('dimTypeSel').checked && !$('dimTypeEach').checked);
    // 數量的選項跟著直徑符號
    var sym = diaSym(), opts = $('dimCount').options;
    opts[1].text = '4\u00D7' + sym + '3';
    opts[2].text = '4-' + sym + '3';
    save();
    renderInfo();
  }

  FIELDS.forEach(function (id) {
    var isStyle = STYLE.indexOf(id) >= 0;
    var handler = function () {
      changed();
      if (isStyle) scheduleApply();
    };
    $(id).addEventListener('input', handler);
    $(id).addEventListener('change', handler);
  });
  // 量寬高的範圍改了：重新讀取選取物件的寬高
  $('dimVisible').addEventListener('change', refresh);
  $('dimFontFamily').addEventListener('change', function () {
    fillStyles(null);
    setFontName($('dimFontFamily').value ? $('dimFontStyle').value : '');
  });
  $('dimFontStyle').addEventListener('change', function () { setFontName($('dimFontStyle').value); });
  $('btnDimFonts').addEventListener('click', function () { if (App.reloadFonts) App.reloadFonts(); });

  $('btnDimAdd').addEventListener('click', addAll);
  $('btnDimSync').addEventListener('click', function () { restyle('dimSync'); });
  $('btnDimRestyle').addEventListener('click', function () { restyle('dimRestyle'); });
  $('btnDimSelectAll').addEventListener('click', function () { call('dimSelectAll', null, function (r) { App.setStatus(r.msg, !r.ok); }); });
  $('btnDimToggle').addEventListener('click', function () { call('dimToggle', null, function (r) { App.setStatus(r.msg, !r.ok); }); });
  $('btnDimRefresh').addEventListener('click', refresh);

  // 在 Illustrator 裡換了選取之後，滑鼠移回面板時更新
  App.on('tab', function (name) { if (name === 'dim') refresh(); });
  // 縮放分頁縮放了文件：樣式裡跟大小有關的數值和「距離」用同一個倍率縮放，
  // 之後新增的尺寸、按「同步全部」才會跟縮放後的尺寸一樣大（跟 host.jsx 的 scaleDimStyle 一樣的欄位）
  // 基準也跟著縮放：樣式放大了，適合的物件也變大，自動調整的倍率才不會重複放大
  var SIZE_FIELDS = ['dimFontSize', 'dimLineWidth', 'dimEndSize', 'dimTextGap', 'dimBreakGap', 'dimExtGap', 'dimExtOver', 'dimLeader', 'dimOffset',
    'dimAutoBase'];
  function scaleSize(s) {
    SIZE_FIELDS.forEach(function (id) {
      var el = $(id);
      el.value = Math.round((parseFloat(el.value) || 0) * s * 1000) / 1000;
    });
    changed();
  }
  App.on('scaled', scaleSize);
  App.on('fonts', setFonts);
  var autoRefresh = App.whenIdle(refresh);
  document.documentElement.addEventListener('mouseenter', autoRefresh);
  window.addEventListener('focus', autoRefresh);
  if (cep) {
    try { cep.addEventListener('documentAfterActivate', refresh); } catch (e) {}
  }

  load();
  if (App.fontList) setFonts(App.fontList);
  changed();
  refresh();
})();
