/* CMF Tool — 尺寸分頁：選取物件後標寬度、高度，圓的直徑 Ø，圓弧和圓角的半徑 R */
(function () {
  'use strict';

  var App = window.CMFApp;
  var cep = window.__adobe_cep__;
  var $ = function (id) { return document.getElementById(id); };
  var KEY = 'cmftool:dim';
  // 樣式欄位：記在每個尺寸上，「同步全部」會套用。id 去掉 dim 就是樣式的名稱（dimFontSize → fontSize）
  var STYLE = ['dimFontSize', 'dimTextColor', 'dimFontName', 'dimTextAlign', 'dimTextGap', 'dimLineWidth', 'dimLineColor',
    'dimEndStyle', 'dimEndSize', 'dimExtGap', 'dimExtOver', 'dimUnit', 'dimDecimals', 'dimShowUnit', 'dimRatio'];
  var FIELDS = ['dimWSide', 'dimHSide', 'dimOffset', 'dimAngle', 'dimEach'].concat(STYLE);
  var UNIT_PT = { mm: 72 / 25.4, cm: 72 / 2.54, 'in': 72, pt: 1, px: 1 };
  var DIA = '\u00D8';

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
    if (el.type === 'number') return parseFloat(el.value);
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
    o.extGap = num('dimExtGap', 2);
    o.extOver = num('dimExtOver', 2);
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
  }

  // 只在尺寸分頁開著時讀取
  var loading = false, again = false;
  function refresh() {
    if (!cep || $('page-dim').hidden) return;
    if (loading) { again = true; return; }
    loading = true;
    call('dimInfo', null, function (r) {
      loading = false;
      info = r && r.ok ? r : null;
      renderInfo();
      renderPreview();
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
    renderPreview();
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
  }

  // 預覽用：PostScript 名稱 → CSS 字體
  function cssFont(size) {
    var hit = fontIndex[$('dimFontName').value];
    if (!hit) return '400 ' + size + 'px sans-serif';
    var st = hit[1];
    var w = /thin|hairline|w1/i.test(st) ? 200 : /extralight|ultralight|light|w2|w3/i.test(st) ? 300
      : /semibold|demibold|w6/i.test(st) ? 600 : /extrabold|ultrabold|heavy|black|w8|w9/i.test(st) ? 800
      : /bold|w7/i.test(st) ? 700 : /medium|w5/i.test(st) ? 500 : 400;
    return (/italic|oblique/i.test(st) ? 'italic ' : '') + w + ' ' + size + 'px "' + hit[0].replace(/"/g, '') + '", sans-serif';
  }

  /* ---------- 預覽 ---------- */
  // 跟 host.jsx 的版面規則一樣，只是用 SVG 座標（y 往下）
  var measureCtx = document.createElement('canvas').getContext('2d');

  function isDark(hex) {
    var h = hex.replace('#', '');
    var r = parseInt(h.substr(0, 2), 16), g = parseInt(h.substr(2, 2), 16), b = parseInt(h.substr(4, 2), 16);
    return (0.299 * r + 0.587 * g + 0.114 * b) < 60;
  }

  // 深色介面上，黑色改成淺灰才看得到
  function shown(hex) {
    return document.documentElement.classList.contains('light') || !isDark(hex) ? hex : '#e6e6e6';
  }

  function renderPreview() {
    var st = readStyle(), k = 1.4, parts = [];
    var lw = Math.max(0.6, st.lineWidth * k), L = st.endSize * k, fs = Math.max(6, st.fontSize * k);
    var gap = st.textGap * k, eg = st.extGap * k, eo = st.extOver * k;
    var lc = shown(st.lineColor), tc = shown(st.textColor), font = cssFont(fs), th = fs * 0.72;
    var aligned = st.textAlign !== 'horizontal', kind = st.endStyle;
    var unit = st.showUnit ? unitOf()[0] : '', ratio = ratioOf(st.ratio) || 1;
    measureCtx.font = font;

    function pt(p) { return p[0].toFixed(2) + ' ' + p[1].toFixed(2); }
    function along(p, u, s) { return [p[0] + u[0] * s, p[1] + u[1] * s]; }
    function line(pts) {
      parts.push('<path d="M' + pts.map(pt).join('L') + '" fill="none" stroke="' + lc + '" stroke-width="' + lw + '"/>');
    }
    // u：指向尖端的方向；斜線用尺寸線的方向
    function end(tip, u, how) {
      if (how === 'none' || !(L > 0)) return;
      if (how === 'dot') {
        parts.push('<circle cx="' + tip[0] + '" cy="' + tip[1] + '" r="' + L * 0.3 + '" fill="' + lc + '"/>');
      } else if (how === 'tick') {
        var v = [(u[0] + u[1]) * Math.SQRT1_2, (u[1] - u[0]) * Math.SQRT1_2]; // 往右上 45°（y 往下）
        parts.push('<path d="M' + pt(along(tip, v, -L / 2)) + 'L' + pt(along(tip, v, L / 2)) + '" stroke="' + lc +
          '" stroke-width="' + lw * 2 + '"/>');
      } else {
        var hw = L * 0.18, B = along(tip, u, -L);
        parts.push('<path d="M' + pt(tip) + 'L' + pt([B[0] - u[1] * hw, B[1] + u[0] * hw]) + 'L' +
          pt([B[0] + u[1] * hw, B[1] - u[0] * hw]) + 'Z" fill="' + lc + '"/>');
      }
    }
    // c：文字中心；angle：逆時針的角度
    function text(label, c, angle) {
      parts.push('<text x="' + c[0].toFixed(2) + '" y="' + c[1].toFixed(2) + '" transform="rotate(' + (-angle) + ' ' + pt(c) +
        ')" fill="' + tc + '" style=\'font:' + font + '\' text-anchor="middle" dominant-baseline="central">' + esc(label) + '</text>');
    }
    function readAngle(u) { // SVG 方向 → 逆時針角度，(-90, 90]
      var a = Math.atan2(-u[1], u[0]) * 180 / Math.PI;
      return a > 90.01 ? a - 180 : a <= -89.99 ? a + 180 : a;
    }
    function textCenter(mid, n, tw, isAligned) {
      var ext = isAligned ? th / 2 : tw / 2 * Math.abs(n[0]) + th / 2 * Math.abs(n[1]);
      return along(mid, n, gap + ext);
    }
    function label(v, prefix) { return (prefix || '') + fmt(v * ratio, decimals()) + unit; }

    function linear(p1, p2, a, n, off, value) {
      var Q1 = along(p1, n, off), Q2 = along(p2, n, off);
      [[p1, Q1], [p2, Q2]].forEach(function (e) { if (off - eg > 0.5) line([along(e[0], n, eg), along(e[1], n, eo)]); });
      var len = Math.hypot(Q2[0] - Q1[0], Q2[1] - Q1[1]), inside = len >= L * 2.6;
      if (kind === 'arrow') line(inside ? [along(Q1, a, L * 0.9), along(Q2, a, -L * 0.9)] : [along(Q1, a, -L * 2.2), along(Q2, a, L * 2.2)]);
      else line(kind === 'tick' ? [along(Q1, a, -L * 0.6), along(Q2, a, L * 0.6)] : [Q1, Q2]);
      if (kind === 'tick') { end(Q1, a, kind); end(Q2, a, kind); }
      else {
        var s = kind === 'arrow' && !inside ? 1 : -1;
        end(Q1, [a[0] * s, a[1] * s], kind);
        end(Q2, [-a[0] * s, -a[1] * s], kind);
      }
      var lb = label(value), tw = measureCtx.measureText(lb).width;
      text(lb, textCenter([(Q1[0] + Q2[0]) / 2, (Q1[1] + Q2[1]) / 2], n, tw, aligned), aligned ? readAngle(a) : 0);
    }

    // 產品：圓角矩形，寬標在上、高標在右（跟面板設定的位置一樣）
    var x0 = 22, x1 = 118, y0 = 34, y1 = 74, off = 14;
    var wTop = $('dimWSide').value !== 'bottom', hRight = $('dimHSide').value !== 'left';
    if (!wTop) { y0 = 12; y1 = 52; }
    if (!hRight) { x0 = 50; x1 = 146; }
    parts.push('<rect x="' + x0 + '" y="' + y0 + '" width="' + (x1 - x0) + '" height="' + (y1 - y0) + '" rx="7" fill="var(--product)"/>');
    if (wTop) linear([x0, y0], [x1, y0], [1, 0], [0, -1], off, 120);
    else linear([x0, y1], [x1, y1], [1, 0], [0, 1], off, 120);
    if (hRight) linear([x1, y1], [x1, y0], [0, -1], [1, 0], off, 50);
    else linear([x0, y1], [x0, y0], [0, -1], [-1, 0], off, 50);

    // 圓：直徑
    var C = [182, 48], r = 20, ang = num('dimAngle', 45, -1e9) * Math.PI / 180, u = [Math.cos(ang), -Math.sin(ang)];
    var how = kind === 'tick' ? 'arrow' : kind, lb = label(40, DIA), tw = measureCtx.measureText(lb).width;
    parts.push('<circle cx="' + C[0] + '" cy="' + C[1] + '" r="' + r + '" fill="var(--product)"/>');
    var span = aligned ? tw : 2 * (tw / 2 * Math.abs(u[0]) + th / 2 * Math.abs(u[1]));
    var endRoom = how === 'none' ? 0 : L * 1.2, M = along(C, u, r);
    if (r * 2 >= span + endRoom * 2 + gap * 4) {
      var M1 = along(C, u, -r), ta = readAngle(u), up = [-Math.sin(ta * Math.PI / 180), -Math.cos(ta * Math.PI / 180)];
      line(how === 'arrow' ? [along(M1, u, L * 0.9), along(M, u, -L * 0.9)] : [M1, M]);
      end(M1, [-u[0], -u[1]], how);
      end(M, u, how);
      text(lb, textCenter(C, up, tw, aligned), aligned ? ta : 0);
    } else {
      var K = along(M, u, Math.max(8, L * 1.5)), sx = u[0] < -0.0001 ? -1 : 1, shelf = tw + gap * 2;
      if (K[0] + sx * shelf > 256 || K[0] + sx * shelf < 4) sx = -sx; // 預覽框放不下就換邊
      line([how === 'arrow' ? along(M, u, L * 0.9) : M, K, [K[0] + sx * shelf, K[1]]]);
      end(M, [-u[0], -u[1]], how);
      text(lb, [K[0] + sx * shelf / 2, K[1] - gap - th / 2], 0);
    }
    $('dimPreview').innerHTML = parts.join('');
  }

  /* ---------- 標註 ---------- */
  function add(kind) {
    if (busy) return;
    if (!(ratioOf($('dimRatio').value) > 0)) { App.setStatus('比例請寫成 1:2 這樣的格式', true); return; }
    busy = true;
    $('dimTools').classList.add('busy');
    App.setStatus('標註中…');
    call('dimAdd', {
      kind: kind,
      offset: num('dimOffset', 12),
      wSide: $('dimWSide').value,
      hSide: $('dimHSide').value,
      angle: num('dimAngle', 45, -1e9),
      each: $('dimEach').checked,
      style: JSON.stringify(readStyle())
    }, function (r) {
      busy = false;
      $('dimTools').classList.remove('busy');
      App.setStatus(r.msg, !r.ok);
    });
    if (!cep) { busy = false; $('dimTools').classList.remove('busy'); }
  }

  function restyle(fn) {
    if (!(ratioOf($('dimRatio').value) > 0)) { App.setStatus('比例請寫成 1:2 這樣的格式', true); return; }
    call(fn, readStyle(), function (r) { App.setStatus(r.msg, !r.ok); });
  }

  /* ---------- 事件 ---------- */
  function changed() {
    $('dimRatio').classList.toggle('invalid', !(ratioOf($('dimRatio').value) > 0));
    save();
    renderInfo();
    renderPreview();
  }

  FIELDS.forEach(function (id) {
    $(id).addEventListener('input', changed);
    $(id).addEventListener('change', changed);
  });
  $('dimFontFamily').addEventListener('change', function () {
    fillStyles(null);
    setFontName($('dimFontFamily').value ? $('dimFontStyle').value : '');
  });
  $('dimFontStyle').addEventListener('change', function () { setFontName($('dimFontStyle').value); });
  $('btnDimFonts').addEventListener('click', function () { if (App.reloadFonts) App.reloadFonts(); });

  Array.prototype.forEach.call(document.querySelectorAll('#dimTools .tool'), function (b) {
    b.addEventListener('click', function () { add(b.getAttribute('data-kind')); });
  });
  $('btnDimSync').addEventListener('click', function () { restyle('dimSync'); });
  $('btnDimRestyle').addEventListener('click', function () { restyle('dimRestyle'); });
  $('btnDimSelectAll').addEventListener('click', function () { call('dimSelectAll', null, function (r) { App.setStatus(r.msg, !r.ok); }); });
  $('btnDimToggle').addEventListener('click', function () { call('dimToggle', null, function (r) { App.setStatus(r.msg, !r.ok); }); });
  $('btnDimRefresh').addEventListener('click', refresh);

  // 在 Illustrator 裡換了選取之後，滑鼠移回面板時更新
  App.on('tab', function (name) { if (name === 'dim') refresh(); });
  // 縮放分頁縮放了全部工作區：樣式裡跟大小有關的數值和「距離」用同一個倍率縮放，
  // 之後新增的尺寸、按「同步全部」才會跟縮放後的尺寸一樣大（跟 host.jsx 的 scaleDimStyle 一樣的欄位）
  var SIZE_FIELDS = ['dimFontSize', 'dimLineWidth', 'dimEndSize', 'dimTextGap', 'dimExtGap', 'dimExtOver', 'dimOffset'];
  App.on('scaled', function (s) {
    SIZE_FIELDS.forEach(function (id) {
      var el = $(id);
      el.value = Math.round((parseFloat(el.value) || 0) * s * 1000) / 1000;
    });
    changed();
  });
  App.on('fonts', setFonts);
  App.on('theme', renderPreview);
  document.documentElement.addEventListener('mouseenter', refresh);
  window.addEventListener('focus', refresh);
  if (cep) {
    try { cep.addEventListener('documentAfterActivate', refresh); } catch (e) {}
  }

  load();
  if (App.fontList) setFonts(App.fontList);
  changed();
  refresh();
})();
