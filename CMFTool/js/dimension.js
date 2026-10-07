/* CMF Tool — 尺寸分頁：選取物件後標寬度、高度，圓的直徑 Ø，圓弧和圓角的半徑 R */
(function () {
  'use strict';

  var App = window.CMFApp;
  var cep = window.__adobe_cep__;
  var $ = function (id) { return document.getElementById(id); };
  var KEY = 'cmftool:dim';
  // 樣式欄位：記在每個尺寸上，「同步全部」會套用。id 去掉 dim 就是樣式的名稱（dimFontSize → fontSize）
  var STYLE = ['dimFontSize', 'dimTextColor', 'dimFontName', 'dimTextAlign', 'dimTextPos', 'dimTextGap', 'dimTextBg', 'dimTextBgColor',
    'dimLineWidth', 'dimLineColor', 'dimEndStyle', 'dimEndSize', 'dimArrowPos', 'dimExtLine', 'dimExtGap', 'dimExtOver',
    'dimUnit', 'dimDecimals', 'dimShowUnit', 'dimRatio', 'dimDiaSym', 'dimCount', 'dimPrefix', 'dimSuffix',
    'dimRadMode', 'dimLeader', 'dimShelf', 'dimCenterMark'];
  var FIELDS = ['dimWSide', 'dimHSide', 'dimOffset', 'dimAngle', 'dimEach', 'dimVisible', 'dimAutoApply'].concat(STYLE);
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
    o.leader = num('dimLeader', 12);
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
    scheduleApply();
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
    // 有底色時文字畫在底色上，用原本的顏色
    var lc = shown(st.lineColor), tc = st.textBg === 'box' ? st.textColor : shown(st.textColor), font = cssFont(fs), th = fs * 0.72;
    var aligned = st.textAlign !== 'horizontal', kind = st.endStyle, middle = st.textPos === 'middle';
    var unit = st.showUnit ? unitOf()[0] : '', ratio = ratioOf(st.ratio) || 1;
    var pad = Math.max(0, Math.min(fs * 0.2, gap * 0.75)), minLen = Math.max(L, 2 * k);
    measureCtx.font = font;

    function pt(p) { return p[0].toFixed(2) + ' ' + p[1].toFixed(2); }
    function along(p, u, s) { return [p[0] + u[0] * s, p[1] + u[1] * s]; }
    function dot(p, q) { return p[0] * q[0] + p[1] * q[1]; }
    function line(pts, width) {
      parts.push('<path d="M' + pts.map(pt).join('L') + '" fill="none" stroke="' + lc + '" stroke-width="' + (width || lw) + '"/>');
    }
    function arrowLike(how) { return how === 'arrow' || how === 'open'; }
    function trim(how) { return how === 'arrow' ? L * 0.9 : 0; }
    // u：指向尖端的方向；斜線用尺寸線的方向
    function end(tip, u, how) {
      if (how === 'none' || !(L > 0)) return;
      if (how === 'dot') {
        parts.push('<circle cx="' + tip[0] + '" cy="' + tip[1] + '" r="' + L * 0.3 + '" fill="' + lc + '"/>');
      } else if (how === 'tick') {
        var v = [(u[0] + u[1]) * Math.SQRT1_2, (u[1] - u[0]) * Math.SQRT1_2]; // 往右上 45°（y 往下）
        parts.push('<path d="M' + pt(along(tip, v, -L / 2)) + 'L' + pt(along(tip, v, L / 2)) + '" stroke="' + lc +
          '" stroke-width="' + lw * 2 + '"/>');
      } else if (how === 'open') {
        var ow = L * 0.27, O = along(tip, u, -L);
        parts.push('<path d="M' + pt([O[0] - u[1] * ow, O[1] + u[0] * ow]) + 'L' + pt(tip) + 'L' + pt([O[0] + u[1] * ow, O[1] - u[0] * ow]) +
          '" fill="none" stroke="' + lc + '" stroke-width="' + lw + '" stroke-linejoin="round"/>');
      } else {
        var hw = L * 0.18, B = along(tip, u, -L);
        parts.push('<path d="M' + pt(tip) + 'L' + pt([B[0] - u[1] * hw, B[1] + u[0] * hw]) + 'L' +
          pt([B[0] + u[1] * hw, B[1] - u[0] * hw]) + 'Z" fill="' + lc + '"/>');
      }
    }
    // c：文字中心；angle：逆時針的角度；有底色時先畫一塊底
    function text(lb, c, angle) {
      var tw = measureCtx.measureText(lb).width, rot = ' transform="rotate(' + (-angle) + ' ' + pt(c) + ')"';
      if (st.textBg === 'box') {
        parts.push('<rect x="' + (c[0] - tw / 2 - pad).toFixed(2) + '" y="' + (c[1] - th / 2 - pad).toFixed(2) + '" width="' +
          (tw + pad * 2).toFixed(2) + '" height="' + (th + pad * 2).toFixed(2) + '"' + rot + ' fill="' + esc(st.textBgColor) + '"/>');
      }
      parts.push('<text x="' + c[0].toFixed(2) + '" y="' + c[1].toFixed(2) + '"' + rot + ' fill="' + tc + '" style=\'font:' + font +
        '\' text-anchor="middle" dominant-baseline="central">' + esc(lb) + '</text>');
    }
    function readAngle(u) { // SVG 方向 → 逆時針角度，(-90, 90]
      var a = Math.atan2(-u[1], u[0]) * 180 / Math.PI;
      return a > 90.01 ? a - 180 : a <= -89.99 ? a + 180 : a;
    }
    function textCenter(mid, n, tw, isAligned) {
      var ext = isAligned ? th / 2 : tw / 2 * Math.abs(n[0]) + th / 2 * Math.abs(n[1]);
      return along(mid, n, gap + ext);
    }
    function textAlong(tw, v, isAligned) { return isAligned ? tw / 2 : tw / 2 * Math.abs(v[0]) + th / 2 * Math.abs(v[1]); }
    function label(v, sym, q) {
      var s = (sym || '') + fmt(v * ratio, decimals()) + unit;
      if (q > 1 && st.count === 'x') s = q + '×' + s;
      else if (q > 1 && st.count === 'dash') s = q + '-' + s;
      return (st.prefix || '') + s + (st.suffix || '');
    }
    // 文字放在線中間：p → q（方向 a）的線在 mid 前後各空 half；放不下回傳 false，線由呼叫的人畫
    function split(p, q, a, mid, half) {
      if (dot([mid[0] - p[0], mid[1] - p[1]], a) - half < minLen || dot([q[0] - mid[0], q[1] - mid[1]], a) - half < minLen) return false;
      line([p, along(mid, a, -half)]);
      line([along(mid, a, half), q]);
      return true;
    }

    function linear(p1, p2, a, n, off, value) {
      var Q1 = along(p1, n, off), Q2 = along(p2, n, off);
      if (st.extLine !== 'none') {
        var ew = st.extLine === 'thin' ? Math.max(0.4, lw / 2) : lw;
        [[p1, Q1], [p2, Q2]].forEach(function (e) { if (off - eg > 0.5) line([along(e[0], n, eg), along(e[1], n, eo)], ew); });
      }
      var len = Math.hypot(Q2[0] - Q1[0], Q2[1] - Q1[1]), pts;
      var inside = st.arrowPos === 'inside' ? true : st.arrowPos === 'outside' ? false : len >= L * 2.6;
      if (arrowLike(kind)) pts = inside ? [along(Q1, a, trim(kind)), along(Q2, a, -trim(kind))] : [along(Q1, a, -L * 2.2), along(Q2, a, L * 2.2)];
      else if (kind === 'tick') pts = [along(Q1, a, -L * 0.6), along(Q2, a, L * 0.6)];
      else pts = [Q1, Q2];
      var lb = label(value), tw = measureCtx.measureText(lb).width, mid = [(Q1[0] + Q2[0]) / 2, (Q1[1] + Q2[1]) / 2];
      var isSplit = middle && (inside || !arrowLike(kind)) && split(pts[0], pts[1], a, mid, textAlong(tw, a, aligned) + gap);
      if (!isSplit) line(pts);
      if (kind === 'tick') { end(Q1, a, kind); end(Q2, a, kind); }
      else {
        var s = arrowLike(kind) && !inside ? 1 : -1;
        end(Q1, [a[0] * s, a[1] * s], kind);
        end(Q2, [-a[0] * s, -a[1] * s], kind);
      }
      text(lb, isSplit ? mid : textCenter(mid, n, tw, aligned), aligned ? readAngle(a) : 0);
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

    // 圓：直徑，用面板選的標法（自動 = 這個大小放不放得下）
    var mode = st.radMode, inner = mode === 'inside';
    var C = inner ? [192, 42] : [182, 50], r = inner ? 28 : 20;
    var ang = num('dimAngle', 45, -1e9) * Math.PI / 180, u = [Math.cos(ang), -Math.sin(ang)], back = [-u[0], -u[1]];
    var how = kind === 'tick' ? 'arrow' : kind, lb = label(40, diaSym()), tw = measureCtx.measureText(lb).width;
    if (mode !== 'leader' && mode !== 'through' && mode !== 'inside') {
      var span = aligned ? tw : 2 * (tw / 2 * Math.abs(u[0]) + th / 2 * Math.abs(u[1]));
      mode = r * 2 >= span + (how === 'none' ? 0 : L * 1.2) * 2 + gap * 4 ? 'inside' : 'leader';
    }
    // 引線：預覽框裡最長 30。文字太長時把圓往旁邊挪（跟 Illustrator 一樣不換邊，挪不下就超出預覽框）
    var lead = Math.min(30, Math.max(st.leader * k, L * 1.5, 2 * k)), sx = u[0] < -0.0001 ? -1 : 1;
    var shelf = st.shelf === 'end' || st.shelf === 'none' ? st.shelf : 'over', sl = Math.max(L * 1.5, fs * 0.6);
    var need = shelf === 'over' ? tw + gap * 2 : (shelf === 'end' ? sl : 0) + gap + tw;
    if (mode !== 'inside') {
      var lo = 150 + r, hi = 256 - r, reach = u[0] * (r + lead) + sx * need;
      if (sx > 0) hi = Math.min(hi, 252 - reach); else lo = Math.max(lo, 150 - reach);
      C[0] = sx > 0 ? Math.max(150 + r, Math.min(hi, C[0])) : Math.min(256 - r, Math.max(lo, C[0]));
    }
    var M = along(C, u, r), M1 = along(C, u, -r);
    parts.push('<circle cx="' + C[0] + '" cy="' + C[1] + '" r="' + r + '" fill="var(--product)"/>');
    if (mode === 'inside') {
      var ta = readAngle(u), up = [-Math.sin(ta * Math.PI / 180), -Math.cos(ta * Math.PI / 180)];
      var p = along(M1, u, trim(how)), q = along(M, u, -trim(how));
      var isSplit = middle && split(p, q, u, C, textAlong(tw, u, aligned) + gap);
      if (!isSplit) line([p, q]);
      end(M1, back, how);
      end(M, u, how);
      text(lb, isSplit ? C : textCenter(C, up, tw, aligned), aligned ? ta : 0);
    } else {
      // 引線（箭頭從外面指向圓心）或穿過圓心（箭頭往外指），拉到 K 之後轉水平
      var K = along(M, u, lead);
      var pts = [mode === 'through' ? along(M1, u, trim(how)) : along(M, u, trim(how)), K];
      if (mode === 'through') { end(M1, back, how); end(M, u, how); }
      else end(M, back, how);
      var c;
      if (shelf === 'over') {
        pts.push([K[0] + sx * (tw + gap * 2), K[1]]);
        c = [K[0] + sx * (tw / 2 + gap), K[1] - gap - th / 2];
      } else {
        var S = shelf === 'end' ? [K[0] + sx * sl, K[1]] : K;
        if (shelf === 'end') pts.push(S);
        c = [S[0] + sx * (gap + tw / 2), S[1]];
      }
      line(pts);
      text(lb, c, 0);
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
      visible: $('dimVisible').checked,
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
    // 數量的選項跟著直徑符號
    var sym = diaSym(), opts = $('dimCount').options;
    opts[1].text = '4\u00D7' + sym + '3';
    opts[2].text = '4-' + sym + '3';
    save();
    renderInfo();
    renderPreview();
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
  var SIZE_FIELDS = ['dimFontSize', 'dimLineWidth', 'dimEndSize', 'dimTextGap', 'dimExtGap', 'dimExtOver', 'dimLeader', 'dimOffset'];
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
