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
    'dimRadMode', 'dimLeader', 'dimShelf', 'dimCenterMark', 'dimLenEnds', 'dimLenSide', 'dimAuto', 'dimAutoBase'];
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
    // 寬高的尺寸線離物件的距離：新增時用，「同步全部」、「套用選取」也把寬高的尺寸改成這個距離（host.jsx 的 rebuildDim）
    o.offset = num('dimOffset', 12);
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
    // 線條的預設顏色改成深灰：存著的還是舊的預設（黑色）就換掉，只做一次（沒有存過設定的也記下來，之後選的黑色不會被換掉）
    var firstFix = false;
    try { firstFix = !localStorage.getItem('cmftool:dimDefaults2'); localStorage.setItem('cmftool:dimDefaults2', '1'); } catch (e) {}
    if (!o) return;
    if (firstFix && String(o.dimLineColor || '').toLowerCase() === '#000000') o.dimLineColor = '#404040';
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
      if ($('dimUnit').value === 'doc') renderFigure();   // 文件單位
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
    syncFontSelects();
    if (fc) renderFigure();   // 圖上的字體
  }

  // 依目前的 PostScript 名稱，把字體、字重兩個選單切到對應的位置
  function syncFontSelects() {
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

  /* ---------- 兩點距離 ---------- */
  // 按「兩點距離」切到鋼筆，點起點、點終點就標一個（host.jsx 的 dimLenPoll），標好就結束、切回選取工具；
  // 點第二下之前再按一次或 Esc 取消。
  // 跟「新增標註」都用鋼筆，開始一個時會結束另一個（App 的 'pen' 事件）
  var lenOn = false, lenTimer = null, lenPolling = false, lenIdle = 0;
  var LEN_IDLE = 60000;   // 閒置太久自動結束，避免把之後畫的線當成兩點距離

  function lenButton(on) {
    $('btnDimLen').classList.toggle('active', on);
    $('btnDimLen').setAttribute('aria-pressed', String(on));
  }

  function startLen() {
    if (!cep) { App.setStatus('請在 Illustrator 中開啟此面板', true); return; }
    if (!(ratioOf($('dimRatio').value) > 0)) { App.setStatus('比例請寫成 1:2 這樣的格式', true); return; }
    App.emit('pen', 'dim');
    call('dimLenBegin', null, function (r) {
      if (!r.ok) { App.setStatus(r.msg, true); return; }
      lenOn = true;
      App.busy.dimLen = true;
      lenIdle = Date.now();
      lenButton(true);
      if (figMode !== 'len') setFigMode('len');   // 示範框換成兩點距離，看得到現在的樣式
      App.setStatus(r.msg || ($('dimLenEnds').value === 'one' ? '點起點（箭頭指的地方）→ 點終點' : '點起點 → 點終點') +
        '，按住 Shift 鎖水平、垂直。Esc 取消');
      lenTimer = setInterval(pollLen, 200);
    });
  }

  // msg：結束的原因（省略 = 「已結束兩點距離」）；keepTool：不切回選取工具
  function stopLen(keepTool, msg) {
    if (!lenOn) return;
    clearInterval(lenTimer);
    lenTimer = null;
    lenOn = false;
    App.busy.dimLen = false;
    lenButton(false);
    call('dimLenEnd', { keepTool: !!keepTool }, function (r) { if (msg !== false) App.setStatus(msg || r.msg, !r.ok); });
  }

  function pollLen() {
    if (lenPolling || !lenOn) return;
    lenPolling = true;
    call('dimLenPoll', readStyle(), function (r) {
      lenPolling = false;
      if (!lenOn) return;
      if (r.state === 'drawing') { lenIdle = Date.now(); if (!r.ok) App.setStatus(r.msg, true); return; }
      if (r.state === 'done') { stopLen(false, r.msg); return; }   // 標好一個就結束，切回選取工具
      if (r.state === 'other') { stopLen(true, r.msg); return; }   // 選了其他物件：取消，工具維持使用者選的
      if (r.state !== 'idle') { App.setStatus(r.msg, true); stopLen(true, false); return; }
      if (Date.now() - lenIdle > LEN_IDLE) stopLen(true, '閒置過久，已結束兩點距離');
    });
  }

  // 每個部位的設定（樣式的名稱）：示範框上選著一個部位時，「同步全部」、「套用選取」、自動套用、「吸取樣式」只處理這些
  var PART_FIELDS = {
    text: ['fontSize', 'textColor', 'fontName', 'textAlign', 'textPos', 'textGap', 'breakGap', 'textBg', 'textBgColor',
      'unit', 'decimals', 'showUnit', 'ratio', 'diaSym', 'count', 'prefix', 'suffix'],
    line: ['lineWidth', 'lineColor', 'offset', 'radMode', 'leader', 'shelf', 'centerMark', 'lenSide'],
    end: ['endStyle', 'endSize', 'arrowPos', 'lenEnds'],
    ext: ['extLine', 'extGap', 'extOver']
  };

  // 送給 host.jsx 的樣式：選著一個部位時帶上 only（那個部位的欄位），其他設定每個尺寸維持原樣
  function syncStyle() {
    var o = readStyle(), part = fc.current();
    if (part && PART_FIELDS[part]) { o.only = PART_FIELDS[part]; o.onlyName = partName(part); }
    return o;
  }

  function restyle(fn) {
    if (!(ratioOf($('dimRatio').value) > 0)) { App.setStatus('比例請寫成 1:2 這樣的格式', true); return; }
    call(fn, syncStyle(), function (r) { App.setStatus(r.msg, !r.ok); });
  }

  // 「修改時自動套用到選取的尺寸」：停手一下才套用；沒有選取尺寸時不提示
  var applyTimer = null;
  // full：不管選著的部位，整份樣式（「重設」大小一次改好幾個部位）
  function scheduleApply(full) {
    if (!$('dimAutoApply').checked || !cep) return;
    clearTimeout(applyTimer);
    var st = full ? readStyle() : syncStyle();
    applyTimer = setTimeout(function () {
      if (!(ratioOf($('dimRatio').value) > 0)) return;
      call('dimRestyle', st, function (r) { if (r.ok) App.setStatus(r.msg); });
    }, 600);
  }

  /* ---------- 吸取樣式 ---------- */
  // 選一個尺寸按「吸取樣式」：面板的樣式改成這個尺寸記錄的樣式（「套用選取」寫進去的那一份），
  // 之後新增的尺寸、按「同步全部」就跟它一樣。只改面板，不會觸發「修改時自動套用到選取的尺寸」。
  // 記錄裡沒有的項目（舊版的尺寸）或不認得的值，用 host.jsx 遇到時的做法，才會跟這個尺寸重新產生時一樣
  var PICK_DEFAULT = { textColor: '#000000', fontName: '', textAlign: 'aligned', textPos: 'above', textBg: 'none',
    textBgColor: '#ffffff', lineColor: '#000000', endStyle: 'arrow', arrowPos: 'auto', extLine: 'same', unit: 'doc',
    showUnit: false, ratio: '1:1', diaSym: 'slash', count: 'none', prefix: '', suffix: '', radMode: 'auto', shelf: 'over',
    centerMark: true, lenEnds: 'both', lenSide: 'a' };

  // r = CMF.pickStyle()：{ dim, dimGeom, dimName, dims }
  // 示範框上選著一個部位時，只吸取那個部位的設定
  function applyPicked(r) {
    var st = r.dim, gm = r.dimGeom || {}, lin = gm.t === 'lin' && !gm.free;   // 兩點距離沒有「距離」
    var part = fc.current(), keys = part && PART_FIELDS[part], want = function (k) { return !keys || keys.indexOf(k) >= 0; };
    STYLE.forEach(function (id) {
      if (id === 'dimAuto' || id === 'dimAutoBase') return;   // 自動調整是新增時的做法，不算樣式；基準在下面
      var key = id.charAt(3).toLowerCase() + id.slice(4), v = st[key];
      if (!want(key)) return;
      if (key === 'breakGap' && v === undefined) v = st.textGap;   // 舊的樣式沒有「斷開留空」，用「文字離線」
      // 舊版的直徑、半徑沒有「引線」，長度記在幾何的 off（同心圓的內圈多拉了一段，不算）
      if (key === 'leader' && v === undefined && !lin && gm.x === undefined && !gm.out) v = gm.off;
      if (key === 'ratio' && !(ratioOf(String(v)) > 0)) v = undefined;   // 格式不對時 host.jsx 當成 1:1
      if (!App.setField($(id), v) && PICK_DEFAULT[key] !== undefined) App.setField($(id), PICK_DEFAULT[key]);
    });
    // 寬高的「距離」；依物件大小自動調整時，基準改成這個尺寸量的物件大小：一樣大的物件標出來就跟它一樣
    if (lin && want('offset')) App.setField($('dimOffset'), gm.off);
    if (gm.free && gm.side && want('lenSide')) App.setField($('dimLenSide'), gm.side);   // 兩點距離：文字實際在哪一側
    if ($('dimAuto').checked && gm.z > 0 && !keys) App.setField($('dimAutoBase'), gm.z / UNIT_PT.mm);
    syncFontSelects();
    changed();
    var msg = '已吸取' + (r.dimName ? '「' + r.dimName + '」' : '尺寸') + '的' + (keys ? partName(part) : '樣式');
    if (r.dims > 1) msg += '（選了 ' + r.dims + ' 個尺寸，吸取的是這一個）';
    if (st.fontName && Object.keys(fontIndex).length && !fontIndex[st.fontName]) msg += '；字體清單裡沒有「' + st.fontName + '」';
    App.setStatus(msg);
  }

  function pickStyle() {
    call('pickStyle', null, function (r) {
      if (r.ok && r.dim) applyPicked(r);
      else if (r.ok && r.callout) { App.showTab('callout'); App.emit('pickedCallout', r); }   // 選的是標註：改標註分頁的樣式
      else App.setStatus(r.msg || '請先選取要吸取樣式的尺寸或標註', true);
    });
  }

  /* ---------- 樣式的示範框 ---------- */
  // 照目前的樣式畫一個尺寸，畫法跟 host.jsx 的 buildLinear、buildRadial 一樣（座標 y 朝上，輸出時翻過來）。
  // 左上角切換：矩形 = 上面標寬度、右邊標高度；圓 = 標直徑（標法、角度、引線照卡片裡的設定）；
  // 兩點距離 = 兩塊物件中間的空隙標一個箭頭（沒有延伸線，箭頭只畫一端時畫在左邊的起點）。
  // 點文字、尺寸線（引線）、端點、延伸線，打開那個部位的設定卡片（js/figcard.js，跟標註的圖一樣）；
  // 在部位上按住左右拖曳改大小：文字 = 字級、線 = 線寬、端點 = 大小、延伸線 = 留空（js/numfield.js 的拖曳把手）。
  // 圖上的大小 = pt × 2，太小、太大的畫成看得到、放得下的大小；物件自動縮小、置中，整個尺寸放得進框裡。
  // 標的數字是示範用的大小（寬 42.5、高 30、直徑 24、兩點距離 12 mm），照面板的單位、比例、小數、符號和前後綴顯示
  var FIG = 240, K = 2, BAND = 30;           // BAND：上面留給左上角的切換按鈕和提示
  var SAMPLE = { w: 42.5, h: 30, d: 24, len: 12 };
  var FIG_KEY = 'cmftool:dimFig';
  var PART_NAMES = { text: '文字', line: '尺寸線', end: '端點', ext: '延伸線' };
  var ZOOM_C = [10 + (FIG - 20) / 2, FIG - 10 - (FIG - BAND - 10) / 2];   // 整張圖縮小時（f.zoom）的中心
  var figMode = 'rect', drawnRad = null;     // drawnRad：圓實際畫的標法（自動 → inside 或 leader）
  var drawn = null;                          // 畫出來的圖（矩形四邊的圓點照它放）
  function partName(p) { return p === 'line' && figMode === 'circle' && drawnRad !== 'inside' ? '引線' : PART_NAMES[p]; }
  var fc = App.figCard({ svg: $('dimPreview'), card: $('dimPop'), title: $('dimPopTitle'), names: partName, first: 'text',
    key: 'cmftool:dimPart', hintKey: 'cmftool:dimPartHint' });

  function vsub(p, q) { return [p[0] - q[0], p[1] - q[1]]; }
  function vdot(p, q) { return p[0] * q[0] + p[1] * q[1]; }
  function vlen(p) { return Math.sqrt(p[0] * p[0] + p[1] * p[1]); }
  function along(p, u, d) { return [p[0] + u[0] * d, p[1] + u[1] * d]; }
  function dirOf(deg) { var a = deg * Math.PI / 180; return [Math.cos(a), Math.sin(a)]; }
  function clamp(v, lo, hi) { return isFinite(v) ? Math.max(lo, Math.min(hi, v)) : lo; }
  // 線的方向 → 文字角度：(-90°, 90°]，文字由左往右、由下往上讀
  function readAngle(u) {
    var a = Math.atan2(u[1], u[0]) * 180 / Math.PI;
    if (a > 90.01) a -= 180; else if (a <= -89.99) a += 180;
    return a;
  }

  // 圖上用的樣式（圖上的單位）
  // 太大時字、線、端點、間距全部用同一個倍率 z 縮小，比例才跟實際標出來的一樣；太小的才各自放大到看得到
  function figStyle() {
    var st = readStyle(), mode = st.radMode;
    var z = Math.min(1, 40 / (st.fontSize * K), 24 / Math.max(st.endSize * K, 1e-6), 6 / (st.lineWidth * K),
      20 / Math.max(st.textGap * K, st.breakGap * K, st.extGap * K, st.extOver * K, 1e-6));
    var k = K * z, lw = Math.max(0.6, st.lineWidth * k), gap = st.textGap * k;
    return {
      st: st, lw: lw, col: st.lineColor, fs: Math.max(8, st.fontSize * k), L: Math.max(3, st.endSize * k), end: st.endStyle,
      textGap: gap, besideGap: gap + lw / 2, breakGap: st.breakGap * k,
      extGap: st.extGap * k, extOver: st.extOver * k, extLine: st.extLine,
      arrowPos: st.arrowPos, aligned: st.textAlign !== 'horizontal', textPos: st.textPos,
      leader: clamp(st.leader * k, 0, 60), shelf: st.shelf === 'end' || st.shelf === 'none' ? st.shelf : 'over',
      radMode: mode === 'leader' || mode === 'through' || mode === 'inside' ? mode : 'auto',
      off: clamp(num('dimOffset', 12) * k, 6, 60), angle: num('dimAngle', 45, -1e9), lenEnds: st.lenEnds,
      font: App.previewFont ? App.previewFont(st.fontName) : { family: 'sans-serif', weight: 400, italic: false }
    };
  }

  // 示範的數字（跟 host.jsx 的 dimLabel 一樣的格式；文件單位還不知道時用 mm）
  function sampleLabel(mm, sym) {
    var u = $('dimUnit').value === 'doc' && !(info && info.doc) ? ['mm', UNIT_PT.mm] : unitOf();
    var v = mm * UNIT_PT.mm * (ratioOf($('dimRatio').value) || 1) / (u[1] || 1);
    return $('dimPrefix').value + (sym || '') + fmt(v, decimals()) + ($('dimShowUnit').checked ? u[0] : '') + $('dimSuffix').value;
  }

  // 文字的大小（半寬、半高）：用畫布量寬度；高度約是字級的 0.72（中文等再高一點）
  var meter = null;
  try { meter = document.createElement('canvas').getContext('2d'); } catch (e) {}
  function measure(label, s) {
    var tall = /[^\x00-\x7F\u00B1\u00D7\u00D8\u03C6\u2300]/.test(label), w = label.length * s.fs * 0.56;
    if (meter) {
      meter.font = (s.font.italic ? 'italic ' : '') + s.font.weight + ' ' + s.fs + 'px ' + s.font.family;
      w = meter.measureText(label).width;
    }
    return { label: label, hw: w / 2, hh: s.fs * (tall ? 0.44 : 0.36) };
  }

  function textAlong(tx, v, aligned) { return aligned ? tx.hw : tx.hw * Math.abs(v[0]) + tx.hh * Math.abs(v[1]); }
  function textCenter(mid, n, tx, aligned, gap) {
    return along(mid, n, gap + (aligned ? tx.hh : tx.hw * Math.abs(n[0]) + tx.hh * Math.abs(n[1])));
  }
  // 文字放在線中間：p → q（方向 a）在 mid 前後各空 half；兩邊剩不到 minLen 時不斷開（null）
  function splitPts(p, q, a, mid, half, minLen) {
    if (vdot(vsub(mid, p), a) - half < minLen || vdot(vsub(q, mid), a) - half < minLen) return null;
    return [[p, along(mid, a, -half)], [along(mid, a, half), q]];
  }

  function newFig() { return { line: [], end: [], ext: [], text: [], shape: null }; }

  // 寬度、高度和兩點距離（gm.free：距離 0、沒有延伸線）（host.jsx 的 buildLinear）
  function linDim(f, gm, s, label) {
    var a = gm.a, n = gm.n, L = s.L, kind = s.end, tr = kind === 'arrow' ? L * 0.9 : 0;
    var level = Math.max(vdot(gm.p1, n), vdot(gm.p2, n)) + (gm.free ? 0 : s.off), P = [gm.p1, gm.p2], Q = [], i;
    for (i = 0; i < 2; i++) {
      var dist = level - vdot(P[i], n);
      Q.push(along(P[i], n, dist));
      if (dist - s.extGap > 0.1) f.ext.push({ pts: [along(P[i], n, s.extGap), along(Q[i], n, s.extOver)], w: s.extLine === 'thin' ? s.lw / 2 : s.lw });
    }
    if (vdot(vsub(Q[1], Q[0]), a) < 0) Q.reverse();
    var Q1 = Q[0], Q2 = Q[1], len = vlen(vsub(Q2, Q1)), arrowLike = kind === 'arrow' || kind === 'open';
    var one = gm.free && s.lenEnds === 'one';   // 兩點距離、箭頭只畫起點 Q1
    var inside = one || (s.arrowPos === 'inside' ? true : s.arrowPos === 'outside' ? false : len >= L * 2.6), pts;
    if (one) pts = [arrowLike ? along(Q1, a, tr) : Q1, Q2];
    else if (arrowLike) pts = inside ? [along(Q1, a, tr), along(Q2, a, -tr)] : [along(Q1, a, -L * 2.2), along(Q2, a, L * 2.2)];
    else if (kind === 'tick') pts = [along(Q1, a, -L * 0.6), along(Q2, a, L * 0.6)];
    else pts = [Q1, Q2];
    var tx = measure(label, s), mid = [(Q1[0] + Q2[0]) / 2, (Q1[1] + Q2[1]) / 2], split = null;
    if (s.textPos === 'middle' && (inside || !arrowLike)) {
      split = splitPts(pts[0], pts[1], a, mid, textAlong(tx, a, s.aligned) + s.breakGap, Math.max(L, 2 * K));
    }
    f.line.push.apply(f.line, split || [pts]);
    if (one) f.end.push({ tip: Q1, u: kind === 'tick' ? a : [-a[0], -a[1]] });
    else if (kind === 'tick') f.end.push({ tip: Q1, u: a }, { tip: Q2, u: a });
    else {
      var sg = arrowLike && !inside ? 1 : -1;
      f.end.push({ tip: Q1, u: [a[0] * sg, a[1] * sg] }, { tip: Q2, u: [-a[0] * sg, -a[1] * sg] });
    }
    f.text.push({ tx: tx, c: split ? mid : textCenter(mid, n, tx, s.aligned, s.besideGap), ang: s.aligned ? readAngle(a) : 0 });
  }

  // 直徑（host.jsx 的 buildRadial）
  function diaDim(f, c, r, u, s, label) {
    var L = s.L, kind = s.end === 'tick' ? 'arrow' : s.end, gap = s.textGap, tr = kind === 'arrow' ? L * 0.9 : 0, back = [-u[0], -u[1]];
    var tx = measure(label, s), span = s.aligned ? tx.hw * 2 : 2 * (tx.hw * Math.abs(u[0]) + tx.hh * Math.abs(u[1]));
    var endRoom = kind === 'none' ? 0 : L * 1.2, M = along(c, u, r), M1 = along(c, u, -r), mode = s.radMode, tc, ang = 0;
    if (mode === 'auto') mode = r * 2 >= span + endRoom * 2 + gap * 4 ? 'inside' : 'leader';
    if (mode === 'inside') {
      ang = s.aligned ? readAngle(u) : 0;
      var p = along(M1, u, tr), q = along(M, u, -tr);
      var split = s.textPos === 'middle' ? splitPts(p, q, u, c, textAlong(tx, u, s.aligned) + s.breakGap, Math.max(L, 2 * K)) : null;
      f.line.push.apply(f.line, split || [[p, q]]);
      tc = split ? c : textCenter(c, dirOf(readAngle(u) + 90), tx, s.aligned, s.besideGap);
      f.end.push({ tip: M1, u: back, kind: kind }, { tip: M, u: u, kind: kind });
    } else {
      var Kp = along(M, u, Math.max(s.leader, L * 1.5, 2 * K)), sx = u[0] < -0.0001 ? -1 : 1, pts;
      if (mode === 'through') {
        pts = [along(M1, u, tr), Kp];
        f.end.push({ tip: M1, u: back, kind: kind }, { tip: M, u: u, kind: kind });
      } else {
        pts = [along(M, u, tr), Kp];
        f.end.push({ tip: M, u: back, kind: kind });
      }
      if (s.shelf === 'over') {
        pts.push([Kp[0] + sx * (tx.hw * 2 + gap * 2), Kp[1]]);
        tc = [Kp[0] + sx * (tx.hw + gap), Kp[1] + s.besideGap + tx.hh];
      } else {
        var S = Kp;
        if (s.shelf === 'end') { S = [Kp[0] + sx * Math.max(L * 1.5, s.fs * 0.6), Kp[1]]; pts.push(S); }
        tc = [S[0] + sx * (gap + tx.hw), S[1]];
      }
      f.line.push(pts);
    }
    f.text.push({ tx: tx, c: tc, ang: ang });
    return mode;
  }

  // 外框 [左, 下, 右, 上]
  function bounds(f, s) {
    var b = [Infinity, Infinity, -Infinity, -Infinity];
    var add = function (p) { b[0] = Math.min(b[0], p[0]); b[1] = Math.min(b[1], p[1]); b[2] = Math.max(b[2], p[0]); b[3] = Math.max(b[3], p[1]); };
    f.line.concat(f.ext.map(function (e) { return e.pts; })).forEach(function (pts) { pts.forEach(add); });
    f.end.forEach(function (e) { add([e.tip[0] - s.L, e.tip[1] - s.L]); add([e.tip[0] + s.L, e.tip[1] + s.L]); });
    f.text.forEach(function (t) {
      var a = t.ang * Math.PI / 180, cs = Math.abs(Math.cos(a)), sn = Math.abs(Math.sin(a));
      var w = t.tx.hw * cs + t.tx.hh * sn + 2, h = t.tx.hw * sn + t.tx.hh * cs + 2;
      add([t.c[0] - w, t.c[1] - h]); add([t.c[0] + w, t.c[1] + h]);
    });
    (f.shape.blocks || [f.shape]).forEach(function (sh) {
      if (sh.r) { add([sh.c[0] - sh.r, sh.c[1] - sh.r]); add([sh.c[0] + sh.r, sh.c[1] + sh.r]); }
      else { add([sh.x, sh.y]); add([sh.x + sh.w, sh.y + sh.h]); }
    });
    return b;
  }

  // 物件先畫大一點，量外框：放不下就縮小物件（shrink 算新的 size），最後移到框的中間（上面留 BAND）。
  // 物件縮到最小還是放不下（字、箭頭、距離很大）：整張圖等比縮小（f.zoom，dimFigure 畫的時候套用）
  function fitFigure(make, size, min, shrink, s) {
    var aw = FIG - 20, ah = FIG - BAND - 10, f, b, i;
    for (i = 0; ; i++) {
      f = make(size, 0, 0);
      b = bounds(f, s);
      var ox = b[2] - b[0] - aw, oy = b[3] - b[1] - ah;
      if ((ox <= 0.5 && oy <= 0.5) || size <= min || i === 5) break;
      size = Math.max(min, shrink(size, Math.max(0, ox), Math.max(0, oy)));
    }
    var bw = b[2] - b[0], bh = b[3] - b[1];
    if (bw > aw + 0.5 || bh > ah + 0.5) {
      f.zoom = { k: Math.min(aw / bw, ah / bh), c: [(b[0] + b[2]) / 2, (b[1] + b[3]) / 2] };
      return f;
    }
    return make(size, 10 + (aw - bw) / 2 - b[0], 10 + (ah - bh) / 2 - b[1]);
  }

  // 矩形：size = 寬，高 = 寬 × 0.72。寬高畫在四邊的圓點選的那幾邊；都沒選時畫上、右，淡淡的（f.faint，還是點得到）
  function rectFigure(s) {
    var on = { top: $('dimPosTop').checked, bottom: $('dimPosBottom').checked, left: $('dimPosLeft').checked, right: $('dimPosRight').checked };
    var faint = !on.top && !on.bottom && !on.left && !on.right;
    if (faint) on = { top: true, right: true };
    var fig = fitFigure(function (W, dx, dy) {
      var f = newFig(), H = W * 0.72, x1 = dx + W, y1 = dy + H, w = sampleLabel(SAMPLE.w), h = sampleLabel(SAMPLE.h);
      f.shape = { x: dx, y: dy, w: W, h: H };
      if (on.top) linDim(f, { p1: [dx, y1], p2: [x1, y1], a: [1, 0], n: [0, 1] }, s, w);
      if (on.bottom) linDim(f, { p1: [dx, dy], p2: [x1, dy], a: [1, 0], n: [0, -1] }, s, w);
      if (on.left) linDim(f, { p1: [dx, dy], p2: [dx, y1], a: [0, 1], n: [-1, 0] }, s, h);
      if (on.right) linDim(f, { p1: [x1, dy], p2: [x1, y1], a: [0, 1], n: [1, 0] }, s, h);
      return f;
    }, 220, 50, function (W, ox, oy) { return W - Math.max(ox, oy / 0.72) - 0.5; }, s);
    fig.faint = faint;
    return fig;
  }

  // 圓：size = 半徑
  function circleFigure(s) {
    var u = dirOf(s.angle);
    return fitFigure(function (r, dx, dy) {
      var f = newFig(), c = [dx, dy];
      f.shape = { c: c, r: r };
      f.mode = diaDim(f, c, r, u, s, sampleLabel(SAMPLE.d, diaSym()));
      return f;
    }, 100, 18, function (r, ox, oy) { return r - Math.max(ox, oy) * 0.55 - 0.5; }, s);
  }

  // 兩點距離：size = 整張圖的寬，兩塊物件（各寬 30%、高 45%）中間的空隙標一個箭頭
  function lenFigure(s) {
    return fitFigure(function (W, dx, dy) {
      var f = newFig(), bw = W * 0.3, H = W * 0.45, y = dy + H / 2;
      f.shape = { blocks: [{ x: dx, y: dy, w: bw, h: H }, { x: dx + W - bw, y: dy, w: bw, h: H }] };
      linDim(f, { p1: [dx + bw, y], p2: [dx + W - bw, y], a: [1, 0], n: [0, s.st.lenSide === 'b' ? -1 : 1], free: true }, s, sampleLabel(SAMPLE.len));
      return f;
    }, 220, 60, function (W, ox, oy) { return W - Math.max(ox, oy / 0.45) - 0.5; }, s);
  }

  function esc2(t) { return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
  function P(p) { return p[0].toFixed(2) + ',' + (FIG - p[1]).toFixed(2); }
  function X(v) { return v.toFixed(2); }

  // 端點（host.jsx 的 dimEnd）：tip 尖端，u 指向尖端的方向
  function endMark(e, s) {
    var L = s.L, t = e.tip, u = e.u, kind = e.kind || s.end, col = s.col, perp = [-u[1], u[0]];
    if (kind === 'none') return '<circle class="ghost" cx="' + X(t[0]) + '" cy="' + X(FIG - t[1]) + '" r="4"/>';
    if (kind === 'dot') return '<circle cx="' + X(t[0]) + '" cy="' + X(FIG - t[1]) + '" r="' + X(L * 0.3) + '" fill="' + col + '"/>';
    if (kind === 'tick') {
      var v = [(u[0] - u[1]) * Math.SQRT1_2, (u[0] + u[1]) * Math.SQRT1_2];
      return '<polyline points="' + P(along(t, v, -L / 2)) + ' ' + P(along(t, v, L / 2)) + '" stroke="' + col + '" stroke-width="' + X(s.lw * 2) + '" fill="none"/>';
    }
    var B = along(t, u, -L);
    if (kind === 'open') {
      var ow = L * 0.27;
      return '<polyline points="' + P(along(B, perp, ow)) + ' ' + P(t) + ' ' + P(along(B, perp, -ow)) + '" fill="none" stroke="' + col +
        '" stroke-width="' + X(s.lw) + '" stroke-linejoin="round"/>';
    }
    var hw = L * 0.18;
    return '<path d="M' + P(t) + 'L' + P(along(B, perp, hw)) + 'L' + P(along(B, perp, -hw)) + 'Z" fill="' + col + '"/>';
  }

  // 選到的線：沿著線的一條寬的淡色線（.hl.glow；寬度寫在 style 才蓋得過 css 的 .hl 線寬）
  function dimFigure() {
    var s = figStyle(), f = figMode === 'circle' ? circleFigure(s) : (figMode === 'len' ? lenFigure(s) : rectFigure(s)), out = [];
    drawn = f;
    drawnRad = f.mode || null;
    // 拖曳改哪個欄位：[欄位 id, 顯示的名稱, 單位]；端點「無」、不畫延伸線時不能拖
    var size = {
      text: ['dimFontSize', '字級', 'pt'],
      line: ['dimLineWidth', '線寬', 'pt'],
      end: s.end === 'none' ? null : ['dimEndSize', '端點大小', 'pt'],
      ext: s.extLine === 'none' ? null : ['dimExtGap', '延伸線留空', 'pt']
    };
    var group = function (name, body) {
      var z = size[name];
      return '<g data-part="' + name + '" tabindex="0" role="button" aria-label="' + partName(name) + '的設定"' +
        (z ? ' data-scrub-for="' + z[0] + '" data-scrub-label="' + z[1] + '" data-scrub-unit="' + z[2] + '"' : '') +
        (fc.marked(name) ? ' class="sel"' : '') + '>' + body + '</g>';
    };
    var poly = function (pts) { return pts.map(P).join(' '); };

    out.push((f.shape.blocks || [f.shape]).map(function (sh) {
      return sh.r ? '<circle class="prod" cx="' + X(sh.c[0]) + '" cy="' + X(FIG - sh.c[1]) + '" r="' + X(sh.r) + '"/>'
        : '<rect class="prod" x="' + X(sh.x) + '" y="' + X(FIG - sh.y - sh.h) + '" width="' + X(sh.w) + '" height="' + X(sh.h) + '" rx="' + X(Math.min(sh.w, sh.h) * 0.12) + '"/>';
    }).join(''));

    if (f.ext.length) {
      var ghost = s.extLine === 'none';
      out.push(group('ext',
        f.ext.map(function (e) { return '<polyline class="hl glow" points="' + poly(e.pts) + '" style="stroke-width:' + X(e.w + 7) + 'px"/>'; }).join('') +
        f.ext.map(function (e) {
          return ghost ? '<polyline class="ghost" points="' + poly(e.pts) + '"/>'
            : '<polyline points="' + poly(e.pts) + '" fill="none" stroke="' + s.col + '" stroke-width="' + X(e.w) + '"/>';
        }).join('') +
        f.ext.map(function (e) { return '<polyline class="hit" points="' + poly(e.pts) + '" stroke-width="12"/>'; }).join('')));
    }

    out.push(group('line',
      f.line.map(function (pts) { return '<polyline class="hl glow" points="' + poly(pts) + '" style="stroke-width:' + X(s.lw + 7) + 'px"/>'; }).join('') +
      f.line.map(function (pts) {
        return '<polyline points="' + poly(pts) + '" fill="none" stroke="' + s.col + '" stroke-width="' + X(s.lw) + '" stroke-miterlimit="10"/>';
      }).join('') +
      f.line.map(function (pts) { return '<polyline class="hit" points="' + poly(pts) + '" stroke-width="14"/>'; }).join('')));

    // 端點的框：圍住箭頭（圓點、斜線在尖端）；點擊範圍順著箭頭（斜線順著斜線），才不會蓋到旁邊的延伸線
    var er = Math.max(s.L * 0.6, 5) + 2;
    var kindOf = function (e) { return e.kind || s.end; };
    var centerOf = function (e) { var k = kindOf(e); return k === 'arrow' || k === 'open' ? along(e.tip, e.u, -s.L / 2) : e.tip; };
    var quad = function (c, u, hl, hw) {
      var n = [-u[1], u[0]];
      return '<polygon class="hit" points="' + [[hl, hw], [hl, -hw], [-hl, -hw], [-hl, hw]].map(function (k) {
        return P(along(along(c, u, k[0]), n, k[1]));
      }).join(' ') + '"/>';
    };
    var endHit = function (e) {
      var k = kindOf(e), c = centerOf(e);
      if (k === 'arrow' || k === 'open') return quad(c, e.u, s.L / 2 + 3, Math.max(s.L * 0.3, 5));
      if (k === 'tick') return quad(c, [(e.u[0] - e.u[1]) * Math.SQRT1_2, (e.u[0] + e.u[1]) * Math.SQRT1_2], s.L / 2 + 3, 5);
      var r = k === 'dot' ? Math.max(s.L * 0.3 + 3, 6) : 7;
      return quad(c, [1, 0], r, r);
    };
    out.push(group('end',
      f.end.map(function (e) { var c = centerOf(e); return '<circle class="hl" cx="' + X(c[0]) + '" cy="' + X(FIG - c[1]) + '" r="' + X(er) + '"/>'; }).join('') +
      f.end.map(function (e) { return endMark(e, s); }).join('') +
      f.end.map(endHit).join('')));

    var st = s.st, font = s.font, bgPad = Math.max(0, Math.min(s.fs * 0.2, Math.min(s.textGap, s.breakGap) * 0.75));
    out.push(group('text', f.text.map(function (t) {
      var cx = X(t.c[0]), cy = X(FIG - t.c[1]), rot = Math.abs(t.ang) > 0.01 ? ' transform="rotate(' + X(-t.ang) + ' ' + cx + ' ' + cy + ')"' : '';
      var box = function (pad) {
        var w = t.tx.hw + pad, h = t.tx.hh + pad;
        return 'x="' + X(t.c[0] - w) + '" y="' + X(FIG - t.c[1] - h) + '" width="' + X(w * 2) + '" height="' + X(h * 2) + '"' + rot;
      };
      return (st.textBg === 'box' ? '<rect ' + box(bgPad) + ' fill="' + st.textBgColor + '"/>' : '') +
        '<rect class="hl" ' + box(2) + ' rx="2"/>' +
        '<text x="' + cx + '" y="' + cy + '" fill="' + st.textColor + '" font-size="' + X(s.fs) + '" font-family="' + esc(font.family) +
        '" font-weight="' + font.weight + '" font-style="' + (font.italic ? 'italic' : 'normal') +
        '" text-anchor="middle" dominant-baseline="central"' + rot + '>' + esc2(t.tx.label) + '</text>' +
        '<rect class="hit" ' + box(4) + '/>';
    }).join('')));
    // 提示在左上角的切換和右上角的「重設」中間，不跟著縮放
    var hint = fc.hint() ? '<text class="fig-hint" x="206" y="20" text-anchor="end">點部位調整，拖曳改大小</text>' : '';
    if (f.faint) out = [out[0], '<g opacity=".35">' + out.slice(1).join('') + '</g>'];
    if (!f.zoom) return hint + out.join('');
    var z = f.zoom;
    return hint + '<g transform="translate(' + X(ZOOM_C[0]) + ' ' + X(ZOOM_C[1]) + ') scale(' + z.k.toFixed(4) + ') translate(' + X(-z.c[0]) + ' ' + X(-(FIG - z.c[1])) + ')">' +
      out.join('') + '</g>';
  }

  // 矩形四邊的圓點（選寬高標在哪一邊）放在畫出來的矩形每一邊的中間；標圓時不顯示
  function placeDots() {
    var pad = $('dimPad'), sh = drawn && drawn.shape, z = drawn && drawn.zoom;
    pad.hidden = figMode !== 'rect' || !sh;
    if (pad.hidden) return;
    var mx = sh.x + sh.w / 2, my = sh.y + sh.h / 2;
    [['dimPosTop', mx, sh.y + sh.h], ['dimPosBottom', mx, sh.y], ['dimPosLeft', sh.x, my], ['dimPosRight', sh.x + sh.w, my]].forEach(function (d) {
      var x = d[1], y = FIG - d[2];   // 圖的座標（y 往上）→ 示範框的座標，有縮放時一樣縮放
      if (z) { x = ZOOM_C[0] + z.k * (x - z.c[0]); y = ZOOM_C[1] + z.k * (y - (FIG - z.c[1])); }
      var dot = pad.querySelector('label[for="' + d[0] + '"]');
      dot.style.left = (x / FIG * 100) + '%';
      dot.style.top = (y / FIG * 100) + '%';
    });
  }

  function renderFigure() {
    $('dimPreview').innerHTML = dimFigure();
    placeDots();
    fc.update();
  }

  var FIG_MODES = { rect: 'dimFigRect', circle: 'dimFigCircle', len: 'dimFigLen' };
  function showFigMode() {
    Object.keys(FIG_MODES).forEach(function (m) { $(FIG_MODES[m]).checked = figMode === m; });
    $('dimPop').setAttribute('data-mode', figMode);
  }
  function setFigMode(m) {
    figMode = FIG_MODES[m] ? m : 'rect';
    showFigMode();
    try { localStorage.setItem(FIG_KEY, figMode); } catch (e) {}
    renderFigure();
  }
  ['dimFigRect', 'dimFigCircle', 'dimFigLen'].forEach(function (id) {
    $(id).addEventListener('change', function () { if ($(id).checked) setFigMode($(id).value); });
  });

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
    renderFigure();
  }

  FIELDS.forEach(function (id) {
    var isStyle = STYLE.indexOf(id) >= 0 || id === 'dimOffset';   // 距離也自動套用到選取的尺寸
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

  $('btnDimAdd').addEventListener('click', function () { if (lenOn) stopLen(false, false); addAll(); });
  $('btnDimLen').addEventListener('click', function () { if (lenOn) stopLen(false); else startLen(); });
  App.on('pen', function (who) { if (who !== 'dim') stopLen(true, false); });   // 開始新增標註
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape' && lenOn && document.activeElement.tagName !== 'INPUT') stopLen(false);
  });
  $('btnDimSync').addEventListener('click', function () { restyle('dimSync'); });
  $('btnDimRestyle').addEventListener('click', function () { restyle('dimRestyle'); });
  $('btnDimPickStyle').addEventListener('click', pickStyle);
  App.on('pickedDim', applyPicked);   // 在標註分頁按「吸取樣式」、選的是尺寸
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

  // 示範框右上角的「重設」：這些大小改回預設值（index.html 寫的值），顏色、字體、數字格式、標法不動。
  // 跟改其他樣式一樣：存起來，勾了「修改時自動套用到選取的尺寸」就套用到選取的尺寸
  // 預設值依目前工作區的大小等比放大縮小（App.artboardFactor：A4 = 預設值；自動調整的基準也一起）
  function resetSizes() {
    App.artboardFactor(function (k, note) {
      SIZE_FIELDS.forEach(function (id) { var el = $(id); el.value = App.scaledDefault(el, k); });
      changed();
      scheduleApply(true);
      App.setStatus('大小改回預設' + (note ? '（' + note + '）' : '') + '：字級 ' + $('dimFontSize').value + '、線寬 ' + $('dimLineWidth').value +
        '、端點 ' + $('dimEndSize').value + '、距離 ' + $('dimOffset').value + ' pt');
    });
  }
  $('btnDimReset').addEventListener('click', resetSizes);
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
  try { figMode = localStorage.getItem(FIG_KEY) || 'rect'; } catch (e) {}
  if (!FIG_MODES[figMode]) figMode = 'rect';
  showFigMode();
  if (App.fontList) setFonts(App.fontList);
  changed();
  refresh();
})();
