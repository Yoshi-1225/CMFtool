/* CMF Tool — 標註分頁 */
(function () {
  var App = window.CMFApp;
  var cep = window.__adobe_cep__;
  var KEY = "cmfCalloutSettings";
  var STYLE_FIELDS = ["fontSize", "textColor", "fontName", "badge", "badgeColor", "badgeTextColor", "badgePadding",
    "badgeStrokeWidth", "badgeStrokeColor", "lineWidth", "lineColor", "endStyle", "endSize", "styleName", "styleScale", "gap",
    "chipSize", "chipBorder", "chipBorderColor"];
  // kind：新增 number 編號、swatch 色號（色塊 swColor + 文字 swText；swSpot = 滴管取到的色票名稱）、
  //   note 文字說明（noteText，可以好幾行）。三種各自有一套樣式（kindStyles），欄位裡是目前那一種的
  var FIELDS = ["lineMode", "continuous", "startMode", "startNumber", "sortMode", "sameNumber", "autoSync",
    "kind", "swColor", "swText", "swSpot", "swNoLine", "noteText", "noteArrowOnly"].concat(STYLE_FIELDS);
  var KIND_LABEL = { number: "編號", swatch: "色號", note: "說明" };
  var KIND_ADD = { number: "新增標註", swatch: "新增色號", note: "新增說明" };
  var MODE_TIPS = {
    straight: "點目標點 → 點{label}位置",
    elbow: "點目標點 → 點轉折點 → 點{label}位置",
    free: "左鍵持續加點，按「完成這條線」或快捷鍵完成"
  };
  function kindOf() { var k = $("kind").value; return KIND_LABEL[k] ? k : "number"; }
  // 文字說明「只畫箭頭」：沒有文字
  function arrowOnly() { return kindOf() === "note" && $("noteArrowOnly").checked; }
  function addLabel() { return arrowOnly() ? "新增箭頭" : KIND_ADD[kindOf()]; }
  // 色號「沒有箭頭」：點一下就放
  function noLine() { return kindOf() === "swatch" && $("swNoLine").checked; }
  function modeTip() {
    if (noLine()) return "點一下要放色號的位置（色塊的左邊）";
    return MODE_TIPS[$("lineMode").value].replace("{label}", arrowOnly() ? "箭頭尾端" : KIND_LABEL[kindOf()]);
  }
  var $ = function (id) { return document.getElementById(id); };

  // ---------- 與 ExtendScript 溝通 ----------
  function extensionPath() {
    if (!cep) return "";
    var p = decodeURI(cep.getSystemPath("extension"));
    return /^file:\/\/\/[A-Za-z]:/.test(p) ? p.replace("file:///", "") : p.replace("file://", "");
  }

  // JSON 物件 → ExtendScript 字串常值（U+2028/2029 在 ExtendScript 字串裡不合法）
  function q(obj) {
    return JSON.stringify(JSON.stringify(obj)).replace(/\u2028/g, "\\u2028").replace(/\u2029/g, "\\u2029");
  }

  function run(expr, cb, quiet) {
    if (!cep) { setStatus("請在 Illustrator 中開啟此面板", true); return; }
    cep.evalScript(expr, function (raw) {
      var r;
      try { r = JSON.parse(raw); } catch (e) { r = { ok: false, msg: "腳本錯誤：" + raw }; }
      if (!quiet && r.msg) setStatus(r.msg, !r.ok);
      if (cb) cb(r);
    });
  }

  function callHost(fn, obj, cb, quiet) {
    run("CMF." + fn + "(" + (obj ? q(obj) : "") + ")", cb, quiet);
  }

  function setStatus(msg, isError) { App.setStatus(msg, isError); }

  // ---------- 設定 ----------
  function readSettings() {
    var o = {};
    FIELDS.forEach(function (k) {
      var el = $(k);
      if (el.type === "checkbox") o[k] = el.checked;
      else if (App.isNum(el)) o[k] = parseFloat(el.value) || 0;
      else o[k] = el.value;
    });
    if (o.kind === "note") o.noteEnd = o.endStyle;   // host.jsx 的文字說明用 noteEnd 當端點
    o.hostPath = extensionPath() + "/jsx/host.jsx";
    return o;
  }

  // 編號、色號、文字說明各自的樣式：{ number: {...}, swatch: {...}, note: {...} }。
  // 欄位裡放的是目前那一種的；切換種類時先存起來，再換成另一種的（第一次切到的種類從目前的樣式複製，文字說明的端點改成箭頭）。
  // 「同步全部」、「套用選取」送出目前那一種的樣式，host.jsx 只改同一種的標註
  var STYLES_KEY = "cmfCalloutKindStyles";
  var kindStyles = {};
  try { kindStyles = JSON.parse(localStorage.getItem(STYLES_KEY) || "{}") || {}; } catch (e) { kindStyles = {}; }

  function storeKindStyle() {
    var o = readSettings(), st = {};
    STYLE_FIELDS.forEach(function (k) { st[k] = o[k]; });
    kindStyles[kindOf()] = st;
    try { localStorage.setItem(STYLES_KEY, JSON.stringify(kindStyles)); } catch (e) {}
  }

  function switchKind(k) {
    if (!KIND_LABEL[k] || k === kindOf()) return;
    storeKindStyle();
    $("kind").value = k;
    if (kindStyles[k]) applySettings(kindStyles[k]);
    else if (k === "note") $("endStyle").value = "arrow";
    syncFontSelects();
    renderPreview();
    saveSettings();
  }

  function applySettings(o) {
    FIELDS.forEach(function (k) {
      if (!o || o[k] === undefined) return;
      var el = $(k);
      if (el.type === "checkbox") el.checked = !!o[k]; else el.value = o[k];
    });
  }

  // 面板自己的設定沒有時（第一次開啟），沿用快捷鍵腳本用的設定檔
  function loadSettings(done) {
    var saved = null;
    try { saved = localStorage.getItem(KEY); } catch (e) {}
    if (saved || !cep) {
      try { applySettings(JSON.parse(saved || "{}")); } catch (e) {}
      migrateDefaults();
      done();
      return;
    }
    run("CMF.loadSettings()", function (r) {
      if (r && r.ok && r.settings) applySettings(r.settings);
      migrateDefaults();
      done();
    }, true);
  }

  // 預設值改過一次（線、箭頭深灰；編號有圓圈：白底、深灰框、黑字、內距 3；色塊無框；色號文字 Pantone）。
  // 存著的設定還是舊的預設值的，換成新的；自己改過的不動。只做一次
  var DEFAULTS_KEY = "cmftool:calloutDefaults2";
  var OLD_DEFAULTS = { lineColor: "#2b7bd6", chipBorder: 0.25, swText: "" };
  var NEW_DEFAULTS = { lineColor: "#404040", chipBorder: 0, swText: "Pantone" };
  var OLD_BADGE = { badge: false, badgeColor: "#2b7bd6", badgeTextColor: "#ffffff", badgePadding: 2, badgeStrokeWidth: 0, badgeStrokeColor: "#000000" };
  var NEW_BADGE = { badge: true, badgeColor: "#ffffff", badgeTextColor: "#000000", badgePadding: 3, badgeStrokeWidth: 0.5, badgeStrokeColor: "#404040" };
  function same(a, b) {
    if (typeof b === "boolean") return !!a === b;
    if (typeof b === "number") return Math.abs(parseFloat(a) - b) < 1e-6;
    return String(a === undefined || a === null ? "" : a).toLowerCase() === b;
  }
  function migrateDefaults() {
    try { if (localStorage.getItem(DEFAULTS_KEY)) return; localStorage.setItem(DEFAULTS_KEY, "1"); } catch (e) { return; }
    // get(k) / put(k, v)：一份設定（面板的欄位，或 kindStyles 裡存的另外兩種）
    var fix = function (get, put) {
      Object.keys(OLD_DEFAULTS).forEach(function (k) { if (get(k) !== undefined && same(get(k), OLD_DEFAULTS[k])) put(k, NEW_DEFAULTS[k]); });
      // 圓圈：還是舊的預設（沒有圓圈底、藍色）才整組換成新的
      if (Object.keys(OLD_BADGE).every(function (k) { return get(k) === undefined || same(get(k), OLD_BADGE[k]); })) {
        Object.keys(NEW_BADGE).forEach(function (k) { put(k, NEW_BADGE[k]); });
      }
    };
    fix(function (k) { var el = $(k); return el.type === "checkbox" ? el.checked : el.value; },
      function (k, v) { var el = $(k); if (el.type === "checkbox") el.checked = !!v; else el.value = v; });
    Object.keys(kindStyles).forEach(function (kind) {
      var st = kindStyles[kind];
      fix(function (k) { return st[k]; }, function (k, v) { if (k !== "swText") st[k] = v; });
    });
    try { localStorage.setItem(STYLES_KEY, JSON.stringify(kindStyles)); } catch (e) {}
  }

  function saveLocal() {
    try { localStorage.setItem(KEY, JSON.stringify(readSettings())); } catch (e) {}
    storeKindStyle();
  }

  var saveTimer = null;
  function saveSettings() {
    var o = readSettings();
    storeKindStyle();
    try { localStorage.setItem(KEY, JSON.stringify(o)); } catch (e) {}
    clearTimeout(saveTimer);
    saveTimer = setTimeout(function () { if (cep) callHost("saveSettings", o, null, true); }, 400);
  }

  // ---------- 標註的圖：點部位，打開它的設定 ----------
  // 樣式區塊裡畫一個標註（照目前的樣式），編號、圓圈、引線、端點都可以點：打開那個部位的設定卡片，
  // 卡片放在圖的下面，不會蓋住圖（js/figcard.js）。
  // 圖上的大小 = pt × 2；太小、太大的數值畫成看得到、放得下的大小。每個部位有比畫出來還大的點擊範圍，
  // 沒有圓圈底、端點是「無」時，滑過圖會出現虛線，一樣點得到。
  // 在部位上按住左右拖曳，直接改它的大小（js/numfield.js 的拖曳把手）：編號 = 字級、圓圈 = 內距、引線 = 線寬、
  // 端點 = 大小（繪圖樣式 = %）。第一次在圖上點過之後，角落的提示就不再出現
  var PART_NAMES = { text: "文字", badge: "圓圈", chip: "色塊", line: "引線", end: "端點" };
  // 每個部位的設定：圖上選著一個部位時，「同步全部」、「套用選取」、⚡ 自動套用、「吸取樣式」只處理這些
  var PART_FIELDS = {
    text: ["fontSize", "textColor", "fontName", "badgeTextColor"],
    badge: ["badge", "badgeColor", "badgePadding", "badgeStrokeWidth", "badgeStrokeColor"],
    chip: ["chipSize", "chipBorder", "chipBorderColor"],
    line: ["lineWidth", "lineColor", "gap"],
    end: ["endStyle", "endSize", "styleName", "styleScale", "noteEnd"]
  };
  var fc = App.figCard({ svg: $("preview"), card: $("stylePop"), title: $("popTitle"), names: PART_NAMES, first: "text",
    key: "cmftool:calloutPart", hintKey: "cmftool:partHint" });

  // 顏色欄位外面包著滴管按鈕的框（見「取色」）
  function boxOf(id) {
    var el = $(id);
    return el.parentNode && el.parentNode.classList.contains("colorbox") ? el.parentNode : el;
  }

  function clamp(v, lo, hi) { return isFinite(v) ? Math.max(lo, Math.min(hi, v)) : lo; }

  function renderPreview() {
    var o = readSettings(), kind = kindOf(), sw = kind !== "number", note = kind === "note";
    $("badgeRows").classList.toggle("hidden", !o.badge);
    $("styleRow").classList.toggle("hidden", o.endStyle !== "style");
    $("startNumber").disabled = o.startMode === "auto";
    $("badgeStrokeColor").disabled = !(o.badgeStrokeWidth > 0);
    $("chipBorderColor").disabled = !(o.chipBorder > 0);
    // 編號的顏色：有圓圈底時是圓圈裡的數字顏色；色號的文字一律用文字顏色
    boxOf("textColor").classList.toggle("hidden", !!o.badge && !sw);
    boxOf("badgeTextColor").classList.toggle("hidden", !o.badge || sw);
    $("kindNumber").checked = kind === "number";
    $("kindSwatch").checked = kind === "swatch";
    $("kindNote").checked = note;
    $("swRow").hidden = kind !== "swatch";
    $("noteRow").hidden = !note;
    $("numSect").hidden = sw;   // 色號、文字說明沒有編號，不顯示編號表
    if (!adding) $("btnAdd").textContent = addLabel();
    $("noteText").disabled = arrowOnly();
    updateFinishButton();
    $("preview").innerHTML = figure(o);
    fc.update();
  }

  function figure(o) {
    var K = 2, out = [];
    // 照實際的比例畫，但整張圖自動縮放到合適的大小：編號（圓圈）、色號、文字說明縮放到差不多一樣大，
    // 字、圓圈、色塊、引線、端點、間距都用同一個倍率 z，比例才跟實際標出來的一樣。
    // 端點、線寬、間距放大後太大時 z 再縮小一點；太小的才各自放大到看得到
    var num = function (v) { v = Number(v); return isFinite(v) && v > 0 ? v * K : 0; };
    var fs0 = num(o.fontSize) || 9, es0 = num(o.endSize), lw0 = num(o.lineWidth) || 0.8, gap0 = num(o.gap);
    var pad0 = num(o.badgePadding), bsw0 = num(o.badgeStrokeWidth), cs0 = num(o.chipSize) || fs0;
    var R0 = fs0 * 0.36 + pad0 + bsw0 / 2;   // 圓圈的外緣（半徑）
    var badgeFig = o.badge && o.kind !== "swatch" && o.kind !== "note";
    var z = badgeFig ? 20 / R0 : (o.kind === "swatch" ? 24 / Math.max(fs0, cs0) : 26 / fs0);
    z = Math.min(z, es0 ? 40 / (es0 * 1.6) : z, 8 / lw0, gap0 ? 24 / gap0 : z);
    var lw = Math.max(0.8, lw0 * z), es = Math.max(3, es0 * z), fs = Math.max(9, fs0 * z), gap = gap0 * z, pad = pad0 * z;
    var bsw = bsw0 ? Math.max(0.75, bsw0 * z) : 0;
    // 正方形的圖：編號端在右上 → (轉折) → 目標點（落在左下角的產品上）；最大的圓圈也放得下
    // 色號、文字說明：標籤比較寬，編號端放左邊一點
    var sw = o.kind === "swatch" || o.kind === "note", bare = o.kind === "swatch" && o.swNoLine;   // bare：沒有箭頭的色號
    var pts = bare ? [[40, 112], [39, 112]] : sw ? (o.lineMode === "straight" ? [[118, 76], [100, 182]] : [[96, 58], [64, 58], [100, 182]])
      : (o.lineMode === "straight" ? [[176, 84], [100, 182]] : [[156, 64], [118, 64], [100, 182]]);
    var S = pts[0], S1 = pts[1], E = pts[pts.length - 1], E1 = pts[pts.length - 2];
    // 拖曳改哪個欄位：[欄位 id, 顯示的名稱, 單位]；沒有圓圈底、端點是「無」時不能拖
    var size = {
      text: ["fontSize", "字級", "pt"],
      badge: o.badge ? ["badgePadding", "圓圈內距", "pt"] : null,
      chip: ["chipSize", "色塊大小", "pt"],
      line: ["lineWidth", "線寬", "pt"],
      end: o.endStyle === "style" ? ["styleScale", "繪圖樣式", "%"] : o.endStyle === "none" ? null : ["endSize", "端點大小", "pt"]
    };
    var group = function (name, body) {
      var z = size[name];
      return '<g data-part="' + name + '" tabindex="0" role="button" aria-label="' + PART_NAMES[name] + '的設定"' +
        (z ? ' data-scrub-for="' + z[0] + '" data-scrub-label="' + z[1] + '" data-scrub-unit="' + z[2] + '"' : "") +
        (fc.marked(name) ? ' class="sel"' : "") + ">" + body + "</g>";
    };
    var xy = function (p) { return p[0].toFixed(2) + "," + p[1].toFixed(2); };

    out.push('<path class="prod" d="M0 150 C 60 152, 120 180, 150 240 L0 240 Z"/>');
    if (fc.hint()) out.push('<text class="fig-hint" x="8" y="15">點一下調整設定，左右拖曳改大小</text>');

    // 引線（箭頭時線縮到箭頭底部，跟 host.jsx 一樣）
    var ux = E[0] - E1[0], uy = E[1] - E1[1], L = Math.sqrt(ux * ux + uy * uy) || 1;
    ux /= L; uy /= L;
    var tip = o.endStyle === "arrow" ? [E[0] - ux * es * 1.6, E[1] - uy * es * 1.6] : E;
    var full = pts.map(xy).join(" "), drawn = pts.slice(0, -1).concat([tip]).map(xy).join(" ");
    if (!bare) out.push(group("line",
      '<polyline class="hl glow" points="' + full + '" style="stroke-width:' + (lw + 7) + 'px"/>' +   // style：蓋過 css 的 .hl 線寬
      '<polyline points="' + drawn + '" fill="none" stroke="' + o.lineColor + '" stroke-width="' + lw + '" stroke-linejoin="round"/>' +
      '<polyline class="hit" points="' + full + '" stroke-width="14"/>'));

    // 端點
    var mark, len = es * 1.6, B = [E[0] - ux * len, E[1] - uy * len];
    if (o.endStyle === "arrow") {
      var hw = es * 0.5;
      mark = '<path d="M' + xy(E) + "L" + xy([B[0] - uy * hw, B[1] + ux * hw]) + "L" + xy([B[0] + uy * hw, B[1] - ux * hw]) + 'Z" fill="' + o.lineColor + '"/>';
    } else if (o.endStyle === "dot") {
      mark = '<circle cx="' + E[0] + '" cy="' + E[1] + '" r="' + es / 2 + '" fill="' + o.lineColor + '"/>';
    } else if (o.endStyle === "style") {
      // 繪圖樣式：實際的樣子看不到，用開放箭頭代表
      var ow = es * 0.6;
      mark = '<polyline points="' + xy([B[0] - uy * ow, B[1] + ux * ow]) + " " + xy(E) + " " + xy([B[0] + uy * ow, B[1] - ux * ow]) +
        '" fill="none" stroke="' + o.lineColor + '" stroke-width="' + lw + '" stroke-linejoin="round" stroke-linecap="round"/>';
    } else {
      mark = '<circle class="ghost" cx="' + E[0] + '" cy="' + E[1] + '" r="4"/>';
    }
    var er = Math.max(es * 1.2, 7) + 3;
    if (!bare) out.push(group("end", '<circle class="hl" cx="' + E[0] + '" cy="' + E[1] + '" r="' + er + '"/>' + mark +
      '<circle class="hit" cx="' + E[0] + '" cy="' + E[1] + '" r="' + Math.max(er, 12) + '"/>'));

    // 編號：沿線的第一段往外推（跟 host.jsx 的 buildCallout 一樣）
    var dx = S[0] - S1[0], dy = S[1] - S1[1], dl = Math.sqrt(dx * dx + dy * dy) || 1;
    dx /= dl; dy /= dl;
    if (sw) { out.push((o.kind === "note" ? noteFigure : swatchFigure)(o, S, [dx, dy], fs, gap, group, z)); return out.join(""); }
    var tw = fs * 0.3, th = fs * 0.36, R = Math.max(tw, th) + pad, C;
    if (o.badge) {
      C = [S[0] + dx * (gap + R + bsw / 2), S[1] + dy * (gap + R + bsw / 2)];
    } else {
      var ext = Math.min(Math.abs(dx) > 0.001 ? tw / Math.abs(dx) : 1e9, Math.abs(dy) > 0.001 ? th / Math.abs(dy) : 1e9);
      C = [S[0] + dx * (gap + ext), S[1] + dy * (gap + ext)];
    }
    var cx = C[0].toFixed(2), cy = C[1].toFixed(2);
    var rr = o.badge ? R + bsw / 2 : Math.max(tw, th) + 5;   // 圓圈（沒有圓圈底時是虛線圈）的外緣
    out.push(group("badge", '<circle class="hl" cx="' + cx + '" cy="' + cy + '" r="' + (rr + 3) + '"/>' +
      (o.badge
        ? '<circle cx="' + cx + '" cy="' + cy + '" r="' + R + '" fill="' + o.badgeColor + '"' +
          (bsw ? ' stroke="' + o.badgeStrokeColor + '" stroke-width="' + bsw + '"' : "") + "/>"
        : '<circle class="ghost" cx="' + cx + '" cy="' + cy + '" r="' + rr + '"/>') +
      '<circle class="hit" cx="' + cx + '" cy="' + cy + '" r="' + (rr + 6) + '"/>'));

    var pf = previewFont(o);
    var font = pf.family + '" font-weight="' + pf.weight + '" font-style="' + (pf.italic ? "italic" : "normal");
    var box = 'x="' + (C[0] - tw - 2).toFixed(2) + '" y="' + (C[1] - th - 2).toFixed(2) + '" width="' + (tw * 2 + 4).toFixed(2) +
      '" height="' + (th * 2 + 4).toFixed(2) + '"';
    out.push(group("text", '<rect class="hl" ' + box + ' rx="2"/>' +
      '<text x="' + cx + '" y="' + cy + '" fill="' + (o.badge ? o.badgeTextColor : o.textColor) + '" font-size="' + fs +
      '" font-family="' + font + '" text-anchor="middle" dominant-baseline="central">3</text>' +
      '<rect class="hit" ' + box + "/>"));
    return out.join("");
  }

  // 色號的標籤：[色塊][文字]（host.jsx 的 swatchLabel）。放不進圖裡時字和色塊一起縮小
  var meter = null;
  try { meter = document.createElement("canvas").getContext("2d"); } catch (e) {}
  function textWidth(text, fs, pf) {
    if (!meter) return text.length * fs * 0.6;
    meter.font = (pf.italic ? "italic " : "") + pf.weight + " " + fs + "px " + pf.family;
    return meter.measureText(text).width;
  }

  function swLabel(o) { return String(o.swText || "").trim() || o.swSpot || String(o.swColor).toUpperCase(); }

  // z：圖的倍率（figure 算的，跟字、線一起縮小）
  function swatchFigure(o, S, dir, fs, gap, group, z) {
    var K = 2 * z, pf = previewFont(o), label = swLabel(o);
    var cs = Math.max(6, (Number(o.chipSize) || Number(o.fontSize)) * K), bw = o.chipBorder > 0 ? Math.max(0.75, o.chipBorder * K) : 0;
    var tw = textWidth(label, fs, pf), outer = cs + bw, cg = Math.max(cs * 0.4, fs * 0.3);
    // 能用的寬度：直線（文字在線頭上方）左右對稱；折線往右
    var room = Math.abs(dir[0]) > 0.7 ? 236 - S[0] - gap : 2 * Math.min(S[0], 240 - S[0]) - 8;
    var k = Math.min(1, room / (outer + cg + tw));
    fs *= k; cs *= k; bw *= k; tw *= k; outer *= k; cg *= k;
    var hw = (outer + cg + tw) / 2, hh = Math.max(outer, fs * 0.72) / 2;
    var ext = Math.min(Math.abs(dir[0]) > 0.001 ? hw / Math.abs(dir[0]) : 1e9, Math.abs(dir[1]) > 0.001 ? hh / Math.abs(dir[1]) : 1e9);
    var C = [S[0] + dir[0] * (gap + ext), S[1] + dir[1] * (gap + ext)], left = C[0] - hw;
    var f = function (v) { return v.toFixed(2); };
    var cx = left + bw / 2, cy = C[1] - cs / 2, tx = left + outer + cg, th = fs * 0.36;
    var chipBox = 'x="' + f(left - 2) + '" y="' + f(C[1] - outer / 2 - 2) + '" width="' + f(outer + 4) + '" height="' + f(outer + 4) + '"';
    var textBox = 'x="' + f(tx - 2) + '" y="' + f(C[1] - th - 2) + '" width="' + f(tw + 4) + '" height="' + f(th * 2 + 4) + '"';
    return group("chip", '<rect class="hl" ' + chipBox + ' rx="2"/>' +
        '<rect x="' + f(cx) + '" y="' + f(cy) + '" width="' + f(cs) + '" height="' + f(cs) + '" fill="' + o.swColor + '"' +
        (bw ? ' stroke="' + o.chipBorderColor + '" stroke-width="' + f(bw) + '"' : "") + "/>" +
        '<rect class="hit" ' + chipBox + "/>") +
      group("text", '<rect class="hl" ' + textBox + ' rx="2"/>' +
        '<text x="' + f(tx) + '" y="' + f(C[1]) + '" fill="' + o.textColor + '" font-size="' + f(fs) + '" font-family="' + pf.family +
        '" font-weight="' + pf.weight + '" font-style="' + (pf.italic ? "italic" : "normal") + '" dominant-baseline="central">' +
        label.replace(/&/g, "&amp;").replace(/</g, "&lt;") + "</text>" +
        '<rect class="hit" ' + textBox + "/>");
  }

  // 文字說明：一段文字（好幾行），往右靠左、往左靠右、上下置中（host.jsx 的 swatchLabel）
  function noteFigure(o, S, dir, fs, gap, group) {
    if (o.noteArrowOnly) return "";   // 只畫箭頭：沒有文字
    var pf = previewFont(o), lines = (String(o.noteText || "").replace(/\r/g, "").trim() || "說明").split("\n");
    var widths = lines.map(function (l) { return textWidth(l, fs, pf); });
    var tw = Math.max.apply(null, widths), lh = fs * 1.2;
    var room = Math.abs(dir[0]) > 0.7 ? 236 - S[0] - gap : 2 * Math.min(S[0], 240 - S[0]) - 8;
    var k = Math.min(1, room / tw, 150 / (lines.length * lh));
    fs *= k; tw *= k; lh *= k;
    var hw = tw / 2, hh = lines.length * lh / 2;
    var ext = Math.min(Math.abs(dir[0]) > 0.001 ? hw / Math.abs(dir[0]) : 1e9, Math.abs(dir[1]) > 0.001 ? hh / Math.abs(dir[1]) : 1e9);
    var C = [S[0] + dir[0] * (gap + ext), S[1] + dir[1] * (gap + ext)];
    var anchor = Math.abs(dir[0]) < 0.3 ? "middle" : dir[0] > 0 ? "start" : "end";
    var x = anchor === "middle" ? C[0] : anchor === "start" ? C[0] - hw : C[0] + hw;
    var f = function (v) { return v.toFixed(2); };
    var box = 'x="' + f(C[0] - hw - 2) + '" y="' + f(C[1] - hh - 2) + '" width="' + f(tw + 4) + '" height="' + f(hh * 2 + 4) + '"';
    return group("text", '<rect class="hl" ' + box + ' rx="2"/>' +
      '<text fill="' + o.textColor + '" font-size="' + f(fs) + '" font-family="' + pf.family + '" font-weight="' + pf.weight +
      '" font-style="' + (pf.italic ? "italic" : "normal") + '" text-anchor="' + anchor + '">' +
      lines.map(function (l, i) {
        return '<tspan x="' + f(x) + '" y="' + f(C[1] - hh + lh * (i + 0.5)) + '" dominant-baseline="central">' +
          l.replace(/&/g, "&amp;").replace(/</g, "&lt;") + "</tspan>";
      }).join("") + "</text>" +
      '<rect class="hit" ' + box + "/>");
  }

  // ---------- 字體選單 ----------
  // fonts：{ 家族: [[字重, PostScript 名稱], ...] }，快取在 localStorage，避免每次開面板都掃描
  var FONT_KEY = "cmfCalloutFonts";
  var fonts = {}, fontIndex = {}; // fontIndex：PostScript 名稱 → [家族, 字重]

  function setFontData(list) {
    fonts = {}; fontIndex = {};
    list.forEach(function (f) {
      if (!fonts[f[0]]) fonts[f[0]] = [];
      fonts[f[0]].push([f[1], f[2]]);
      fontIndex[f[2]] = [f[0], f[1]];
    });
    var fam = $("fontFamily");
    var families = Object.keys(fonts).sort(function (a, b) { return a.localeCompare(b); });
    var html = '<option value="">預設字體</option>';
    families.forEach(function (name) {
      html += '<option value="' + name.replace(/"/g, "&quot;") + '">' + name.replace(/</g, "&lt;") + '</option>';
    });
    fam.innerHTML = html;
    syncFontSelects();
    App.fontList = list;      // 尺寸分頁共用字體清單
    App.emit("fonts", list);
  }

  // 依目前的 PostScript 名稱，把兩個選單切到對應位置
  function syncFontSelects() {
    var hit = fontIndex[$("fontName").value];
    $("fontFamily").value = hit ? hit[0] : "";
    fillStyles(hit ? hit[1] : null);
  }

  function fillStyles(selected) {
    var fam = $("fontFamily").value, st = $("fontStyle");
    var styles = fonts[fam] || [];
    st.disabled = !styles.length;
    st.innerHTML = styles.map(function (s) {
      return '<option value="' + s[1].replace(/"/g, "&quot;") + '">' + s[0].replace(/</g, "&lt;") + '</option>';
    }).join("");
    if (!styles.length) return;
    var pick = null;
    styles.forEach(function (s) { if (s[0] === selected) pick = s[1]; });
    if (!pick) {
      // 換家族時，優先選一般字重
      styles.forEach(function (s) { if (!pick && /^(regular|roman|book|normal|w3|medium)$/i.test(s[0])) pick = s[1]; });
    }
    st.value = pick || styles[0][1];
  }

  function setFontName(name) {
    var el = $("fontName");
    if (el.value === name) return;
    el.value = name;
    el.dispatchEvent(new Event("change"));
  }

  function loadFonts(force) {
    if (!force) {
      try {
        var cached = JSON.parse(localStorage.getItem(FONT_KEY) || "null");
        if (cached && cached.length) { setFontData(cached); return; }
      } catch (e) {}
    }
    if (!cep) return;
    setStatus("讀取字體清單中…");
    run("CMF.listFonts()", function (r) {
      if (!r.ok || !r.fonts) return;
      try { localStorage.setItem(FONT_KEY, JSON.stringify(r.fonts)); } catch (e) {}
      setFontData(r.fonts);
    });
  }

  $("fontFamily").addEventListener("change", function () {
    fillStyles(null);
    setFontName($("fontFamily").value ? $("fontStyle").value : "");
  });
  $("fontStyle").addEventListener("change", function () { setFontName($("fontStyle").value); });
  $("btnFonts").addEventListener("click", function () { loadFonts(true); });
  App.reloadFonts = function () { loadFonts(true); };

  // 預覽用：PostScript 名稱 → CSS 字體設定（尺寸分頁的圖也用）
  function previewFont(o) { return App.previewFont(o.fontName); }
  App.previewFont = function (name) {
    var hit = fontIndex[name];
    if (!hit) return { family: "sans-serif", weight: 400, italic: false };
    var st = hit[1];
    var weight = /thin|hairline|w1/i.test(st) ? 200 : /extralight|ultralight|light|w2|w3/i.test(st) ? 300
      : /semibold|demibold|w6/i.test(st) ? 600 : /extrabold|ultrabold|heavy|black|w8|w9/i.test(st) ? 800
      : /bold|w7/i.test(st) ? 700 : /medium|w5/i.test(st) ? 500 : 400;
    return { family: "'" + hit[0].replace(/'/g, "") + "', sans-serif", weight: weight, italic: /italic|oblique/i.test(st) };
  };

  // ---------- 取色 ----------
  // 每個顏色欄位旁加一個滴管按鈕：按下後點畫布上的物件，取它的顏色
  var DROPPER = '<svg viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><path d="M13.7 2.3a2 2 0 0 0-2.8 0L9.2 4 8.5 3.3 7.1 4.7l.7.7-5.1 5.1-.5 2.4-.9.9 1.4 1.4.9-.9 2.4-.5 5.1-5.1.7.7 1.4-1.4-.7-.7 1.7-1.7a2 2 0 0 0 0-2.8zM5.1 12.3l-1.3.3.3-1.3 5.1-5.1 1 1z"/></svg>';
  var STROKE_FIRST = { lineColor: true, badgeStrokeColor: true, dimLineColor: true };
  var picking = null, pickTimer = null, pickBusy = false, pickStart = 0;

  Array.prototype.forEach.call(document.querySelectorAll('input[type=color]'), function (input) {
    var box = document.createElement("span");
    box.className = "colorbox";
    input.parentNode.insertBefore(box, input);
    box.appendChild(input);
    var btn = document.createElement("button");
    btn.className = "pick";
    btn.type = "button";
    btn.title = "從畫布上的物件取色";
    btn.innerHTML = DROPPER;
    btn.addEventListener("click", function () {
      if (picking === input.id) { stopPick("已取消取色"); return; }
      startPick(input.id, btn);
    });
    box.appendChild(btn);
    // 顏色欄位停用時，滴管也停用
    new MutationObserver(function () { btn.disabled = input.disabled; })
      .observe(input, { attributes: true, attributeFilter: ["disabled"] });
  });

  function startPick(id, btn) {
    if (!cep) { setStatus("請在 Illustrator 中開啟此面板", true); return; }
    if (adding) stopAdd(true);
    if (picking) stopPick();
    callHost("beginPick", null, function (r) {
      if (!r.ok) return;
      picking = id;
      App.busy.pick = true;
      pickStart = Date.now();
      btn.classList.add("active");
      setStatus("點一下畫布上要取色的物件（再按一次滴管或 Esc 取消）");
      pickTimer = setInterval(pollPick, 150);
    });
  }

  function stopPick(msg, isError) {
    clearInterval(pickTimer);
    pickTimer = null;
    picking = null;
    App.busy.pick = false;
    Array.prototype.forEach.call(document.querySelectorAll(".pick.active"), function (b) { b.classList.remove("active"); });
    if (cep) cep.evalScript("CMF.endPick()", function () {});
    if (msg) setStatus(msg, isError);
  }

  function pollPick() {
    if (!picking || pickBusy) return;
    if (Date.now() - pickStart > 20000) { stopPick("取色逾時，已取消"); return; }
    pickBusy = true;
    var id = picking;
    run("CMF.pollPick(" + q({ prefer: STROKE_FIRST[id] ? "stroke" : "fill" }) + ")", function (r) {
      pickBusy = false;
      if (picking !== id || !r || r.state !== "done") return;
      if (r.ok && r.hex) {
        var input = $(id);
        input.value = r.hex;
        input.dispatchEvent(new Event("input"));
        input.dispatchEvent(new Event("change"));
        if (id === "swColor") pickedSwatch(r.spot || "");
      }
      stopPick(r.msg, !r.ok);
    }, true);
  }

  // 色號的滴管：記下色票名稱（色塊用文件裡的色票），文字空白或還是上次自動帶入的名稱時，換成這個名稱
  var autoText = "Pantone";   // 預設的文字也算自動帶入的：滴管取到色票名稱時換掉
  function pickedSwatch(spot) {
    $("swSpot").value = spot;
    var t = $("swText");
    if (spot && (!t.value.trim() || t.value === autoText)) { t.value = spot; autoText = spot; }
    renderPreview();
    saveSettings();
  }
  // 自己改了色塊的顏色：不是色票了
  $("swColor").addEventListener("input", function () { $("swSpot").value = ""; });

  // ---------- 編號表 ----------
  // 有指定 CMF 清單時，每一列可以選擇對應的物件（名稱來自 Excel）
  var refreshing = false, refreshAgain = false;
  function refreshTable() {
    if (!cep) return;
    if (refreshing) { refreshAgain = true; return; }
    refreshing = true;
    run("CMF.listNumbers()", function (r) {
      var info = App.excel ? App.excel.cmfInfo() : Promise.resolve(null);
      info.then(function (list) {
        renderTable(r, list);
        if (App.excel) App.excel.numbersChanged(r);
      }).catch(function () { renderTable(r, null); }).then(function () {
        refreshing = false;
        if (refreshAgain) { refreshAgain = false; refreshTable(); }
      });
    }, true);
  }

  function option(value, text, selected) {
    var o = document.createElement("option");
    o.value = value;
    o.textContent = text;
    o.selected = !!selected;
    return o;
  }

  function renderTable(r, list) {
    var box = $("numTable");
    var items = (r && r.items) || [];
    var names = list && !list.error ? list.names : null;
    var total = items.reduce(function (s, it) { return s + it.count; }, 0);

    var info = $("tableInfo"), meta = "";
    if (items.length) meta = items.length + " 個" + (total > items.length ? "，" + total + " 處" : "");
    info.classList.remove("warn");
    info.title = "";
    if (list && list.error) {
      meta = list.error;
      info.classList.add("warn");
    } else if (names) {
      var linked = items.filter(function (it) { return it.key; }).length;
      if (items.length) meta += " · 已對應 " + linked;
      info.title = "CMF 清單：" + list.label;
    }
    info.textContent = meta;
    $("btnLinkExcel").hidden = !!list;

    // 正在編輯某一格時不重畫，避免打字被打斷
    var active = document.activeElement;
    if (box.contains(active) && (active.tagName === "INPUT" || active.tagName === "SELECT")) return;

    if (!items.length) {
      box.innerHTML = '<div class="empty">還沒有標註</div>';
      return;
    }
    var known = {}, owner = {};
    if (names) names.forEach(function (n) { known[n.name] = true; });
    items.forEach(function (it) { if (it.key && owner[it.key] === undefined) owner[it.key] = it.num; });

    box.innerHTML = "";
    items.forEach(function (it) {
      var row = document.createElement("div");
      row.className = "numrow";
      row.title = "點一下選取並移到這個標註";

      var input = document.createElement("input");
      input.type = "number";
      input.min = "1";
      input.setAttribute("data-scrub", "off");   // 不能拖曳、滾輪改：改了就會重新編號
      input.value = it.num;
      input.setAttribute("aria-label", "編號 " + it.num);
      input.title = "改號碼後按 Enter；號碼已存在時兩個互換";
      input.addEventListener("click", function (e) { e.stopPropagation(); });
      input.addEventListener("keydown", function (e) {
        if (e.key === "Enter") input.blur();
        if (e.key === "Escape") { input.value = it.num; input.blur(); }
      });
      input.addEventListener("change", function () {
        var to = parseInt(input.value, 10);
        if (!(to > 0) || to === it.num) { input.value = it.num; return; }
        callHost("changeNumber", { from: it.num, to: to }, function () { setTimeout(refreshTable, 0); });
      });
      row.appendChild(input);

      if (names) {
        var sel = document.createElement("select");
        sel.setAttribute("aria-label", "編號 " + it.num + " 對應的物件");
        sel.appendChild(option("", "—", !it.key));
        names.forEach(function (n) {
          var other = owner[n.name] !== undefined && owner[n.name] !== it.num ? "（" + owner[n.name] + "）" : "";
          sel.appendChild(option(n.name, n.name + other, n.name === it.key));
        });
        if (it.key && !known[it.key]) {
          sel.appendChild(option(it.key, it.key + "（清單中沒有）", true));
          row.classList.add("is-missing");
        }
        sel.classList.toggle("is-empty", !it.key);
        sel.title = it.key || "選擇對應的物件";
        sel.addEventListener("click", function (e) { e.stopPropagation(); });
        sel.addEventListener("change", function () {
          callHost("setLink", { num: it.num, key: sel.value }, function () {
            sel.blur();
            setTimeout(refreshTable, 0);
          });
        });
        row.appendChild(sel);
      } else {
        row.appendChild(document.createElement("span"));
      }

      var count = document.createElement("span");
      count.className = "count";
      count.textContent = it.count > 1 ? "×" + it.count : "";
      count.title = it.count > 1 ? it.count + " 處使用這個編號" : "";
      row.appendChild(count);

      row.addEventListener("click", function () { callHost("selectNumber", { num: it.num, center: true }); });
      box.appendChild(row);
    });
  }

  // ---------- 互動新增 ----------
  function updateFinishButton() {
    $("btnFinish").classList.toggle("hidden", !(adding && $("lineMode").value === "free" && !noLine()));
  }

  function finishFree() {
    if (!adding) return;
    run("CMF.finishFree(" + q(readSettings()) + ")", function (r) {
      if (!r.ok) return;
      afterAdded(r);
      if ($("continuous").checked) setStatus(r.msg + "，繼續點下一個目標點");
      else stopAdd(false);
    });
  }

  var adding = false, pollTimer = null, polling = false, idleSince = 0;
  var IDLE_LIMIT = 45000; // 連續模式閒置太久自動結束，避免誤轉換其他線段

  function startAdd() {
    if (!cep) { setStatus("請在 Illustrator 中開啟此面板", true); return; }
    if (picking) stopPick();
    App.emit("pen", "callout");   // 尺寸分頁的「兩點距離」也用鋼筆：先結束它
    callHost("beginAdd", null, function (r) {
      if (!r.ok) return;
      adding = true;
      App.busy.callout = true;
      idleSince = Date.now();
      $("btnAdd").textContent = "結束新增";
      $("btnAdd").classList.add("active");
      updateFinishButton();
      if (!r.msg) setStatus(modeTip());
      pollTimer = setInterval(poll, 200);
    });
  }

  function stopAdd(keepTool) {
    clearInterval(pollTimer);
    pollTimer = null;
    adding = false;
    App.busy.callout = false;
    $("btnAdd").textContent = addLabel();
    $("btnAdd").classList.remove("active");
    updateFinishButton();
    run("CMF.endAdd(" + (keepTool ? "true" : "false") + "," + q(readSettings()) + ")", afterAdded);
  }

  function afterAdded(r) {
    if (r && r.ok && r.next && $("startMode").value === "manual") {
      $("startNumber").value = r.next;
      saveSettings();
    }
    refreshTable();
  }

  function poll() {
    if (polling || !adding) return;
    polling = true;
    run("CMF.pollAdd(" + q(readSettings()) + ")", function (r) {
      polling = false;
      if (!adding || !r) return;
      if (r.state === "drawing") { idleSince = Date.now(); return; }
      if (r.state === "error") { setStatus(r.msg, true); stopAdd(true); return; }
      if (r.state === "done") {
        idleSince = Date.now();
        afterAdded(r);
        if ($("continuous").checked) {
          setStatus(r.msg + "，繼續點下一個目標點");
        } else {
          setStatus(r.msg);
          stopAdd(false);
        }
        return;
      }
      if ($("continuous").checked && Date.now() - idleSince > IDLE_LIMIT) {
        setStatus("閒置過久，已結束新增");
        stopAdd(true);
      }
    }, true);
  }

  // 送給「同步全部」、「套用選取」的樣式：圖上選著一個部位時，帶上 only（那個部位的欄位），其他設定每個標註維持原樣。
  // full = 不管選著的部位，整份樣式（「重設」大小一次改好幾個部位）
  function syncStyle(full) {
    var o = readSettings(), part = full ? null : fc.current();
    if (part && PART_FIELDS[part]) { o.only = PART_FIELDS[part]; o.onlyName = PART_NAMES[part]; }
    return o;
  }

  // ---------- 修改樣式時自動套用到選取的標註（⚡） ----------
  var syncTimer = null;
  function scheduleSync(full) {
    if (!$("autoSync").checked || !cep) return;
    clearTimeout(syncTimer);
    var st = syncStyle(full);
    // 只改選取的標註（跟尺寸分頁的 ⚡ 一樣）；沒有選取標註時不提示
    syncTimer = setTimeout(function () { callHost("restyle", st, function (r) { if (r.ok) setStatus(r.msg); }, true); }, 700);
  }

  // ---------- 吸取樣式 ----------
  // 選一個標註按「吸取樣式」：面板的樣式改成這個標註記錄的樣式，之後新增的標註、按「同步全部」就跟它一樣。
  // 只改面板（快捷鍵腳本用的設定檔也一起存），不會觸發「修改樣式時自動套用到選取的標註」。
  // 記錄裡沒有的項目（舊版的標註）或不認得的值，用 host.jsx 遇到時的做法
  var PICK_DEFAULT = { textColor: "#000000", fontName: "", badge: false, badgeStrokeWidth: 0, lineColor: "#000000",
    endStyle: "none", styleScale: 100 };

  // r = CMF.pickStyle()：{ callout, num, label, kind, callouts }。吸取到面板目前那一種的樣式（不切換種類）：
  // 例如在文字說明吸取編號標註的箭頭。同步不會互相影響，吸取可以跨種類。
  // 圖上選著一個部位時只吸取那個部位的設定
  function applyPicked(r) {
    var st = r.callout, part = fc.current(), keys = part && PART_FIELDS[part];
    // 舊版的文字說明，端點記在 noteEnd
    if (r.kind === "note" && st.noteEnd) st.endStyle = st.noteEnd;
    STYLE_FIELDS.forEach(function (k) {
      if (keys && keys.indexOf(k) < 0) return;
      if (!App.setField($(k), st[k]) && PICK_DEFAULT[k] !== undefined) App.setField($(k), PICK_DEFAULT[k]);
    });
    syncFontSelects();
    renderPreview();
    saveSettings();
    var msg = "已吸取" + (r.num !== undefined ? "標註 " + r.num + " " : r.label || "標註") + "的" + (keys ? PART_NAMES[part] : "樣式");
    if (r.callouts > 1) msg += "（選了 " + r.callouts + " 個標註，吸取的是這一個）";
    if (st.fontName && Object.keys(fontIndex).length && !fontIndex[st.fontName]) msg += "；字體清單裡沒有「" + st.fontName + "」";
    setStatus(msg);
  }

  function pickStyle() {
    if (picking) stopPick();
    callHost("pickStyle", null, function (r) {
      if (r.ok && r.callout) applyPicked(r);
      else if (r.ok && r.dim) { App.showTab("dim"); App.emit("pickedDim", r); }   // 選的是尺寸：改尺寸分頁的樣式
      else setStatus(r.msg || "請先選取要吸取樣式的標註或尺寸", true);
    }, true);
  }

  // ---------- 事件 ----------
  // 打字時（input）只更新預覽和面板的設定；按 Enter、離開欄位或改選項時（change）才呼叫 Illustrator
  // 存設定檔、自動同步（打字中途呼叫 Illustrator 會搶走鍵盤焦點，見 App.typing）
  FIELDS.forEach(function (k) {
    var isStyle = STYLE_FIELDS.indexOf(k) >= 0;
    $(k).addEventListener("input", function () {
      renderPreview();
      saveLocal();
    });
    $(k).addEventListener("change", function () {
      renderPreview();
      saveSettings();
      if (isStyle) scheduleSync();
    });
  });

  // 編號／色號／文字說明：存在 kind 裡（設定檔、快捷鍵腳本也看得到）。新增中切換，下一條線就用新的種類
  ["kindNumber", "kindSwatch", "kindNote"].forEach(function (id) {
    $(id).addEventListener("change", function () {
      if (!$(id).checked) return;
      switchKind($(id).value);
      if (adding) setStatus(modeTip());
    });
  });
  $("btnAdd").addEventListener("click", function () { if (adding) stopAdd(false); else startAdd(); });
  $("btnFinish").addEventListener("click", finishFree);
  App.on("pen", function (who) { if (who !== "callout" && adding) stopAdd(true); });
  document.addEventListener("keydown", function (e) {
    if (e.key === "Escape" && picking) { stopPick("已取消取色"); return; }
    if (e.key === "Escape" && adding && !/^(INPUT|TEXTAREA)$/.test(document.activeElement.tagName)) stopAdd(false);
  });

  $("btnConvert").addEventListener("click", function () { callHost("convert", readSettings(), afterAdded); });
  $("btnRenumber").addEventListener("click", function () { callHost("renumber", readSettings(), refreshTable); });
  $("btnRefresh").addEventListener("click", refreshTable);
  $("btnLinkExcel").addEventListener("click", function () {
    App.showTab("excel");
    var box = $("cmfBox");
    if (box) { box.open = true; box.scrollIntoView(); }
  });
  $("btnSync").addEventListener("click", function () { callHost("syncAll", syncStyle()); });
  $("btnRestyle").addEventListener("click", function () { callHost("restyle", syncStyle()); });
  $("btnRelayout").addEventListener("click", function () { callHost("relayout"); });
  $("btnAlignRow").addEventListener("click", function () { callHost("alignLabels", { dir: "row" }); });
  $("btnAlignCol").addEventListener("click", function () { callHost("alignLabels", { dir: "col" }); });
  $("btnPickStyle").addEventListener("click", pickStyle);
  App.on("pickedCallout", applyPicked);   // 在尺寸分頁按「吸取樣式」、選的是標註
  $("btnSelectAll").addEventListener("click", function () { callHost("selectAll"); });
  $("btnToggle").addEventListener("click", function () { callHost("toggleLayer"); });
  $("btnSetNumber").addEventListener("click", function () {
    callHost("setNumber", { num: parseInt($("setNumberValue").value, 10) }, refreshTable);
  });

  // 滑鼠移回面板或切換文件時，更新編號表（使用者可能在畫布上刪除或複製了標註）
  var lastEnter = 0;
  document.body.addEventListener("mouseenter", App.whenIdle(function () {
    if (adding || Date.now() - lastEnter < 800) return;
    lastEnter = Date.now();
    refreshTable();
  }));

  if (cep) {
    try { cep.addEventListener("com.cmf.callout.add", function () { if (adding) stopAdd(false); else startAdd(); }); } catch (e) {}
    try { cep.addEventListener("com.cmf.callout.finish", finishFree); } catch (e) {}
    try { cep.addEventListener("documentAfterActivate", refreshTable); } catch (e) {}
  }
  App.on("theme", renderPreview);
  App.on("numbers", refreshTable);   // Excel 分頁改了 CMF 清單或自動對應之後

  // 縮放分頁縮放了文件：樣式裡跟尺寸有關的數值用同一個倍率縮放
  // （跟 host.jsx 的 scaleNote 一樣的欄位；只存檔，不觸發自動同步）
  var SIZE_FIELDS = ["fontSize", "badgePadding", "badgeStrokeWidth", "lineWidth", "endSize", "styleScale", "gap", "chipSize", "chipBorder"];
  App.on("scaled", function (s) {
    SIZE_FIELDS.forEach(function (k) {
      var el = $(k);
      el.value = Math.round((parseFloat(el.value) || 0) * s * 1000) / 1000;
    });
    // 另外兩種存起來的樣式也一起縮放
    Object.keys(kindStyles).forEach(function (kind) {
      if (kind === kindOf()) return;
      SIZE_FIELDS.forEach(function (k) {
        var v = parseFloat(kindStyles[kind][k]);
        if (isFinite(v)) kindStyles[kind][k] = Math.round(v * s * 1000) / 1000;
      });
    });
    renderPreview();
    saveSettings();
  });

  // 圖右上角的「重設」：這些大小改回預設值（index.html 寫的值），顏色、字體、圓圈底開關、端點樣式不動。
  // 跟改其他樣式一樣：存起來，勾了「修改樣式時自動套用到選取的標註」就套用到選取的標註
  // 預設值依目前工作區的大小等比放大縮小（App.artboardFactor：A4 = 預設值）
  $("btnReset").addEventListener("click", function () {
    App.artboardFactor(function (k, note) {
      SIZE_FIELDS.forEach(function (id) { var el = $(id); el.value = App.scaledDefault(el, k); });
      renderPreview();
      saveSettings();
      scheduleSync(true);
      setStatus("大小改回預設" + (note ? "（" + note + "）" : "") + "：字級 " + $("fontSize").value + "、線寬 " + $("lineWidth").value +
        "、端點 " + $("endSize").value + "、圓圈內距 " + $("badgePadding").value + " pt");
    });
  });

  loadSettings(function () {
    loadFonts(false);
    renderPreview();
    saveSettings();
    refreshTable();
  });
  if (!cep) setStatus("預覽模式：請在 Illustrator 中使用", true);
})();
