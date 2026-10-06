/*
 * Excel 同步 — Illustrator 端 (ExtendScript / ES3)
 * 綁定方式：文字物件的「名稱」設為  xl:儲存格位址
 *   例：xl:B3、xl:工作表1!B3、xl:'價格 表'!C12
 * 名稱會顯示在圖層面板，可以直接在那裡檢查或手動修改。
 * 表格是一個群組，名稱為 xltable:範圍，備註（note）存 JSON：
 *   { mode: "excel" | "build", w, h, grid, cmf: { num, name, head, rest } }
 *   cmf：這個表格是 CMF 清單，依標註編號排序（見 js/table.js 的 cmfPlan）
 * 由 CMF Tool 面板在啟動時載入（$.evalFile）。
 */

var XL_PREFIX = "xl:";
var XL_TABLE = "xltable:";

/* ---------- 小工具（ExtendScript 沒有 JSON） ---------- */

function _q(s) {
    s = String(s);
    var out = '"';
    for (var i = 0; i < s.length; i++) {
        var c = s.charAt(i), code = s.charCodeAt(i);
        if (c === '"') out += '\\"';
        else if (c === "\\") out += "\\\\";
        else if (c === "\n") out += "\\n";
        else if (c === "\r") out += "\\r";
        else if (c === "\t") out += "\\t";
        else if (code < 32) out += "\\u" + ("000" + code.toString(16)).slice(-4);
        else out += c;
    }
    return out + '"';
}

function _arr(list) {
    var parts = [];
    for (var i = 0; i < list.length; i++) parts.push(_q(list[i]));
    return "[" + parts.join(",") + "]";
}

/* 簡單的 JSON 輸出（數字、字串、布林、null、陣列、物件） */
function _json(v) {
    if (v === null || v === undefined) return "null";
    if (typeof v === "number") return isFinite(v) ? String(v) : "null";
    if (typeof v === "boolean") return v ? "true" : "false";
    if (typeof v === "string") return _q(v);
    var parts = [], k;
    if (v instanceof Array) {
        for (k = 0; k < v.length; k++) parts.push(_json(v[k]));
        return "[" + parts.join(",") + "]";
    }
    for (k in v) if (v.hasOwnProperty(k) && v[k] !== undefined) parts.push(_q(k) + ":" + _json(v[k]));
    return "{" + parts.join(",") + "}";
}

function _refOf(item) {
    var n = item.name;
    if (n && n.indexOf(XL_PREFIX) === 0) return n.substring(XL_PREFIX.length);
    return null;
}

function _docKey(doc) {
    try { return doc.fullName.fsName; } catch (e) { return doc.name; }
}

/* .ai 所在資料夾；尚未存檔的文件回傳空字串 */
function _docFolder(doc) {
    try {
        var p = doc.path;
        return (p && p.fsName) ? p.fsName : "";
    } catch (e) { return ""; }
}

/* 從目前選取中找出所有文字物件（含群組內、正在編輯的文字） */
function _selectedTextFrames(doc) {
    var result = [];
    var sel = doc.selection;
    if (!sel) return result;

    // 用文字工具點在文字裡時，selection 是 TextRange
    if (sel.typename === "TextRange") {
        var frames = sel.story.textFrames;
        for (var t = 0; t < frames.length; t++) result.push(frames[t]);
        return result;
    }

    function walk(item) {
        if (item.typename === "TextFrame") result.push(item);
        else if (item.typename === "GroupItem") {
            for (var g = 0; g < item.pageItems.length; g++) walk(item.pageItems[g]);
        }
    }
    for (var i = 0; i < sel.length; i++) walk(sel[i]);
    return result;
}

/* ---------- 面板呼叫的函式 ---------- */

function es_context() {
    if (app.documents.length === 0) return '{"doc":null}';
    var doc = app.activeDocument;
    var frames = _selectedTextFrames(doc);
    var refs = [], seen = {};
    for (var i = 0; i < frames.length; i++) {
        var r = _refOf(frames[i]);
        if (r !== null && !seen[r]) { seen[r] = true; refs.push(r); }
    }
    return '{"doc":' + _q(_docKey(doc)) +
           ',"docName":' + _q(doc.name) +
           ',"folder":' + _q(_docFolder(doc)) +
           ',"selected":' + frames.length +
           ',"refs":' + _arr(refs) + "}";
}

function es_bind(ref) {
    if (app.documents.length === 0) return '{"error":"沒有開啟的文件"}';
    var frames = _selectedTextFrames(app.activeDocument);
    if (frames.length === 0) return '{"error":"請先選取文字物件"}';
    for (var i = 0; i < frames.length; i++) frames[i].name = XL_PREFIX + ref;
    return '{"count":' + frames.length + "}";
}

function es_unbind() {
    if (app.documents.length === 0) return '{"error":"沒有開啟的文件"}';
    var frames = _selectedTextFrames(app.activeDocument);
    var n = 0;
    for (var i = 0; i < frames.length; i++) {
        if (_refOf(frames[i]) !== null) { frames[i].name = ""; n++; }
    }
    return '{"count":' + n + "}";
}

/* 列出文件中用到的儲存格（表格外的）和表格範圍 */
function es_refs() {
    if (app.documents.length === 0) return '{"doc":null,"refs":[],"tables":[]}';
    var doc = app.activeDocument;
    var refs = [], seen = {};
    for (var i = 0; i < doc.textFrames.length; i++) {
        var tf = doc.textFrames[i];
        var r = _refOf(tf);
        if (r !== null && !seen[r] && !_inTable(tf)) { seen[r] = true; refs.push(r); }
    }
    // 同一個範圍可能有好幾個表格（例如複製到別的工作區域），只要其中一個是 CMF 清單就算
    var tables = [], byLabel = {}, groups = _tableGroups(doc);
    for (var g = 0; g < groups.length; g++) {
        var label = groups[g].name.substring(XL_TABLE.length), meta = _meta(groups[g]);
        var t = byLabel.hasOwnProperty(label) ? byLabel[label] : null;
        if (!t) { t = byLabel[label] = { label: label, mode: meta.mode || "build", cmf: null }; tables.push(t); }
        if (!t.cmf && meta.cmf) t.cmf = meta.cmf;
    }
    return '{"doc":' + _q(_docKey(doc)) + ',"refs":' + _arr(refs) + ',"tables":' + _json(tables) + "}";
}

/* 把值寫進文字物件。values = { "B3": "1,280", ... } */
function es_apply(valuesJson) {
    if (app.documents.length === 0) return '{"error":"沒有開啟的文件"}';
    var values = eval("(" + valuesJson + ")");
    var doc = app.activeDocument;
    var changed = 0, same = 0, failed = [], failedSeen = {};

    for (var i = 0; i < doc.textFrames.length; i++) {
        var tf = doc.textFrames[i];
        var ref = _refOf(tf);
        if (ref === null || !values.hasOwnProperty(ref) || _inTable(tf)) continue;

        // Illustrator 的換行是 \r
        var v = String(values[ref]).replace(/\r\n|\n/g, "\r");
        if (tf.contents === v) { same++; continue; }
        try {
            tf.contents = v;
            changed++;
        } catch (e) {
            // 通常是物件或圖層被鎖定／隱藏
            if (!failedSeen[ref]) { failedSeen[ref] = true; failed.push(ref); }
        }
    }
    if (changed > 0) app.redraw();
    return '{"changed":' + changed + ',"same":' + same + ',"failed":' + _arr(failed) + "}";
}

/* 選取所有綁定到某個儲存格的物件，方便找東西 */
function es_selectRef(ref) {
    if (app.documents.length === 0) return '{"count":0}';
    var doc = app.activeDocument;
    doc.selection = null;
    var n = 0;
    for (var i = 0; i < doc.textFrames.length; i++) {
        var tf = doc.textFrames[i];
        if (_refOf(tf) === ref) {
            try { tf.selected = true; n++; } catch (e) {}
        }
    }
    return '{"count":' + n + "}";
}

/* ====================================================================
 * 表格：每格一個文字物件，加上填色、框線；整組是一個群組
 * table = { label, width, height, grid, cells:[...], borders:[...] }
 * 座標單位 pt，從表格左上角起算，y 往下為正
 * ==================================================================== */

function _inTable(item) {
    var p = item.parent;
    while (p && p.typename === "GroupItem") {
        if (p.name.indexOf(XL_TABLE) === 0) return true;
        p = p.parent;
    }
    return false;
}

/* 表格群組備註裡的 JSON；讀不到就回傳空物件 */
function _meta(g) {
    try {
        var m = eval("(" + g.note + ")");
        return (m && typeof m === "object") ? m : {};
    } catch (e) { return {}; }
}

/* "excel" = 由 Excel 複製貼上；"build" = 外掛自己畫的 */
function _tableMode(g) {
    return _meta(g).mode || "build";
}

/* 換掉表格時，沿用舊表格的 CMF 清單設定 */
function _keepCmf(fresh, oldMeta) {
    if (!oldMeta || !oldMeta.cmf) return;
    var m = _meta(fresh);
    m.cmf = oldMeta.cmf;
    fresh.note = _json(m);
}

function _tableGroups(doc) {
    var list = [];
    for (var i = 0; i < doc.groupItems.length; i++) {
        if (doc.groupItems[i].name.indexOf(XL_TABLE) === 0) list.push(doc.groupItems[i]);
    }
    return list;
}

/* ---------- 顏色 ---------- */

var _colorCache = {};
function _color(hex) {
    var doc = app.activeDocument;
    var cmyk = doc.documentColorSpace === DocumentColorSpace.CMYK;
    var key = (cmyk ? "c" : "r") + hex;
    if (_colorCache[key]) return _colorCache[key];
    var r = parseInt(hex.substr(0, 2), 16), g = parseInt(hex.substr(2, 2), 16), b = parseInt(hex.substr(4, 2), 16);
    var col;
    if (cmyk) {
        col = new CMYKColor();
        if (r === 0 && g === 0 && b === 0) {          // 黑色用單色黑 K100，不要四色黑
            col.cyan = 0; col.magenta = 0; col.yellow = 0; col.black = 100;
        } else {
            var v = app.convertSampleColor(ImageColorSpace.RGB, [r, g, b], ImageColorSpace.CMYK, ColorConvertPurpose.defaultpurpose);
            col.cyan = v[0]; col.magenta = v[1]; col.yellow = v[2]; col.black = v[3];
        }
    } else {
        col = new RGBColor();
        col.red = r; col.green = g; col.blue = b;
    }
    _colorCache[key] = col;
    return col;
}

/* ---------- 字型 ---------- */

var _fontIndex = null, _fontCache = {};

function _fontFamilies() {
    if (_fontIndex) return _fontIndex;
    _fontIndex = {};
    var all = app.textFonts;
    for (var i = 0; i < all.length; i++) {
        var f = all[i], fam;
        try { fam = f.family.toLowerCase(); } catch (e) { continue; }
        (_fontIndex[fam] || (_fontIndex[fam] = [])).push(f);
    }
    return _fontIndex;
}

function _pickStyle(list, bold, italic, preferStyle) {
    var best = null, bestScore = -1;
    for (var i = 0; i < list.length; i++) {
        var s = String(list[i].style).toLowerCase();
        var isBold = /bold|black|heavy|semibold|demi|w[6-9]/.test(s);
        var isItalic = /italic|oblique/.test(s);
        var score = (isBold === bold ? 4 : 0) + (isItalic === italic ? 4 : 0);
        if (preferStyle && s === preferStyle) score += 8;
        if (/^(regular|normal|roman|book|medium|w3|w4)$/.test(s) || (bold && s === "bold")) score += 1;
        if (score > bestScore) { bestScore = score; best = list[i]; }
    }
    return best;
}

function _findFont(names, bold, italic) {
    var key = names.join("|") + "|" + bold + "|" + italic;
    if (_fontCache.hasOwnProperty(key)) return _fontCache[key];
    var fams = _fontFamilies(), found = null;
    for (var i = 0; i < names.length && !found; i++) {
        var n = String(names[i]);
        var list = fams[n.toLowerCase()];
        if (list) { found = _pickStyle(list, bold, italic, null); break; }
        // 例如 "Microsoft JhengHei Light" → 家族 "Microsoft JhengHei" + 樣式 "Light"
        var cut = n.lastIndexOf(" ");
        if (cut > 0) {
            list = fams[n.substring(0, cut).toLowerCase()];
            if (list) found = _pickStyle(list, bold, italic, n.substring(cut + 1).toLowerCase());
        }
        if (!found) { try { found = app.textFonts.getByName(n.replace(/\s+/g, "")); } catch (e2) {} }
    }
    _fontCache[key] = found;
    return found;
}

/* ---------- 建立表格 ---------- */

function _applyText(range, c, missing) {
    var ca = range.characterAttributes;
    ca.size = c.font.sz;
    var f = _findFont(c.font.names, c.font.b, c.font.i);
    if (f) ca.textFont = f; else missing[c.font.names[0]] = true;
    ca.fillColor = _color(c.font.color || "000000");
    ca.underline = !!c.font.u;
    ca.strikeThrough = !!c.font.st;
    range.paragraphAttributes.justification =
        c.ha === "right" ? Justification.RIGHT : c.ha === "center" ? Justification.CENTER : Justification.LEFT;
}

function _line(group, x1, y1, x2, y2, b) {
    var p = group.pathItems.add();
    p.setEntirePath([[x1, y1], [x2, y2]]);
    p.filled = false;
    p.stroked = true;
    p.strokeWidth = b.w;
    p.strokeColor = _color(b.color);
    if (b.dash) p.strokeDashes = b.dash;
    return p;
}

function _buildTable(t, left, top, missing) {
    var doc = app.activeDocument;
    var group = doc.groupItems.add();
    group.name = XL_TABLE + t.label;
    group.note = '{"w":' + t.width + ',"h":' + t.height + ',"grid":' + (t.grid ? "true" : "false") + "}";

    // 外框（看不見），用來記住表格的位置和縮放
    var frame = group.pathItems.rectangle(top, left, t.width, t.height);
    frame.filled = false;
    frame.stroked = false;
    frame.name = "xlframe";

    var i, c;
    var hasFill = false;
    for (i = 0; i < t.cells.length; i++) if (t.cells[i].fill) { hasFill = true; break; }
    if (hasFill) {
        var fills = group.groupItems.add();
        fills.name = "填色";
        for (i = 0; i < t.cells.length; i++) {
            c = t.cells[i];
            if (!c.fill) continue;
            var fr = fills.pathItems.rectangle(top - c.y, left + c.x, c.w, c.h);
            fr.stroked = false;
            fr.filled = true;
            fr.fillColor = _color(c.fill);
        }
    }

    if (t.grid) {
        var grid = group.groupItems.add();
        grid.name = "格線";
        for (i = 0; i < t.cells.length; i++) {
            c = t.cells[i];
            var gr = grid.pathItems.rectangle(top - c.y, left + c.x, c.w, c.h);
            gr.filled = false;
            gr.stroked = true;
            gr.strokeWidth = 0.5;
            gr.strokeColor = _color("D4D4D4");
        }
    }

    if (t.borders.length) {
        var lines = group.groupItems.add();
        lines.name = "框線";
        for (i = 0; i < t.borders.length; i++) {
            var b = t.borders[i];
            var x1 = left + b.x1, y1 = top - b.y1, x2 = left + b.x2, y2 = top - b.y2;
            if (b.double) {                                   // 雙線
                var dx = (y1 === y2) ? 0 : 0.75, dy = (y1 === y2) ? 0.75 : 0;
                _line(lines, x1 - dx, y1 - dy, x2 - dx, y2 - dy, b);
                _line(lines, x1 + dx, y1 + dy, x2 + dx, y2 + dy, b);
            } else {
                _line(lines, x1, y1, x2, y2, b);
            }
        }
    }

    var pad = 2.25;
    for (i = 0; i < t.cells.length; i++) {
        c = t.cells[i];
        var text = String(c.text).replace(/\r\n|\n/g, "\r");
        var sz = c.font.sz, tf;
        var cellTop = top - c.y, cellBottom = top - c.y - c.h;

        if (c.wrap) {
            // 自動換行：用區域文字，寬度等於儲存格
            var box = group.pathItems.rectangle(cellTop - 1, left + c.x + pad, Math.max(c.w - pad * 2, 1), Math.max(c.h - 2, 1));
            tf = group.textFrames.areaText(box);
        } else {
            var lineCount = text.split("\r").length, lead = sz * 1.2;
            var ax = c.ha === "right" ? left + c.x + c.w - pad - c.indent
                   : c.ha === "center" ? left + c.x + c.w / 2
                   : left + c.x + pad + c.indent;
            var ay = c.va === "top" ? cellTop - 1.5 - sz * 0.88
                   : c.va === "center" ? (cellTop + cellBottom) / 2 - sz * 0.32 + (lineCount - 1) * lead / 2
                   : cellBottom + sz * 0.22 + 1 + (lineCount - 1) * lead;
            try { tf = group.textFrames.pointText([ax, ay]); }
            catch (e) { tf = doc.textFrames.pointText([ax, ay]); tf.move(group, ElementPlacement.PLACEATEND); }
        }

        tf.contents = text === "" ? " " : text;     // 空字串無法設定樣式，先放空白
        _applyText(tf.textRange, c, missing);
        if (text === "") tf.contents = "";
        tf.name = XL_PREFIX + c.ref;
    }
    return group;
}

function _missingList(missing) {
    var list = [];
    for (var k in missing) if (missing.hasOwnProperty(k)) list.push(k);
    return list;
}

/* 匯入新表格，放在目前畫面中央 */
function es_importTable(tableJson) {
    if (app.documents.length === 0) return '{"error":"沒有開啟的文件"}';
    var t = eval("(" + tableJson + ")");
    var doc = app.activeDocument;
    var center = doc.activeView.centerPoint;
    var missing = {};
    var group = _buildTable(t, center[0] - t.width / 2, center[1] + t.height / 2, missing);
    doc.selection = null;
    group.selected = true;
    app.redraw();
    return '{"count":' + t.cells.length + ',"missingFonts":' + _arr(_missingList(missing)) + "}";
}

/* 更新：依 Excel 重建每個表格，保留原本的位置和縮放比例 */
function es_rebuildTables(dataJson) {
    if (app.documents.length === 0) return '{"error":"沒有開啟的文件"}';
    var data = eval("(" + dataJson + ")");
    var doc = app.activeDocument;
    var groups = _tableGroups(doc);
    var done = 0, failed = [], missing = {};

    for (var i = 0; i < groups.length; i++) {
        var old = groups[i];
        var label = old.name.substring(XL_TABLE.length);
        if (!data.hasOwnProperty(label)) continue;
        var t = data[label], fresh = null;
        try {
            var meta = null, frame = null;
            try { meta = eval("(" + old.note + ")"); } catch (e) {}
            for (var k = 0; k < old.pathItems.length; k++) {
                if (old.pathItems[k].name === "xlframe") { frame = old.pathItems[k]; break; }
            }
            var gb = (frame || old).geometricBounds;          // [左, 上, 右, 下]
            var scale = (frame && meta && meta.w) ? (gb[2] - gb[0]) / meta.w : 1;
            if (meta && meta.hasOwnProperty("grid")) t.grid = meta.grid;

            var fresh = _buildTable(t, gb[0], gb[1], missing);
            if (Math.abs(scale - 1) > 0.0001) {
                var pct = scale * 100;
                fresh.resize(pct, pct, true, true, true, true, pct, Transformation.TOPLEFT);
            }
            // 對齊回原本外框的左上角
            for (k = 0; k < fresh.pathItems.length; k++) {
                if (fresh.pathItems[k].name === "xlframe") {
                    var nb = fresh.pathItems[k].geometricBounds;
                    fresh.translate(gb[0] - nb[0], gb[1] - nb[1]);
                    break;
                }
            }
            fresh.move(old, ElementPlacement.PLACEBEFORE);
            _keepCmf(fresh, meta);
            old.remove();
            done++;
        } catch (err) {
            failed.push(label);
            if (fresh) { try { fresh.remove(); } catch (e3) {} }
        }
    }
    app.redraw();
    return '{"count":' + done + ',"failed":' + _arr(failed) + ',"missingFonts":' + _arr(_missingList(missing)) + "}";
}

/* 指定 CMF 清單（依標註排序）的表格。cfgJson 空字串 = 取消
 * 一份文件只有一個 CMF 清單：其他表格的設定會被清掉 */
function es_setCmf(label, cfgJson) {
    if (app.documents.length === 0) return '{"error":"沒有開啟的文件"}';
    var cfg = cfgJson ? eval("(" + cfgJson + ")") : null;
    var groups = _tableGroups(app.activeDocument), n = 0, failed = 0;
    for (var i = 0; i < groups.length; i++) {
        var g = groups[i], m = _meta(g), mine = cfg && g.name === XL_TABLE + label;
        if (!mine && !m.cmf) continue;
        if (mine) m.cmf = cfg; else delete m.cmf;
        try { g.note = _json(m); if (mine) n++; } catch (e) { failed++; }
    }
    return '{"count":' + n + ',"failed":' + failed + "}";
}

function es_selectTable(label) {
    if (app.documents.length === 0) return '{"count":0}';
    var doc = app.activeDocument;
    doc.selection = null;
    var groups = _tableGroups(doc), n = 0;
    for (var i = 0; i < groups.length; i++) {
        if (groups[i].name === XL_TABLE + label) {
            try { groups[i].selected = true; n++; } catch (e) {}
        }
    }
    return '{"count":' + n + "}";
}

/* ====================================================================
 * Excel 原生外觀：Excel 已經把範圍複製到剪貼簿，這裡貼上
 * mode "new"：放在畫面中央；"replace"：取代所有同一個範圍的表格，保留位置和縮放
 * ==================================================================== */
/* 剪貼簿可能暫時被別的程式佔用（OneDrive、剪貼簿工具，或 Excel 還在準備資料）：等一下再試 */
function _paste(doc) {
    var last = null;
    for (var t = 0; t < 5; t++) {
        try {
            doc.selection = null;
            app.paste();
        } catch (e) { last = e; }
        var sel = doc.selection;
        if (sel && sel.length) return sel;
        $.sleep(300);
    }
    throw new Error(last ? last.message : "剪貼簿裡沒有 Excel 的內容");
}

function es_pasteTable(label, mode) {
    try { return _pasteTable(label, mode); }
    catch (e) { return '{"error":' + _q("表格沒有更新：" + e.message) + "}"; }
}

function _pasteTable(label, mode) {
    if (app.documents.length === 0) return '{"error":"沒有開啟的文件"}';
    var doc = app.activeDocument;
    var name = XL_TABLE + label;
    var i, targets = [];

    if (mode === "replace") {
        var groups = _tableGroups(doc);
        for (i = 0; i < groups.length; i++) if (groups[i].name === name) targets.push(groups[i]);
        if (!targets.length) return '{"count":0,"failed":[]}';
    }

    var sel;
    try { sel = _paste(doc); }
    catch (e) { return '{"error":' + _q("無法貼上 Excel 複製的內容（" + e.message + "）") + ',"retry":true}'; }
    if (sel.length > 1 || sel[0].typename !== "GroupItem") {
        app.executeMenuCommand("group");
        sel = doc.selection;
    }
    var pasted = sel[0];
    var pb = pasted.geometricBounds;
    pasted.name = name;
    pasted.note = '{"mode":"excel","w":' + (pb[2] - pb[0]) + ',"h":' + (pb[1] - pb[3]) + "}";

    if (mode !== "replace") {
        app.redraw();
        return '{"count":1,"failed":[]}';
    }

    var done = 0, failed = [];
    for (i = 0; i < targets.length; i++) {
        var old = targets[i], copy = null;
        try {
            var meta = null;
            try { meta = eval("(" + old.note + ")"); } catch (e) {}
            var ob = old.geometricBounds;                       // [左, 上, 右, 下]
            var scale = (meta && meta.w) ? (ob[2] - ob[0]) / meta.w : 1;

            copy = pasted.duplicate(old, ElementPlacement.PLACEBEFORE);
            if (Math.abs(scale - 1) > 0.0001) {
                var pct = scale * 100;
                copy.resize(pct, pct, true, true, true, true, pct, Transformation.TOPLEFT);
            }
            var nb = copy.geometricBounds;
            copy.translate(ob[0] - nb[0], ob[1] - nb[1]);       // 對齊原本的左上角
            copy.name = pasted.name;
            copy.note = pasted.note;
            _keepCmf(copy, meta);
            old.remove();
            done++;
        } catch (err) {
            failed.push(label);
            if (copy) { try { copy.remove(); } catch (e2) {} }
        }
    }
    pasted.remove();
    doc.selection = null;
    app.redraw();
    return '{"count":' + done + ',"failed":' + _arr(failed) + "}";
}
