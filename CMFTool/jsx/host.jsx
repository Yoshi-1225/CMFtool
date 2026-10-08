/*
 * CMF Callout - Illustrator ExtendScript 核心：標註、尺寸、縮放
 * 對外函式接收 JSON 字串、回傳 JSON 字串：{"ok":true,"msg":"...","next":5,"state":"done"}
 * 注意：ExtendScript 是 ES3，沒有 JSON / let / 箭頭函式 / Array.indexOf
 *
 * 每個標註是一個群組，備註（note）存著：CMF_CALLOUT|編號|k=對應的物件名稱|樣式 JSON
 * 「k=」那一段只有對應過 Excel 清單的物件才有，名稱用 encodeURIComponent 編碼。
 * 尺寸也是群組，備註存著：CMF_DIM|幾何 JSON|樣式 JSON（見「尺寸標註」）
 */
$.global.CMF = (function () {
    var LAYER_NAME = "CMF Callouts";
    var TAG = "CMF_CALLOUT";
    var SETTINGS_DIR = Folder.userData + "/CMFCallout";
    var SETTINGS_FILE = SETTINGS_DIR + "/settings.json";

    // ---------- 共用 ----------
    function parse(s) { return eval("(" + s + ")"); }

    function esc(s) {
        return String(s).replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/[\r\n\t]+/g, " ")
            .replace(/[\x00-\x1f]/g, "");
    }

    function res(ok, msg, next, state) {
        var out = '{"ok":' + (ok ? "true" : "false") + ',"msg":"' + esc(msg) + '"';
        if (next !== undefined && next !== null) out += ',"next":' + next;
        if (state) out += ',"state":"' + state + '"';
        return out + "}";
    }

    function getDoc() {
        if (app.documents.length === 0) throw new Error("沒有開啟的文件");
        return app.activeDocument;
    }

    function hexToColor(hex, d) {
        hex = String(hex || "#000000").replace("#", "");
        if (hex.length === 3) hex = hex.charAt(0) + hex.charAt(0) + hex.charAt(1) + hex.charAt(1) + hex.charAt(2) + hex.charAt(2);
        var r = parseInt(hex.substr(0, 2), 16), g = parseInt(hex.substr(2, 2), 16), b = parseInt(hex.substr(4, 2), 16);
        if (d.documentColorSpace == DocumentColorSpace.CMYK) {
            var c = new CMYKColor();
            var k = 1 - Math.max(r, g, b) / 255;
            if (k >= 0.9999) { c.cyan = 0; c.magenta = 0; c.yellow = 0; c.black = 100; return c; }
            c.cyan = (1 - r / 255 - k) / (1 - k) * 100;
            c.magenta = (1 - g / 255 - k) / (1 - k) * 100;
            c.yellow = (1 - b / 255 - k) / (1 - k) * 100;
            c.black = k * 100;
            return c;
        }
        var rgb = new RGBColor();
        rgb.red = r; rgb.green = g; rgb.blue = b;
        return rgb;
    }

    // name 省略 = 標註圖層
    // 重建標註、尺寸前，暫時解除它和上層圖層、群組的鎖定和隱藏（鎖定、隱藏的東西改不了），做完用 restoreTemp 還原
    function openAncestors(it) {
        var saved = [];
        for (var p = it.parent; p && p.typename !== "Document"; p = p.parent) {
            setTemp(p, "locked", false, saved);
            setTemp(p, p.typename === "Layer" ? "visible" : "hidden", p.typename === "Layer", saved);
        }
        setTemp(it, "locked", false, saved);
        setTemp(it, "hidden", false, saved);
        return saved;
    }

    function restoreTemp(saved) {
        for (var i = saved.length - 1; i >= 0; i--) { try { saved[i][0][saved[i][1]] = saved[i][2]; } catch (e) {} }
    }

    // 換成重建好的新群組時，保留原本的鎖定、隱藏
    function keepState(ng, locked, hidden) {
        try { if (hidden) ng.hidden = true; if (locked) ng.locked = true; } catch (e) {}
    }

    function getLayer(d, create, name) {
        var L = null;
        name = name || LAYER_NAME;
        try { L = d.layers.getByName(name); } catch (e) { L = null; }
        if (!L && create) { L = d.layers.add(); L.name = name; }
        if (L && create) { L.locked = false; L.visible = true; }
        return L;
    }

    function isCallout(it) {
        return it && it.typename === "GroupItem" && it.note && String(it.note).indexOf(TAG + "|") === 0;
    }

    function noteParts(g) {
        var p = String(g.note).split("|"), rest = p.slice(2), key = "";
        if (rest.length > 1 && rest[0].indexOf("k=") === 0) {
            try { key = decodeURIComponent(rest[0].substring(2)); } catch (e) { key = ""; }
            rest = rest.slice(1);
        }
        return { num: parseInt(p[1], 10), key: key, style: rest.join("|") };
    }

    function makeNote(num, key, styleStr) {
        return TAG + "|" + num + "|" + (key ? "k=" + encodeURIComponent(key) + "|" : "") + styleStr;
    }

    // 圖層面板上顯示：CMF 3 外殼
    function groupName(num, key) { return "CMF " + num + (key ? " " + key : ""); }

    function findChild(g, name) {
        for (var i = 0; i < g.pageItems.length; i++) if (g.pageItems[i].name === name) return g.pageItems[i];
        return null;
    }

    function contains(arr, it) {
        for (var i = 0; i < arr.length; i++) if (arr[i] === it) return true;
        return false;
    }

    function allCallouts(d) {
        var out = [], gs = d.groupItems;
        for (var i = 0; i < gs.length; i++) if (isCallout(gs[i])) out.push(gs[i]);
        return out;
    }

    function calloutsWithNumber(d, n) {
        var cs = allCallouts(d), out = [];
        for (var i = 0; i < cs.length; i++) if (noteParts(cs[i]).num === n) out.push(cs[i]);
        return out;
    }

    function selectedCallouts(d) {
        var out = [], sel = d.selection;
        if (!sel) return out;
        for (var i = 0; i < sel.length; i++) {
            var it = sel[i];
            while (it && it.typename !== "Layer" && !isCallout(it)) it = it.parent;
            if (isCallout(it) && !contains(out, it)) out.push(it);
        }
        return out;
    }

    function maxNumber(d) {
        var cs = allCallouts(d), m = 0;
        for (var i = 0; i < cs.length; i++) { var n = noteParts(cs[i]).num; if (n > m) m = n; }
        return m;
    }

    function norm(dx, dy, fx, fy) {
        var l = Math.sqrt(dx * dx + dy * dy);
        if (l < 0.0001) return [fx, fy];
        return [dx / l, dy / l];
    }

    // 文字實際字形外框（轉外框量測後刪除）
    function glyphBounds(tf) {
        var ol = tf.duplicate().createOutline();
        var b = ol.geometricBounds; // [left, top, right, bottom]
        ol.remove();
        return b;
    }

    // 反轉路徑方向（保留曲線把手）
    function reversePath(path) {
        var pts = path.pathPoints, n = pts.length, data = [], i;
        for (i = 0; i < n; i++) data.push([pts[i].anchor, pts[i].leftDirection, pts[i].rightDirection, pts[i].pointType]);
        for (i = 0; i < n; i++) {
            var src = data[n - 1 - i], p = pts[i];
            p.anchor = src[0];
            p.leftDirection = src[2];
            p.rightDirection = src[1];
            p.pointType = src[3];
        }
    }

    // ---------- 排序 ----------
    // getPt 回傳 [x, y]；Illustrator 座標 y 往上為正
    function sortByMode(arr, getPt, mode) {
        if (mode === "order" || arr.length < 2) return arr;
        var w = [], i, minx = 1e9, maxx = -1e9, miny = 1e9, maxy = -1e9;
        for (i = 0; i < arr.length; i++) {
            var p = getPt(arr[i]);
            w.push({ it: arr[i], x: p[0], y: p[1], idx: i });
            if (p[0] < minx) minx = p[0]; if (p[0] > maxx) maxx = p[0];
            if (p[1] < miny) miny = p[1]; if (p[1] > maxy) maxy = p[1];
        }
        var cx = (minx + maxx) / 2, cy = (miny + maxy) / 2, tol = 8;
        for (i = 0; i < w.length; i++) {
            var a = Math.atan2(w[i].x - cx, w[i].y - cy); // 12 點鐘 = 0，順時針遞增
            if (a < 0) a += Math.PI * 2;
            w[i].a = a;
        }
        w.sort(function (p, q) {
            if (mode === "clockwise") return (p.a - q.a) || (p.idx - q.idx);
            if (mode === "topdown") return Math.abs(p.y - q.y) > tol ? q.y - p.y : p.x - q.x;
            if (mode === "leftright") return Math.abs(p.x - q.x) > tol ? p.x - q.x : q.y - p.y;
            return p.idx - q.idx;
        });
        var out = [];
        for (i = 0; i < w.length; i++) out.push(w[i].it);
        return out;
    }

    function labelPt(g) {
        var line = findChild(g, "CMF_Line");
        return line ? line.pathPoints[0].anchor : [g.left, g.top];
    }

    // ---------- 建立單一標註 ----------
    // line：開放路徑，第一個錨點 = 編號端，最後一個錨點 = 指向產品端
    // （使用者畫的方向相反，呼叫前會先 reversePath）
    // key：對應的物件名稱（可省略）
    function buildCallout(d, line, num, styleStr, container, key) {
        var o = parse(styleStr);
        var pts = line.pathPoints, n = pts.length;
        var S = pts[0].anchor, S1 = pts[1].anchor;
        var E = pts[n - 1].anchor, E1 = pts[n - 2].anchor;
        var lineColor = hexToColor(o.lineColor, d);

        var g = container.groupItems.add();
        g.name = groupName(num, key);
        g.note = makeNote(num, key, styleStr);

        line.move(g, ElementPlacement.PLACEATEND);
        line.name = "CMF_Line";
        line.filled = false;
        line.stroked = true;
        line.strokeColor = lineColor;
        line.strokeWidth = Number(o.lineWidth);
        line.strokeDashes = [];
        line.strokeJoin = StrokeJoin.ROUNDENDJOIN;

        // 端點
        var size = Number(o.endSize);
        if (o.endStyle === "dot") {
            var r = size / 2;
            var dot = g.pathItems.ellipse(E[1] + r, E[0] - r, r * 2, r * 2);
            dot.name = "CMF_End";
            dot.stroked = false; dot.filled = true; dot.fillColor = lineColor;
        } else if (o.endStyle === "arrow") {
            var u = norm(E[0] - E1[0], E[1] - E1[1], 1, 0);
            var len = size * 1.6, hw = size * 0.5;
            var B = [E[0] - u[0] * len, E[1] - u[1] * len];
            var tri = g.pathItems.add();
            tri.setEntirePath([[E[0], E[1]], [B[0] - u[1] * hw, B[1] + u[0] * hw], [B[0] + u[1] * hw, B[1] - u[0] * hw]]);
            tri.closed = true;
            tri.name = "CMF_End";
            tri.stroked = false; tri.filled = true; tri.fillColor = lineColor;
            // 線縮到箭頭底部，避免線頭穿出箭頭尖
            var last = line.pathPoints[n - 1];
            last.anchor = B; last.leftDirection = B; last.rightDirection = B;
        } else if (o.endStyle === "style") {
            var gs;
            try { gs = d.graphicStyles.getByName(o.styleName); } catch (e) { gs = null; }
            if (!gs) throw new Error("找不到繪圖樣式「" + o.styleName + "」，請先在繪圖樣式面板建立");
            gs.applyTo(line);
            // 縮放過的文件：繪圖樣式的線寬、箭頭跟著放大縮小（舊的標註沒有 styleScale = 100%）
            var pct = Number(o.styleScale);
            if (pct > 0 && Math.abs(pct - 100) > 0.001) {
                var prefs = [];
                setPref("scaleLineWeight", true, prefs);
                try { line.resize(100, 100, true, true, true, true, pct, Transformation.CENTER); }
                finally { for (var pi = 0; pi < prefs.length; pi++) { try { app.preferences.setBooleanPreference(prefs[pi][0], prefs[pi][1]); } catch (eP) {} } }
            }
        }

        // 編號文字
        var t = g.textFrames.add();
        t.name = "CMF_Label";
        t.contents = String(num);
        var ca = t.textRange.characterAttributes;
        ca.size = Number(o.fontSize);
        if (o.fontName) {
            try { ca.textFont = app.textFonts.getByName(o.fontName); } catch (e2) { /* 找不到字體就用預設 */ }
        }
        ca.fillColor = hexToColor(o.badge ? o.badgeTextColor : o.textColor, d);
        t.textRange.paragraphAttributes.justification = Justification.CENTER;

        var b = glyphBounds(t);
        var cx = (b[0] + b[2]) / 2, cy = (b[1] + b[3]) / 2;
        var hwT = (b[2] - b[0]) / 2, hhT = (b[1] - b[3]) / 2;

        // 從線的第一段往外推，算出編號中心
        var dir = norm(S[0] - S1[0], S[1] - S1[1], -1, 0);
        var gap = Number(o.gap), extent, R = 0;
        var bsw = o.badge ? Number(o.badgeStrokeWidth) || 0 : 0;
        if (o.badge) {
            R = Math.max(hwT, hhT) + Number(o.badgePadding);
            extent = R + bsw / 2; // 邊框有一半在圓外
        } else {
            var ex = Math.abs(dir[0]) > 0.0001 ? hwT / Math.abs(dir[0]) : 1e9;
            var ey = Math.abs(dir[1]) > 0.0001 ? hhT / Math.abs(dir[1]) : 1e9;
            extent = Math.min(ex, ey);
        }
        var C = [S[0] + dir[0] * (gap + extent), S[1] + dir[1] * (gap + extent)];
        t.translate(C[0] - cx, C[1] - cy);

        if (o.badge) {
            var circ = g.pathItems.ellipse(C[1] + R, C[0] - R, R * 2, R * 2);
            circ.name = "CMF_Badge";
            circ.filled = true;
            circ.fillColor = hexToColor(o.badgeColor, d);
            if (bsw > 0) {
                circ.stroked = true;
                circ.strokeWidth = bsw;
                circ.strokeColor = hexToColor(o.badgeStrokeColor, d);
                circ.strokeDashes = [];
            } else {
                circ.stroked = false;
            }
            circ.move(t, ElementPlacement.PLACEAFTER);
        }
        return g;
    }

    // 拆掉舊標註，用同一條線重建；沒指定 key 時保留原本對應的物件
    function rebuild(d, g, num, styleStr, key) {
        if (key === undefined) key = noteParts(g).key;
        var line = findChild(g, "CMF_Line");
        if (!line) return null;
        var end = findChild(g, "CMF_End");
        if (end && end.typename === "PathItem" && end.closed && end.pathPoints.length === 3) {
            var tip = end.pathPoints[0].anchor; // 還原箭頭尖端位置
            var last = line.pathPoints[line.pathPoints.length - 1];
            last.anchor = tip; last.leftDirection = tip; last.rightDirection = tip;
        }
        var locked = g.locked, hidden = g.hidden, saved = openAncestors(g);
        try {
            var container = g.parent; // 留在原本的圖層或群組裡
            line.move(g, ElementPlacement.PLACEBEFORE);
            g.remove();
            var ng = buildCallout(d, line, num, styleStr, container, key);
            keepState(ng, locked, hidden);
            return ng;
        } finally { restoreTemp(saved); }
    }

    // 一個一個重建，有一個失敗也繼續做下一個；styleOf(note) 回傳要用的樣式
    function rebuildMany(d, cs, styleOf) {
        var r = { made: [], failed: 0, why: "" };
        for (var i = 0; i < cs.length; i++) {
            try {
                var p = noteParts(cs[i]), ng = rebuild(d, cs[i], p.num, styleOf(p));
                if (ng) r.made.push(ng); else r.failed++;
            } catch (e) { r.failed++; if (!r.why) r.why = e.message; }
        }
        return r;
    }

    function failNote(r) { return r.failed ? "（" + r.failed + " 個無法更新" + (r.why ? "：" + r.why : "") + "）" : ""; }

    function nextNumber(d, o) {
        return o.startMode === "auto" ? maxNumber(d) + 1 : parseInt(o.startNumber, 10) || 1;
    }

    // ---------- 互動新增 ----------
    // 面板按「新增」後切到鋼筆工具，鋼筆本身就有預覽線。
    // 面板每 200ms 呼叫 pollAdd，依線條模式判斷何時完成：
    //   straight：2 個錨點完成
    //   elbow   ：3 個錨點完成
    //   free    ：一直加點，按「完成這條線」/ 快捷鍵，或路徑被取消選取時完成
    //
    // 自由折線畫線過程中「完全不修改路徑」（改名稱等動作會干擾鋼筆），
    // 改用第一個錨點的座標當指紋來辨認這條線。
    var pendingFP = null; // { x, y, layer }

    function selectToolSafe(name) {
        try { app.selectTool(name); return true; } catch (e) { return false; }
    }

    function matchFP(it) {
        if (!pendingFP || !it) return false;
        try {
            if (it.typename !== "PathItem" || it.closed || it.pathPoints.length < 1) return false;
            var a = it.pathPoints[0].anchor;
            return Math.abs(a[0] - pendingFP.x) < 0.01 && Math.abs(a[1] - pendingFP.y) < 0.01;
        } catch (e) { return false; }
    }

    // 找回正在畫的那條線：先看目前選取，再從圖層最上層往下找
    function findPending(d) {
        if (!pendingFP) return null;
        var sel = d.selection, i;
        if (sel && sel.length) for (i = 0; i < sel.length; i++) if (matchFP(sel[i])) return sel[i];
        try {
            var items = pendingFP.layer.pathItems, n = Math.min(items.length, 60);
            for (i = 0; i < n; i++) if (matchFP(items[i]) && !isCallout(items[i].parent)) return items[i];
        } catch (e) {}
        return null;
    }

    function finalize(d, path, styleStr) {
        var o = parse(styleStr);
        reversePath(path); // 鋼筆第一點是目標點，buildCallout 需要第一點是編號端
        var num = nextNumber(d, o);
        buildCallout(d, path, num, styleStr, getLayer(d, true));
        d.selection = null; // 取消選取，鋼筆下一次點擊會開始新路徑
        return num;
    }

    function doneRes(num) { return res(true, "已新增標註 " + num, num + 1, "done"); }

    function beginAdd() {
        try {
            var d = getDoc();
            pendingFP = null;
            d.selection = null; // 清空選取，之後出現的選取路徑就一定是新畫的
            var ok = selectToolSafe("Adobe Pen Tool");
            return res(true, ok ? "" : "請按 P 切換到鋼筆工具", null, "ready");
        } catch (e) { return res(false, e.message); }
    }

    function pollAdd(styleStr) {
        try {
            if (app.documents.length === 0) { pendingFP = null; return res(true, "", null, "idle"); }
            var d = app.activeDocument, o = parse(styleStr), mode = o.lineMode || "straight";
            var sel = d.selection;

            if (mode === "free" && pendingFP) {
                // 還選取著同一條線 = 還在畫
                if (sel && sel.length === 1 && matchFP(sel[0])) return res(true, "", null, "drawing");
                // 不再被選取 = Ctrl/⌘ + 點空白處結束了
                var p = findPending(d);
                pendingFP = null;
                if (p && p.pathPoints.length >= 2) return doneRes(finalize(d, p, styleStr));
                return res(true, "", null, "idle");
            }

            if (!sel || sel.length !== 1) return res(true, "", null, "idle");
            var it = sel[0];
            if (it.typename !== "PathItem" || it.closed || isCallout(it.parent)) return res(true, "", null, "idle");
            var cnt = it.pathPoints.length;

            if (mode === "free") {
                if (cnt >= 2) {
                    var a = it.pathPoints[0].anchor;
                    pendingFP = { x: a[0], y: a[1], layer: it.layer };
                }
                return res(true, "", null, "drawing");
            }

            var need = mode === "elbow" ? 3 : 2;
            if (cnt < need) return res(true, "", null, "drawing");
            return doneRes(finalize(d, it, styleStr));
        } catch (e) { pendingFP = null; return res(false, e.message, null, "error"); }
    }

    // 自由折線：「完成這條線」按鈕 / 快捷鍵
    function finishFree(styleStr) {
        try {
            var d = getDoc(), p = findPending(d), sel = d.selection;
            pendingFP = null;
            // 輪詢還沒抓到指紋時，直接用目前選取的線
            if (!p && sel && sel.length === 1 && sel[0].typename === "PathItem" && !sel[0].closed && !isCallout(sel[0].parent)) p = sel[0];
            if (!p) return res(false, "沒有正在畫的線", null, "idle");
            if (p.pathPoints.length < 2) return res(false, "至少要點 2 個點", null, "drawing");
            return doneRes(finalize(d, p, styleStr));
        } catch (e) { pendingFP = null; return res(false, e.message, null, "error"); }
    }

    // keepTool：不切回選取工具；styleStr：自由折線結束時，把畫到一半的線也完成
    function endAdd(keepTool, styleStr) {
        try {
            var msg = "已結束新增", next = null;
            if (app.documents.length > 0) {
                var d = app.activeDocument, p = findPending(d);
                if (p && p.pathPoints.length >= 2 && styleStr) {
                    var num = finalize(d, p, styleStr);
                    msg = "已新增標註 " + num + "，結束新增";
                    next = num + 1;
                } else {
                    var sel = d.selection;
                    // 只點了第一下就結束時，刪掉留下的單一錨點
                    if (sel && sel.length === 1 && sel[0].typename === "PathItem" && sel[0].pathPoints.length < 2) sel[0].remove();
                }
            }
            pendingFP = null;
            if (!keepTool) selectToolSafe("Adobe Select Tool");
            return res(true, msg, next, "idle");
        } catch (e) { pendingFP = null; return res(false, e.message); }
    }

    // ---------- 轉換已畫好的線 ----------
    function convert(styleStr) {
        try {
            var d = getDoc(), o = parse(styleStr), sel = d.selection, lines = [], i;
            if (!sel || sel.length === 0) return res(false, "請先選取要轉換的線段");
            for (i = 0; i < sel.length; i++) {
                var it = sel[i];
                if (it.typename === "PathItem" && !it.closed && it.pathPoints.length >= 2 && !isCallout(it.parent)) lines.push(it);
            }
            if (lines.length === 0) return res(false, "選取中沒有可用的開放路徑（請用鋼筆或線段工具畫線）");

            lines.reverse(); // 選取順序是由上到下的堆疊順序，反轉後約等於畫線順序
            // 與互動新增一致：線的起點 = 目標點，終點 = 編號位置
            for (i = 0; i < lines.length; i++) reversePath(lines[i]);
            lines = sortByMode(lines, function (l) { return l.pathPoints[0].anchor; }, o.sortMode);

            var start = nextNumber(d, o);
            var layer = getLayer(d, true), made = [];
            for (i = 0; i < lines.length; i++) {
                made.push(buildCallout(d, lines[i], o.sameNumber ? start : start + i, styleStr, layer));
            }
            d.selection = made;
            var next = o.sameNumber ? start + 1 : start + lines.length;
            return res(true, "已建立 " + made.length + " 個標註", next);
        } catch (e) { return res(false, e.message); }
    }

    // ---------- 編號 ----------
    // 依位置重新編號；共用同一號碼的標註會維持共用
    function renumber(optStr) {
        try {
            var d = getDoc(), o = parse(optStr), cs = allCallouts(d), i;
            if (cs.length === 0) return res(false, "文件中沒有標註");
            if (o.sortMode === "order") {
                cs.sort(function (a, b) { return noteParts(a).num - noteParts(b).num; });
            } else {
                cs = sortByMode(cs, labelPt, o.sortMode);
            }
            var oldNums = [], newNums = [], next = 1;
            for (i = 0; i < cs.length; i++) {
                var p = noteParts(cs[i]), k, found = -1;
                for (k = 0; k < oldNums.length; k++) if (oldNums[k] === p.num) { found = k; break; }
                var nn;
                if (found >= 0) nn = newNums[found];
                else { nn = next++; oldNums.push(p.num); newNums.push(nn); }
                if (nn !== p.num) rebuild(d, cs[i], nn, p.style);
            }
            d.selection = null;
            return res(true, "已重新編號 1–" + (next - 1), next);
        } catch (e) { return res(false, e.message); }
    }

    function setNumber(optStr) {
        try {
            var d = getDoc(), o = parse(optStr), cs = selectedCallouts(d), made = [], i;
            if (cs.length === 0) return res(false, "請先選取標註");
            var n = parseInt(o.num, 10);
            if (!(n > 0)) return res(false, "編號需為正整數");
            // 這個編號已經對應某個物件時，選取的標註也改成對應它
            var others = calloutsWithNumber(d, n), key;
            for (i = 0; i < others.length; i++) {
                if (!contains(cs, others[i]) && noteParts(others[i]).key) { key = noteParts(others[i]).key; break; }
            }
            for (i = 0; i < cs.length; i++) made.push(rebuild(d, cs[i], n, noteParts(cs[i]).style, key));
            d.selection = made;
            return res(true, "已將 " + made.length + " 個標註設為 " + n);
        } catch (e) { return res(false, e.message); }
    }

    function docKey(d) {
        try { return d.fullName.fsName; } catch (e) { return d.name; }
    }

    // 編號表：列出每個編號、數量和對應的物件
    function listNumbers() {
        try {
            if (app.documents.length === 0) return '{"ok":true,"doc":null,"items":[]}';
            var d = app.activeDocument, cs = allCallouts(d), nums = [], counts = {}, keys = {}, i;
            for (i = 0; i < cs.length; i++) {
                var p = noteParts(cs[i]), n = p.num;
                if (counts[n] === undefined) { counts[n] = 0; keys[n] = ""; nums.push(n); }
                counts[n]++;
                if (!keys[n] && p.key) keys[n] = p.key;
            }
            nums.sort(function (a, b) { return a - b; });
            var parts = [];
            for (i = 0; i < nums.length; i++) {
                parts.push('{"num":' + nums[i] + ',"count":' + counts[nums[i]] + ',"key":"' + esc(keys[nums[i]]) + '"}');
            }
            return '{"ok":true,"doc":"' + esc(docKey(d)) + '","items":[' + parts.join(",") + ']}';
        } catch (e) { return res(false, e.message); }
    }

    // ---------- 對應 Excel 清單的物件 ----------
    // 只改備註和名稱，不重建標註
    function setKey(g, key) {
        var p = noteParts(g);
        if (p.key === key) return;
        try { g.layer.locked = false; } catch (e) {}
        g.note = makeNote(p.num, key, p.style);
        g.name = groupName(p.num, key);
    }

    // 物件名稱：換行和連續空白收成一個空白（跟面板比對名稱的方式一致）
    function cleanKey(k) { return String(k || "").replace(/\s+/g, " ").replace(/^ | $/g, ""); }

    // 一個物件只對應一個編號：其他編號原本對到它的，取消對應
    function applyLink(d, num, key) {
        var cs = allCallouts(d), hit = 0, freed = {}, freedList = [];
        for (var i = 0; i < cs.length; i++) {
            var p = noteParts(cs[i]);
            if (p.num === num) { setKey(cs[i], key); hit++; }
            else if (key && p.key === key) {
                setKey(cs[i], "");
                if (!freed[p.num]) { freed[p.num] = true; freedList.push(p.num); }
            }
        }
        return { hit: hit, freed: freedList };
    }

    // o = { num, key }；key 空字串 = 取消對應
    function setLink(optStr) {
        try {
            var d = getDoc(), o = parse(optStr), num = parseInt(o.num, 10), key = cleanKey(o.key);
            var r = applyLink(d, num, key);
            if (!r.hit) return res(false, "找不到編號 " + num);
            if (!key) return res(true, "編號 " + num + " 已取消對應");
            return res(true, "編號 " + num + " 對應「" + key + "」" +
                (r.freed.length ? "（編號 " + r.freed.join("、") + " 改為未對應）" : ""));
        } catch (e) { return res(false, e.message); }
    }

    // o = { links: [{ num, key }] }：一次設定多個（自動對應用）
    function setLinks(optStr) {
        try {
            var d = getDoc(), o = parse(optStr), n = 0;
            for (var i = 0; i < o.links.length; i++) {
                if (applyLink(d, parseInt(o.links[i].num, 10), cleanKey(o.links[i].key)).hit) n++;
            }
            return res(true, n ? "已自動對應 " + n + " 個編號" : "沒有可以自動對應的編號");
        } catch (e) { return res(false, e.message); }
    }

    // 選取某個編號的所有標註，並把畫面移過去
    function selectNumber(optStr) {
        try {
            var d = getDoc(), o = parse(optStr), cs = calloutsWithNumber(d, parseInt(o.num, 10));
            if (cs.length === 0) return res(false, "找不到編號 " + o.num);
            d.selection = cs;
            if (o.center) {
                var b = cs[0].geometricBounds;
                d.views[0].centerPoint = [(b[0] + b[2]) / 2, (b[1] + b[3]) / 2];
            }
            return res(true, "已選取編號 " + o.num + "（" + cs.length + " 個）");
        } catch (e) { return res(false, e.message); }
    }

    // 編號表改號：目標號碼已存在時，兩個號碼互換
    function changeNumber(optStr) {
        try {
            var d = getDoc(), o = parse(optStr), from = parseInt(o.from, 10), to = parseInt(o.to, 10), i;
            if (!(to > 0)) return res(false, "編號需為正整數");
            if (from === to) return res(true, "");
            var a = calloutsWithNumber(d, from), b = calloutsWithNumber(d, to);
            if (a.length === 0) return res(false, "找不到編號 " + from);
            for (i = 0; i < a.length; i++) rebuild(d, a[i], to, noteParts(a[i]).style);
            for (i = 0; i < b.length; i++) rebuild(d, b[i], from, noteParts(b[i]).style);
            d.selection = null;
            return res(true, b.length ? from + " 與 " + to + " 已互換" : "已將 " + from + " 改為 " + to);
        } catch (e) { return res(false, e.message); }
    }

    // ---------- 樣式 ----------
    // 同步全部：文件中所有標註套用目前樣式
    function syncAll(styleStr) {
        try {
            var d = getDoc(), cs = allCallouts(d);
            if (cs.length === 0) return res(true, "文件中沒有標註");
            var sel = selectedCallouts(d);
            var r = rebuildMany(d, cs, function () { return styleStr; });
            if (sel.length === 0) d.selection = null;
            return res(r.made.length > 0, "已同步 " + r.made.length + " 個標註的樣式" + failNote(r));
        } catch (e) { return res(false, e.message); }
    }

    // 套用到選取的標註
    function restyle(styleStr) {
        try {
            var d = getDoc(), cs = selectedCallouts(d);
            if (cs.length === 0) return res(false, "請先選取標註");
            var r = rebuildMany(d, cs, function () { return styleStr; });
            d.selection = r.made;
            return res(r.made.length > 0, "已更新 " + r.made.length + " 個標註的樣式" + failNote(r));
        } catch (e) { return res(false, e.message); }
    }

    // 移動過線條錨點後，用原本樣式重新排版
    function relayout() {
        try {
            var d = getDoc(), cs = selectedCallouts(d);
            if (cs.length === 0) return res(false, "請先選取要重新排版的標註");
            var r = rebuildMany(d, cs, function (p) { return p.style; });
            d.selection = r.made;
            return res(r.made.length > 0, "已重新排版 " + r.made.length + " 個標註" + failNote(r));
        } catch (e) { return res(false, e.message); }
    }

    function selectAll() {
        try {
            var d = getDoc(), cs = allCallouts(d);
            if (cs.length === 0) return res(false, "文件中沒有標註");
            d.selection = cs;
            return res(true, "已選取 " + cs.length + " 個標註");
        } catch (e) { return res(false, e.message); }
    }

    function toggleLayer() {
        try {
            var L = getLayer(getDoc(), false);
            if (!L) return res(false, "還沒有「" + LAYER_NAME + "」圖層");
            L.visible = !L.visible;
            return res(true, L.visible ? "標註圖層已顯示" : "標註圖層已隱藏");
        } catch (e) { return res(false, e.message); }
    }

    // ---------- 取色 ----------
    // CEP 面板內建的滴管無法吸取面板以外的畫面，改成「點畫布上的物件取它的顏色」
    function to2(v) {
        v = Math.max(0, Math.min(255, Math.round(v)));
        return (v < 16 ? "0" : "") + v.toString(16);
    }

    function colorToHex(c) {
        if (!c) return null;
        var t = c.typename;
        if (t === "RGBColor") return "#" + to2(c.red) + to2(c.green) + to2(c.blue);
        if (t === "CMYKColor") {
            var k = 1 - c.black / 100;
            return "#" + to2(255 * (1 - c.cyan / 100) * k) + to2(255 * (1 - c.magenta / 100) * k) + to2(255 * (1 - c.yellow / 100) * k);
        }
        if (t === "GrayColor") { var g = 255 * (1 - c.gray / 100); return "#" + to2(g) + to2(g) + to2(g); }
        if (t === "SpotColor") {
            var base = colorToHex(c.spot.color);
            if (!base) return null;
            var tint = c.tint / 100, out = "#";
            for (var i = 0; i < 3; i++) out += to2(255 - (255 - parseInt(base.substr(1 + i * 2, 2), 16)) * tint);
            return out;
        }
        if (t === "GradientColor") return colorToHex(c.gradient.gradientStops[0].color);
        return null; // 無色、圖樣
    }

    // prefer："fill" 或 "stroke"，找不到時改用另一個
    function itemColor(it, prefer) {
        var t = it.typename, i, c;
        if (t === "TextFrame" || t === "TextRange") {
            var tr = t === "TextFrame" ? it.textRange : it;
            try { return colorToHex(tr.characterAttributes.fillColor); } catch (e) { return null; }
        }
        if (t === "PathItem") {
            var f = it.filled ? colorToHex(it.fillColor) : null;
            var s = it.stroked ? colorToHex(it.strokeColor) : null;
            return prefer === "stroke" ? (s || f) : (f || s);
        }
        if (t === "CompoundPathItem" && it.pathItems.length) return itemColor(it.pathItems[0], prefer);
        if (t === "GroupItem") {
            for (i = 0; i < it.pageItems.length; i++) { c = itemColor(it.pageItems[i], prefer); if (c) return c; }
        }
        return null;
    }

    function beginPick() {
        try {
            var d = getDoc();
            d.selection = null;
            selectToolSafe("Adobe Direct Select Tool"); // 直接選取，點群組內的物件也能取到單一顏色
            return res(true, "", null, "ready");
        } catch (e) { return res(false, e.message); }
    }

    function pollPick(optStr) {
        try {
            if (app.documents.length === 0) return res(true, "", null, "idle");
            var d = app.activeDocument, o = parse(optStr), sel = d.selection;
            if (!sel) return res(true, "", null, "idle");
            var isText = sel.typename === "TextRange"; // 用文字工具選取文字時，selection 是 TextRange 而不是陣列
            if (!isText && sel.length === 0) return res(true, "", null, "idle");
            var it = isText ? sel : sel[0];
            var hex = itemColor(it, o.prefer);
            d.selection = null;
            if (!hex) return res(false, "這個物件沒有可取用的顏色（無色或圖樣）", null, "done");
            return '{"ok":true,"msg":"已取色 ' + hex + '","hex":"' + hex + '","state":"done"}';
        } catch (e) { return res(false, e.message, null, "done"); }
    }

    function endPick() {
        selectToolSafe("Adobe Select Tool");
        return res(true, "", null, "idle");
    }

    // ---------- 字體清單 ----------
    // 回傳 [家族, 字重, PostScript 名稱]；字體很多時需要幾秒，面板會快取結果
    function listFonts() {
        try {
            var f = app.textFonts, n = f.length, parts = [];
            for (var i = 0; i < n; i++) {
                var t = f[i];
                parts.push('["' + esc(t.family) + '","' + esc(t.style) + '","' + esc(t.name) + '"]');
            }
            return '{"ok":true,"msg":"已載入 ' + n + ' 個字體","fonts":[' + parts.join(",") + ']}';
        } catch (e) { return res(false, e.message); }
    }

    // ---------- 依物件尺寸縮放整份文件 ----------
    // 選一個物件、指定它的目標寬度或高度：物件、文字、線寬、效果和工作區一起等比縮放，
    // 看起來跟原本一樣，只有尺寸改變。範圍是物件所在的工作區，或全部工作區。
    // 寬高跟「變形」面板一樣：有勾「使用預視邊界」就含線寬，剪裁群組用遮色片的範圍。
    var UNITS = {
        Points: ["pt", 1], Picas: ["pc", 12], Inches: ["in", 72], Millimeters: ["mm", 72 / 25.4],
        Centimeters: ["cm", 72 / 2.54], Pixels: ["px", 1], Qs: ["Q", 72 / 25.4 / 4],
        Feet: ["ft", 864], Yards: ["yd", 2592], Meters: ["m", 72 / 0.0254]
    };
    var MAX_ARTBOARD = 16383; // Illustrator 工作區的上限 (pt)

    function docUnit(d) {
        var u = null;
        try { u = UNITS[String(d.rulerUnits).replace(/^.*\./, "")]; } catch (e) {}
        return u || UNITS.Points;
    }

    // 大型畫布文件（Illustrator 2019 以後）腳本裡的數值是實際尺寸的 1/scaleFactor
    function docScaleFactor(d) {
        var f = 1;
        try { f = Number(d.scaleFactor) || 1; } catch (e) {}
        return f;
    }

    function usePreviewBounds() {
        try { return app.preferences.getBooleanPreference("includeStrokeInBounds"); } catch (e) { return false; }
    }

    function clipPathOf(g) {
        for (var i = 0; i < g.pageItems.length; i++) {
            var c = g.pageItems[i];
            if (c.typename === "PathItem" && c.clipping) return c;
            if (c.typename === "CompoundPathItem" && c.pathItems.length && c.pathItems[0].clipping) return c;
        }
        return null;
    }

    // [左, 上, 右, 下]；剪裁群組用遮色片的範圍
    function boundsOf(it, preview) {
        if (it.typename === "GroupItem" && it.clipped) {
            var c = clipPathOf(it);
            if (c) return c.geometricBounds;
        }
        return preview ? it.visibleBounds : it.geometricBounds;
    }

    function unionBounds(items, preview) {
        var u = null;
        for (var i = 0; i < items.length; i++) {
            var b = boundsOf(items[i], preview);
            if (!u) { u = [b[0], b[1], b[2], b[3]]; continue; }
            if (b[0] < u[0]) u[0] = b[0];
            if (b[1] > u[1]) u[1] = b[1];
            if (b[2] > u[2]) u[2] = b[2];
            if (b[3] < u[3]) u[3] = b[3];
        }
        return u;
    }

    // 用文字工具點在文字裡時，selection 是 TextRange：改用它所在的文字框
    function selectedItems(d) {
        var sel = d.selection, out = [], i;
        if (!sel) return out;
        if (sel.typename === "TextRange") {
            try { for (i = 0; i < sel.story.textFrames.length; i++) out.push(sel.story.textFrames[i]); } catch (e) {}
            return out;
        }
        for (i = 0; i < sel.length; i++) out.push(sel[i]);
        return out;
    }

    // 重疊面積；pad 讓寬或高是 0 的物件（水平線、垂直線）也算得到
    function overlapArea(b, r, pad) {
        var w = Math.min(b[2] + pad, r[2]) - Math.max(b[0] - pad, r[0]);
        var h = Math.min(b[1] + pad, r[1]) - Math.max(b[3] - pad, r[3]);
        return w > 0 && h > 0 ? w * h : 0;
    }

    // 物件屬於重疊面積最大的工作區；不在任何工作區上 = -1
    function artboardOf(d, b) {
        var best = -1, most = 0;
        for (var i = 0; i < d.artboards.length; i++) {
            var a = overlapArea(b, d.artboards[i].artboardRect, 0.01);
            if (a > most) { most = a; best = i; }
        }
        return best;
    }

    // 圖層裡最上層的物件（群組裡的物件 → 整個群組）
    function topItem(it) {
        while (it.parent && it.parent.typename !== "Layer" && it.parent.typename !== "Document") it = it.parent;
        return it;
    }

    function collectTop(layers, out) {
        for (var i = 0; i < layers.length; i++) {
            var items = layers[i].pageItems;
            for (var k = 0; k < items.length; k++) out.push(items[k]);
            collectTop(layers[i].layers, out);
        }
    }

    // 暫時改屬性（解除鎖定、顯示），結束後依相反順序還原
    function setTemp(obj, prop, val, saved) {
        try {
            if (obj[prop] !== val) { saved.push([obj, prop, obj[prop]]); obj[prop] = val; }
        } catch (e) {}
    }

    function openLayers(layers, saved) {
        for (var i = 0; i < layers.length; i++) {
            setTemp(layers[i], "locked", false, saved);
            setTemp(layers[i], "visible", true, saved); // 隱藏圖層中的物件無法修改
            openLayers(layers[i].layers, saved);
        }
    }

    function setPref(name, val, saved) {
        try {
            var old = app.preferences.getBooleanPreference(name);
            if (old !== val) { app.preferences.setBooleanPreference(name, val); saved.push([name, old]); }
        } catch (e) {}
    }

    // 標準座標（跟 artboardRect 一致），結束後還原使用者原本的設定
    function useDocCoords() {
        try {
            var old = app.coordinateSystem;
            app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;
            return old;
        } catch (e) { return null; }
    }

    function restoreCoords(old) {
        if (old !== null) { try { app.coordinateSystem = old; } catch (e) {} }
    }

    function mapPt(p, origin, s) {
        return [origin[0] + (p[0] - origin[0]) * s, origin[1] + (p[1] - origin[1]) * s];
    }

    // 以 origin 為基準等比縮放：線寬、圖樣、漸層一起縮放，縮放後再把左上角對到正確位置
    function scaleItem(it, s, origin) {
        var b0 = it.geometricBounds, pct = s * 100;
        it.resize(pct, pct, true, true, true, true, pct, Transformation.TOPLEFT);
        var b1 = it.geometricBounds, to = mapPt([b0[0], b0[1]], origin, s);
        it.translate(to[0] - b1[0], to[1] - b1[1], true, true, true, true);
    }

    // 標註備註裡跟尺寸有關的樣式也跟著縮放，之後重新編號、重新排版才會維持縮放後的大小
    function scaleNote(g, s) {
        var p = noteParts(g);
        var re = /"(fontSize|badgePadding|badgeStrokeWidth|lineWidth|endSize|styleScale|gap)"\s*:\s*(-?[0-9.]+(?:[eE][-+]?[0-9]+)?)/g;
        var style = p.style.replace(re, function (m, k, v) {
            return '"' + k + '":' + roundTo(Number(v) * s, 3);
        });
        // 舊的標註沒有 styleScale（= 100%）：補上，之後重新產生時繪圖樣式的線寬才會維持縮放後的大小
        if (!/"styleScale"\s*:/.test(style)) style = style.replace(/\}\s*$/, ',"styleScale":' + roundTo(100 * s, 3) + "}");
        if (style !== p.style) g.note = makeNote(p.num, p.key, style);
    }

    function jsonNum(v) { return isFinite(v) ? String(v) : "0"; }

    function roundTo(v, n) { var f = Math.pow(10, n); return Math.round(v * f) / f; }

    // 面板顯示用：選取物件的寬高 (pt)、文件單位、所在的工作區
    function scaleInfo() {
        var coords = useDocCoords();
        try {
            if (app.documents.length === 0) return '{"ok":true,"doc":false}';
            var d = app.activeDocument, u = docUnit(d), f = docScaleFactor(d), preview = usePreviewBounds();
            var items = selectedItems(d), b = items.length ? unionBounds(items, preview) : null;
            var ab = b ? artboardOf(d, b) : d.artboards.getActiveArtboardIndex();
            var out = '{"ok":true,"doc":true,"unit":"' + u[0] + '","unitPt":' + jsonNum(u[1]) +
                ',"preview":' + (preview ? "true" : "false") + ',"sel":' + items.length + ',"artboards":' + d.artboards.length;
            if (b) out += ',"w":' + jsonNum((b[2] - b[0]) * f) + ',"h":' + jsonNum((b[1] - b[3]) * f);
            if (ab >= 0) out += ',"ab":' + ab + ',"abName":"' + esc(d.artboards[ab].name) + '"';
            return out + "}";
        } catch (e) { return res(false, e.message); }
        finally { restoreCoords(coords); }
    }

    // o = { side: "w" | "h" | "pct", value, scope: "artboard" | "all" }
    // value：w、h 是選取物件的目標尺寸 (pt)，pct 是百分比
    function scaleDoc(optStr) {
        var coords = useDocCoords(), saved = [], prefs = [], i, j;
        try {
            var d = getDoc(), o = parse(optStr), u = docUnit(d), f = docScaleFactor(d);
            var items = selectedItems(d), b = items.length ? unionBounds(items, usePreviewBounds()) : null;
            var s = Number(o.value) / 100;
            if (o.side !== "pct") {
                if (!b) return res(false, "請先選取一個物件");
                var cur = (o.side === "h" ? b[1] - b[3] : b[2] - b[0]) * f;
                if (!(cur > 0.001)) return res(false, "選取的物件" + (o.side === "h" ? "高度" : "寬度") + "是 0，無法依它縮放");
                s = Number(o.value) / cur;
            }
            if (!(s > 0) || !isFinite(s)) return res(false, "請輸入大於 0 的數字");
            if (Math.abs(s - 1) < 0.000001) return res(true, "尺寸已經一樣，不需要縮放");

            // 縮放哪些工作區、以哪裡為基準：單一工作區時它的左上角不動；全部時以所有工作區的中心為基準
            var all = o.scope === "all", boards = [], ab = -1, origin, r;
            if (all) {
                var ub = null;
                for (i = 0; i < d.artboards.length; i++) {
                    boards.push(i);
                    r = d.artboards[i].artboardRect;
                    ub = ub ? [Math.min(ub[0], r[0]), Math.max(ub[1], r[1]), Math.max(ub[2], r[2]), Math.min(ub[3], r[3])]
                        : [r[0], r[1], r[2], r[3]];
                }
                origin = [(ub[0] + ub[2]) / 2, (ub[1] + ub[3]) / 2];
            } else {
                ab = b ? artboardOf(d, b) : d.artboards.getActiveArtboardIndex();
                if (ab < 0) return res(false, "選取的物件不在任何工作區上，請改選「全部工作區」");
                boards.push(ab);
                r = d.artboards[ab].artboardRect;
                origin = [r[0], r[1]];
            }

            // 先確認工作區的新尺寸可以用，再開始改
            var oldRects = [], newRects = [];
            for (i = 0; i < boards.length; i++) {
                r = d.artboards[boards[i]].artboardRect;
                var tl = mapPt([r[0], r[1]], origin, s), br = mapPt([r[2], r[3]], origin, s);
                var nr = [tl[0], tl[1], br[0], br[1]];
                if (nr[2] - nr[0] > MAX_ARTBOARD + 0.5 || nr[1] - nr[3] > MAX_ARTBOARD + 0.5) {
                    return res(false, "放大後工作區會超過 Illustrator 的上限 " + roundTo(MAX_ARTBOARD * f / u[1], 1) + " " + u[0]);
                }
                if (nr[2] - nr[0] < 1 || nr[1] - nr[3] < 1) return res(false, "縮小後工作區太小了");
                oldRects.push(r);
                newRects.push(nr);
            }

            // 縮放的物件：圖層裡最上層的物件。單一工作區時只取屬於這個工作區的，選取的物件一定算進去
            setPref("scaleLineWeight", true, prefs); // 縮放線條和效果
            setPref("scaleCorners", true, prefs);    // 縮放圓角
            openLayers(d.layers, saved);
            var every = d.pageItems, n = every.length;
            for (i = 0; i < n; i++) {
                setTemp(every[i], "locked", false, saved);
                setTemp(every[i], "hidden", false, saved);
            }
            var units = [], picked = [];
            for (i = 0; i < items.length; i++) {
                var t = topItem(items[i]);
                if (!contains(picked, t)) picked.push(t);
            }
            var inScope = function (top) {
                return all || contains(picked, top) || artboardOf(d, boundsOf(top, false)) === ab;
            };
            collectTop(d.layers, units);
            if (!all) {
                var keep = [];
                for (i = 0; i < units.length; i++) if (inScope(units[i])) keep.push(units[i]);
                units = keep;
            }
            // 先判斷標註、尺寸在不在範圍內（縮放後位置就變了）
            var callouts = allCallouts(d), scaledCallouts = [];
            for (i = 0; i < callouts.length; i++) if (inScope(topItem(callouts[i]))) scaledCallouts.push(callouts[i]);
            var dims = allDims(d), scaledDims = [];
            for (i = 0; i < dims.length; i++) if (inScope(topItem(dims[i]))) scaledDims.push(dims[i]);

            var done = [];
            for (i = 0; i < boards.length; i++) {
                var A = d.artboards[boards[i]];
                try { A.artboardRect = newRects[i]; done.push(i); }
                catch (eA) {
                    for (j = done.length - 1; j >= 0; j--) {
                        try { d.artboards[boards[done[j]]].artboardRect = oldRects[done[j]]; } catch (eB) {}
                    }
                    return res(false, "工作區「" + A.name + "」無法縮放到這個大小（" + eA.message + "）");
                }
            }

            var failed = 0;
            for (i = 0; i < units.length; i++) {
                try { scaleItem(units[i], s, origin); } catch (eI) { failed++; }
            }
            for (i = 0; i < scaledCallouts.length; i++) {
                try { scaleNote(scaledCallouts[i], s); } catch (eN) {}
            }
            // 尺寸重建：數字改成縮放後的尺寸，字級、線寬跟著縮放
            for (i = 0; i < scaledDims.length; i++) {
                try { rebuildDim(d, scaledDims[i], scaleDimStyle(dimNote(scaledDims[i]).style, s)); } catch (eD) {}
            }

            // 畫面跟著縮放，看起來跟縮放前一樣
            try {
                var v = d.views[0], c = v.centerPoint;
                v.zoom = Math.max(0.0313, Math.min(640, v.zoom / s));
                v.centerPoint = mapPt(c, origin, s);
            } catch (eV) {}

            var msg = "已縮放 " + roundTo(s * 100, 2) + "%：" +
                (all ? "全部 " + boards.length + " 個工作區" : "工作區「" + d.artboards[ab].name + "」") +
                "和 " + (units.length - failed) + " 個物件";
            if (failed) msg += "（" + failed + " 個無法縮放）";
            if (!all) {
                var hits = [];
                for (i = 0; i < d.artboards.length; i++) {
                    if (i === ab) continue;
                    var other = d.artboards[i].artboardRect;
                    if (overlapArea(other, newRects[0], 0) > 1 && overlapArea(other, oldRects[0], 0) <= 1) {
                        hits.push("「" + d.artboards[i].name + "」");
                    }
                }
                if (hits.length) msg += "；跟工作區" + hits.join("、") + "重疊了";
            }
            // factor：面板縮放全部工作區後，標註樣式也用同一個倍率縮放
            return '{"ok":true,"msg":"' + esc(msg) + '","factor":' + jsonNum(s) + "}";
        } catch (e) { return res(false, e.message); }
        finally {
            for (i = saved.length - 1; i >= 0; i--) { try { saved[i][0][saved[i][1]] = saved[i][2]; } catch (eR) {} }
            for (i = 0; i < prefs.length; i++) { try { app.preferences.setBooleanPreference(prefs[i][0], prefs[i][1]); } catch (eP) {} }
            restoreCoords(coords);
        }
    }

    // ---------- 交換兩個物件的位置和大小 ----------
    // 選兩個物件（圖片、群組都可以）按一下：A 換到 B 的位置和大小，B 換到 A 的。
    // 保持比例時放進對方的外框裡、置中，不會變形；不保持比例時，寬高完全換成對方的。
    // 外框跟「變形」面板一樣（預視邊界、剪裁群組用遮色片的範圍）。圖層和上下順序不變。
    function isInside(it, outer) {
        for (var p = it.parent; p && p.typename !== "Layer" && p.typename !== "Document"; p = p.parent) {
            if (p === outer) return true;
        }
        return false;
    }

    function fitInto(it, box, keepRatio, preview) {
        var b = boundsOf(it, preview), w = b[2] - b[0], h = b[1] - b[3];
        var sx = w > 0.001 ? (box[2] - box[0]) / w : null, sy = h > 0.001 ? (box[1] - box[3]) / h : null;
        if (keepRatio || sx === null || sy === null) {
            // 寬或高是 0 的物件（水平線、垂直線）只看另一邊
            var k = sx === null ? sy : sy === null ? sx : Math.min(sx, sy);
            sx = sy = k === null ? 1 : k;
        }
        if (Math.abs(sx - 1) > 0.000001 || Math.abs(sy - 1) > 0.000001) {
            it.resize(sx * 100, sy * 100, true, true, true, true, Math.sqrt(sx * sy) * 100, Transformation.CENTER);
        }
        var nb = boundsOf(it, preview);
        it.translate((box[0] + box[2]) / 2 - (nb[0] + nb[2]) / 2, (box[1] + box[3]) / 2 - (nb[1] + nb[3]) / 2, true, true, true, true);
    }

    // o = { keepRatio }
    function swapItems(optStr) {
        try {
            var d = getDoc(), o = parse(optStr), items = selectedItems(d);
            if (items.length !== 2) {
                var hint = items.length === 1 && items[0].typename === "GroupItem" ? "；群組裡的物件請用群組選取或直接選取工具選" : "";
                return res(false, "請選取兩個物件（目前選了 " + items.length + " 個）" + hint);
            }
            var a = items[0], b = items[1];
            if (isInside(a, b) || isInside(b, a)) return res(false, "選取的兩個物件一個在另一個的群組裡，請改選兩個各自獨立的物件");
            var preview = usePreviewBounds(), keep = o.keepRatio !== false;
            var ba = boundsOf(a, preview), bb = boundsOf(b, preview);
            fitInto(a, bb, keep, preview);
            fitInto(b, ba, keep, preview);
            var p = (ba[2] - ba[0]) * (bb[1] - bb[3]), q = (bb[2] - bb[0]) * (ba[1] - ba[3]);
            var sameRatio = Math.abs(p - q) <= 0.005 * Math.max(Math.abs(p), Math.abs(q));
            return res(true, "已交換兩個物件的位置和大小" + (keep && !sameRatio ? "（比例不同，保持比例放進對方的範圍、置中）" : ""));
        } catch (e) { return res(false, e.message); }
    }

    // ---------- 尺寸標註 ----------
    // 選取物件後標寬度、高度（外框尺寸），或圓的直徑 Ø、圓弧和圓角的半徑 R。
    // 每個尺寸是一個群組，放在「CMF Dimensions」圖層，備註存著：CMF_DIM|幾何 JSON|樣式 JSON
    //   寬高：{ t:"lin", p1, p2 量測點（延伸線起點）, a 量測方向, n 尺寸線在哪一側, off 尺寸線離量測點的距離 }
    //   直徑、半徑：{ t:"dia" | "rad", c 圓心, r 半徑, u 標註方向（圓心 → 箭頭指的點）, out = 1 一律用引線,
    //     x 引線比樣式的「引線」多拉的長度（同心圓的內圈）, q 同樣大小的圓有幾個 }
    //     舊版的尺寸沒有 x，引線長度記在 off
    //   ref：建立時 DIM_Line 頭尾兩個錨點的位置
    //   z：量的物件較長的一邊（新增時「依物件大小自動調整樣式」算出來的依據，跟著尺寸一起移動、縮放）
    // 重建時比對 DIM_Line 現在的位置，算出尺寸被移動、等比縮放、旋轉了多少，套用到記錄的點上，
    // 所以搬過的尺寸重建後還在原地，縮放整份文件後數字也會跟著變。
    var DIM_LAYER = "CMF Dimensions";
    var DIM_TAG = "CMF_DIM";
    var DIM_UNITS = { mm: ["mm", 72 / 25.4], cm: ["cm", 72 / 2.54], "in": ["in", 72], pt: ["pt", 1], px: ["px", 1] };
    var DIA_SYMS = { slash: "\u00D8", phi: "\u03C6", sign: "\u2300" };
    var DIA = DIA_SYMS.slash;
    var MAX_SEGS = 4000;  // 找圓弧時最多看幾段曲線，避免複雜的圖跑太久
    var MAX_ARCS = 20;    // 一個物件最多標幾種半徑

    function isDim(it) {
        return it && it.typename === "GroupItem" && it.note && String(it.note).indexOf(DIM_TAG + "|") === 0;
    }

    function allDims(d) {
        var out = [], gs = d.groupItems;
        for (var i = 0; i < gs.length; i++) if (isDim(gs[i])) out.push(gs[i]);
        return out;
    }

    // 往上找第一個符合 test 的群組（自己也算）
    function ownerOf(it, test) {
        while (it && it.typename !== "Layer" && it.typename !== "Document") {
            if (test(it)) return it;
            it = it.parent;
        }
        return null;
    }

    function selectedDims(d) {
        var out = [], items = selectedItems(d);
        for (var i = 0; i < items.length; i++) {
            var g = ownerOf(items[i], isDim);
            if (g && !contains(out, g)) out.push(g);
        }
        return out;
    }

    // 要量的物件：選取中扣掉尺寸和標註
    function measuredItems(d) {
        var out = [], items = selectedItems(d);
        for (var i = 0; i < items.length; i++) {
            if (!ownerOf(items[i], isDim) && !ownerOf(items[i], isCallout)) out.push(items[i]);
        }
        return out;
    }

    function dimNote(g) {
        var s = String(g.note), i = s.indexOf("|"), j = s.indexOf("|", i + 1);
        return { geom: s.substring(i + 1, j), style: s.substring(j + 1) };
    }

    // ---- 向量 ----
    function vsub(p, q) { return [p[0] - q[0], p[1] - q[1]]; }
    function vdot(p, q) { return p[0] * q[0] + p[1] * q[1]; }
    function vlen(p) { return Math.sqrt(p[0] * p[0] + p[1] * p[1]); }
    function along(p, u, s) { return [p[0] + u[0] * s, p[1] + u[1] * s]; }
    function dirOf(deg) { var a = deg * Math.PI / 180; return [Math.cos(a), Math.sin(a)]; }

    // 線的方向 → 文字角度：(-90°, 90°]，文字由左往右、由下往上讀
    function readAngle(u) {
        var a = Math.atan2(u[1], u[0]) * 180 / Math.PI;
        if (a > 90.01) a -= 180;
        else if (a <= -89.99) a += 180;
        return a;
    }

    // ---- 數字 ----
    // 小數最多 dec 位，去掉尾數的 0
    function fmtNum(v, dec) {
        var s = v.toFixed(dec);
        if (s.indexOf(".") >= 0) s = s.replace(/0+$/, "").replace(/\.$/, "");
        return s === "-0" ? "0" : s;
    }

    // 圖面比例 1:2 → 實際尺寸 = 圖上尺寸 × 2；只寫一個數字 n = 1:n
    function ratioOf(text) {
        var m = /^\s*(\d*\.?\d+)\s*[:\uFF1A\/]\s*(\d*\.?\d+)\s*$/.exec(String(text || ""));
        if (m) {
            var a = parseFloat(m[1]), b = parseFloat(m[2]);
            return a > 0 && b > 0 ? b / a : 1;
        }
        var n = parseFloat(text);
        return n > 0 ? n : 1;
    }

    // 圖上長度（腳本座標）→ 標示的文字：前綴 + 數量（4×）+ 符號（Ø、R）+ 數字 + 單位 + 後綴
    function dimLabel(d, st, len, sym, q) {
        var u = st.unit === "doc" || !DIM_UNITS[st.unit] ? docUnit(d) : DIM_UNITS[st.unit];
        var dec = Math.max(0, Math.min(4, parseInt(st.decimals, 10) || 0));
        var v = len * docScaleFactor(d) * ratioOf(st.ratio) / u[1];
        var s = (sym || "") + fmtNum(v, dec) + (st.showUnit ? u[0] : "");
        if (q > 1 && st.count === "x") s = q + "\u00D7" + s;
        else if (q > 1 && st.count === "dash") s = q + "-" + s;
        return String(st.prefix || "") + s + String(st.suffix || "");
    }

    function diaSym(st) { return DIA_SYMS[st.diaSym] || DIA; }

    // 實心箭頭的尖端比線細，線要縮短一點才不會凸出去
    function arrowLike(kind) { return kind === "arrow" || kind === "open"; }
    function arrowTrim(kind, L) { return kind === "arrow" ? L * 0.9 : 0; }

    // ---- 畫 ----
    function dimPath(g, pts, st, col, name) {
        var p = g.pathItems.add();
        p.setEntirePath(pts);
        p.name = name;
        p.filled = false;
        p.stroked = true;
        p.strokeColor = col;
        p.strokeWidth = Number(st.lineWidth);
        p.strokeDashes = [];
        p.strokeCap = StrokeCap.BUTTENDCAP;
        p.strokeJoin = StrokeJoin.MITERENDJOIN;
        return p;
    }

    // 端點：tip 是尖端，u 是指向尖端的方向（斜線用尺寸線的方向）
    function dimEnd(g, tip, u, st, col, kind) {
        var L = Number(st.endSize);
        if (!(L > 0) || kind === "none") return;
        if (kind === "dot") {
            var r = L * 0.3;
            var dot = g.pathItems.ellipse(tip[1] + r, tip[0] - r, r * 2, r * 2);
            dot.name = "DIM_End";
            dot.stroked = false; dot.filled = true; dot.fillColor = col;
        } else if (kind === "tick") {
            // 建築圖的斜線：跟尺寸線夾 45°
            var v = [(u[0] - u[1]) * Math.SQRT1_2, (u[0] + u[1]) * Math.SQRT1_2];
            var tk = dimPath(g, [along(tip, v, -L / 2), along(tip, v, L / 2)], st, col, "DIM_End");
            tk.strokeWidth = Number(st.lineWidth) * 2;
        } else if (kind === "open") {
            // 開放箭頭：兩條線，夾角約 30°
            var ow = L * 0.27, OB = along(tip, u, -L), ox = -u[1] * ow, oy = u[0] * ow;
            var op = dimPath(g, [[OB[0] + ox, OB[1] + oy], [tip[0], tip[1]], [OB[0] - ox, OB[1] - oy]], st, col, "DIM_End");
            op.strokeJoin = StrokeJoin.ROUNDENDJOIN; // 尖角接合會讓尖端凸出去
        } else {
            var hw = L * 0.18, B = along(tip, u, -L), px = -u[1] * hw, py = u[0] * hw;
            var tri = g.pathItems.add();
            tri.setEntirePath([[tip[0], tip[1]], [B[0] + px, B[1] + py], [B[0] - px, B[1] - py]]);
            tri.closed = true;
            tri.name = "DIM_End";
            tri.stroked = false; tri.filled = true; tri.fillColor = col;
        }
    }

    // 建立文字，量好字形的大小（還沒旋轉）：hw 半寬、hh 半高
    function dimText(d, g, label, st) {
        var t = g.textFrames.add();
        t.name = "DIM_Text";
        t.contents = label;
        var ca = t.textRange.characterAttributes;
        ca.size = Number(st.fontSize);
        if (st.fontName) {
            try { ca.textFont = app.textFonts.getByName(st.fontName); } catch (e) { /* 找不到字體就用預設 */ }
        }
        ca.fillColor = hexToColor(st.textColor, d);
        t.textRange.paragraphAttributes.justification = Justification.CENTER;
        var b = glyphBounds(t);
        return { t: t, b: b, hw: (b[2] - b[0]) / 2, hh: (b[1] - b[3]) / 2, d: d, g: g, st: st };
    }

    // 文字離線：文字跟線之間空出來的距離，從線的邊緣算（線寬的一半要加上去）
    function besideGap(st) { return Number(st.textGap) + Number(st.lineWidth) / 2; }

    // 斷開留空：文字放在線中間時，文字兩邊跟線頭之間空出來的距離；舊的樣式沒有，用「文字離線」
    function breakGap(st) {
        var v = Number(st.breakGap);
        return st.breakGap !== undefined && st.breakGap !== null && st.breakGap !== "" && v >= 0 ? v : Number(st.textGap);
    }

    // 文字放在 mid 的 n 那一側、離線 gap；aligned = 沿著線（文字跟 n 垂直）
    function textCenter(mid, n, tx, aligned, gap) {
        var ext = aligned ? tx.hh : tx.hw * Math.abs(n[0]) + tx.hh * Math.abs(n[1]);
        return along(mid, n, gap + ext);
    }

    // 文字在 v 方向佔的長度的一半
    function textAlong(tx, v, aligned) {
        return aligned ? tx.hw : tx.hw * Math.abs(v[0]) + tx.hh * Math.abs(v[1]);
    }

    // 旋轉後重新量字形，把字形中心移到 c。文字移到群組最上層；有底色時在文字後面墊一塊
    function placeText(tx, c, angle) {
        var b = tx.b, st = tx.st;
        if (Math.abs(angle) > 0.01) {
            tx.t.rotate(angle);
            b = glyphBounds(tx.t);
        }
        tx.t.translate(c[0] - (b[0] + b[2]) / 2, c[1] - (b[1] + b[3]) / 2);
        tx.t.move(tx.g, ElementPlacement.PLACEATBEGINNING);
        if (st.textBg !== "box") return;
        // 留的邊不超過文字離線、斷開留空，才不會蓋到尺寸線
        var pad = Math.max(0, Math.min(Number(st.fontSize) * 0.2, Math.min(Number(st.textGap), breakGap(st)) * 0.75));
        var w = tx.hw + pad, h = tx.hh + pad;
        var bg = tx.g.pathItems.rectangle(c[1] + h, c[0] - w, w * 2, h * 2);
        bg.name = "DIM_TextBg";
        bg.stroked = false;
        bg.filled = true;
        bg.fillColor = hexToColor(st.textBgColor || "#ffffff", tx.d);
        if (Math.abs(angle) > 0.01) bg.rotate(angle);
        bg.move(tx.t, ElementPlacement.PLACEAFTER);
    }

    // 文字放在線中間、把線斷開：p → q（方向 a）的線在 mid 前後各空 half，兩邊剩不到 minLen 時不斷開、回傳 null
    // 前半段叫 DIM_Line（重建時用它的頭尾判斷尺寸被移動了多少）
    function splitLine(g, p, q, a, mid, half, minLen, st, col) {
        if (vdot(vsub(mid, p), a) - half < minLen || vdot(vsub(q, mid), a) - half < minLen) return null;
        var first = dimPath(g, [p, along(mid, a, -half)], st, col, "DIM_Line");
        dimPath(g, [along(mid, a, half), q], st, col, "DIM_Line2");
        return first;
    }

    // 寬度、高度：回傳 DIM_Line
    function buildLinear(d, g, gm, st) {
        var a = gm.a, n = gm.n, col = hexToColor(st.lineColor, d), i;
        var L = Number(st.endSize), kind = st.endStyle, gapE = Number(st.extGap), over = Number(st.extOver);
        var gap = besideGap(st), tr = arrowTrim(kind, L);
        var level = Math.max(vdot(gm.p1, n), vdot(gm.p2, n)) + Number(gm.off);
        var P = [gm.p1, gm.p2], Q = [];
        for (i = 0; i < 2; i++) {
            var dist = level - vdot(P[i], n);
            Q.push(along(P[i], n, dist));
            if (dist - gapE > 0.1 && st.extLine !== "none") {
                var ext = dimPath(g, [along(P[i], n, gapE), along(Q[i], n, over)], st, col, "DIM_Ext");
                if (st.extLine === "thin") ext.strokeWidth = Number(st.lineWidth) / 2;
            }
        }
        if (vdot(vsub(Q[1], Q[0]), a) < 0) Q.reverse();
        var Q1 = Q[0], Q2 = Q[1], len = vlen(vsub(Q2, Q1));

        // 箭頭放在延伸線內側往外指，或外側往內指；自動 = 放不下時放外側
        var inside = st.arrowPos === "inside" ? true : st.arrowPos === "outside" ? false : len >= L * 2.6, pts;
        if (arrowLike(kind)) pts = inside ? [along(Q1, a, tr), along(Q2, a, -tr)] : [along(Q1, a, -L * 2.2), along(Q2, a, L * 2.2)];
        else if (kind === "tick") pts = [along(Q1, a, -L * 0.6), along(Q2, a, L * 0.6)];
        else pts = [Q1, Q2];

        var tx = dimText(d, g, dimLabel(d, st, len), st), aligned = st.textAlign !== "horizontal";
        var mid = [(Q1[0] + Q2[0]) / 2, (Q1[1] + Q2[1]) / 2], line = null;
        // 文字放在線中間；箭頭在外面（線很短）或放不下時，放在線的旁邊
        if (st.textPos === "middle" && (inside || !arrowLike(kind))) {
            line = splitLine(g, pts[0], pts[1], a, mid, textAlong(tx, a, aligned) + breakGap(st), Math.max(L, 2), st, col);
        }
        var split = !!line;
        if (!line) line = dimPath(g, pts, st, col, "DIM_Line");
        if (kind === "tick") {
            dimEnd(g, Q1, a, st, col, kind);
            dimEnd(g, Q2, a, st, col, kind);
        } else {
            var s = arrowLike(kind) && !inside ? 1 : -1;
            dimEnd(g, Q1, [a[0] * s, a[1] * s], st, col, kind);
            dimEnd(g, Q2, [-a[0] * s, -a[1] * s], st, col, kind);
        }
        placeText(tx, split ? mid : textCenter(mid, n, tx, aligned, gap), aligned ? readAngle(a) : 0);
        return line;
    }

    // 直徑、半徑的標法：leader 引線拉到外面、through 穿過圓心再拉出去、inside 畫在圓裡面、auto 放得下就畫在圓裡面
    function radMode(st) {
        var m = st.radMode;
        return m === "leader" || m === "through" || m === "inside" ? m : "auto";
    }

    // 引線長度 = 樣式的「引線」+ 同心圓內圈多拉的 x
    // 舊版的尺寸沒有 x、長度記在 off：同心圓的內圈維持原本的長度，其他的改用樣式的「引線」
    function leaderLen(gm, st) {
        var base = Number(st.leader), has = st.leader !== undefined && st.leader !== null && st.leader !== "" && base >= 0;
        if (gm.x !== undefined) return (has ? base : 12) + Number(gm.x);
        return has && !gm.out ? base : Number(gm.off) || 0;
    }

    // 直徑、半徑：回傳 DIM_Line
    function buildRadial(d, g, gm, st) {
        var c = gm.c, r = gm.r, u = gm.u, dia = gm.t === "dia", col = hexToColor(st.lineColor, d);
        var L = Number(st.endSize), kind = st.endStyle === "tick" ? "arrow" : st.endStyle, gap = Number(st.textGap);
        var tr = arrowTrim(kind, L), back = [-u[0], -u[1]];
        var tx = dimText(d, g, dimLabel(d, st, dia ? r * 2 : r, dia ? diaSym(st) : "R", gm.q), st);
        var aligned = st.textAlign !== "horizontal";
        var span = aligned ? tx.hw * 2 : 2 * (tx.hw * Math.abs(u[0]) + tx.hh * Math.abs(u[1]));
        var endRoom = kind === "none" ? 0 : L * 1.2, arm = Math.max(L * 0.6, Number(st.lineWidth) * 3);
        var M = along(c, u, r), M1 = along(c, u, -r), mode = radMode(st), line = null, pts, tc, ang = 0;

        // 同心圓的內圈一律用引線，文字才不會跟外圈的疊在一起
        if (gm.out) mode = "leader";
        else if (mode === "auto") {
            mode = (dia ? r * 2 >= span + endRoom * 2 + gap * 4 : r >= span + endRoom + arm + gap * 3) ? "inside" : "leader";
        }
        // 半徑的線從圓心出發時，圓心畫十字
        if (!dia && mode !== "leader" && st.centerMark !== false) {
            dimPath(g, [[c[0] - arm, c[1]], [c[0] + arm, c[1]]], st, col, "DIM_Center");
            dimPath(g, [[c[0], c[1] - arm], [c[0], c[1] + arm]], st, col, "DIM_Center");
        }

        if (mode === "inside") {
            ang = aligned ? readAngle(u) : 0;
            var p = dia ? along(M1, u, tr) : c, q = along(M, u, -tr);
            var mid = dia ? c : along(c, u, (r - endRoom + arm) / 2);
            if (st.textPos === "middle") line = splitLine(g, p, q, u, mid, textAlong(tx, u, aligned) + breakGap(st), Math.max(L, 2), st, col);
            if (line) tc = mid;
            else {
                line = dimPath(g, [p, q], st, col, "DIM_Line");
                tc = textCenter(mid, dirOf(readAngle(u) + 90), tx, aligned, besideGap(st));
            }
            if (dia) dimEnd(g, M1, back, st, col, kind);
            dimEnd(g, M, u, st, col, kind);
        } else {
            // 引線：箭頭從外面指向圓心；穿過圓心：線從圓心（直徑從對面的圓周）穿出來，箭頭往外指
            // 線拉到圓外 K 之後轉水平，文字放在水平線那一邊
            var K = along(M, u, Math.max(leaderLen(gm, st), L * 1.5, 2)), sx = u[0] < -0.0001 ? -1 : 1;
            var shelf = st.shelf === "end" || st.shelf === "none" ? st.shelf : "over";
            if (mode === "through") {
                pts = [dia ? along(M1, u, tr) : c, K];
                if (dia) dimEnd(g, M1, back, st, col, kind);
                dimEnd(g, M, u, st, col, kind);
            } else {
                pts = [along(M, u, tr), K];
                dimEnd(g, M, back, st, col, kind);
            }
            if (shelf === "over") {
                // 文字站在水平線上，水平線跟文字一樣長
                pts.push([K[0] + sx * (tx.hw * 2 + gap * 2), K[1]]);
                tc = [K[0] + sx * (tx.hw + gap), K[1] + besideGap(st) + tx.hh];
            } else {
                // 文字接在水平線（或引線）的後面，跟線的末端一樣高
                var S = K;
                if (shelf === "end") { S = [K[0] + sx * Math.max(L * 1.5, Number(st.fontSize) * 0.6), K[1]]; pts.push(S); }
                tc = [S[0] + sx * (gap + tx.hw), S[1]];
            }
            line = dimPath(g, pts, st, col, "DIM_Line");
        }
        placeText(tx, tc, ang);
        return line;
    }

    function ptStr(p) { return "[" + jsonNum(p[0]) + "," + jsonNum(p[1]) + "]"; }

    function geomStr(gm) {
        var s = '{"t":"' + gm.t + '"';
        if (gm.t === "lin") s += ',"p1":' + ptStr(gm.p1) + ',"p2":' + ptStr(gm.p2) + ',"a":' + ptStr(gm.a) + ',"n":' + ptStr(gm.n);
        else s += ',"c":' + ptStr(gm.c) + ',"r":' + jsonNum(gm.r) + ',"u":' + ptStr(gm.u) + (gm.out ? ',"out":1' : "") +
            (gm.x !== undefined ? ',"x":' + jsonNum(gm.x) : "") + (gm.q > 1 ? ',"q":' + gm.q : "");
        if (gm.z > 0) s += ',"z":' + jsonNum(gm.z);
        return s + ',"off":' + jsonNum(gm.off) + ',"ref":[' + ptStr(gm.ref[0]) + "," + ptStr(gm.ref[1]) + "]}";
    }

    function buildDim(d, container, gm, styleStr) {
        var st = parse(styleStr), g = container.groupItems.add(), line;
        try {
            line = gm.t === "lin" ? buildLinear(d, g, gm, st) : buildRadial(d, g, gm, st);
        } catch (e) {
            try { g.remove(); } catch (e2) {}
            throw e;
        }
        var pts = line.pathPoints, a = pts[0].anchor, b = pts[pts.length - 1].anchor;
        gm.ref = [[a[0], a[1]], [b[0], b[1]]];
        g.note = DIM_TAG + "|" + geomStr(gm) + "|" + styleStr;
        g.name = "尺寸 " + findChild(g, "DIM_Text").contents;
        return g;
    }

    // 兩組對應點 → 相似變換（移動、等比縮放、旋轉）：p' = k·p + b，用複數乘法
    function similarity(r0, r1, c0, c1) {
        var dx = r1[0] - r0[0], dy = r1[1] - r0[1], ex = c1[0] - c0[0], ey = c1[1] - c0[1];
        var dd = dx * dx + dy * dy;
        if (dd < 1e-12 || ex * ex + ey * ey < 1e-12) return null;
        var kr = (ex * dx + ey * dy) / dd, ki = (ey * dx - ex * dy) / dd;
        return { kr: kr, ki: ki, s: Math.sqrt(kr * kr + ki * ki),
            bx: c0[0] - (kr * r0[0] - ki * r0[1]), by: c0[1] - (ki * r0[0] + kr * r0[1]) };
    }

    function simPt(T, p) { return [T.kr * p[0] - T.ki * p[1] + T.bx, T.ki * p[0] + T.kr * p[1] + T.by]; }
    function simDir(T, v) { return [(T.kr * v[0] - T.ki * v[1]) / T.s, (T.ki * v[0] + T.kr * v[1]) / T.s]; }

    function simGeom(gm, T) {
        var o = { t: gm.t, off: gm.off * T.s };
        if (gm.t === "lin") {
            o.p1 = simPt(T, gm.p1); o.p2 = simPt(T, gm.p2); o.a = simDir(T, gm.a); o.n = simDir(T, gm.n);
        } else {
            o.c = simPt(T, gm.c); o.r = gm.r * T.s; o.u = simDir(T, gm.u); o.out = gm.out; o.q = gm.q;
            if (gm.x !== undefined) o.x = gm.x * T.s;
        }
        if (gm.z > 0) o.z = gm.z * T.s;
        return o;
    }

    // 用尺寸自己的記錄重建；styleStr 省略 = 原本的樣式。「同步全部」、「套用選取」一律照面板的樣式，
    // 不再依物件大小調整（自動調整只在新增時作用），所有尺寸的字級、線寬才會一致
    function rebuildDim(d, g, styleStr) {
        var p = dimNote(g), gm = parse(p.geom), line = findChild(g, "DIM_Line");
        if (!line || !gm.ref) return null;
        var pts = line.pathPoints;
        var T = similarity(gm.ref[0], gm.ref[1], pts[0].anchor, pts[pts.length - 1].anchor);
        if (T) gm = simGeom(gm, T);
        var locked = g.locked, hidden = g.hidden, saved = openAncestors(g);
        try {
            var ng = buildDim(d, g.layer, gm, styleStr || p.style);
            ng.move(g, ElementPlacement.PLACEBEFORE); // 留在原本的位置（圖層、群組、上下順序）
            g.remove();
            keepState(ng, locked, hidden);
            return ng;
        } finally { restoreTemp(saved); }
    }

    // 縮放整份文件時，尺寸的字級、線寬等也跟著縮放
    function scaleDimStyle(style, s) {
        var re = /"(fontSize|lineWidth|endSize|extGap|extOver|textGap|breakGap|leader)"\s*:\s*(-?[0-9.]+(?:[eE][-+]?[0-9]+)?)/g;
        return style.replace(re, function (m, k, v) {
            return '"' + k + '":' + roundTo(Number(v) * s, 3);
        });
    }

    // 「依物件大小自動調整樣式」：面板的樣式是給 autoBase (pt) 大的物件用的（樣式裡的 autoBase，0 = 關閉）。
    // 物件較長的一邊 z 跟它不同時，字級、線寬、箭頭、間距、引線和「距離」等比縮放；最小縮到一半
    var AUTO_MIN = 0.5, AUTO_MAX = 50;

    function autoFactor(d, styleStr, z) {
        var m = /"autoBase"\s*:\s*([0-9.eE+-]+)/.exec(styleStr), base = m ? Number(m[1]) : 0;
        if (!(base > 0) || !(z > 0)) return 1;
        return Math.max(AUTO_MIN, Math.min(AUTO_MAX, z * docScaleFactor(d) / base));
    }

    function autoStyle(d, styleStr, z) {
        var k = autoFactor(d, styleStr, z);
        return Math.abs(k - 1) < 0.000001 ? styleStr : scaleDimStyle(styleStr, k);
    }

    function longSide(b) { return Math.max(b[2] - b[0], b[1] - b[3]); }

    // ---- 寬高 ----
    // b = [左, 上, 右, 下]；寬度標在上方或下方，高度標在右側或左側
    function linGeom(b, side, off) {
        if (side === "bottom") return { t: "lin", p1: [b[0], b[3]], p2: [b[2], b[3]], a: [1, 0], n: [0, -1], off: off };
        if (side === "left") return { t: "lin", p1: [b[0], b[3]], p2: [b[0], b[1]], a: [0, 1], n: [-1, 0], off: off };
        if (side === "right") return { t: "lin", p1: [b[2], b[3]], p2: [b[2], b[1]], a: [0, 1], n: [1, 0], off: off };
        return { t: "lin", p1: [b[0], b[1]], p2: [b[2], b[1]], a: [1, 0], n: [0, 1], off: off };
    }

    // ---- 找圓弧 ----
    function bezPt(p0, p1, p2, p3, t) {
        var m = 1 - t, a = m * m * m, b = 3 * m * m * t, c = 3 * m * t * t, e = t * t * t;
        return [a * p0[0] + b * p1[0] + c * p2[0] + e * p3[0], a * p0[1] + b * p1[1] + c * p2[1] + e * p3[1]];
    }

    // 通過三點的圓；三點共線 = null
    function circle3(a, b, c) {
        var bx = b[0] - a[0], by = b[1] - a[1], cx = c[0] - a[0], cy = c[1] - a[1];
        var dd = 2 * (bx * cy - by * cx);
        if (Math.abs(dd) < 1e-9) return null;
        var b2 = bx * bx + by * by, c2 = cx * cx + cy * cy;
        var ux = (cy * b2 - by * c2) / dd, uy = (bx * c2 - cx * b2) / dd;
        return { c: [a[0] + ux, a[1] + uy], r: Math.sqrt(ux * ux + uy * uy) };
    }

    function arcTol(r) { return Math.max(0.02, r * 0.004); }

    function wrapAngle(a) {
        while (a < 0) a += Math.PI * 2;
        while (a >= Math.PI * 2) a -= Math.PI * 2;
        return a;
    }

    // 一段曲線是圓弧時回傳 { c, r, a0, sw }：a0 起點角度，sw 掃過的角度（弧度，逆時針為正）
    function segArc(p0, p1, p2, p3) {
        if (vlen(vsub(p3, p0)) < 0.01) return null;
        if (vlen(vsub(p1, p0)) < 0.0001 && vlen(vsub(p2, p3)) < 0.0001) return null; // 直線
        var m = bezPt(p0, p1, p2, p3, 0.5), k = circle3(p0, m, p3);
        if (!k || k.r > 100000) return null;
        var tol = arcTol(k.r);
        for (var t = 0.125; t < 1; t += 0.25) {
            if (Math.abs(vlen(vsub(bezPt(p0, p1, p2, p3, t), k.c)) - k.r) > tol) return null;
        }
        var a0 = Math.atan2(p0[1] - k.c[1], p0[0] - k.c[0]), a1 = Math.atan2(p3[1] - k.c[1], p3[0] - k.c[0]);
        var ccw = (p0[0] - k.c[0]) * (m[1] - k.c[1]) - (p0[1] - k.c[1]) * (m[0] - k.c[0]) > 0;
        return { c: k.c, r: k.r, a0: a0, sw: ccw ? wrapAngle(a1 - a0) : -wrapAngle(a0 - a1) };
    }

    function sameCircle(p, q) {
        var tol = Math.max(0.05, Math.max(p.r, q.r) * 0.005);
        return Math.abs(p.r - q.r) <= tol && vlen(vsub(p.c, q.c)) <= tol;
    }

    // 直接選取工具：這一段被選到了嗎（起點的右側、終點的左側）
    function segSelected(s0, s1) {
        var A = PathPointSelection.ANCHORPOINT, LR = PathPointSelection.LEFTRIGHTDIRECTION;
        return (s0 == A || s0 == LR || s0 == PathPointSelection.RIGHTDIRECTION) &&
            (s1 == A || s1 == LR || s1 == PathPointSelection.LEFTDIRECTION);
    }

    // 路徑上的圓弧：相鄰、同一個圓的段落合併成一個弧
    // useSel：路徑本身被直接選取時，只看選到的段落
    function pathArcs(p, useSel, budget) {
        var pts = p.pathPoints, n = pts.length, i, out = [];
        if (n < 2) return out;
        var cnt = p.closed ? n : n - 1;
        if (budget.n + cnt > MAX_SEGS) { budget.over = true; return out; }
        budget.n += cnt;

        var data = [], sel = [], partial = false, any = false, NONE = PathPointSelection.NOSELECTION;
        for (i = 0; i < n; i++) {
            var q = pts[i];
            data.push([q.anchor, q.leftDirection, q.rightDirection]);
            if (useSel) {
                var s = q.selected;
                sel.push(s);
                if (s != PathPointSelection.ANCHORPOINT) partial = true;
                if (s != NONE) any = true;
            }
        }
        var use = [];
        for (i = 0; i < cnt; i++) use.push(true);
        partial = partial && any;
        if (partial) {
            var hit = false;
            for (i = 0; i < cnt; i++) { use[i] = segSelected(sel[i], sel[(i + 1) % n]); if (use[i]) hit = true; }
            // 只點了錨點：用跟它相連的段落
            if (!hit) for (i = 0; i < cnt; i++) use[i] = sel[i] != NONE || sel[(i + 1) % n] != NONE;
        }

        var cur = null;
        for (i = 0; i < cnt; i++) {
            var A = data[i], B = data[(i + 1) % n];
            var arc = use[i] ? segArc(A[0], A[2], B[1], B[0]) : null;
            if (arc && cur && cur.last === i - 1 && sameCircle(cur, arc) && (cur.sw > 0) === (arc.sw > 0)) {
                cur.sw += arc.sw;
                cur.last = i;
            } else if (arc) {
                cur = { c: arc.c, r: arc.r, a0: arc.a0, sw: arc.sw, first: i, last: i, sel: partial };
                out.push(cur);
            } else {
                cur = null;
            }
        }
        // 封閉路徑：最後一個弧接回第一個
        if (p.closed && out.length > 1) {
            var f = out[0], l = out[out.length - 1];
            if (f.first === 0 && l.last === cnt - 1 && sameCircle(f, l) && (f.sw > 0) === (l.sw > 0)) {
                l.sw += f.sw;
                out.shift();
            }
        }
        for (i = 0; i < out.length; i++) {
            out[i].full = Math.abs(out[i].sw) >= Math.PI * 2 - 0.05;
            if (out[i].full) out[i].sw = out[i].sw > 0 ? Math.PI * 2 : -Math.PI * 2;
        }
        return out;
    }

    function collectPaths(it, out) {
        var t = it.typename, i;
        if (t === "PathItem") out.push(it);
        else if (t === "CompoundPathItem") for (i = 0; i < it.pathItems.length; i++) out.push(it.pathItems[i]);
        else if (t === "GroupItem") for (i = 0; i < it.pageItems.length; i++) collectPaths(it.pageItems[i], out);
    }

    // 線條超出路徑多少：看 visibleBounds 和 geometricBounds 的差，線條對齊內側、外側也算得對；
    // 最多算到線寬（效果撐大外框時不會多算）
    function strokeOut(p) {
        try {
            if (!p.stroked || !(p.strokeWidth > 0)) return 0;
            var gb = p.geometricBounds, vb = p.visibleBounds;
            var e = Math.min(vb[2] - vb[0] - (gb[2] - gb[0]), vb[1] - vb[3] - (gb[1] - gb[3])) / 2;
            return Math.max(0, Math.min(p.strokeWidth, e));
        } catch (err) { return 0; }
    }

    // visible：「含線寬」，直徑、半徑量到線條的外緣
    function itemArcs(it, budget, visible) {
        var paths = [], out = [], i, k;
        collectPaths(it, paths);
        for (i = 0; i < paths.length && !budget.over; i++) {
            var arcs = pathArcs(paths[i], paths[i] === it, budget), ext = visible ? strokeOut(paths[i]) : 0;
            for (k = 0; k < arcs.length; k++) {
                arcs[k].r += ext;
                out.push(arcs[k]);
            }
        }
        return out;
    }

    function rotDir(u, deg) {
        var a = deg * Math.PI / 180, c = Math.cos(a), s = Math.sin(a);
        return [u[0] * c - u[1] * s, u[0] * s + u[1] * c];
    }

    // 標註方向：整圓用指定的角度，圓弧用弧的中間
    // 直徑標半圓以上的弧時，線跟弧的中間垂直，兩端才都落在弧上（選靠近指定角度的那一邊）
    function arcDir(arc, angle, dia) {
        if (arc.full) return dirOf(angle);
        var m = arc.a0 + arc.sw / 2, u = [Math.cos(m), Math.sin(m)];
        if (dia && Math.abs(arc.sw) >= Math.PI - 0.05) {
            var p = [-u[1], u[0]];
            return vdot(p, dirOf(angle)) >= 0 ? p : [u[1], -u[0]];
        }
        return u;
    }

    // 一個物件要標的圓弧：同樣大小只標一個，挑物件上最靠「角度」那一側的（45° = 右上角）
    // 直徑只看整圓、半圓以上的弧和直接選取的弧；半徑看全部
    // n：同樣大小的有幾個（標「4×Ø3」用）；直徑只數整圓，長圓孔兩端的半圓不算兩個孔
    function pickArcs(arcs, dia, b, angle) {
        var picks = [], i, k, want = dirOf(angle);
        var cx = (b[0] + b[2]) / 2, cy = (b[1] + b[3]) / 2;
        var W = Math.max(b[2] - b[0], 0.001), H = Math.max(b[1] - b[3], 0.001);
        for (i = 0; i < arcs.length; i++) {
            var a = arcs[i];
            if (dia && !a.full && !a.sel && Math.abs(a.sw) < Math.PI - 0.05) continue;
            var u = arcDir(a, angle, dia), M = along(a.c, u, a.r);
            var cand = { c: a.c, r: a.r, u: u, score: vdot([(M[0] - cx) / W, (M[1] - cy) / H], want), n: 0 }, same = -1;
            for (k = 0; k < picks.length; k++) {
                if (Math.abs(picks[k].r - a.r) <= Math.max(0.05, a.r * 0.005)) { same = k; break; }
            }
            if (same < 0) {
                if (picks.length < MAX_ARCS) { picks.push(cand); same = picks.length - 1; }
            } else if (cand.score > picks[same].score + 1e-6) {
                cand.n = picks[same].n;
                picks[same] = cand;
            }
            if (same >= 0 && (!dia || a.full)) picks[same].n++;
        }
        picks.sort(function (p, q) { return q.r - p.r; });
        return picks;
    }

    // ---- 對外 ----
    // o = { jobs: [{ kind: "w" | "h", side: "top" | "bottom" | "left" | "right", each }, { kind: "dia" | "rad" }],
    //       offset 寬高的尺寸線距離, angle, visible 含線寬（寬高、直徑、半徑都量到線條外緣）, style: 樣式 JSON 字串 }
    //   each：每個物件各自標；否則標選取的物件合起來的外框。一次呼叫做完全部
    //   舊的寫法 { kind: "w" | "h" | "wh" | "dia" | "rad", wSide, hSide, each } 也可以用
    // 直徑、半徑的引線長度在樣式裡（「同步全部」會一起改）
    function dimJobs(o) {
        if (o.jobs) return o.jobs;
        var k = o.kind, out = [];
        if (k === "w" || k === "wh") out.push({ kind: "w", side: o.wSide === "bottom" ? "bottom" : "top", each: o.each });
        if (k === "h" || k === "wh") out.push({ kind: "h", side: o.hSide === "left" ? "left" : "right", each: o.each });
        if (k === "dia" || k === "rad") out.push({ kind: k });
        return out;
    }

    function dimAdd(optStr) {
        try {
            var d = getDoc(), o = parse(optStr), jobs = dimJobs(o), i, j, q;
            var items = measuredItems(d), off = Number(o.offset) || 0, made = [], zero = 0, noArc = [], done = {};
            var firstK = 0;   // 第一個尺寸用的自動調整倍率：面板把樣式改成這個尺寸實際用的數值
            if (!items.length) {
                return res(false, selectedItems(d).length ? "選取的是尺寸或標註，請選取要量的物件" : "請先選取要標尺寸的物件");
            }
            var layer = getLayer(d, true, DIM_LAYER), budget = { n: 0, over: false }, angle = Number(o.angle) || 0;
            for (q = 0; q < jobs.length; q++) {
                var kind = jobs[q].kind, side = jobs[q].side;
                if (kind === "w" || kind === "h") {
                    var sets = [];
                    if (jobs[q].each) { for (i = 0; i < items.length; i++) sets.push([items[i]]); }
                    else sets.push(items);
                    if (kind === "w") side = side === "bottom" ? "bottom" : "top";
                    else side = side === "left" ? "left" : "right";
                    for (i = 0; i < sets.length; i++) {
                        var b = unionBounds(sets[i], !!o.visible);
                        // 選範圍和各物件一起標、只選一個物件時，同一個位置只標一次
                        var key = side + "|" + roundTo(b[0], 2) + "|" + roundTo(b[1], 2) + "|" + roundTo(b[2], 2) + "|" + roundTo(b[3], 2);
                        if (done[key]) continue;
                        done[key] = true;
                        if (!((kind === "w" ? b[2] - b[0] : b[1] - b[3]) > 0.001)) { zero++; continue; }
                        var z = longSide(b), kz = autoFactor(d, o.style, z), gw = linGeom(b, side, off * kz);
                        gw.z = z;
                        if (!firstK) firstK = kz;
                        made.push(buildDim(d, layer, gw, autoStyle(d, o.style, z)));
                    }
                } else if (kind === "dia" || kind === "rad") {
                    var dia = kind === "dia", before = made.length;
                    for (i = 0; i < items.length; i++) {
                        var ib = boundsOf(items[i], false), iz = longSide(ib), ist = autoStyle(d, o.style, iz);
                        var picks = pickArcs(itemArcs(items[i], budget, !!o.visible), dia, ib, angle);
                        if (picks.length && !firstK) firstK = autoFactor(d, o.style, iz);
                        for (j = 0; j < picks.length; j++) {
                            var pk = picks[j], gm = { t: dia ? "dia" : "rad", c: pk.c, r: pk.r, u: pk.u, off: 0, x: 0, q: Math.max(1, pk.n), z: iz };
                            // 同心圓（例如圓環的內外圈）：最大的照常標，小的轉 45°、用引線拉到最外圈外面，文字才不會疊在一起
                            var inner = 0, outerR = pk.r;
                            for (var k = 0; k < j; k++) {
                                if (vlen(vsub(picks[k].c, pk.c)) <= Math.max(0.05, pk.r * 0.005)) { inner++; outerR = Math.max(outerR, picks[k].r); }
                            }
                            if (inner) { gm.u = rotDir(pk.u, -45 * inner); gm.out = 1; gm.x = outerR - pk.r; }
                            made.push(buildDim(d, layer, gm, ist));
                        }
                    }
                    if (made.length === before) noArc.push(kind);
                }
            }
            var NO_ARC = { dia: "選取的物件裡沒有圓（圓角請用「半徑 R」）", rad: "選取的物件裡沒有圓弧或圓角" };
            if (!made.length) {
                if (budget.over) return res(false, "選取的物件太複雜，請只選要標的圓或圓弧");
                if (noArc.length) return res(false, NO_ARC[noArc[0]]);
                return res(false, zero ? "選取的物件寬度或高度是 0" : "沒有要標的尺寸");
            }
            var names = [];
            for (i = 0; i < made.length && i < 4; i++) names.push(findChild(made[i], "DIM_Text").contents);
            var msg = "已標註 " + names.join("、") + (made.length > 4 ? " 等 " + made.length + " 個尺寸" : "");
            if (zero) msg += "（" + zero + " 個是 0，沒有標）";
            for (i = 0; i < noArc.length; i++) msg += "（" + (noArc[i] === "dia" ? "沒有找到圓" : "沒有找到圓弧或圓角") + "）";
            if (budget.over) msg += "（圖太複雜，只找了一部分的圓弧）";
            return '{"ok":true,"msg":"' + esc(msg) + '","k":' + jsonNum(firstK || 1) + "}";
        } catch (e) { return res(false, e.message); }
    }

    // 面板顯示用：選取物件合起來的寬高（pt）、文件單位；o.visible = 含線寬
    function dimInfo(optStr) {
        try {
            if (app.documents.length === 0) return '{"ok":true,"doc":false}';
            var o = optStr ? parse(optStr) : {};
            var d = app.activeDocument, u = docUnit(d), f = docScaleFactor(d), items = measuredItems(d);
            var out = '{"ok":true,"doc":true,"sel":' + items.length + ',"docUnit":"' + u[0] + '","docUnitPt":' + jsonNum(u[1]);
            if (items.length) {
                var b = unionBounds(items, !!o.visible);
                out += ',"w":' + jsonNum((b[2] - b[0]) * f) + ',"h":' + jsonNum((b[1] - b[3]) * f);
            }
            return out + "}";
        } catch (e) { return res(false, e.message); }
    }

    function rebuildDims(d, gs, styleStr) {
        var r = { made: [], failed: 0, why: "" };
        for (var i = 0; i < gs.length; i++) {
            var ng = null;
            try { ng = rebuildDim(d, gs[i], styleStr); } catch (e) { if (!r.why) r.why = e.message; }
            if (ng) r.made.push(ng); else r.failed++;
        }
        return r;
    }

    // 文件中所有尺寸套用目前樣式（位置不變）
    function dimSync(styleStr) {
        try {
            var d = getDoc(), gs = allDims(d);
            if (gs.length === 0) return res(true, "文件中沒有尺寸");
            var r = rebuildDims(d, gs, styleStr);
            return res(r.made.length > 0, "已同步 " + r.made.length + " 個尺寸的樣式" + failNote(r));
        } catch (e) { return res(false, e.message); }
    }

    function dimRestyle(styleStr) {
        try {
            var d = getDoc(), gs = selectedDims(d);
            if (gs.length === 0) return res(false, "請先選取尺寸");
            var r = rebuildDims(d, gs, styleStr);
            d.selection = r.made;
            return res(r.made.length > 0, "已更新 " + r.made.length + " 個尺寸的樣式" + failNote(r));
        } catch (e) { return res(false, e.message); }
    }

    function dimSelectAll() {
        try {
            var d = getDoc(), gs = allDims(d);
            if (gs.length === 0) return res(false, "文件中沒有尺寸");
            d.selection = gs;
            return res(true, "已選取 " + gs.length + " 個尺寸");
        } catch (e) { return res(false, e.message); }
    }

    function dimToggle() {
        try {
            var L = getLayer(getDoc(), false, DIM_LAYER);
            if (!L) return res(false, "還沒有「" + DIM_LAYER + "」圖層");
            L.visible = !L.visible;
            return res(true, L.visible ? "尺寸圖層已顯示" : "尺寸圖層已隱藏");
        } catch (e) { return res(false, e.message); }
    }

    // ---------- 設定檔（給快捷鍵腳本使用） ----------
    // 面板第一次開啟（還沒有自己的設定）時，沿用快捷鍵腳本用的設定檔
    function loadSettings() {
        try {
            var file = new File(SETTINGS_FILE);
            if (!file.exists) return res(false, "");
            file.encoding = "UTF-8";
            file.open("r");
            var str = file.read();
            file.close();
            parse(str); // 確認是有效的 JSON
            return '{"ok":true,"settings":' + str + '}';
        } catch (e) { return res(false, ""); }
    }

    function saveSettings(str) {
        try {
            var f = new Folder(SETTINGS_DIR);
            if (!f.exists) f.create();
            var file = new File(SETTINGS_FILE);
            file.encoding = "UTF-8";
            file.open("w"); file.write(str); file.close();
            return res(true, "saved");
        } catch (e) { return res(false, e.message); }
    }

    return {
        beginAdd: beginAdd,
        pollAdd: pollAdd,
        endAdd: endAdd,
        finishFree: finishFree,
        convert: convert,
        renumber: renumber,
        setNumber: setNumber,
        listNumbers: listNumbers,
        selectNumber: selectNumber,
        changeNumber: changeNumber,
        syncAll: syncAll,
        restyle: restyle,
        relayout: relayout,
        selectAll: selectAll,
        toggleLayer: toggleLayer,
        listFonts: listFonts,
        beginPick: beginPick,
        pollPick: pollPick,
        endPick: endPick,
        setLink: setLink,
        setLinks: setLinks,
        scaleInfo: scaleInfo,
        scaleDoc: scaleDoc,
        swapItems: swapItems,
        dimAdd: dimAdd,
        dimInfo: dimInfo,
        dimSync: dimSync,
        dimRestyle: dimRestyle,
        dimSelectAll: dimSelectAll,
        dimToggle: dimToggle,
        loadSettings: loadSettings,
        saveSettings: saveSettings
    };
})();
