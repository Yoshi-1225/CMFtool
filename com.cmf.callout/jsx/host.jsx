/*
 * CMF Callout - Illustrator ExtendScript 核心
 * 對外函式接收 JSON 字串、回傳 JSON 字串：{"ok":true,"msg":"...","next":5,"state":"done"}
 * 注意：ExtendScript 是 ES3，沒有 JSON / let / 箭頭函式 / Array.indexOf
 */
$.global.CMF = (function () {
    var LAYER_NAME = "CMF Callouts";
    var TAG = "CMF_CALLOUT";
    var SETTINGS_DIR = Folder.userData + "/CMFCallout";
    var SETTINGS_FILE = SETTINGS_DIR + "/settings.json";

    // ---------- 共用 ----------
    function parse(s) { return eval("(" + s + ")"); }

    function esc(s) { return String(s).replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\n/g, " "); }

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

    function getLayer(d, create) {
        var L = null;
        try { L = d.layers.getByName(LAYER_NAME); } catch (e) { L = null; }
        if (!L && create) { L = d.layers.add(); L.name = LAYER_NAME; }
        if (L && create) { L.locked = false; L.visible = true; }
        return L;
    }

    function isCallout(it) {
        return it && it.typename === "GroupItem" && it.note && String(it.note).indexOf(TAG + "|") === 0;
    }

    function noteParts(g) {
        var p = String(g.note).split("|");
        return { num: parseInt(p[1], 10), style: p.slice(2).join("|") };
    }

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
    function buildCallout(d, line, num, styleStr, container) {
        var o = parse(styleStr);
        var pts = line.pathPoints, n = pts.length;
        var S = pts[0].anchor, S1 = pts[1].anchor;
        var E = pts[n - 1].anchor, E1 = pts[n - 2].anchor;
        var lineColor = hexToColor(o.lineColor, d);

        var g = container.groupItems.add();
        g.name = "CMF " + num;
        g.note = TAG + "|" + num + "|" + styleStr;

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

    // 拆掉舊標註，用同一條線重建
    function rebuild(d, g, num, styleStr) {
        var line = findChild(g, "CMF_Line");
        if (!line) return null;
        var end = findChild(g, "CMF_End");
        if (end && end.typename === "PathItem" && end.closed && end.pathPoints.length === 3) {
            var tip = end.pathPoints[0].anchor; // 還原箭頭尖端位置
            var last = line.pathPoints[line.pathPoints.length - 1];
            last.anchor = tip; last.leftDirection = tip; last.rightDirection = tip;
        }
        var container = g.layer;
        container.locked = false;
        container.visible = true; // 隱藏圖層中的物件無法修改
        line.move(g, ElementPlacement.PLACEBEFORE);
        g.remove();
        return buildCallout(d, line, num, styleStr, container);
    }

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
            var d = getDoc(), o = parse(optStr), cs = selectedCallouts(d), made = [];
            if (cs.length === 0) return res(false, "請先選取標註");
            var n = parseInt(o.num, 10);
            if (!(n > 0)) return res(false, "編號需為正整數");
            for (var i = 0; i < cs.length; i++) made.push(rebuild(d, cs[i], n, noteParts(cs[i]).style));
            d.selection = made;
            return res(true, "已將 " + made.length + " 個標註設為 " + n);
        } catch (e) { return res(false, e.message); }
    }

    // 編號表：列出每個編號與數量
    function listNumbers() {
        try {
            if (app.documents.length === 0) return '{"ok":true,"items":[]}';
            var cs = allCallouts(app.activeDocument), nums = [], counts = {}, i;
            for (i = 0; i < cs.length; i++) {
                var n = noteParts(cs[i]).num;
                if (counts[n] === undefined) { counts[n] = 0; nums.push(n); }
                counts[n]++;
            }
            nums.sort(function (a, b) { return a - b; });
            var parts = [];
            for (i = 0; i < nums.length; i++) parts.push('{"num":' + nums[i] + ',"count":' + counts[nums[i]] + '}');
            return '{"ok":true,"items":[' + parts.join(",") + ']}';
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
            for (var i = 0; i < cs.length; i++) rebuild(d, cs[i], noteParts(cs[i]).num, styleStr);
            if (sel.length === 0) d.selection = null;
            return res(true, "已同步 " + cs.length + " 個標註的樣式");
        } catch (e) { return res(false, e.message); }
    }

    // 套用到選取的標註
    function restyle(styleStr) {
        try {
            var d = getDoc(), cs = selectedCallouts(d), made = [];
            if (cs.length === 0) return res(false, "請先選取標註");
            for (var i = 0; i < cs.length; i++) made.push(rebuild(d, cs[i], noteParts(cs[i]).num, styleStr));
            d.selection = made;
            return res(true, "已更新 " + made.length + " 個標註的樣式");
        } catch (e) { return res(false, e.message); }
    }

    // 移動過線條錨點後，用原本樣式重新排版
    function relayout() {
        try {
            var d = getDoc(), cs = selectedCallouts(d), made = [];
            if (cs.length === 0) return res(false, "請先選取要重新排版的標註");
            for (var i = 0; i < cs.length; i++) {
                var p = noteParts(cs[i]);
                made.push(rebuild(d, cs[i], p.num, p.style));
            }
            d.selection = made;
            return res(true, "已重新排版 " + made.length + " 個標註");
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

    // ---------- 設定檔（給快捷鍵腳本使用） ----------
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
        saveSettings: saveSettings
    };
})();
