/*
 * Excel 同步 — 表格版面與樣式（Node 模組）
 * SheetJS 免費版讀不到樣式，所以這裡直接解析 .xlsx 裡的 XML：
 *   styles.xml（字型、填色、框線、對齊）、theme1.xml（佈景主題色彩）、各工作表的 XML（欄寬列高、選取範圍）
 */
'use strict';

module.exports = function (XLSX) {
  var PX = 0.75;          // Excel 100% 時 1px = 0.75pt
  var MAX_CELLS = 2000;

  /* ---------- 儲存格位址 ---------- */

  function quoteSheet(name) {
    return /[^A-Za-z0-9_\u4e00-\u9fff]/.test(name) ? "'" + name.replace(/'/g, "''") + "'" : name;
  }

  // "A1:D10" / "工作表1!A1:D10" / "'價格 表'!B2" / 反向選取也可以
  function parseRange(input) {
    var str = String(input).trim(), i = str.lastIndexOf('!'), sheet = null, body = str;
    if (i >= 0) {
      sheet = str.slice(0, i).replace(/^'(.*)'$/, '$1').replace(/''/g, "'");
      body = str.slice(i + 1);
    }
    body = body.replace(/\$/g, '').toUpperCase();
    if (!/^[A-Z]{1,3}[1-9][0-9]{0,6}(:[A-Z]{1,3}[1-9][0-9]{0,6})?$/.test(body) || sheet === '') return null;
    var rg = XLSX.utils.decode_range(body.indexOf(':') < 0 ? body + ':' + body : body);
    return {
      sheet: sheet,
      s: { r: Math.min(rg.s.r, rg.e.r), c: Math.min(rg.s.c, rg.e.c) },
      e: { r: Math.max(rg.s.r, rg.e.r), c: Math.max(rg.s.c, rg.e.c) }
    };
  }

  function cellText(ws, addr) {
    var c = ws[addr];
    if (!c) {
      var pos = XLSX.utils.decode_cell(addr);
      (ws['!merges'] || []).some(function (m) {
        if (pos.r >= m.s.r && pos.r <= m.e.r && pos.c >= m.s.c && pos.c <= m.e.c) {
          c = ws[XLSX.utils.encode_cell(m.s)];
          return true;
        }
        return false;
      });
    }
    if (!c || c.v == null) return '';
    return c.w != null ? c.w : String(c.v);
  }

  /* ---------- 讀 .xlsx 裡的原始 XML ---------- */

  function stripNs(xml) { return xml.replace(/<(\/?)[A-Za-z0-9_]+:/g, '<$1'); }

  function unescapeXml(s) {
    return s == null ? s : s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'").replace(/&amp;/g, '&');
  }

  function fileText(wb, p) {
    if (!wb.files || !p) return '';
    var k = p.replace(/^\//, '');
    var f = wb.files[k] || wb.files['/' + k];
    if (!f || !f.content) return '';
    return stripNs(Buffer.from(f.content).toString('utf8'));
  }

  function attr(tag, name) {
    if (!tag) return null;
    var m = tag.match(new RegExp('(?:^|\\s)' + name.replace(':', '\\:') + '="([^"]*)"'));
    return m ? unescapeXml(m[1]) : null;
  }

  function tags(xml, name) { return xml.match(new RegExp('<' + name + '\\b[^>]*>', 'g')) || []; }

  function elements(xml, name) {
    return xml.match(new RegExp('<' + name + '\\b[^>]*?(?:/>|>[\\s\\S]*?</' + name + '>)', 'g')) || [];
  }

  function section(xml, name) {
    var m = xml.match(new RegExp('<' + name + '\\b[^>]*>([\\s\\S]*?)</' + name + '>'));
    return m ? m[1] : '';
  }

  // 工作表名稱 → XML 路徑（照 workbook.xml 的關聯，不靠檔名順序）
  function sheetPath(wb, name) {
    var book = fileText(wb, 'xl/workbook.xml');
    var rels = fileText(wb, 'xl/_rels/workbook.xml.rels');
    var sheet = tags(book, 'sheet').filter(function (t) { return attr(t, 'name') === name; })[0];
    var rel = sheet && tags(rels, 'Relationship').filter(function (t) { return attr(t, 'Id') === attr(sheet, 'r:id'); })[0];
    var target = rel && attr(rel, 'Target');
    if (target) return target.charAt(0) === '/' ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
    var idx = wb.SheetNames.indexOf(name);
    return wb.Directory && wb.Directory.sheets ? wb.Directory.sheets[idx] : null;
  }

  function sheetXml(wb, name) {
    wb._xmlCache = wb._xmlCache || {};
    if (!(name in wb._xmlCache)) wb._xmlCache[name] = fileText(wb, sheetPath(wb, name));
    return wb._xmlCache[name];
  }

  /* ---------- Excel 存檔時選取的範圍 ---------- */

  function savedSelection(wb) {
    var view = tags(fileText(wb, 'xl/workbook.xml'), 'workbookView')[0];
    var tab = parseInt(attr(view, 'activeTab'), 10) || 0;
    var name = wb.SheetNames[tab] || wb.SheetNames[0];
    var sv = elements(sheetXml(wb, name), 'sheetView')[0] || '';
    var pane = tags(sv, 'pane')[0];
    var activePane = pane ? attr(pane, 'activePane') : null;
    var sels = tags(sv, 'selection');
    var sel = sels.filter(function (t) { return attr(t, 'pane') === activePane; })[0] || sels[0];
    var ref = sel ? (attr(sel, 'sqref') || attr(sel, 'activeCell')) : null;
    return { sheet: name, range: (ref || 'A1').split(' ')[0] };
  }

  /* ---------- 顏色 ---------- */

  var INDEXED = [
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
    '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
    '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
    '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
    '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
    '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'
  ];

  function parseTheme(xml) {
    var scheme = section(xml, 'clrScheme');
    var pick = function (n) {
      var el = section(scheme, n);
      var t = tags(el, 'srgbClr')[0], s = tags(el, 'sysClr')[0];
      return (t && attr(t, 'val')) || (s && attr(s, 'lastClr')) || null;
    };
    // Excel 的佈景主題索引：0 背景1、1 文字1、2 背景2、3 文字2、4–9 輔色1–6
    return [pick('lt1') || 'FFFFFF', pick('dk1') || '000000', pick('lt2') || 'E7E6E6', pick('dk2') || '44546A',
            pick('accent1'), pick('accent2'), pick('accent3'), pick('accent4'), pick('accent5'), pick('accent6'),
            pick('hlink'), pick('folHlink')];
  }

  function tint(hex, t) {
    if (!t) return hex;
    var r = parseInt(hex.substr(0, 2), 16) / 255, g = parseInt(hex.substr(2, 2), 16) / 255, b = parseInt(hex.substr(4, 2), 16) / 255;
    var max = Math.max(r, g, b), min = Math.min(r, g, b), l = (max + min) / 2, h = 0, s = 0, d = max - min;
    if (d) {
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
      h /= 6;
    }
    l = t < 0 ? l * (1 + t) : l * (1 - t) + t;
    var hue = function (p, q, x) {
      if (x < 0) x += 1; if (x > 1) x -= 1;
      if (x < 1 / 6) return p + (q - p) * 6 * x;
      if (x < 1 / 2) return q;
      if (x < 2 / 3) return p + (q - p) * (2 / 3 - x) * 6;
      return p;
    };
    var out;
    if (!s) out = [l, l, l];
    else {
      var q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
      out = [hue(p, q, h + 1 / 3), hue(p, q, h), hue(p, q, h - 1 / 3)];
    }
    return out.map(function (v) { return ('0' + Math.round(v * 255).toString(16)).slice(-2); }).join('').toUpperCase();
  }

  function colorOf(tag, st) {
    if (!tag) return null;
    var rgb = attr(tag, 'rgb'), theme = attr(tag, 'theme'), idx = attr(tag, 'indexed'), hex = null;
    if (rgb) hex = rgb.length === 8 ? rgb.slice(2) : rgb;
    else if (theme != null) hex = st.theme[+theme] || null;
    else if (idx != null) hex = +idx === 64 ? '000000' : +idx === 65 ? 'FFFFFF' : (st.indexed[+idx] || null);
    else if (attr(tag, 'auto')) hex = '000000';
    if (!hex) return null;
    return tint(hex.toUpperCase(), parseFloat(attr(tag, 'tint')) || 0);
  }

  /* ---------- styles.xml ---------- */

  function flag(xml, name) {
    var t = tags(xml, name)[0];
    if (!t) return false;
    var v = attr(t, 'val');
    return v == null || !/^(0|false|none)$/i.test(v);
  }

  function parseStyles(wb) {
    if (wb._styles) return wb._styles;
    var xml = fileText(wb, 'xl/styles.xml');
    var themeFile = Object.keys(wb.files || {}).filter(function (k) { return /^\/?xl\/theme\/theme\d*\.xml$/.test(k); }).sort()[0];
    var st = { theme: parseTheme(fileText(wb, themeFile)), indexed: INDEXED.slice() };

    var custom = tags(section(xml, 'indexedColors'), 'rgbColor');
    custom.forEach(function (t, i) { var v = attr(t, 'rgb'); if (v) st.indexed[i] = v.slice(-6); });

    st.fonts = elements(section(xml, 'fonts'), 'font').map(function (f) {
      return {
        name: attr(tags(f, 'name')[0], 'val') || 'Calibri',
        sz: parseFloat(attr(tags(f, 'sz')[0], 'val')) || 11,
        b: flag(f, 'b'), i: flag(f, 'i'), u: flag(f, 'u'), st: flag(f, 'strike'),
        color: colorOf(tags(f, 'color')[0], st) || '000000'
      };
    });
    if (!st.fonts.length) st.fonts = [{ name: 'Calibri', sz: 11, b: false, i: false, u: false, st: false, color: '000000' }];

    st.fills = elements(section(xml, 'fills'), 'fill').map(function (f) {
      var pf = tags(f, 'patternFill')[0];
      if (pf) {
        var type = attr(pf, 'patternType');
        if (!type || type === 'none' || type === 'gray125') return null;
        return colorOf(tags(f, 'fgColor')[0], st) || colorOf(tags(f, 'bgColor')[0], st);
      }
      var stop = tags(f, 'color')[0];         // 漸層填色：取第一個顏色
      return stop ? colorOf(stop, st) : null;
    });

    st.borders = elements(section(xml, 'borders'), 'border').map(function (b) {
      var side = function (names) {
        for (var i = 0; i < names.length; i++) {
          var el = elements(b, names[i])[0];
          var style = el && attr(el, 'style');
          if (style && style !== 'none') return { style: style, color: colorOf(tags(el, 'color')[0], st) || '000000' };
        }
        return null;
      };
      return { left: side(['left', 'start']), right: side(['right', 'end']), top: side(['top']), bottom: side(['bottom']) };
    });

    st.xfs = elements(section(xml, 'cellXfs'), 'xf').map(function (x) {
      var open = tags(x, 'xf')[0], al = tags(x, 'alignment')[0];
      return {
        font: +attr(open, 'fontId') || 0, fill: +attr(open, 'fillId') || 0, border: +attr(open, 'borderId') || 0,
        ha: al ? attr(al, 'horizontal') : null, va: al ? attr(al, 'vertical') : null,
        wrap: al ? /^(1|true)$/.test(attr(al, 'wrapText') || '') : false,
        indent: al ? +attr(al, 'indent') || 0 : 0
      };
    });
    if (!st.xfs.length) st.xfs = [{ font: 0, fill: 0, border: 0, ha: null, va: null, wrap: false, indent: 0 }];

    return (wb._styles = st);
  }

  // 每格用哪個樣式：儲存格 > 列 > 欄
  function sheetStyleIndex(xml) {
    var cell = {}, row = {}, col = [];
    var re = /<c\b[^>]*>/g, m;
    while ((m = re.exec(xml))) {
      var r = attr(m[0], 'r'), s = attr(m[0], 's');
      if (r && s) cell[r] = +s;
    }
    tags(xml, 'row').forEach(function (t) {
      if (/^(1|true)$/.test(attr(t, 'customFormat') || '')) row[+attr(t, 'r') - 1] = +attr(t, 's') || 0;
    });
    tags(section(xml, 'cols'), 'col').forEach(function (t) {
      var st = attr(t, 'style');
      if (st == null) return;
      for (var c = +attr(t, 'min') - 1; c <= +attr(t, 'max') - 1; c++) col[c] = +st;
    });
    return function (r, c, addr) {
      if (addr in cell) return cell[addr];
      if (r in row) return row[r];
      return col[c] != null ? col[c] : 0;
    };
  }

  /* ---------- 字型名稱：Excel 常見的中文名稱 → Illustrator 的字型家族 ---------- */

  var FONT_ALIASES = {
    '新細明體': ['PMingLiU'], '細明體': ['MingLiU'], '標楷體': ['DFKai-SB', 'BiauKai'],
    '微軟正黑體': ['Microsoft JhengHei'], '微軟正黑體 Light': ['Microsoft JhengHei Light'],
    '微软雅黑': ['Microsoft YaHei'], '微軟雅黑': ['Microsoft YaHei'], '宋体': ['SimSun'], '新宋体': ['NSimSun'],
    '黑体': ['SimHei'], '楷体': ['KaiTi'], '仿宋': ['FangSong'], '等线': ['DengXian'], '等線': ['DengXian'],
    '游ゴシック': ['Yu Gothic'], '游明朝': ['Yu Mincho'], 'メイリオ': ['Meiryo'],
    'ＭＳ Ｐゴシック': ['MS PGothic'], 'ＭＳ ゴシック': ['MS Gothic'], 'ＭＳ 明朝': ['MS Mincho'], '맑은 고딕': ['Malgun Gothic']
  };

  /* ---------- 框線 ---------- */

  var BORDER = {
    hair: { w: 0.4 }, thin: { w: 0.75 }, medium: { w: 1.5 }, thick: { w: 2.25 },
    double: { w: 0.75, double: true },
    dotted: { w: 0.75, dash: [0.75, 1.5] }, dashed: { w: 0.75, dash: [2.25, 1.5] },
    mediumDashed: { w: 1.5, dash: [4.5, 2.25] },
    dashDot: { w: 0.75, dash: [3, 1.5, 0.75, 1.5] }, mediumDashDot: { w: 1.5, dash: [4.5, 2.25, 1.5, 2.25] },
    dashDotDot: { w: 0.75, dash: [3, 1.5, 0.75, 1.5, 0.75, 1.5] },
    mediumDashDotDot: { w: 1.5, dash: [4.5, 2.25, 1.5, 2.25, 1.5, 2.25] },
    slantDashDot: { w: 1.5, dash: [4.5, 1.5, 1.5, 1.5] }
  };

  function weight(b) {
    var d = BORDER[b.style] || BORDER.thin;
    return d.w + (d.double ? 1.5 : 0) + (d.dash ? 0 : 0.01);
  }

  /* ---------- 建立表格資料 ---------- */

  function buildTable(wb, sel, opts) {
    opts = opts || {};
    var sheetName = sel.sheet || wb.SheetNames[0];
    var ws = wb.Sheets[sheetName];
    if (!ws) throw new Error('找不到工作表：' + sheetName);

    var count = (sel.e.r - sel.s.r + 1) * (sel.e.c - sel.s.c + 1);
    if (count > MAX_CELLS) throw new Error('範圍太大（' + count + ' 格），一次最多 ' + MAX_CELLS + ' 格');

    var st = parseStyles(wb);
    var xml = sheetXml(wb, sheetName);
    var styleAt = sheetStyleIndex(xml);
    var fmt = tags(xml, 'sheetFormatPr')[0];
    var baseFont = st.fonts[0];
    var mdw = Math.round(baseFont.sz * 0.64) || 7;                // 預設字型的數字寬度 (px)
    var defColW = parseFloat(attr(fmt, 'defaultColWidth'));
    var baseColW = parseFloat(attr(fmt, 'baseColWidth')) || 8;
    var defColPx = defColW ? defColW * mdw : Math.ceil((baseColW * mdw + 5) / 8) * 8;
    var defRowPt = parseFloat(attr(fmt, 'defaultRowHeight')) || Math.round(baseFont.sz * 1.36 * 4) / 4;

    var cols = ws['!cols'] || [], rows = ws['!rows'] || [];
    var colW = function (c) {
      var k = cols[c];
      if (k && k.hidden) return 0;
      return ((k && k.width != null) ? k.width * mdw : defColPx) * PX;
    };
    var rowH = function (r) {
      var k = rows[r];
      if (k && k.hidden) return 0;
      return (k && k.hpt) ? k.hpt : defRowPt;
    };

    var xs = {}, ys = {}, x = 0, y = 0, r, c;
    for (c = sel.s.c; c <= sel.e.c + 1; c++) { xs[c] = x; if (c <= sel.e.c) x += colW(c); }
    for (r = sel.s.r; r <= sel.e.r + 1; r++) { ys[r] = y; if (r <= sel.e.r) y += rowH(r); }

    // 合併儲存格
    var span = {}, mergeOf = {};
    (ws['!merges'] || []).forEach(function (m, mi) {
      var s = { r: Math.max(m.s.r, sel.s.r), c: Math.max(m.s.c, sel.s.c) };
      var e = { r: Math.min(m.e.r, sel.e.r), c: Math.min(m.e.c, sel.e.c) };
      if (s.r > e.r || s.c > e.c) return;
      span[s.r + ',' + s.c] = e;
      for (var rr = s.r; rr <= e.r; rr++) for (var cc = s.c; cc <= e.c; cc++) mergeOf[rr + ',' + cc] = mi;
    });
    var sameMerge = function (r1, c1, r2, c2) {
      var a = mergeOf[r1 + ',' + c1], b = mergeOf[r2 + ',' + c2];
      return a != null && a === b;
    };
    var xfAt = function (r, c) {
      var addr = XLSX.utils.encode_cell({ r: r, c: c });
      return st.xfs[styleAt(r, c, addr)] || st.xfs[0];
    };

    /* 文字與填色 */
    var prefix = quoteSheet(sheetName) + '!';
    var cells = [];
    for (r = sel.s.r; r <= sel.e.r; r++) {
      for (c = sel.s.c; c <= sel.e.c; c++) {
        var key = r + ',' + c;
        if (mergeOf[key] != null && !span[key]) continue;          // 合併範圍中被蓋住的格
        var end = span[key] || { r: r, c: c };
        var w = xs[end.c + 1] - xs[c], h = ys[end.r + 1] - ys[r];
        if (w <= 0 || h <= 0) continue;                          // 隱藏的欄或列

        var addr = XLSX.utils.encode_cell({ r: r, c: c });
        var cell = ws[addr];
        var xf = xfAt(r, c);
        var font = st.fonts[xf.font] || baseFont;

        var ha = xf.ha;
        if (!ha || ha === 'general') {
          var t = cell && cell.t;
          ha = (t === 'n' || t === 'd') ? 'right' : (t === 'b' || t === 'e') ? 'center' : 'left';
        } else if (ha === 'centerContinuous' || ha === 'distributed') ha = 'center';
        else if (ha !== 'center' && ha !== 'right') ha = 'left';

        cells.push({
          ref: prefix + addr,
          text: cellText(ws, addr),
          x: xs[c], y: ys[r], w: w, h: h,
          ha: ha, va: xf.va === 'top' ? 'top' : xf.va === 'center' ? 'center' : 'bottom',
          wrap: xf.wrap, indent: xf.indent * mdw * 1.5 * PX,
          fill: st.fills[xf.fill] || null,
          font: {
            names: [font.name].concat(FONT_ALIASES[font.name] || []),
            sz: font.sz, b: font.b, i: font.i, u: font.u, st: font.st, color: font.color
          }
        });
      }
    }

    /* 框線：相鄰兩格共用的邊取較粗的那條，再把同一直線上的連續線段接起來 */
    var edges = {};
    var put = function (k, b, x1, y1, x2, y2) {
      if (!b || (x1 === x2 && y1 === y2)) return;
      if (!edges[k] || weight(b) > weight(edges[k].b)) edges[k] = { b: b, x1: x1, y1: y1, x2: x2, y2: y2 };
    };
    for (r = sel.s.r; r <= sel.e.r; r++) {
      for (c = sel.s.c; c <= sel.e.c; c++) {
        if (colW(c) === 0 || rowH(r) === 0) continue;
        var bd = st.borders[xfAt(r, c).border];
        if (!bd) continue;
        var x0 = xs[c], x1 = xs[c + 1], y0 = ys[r], y1 = ys[r + 1];
        if (!(r > sel.s.r && sameMerge(r, c, r - 1, c))) put('h' + r + ',' + c, bd.top, x0, y0, x1, y0);
        if (!(r < sel.e.r && sameMerge(r, c, r + 1, c))) put('h' + (r + 1) + ',' + c, bd.bottom, x0, y1, x1, y1);
        if (!(c > sel.s.c && sameMerge(r, c, r, c - 1))) put('v' + r + ',' + c, bd.left, x0, y0, x0, y1);
        if (!(c < sel.e.c && sameMerge(r, c, r, c + 1))) put('v' + r + ',' + (c + 1), bd.right, x1, y0, x1, y1);
      }
    }
    var segs = Object.keys(edges).map(function (k) { return edges[k]; });
    segs.sort(function (a, b) {
      var ha = a.y1 === a.y2, hb = b.y1 === b.y2;
      if (ha !== hb) return ha ? -1 : 1;
      return ha ? (a.y1 - b.y1 || a.x1 - b.x1) : (a.x1 - b.x1 || a.y1 - b.y1);
    });
    var borders = [];
    segs.forEach(function (s) {
      var last = borders[borders.length - 1];
      var same = last && last.style === s.b.style && last.color === s.b.color;
      if (same && s.y1 === s.y2 && last.y1 === last.y2 && last.y1 === s.y1 && Math.abs(last.x2 - s.x1) < 0.01) { last.x2 = s.x2; return; }
      if (same && s.x1 === s.x2 && last.x1 === last.x2 && last.x1 === s.x1 && Math.abs(last.y2 - s.y1) < 0.01) { last.y2 = s.y2; return; }
      var d = BORDER[s.b.style] || BORDER.thin;
      borders.push({ style: s.b.style, color: s.b.color, w: d.w, dash: d.dash || null, double: !!d.double,
                     x1: s.x1, y1: s.y1, x2: s.x2, y2: s.y2 });
    });
    borders.forEach(function (b) { delete b.style; });

    return {
      label: prefix + XLSX.utils.encode_range(sel.s, sel.e),
      width: x, height: y,
      grid: !!opts.grid,
      cells: cells,
      borders: borders
    };
  }

  return {
    MAX_CELLS: MAX_CELLS,
    quoteSheet: quoteSheet,
    parseRange: parseRange,
    cellText: cellText,
    savedSelection: savedSelection,
    buildTable: buildTable
  };
};
