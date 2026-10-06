/* CMF Tool — 標註分頁 */
(function () {
  var App = window.CMFApp;
  var cep = window.__adobe_cep__;
  var KEY = "cmfCalloutSettings";
  var STYLE_FIELDS = ["fontSize", "textColor", "fontName", "badge", "badgeColor", "badgeTextColor", "badgePadding",
    "badgeStrokeWidth", "badgeStrokeColor", "lineWidth", "lineColor", "endStyle", "endSize", "styleName", "gap"];
  var FIELDS = ["lineMode", "continuous", "startMode", "startNumber", "sortMode", "sameNumber", "autoSync"].concat(STYLE_FIELDS);
  var MODE_TIPS = {
    straight: "點目標點 → 點編號位置",
    elbow: "點目標點 → 點轉折點 → 點編號位置",
    free: "左鍵持續加點，按「完成這條線」或快捷鍵完成"
  };
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
      else if (el.type === "number") o[k] = parseFloat(el.value) || 0;
      else o[k] = el.value;
    });
    o.hostPath = extensionPath() + "/jsx/host.jsx";
    return o;
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
      done();
      return;
    }
    run("CMF.loadSettings()", function (r) {
      if (r && r.ok && r.settings) applySettings(r.settings);
      done();
    }, true);
  }

  var saveTimer = null;
  function saveSettings() {
    var o = readSettings();
    try { localStorage.setItem(KEY, JSON.stringify(o)); } catch (e) {}
    clearTimeout(saveTimer);
    saveTimer = setTimeout(function () { if (cep) callHost("saveSettings", o, null, true); }, 400);
  }

  // ---------- 預覽 ----------
  function isDark(hex) {
    var h = hex.replace("#", "");
    var r = parseInt(h.substr(0, 2), 16), g = parseInt(h.substr(2, 2), 16), b = parseInt(h.substr(4, 2), 16);
    return (0.299 * r + 0.587 * g + 0.114 * b) < 60;
  }

  function renderPreview() {
    var o = readSettings();
    $("badgeRows").classList.toggle("hidden", !o.badge);
    $("styleRow").classList.toggle("hidden", o.endStyle !== "style");
    $("startNumber").disabled = o.startMode === "auto";
    $("badgeStrokeColor").disabled = !(o.badgeStrokeWidth > 0);
    updateFinishButton();

    var scale = 1.6;
    // 編號端 → (轉折) → 目標點
    var pts = o.lineMode === "straight"
      ? [[70, 30], [196, 52]]
      : [[62, 24], [128, 24], [196, 52]];
    var S = pts[0], S1 = pts[1], E = pts[pts.length - 1], E1 = pts[pts.length - 2];
    var lw = Math.max(0.6, o.lineWidth * scale), es = o.endSize * scale;
    var parts = ['<path d="M178 84 C 186 50, 210 26, 260 20 L260 84 Z" fill="var(--product)"/>'];

    var ux = E[0] - E1[0], uy = E[1] - E1[1], L = Math.sqrt(ux * ux + uy * uy);
    ux /= L; uy /= L;
    var lineEnd = E;
    if (o.endStyle === "arrow") {
      var len = es * 1.6, hw = es * 0.5, bx = E[0] - ux * len, by = E[1] - uy * len;
      lineEnd = [bx, by];
      parts.push('<path d="M' + E[0] + ' ' + E[1] + ' L' + (bx - uy * hw) + ' ' + (by + ux * hw) +
        ' L' + (bx + uy * hw) + ' ' + (by - ux * hw) + 'Z" fill="' + o.lineColor + '"/>');
    }
    var poly = pts.slice(0, -1).concat([lineEnd]).map(function (p) { return p.join(","); }).join(" ");
    parts.push('<polyline points="' + poly + '" fill="none" stroke-linejoin="round" stroke="' + o.lineColor +
      '" stroke-width="' + lw + '"/>');
    if (o.endStyle === "dot") {
      parts.push('<circle cx="' + E[0] + '" cy="' + E[1] + '" r="' + es / 2 + '" fill="' + o.lineColor + '"/>');
    }

    // 編號沿第一段的反方向推出去
    var dx = S[0] - S1[0], dy = S[1] - S1[1], dl = Math.sqrt(dx * dx + dy * dy);
    dx /= dl; dy /= dl;
    var fs = Math.max(6, o.fontSize * scale), gap = o.gap * scale;
    var pf = previewFont(o);
    var font = pf.family + '" font-weight="' + pf.weight + '" font-style="' + (pf.italic ? "italic" : "normal");
    if (o.badge) {
      var R = fs * 0.42 + o.badgePadding * scale;
      var bsw = (o.badgeStrokeWidth || 0) * scale;
      var cx = S[0] + dx * (gap + R + bsw / 2), cy = S[1] + dy * (gap + R + bsw / 2);
      parts.push('<circle cx="' + cx + '" cy="' + cy + '" r="' + R + '" fill="' + o.badgeColor + '"' +
        (bsw > 0 ? ' stroke="' + o.badgeStrokeColor + '" stroke-width="' + bsw + '"' : '') + '/>');
      parts.push('<text x="' + cx + '" y="' + cy + '" fill="' + o.badgeTextColor + '" font-size="' + fs +
        '" font-family="' + font + '" text-anchor="middle" dominant-baseline="central">3</text>');
    } else {
      var w = fs * 0.3, h = fs * 0.36;
      var ext = Math.min(Math.abs(dx) > 0.001 ? w / Math.abs(dx) : 1e9, Math.abs(dy) > 0.001 ? h / Math.abs(dy) : 1e9);
      var tx = S[0] + dx * (gap + ext), ty = S[1] + dy * (gap + ext);
      var light = document.documentElement.classList.contains("light");
      var tc = light || !isDark(o.textColor) ? o.textColor : "#e6e6e6";
      parts.push('<text x="' + tx + '" y="' + ty + '" fill="' + tc + '" font-size="' + fs +
        '" font-family="' + font + '" text-anchor="middle" dominant-baseline="central">3</text>');
    }
    $("preview").innerHTML = parts.join("");
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

  // 預覽用：PostScript 名稱 → CSS 字體設定
  function previewFont(o) {
    var hit = fontIndex[o.fontName];
    if (!hit) return { family: "sans-serif", weight: 400, italic: false };
    var st = hit[1];
    var weight = /thin|hairline|w1/i.test(st) ? 200 : /extralight|ultralight|light|w2|w3/i.test(st) ? 300
      : /semibold|demibold|w6/i.test(st) ? 600 : /extrabold|ultrabold|heavy|black|w8|w9/i.test(st) ? 800
      : /bold|w7/i.test(st) ? 700 : /medium|w5/i.test(st) ? 500 : 400;
    return { family: "'" + hit[0].replace(/'/g, "") + "', sans-serif", weight: weight, italic: /italic|oblique/i.test(st) };
  }

  // ---------- 取色 ----------
  // 每個顏色欄位旁加一個滴管按鈕：按下後點畫布上的物件，取它的顏色
  var DROPPER = '<svg viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><path d="M13.7 2.3a2 2 0 0 0-2.8 0L9.2 4 8.5 3.3 7.1 4.7l.7.7-5.1 5.1-.5 2.4-.9.9 1.4 1.4.9-.9 2.4-.5 5.1-5.1.7.7 1.4-1.4-.7-.7 1.7-1.7a2 2 0 0 0 0-2.8zM5.1 12.3l-1.3.3.3-1.3 5.1-5.1 1 1z"/></svg>';
  var STROKE_FIRST = { lineColor: true, badgeStrokeColor: true };
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
      }
      stopPick(r.msg, !r.ok);
    }, true);
  }

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
    $("btnFinish").classList.toggle("hidden", !(adding && $("lineMode").value === "free"));
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
    callHost("beginAdd", null, function (r) {
      if (!r.ok) return;
      adding = true;
      idleSince = Date.now();
      $("btnAdd").textContent = "結束新增";
      $("btnAdd").classList.add("active");
      updateFinishButton();
      if (!r.msg) setStatus(MODE_TIPS[$("lineMode").value]);
      pollTimer = setInterval(poll, 200);
    });
  }

  function stopAdd(keepTool) {
    clearInterval(pollTimer);
    pollTimer = null;
    adding = false;
    $("btnAdd").textContent = "新增標註";
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

  // ---------- 樣式自動同步 ----------
  var syncTimer = null;
  function scheduleSync() {
    if (!$("autoSync").checked || !cep) return;
    clearTimeout(syncTimer);
    syncTimer = setTimeout(function () { callHost("syncAll", readSettings()); }, 700);
  }

  // ---------- 事件 ----------
  FIELDS.forEach(function (k) {
    var isStyle = STYLE_FIELDS.indexOf(k) >= 0;
    var handler = function () {
      renderPreview();
      saveSettings();
      if (isStyle) scheduleSync();
    };
    $(k).addEventListener("input", handler);
    $(k).addEventListener("change", handler);
  });

  $("btnAdd").addEventListener("click", function () { if (adding) stopAdd(false); else startAdd(); });
  $("btnFinish").addEventListener("click", finishFree);
  document.addEventListener("keydown", function (e) {
    if (e.key === "Escape" && picking) { stopPick("已取消取色"); return; }
    if (e.key === "Escape" && adding && document.activeElement.tagName !== "INPUT") stopAdd(false);
  });

  $("btnConvert").addEventListener("click", function () { callHost("convert", readSettings(), afterAdded); });
  $("btnRenumber").addEventListener("click", function () { callHost("renumber", readSettings(), refreshTable); });
  $("btnRefresh").addEventListener("click", refreshTable);
  $("btnLinkExcel").addEventListener("click", function () {
    App.showTab("excel");
    var box = $("cmfBox");
    if (box) { box.open = true; box.scrollIntoView(); }
  });
  $("btnSync").addEventListener("click", function () { callHost("syncAll", readSettings()); });
  $("btnRestyle").addEventListener("click", function () { callHost("restyle", readSettings()); });
  $("btnRelayout").addEventListener("click", function () { callHost("relayout"); });
  $("btnSelectAll").addEventListener("click", function () { callHost("selectAll"); });
  $("btnToggle").addEventListener("click", function () { callHost("toggleLayer"); });
  $("btnSetNumber").addEventListener("click", function () {
    callHost("setNumber", { num: parseInt($("setNumberValue").value, 10) }, refreshTable);
  });

  // 滑鼠移回面板或切換文件時，更新編號表（使用者可能在畫布上刪除或複製了標註）
  var lastEnter = 0;
  document.body.addEventListener("mouseenter", function () {
    if (adding || Date.now() - lastEnter < 800) return;
    lastEnter = Date.now();
    refreshTable();
  });

  if (cep) {
    try { cep.addEventListener("com.cmf.callout.add", function () { if (adding) stopAdd(false); else startAdd(); }); } catch (e) {}
    try { cep.addEventListener("com.cmf.callout.finish", finishFree); } catch (e) {}
    try { cep.addEventListener("documentAfterActivate", refreshTable); } catch (e) {}
  }
  App.on("theme", renderPreview);
  App.on("numbers", refreshTable);   // Excel 分頁改了 CMF 清單或自動對應之後

  loadSettings(function () {
    loadFonts(false);
    renderPreview();
    saveSettings();
    refreshTable();
  });
  if (!cep) setStatus("預覽模式：請在 Illustrator 中使用", true);
})();
