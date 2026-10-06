/* CMF Tool — Excel 分頁（CEP，含 Node.js）：綁定文字、匯入表格、CMF 清單依標註排序 */
(function () {
  'use strict';

  var App = window.CMFApp;
  var cs = new CSInterface();
  var nodeRequire = (typeof cep_node !== 'undefined' && cep_node.require) ? cep_node.require : require;
  var fs = nodeRequire('fs');
  var path = nodeRequire('path');
  var EXT = cs.getSystemPath(SystemPath.EXTENSION);
  var XLSX = nodeRequire(path.join(EXT, 'lib', 'xlsx.full.min.js'));
  var Table = nodeRequire(path.join(EXT, 'js', 'table.js'))(XLSX);
  var os = nodeRequire('os');
  var childProcess = nodeRequire('child_process');
  var IS_WIN = os.platform() === 'win32';

  var $ = function (id) { return document.getElementById(id); };
  var ui = {
    fileName: $('fileName'), fileDir: $('fileDir'), btnPick: $('btnPick'),
    fileChoices: $('fileChoices'), fileSource: $('fileSource'), btnAuto: $('btnAuto'),
    selInfo: $('selInfo'), cellRef: $('cellRef'), btnBind: $('btnBind'), btnUnbind: $('btnUnbind'),
    btnUpdate: $('btnUpdate'), chkAuto: $('chkAuto'), result: $('result'),
    refList: $('refList'), refCount: $('refCount'),
    rangeRef: $('rangeRef'), btnImport: $('btnImport'), chkGrid: $('chkGrid'),
    cmfTable: $('cmfTable'), cmfFields: $('cmfFields'), cmfNum: $('cmfNum'), cmfName: $('cmfName'),
    cmfHead: $('cmfHead'), cmfRest: $('cmfRest'), cmfAuto: $('cmfAuto'), cmfInfo: $('cmfInfo'),
    btnCmfMatch: $('btnCmfMatch')
  };

  var state = {
    docKey: null, docFolder: '', excelPath: null,
    source: null,            // 'manual' = 使用者選的；'auto' = 從 .ai 同資料夾找到的
    candidates: [], candidatesKey: null,
    selectedRefs: [], tables: [], busy: false, watching: null, debounce: null
  };

  // CMF 清單：sig = 每份文件目前「編號=物件」的對應，變了才重排
  var cmf = { sig: {}, timer: null, pending: false, book: null };

  /* ---------- 呼叫 Illustrator ---------- */

  function lit(v) {
    return JSON.stringify(v).replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
  }

  function host(fn, args) {
    var call = fn + '(' + (args || []).map(lit).join(',') + ')';
    return new Promise(function (resolve, reject) {
      cs.evalScript(call, function (res) {
        if (res === 'EvalScript error.') return reject(new Error('Illustrator 腳本執行失敗（' + fn + '）'));
        try { resolve(JSON.parse(res)); } catch (e) { reject(new Error('無法解析 Illustrator 回傳：' + res)); }
      });
    });
  }

  // Illustrator 端的 Excel 函式（manifest 只能指定一個腳本，標註用的 host.jsx 由 CEP 載入）
  cs.evalScript('$.evalFile(' + lit(EXT + '/jsx/excel.jsx') + ')');

  /* ---------- 儲存格位址 ---------- */

  // "B3" / "$B$3" / "工作表1!B3" / "'價格 表'!C12"
  function parseRef(ref) {
    var i = ref.lastIndexOf('!');
    var sheet = null, cell = ref;
    if (i >= 0) {
      sheet = ref.slice(0, i).replace(/^'(.*)'$/, '$1').replace(/''/g, "'");
      cell = ref.slice(i + 1);
    }
    cell = cell.replace(/\$/g, '').toUpperCase();
    if (!/^[A-Z]{1,3}[1-9][0-9]{0,6}$/.test(cell) || sheet === '') return null;
    return { sheet: sheet, cell: cell };
  }

  function normalizeRef(input) {
    var p = parseRef(input.trim());
    if (!p) return null;
    return p.sheet ? Table.quoteSheet(p.sheet) + '!' + p.cell : p.cell;
  }

  /* ---------- 讀 Excel ---------- */

  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  // Excel 存檔的瞬間檔案可能還在寫入，失敗就稍等重試
  function readWorkbook(file, withLayout) {
    var attempt = 0;
    // withLayout：連欄寬、列高、數字格式和原始 XML 一起讀（匯入表格時需要）
    var opts = withLayout ? { type: 'buffer', cellStyles: true, cellNF: true, bookFiles: true } : { type: 'buffer' };
    function tryRead() {
      try {
        return Promise.resolve(XLSX.read(fs.readFileSync(file), opts));
      } catch (err) {
        if (++attempt >= 4) return Promise.reject(err);
        return sleep(500).then(tryRead);
      }
    }
    return tryRead();
  }

  // 編號表每次滑鼠移進面板都會讀物件名稱，檔案沒變就用上次讀的
  function cachedWorkbook(file) {
    var st;
    try { st = fs.statSync(file); } catch (e) { return Promise.reject(new Error('找不到檔案：' + path.basename(file))); }
    var key = file + '|' + st.mtimeMs + '|' + st.size;
    if (cmf.book && cmf.book.key === key) return Promise.resolve(cmf.book.wb);
    return readWorkbook(file, true).then(function (wb) {
      cmf.book = { key: key, wb: wb };
      return wb;
    });
  }

  function friendly(err) {
    if (err && (err.code === 'EBUSY' || err.code === 'EPERM' || err.code === 'EACCES')) {
      return new Error('Excel 檔案暫時無法讀取，請確認已存檔後再試一次');
    }
    return err;
  }

  function collectValues(wb, refs) {
    var values = {}, missing = [], empty = [];
    refs.forEach(function (ref) {
      var p = parseRef(ref);
      var sheetName = p && (p.sheet || wb.SheetNames[0]);
      var ws = p && wb.Sheets[sheetName];
      if (!ws) { missing.push(ref); return; }
      var text = Table.cellText(ws, p.cell);
      if (text === '') empty.push(ref);
      values[ref] = text;
    });
    return { values: values, missing: missing, empty: empty };
  }

  /* ---------- 每份文件記住自己的 Excel 檔 ---------- */

  function storageKey(docKey) { return 'excelsync:path:' + docKey; }

  function loadPathFor(docKey) {
    try { return localStorage.getItem(storageKey(docKey)); } catch (e) { return null; }
  }

  function savePathFor(docKey, file) {
    try { localStorage.setItem(storageKey(docKey), file); } catch (e) {}
  }

  function clearPathFor(docKey) {
    try { localStorage.removeItem(storageKey(docKey)); } catch (e) {}
  }

  /* ---------- 自動尋找 .ai 同資料夾的 Excel ---------- */

  function findExcelNear(docKey, folder) {
    if (!folder) return [];
    var names;
    try { names = fs.readdirSync(folder); } catch (e) { return []; }
    var base = path.basename(docKey, path.extname(docKey)).toLowerCase();
    return names
      .filter(function (n) { return /\.(xlsx|xlsm|xls)$/i.test(n) && n.indexOf('~$') !== 0; }) // ~$ 開頭是 Excel 開啟中的暫存檔
      .map(function (n) {
        return { full: path.join(folder, n), name: n,
                 same: path.basename(n, path.extname(n)).toLowerCase() === base };
      })
      .sort(function (a, b) { return (b.same - a.same) || a.name.localeCompare(b.name); });
  }

  // 決定這份文件要用哪個 Excel：手動選過的優先，否則自動找
  function resolveExcel(ctx) {
    var docChanged = ctx.doc !== state.docKey;
    state.docKey = ctx.doc;
    state.docFolder = ctx.folder;

    var stored = loadPathFor(ctx.doc);
    if (stored && fs.existsSync(stored)) {
      if (docChanged || state.excelPath !== stored || state.source !== 'manual') {
        state.candidates = [];
        state.candidatesKey = null;
        setExcelPath(stored, 'manual');
      }
      return;
    }

    var found = findExcelNear(ctx.doc, ctx.folder);
    var key = found.map(function (f) { return f.full; }).join('|');
    if (!docChanged && state.source === 'auto' && key === state.candidatesKey) return;  // 沒變化

    state.candidates = found;
    state.candidatesKey = key;
    var pick = null;
    if (found.length === 1) pick = found[0].full;          // 只有一個：直接用
    else if (found.length && found[0].same) pick = found[0].full;  // 多個：用跟 .ai 同名的
    setExcelPath(pick, 'auto');                            // 多個又沒有同名的：讓使用者從清單選
  }

  /* ---------- 畫面 ---------- */

  function showFile() {
    var f = state.excelPath;
    if (!f) {
      ui.fileName.textContent = '尚未選擇檔案';
      ui.fileName.classList.add('is-empty');
      ui.fileName.title = '';
      ui.fileDir.textContent = '';
      ui.fileDir.title = '';
    } else {
      ui.fileName.textContent = path.basename(f);
      ui.fileName.classList.toggle('is-empty', false);
      ui.fileName.title = f;
      ui.fileDir.title = path.dirname(f);
      ui.fileDir.textContent = path.basename(path.dirname(f));   // 只顯示資料夾名稱，完整路徑在提示裡
    }

    var msg = '';
    if (!state.docKey) msg = '';
    else if (state.source === 'manual') msg = '手動選擇';
    else if (!state.docFolder) msg = '.ai 存檔後會自動找同資料夾的 Excel';
    else if (f) msg = '自動：與 .ai 同資料夾';
    else if (state.candidates.length > 1) msg = '資料夾裡有 ' + state.candidates.length + ' 個 Excel，請選一個';
    else msg = '同資料夾沒有 Excel';
    ui.fileSource.textContent = msg;
    ui.btnAuto.hidden = !(state.source === 'manual' && state.docFolder);

    // 同資料夾有多個 Excel 時顯示清單
    var showList = state.source === 'auto' && state.candidates.length > 1;
    ui.fileChoices.hidden = !showList;
    if (showList) {
      ui.fileChoices.innerHTML = '';
      if (!f) ui.fileChoices.appendChild(new Option('選擇 Excel 檔…', ''));
      state.candidates.forEach(function (c) {
        ui.fileChoices.appendChild(new Option(c.name, c.full, false, c.full === f));
      });
    }
  }

  function showResult(lines) {
    ui.result.innerHTML = '';
    lines.forEach(function (l) {
      var p = document.createElement('p');
      p.className = l[0];
      p.textContent = l[1];
      ui.result.appendChild(p);
    });
    ui.result.hidden = false;
  }

  function showError(err) {
    var msg = (err && err.message) || String(err);
    showResult([['error', msg]]);
    App.setStatus(msg, true);
  }

  function chip(text, cls, onClick, title) {
    var b = document.createElement('button');
    b.type = 'button';
    b.className = cls;
    b.textContent = text;
    if (title) b.title = title;
    b.addEventListener('click', onClick);
    ui.refList.appendChild(b);
  }

  function renderRefs(refs, tables) {
    tables = tables || [];
    ui.refList.innerHTML = '';
    var total = refs.length + tables.length;
    ui.refCount.textContent = total ? total : '';
    if (!total) {
      ui.refList.innerHTML = '<span class="muted">沒有綁定</span>';
      return;
    }
    tables.forEach(function (t) {
      var label = t.label;
      chip('表格 ' + label, 'ref-chip is-table' + (t.cmf ? ' is-cmf' : ''), function () {
        host('es_selectTable', [label]).then(refreshContext).catch(showError);
      }, t.cmf ? 'CMF 清單（依標註排序）' : '點一下選取');
    });
    var LIMIT = 40, sorted = refs.slice().sort(compareRefs);
    sorted.slice(0, LIMIT).forEach(function (ref) {
      chip(ref, 'ref-chip' + (state.selectedRefs.indexOf(ref) >= 0 ? ' is-selected' : ''), function () {
        host('es_selectRef', [ref]).then(refreshContext).catch(showError);
      }, '點一下選取');
    });
    if (sorted.length > LIMIT) {
      var more = document.createElement('span');
      more.className = 'muted';
      more.textContent = '還有 ' + (sorted.length - LIMIT) + ' 個';
      ui.refList.appendChild(more);
    }
  }

  // 依工作表、欄、列排序
  function compareRefs(a, b) {
    var pa = parseRef(a), pb = parseRef(b);
    if (!pa || !pb) return a < b ? -1 : 1;
    var sa = pa.sheet || '', sb = pb.sheet || '';
    if (sa !== sb) return sa < sb ? -1 : 1;
    var ca = XLSX.utils.decode_cell(pa.cell), cb = XLSX.utils.decode_cell(pb.cell);
    return ca.c - cb.c || ca.r - cb.r;
  }

  // 同時有好幾個地方要求更新時，共用同一次
  var ctxPromise = null;
  function refreshContext() {
    if (!ctxPromise) {
      ctxPromise = loadContext().then(function (r) { ctxPromise = null; return r; },
                                      function (e) { ctxPromise = null; throw e; });
    }
    return ctxPromise;
  }

  function loadContext() {
    return Promise.all([host('es_context'), host('es_refs')]).then(function (r) {
      var ctx = r[0], all = r[1];
      if (!ctx.doc) {
        if (state.docKey || state.excelPath) {
          state.docKey = null;
          state.docFolder = '';
          state.candidates = [];
          state.candidatesKey = null;
          setExcelPath(null, null);
        }
        state.tables = [];
        ui.selInfo.textContent = '請先開啟 Illustrator 文件';
        renderRefs([]);
        renderCmf([]);
        return { refs: [], tables: [] };
      }
      resolveExcel(ctx);
      state.selectedRefs = ctx.refs;
      state.tables = all.tables;
      if (ctx.selected === 0) ui.selInfo.textContent = '請在 Illustrator 中選取文字物件';
      else if (ctx.refs.length === 0) ui.selInfo.textContent = '已選取 ' + ctx.selected + ' 個文字，尚未綁定';
      else ui.selInfo.textContent = '已選取 ' + ctx.selected + ' 個文字：' + ctx.refs.join('、');
      if (ctx.refs.length === 1 && document.activeElement !== ui.cellRef) ui.cellRef.value = ctx.refs[0];
      renderRefs(all.refs, all.tables);
      renderCmf(all.tables);
      return all;
    });
  }

  /* ---------- 動作 ---------- */

  function setExcelPath(file, source) {
    state.excelPath = file || null;
    state.source = source;
    if (source === 'manual' && state.docKey && file) savePathFor(state.docKey, file);
    showFile();
    syncWatcher();
  }

  function useAuto() {
    if (!state.docKey) return;
    clearPathFor(state.docKey);
    state.source = null;
    state.candidatesKey = null;
    ui.result.hidden = true;
    refreshContext().then(afterFileChanged).catch(showError);
  }

  // 換了 Excel 檔：標註分頁的物件清單也要換
  function afterFileChanged() { App.emit('numbers'); }

  function pickFile() {
    var startDir = state.excelPath ? path.dirname(state.excelPath) : (state.docFolder || '');
    var dlg = window.cep.fs.showOpenDialogEx || window.cep.fs.showOpenDialog;
    var res = dlg(false, false, '選擇 Excel 檔案', startDir, ['xlsx', 'xlsm', 'xls']);
    if (res.err || !res.data || !res.data.length) return;
    if (!state.docKey) { showError(new Error('請先開啟 Illustrator 文件')); return; }
    setExcelPath(res.data[0], 'manual');
    ui.result.hidden = true;
    afterFileChanged();
  }

  function bind() {
    var ref = normalizeRef(ui.cellRef.value);
    if (!ref) { showError(new Error('儲存格位址格式不對，例如 B3 或 工作表1!B3')); ui.cellRef.focus(); return; }
    ui.cellRef.value = ref;
    host('es_bind', [ref]).then(function (r) {
      if (r.error) throw new Error(r.error);
      App.setStatus('已將 ' + r.count + ' 個文字綁定到 ' + ref);
      return refreshContext();
    }).catch(showError);
  }

  function unbind() {
    host('es_unbind').then(function (r) {
      if (r.error) throw new Error(r.error);
      App.setStatus(r.count ? '已解除 ' + r.count + ' 個文字的綁定' : '選取的物件沒有綁定');
      return refreshContext();
    }).catch(showError);
  }

  /* ---------- CMF 清單 ---------- */

  function cmfTableOf(tables) {
    return (tables || []).filter(function (t) { return t.cmf; })[0] || null;
  }

  function rangeWidth(sel) { return sel.e.c - sel.s.c + 1; }

  // 標註編號 → 物件名稱
  function linksOf(nums) {
    return ((nums && nums.items) || []).filter(function (it) { return it.key; })
      .map(function (it) { return { num: it.num, key: it.key }; });
  }

  // CMF 清單的內容：物件名稱、欄位標題（給標註分頁的下拉選單和這裡的欄位選單）
  function cmfList(tables) {
    var t = cmfTableOf(tables);
    if (!t) return Promise.resolve(null);
    var sel = Table.parseRange(t.label);
    var info = { label: t.label, cfg: null, width: 1, names: [], titles: [], head: 0, error: null };
    if (!sel) { info.error = '表格範圍不對'; return Promise.resolve(info); }
    info.width = rangeWidth(sel);
    info.cfg = Table.cmfConfig(t.cmf, info.width);
    if (!state.excelPath) { info.error = '找不到 Excel 檔'; return Promise.resolve(info); }
    return cachedWorkbook(state.excelPath).then(function (wb) {
      var rows = Table.cmfRows(wb, sel, info.cfg);
      info.names = Table.cmfNames(rows, info.cfg);
      info.head = Table.cmfHead(rows, info.cfg);
      info.titles = Table.columnTitles(wb, sel, info.head);
      return info;
    }).catch(function (err) {
      info.error = friendly(err).message;
      return info;
    });
  }

  // 依文件裡的表格更新 CMF 區塊
  var cmfRender = 0;
  function renderCmf(tables) {
    var t = cmfTableOf(tables);
    if (document.activeElement !== ui.cmfTable) {
      ui.cmfTable.innerHTML = '';
      ui.cmfTable.appendChild(new Option(tables.length ? '不使用' : '文件中沒有表格', ''));
      tables.forEach(function (x) { ui.cmfTable.appendChild(new Option(x.label, x.label)); });
      ui.cmfTable.value = t ? t.label : '';
      ui.cmfTable.disabled = !tables.length;
    }
    ui.cmfFields.hidden = !t;
    if (!t) return;

    var token = ++cmfRender;
    cmfList(tables).then(function (info) {
      if (token !== cmfRender || !info) return;
      var fill = function (select, value) {
        if (document.activeElement === select) return;
        select.innerHTML = '';
        for (var i = 0; i < info.width; i++) {
          var col = info.titles[i] || { col: '', title: '' };
          var text = (col.col || String(i + 1)) + (col.title ? '  ' + col.title : '');
          select.appendChild(new Option(text, String(i), false, i === value));
        }
      };
      fill(ui.cmfNum, info.cfg.num);
      fill(ui.cmfName, info.cfg.name);
      if (document.activeElement !== ui.cmfHead) ui.cmfHead.value = info.cfg.head == null ? '' : String(info.cfg.head);
      if (document.activeElement !== ui.cmfRest) ui.cmfRest.value = info.cfg.rest;
      ui.cmfInfo.textContent = info.error ? info.error : info.names.length + ' 個物件';
      ui.cmfInfo.title = info.error ? '' : info.names.map(function (n) { return n.name; }).join('、');
    });
  }

  function readCmfForm() {
    return {
      num: parseInt(ui.cmfNum.value, 10) || 0,
      name: ui.cmfName.value === '' ? 1 : parseInt(ui.cmfName.value, 10),
      head: ui.cmfHead.value === '' ? null : parseInt(ui.cmfHead.value, 10),
      rest: ui.cmfRest.value || 'continue'
    };
  }

  // 改了 CMF 設定：存到表格上，馬上重排一次；取消時把原本的表格改回 Excel 的排列
  function saveCmf(label, cfg, restore) {
    return host('es_setCmf', [label || '', cfg ? JSON.stringify(cfg) : '']).then(function (r) {
      if (r.error) throw new Error(r.error);
      return Promise.all([refreshContext(), rememberNumbers()]);   // 記下目前的對應，避免標註分頁重新整理時又排一次
    }).then(function () {
      App.emit('numbers');
      if (label) return update(false, label);
    }).then(function () {
      if (restore && restore !== label) return update(false, restore);   // 一次只能更新一個，依序執行
    }).catch(showError);
  }

  function onCmfTable() {
    var label = ui.cmfTable.value, prev = cmfTableOf(state.tables);
    var sel = label && Table.parseRange(label);
    saveCmf(label, label ? Table.cmfConfig({}, sel ? rangeWidth(sel) : 2) : null, prev && prev.label);
  }

  function onCmfField() {
    var t = cmfTableOf(state.tables);
    if (t) saveCmf(t.label, readCmfForm());
  }

  // 標註分頁每次重新整理編號表時呼叫：對應有變就重排表格
  function numbersChanged(nums) {
    if (!nums || !nums.doc) return;
    var sig = linksOf(nums).map(function (l) { return l.num + '=' + l.key; }).join('|');
    var prev = cmf.sig[nums.doc];
    cmf.sig[nums.doc] = sig;
    if (prev === undefined || prev === sig || !ui.cmfAuto.checked) return;
    if (!cmfTableOf(state.tables)) return;
    clearTimeout(cmf.timer);
    cmf.timer = setTimeout(function () { update(true, true); }, 500);
  }

  function rememberNumbers() {
    return host('CMF.listNumbers').then(function (nums) {
      if (nums && nums.doc) cmf.sig[nums.doc] = linksOf(nums).map(function (l) { return l.num + '=' + l.key; }).join('|');
      return nums;
    });
  }

  // 還沒對應的編號，對到 Excel 中序號相同的物件
  function autoMatch() {
    refreshContext().then(function (all) {
      return Promise.all([cmfList(all.tables), host('CMF.listNumbers')]);
    }).then(function (r) {
      var info = r[0], items = (r[1] && r[1].items) || [];
      if (!info) throw new Error('請先選擇 CMF 清單的表格');
      if (info.error) throw new Error(info.error);
      if (!items.length) throw new Error('文件中還沒有標註');
      var used = {}, bySerial = {}, links = [];
      items.forEach(function (it) { if (it.key) used[it.key] = true; });
      info.names.forEach(function (n) {
        var k = parseInt(n.serial, 10);
        if (k > 0 && /^\d/.test(n.serial) && !(k in bySerial)) bySerial[k] = n.name;
      });
      items.forEach(function (it) {
        var name = bySerial[it.num];
        if (!it.key && name && !used[name]) { used[name] = true; links.push({ num: it.num, key: name }); }
      });
      if (!links.length) { App.setStatus('沒有可以自動對應的編號'); return null; }
      return host('CMF.setLinks', [JSON.stringify({ links: links })]).then(function (res) {
        if (!res.ok) throw new Error(res.msg);
        App.setStatus(res.msg);
        return rememberNumbers();
      }).then(function () {
        App.emit('numbers');
        return update(false, true);
      });
    }).catch(showError);
  }

  // 排序結果的說明
  function cmfReport(plan, lines) {
    if (!plan.linked) {
      lines.push(['time', 'CMF 清單：還沒有對應的編號，維持 Excel 原本的排列']);
      return;
    }
    lines.push(['ok', 'CMF 清單已依標註排序（' + plan.linked + '／' + plan.total + ' 個物件有對應）']);
    if (plan.missing.length) {
      lines.push(['warn', '清單中沒有：' + plan.missing.map(function (m) { return m.num + ' ' + m.key; }).join('、')]);
    }
    if (plan.dup.length) lines.push(['warn', '對到多個編號，使用最小的：' + plan.dup.join('、')]);
  }

  /* ---------- 從 Excel 更新 ---------- */

  // only：只更新一個表格，不動綁定的文字。true = CMF 清單（標註編號改變時），或指定表格範圍
  function update(auto, only) {
    if (state.busy) {
      if (only === true) cmf.pending = true;
      return Promise.resolve();
    }
    if (!state.excelPath) {
      if (!auto) showError(new Error(state.candidates.length > 1 ? '請先從清單選一個 Excel 檔' : '請先選擇 Excel 檔案'));
      return Promise.resolve();
    }
    if (!fs.existsSync(state.excelPath)) { showError(new Error('找不到檔案：' + state.excelPath)); return Promise.resolve(); }

    state.busy = true;
    ui.btnUpdate.disabled = true;
    var file = state.excelPath;

    return Promise.all([host('es_refs'), host('CMF.listNumbers')]).then(function (res) {
      var r = res[0], links = linksOf(res[1]);
      if (!r.doc) throw new Error('請先開啟 Illustrator 文件');
      // 自動更新時，若使用者切到別的文件，就不要動它
      if (auto && r.doc !== state.docKey) return null;
      var cmfT = cmfTableOf(r.tables);
      var label = only === true ? (cmfT && cmfT.label) : only;
      var tables = only ? r.tables.filter(function (t) { return t.label === label; }) : r.tables;
      var refs = only ? [] : r.refs;
      if (only && !tables.length) return null;
      if (!refs.length && !tables.length) throw new Error('文件中還沒有綁定任何文字或表格');
      App.setStatus(only ? '更新表格中…' : '從 Excel 更新中…');

      var excelTables = tables.filter(function (t) { return t.mode === 'excel'; });
      var builtTables = tables.filter(function (t) { return t.mode !== 'excel'; });
      var cmfExcel = excelTables.filter(function (t) { return t.cmf; });
      var plans = {}, done = {};      // CMF 清單的排序；done = 表格已經換成新的
      var planFor = function (t, rows) {
        var sel = Table.parseRange(t.label);
        return (plans[t.label] = Table.cmfPlan(rows, Table.cmfConfig(t.cmf, rangeWidth(sel)), links));
      };

      return readWorkbook(file, builtTables.length > 0 || cmfExcel.length > 0).then(function (wb) {
        var lines = [['time', new Date().toLocaleTimeString() + (auto ? '（自動）' : '')]];
        var fonts = {};
        var addFonts = function (list) { (list || []).forEach(function (f) { fonts[f] = true; }); };
        var fallback = [];              // Excel 無法排序的 CMF 清單：這次改用內建方式畫

        // 0. 從 Excel 複製的表格：重新複製貼上（CMF 清單由 Excel 先排好再複製）
        var excelStep = Promise.resolve();
        if (excelTables.length) {
          if (!IS_WIN) {
            var skipped = excelTables.length - cmfExcel.length;
            if (skipped) lines.push(['warn', '這台電腦無法透過 Excel 複製，略過 ' + skipped + ' 個表格']);
            fallback = cmfExcel.slice();
          } else {
            var items = excelTables.map(function (t) {
              var p = Table.parseRange(t.label);
              var item = { sheet: p.sheet, range: XLSX.utils.encode_range(p.s, p.e) };
              if (t.cmf) {
                var cfg = Table.cmfConfig(t.cmf, rangeWidth(p));
                item.cmf = { num: cfg.num + 1, name: cfg.name + 1 };
              }
              return item;
            });
            var replaced = 0, failedLabels = [], planError = {};
            excelStep = excelCopy(file, items, function (i, msg) {
              var t = excelTables[i];
              [].concat(msg.warn || []).forEach(function (w) {
                lines.push(['warn', t.label + '：部分格式沒有複製（' + psMessage(w) + '）']);
              });
              return host('es_pasteTable', [t.label, 'replace']).then(function (res) {
                if (res.error) throw new Error(res.error);
                replaced += res.count;
                failedLabels = failedLabels.concat(res.failed);
                if (res.count) done[t.label] = true;
              });
            }, function (i, msg) {
              var plan;
              try { plan = planFor(excelTables[i], [].concat(msg.rows || [])); }
              catch (e) { planError[i] = true; throw e; }
              return plan.changed ? { head: plan.head, order: plan.order, serial: plan.serial, edge: plan.edge } : null;
            }).then(function (res) {
              var fatalShown = false;
              excelTables.forEach(function (t, i) {
                if (done[t.label]) return;
                var why = res.errors[i] || res.fatal;
                if (!why) return;
                if (t.cmf && !planError[i]) {
                  fallback.push(t);
                  lines.push(['warn', '無法透過 Excel 排序（' + psMessage(why) + '），這次改用內建方式畫表格']);
                } else if (res.errors[i]) {
                  lines.push(['error', t.label + '：' + psMessage(res.errors[i])]);
                } else if (!fatalShown) {
                  fatalShown = true;
                  lines.push(['error', '無法透過 Excel 更新表格：' + res.fatal]);
                }
              });
              if (replaced && !only) lines.push(['ok', '從 Excel 更新了 ' + replaced + ' 個表格']);
              if (failedLabels.length) lines.push(['error', '表格無法修改（可能被鎖定或隱藏）：' + failedLabels.join('、')]);
            });
          }
        }

        // 1. 內建方式畫的表格：依 Excel 重建（內容、字型、顏色、框線）
        var tableData = {}, tableErrors = [];
        var addBuilt = function (t, keepExcel) {
          try {
            var sel = Table.parseRange(t.label);
            if (!sel) throw new Error('範圍格式不對');
            var opts = {};
            if (t.cmf) {
              var cfg = Table.cmfConfig(t.cmf, rangeWidth(sel));
              opts.plan = planFor(t, Table.cmfRows(wb, sel, cfg));
              opts.numCol = cfg.num;
            }
            var data = Table.buildTable(wb, sel, opts);
            if (keepExcel) data.keepExcel = true;    // 下次更新仍然先試著透過 Excel 複製
            tableData[t.label] = data;
          } catch (e) { tableErrors.push(t.label + '：' + e.message); }
        };
        var step = excelStep.then(function () {
          builtTables.forEach(function (t) { addBuilt(t, false); });
          fallback.forEach(function (t) { addBuilt(t, true); });
          return Object.keys(tableData).length ? host('es_rebuildTables', [JSON.stringify(tableData)]) : null;
        });

        return step.then(function (tr) {
          if (tr && tr.error) throw new Error(tr.error);
          if (tr) {
            Object.keys(tableData).forEach(function (l) { if (tr.failed.indexOf(l) < 0) done[l] = true; });
            if (!only && tr.count) lines.push(['ok', '重建了 ' + tr.count + ' 個表格']);
            if (tr.failed.length) lines.push(['error', '表格無法修改（可能被鎖定或隱藏）：' + tr.failed.join('、')]);
            addFonts(tr.missingFonts);
          }
          tableErrors.forEach(function (m) { lines.push(['error', m]); });
          // CMF 清單：表格真的換成新的才回報排序結果
          tables.forEach(function (t) { if (t.cmf && done[t.label] && plans[t.label]) cmfReport(plans[t.label], lines); });

          // 2. 個別綁定的文字：只換內容
          if (!refs.length) return null;
          var report = collectValues(wb, refs);
          return host('es_apply', [JSON.stringify(report.values)]).then(function (res) {
            if (res.error) throw new Error(res.error);
            lines.push(['ok', res.changed ? '更新了 ' + res.changed + ' 個文字' + (res.same ? '，' + res.same + ' 個沒有變動' : '')
                                          : '文字都是最新的（' + res.same + ' 個）']);
            if (report.missing.length) lines.push(['warn', '找不到工作表：' + report.missing.join('、')]);
            if (report.empty.length) lines.push(['warn', '儲存格是空的：' + report.empty.join('、')]);
            if (res.failed.length) lines.push(['error', '無法修改（物件或圖層可能被鎖定、隱藏）：' + res.failed.join('、')]);
          });
        }).then(function () {
          var missingFonts = Object.keys(fonts);
          if (missingFonts.length) lines.push(['warn', 'Illustrator 找不到字型，改用預設字型：' + missingFonts.join('、')]);
          // 只有一行結果時看狀態列就好；有警告、錯誤或好幾項才展開說明
          var notes = lines.filter(function (l) { return l[0] !== 'time'; });
          if (notes.length > 1 || notes.some(function (l) { return l[0] !== 'ok'; })) showResult(lines);
          else ui.result.hidden = true;
          var bad = lines.filter(function (l) { return l[0] === 'error'; })[0];
          var main = lines.filter(function (l) { return l[0] === 'ok'; })[0];
          App.setStatus(bad ? bad[1] : main ? main[1] : lines.length > 1 ? lines[1][1] : '表格已更新', !!bad);
          return refreshContext();
        });
      });
    }).catch(function (err) {
      showError(friendly(err));
    }).then(function () {
      state.busy = false;
      ui.btnUpdate.disabled = false;
      if (cmf.pending) {
        cmf.pending = false;
        return update(true, true);
      }
    });
  }

  /* ---------- 透過 Excel 複製（Windows）：外觀跟手動複製貼上完全一樣 ---------- */

  // items: [{ sheet, range, fallbackSheet, fallbackRange, cmf }]
  // onCopied(i, {sheet, range}) 回傳 Promise，完成貼上後才讓 Excel 複製下一個
  // onRows(i, {rows}) 回傳排序方式（CMF 清單，見 scripts/excel-copy.ps1），null = 不用重排
  function excelCopy(file, items, onCopied, onRows) {
    return new Promise(function (resolve) {
      var jobFile = path.join(os.tmpdir(), 'excelsync-job-' + Date.now() + '.json');
      fs.writeFileSync(jobFile, JSON.stringify({ path: path.resolve(file), items: items }), 'utf8');

      var result = { ready: false, live: false, errors: {}, fatal: null };
      var finished = false, buffer = '', timer = null;
      var child = childProcess.spawn('powershell.exe',
        ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
         '-File', path.join(EXT, 'scripts', 'excel-copy.ps1'), jobFile],
        { windowsHide: true });

      function finish() {
        if (finished) return;
        finished = true;
        clearTimeout(timer);
        try { fs.unlinkSync(jobFile); } catch (e) {}
        if (!result.ready && !result.fatal) result.fatal = '無法啟動 Excel';
        resolve(result);
      }
      function watchdog() {
        clearTimeout(timer);
        timer = setTimeout(function () {
          if (!result.ready) result.fatal = 'Excel 太久沒有回應';
          try { child.kill(); } catch (e) {}
          finish();
        }, 60000);
      }
      function reply(text) { try { child.stdin.write(text + '\n'); } catch (e) {} }

      function handle(msg) {
        watchdog();
        if (msg.ev === 'ready') { result.ready = true; result.live = !!msg.live; }
        else if (msg.ev === 'rows') {
          Promise.resolve().then(function () { return onRows ? onRows(msg.i, msg) : null; })
            .then(function (plan) { reply(plan ? JSON.stringify(plan) : 'SKIP'); })
            .catch(function (err) { result.errors[msg.i] = err.message || String(err); reply('PASS'); });
        }
        else if (msg.ev === 'copied') {
          Promise.resolve().then(function () { return onCopied(msg.i, msg); })
            .then(function () { reply('NEXT'); })
            .catch(function (err) { result.errors[msg.i] = err.message || String(err); reply('STOP'); });
        }
        else if (msg.ev === 'error') result.errors[msg.i] = msg.message;
        else if (msg.ev === 'fatal') result.fatal = msg.message;
        else if (msg.ev === 'end') { try { child.stdin.end(); } catch (e) {} }
      }

      child.stdout.setEncoding('utf8');
      child.stdout.on('data', function (chunk) {
        buffer += chunk.replace(/^﻿/, '');
        var lines = buffer.split(/\r?\n/);
        buffer = lines.pop();
        lines.forEach(function (line) {
          line = line.trim();
          if (!line) return;
          try { handle(JSON.parse(line)); } catch (e) {}
        });
      });
      child.on('error', function (err) { result.fatal = err.message; finish(); });
      child.on('exit', finish);
      watchdog();
    });
  }

  // excel-copy.ps1 的錯誤是「步驟@行號: 訊息」
  var PS_STEPS = { workbook: '建立暫存活頁簿', widths: '欄寬', rows: '複製列', heights: '列高', formulas: '公式',
                   serial: '序號', borders: '框線', range: '範圍', copy: '複製' };
  function psMessage(m) {
    return String(m).replace(/^(\w+)@(\d+): /, function (all, step, line) {
      return (PS_STEPS[step] || step) + '，第 ' + line + ' 行：';
    });
  }

  function excelLabel(msg) {
    return Table.quoteSheet(msg.sheet) + '!' + String(msg.range).replace(/\$/g, '');
  }

  function importTable() {
    if (state.busy) return;
    if (!state.excelPath) { showError(new Error(state.candidates.length > 1 ? '請先從清單選一個 Excel 檔' : '請先選擇 Excel 檔案')); return; }

    var typed = ui.rangeRef.value.trim();
    var sel = null;
    if (typed) {
      sel = Table.parseRange(typed);
      if (!sel) { showError(new Error('範圍格式不對，例如 A1:D10 或 工作表1!A1:D10')); ui.rangeRef.focus(); return; }
    }

    state.busy = true;
    ui.btnImport.disabled = true;
    App.setStatus(IS_WIN ? '正在透過 Excel 複製…' : '正在讀取 Excel…');

    readWorkbook(state.excelPath, true).then(function (wb) {
      var saved = Table.savedSelection(wb);
      if (!IS_WIN) return importBuilt(wb, sel || Table.parseRange(Table.quoteSheet(saved.sheet) + '!' + saved.range), []);

      var item = sel
        ? { sheet: sel.sheet, range: XLSX.utils.encode_range(sel.s, sel.e) }
        : { sheet: null, range: '', fallbackSheet: saved.sheet, fallbackRange: saved.range };
      var label = null;
      return excelCopy(state.excelPath, [item], function (i, msg) {
        label = excelLabel(msg);
        return host('es_pasteTable', [label, 'new']).then(function (r) { if (r.error) throw new Error(r.error); });
      }).then(function (res) {
        if (label && !res.errors[0]) {
          ui.rangeRef.value = label;
          ui.result.hidden = true;
          App.setStatus('已匯入 ' + label + (res.live ? '（Excel 目前開啟的內容）' : ''));
          return refreshContext();
        }
        // Excel 無法使用時，改用外掛自己畫
        var why = res.fatal || res.errors[0] || '未知的錯誤';
        return importBuilt(wb, sel || Table.parseRange(Table.quoteSheet(saved.sheet) + '!' + saved.range),
          [['warn', '無法透過 Excel 複製（' + why + '），改用內建方式匯入，外觀可能略有不同']]);
      });
    }).catch(function (err) {
      showError(friendly(err));
    }).then(function () {
      state.busy = false;
      ui.btnImport.disabled = false;
    });
  }

  // 內建方式：依 Excel 的樣式自己畫出表格（Mac，或 Excel 無法使用時）
  function importBuilt(wb, sel, notes) {
    if (!sel) throw new Error('讀不到 Excel 的選取範圍，請直接輸入範圍');
    var table = Table.buildTable(wb, sel, { grid: ui.chkGrid.checked });
    if (!table.cells.length) throw new Error('這個範圍裡沒有可見的儲存格');
    ui.rangeRef.value = table.label;
    return host('es_importTable', [JSON.stringify(table)]).then(function (r) {
      if (r.error) throw new Error(r.error);
      var lines = notes.slice();
      if (r.missingFonts.length) lines.push(['warn', 'Illustrator 找不到字型，改用預設字型：' + r.missingFonts.join('、')]);
      if (lines.length) showResult(lines); else ui.result.hidden = true;
      App.setStatus('已匯入 ' + table.label + '，共 ' + r.count + ' 格');
      return refreshContext();
    });
  }

  /* ---------- 自動更新：監看檔案修改時間 ---------- */

  function onFileChanged(curr, prev) {
    if (curr.mtimeMs === prev.mtimeMs || curr.mtimeMs === 0) return;
    clearTimeout(state.debounce);
    state.debounce = setTimeout(function () { update(true); }, 800);
  }

  function syncWatcher() {
    var want = ui.chkAuto.checked && state.excelPath;
    if (state.watching && state.watching !== want) {
      fs.unwatchFile(state.watching, onFileChanged);
      state.watching = null;
    }
    if (want && !state.watching) {
      // watchFile 用輪詢；Excel 存檔是「寫暫存檔再改名」，fs.watch 容易漏掉
      fs.watchFile(want, { interval: 1000 }, onFileChanged);
      state.watching = want;
    }
    try { localStorage.setItem('excelsync:auto', ui.chkAuto.checked ? '1' : '0'); } catch (e) {}
  }

  /* ---------- 給標註分頁用 ---------- */

  App.excel = {
    // 編號表的物件下拉選單：沒有 CMF 清單時回傳 null
    cmfInfo: function () { return refreshContext().then(function (all) { return cmfList(all.tables); }); },
    numbersChanged: numbersChanged
  };

  /* ---------- 啟動 ---------- */

  ui.btnPick.addEventListener('click', pickFile);
  ui.btnAuto.addEventListener('click', useAuto);
  ui.fileChoices.addEventListener('change', function () {
    if (ui.fileChoices.value) { setExcelPath(ui.fileChoices.value, 'manual'); afterFileChanged(); }
  });
  ui.btnBind.addEventListener('click', bind);
  ui.btnImport.addEventListener('click', importTable);
  ui.rangeRef.addEventListener('keydown', function (e) { if (e.key === 'Enter') importTable(); });
  ui.btnUnbind.addEventListener('click', unbind);
  ui.btnUpdate.addEventListener('click', function () { update(false); });
  ui.chkAuto.addEventListener('change', syncWatcher);
  ui.cellRef.addEventListener('keydown', function (e) { if (e.key === 'Enter') bind(); });
  ui.cmfTable.addEventListener('change', onCmfTable);
  [ui.cmfNum, ui.cmfName, ui.cmfHead, ui.cmfRest].forEach(function (el) { el.addEventListener('change', onCmfField); });
  ui.cmfAuto.addEventListener('change', function () {
    try { localStorage.setItem('cmftool:cmfAuto', ui.cmfAuto.checked ? '1' : '0'); } catch (e) {}
  });
  ui.btnCmfMatch.addEventListener('click', autoMatch);

  // CEP 沒有「選取改變」事件，所以滑鼠移進面板時更新一次
  var refresh = function () { refreshContext().catch(function () {}); };
  document.documentElement.addEventListener('mouseenter', refresh);
  window.addEventListener('focus', refresh);
  cs.addEventListener('documentAfterActivate', refresh);
  cs.addEventListener('documentAfterDeactivate', refresh);

  try { ui.chkAuto.checked = localStorage.getItem('excelsync:auto') === '1'; } catch (e) {}
  try { ui.cmfAuto.checked = localStorage.getItem('cmftool:cmfAuto') !== '0'; } catch (e) {}
  if (IS_WIN) ui.chkGrid.parentNode.hidden = true;   // 透過 Excel 複製時，外觀完全照 Excel
  refresh();
})();
