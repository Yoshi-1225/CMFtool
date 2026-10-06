/* Excel 同步 — 面板端（CEP，含 Node.js） */
(function () {
  'use strict';

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
    rangeRef: $('rangeRef'), btnImport: $('btnImport'), chkGrid: $('chkGrid')
  };

  var state = {
    docKey: null, docFolder: '', excelPath: null,
    source: null,            // 'manual' = 使用者選的；'auto' = 從 .ai 同資料夾找到的
    candidates: [], candidatesKey: null,
    selectedRefs: [], busy: false, watching: null, debounce: null
  };

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
    // withLayout：連欄寬、列高和原始 XML 一起讀（匯入表格時需要）
    var opts = withLayout ? { type: 'buffer', cellStyles: true, bookFiles: true } : { type: 'buffer' };
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
    } else {
      ui.fileName.textContent = path.basename(f);
      ui.fileName.classList.toggle('is-empty', false);
      ui.fileName.title = f;
      ui.fileDir.textContent = path.dirname(f);
    }

    var msg = '';
    if (!state.docKey) msg = '';
    else if (state.source === 'manual') msg = '手動選擇';
    else if (!state.docFolder) msg = '.ai 存檔後，會自動找同資料夾的 Excel';
    else if (f) msg = '自動選取：與 .ai 在同一個資料夾';
    else if (state.candidates.length > 1) msg = '資料夾裡有 ' + state.candidates.length + ' 個 Excel 檔，請選一個';
    else msg = '同資料夾沒有 Excel 檔，請按「選擇…」';
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
    showResult([['error', err.message || String(err)]]);
  }

  function chip(text, cls, onClick) {
    var b = document.createElement('button');
    b.type = 'button';
    b.className = cls;
    b.textContent = text;
    b.addEventListener('click', onClick);
    ui.refList.appendChild(b);
  }

  function renderRefs(refs, tables) {
    tables = tables || [];
    ui.refList.innerHTML = '';
    var total = refs.length + tables.length;
    ui.refCount.textContent = total ? '(' + total + ')' : '';
    if (!total) {
      ui.refList.innerHTML = '<p class="hint">還沒有綁定任何文字</p>';
      return;
    }
    tables.forEach(function (t) {
      var label = t.label;
      chip('表格 ' + label, 'ref-chip is-table', function () {
        host('es_selectTable', [label]).then(refreshContext).catch(showError);
      });
    });
    var LIMIT = 40, sorted = refs.slice().sort(compareRefs);
    sorted.slice(0, LIMIT).forEach(function (ref) {
      chip(ref, 'ref-chip' + (state.selectedRefs.indexOf(ref) >= 0 ? ' is-selected' : ''), function () {
        host('es_selectRef', [ref]).then(refreshContext).catch(showError);
      });
    });
    if (sorted.length > LIMIT) {
      var more = document.createElement('span');
      more.className = 'hint';
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

  function refreshContext() {
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
        ui.selInfo.textContent = '請先開啟 Illustrator 文件';
        renderRefs([]);
        return;
      }
      resolveExcel(ctx);
      state.selectedRefs = ctx.refs;
      if (ctx.selected === 0) ui.selInfo.textContent = '請在 Illustrator 中選取文字物件';
      else if (ctx.refs.length === 0) ui.selInfo.textContent = '已選取 ' + ctx.selected + ' 個文字物件，尚未綁定';
      else ui.selInfo.textContent = '已選取 ' + ctx.selected + ' 個文字物件，目前綁定 ' + ctx.refs.join('、');
      if (ctx.refs.length === 1 && document.activeElement !== ui.cellRef) ui.cellRef.value = ctx.refs[0];
      renderRefs(all.refs, all.tables);
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
    refreshContext().catch(showError);
  }

  function pickFile() {
    var startDir = state.excelPath ? path.dirname(state.excelPath) : (state.docFolder || '');
    var dlg = window.cep.fs.showOpenDialogEx || window.cep.fs.showOpenDialog;
    var res = dlg(false, false, '選擇 Excel 檔案', startDir, ['xlsx', 'xlsm', 'xls']);
    if (res.err || !res.data || !res.data.length) return;
    if (!state.docKey) { showError(new Error('請先開啟 Illustrator 文件')); return; }
    setExcelPath(res.data[0], 'manual');
    ui.result.hidden = true;
  }

  function bind() {
    var ref = normalizeRef(ui.cellRef.value);
    if (!ref) { showError(new Error('儲存格位址格式不對，例如 B3 或 工作表1!B3')); ui.cellRef.focus(); return; }
    ui.cellRef.value = ref;
    host('es_bind', [ref]).then(function (r) {
      if (r.error) throw new Error(r.error);
      showResult([['ok', '已將 ' + r.count + ' 個文字物件綁定到 ' + ref]]);
      return refreshContext();
    }).catch(showError);
  }

  function unbind() {
    host('es_unbind').then(function (r) {
      if (r.error) throw new Error(r.error);
      showResult([['ok', r.count ? '已解除 ' + r.count + ' 個文字物件的綁定' : '選取的物件沒有綁定']]);
      return refreshContext();
    }).catch(showError);
  }

  function update(auto) {
    if (state.busy) return Promise.resolve();
    if (!state.excelPath) {
      if (!auto) showError(new Error(state.candidates.length > 1 ? '請先從清單選一個 Excel 檔' : '請先選擇 Excel 檔案'));
      return Promise.resolve();
    }
    if (!fs.existsSync(state.excelPath)) { showError(new Error('找不到檔案：' + state.excelPath)); return Promise.resolve(); }

    state.busy = true;
    ui.btnUpdate.disabled = true;
    var file = state.excelPath, report;

    return host('es_refs').then(function (r) {
      if (!r.doc) throw new Error('請先開啟 Illustrator 文件');
      // 自動更新時，若使用者切到別的文件，就不要動它
      if (auto && r.doc !== state.docKey) return null;
      if (!r.refs.length && !r.tables.length) throw new Error('文件中還沒有綁定任何文字或表格');

      var excelTables = r.tables.filter(function (t) { return t.mode === 'excel'; }).map(function (t) { return t.label; });
      var builtTables = r.tables.filter(function (t) { return t.mode !== 'excel'; }).map(function (t) { return t.label; });

      return readWorkbook(file, builtTables.length > 0).then(function (wb) {
        var lines = [['time', new Date().toLocaleTimeString() + (auto ? '（自動）' : '')]];
        var fonts = {};
        var addFonts = function (list) { (list || []).forEach(function (f) { fonts[f] = true; }); };

        // 0. 從 Excel 複製的表格：重新複製貼上
        var excelStep = Promise.resolve();
        if (excelTables.length) {
          if (!IS_WIN) {
            lines.push(['warn', '這台電腦無法透過 Excel 複製，略過 ' + excelTables.length + ' 個表格']);
          } else {
            var items = excelTables.map(function (label) {
              var p = Table.parseRange(label);
              return { sheet: p.sheet, range: XLSX.utils.encode_range(p.s, p.e) };
            });
            var replaced = 0, failedLabels = [];
            excelStep = excelCopy(file, items, function (i) {
              return host('es_pasteTable', [excelTables[i], 'replace']).then(function (res) {
                if (res.error) throw new Error(res.error);
                replaced += res.count;
                failedLabels = failedLabels.concat(res.failed);
              });
            }).then(function (res) {
              if (res.fatal) lines.push(['error', '無法透過 Excel 更新表格：' + res.fatal]);
              if (replaced) lines.push(['ok', '從 Excel 更新了 ' + replaced + ' 個表格']);
              Object.keys(res.errors).forEach(function (i) { lines.push(['error', excelTables[i] + '：' + res.errors[i]]); });
              if (failedLabels.length) lines.push(['error', '表格無法修改（可能被鎖定或隱藏）：' + failedLabels.join('、')]);
            });
          }
        }

        // 1. 內建方式畫的表格：依 Excel 重建（內容、字型、顏色、框線）
        var tableData = {}, tableErrors = [], hasTables = false;
        builtTables.forEach(function (label) {
          try {
            var sel = Table.parseRange(label);
            if (!sel) throw new Error('範圍格式不對');
            tableData[label] = Table.buildTable(wb, sel);
            hasTables = true;
          } catch (e) { tableErrors.push(label + '：' + e.message); }
        });
        var step = excelStep.then(function () {
          return hasTables ? host('es_rebuildTables', [JSON.stringify(tableData)]) : null;
        });

        return step.then(function (tr) {
          if (tr && tr.error) throw new Error(tr.error);
          if (tr) {
            lines.push(['ok', '重建了 ' + tr.count + ' 個表格']);
            if (tr.failed.length) lines.push(['error', '表格無法修改（可能被鎖定或隱藏）：' + tr.failed.join('、')]);
            addFonts(tr.missingFonts);
          }
          tableErrors.forEach(function (m) { lines.push(['warn', m]); });

          // 2. 個別綁定的文字：只換內容
          if (!r.refs.length) return null;
          var report = collectValues(wb, r.refs);
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
          showResult(lines);
          return refreshContext();
        });
      });
    }).catch(function (err) {
      if (err && (err.code === 'EBUSY' || err.code === 'EPERM' || err.code === 'EACCES')) {
        err = new Error('Excel 檔案暫時無法讀取，請確認已存檔後再試一次');
      }
      showError(err);
    }).then(function () {
      state.busy = false;
      ui.btnUpdate.disabled = false;
    });
  }

  /* ---------- 透過 Excel 複製（Windows）：外觀跟手動複製貼上完全一樣 ---------- */

  // items: [{ sheet, range, fallbackSheet, fallbackRange }]
  // onCopied(i, {sheet, range}) 回傳 Promise，完成貼上後才讓 Excel 複製下一個
  function excelCopy(file, items, onCopied) {
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
        buffer += chunk.replace(/^\uFEFF/, '');
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
    showResult([['time', IS_WIN ? '正在透過 Excel 複製…' : '正在讀取 Excel…']]);

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
          showResult([
            ['ok', '已從 Excel 匯入 ' + label + (res.live ? '（使用 Excel 目前開啟的內容）' : '')],
            ['time', '之後按「從 Excel 更新」會重新複製，保留表格的位置和縮放']
          ]);
          return refreshContext();
        }
        // Excel 無法使用時，改用外掛自己畫
        var why = res.fatal || res.errors[0] || '未知的錯誤';
        return importBuilt(wb, sel || Table.parseRange(Table.quoteSheet(saved.sheet) + '!' + saved.range),
          [['warn', '無法透過 Excel 複製（' + why + '），改用內建方式匯入，外觀可能略有不同']]);
      });
    }).catch(function (err) {
      if (err && (err.code === 'EBUSY' || err.code === 'EPERM' || err.code === 'EACCES')) {
        err = new Error('Excel 檔案暫時無法讀取，請確認已存檔後再試一次');
      }
      showError(err);
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
      var lines = notes.concat([
        ['ok', '已匯入 ' + table.label + '，共 ' + r.count + ' 格'],
        ['time', 'Excel 改完存檔後，按「從 Excel 更新」會同步內容和樣式']
      ]);
      if (r.missingFonts.length) lines.push(['warn', 'Illustrator 找不到字型，改用預設字型：' + r.missingFonts.join('、')]);
      showResult(lines);
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

  /* ---------- 配色跟隨 Illustrator 介面亮度 ---------- */

  function applyTheme() {
    var c = cs.getHostEnvironment().appSkinInfo.panelBackgroundColor.color;
    var r = Math.round(c.red), g = Math.round(c.green), b = Math.round(c.blue);
    var dark = (0.299 * r + 0.587 * g + 0.114 * b) < 128;
    var shift = function (d) {
      var f = function (v) { return Math.max(0, Math.min(255, v + d)); };
      return 'rgb(' + f(r) + ',' + f(g) + ',' + f(b) + ')';
    };
    var s = document.documentElement.style;
    s.setProperty('--bg', 'rgb(' + r + ',' + g + ',' + b + ')');
    s.setProperty('--fg', dark ? '#e1e1e1' : '#1f1f1f');
    s.setProperty('--muted', dark ? '#9a9a9a' : '#6e6e6e');
    s.setProperty('--field', dark ? shift(-22) : '#ffffff');
    s.setProperty('--line', dark ? shift(24) : shift(-36));
    s.setProperty('--btn', dark ? shift(18) : shift(-14));
    s.setProperty('--btn-hover', dark ? shift(30) : shift(-26));
    s.setProperty('--accent', dark ? '#21a366' : '#107c41');
    s.setProperty('--warn', dark ? '#e0a43a' : '#9a6200');
    s.setProperty('--error', dark ? '#e5675f' : '#c42b1c');
  }

  /* ---------- 啟動 ---------- */

  ui.btnPick.addEventListener('click', pickFile);
  ui.btnAuto.addEventListener('click', useAuto);
  ui.fileChoices.addEventListener('change', function () {
    if (ui.fileChoices.value) setExcelPath(ui.fileChoices.value, 'manual');
  });
  ui.btnBind.addEventListener('click', bind);
  ui.btnImport.addEventListener('click', importTable);
  ui.rangeRef.addEventListener('keydown', function (e) { if (e.key === 'Enter') importTable(); });
  ui.btnUnbind.addEventListener('click', unbind);
  ui.btnUpdate.addEventListener('click', function () { update(false); });
  ui.chkAuto.addEventListener('change', syncWatcher);
  ui.cellRef.addEventListener('keydown', function (e) { if (e.key === 'Enter') bind(); });

  // CEP 沒有「選取改變」事件，所以滑鼠移進面板時更新一次
  var refresh = function () { refreshContext().catch(function () {}); };
  document.documentElement.addEventListener('mouseenter', refresh);
  window.addEventListener('focus', refresh);
  cs.addEventListener('documentAfterActivate', refresh);
  cs.addEventListener('documentAfterDeactivate', refresh);
  cs.addEventListener(CSInterface.THEME_COLOR_CHANGED_EVENT, applyTheme);

  try { ui.chkAuto.checked = localStorage.getItem('excelsync:auto') === '1'; } catch (e) {}
  if (IS_WIN) ui.chkGrid.parentNode.hidden = true;   // 透過 Excel 複製時，外觀完全照 Excel
  applyTheme();
  refresh();
})();
