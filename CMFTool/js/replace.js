/* CMF Tool — 取代分頁：一次換好幾組文字（含工作區名稱），例如版本號和日期 */
(function () {
  'use strict';

  var App = window.CMFApp;
  var cep = window.__adobe_cep__;
  var $ = function (id) { return document.getElementById(id); };
  var ROWS_KEY = 'cmftool:replaceRows';
  var OPTS_KEY = 'cmftool:replaceOpts';
  var OPTS = ['repText', 'repBoards', 'repCase', 'repSkipCmf', 'repLocked', 'repShift'];

  var rows = [];       // [{ find, rep }]
  var counts = null;   // 上次計算的每一組符合幾處；改了欄位或選項就清掉
  var busy = false;

  function call(fn, obj, cb) {
    if (!cep) { App.setStatus('請在 Illustrator 中開啟此面板', true); return; }
    var arg = obj ? JSON.stringify(JSON.stringify(obj)).replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029') : '';
    cep.evalScript('CMF.' + fn + '(' + arg + ')', function (raw) {
      var r;
      try { r = JSON.parse(raw); } catch (e) { r = { ok: false, msg: '腳本錯誤：' + raw }; }
      cb(r);
    });
  }

  function save() {
    try { localStorage.setItem(ROWS_KEY, JSON.stringify(rows)); } catch (e) {}
  }

  function scope() {
    var el = document.querySelector('input[name=repScope]:checked');
    return el ? el.value : 'all';
  }

  // 會執行的組：尋找和取代為都有填（取代為留空的不執行，免得忘了填就把字刪掉）
  function isOn(r) { return !!(r.find && r.rep && r.find !== r.rep); }

  function active() {
    return rows.map(function (r, i) { return { find: r.find, rep: r.rep, i: i }; }).filter(isOn);
  }

  /* ---------- 自動填入：日期 → 今天（格式一樣）；版本號 → 最後一個數字加一 ---------- */
  function pad(n, w) { n = String(n); while (n.length < w) n = '0' + n; return n; }

  function validDate(y, m, d) { return m >= 1 && m <= 12 && d >= 1 && d <= 31 && y >= 1900 && y <= 2100; }

  function nextValue(text) {
    var now = new Date(), Y = now.getFullYear(), M = now.getMonth() + 1, D = now.getDate(), m;
    // 2026.10.01、2026/10/1、2026-10-01
    m = /(\d{4})([.\/\-])(\d{1,2})\2(\d{1,2})(?!\d)/.exec(text);
    if (m && validDate(+m[1], +m[3], +m[4])) {
      return text.replace(m[0], Y + m[2] + pad(M, m[3].length) + m[2] + pad(D, m[4].length));
    }
    // 2026年10月1日
    m = /(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日/.exec(text);
    if (m && validDate(+m[1], +m[2], +m[3])) {
      return text.replace(m[0], m[0].replace(/(\d{4})(\s*年\s*)(\d{1,2})(\s*月\s*)(\d{1,2})/, function (all, y, s1, mo, s2, da) {
        return Y + s1 + pad(M, mo.length) + s2 + pad(D, da.length);
      }));
    }
    // 10/01/2026（月/日/年）；第一個數字大於 12 時當成 日/月/年
    m = /(?:^|\D)((\d{1,2})([.\/\-])(\d{1,2})\3(\d{4}))(?!\d)/.exec(text);
    if (m) {
      var dm = +m[2] > 12, mo = dm ? +m[4] : +m[2], da = dm ? +m[2] : +m[4];
      if (validDate(+m[5], mo, da)) {
        var a = dm ? D : M, b = dm ? M : D;
        return text.replace(m[1], pad(a, m[2].length) + m[3] + pad(b, m[4].length) + m[3] + Y);
      }
    }
    // 20261001
    m = /(?:^|\D)((\d{4})(\d{2})(\d{2}))(?!\d)/.exec(text);
    if (m && validDate(+m[2], +m[3], +m[4])) return text.replace(m[1], Y + pad(M, 2) + pad(D, 2));
    // 兩位數年份：10/15/26（月/日/年）
    m = /(?:^|[^\d.\/\-])((\d{1,2})\/(\d{1,2})\/(\d{2}))(?![\d\/])/.exec(text);
    if (m) {
      var sy = +m[2] > 12, ym = sy ? +m[3] : +m[2], yd = sy ? +m[2] : +m[3];
      if (validDate(Y, ym, yd)) return text.replace(m[1], pad(sy ? D : M, m[2].length) + '/' + pad(sy ? M : D, m[3].length) + '/' + pad(Y % 100, 2));
    }
    // 沒有年份：10/15（月/日；第一個數字大於 12 時當成 日/月）、10月15日。只認斜線，1.2、1-2 是版本號、料號
    m = /(?:^|[^\d.\/\-])((\d{1,2})\/(\d{1,2}))(?![\d\/])/.exec(text);
    if (m) {
      var sw = +m[2] > 12, mm = sw ? +m[3] : +m[2], dd = sw ? +m[2] : +m[3];
      if (validDate(Y, mm, dd)) return text.replace(m[1], pad(sw ? D : M, m[2].length) + '/' + pad(sw ? M : D, m[3].length));
    }
    m = /(\d{1,2})(\s*月\s*)(\d{1,2})(\s*日)/.exec(text);
    if (m && validDate(Y, +m[1], +m[3])) return text.replace(m[0], pad(M, m[1].length) + m[2] + pad(D, m[3].length) + m[4]);
    // 版本號：最後一段數字加一（V1.2 → V1.3、R09 → R10）
    m = /(\d+)(?!.*\d)/.exec(text);
    if (m) return text.slice(0, m.index) + pad(Number(m[1]) + 1, m[1].length) + text.slice(m.index + m[1].length);
    // 版次是字母：Rev.A → Rev.B
    m = /([A-Ya-y])$/.exec(text);
    if (m && /(^|[^A-Za-z])[A-Za-z]$/.test(text)) return text.slice(0, -1) + String.fromCharCode(m[1].charCodeAt(0) + 1);
    return null;
  }

  /* ---------- 畫面 ---------- */
  function rowEl(r, i) {
    var el = document.createElement('div');
    el.className = 'rep-row';
    el.innerHTML =
      '<div class="rep-cell"><input type="text" class="rep-find" spellcheck="false" autocomplete="off" placeholder="尋找"' +
      ' title="要找的文字。工作區名稱裡的也會找到（有勾「工作區名稱」時）"><span class="rep-n"></span></div>' +
      '<span class="rep-arrow" aria-hidden="true">→</span>' +
      '<input type="text" class="rep-to" spellcheck="false" autocomplete="off" placeholder="取代為"' +
      ' title="換成的文字。留空的這一組不會執行">' +
      '<button class="icon-btn rep-auto" title="✦ 自動填入：尋找的是日期 → 換成今天（格式一樣）；是版本號 → 最後一個數字加一（V1.2 → V1.3）">' +
      '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
      '<path d="M8 2.5 9.3 6.7 13.5 8 9.3 9.3 8 13.5 6.7 9.3 2.5 8 6.7 6.7Z"/></svg></button>' +
      '<button class="icon-btn rep-del" title="刪除這一組">' +
      '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" aria-hidden="true"><path d="M4.5 4.5l7 7M11.5 4.5l-7 7"/></svg></button>';
    var find = el.querySelector('.rep-find'), to = el.querySelector('.rep-to');
    find.value = r.find;
    to.value = r.rep;
    find.setAttribute('aria-label', '尋找（第 ' + (i + 1) + ' 組）');
    to.setAttribute('aria-label', '取代為（第 ' + (i + 1) + ' 組）');

    function edited() {
      r.find = find.value;
      r.rep = to.value;
      counts = null;
      save();
      render();
    }
    find.addEventListener('input', edited);
    to.addEventListener('input', edited);
    [find, to].forEach(function (inp) {
      inp.addEventListener('keydown', function (e) {
        if (e.key === 'Enter') { e.preventDefault(); if (e.ctrlKey || e.metaKey) doReplace(); else count(); }
      });
    });
    el.querySelector('.rep-auto').addEventListener('click', function () {
      var v = find.value ? nextValue(find.value) : null;
      if (v === null) { App.setStatus(find.value ? '看不出「' + find.value + '」是日期或版本號' : '請先在「尋找」填入目前的日期或版本號', true); return; }
      to.value = v;
      edited();
    });
    el.querySelector('.rep-del').addEventListener('click', function () {
      rows.splice(i, 1);
      if (!rows.length) rows.push({ find: '', rep: '' });
      counts = null;
      save();
      build();
    });
    return el;
  }

  function build() {
    var box = $('repRows');
    box.innerHTML = '';
    rows.forEach(function (r, i) { box.appendChild(rowEl(r, i)); });
    render();
  }

  function render() {
    var els = document.querySelectorAll('#repRows .rep-row'), act = active();
    Array.prototype.forEach.call(els, function (el, i) {
      var r = rows[i], n = el.querySelector('.rep-n'), ok = isOn(r);
      var c = counts && ok ? counts[i] : null;
      n.textContent = c === null || c === undefined ? '' : String(c);
      n.classList.toggle('zero', c === 0);
      n.title = c === null || c === undefined ? '' : '符合 ' + c + ' 處';
      el.classList.toggle('is-off', !!r.find && !ok);
      el.querySelector('.rep-del').disabled = rows.length === 1 && !r.find && !r.rep;
    });
    var what = $('repText').checked || $('repBoards').checked;
    var total = counts ? act.reduce(function (s, r) { return s + (counts[r.i] || 0); }, 0) : null;
    $('btnReplace').disabled = busy || !act.length || !what || total === 0;
    $('btnReplace').textContent = busy ? '取代中…' : total ? '取代 ' + total + ' 處' : '取代';
    $('repMeta').textContent = !act.length ? '' : act.length + ' 組';
  }

  function opts(dry) {
    return {
      // 不執行的組送空的，回傳的 counts 才跟 rows 一一對應
      pairs: rows.map(function (r) { return isOn(r) ? { find: r.find, rep: r.rep } : { find: '', rep: '' }; }),
      scope: scope(),
      text: $('repText').checked,
      boards: $('repBoards').checked,
      cs: $('repCase').checked,
      skipCmf: $('repSkipCmf').checked,
      locked: $('repLocked').checked,
      dry: !!dry
    };
  }

  // 只計算有幾處符合，不改文件
  var loading = false, again = false;
  function count(quiet) {
    if (!cep || $('page-replace').hidden || busy) return;
    if (!active().length || !($('repText').checked || $('repBoards').checked)) { counts = null; render(); return; }
    if (loading) { again = true; return; }
    loading = true;
    call('replaceText', opts(true), function (r) {
      loading = false;
      counts = r.ok ? r.counts : null;
      render();
      if (!quiet || !r.ok) App.setStatus(r.msg, !r.ok);
      if (again) { again = false; count(quiet); }
    });
  }

  function doReplace() {
    if (busy || !active().length) return;
    busy = true;
    render();
    App.setStatus('取代中…');
    call('replaceText', opts(false), function (r) {
      busy = false;
      App.setStatus(r.msg, !r.ok || r.failed > 0);
      // 取代後把「取代為」移到「尋找」，下次改版直接按自動填入
      if (r.ok && r.total > 0 && $('repShift').checked) {
        active().forEach(function (a) {
          if (r.counts[a.i] > 0) { rows[a.i].find = rows[a.i].rep; rows[a.i].rep = ''; }
        });
        save();
        build();
      }
      counts = null;
      render();
      count(true);
    });
  }

  /* ---------- 事件 ---------- */
  $('btnRepAdd').addEventListener('click', function () {
    rows.push({ find: '', rep: '' });
    save();
    build();
    var finds = document.querySelectorAll('#repRows .rep-find');
    finds[finds.length - 1].focus();
  });
  $('btnReplace').addEventListener('click', doReplace);
  $('btnRepCount').addEventListener('click', function () { count(false); });

  function saveOpts() {
    var o = { scope: scope() };
    OPTS.forEach(function (id) { o[id] = $(id).checked; });
    try { localStorage.setItem(OPTS_KEY, JSON.stringify(o)); } catch (e) {}
  }
  Array.prototype.forEach.call(document.querySelectorAll('#page-replace input[type=radio], #page-replace .rep-opts input'), function (el) {
    el.addEventListener('change', function () {
      saveOpts();
      if (el.id === 'repShift') return;
      counts = null;
      render();
      count(true);
    });
  });

  try {
    var o = JSON.parse(localStorage.getItem(OPTS_KEY) || 'null');
    if (o) {
      OPTS.forEach(function (id) { if (typeof o[id] === 'boolean') $(id).checked = o[id]; });
      var radio = o.scope && document.querySelector('input[name=repScope][value="' + o.scope + '"]');
      if (radio) radio.checked = true;
    }
  } catch (e) {}
  try {
    var saved = JSON.parse(localStorage.getItem(ROWS_KEY) || 'null');
    if (saved && saved.length) rows = saved.map(function (r) { return { find: String(r.find || ''), rep: String(r.rep || '') }; });
  } catch (e) {}
  if (!rows.length) rows.push({ find: '', rep: '' });

  // 換了文件、在 Illustrator 裡改了東西之後，滑鼠移回面板時重新計算（打字中不呼叫，見 App.whenIdle）
  App.on('tab', function (name) { if (name === 'replace') count(true); });
  var autoCount = App.whenIdle(function () { count(true); });
  document.documentElement.addEventListener('mouseenter', autoCount);
  window.addEventListener('focus', autoCount);
  if (cep) {
    try { cep.addEventListener('documentAfterActivate', function () { count(true); }); } catch (e) {}
  }

  build();
  count(true);
})();
