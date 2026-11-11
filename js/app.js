/* 불량 제품 파악 — 화면 (자료 넣기 · ① 출하 대조 · ② 판독 · ③ 검사 결과 · ④ AI 요약 · ⑤ 스펙 · 설정)
   2차(2026-10-28): 「자주 검사 CHECK SHEET」 양식 · 출하일정 엑셀 · 모델/품번/호기를 검사표 내용으로 식별.
   계산은 js/logic.js(QCLogic), 저장은 js/store.js(QCStore), AI 연결은 js/ai-endpoint.js · js/ai-panel.js.
   뼈대(사이드바 · 표 · 알림 · AI 패널)는 data10-11, PDF 열기(pdf.js)는 data10-03 에서 가져왔습니다. */
(function () {
  'use strict';
  var L = window.QCLogic, S = window.QCStore, SM = window.QCSample, E = window.AIEndpoint, AIP = window.AIPanel, XLSX = window.XLSX;
  /* 판 맞추기 — 브라우저가 옛 logic.js 와 새 app.js 를 섞어 쓰면(Pages 캐시) 버튼과 결과가 어긋날 수 있어, 섞였으면 계산하지 않고 새로고침을 안내합니다.
     판은 scripts/bump-version.mjs 로 index.html(?v=) · logic.js 와 함께 올립니다. */
  var APP_VERSION = '2026-11-11.8';
  if (!L || L.VERSION !== APP_VERSION) {
    var m0 = document.getElementById('main') || document.body;
    m0.innerHTML = '<section class="card" id="versionMismatch"><h1>새 판이 나왔습니다 — 새로고침(Ctrl+F5)해 주세요</h1><p class="alert warn">이 브라우저가 옛 파일과 새 파일을 섞어 불러왔습니다(js/logic.js ' +
      (L && L.VERSION ? L.VERSION : '옛 판') + ' · 화면 ' + APP_VERSION + '). 이대로 계산하면 버튼과 결과가 어긋날 수 있어 멈췄습니다.</p>' +
      '<p>Windows 는 <b>Ctrl+F5</b>, Mac 은 <b>Cmd+Shift+R</b> 을 한 번 눌러 주세요. 쌓아 둔 자료는 지워지지 않습니다.</p></section>';
    return;
  }
  (function () { var ft = document.querySelector('.site-footer .inner') || document.querySelector('.site-footer'); if (!ft) return; var s = document.createElement('span'); s.className = 'ver'; s.textContent = ' · 판 ' + APP_VERSION; ft.appendChild(s); })();
  var db = S.loadDb();
  var ui = { cur: '', page: 1, zoom: 'fit', fileFilter: 'todo', drawMask: false, maskScope: 'sheet', focus: '', aiText: '', aiFor: '', lowConf: [], pendShip: null, pendSpec: null, busy: '' };
  var main = document.getElementById('main');

  /* ── 작은 도구 ───────────────────────────── */
  function h(tag, attrs) {
    var e = document.createElement(tag);
    if (attrs) Object.keys(attrs).forEach(function (k) {
      var v = attrs[k];
      if (v == null || v === false) return;
      if (k === 'class') e.className = v;
      else if (k === 'value') e.value = v;
      else if (k.slice(0, 2) === 'on' && typeof v === 'function') e.addEventListener(k.slice(2), v);
      else if (v === true) e.setAttribute(k, '');
      else e.setAttribute(k, v);
    });
    for (var i = 2; i < arguments.length; i++) add(e, arguments[i]);
    return e;
  }
  function add(e, c) {
    if (c == null || c === false) return;
    if (Array.isArray(c)) { c.forEach(function (x) { add(e, x); }); return; }
    e.appendChild(typeof c === 'object' ? c : document.createTextNode(String(c)));
  }
  function fld(label, control, hint) { return h('label', { class: 'field' }, h('span', null, label), control, hint ? h('small', { class: 'hint' }, hint) : null); }
  var toastTimer = null;
  function toast(msg, err) {
    var t = document.getElementById('toast');
    t.textContent = msg; t.className = 'toast' + (err ? ' error' : ''); t.hidden = false;
    clearTimeout(toastTimer); toastTimer = setTimeout(function () { t.hidden = true; }, 4500);
  }
  function save() {
    var ok = S.saveDb(db);
    document.getElementById('storeNote').textContent = ok ? '이 브라우저에 저장됨 · ' + S.sizeKb() + 'KB' : '저장 안 됨(브라우저 저장소가 막혔거나 가득 참) — 설정에서 JSON 백업을 받아 두세요';
    return ok;
  }
  function download(name, blob) {
    var a = h('a', { href: URL.createObjectURL(blob), download: name });
    document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  }
  function copyText(t) {
    function fallback() { var ta = h('textarea'); ta.value = t; document.body.appendChild(ta); ta.select(); try { document.execCommand('copy'); toast('복사했습니다.'); } catch (e) { toast('복사하지 못했습니다. 직접 선택해 복사해 주세요.', true); } ta.remove(); }
    if (navigator.clipboard && window.isSecureContext) navigator.clipboard.writeText(t).then(function () { toast('복사했습니다.'); }, fallback); else fallback();
  }
  function p2(n) { return (n < 10 ? '0' : '') + n; }
  function stamp() { var d = new Date(); return d.getFullYear() + p2(d.getMonth() + 1) + p2(d.getDate()) + '_' + p2(d.getHours()) + p2(d.getMinutes()); }
  function nowText() { var d = new Date(); return d.getFullYear() + '-' + p2(d.getMonth() + 1) + '-' + p2(d.getDate()) + ' ' + p2(d.getHours()) + ':' + p2(d.getMinutes()); }
  function todayText() { var d = new Date(); return d.getFullYear() + '-' + p2(d.getMonth() + 1) + '-' + p2(d.getDate()); }
  function refDate() { return L.toDate(db.refDate) || todayText(); }
  function clone(o) { return JSON.parse(JSON.stringify(o)); }
  function pageHead(stage, title, lead, right) {
    return h('div', { class: 'page-head' }, h('div', { class: 'titles' }, h('div', { class: 'stage' }, stage), h('h1', null, title), lead ? h('p', { class: 'lead' }, lead) : null), right || null);
  }
  function tile(k, v, sub, cls) { return h('div', { class: 'tile' }, h('div', { class: 'k' }, k), h('div', { class: 'v' + (cls ? ' ' + cls : '') }, v), sub ? h('div', { class: 'note' }, sub) : null); }
  function table(head, rows, opts) {
    opts = opts || {};
    var t = h('table', { class: 'list' + (opts.cls ? ' ' + opts.cls : '') });
    var tr = h('tr');
    head.forEach(function (x, i) { tr.appendChild(h('th', { class: opts.num && opts.num[i] ? 'num' : null, scope: 'col' }, x)); });
    t.appendChild(h('thead', null, tr));
    var tb = h('tbody');
    rows.forEach(function (r) {
      var row = h('tr', { class: r.cls || null });
      (r.cells || r).forEach(function (x, i) { row.appendChild(h('td', { class: ((opts.num && opts.num[i] ? 'num ' : '') + (opts.nw && opts.nw.indexOf(i) >= 0 ? 'nw' : '')).trim() || null }, x)); });
      tb.appendChild(row);
    });
    t.appendChild(tb);
    return h('div', { class: 'table-wrap' }, t);
  }
  function st(status, text) { return h('span', { class: 'st ' + status }, text || L.VERDICT_LABEL[status] || status); }
  var pct = function (v) { return L.pctText(v, 1); };
  function settings() { return L.cleanSettings(db.settings); }
  function analysis() { return L.analyze(db.ships, db.files, db.readings, db.overrides, db.settings, refDate()); }
  function fileById(id) { return db.files.filter(function (f) { return f.id === id; })[0] || null; }
  function ddayCell(n) { return h('span', { class: 'dd' + (n != null && n < 0 ? ' past' : n != null && n <= 3 ? ' soon' : '') }, L.dday(n)); }
  /* 판독 화면으로 — 파일과 쪽(검사표)을 골라 엽니다. manual 이면 바로 칸에 적을 수 있게 첫 빈 칸으로 */
  function readLink(fid, page, text, manual) {
    return h('a', { class: 'btn btn-sm', href: '#/read', onclick: function () { ui.cur = fid; ui.page = page || 1; ui.aiText = ''; ui.lowConf = []; ui.focus = manual ? L.sheetKey(fid, page || 1) : ''; } }, text || '판독');
  }
  function sheetName(f, page) { return L.sheetLabel(f, page || 1); }
  function rotOf(f, n) { return L.normRot((f.rot || {})[n]); }
  function emptyNotice(what) {
    return h('section', { class: 'card empty' },
      h('p', null, what || '아직 넣은 자료가 없습니다.'),
      h('p', null, '출하일정 엑셀과 검사표 PDF 를 넣거나, 가상 샘플로 먼저 둘러봐 주세요.'),
      h('div', { class: 'btn-row', style: 'justify-content:center' },
        h('a', { class: 'btn btn-primary', href: '#/import' }, '자료 넣기로'),
        h('button', { type: 'button', class: 'btn', onclick: loadSample }, '샘플로 시작')));
  }
  var SRC_LABEL = { content: '검사표 내용', filename: '파일 이름(확인 필요)' };
  var ID_LABEL = { ok: '식별됨', hint: '파일 이름으로만', unknown: '식별 안 됨' };

  /* ── pdf.js · PDF 바이트 ────────────────────── */
  // file:// 에서는 브라우저가 Worker 를 막으므로 worker 스크립트를 일반 스크립트로 먼저 넣어 화면 스레드에서 돌게 합니다(data10-03 과 같은 방식).
  var scriptWait = {};
  function loadScript(src, test) {
    if (test()) return Promise.resolve();
    if (scriptWait[src]) return scriptWait[src];
    scriptWait[src] = new Promise(function (res, rej) {
      var s = document.createElement('script'); s.src = src;
      s.onload = function () { res(); };
      s.onerror = function () { delete scriptWait[src]; rej(new Error(src + ' 를 불러오지 못했습니다. 폴더째 받았는지 확인해 주세요.')); };
      document.head.appendChild(s);
    });
    return scriptWait[src];
  }
  function withPdfLib() {
    return loadScript('vendor/pdfjs/pdf.worker.min.js', function () { return !!window.pdfjsWorker; })
      .then(function () { return loadScript('vendor/pdfjs/pdf.min.js', function () { return !!window.pdfjsLib; }); })
      .then(function () { window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdfjs/pdf.worker.min.js'; return window.pdfjsLib; });
  }
  function b64bytes(b64) { var s = atob(b64), a = new Uint8Array(s.length); for (var i = 0; i < s.length; i++) a[i] = s.charCodeAt(i); return a; }
  function pdfBytes(f) {
    if (f.sample) return loadScript('js/sample-pdfs.js?v=' + APP_VERSION, function () { return !!window.QCSamplePdfs; }).then(function () { var b = window.QCSamplePdfs[f.name]; return b ? b64bytes(b) : null; });
    return S.getPdf(f.id);
  }
  var docCache = {}, canvasCache = {};
  function resetCaches() { docCache = {}; canvasCache = {}; }
  function openDoc(f) {
    if (docCache[f.id]) return docCache[f.id];
    docCache[f.id] = Promise.all([withPdfLib(), pdfBytes(f)]).then(function (x) {
      if (!x[1]) throw new Error('nobytes');
      return x[0].getDocument({ data: x[1].slice ? x[1].slice() : x[1], isEvalSupported: false, disableFontFace: false, verbosity: 0 }).promise;
    });
    docCache[f.id].catch(function () { delete docCache[f.id]; });
    return docCache[f.id];
  }
  /* 쪽 그림 — width 픽셀 폭 */
  function pageCanvas(f, n, width) {
    var turn = rotOf(f, n);
    var k = f.id + '#' + n + '@' + width + 'r' + turn;
    if (canvasCache[k]) return canvasCache[k];
    canvasCache[k] = openDoc(f).then(function (pdf) {
      return pdf.getPage(Math.min(n, pdf.numPages)).then(function (page) {
        var rotation = ((page.rotate || 0) + turn) % 360;   // 돌린 방향(쪽마다 저장) — AI 에 보낼 그림도 이 방향
        var vp0 = page.getViewport({ scale: 1, rotation: rotation }), vp = page.getViewport({ scale: width / vp0.width, rotation: rotation });
        var c = document.createElement('canvas'); c.width = Math.round(vp.width); c.height = Math.round(vp.height);
        var g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
        return page.render({ canvasContext: g, viewport: vp }).promise.then(function () { c.dataset.pages = pdf.numPages; return c; });
      });
    });
    canvasCache[k].catch(function () { delete canvasCache[k]; });
    return canvasCache[k];
  }
  /* 이 검사표(쪽)에 쓰는 가릴 영역 — 파일(쪽)별 > 공정별 > 모든 검사표 */
  function procOf(sk) { var r = db.readings[sk]; return r && r.meta ? r.meta.process || '' : ''; }
  function masksOf(sk, proc) { return L.masksFor(db, sk, proc == null ? procOf(sk) : proc); }
  /* AI 에 보낼 그림 — 가릴 영역을 흰색으로 덮은 JPEG(돌린 방향 그대로) */
  function maskedImage(f, n, proc) {
    var ms = masksOf(L.sheetKey(f.id, n), proc).masks;
    return pageCanvas(f, n, 1600).then(function (src) {
      var c = document.createElement('canvas'); c.width = src.width; c.height = src.height;
      var g = c.getContext('2d'); g.drawImage(src, 0, 0);
      g.fillStyle = '#fff';
      ms.forEach(function (m) { g.fillRect(m.x * c.width, m.y * c.height, m.w * c.width, m.h * c.height); });
      return c;
    });
  }
  /* 글자 층 읽기 — 쪽마다 글자 조각을 자리(x·y)와 함께 꺼내 양식 칸으로 읽습니다(한 PDF 에 여러 쪽 · 여러 공정 — 2026-11-10 답 ④).
     글자가 누워 있으면(90° · 180°) 그 방향으로 돌려 읽고 화면도 그 방향으로 세웁니다. 스캔 그림은 글자가 없어 방향을 알 수 없으니 화면에서 돌립니다.
     f.pageText[쪽] = true(글자 층으로 읽음) · false(그림만) · 'partial'(글자는 있으나 양식이 다름). f.text 는 1쪽 값(2차 자료와 같음). */
  function readPage(f, pdf, n, force) {
    return pdf.getPage(n).then(function (page) {
      return page.getTextContent().then(function (tc) {
        var items = tc.items, turn = L.textRotation(items), norm;
        if (turn) { var vpr = page.getViewport({ scale: 1, rotation: turn }); norm = L.normTextItems(L.viewportTextItems(items, vpr.transform, vpr.height), vpr.height); }
        else norm = L.normTextItems(items, page.getViewport({ scale: 1 }).height);
        var r = L.parseSheetText(norm), sk = L.sheetKey(f.id, n);
        f.pageText = f.pageText || {};
        if (norm.length >= 8) {
          var extra = L.normRot(turn - (page.rotate || 0));
          if (extra && (!f.rot || f.rot[n] == null)) { (f.rot = f.rot || {})[n] = extra; (f.rotAuto = f.rotAuto || {})[n] = true; }
        }
        if (r.ok && r.items.length) {
          f.pageText[n] = true;
          if (force || !db.readings[sk]) db.readings[sk] = { meta: r.meta, items: r.items, hogis: r.hogis, problems: r.problems, auto: 'text', warnings: r.warnings, done: true, by: '글자 층 자동', at: nowText() };
        } else f.pageText[n] = r.reason === 'notext' ? false : 'partial';
        return r;
      });
    });
  }
  function autoRead(f) {
    return openDoc(f).then(function (pdf) {
      f.pages = pdf.numPages;
      var seq = Promise.resolve();
      for (var n = 1; n <= pdf.numPages; n++) (function (n) { seq = seq.then(function () { return readPage(f, pdf, n); }); })(n);
      return seq.then(function () { f.text = f.pageText[1]; });
    }).catch(function () { f.text = null; return null; });
  }
  function autoReadAll(list) {
    var n = { text: 0, image: 0, pages: 0 };
    ui.busy = '검사표 PDF 의 글자 층을 읽는 중…'; render();
    return list.reduce(function (p, f) { return p.then(function () { return autoRead(f).then(function () {
      for (var k = 1; k <= L.pagesOf(f); k++) { n.pages++; if ((f.pageText || {})[k] === true) n.text++; else n.image++; }
    }); }); }, Promise.resolve())
      .then(function () { ui.busy = ''; save(); render(); return n; });
  }

  /* ── 샘플 ───────────────────────────────── */
  function loadSample() {
    if ((db.ships.length || db.files.length) && !db.sample && !confirm('지금 자료를 지우고 가상 샘플을 불러올까요?')) return;
    var d = L.emptyDb();
    d.ships = SM.ships.map(function (s) { return Object.assign({}, s); });
    d.shipFile = SM.shipFile;
    d.files = SM.files.map(function (f) { return { id: L.fileId(f.name, f.size), name: f.name, size: f.size, at: '샘플', sample: true, hasPdf: true, pages: f.pages ? f.pages.length : 1 }; });
    d.masks = SM.masks.map(function (m) { return Object.assign({}, m); });
    d.maskProcs = clone(SM.maskProcs || {});
    d.refDate = SM.refDate;
    d.sample = true;
    db = d; ui.cur = ''; ui.aiText = ''; ui.lowConf = [];
    S.clearPdf(); resetCaches(); save();
    if (location.hash !== '#/ship') location.hash = '#/ship';
    autoReadAll(db.files).then(function (n) { toast('가상 샘플 — 출하 ' + db.ships.length + '대 · 검사표 PDF ' + db.files.length + '개 ' + n.pages + '쪽(글자 층 ' + n.text + ' · 그림만 ' + n.image + ').'); });
  }

  /* ── 자료 넣기 ───────────────────────────── */
  function readXlsx(file) {
    return new Promise(function (res, rej) {
      var fr = new FileReader();
      fr.onload = function () {
        try {
          var wb = XLSX.read(new Uint8Array(fr.result), { type: 'array', cellDates: true });
          res(wb.SheetNames.map(function (n) { return { name: n, aoa: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null }) }; }));
        } catch (e) { rej(new Error(file.name + ': 엑셀로 읽지 못했습니다(' + e.message + ')')); }
      };
      fr.onerror = function () { rej(new Error(file.name + ': 파일을 읽지 못했습니다.')); };
      fr.readAsArrayBuffer(file);
    });
  }
  function fieldsOf(kind) { return kind === 'ship' ? L.SHIP_FIELDS : L.SPEC_FIELDS; }
  function pickSheet(sheets, fields) {
    var best = sheets[0], score = -1;
    sheets.forEach(function (s) { var d = L.detectColumns(s.aoa, fields); var sc = Object.keys(d.map).length - d.missing.length * 3; if (sc > score) { score = sc; best = s; } });
    return best;
  }
  function loadTable(kind, file) {
    readXlsx(file).then(function (sheets) {
      var sh = pickSheet(sheets, fieldsOf(kind));
      var p = { kind: kind, file: file.name, sheets: sheets, sheet: sh.name, override: {} };
      parsePending(p);
      if (kind === 'ship') ui.pendShip = p; else ui.pendSpec = p;
      render();
    }, function (e) { toast(e.message, true); });
  }
  /* 출하일정 엑셀 여러 개를 한 번에(11-11 요청 — 소형 · 대형 · 초대형 파일이 따로) — 파일마다 자동으로 읽어 합칩니다.
     칸을 못 찾은 파일은 넣지 않고 알려 줍니다(그 파일만 따로 골라 열 연결을 맞춥니다). */
  function loadShipMany(files) {
    Promise.all(files.map(function (file) {
      return readXlsx(file).then(function (sheets) {
        var sh = pickSheet(sheets, fieldsOf('ship')), p = { kind: 'ship', file: file.name, sheets: sheets, sheet: sh.name, override: {} };
        parsePending(p); return p;
      }, function (e) { return { file: file.name, error: e.message }; });
    })).then(function (ps) {
      var good = ps.filter(function (p) { return !p.error && !p.result.det.missing.length && p.result.rows.length; });
      var bad = ps.filter(function (p) { return good.indexOf(p) < 0; });
      if (!good.length) { toast('읽을 수 있는 출하일정이 없습니다 — 한 개씩 골라 열 연결을 확인해 주세요.', true); return; }
      if (db.sample) { if (!confirm('지금은 가상 샘플입니다. 샘플을 모두 지우고 이 엑셀 ' + good.length + '개로 시작할까요?')) return; startFresh(); }
      else if (db.ships.length && !confirm('지금 출하일정(' + db.ships.length + '대)에 엑셀 ' + good.length + '개를 더할까요? (같은 모델 · 호기는 새 파일 것으로)')) return;
      var before = db.ships.length, names = [];
      good.forEach(function (p) {
        var rows = p.result.rows.map(function (r) { var o = Object.assign({}, r); delete o.row; o.file = p.file; return o; });
        db.ships = db.ships.length ? L.mergeShips(db.ships, rows) : rows; names.push(p.file);
      });
      db.shipFile = (before && db.shipFile ? db.shipFile + ' + ' : '') + names.join(' + '); ui.pendShip = null;
      save();
      toast('출하일정 엑셀 ' + good.length + '개를 넣었습니다 — ' + before + '대 → ' + db.ships.length + '대.' + (bad.length ? ' 읽지 못한 파일 ' + bad.length + '개(' + bad.map(function (p) { return p.file; }).join(', ') + ')는 한 개씩 골라 주세요.' : ''), !!bad.length);
      render();
    });
  }
  function parsePending(p) {
    var sh = p.sheets.filter(function (s) { return s.name === p.sheet; })[0];
    p.result = p.kind === 'ship' ? L.parseShip(sh.aoa, p.override, { year: +refDate().slice(0, 4), only3: p.only3 !== false }) : L.parseSpec(sh.aoa, p.override);
  }
  function startFresh() { db = L.emptyDb(); db.cleared = true; S.clearPdf(); resetCaches(); }
  function pendingCard(p) {
    var fields = p.kind === 'ship' ? L.SHIP_FIELDS.filter(function (f) { return p.only3 === false || ['date', 'model', 'hogi'].indexOf(f.id) >= 0; }) : fieldsOf(p.kind), det = p.result.det, hdr = det.header;
    // 출하일정은 「모델 · 호기 · 출하예정」 세 칸만(2026-11-10 4차 ③) — 끄면 품번 · 수량 · 품명도 읽음
    var only3 = h('input', { type: 'checkbox' }); only3.checked = p.only3 !== false;
    only3.addEventListener('change', function () { p.only3 = only3.checked; p.override = {}; parsePending(p); render(); });
    var sheetSel = h('select', { 'aria-label': '시트' }, p.sheets.map(function (s) { return h('option', { value: s.name, selected: s.name === p.sheet }, s.name); }));
    sheetSel.addEventListener('change', function () { p.sheet = sheetSel.value; p.override = {}; parsePending(p); render(); });
    var grid = h('div', { class: 'map-grid' });
    fields.forEach(function (f) {
      var sel = h('select', { 'aria-label': f.label + ' 열' }, h('option', { value: '' }, '(없음)'));
      hdr.forEach(function (t, i) { sel.appendChild(h('option', { value: String(i), selected: det.map[f.id] === i }, (t || '(빈 머리)') + ' — ' + XLSX.utils.encode_col(i) + '열')); });
      sel.addEventListener('change', function () { p.override[f.id] = sel.value === '' ? null : +sel.value; parsePending(p); render(); });
      grid.appendChild(h('div', { class: 'map-row' + (f.req && det.map[f.id] == null ? ' missing' : '') }, h('span', { class: 'col' }, f.label, f.req ? h('span', { class: 'req' }, ' 필수') : null), sel));
    });
    var ok = !det.missing.length && p.result.rows.length;
    var preview = p.kind === 'ship'
      ? (p.only3 !== false ? table(['출하예정', '모델', '호기'], p.result.rows.slice(0, 8).map(function (r) { return [r.date || '(날짜 없음)', r.model, r.hogi]; }), { nw: [0, 2] })
        : table(['출하일', '모델', '품번', '호기', '수량'], p.result.rows.slice(0, 8).map(function (r) { return [r.date || '(날짜 없음)', r.model, r.part, r.hogi, r.qty == null ? '' : r.qty]; }), { nw: [0, 3] }))
      : table(['품번', '검사 항목(또는 No)', '기준', '방식'], p.result.rows.slice(0, 8).map(function (r) { return [r.part, r.item, L.specText(r), r.kind === 'okng' ? '육안' : '수치']; }));
    return h('section', { class: 'card' },
      h('h2', null, (p.kind === 'ship' ? '출하일정' : '스펙 덮어쓰기 표') + ' 확인 — ' + p.file),
      h('div', { class: 'form-grid' }, fld('시트', sheetSel), h('p', { class: 'note span-all', style: 'margin:0' }, '머리행: ' + (det.headerRow + 1) + '번째 줄 · 읽은 ' + (p.kind === 'ship' ? '호기 ' : '줄 ') + p.result.rows.length + '개')),
      p.kind === 'ship' ? h('label', { class: 'opt', style: 'margin-top:10px' }, only3, h('span', null, h('span', { class: 't' }, '모델 · 호기 · 출하예정 세 칸만 읽기(권장)'), h('span', { class: 's' }, '소형 · 대형 · 초대형 파일마다 다른 상부 · 하부 · Line · 비고 같은 칸은 읽지도 알리지도 않습니다. 품번 칸으로도 짝지으려면 끄세요.'))) : null,
      h('h3', { style: 'margin-top:14px' }, '열 연결 (자동으로 찾은 것 — 다르면 골라 주세요)'), grid,
      det.missing.length ? h('p', { class: 'alert warn' }, '필수 칸을 찾지 못했습니다: ' + det.missing.join(', ')) : null,
      p.result.warnings.length ? h('details', { class: 'warn-list' }, h('summary', null, '알림 ' + p.result.warnings.length + '건'), h('ul', null, p.result.warnings.slice(0, 60).map(function (w) { return h('li', null, (w.row ? w.row + '행: ' : '') + (w.msg || w)); }))) : null,
      h('h3', { style: 'margin-top:14px' }, '미리 보기'), preview,
      h('div', { class: 'btn-row', style: 'margin-top:12px' },
        h('button', { type: 'button', class: 'btn btn-primary', disabled: !ok, onclick: function () {
          if (db.sample) { if (!confirm('지금은 가상 샘플입니다. 샘플을 모두 지우고 이 엑셀로 시작할까요?')) return; var keepA = ui.pendShip, keepB = ui.pendSpec; startFresh(); ui.pendShip = keepA; ui.pendSpec = keepB; }
          if (p.kind === 'ship') { db.ships = p.result.rows.map(function (r) { var o = Object.assign({}, r); delete o.row; o.file = p.file; return o; }); db.shipFile = p.file; ui.pendShip = null; }
          else { db.overrides = p.result.rows.map(function (r) { var o = Object.assign({}, r); delete o.error; return o; }); db.overrideFile = p.file; ui.pendSpec = null; }
          save(); toast('반영했습니다.'); render();
        } }, '이 내용으로 바꾸기'),
        // 출하일정은 소형 · 초대형 · 버켓처럼 파일이 나뉘어 있습니다(2026-11-10 답 ③) — 지금 출하일정에 더하기
        p.kind === 'ship' && db.ships.length && !db.sample ? h('button', { type: 'button', class: 'btn', disabled: !ok, onclick: function () {
          var add0 = p.result.rows.map(function (r) { var o = Object.assign({}, r); delete o.row; o.file = p.file; return o; });
          var before = db.ships.length;
          db.ships = L.mergeShips(db.ships, add0);
          db.shipFile = (db.shipFile ? db.shipFile + ' + ' : '') + p.file; ui.pendShip = null;
          save(); toast('출하일정에 더했습니다 — ' + before + '대 → ' + db.ships.length + '대(같은 모델 · 호기는 새 파일 것으로).'); render();
        } }, '지금 출하일정에 더하기') : null,
        h('button', { type: 'button', class: 'btn', onclick: function () { if (p.kind === 'ship') ui.pendShip = null; else ui.pendSpec = null; render(); } }, '취소')));
  }
  function addPdfs(list, group) {
    var all = Array.prototype.slice.call(list || []), files = all.filter(function (f) { return /\.pdf$/i.test(f.name); });
    if (!files.length) { toast('PDF 파일이 없습니다.', true); return; }
    if (db.sample) { if (!confirm('지금은 가상 샘플입니다. 샘플을 지우고 내 PDF 로 시작할까요? (출하일정도 함께 지워집니다)')) return; startFresh(); }
    var added = [], dup = 0;
    var jobs = files.map(function (f) {
      var id = L.fileId(f.name, f.size);
      if (fileById(id)) { dup++; return Promise.resolve(); }
      return f.arrayBuffer().then(function (buf) {
        return S.putPdf(id, new Uint8Array(buf)).then(function (stored) {
          // 구분(고객사 등) — 적어 준 이름, 없으면 「폴더째 고르기」의 바로 위 폴더 이름(11-11 요청: 고객사별로 나눠 보기)
          var parts = String(f.webkitRelativePath || '').split('/'), g = String(group || '').trim() || (parts.length > 1 ? parts[parts.length - 2] : '');
          var rec = { id: id, name: f.name, size: f.size, at: nowText(), hasPdf: true, stored: stored };
          if (g) rec.group = g.slice(0, 40);
          db.files.push(rec); added.push(rec);
        });
      });
    });
    Promise.all(jobs).then(function () {
      db.files.sort(function (a, b) { return a.name < b.name ? -1 : 1; });
      save();
      return autoReadAll(added);
    }).then(function (n) {
      toast('PDF ' + added.length + '개(' + n.pages + '쪽)를 넣었습니다 — 글자 층에서 읽음 ' + n.text + '쪽 · 그림만(사람 · AI 판독) ' + n.image + '쪽' + (dup ? ' · 이미 있는 ' + dup + '개는 건너뜀' : '') + (all.length - files.length ? ' · PDF 아닌 파일 ' + (all.length - files.length) + '개 건너뜀' : '') + '.');
    }, function (e) { ui.busy = ''; toast('PDF 를 넣지 못했습니다: ' + e.message, true); render(); });
  }
  function groupsOf() { return db.files.map(function (f) { return f.group || ''; }).filter(function (g, i, a) { return g && a.indexOf(g) === i; }).sort(); }
  function kindLabel(f, page) {
    var t = f.pageText && f.pageText[page || 1] !== undefined ? f.pageText[page || 1] : f.text;
    return t === true ? '글자 층' : t === false ? '그림만(스캔)' : t === 'partial' ? '글자 층(양식 다름)' : '확인 전';
  }
  function removeFile(f) {
    db.files = db.files.filter(function (x) { return x !== f; });
    for (var k = 1; k <= L.pagesOf(f); k++) { var sk = L.sheetKey(f.id, k); delete db.readings[sk]; if (db.maskSheets) delete db.maskSheets[sk]; }
    S.delPdf(f.id); save(); render();
  }
  function pageImport() {
    var an = analysis(), c = an.counts;
    var shipIn = h('input', { type: 'file', accept: '.xlsx,.xls,.csv', multiple: true });
    shipIn.addEventListener('change', function () { var fs = Array.prototype.slice.call(shipIn.files || []); if (fs.length === 1) loadTable('ship', fs[0]); else if (fs.length > 1) loadShipMany(fs); shipIn.value = ''; });
    var specIn = h('input', { type: 'file', accept: '.xlsx,.xls,.csv' }); specIn.addEventListener('change', function () { if (specIn.files[0]) loadTable('spec', specIn.files[0]); specIn.value = ''; });
    var grpIn = h('input', { type: 'text', class: 'grp-in', maxlength: 40, placeholder: '구분(고객사 등 · 선택)', 'aria-label': '넣는 검사표의 구분', list: 'grpList', value: ui.addGroup || '' });
    grpIn.addEventListener('input', function () { ui.addGroup = grpIn.value; });
    var pdfIn = h('input', { type: 'file', accept: '.pdf,application/pdf', multiple: true }); pdfIn.addEventListener('change', function () { addPdfs(pdfIn.files, grpIn.value); pdfIn.value = ''; });
    var dirIn = h('input', { type: 'file', webkitdirectory: true, multiple: true }); dirIn.addEventListener('change', function () { addPdfs(dirIn.files, grpIn.value); dirIn.value = ''; });
    var drop = h('div', { class: 'drop' }, h('span', null, 'PDF 를 여기에 끌어 놓거나'), h('label', { class: 'btn btn-primary' }, pdfIn, 'PDF 여러 개 고르기'), h('label', { class: 'btn' }, dirIn, '폴더째 고르기'),
      grpIn, h('datalist', { id: 'grpList' }, groupsOf().map(function (g) { return h('option', { value: g }); })),
      h('span', { class: 'note', style: 'flex-basis:100%' }, '구분을 적고 넣으면 ② 판독에서 그 구분만 골라 볼 수 있습니다. 비워 두고 「폴더째 고르기」를 쓰면 폴더 이름이 구분이 됩니다.'));
    drop.addEventListener('dragover', function (e) { e.preventDefault(); drop.classList.add('over'); });
    drop.addEventListener('dragleave', function () { drop.classList.remove('over'); });
    drop.addEventListener('drop', function (e) { e.preventDefault(); drop.classList.remove('over'); addPdfs(e.dataTransfer.files); });
    function done(ok, text) { return ok ? h('span', { class: 'done' }, text) : null; }
    var idOf = {}, rank = { unknown: 0, hint: 1, ok: 2 };
    an.ids.forEach(function (i) { var o = idOf[i.file.id]; if (!o) { idOf[i.file.id] = { status: i.status, parts: [], hogis: [] }; o = idOf[i.file.id]; } if (rank[i.status] < rank[o.status]) o.status = i.status; var pm = i.meta.part || i.meta.model; if (pm && o.parts.indexOf(pm) < 0) o.parts.push(pm); i.hogis.forEach(function (x) { if (o.hogis.indexOf(x) < 0) o.hogis.push(x); }); });
    return [
      pageHead('자료 넣기', '출하일정 · 검사표 PDF', '출하일정과 공유폴더의 검사표를 넣으면, 검사표 내용으로 모델 · 품번 · 호기를 찾아 출하 예정 호기와 짝짓고(①) 합불을 판정합니다(③). 파일은 이 브라우저 안에서만 읽습니다.',
        h('div', { class: 'btn-row' }, h('button', { type: 'button', class: 'btn btn-primary', onclick: loadSample }, '샘플로 시작'))),
      h('div', { class: 'steps-row', style: 'margin-bottom:20px' },
        h('section', { class: 'step-card' }, h('h3', null, '1. 출하일정 엑셀', done(db.ships.length, db.ships.length + '대')),
          h('p', { class: 'note' }, '「언제 · 어떤 모델/품번 · 몇 호기를 내보내는지」를 읽습니다. 열: 출하일(출하예정) · 모델 · 품번 · 호기 · 수량(이름이 조금 달라도 찾고, 못 찾으면 골라 연결). 「314498~314500」 처럼 적힌 호기는 펼치고, 「8/10」 처럼 연도가 없으면 기준일의 연도로 봅니다. 소형 · 초대형 등 파일이 여러 개면 하나씩 「더하기」로 합칩니다.'),
          db.shipFile ? h('p', { class: 'note' }, '지금: ' + db.shipFile) : null,
          h('label', { class: 'btn' }, shipIn, '엑셀 고르기(여러 개 가능)')),
        h('section', { class: 'step-card' }, h('h3', null, '2. 검사표 PDF', done(db.files.length, db.files.length + '장')),
          h('p', { class: 'note' }, '파일 이름은 따지지 않습니다. 한 PDF 에 여러 쪽(호기 · 공정)이 있으면 쪽마다 검사표 한 장으로 봅니다. 글자 층이 있는 쪽은 칸 자리로 바로 읽고, 스캔 그림은 ② 판독에서 그림을 보며 옮기거나 AI 로 읽습니다.'),
          h('a', { class: 'btn', href: '#/ship' }, '① 출하 대조로')),
        h('section', { class: 'step-card' }, h('h3', null, '3. 스펙 덮어쓰기 표 (선택)', done(db.overrides.length, db.overrides.length + '줄')),
          h('p', { class: 'note' }, '규격은 검사표의 「420±1」 같은 글자에서 바로 읽습니다. 검사표 글자와 다른 기준을 쓰고 싶을 때만 품번 · 검사 항목(또는 No) · 하한 · 상한 표를 넣습니다.'),
          db.overrideFile ? h('p', { class: 'note' }, '지금: ' + db.overrideFile) : null,
          h('label', { class: 'btn' }, specIn, '엑셀 고르기'))),
      ui.pendShip ? pendingCard(ui.pendShip) : null,
      ui.pendSpec ? pendingCard(ui.pendSpec) : null,
      h('section', { class: 'card' }, h('h2', null, '검사표 PDF 넣기'), drop,
        db.files.length ? h('details', { style: 'margin-top:8px', open: db.files.length <= 12 }, h('summary', null, '넣은 파일 ' + db.files.length + '개'),
          table(['파일', '쪽', '종류', '식별', '모델 · 품번 · 호기', ''], db.files.map(function (f) {
            var i = idOf[f.id] || {};
            return [f.name, L.pagesOf(f), kindLabel(f), st(i.status === 'ok' ? 'pass' : i.status === 'hint' ? 'hold' : 'wait', ID_LABEL[i.status] || '-'),
              [(i.parts || []).join(', '), (i.hogis || []).join(', ')].filter(Boolean).join(' · ') || '-',
              h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { removeFile(f); } }, '빼기')];
          }), { nw: [1] })) : null),
      db.ships.length || db.files.length ? h('section', { class: 'card' }, h('h2', null, '지금 상태'),
        h('p', null, '출하 ' + c.ships + '대 → 합격 ' + c.pass + ' · 불합격 ' + c.fail + ' · 미판정 ' + c.undecided + ' · 검사표 없음 ' + c.none + ' · 식별 필요 파일 ' + (c.unknown + c.hint)),
        h('div', { class: 'btn-row' }, h('a', { class: 'btn btn-primary', href: '#/ship' }, '① 출하 대조'), h('a', { class: 'btn', href: '#/read' }, '② 판독'), h('a', { class: 'btn', href: '#/result' }, '③ 검사 결과'))) : null
    ];
  }

  /* ── ① 출하 대조 ─────────────────────────── */
  function shipCells(r, extra) {
    var u = r.unit;
    var undecided = r.status === 'wait' || r.status === 'hold';
    return [r.ship.date || '-', ddayCell(r.days), r.ship.model || '', r.ship.part || '', h('b', null, r.ship.hogi)].concat(extra || [])
      .concat([u ? sheetName(u.file, u.page) : '-', u && u.file.hasPdf !== false ? readLink(u.fileId, u.page, undecided ? '수동 판독' : '보기', undecided) : '']);
  }
  function failText(j) { return j.fails.map(function (x) { return 'No ' + x.key + ' ' + x.raw; }).join(', '); }
  function holdText(r) {
    if (r.status === 'wait') return r.unit.idStatus === 'hint' ? '판독 대기 — 파일 이름으로만 식별' : '판독 대기';
    var j = r.unit.judge, t = [];
    if (j.blanks.length) t.push('빈 칸 ' + j.blanks.map(function (x) { return x.key; }).join(', '));
    if (j.unreads.length) t.push('확인 필요 ' + j.unreads.map(function (x) { return x.key + '(' + x.msg + ')'; }).join(', '));
    if (j.msg) t.push(j.msg);
    return t.join(' · ');
  }
  function pageShip() {
    var head0 = pageHead('①', '출하 대조 — 검사표 없음 · 불합격 · 미판정');
    if (!db.ships.length && !db.files.length) return [head0, emptyNotice()];
    var an = analysis(), c = an.counts;
    var none = L.shipList(an, ['none']), fail = L.shipList(an, ['fail']), und = L.shipList(an, ['wait', 'hold']);
    var idNeed = an.ids.filter(function (i) { return i.status !== 'ok'; });
    var dIn = h('input', { type: 'date', value: refDate(), 'aria-label': '기준일' });
    dIn.addEventListener('change', function () { db.refDate = L.toDate(dIn.value) || ''; save(); render(); });
    var withU = ['출하일', '남은 날', '모델', '품번', '호기'];
    return [
      pageHead('①', '출하 대조 — 검사표 없음 · 불합격 · 미판정', '출하일정의 호기마다 검사 기록을 품번 + 호기로(품번이 없으면 모델 + 호기로) 찾아 짝짓습니다. 출하일이 가까운 순서입니다.',
        h('div', { class: 'ref-date' }, fld('기준일(남은 날 계산)', dIn), h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { db.refDate = ''; save(); render(); } }, '오늘로'))),
      ui.busy ? h('p', { class: 'alert info' }, ui.busy) : null,
      !db.ships.length ? h('p', { class: 'alert warn' }, '출하일정이 없어 「검사표 없음」을 셀 수 없습니다. 자료 넣기에서 출하일정 엑셀을 넣어 주세요.') : null,
      h('div', { class: 'tiles' },
        tile('출하 예정', c.ships + '대', db.shipFile || ''),
        tile('검사표 없음', c.none + '대', '출하 예정인데 짝지은 검사표가 없음', c.none ? 'over' : 'ok'),
        tile('불합격', c.fail + '대', '한 칸이라도 공차 밖 · 육안 NG', c.fail ? 'over' : 'ok'),
        tile('미판정', c.undecided + '대', '판독 대기 ' + c.wait + ' · 보류 ' + c.hold)),
      h('section', { class: 'card' }, h('h2', null, '출하 예정인데 검사표 없음 (' + none.length + ')'),
        none.length ? table(withU.concat(['수량']), none.map(function (r) { return { cls: 'row-missing', cells: [r.ship.date || '-', ddayCell(r.days), r.ship.model || '', r.ship.part || '', h('b', null, r.ship.hogi), r.ship.qty == null ? '' : r.ship.qty] }; }), { nw: [0, 1, 4] })
          : h('p', { class: 'note' }, '출하 예정 호기마다 검사표가 있습니다.'),
        none.length && idNeed.length ? h('p', { class: 'note', style: 'margin-top:8px' }, '아직 식별하지 못한 검사표가 ' + idNeed.length + '장 있습니다(아래). 그 검사표의 호기를 적으면 이 목록에서 빠질 수 있습니다.') : null),
      h('section', { class: 'card' }, h('h2', null, '불합격 (' + fail.length + ')'),
        fail.length ? table(withU.concat(['벗어난 칸', '검사표', '']), fail.map(function (r) { return { cls: 'row-fail', cells: shipCells(r, [failText(r.unit.judge)]) }; }), { nw: [0, 1, 4] })
          : h('p', { class: 'note' }, '불합격인 출하 예정 호기가 없습니다.')),
      h('section', { class: 'card' }, h('h2', null, '미판정 (' + und.length + ')'),
        und.length ? table(withU.concat(['상태', '내용', '검사표', '']), und.map(function (r) {
          return shipCells(r, [st(r.status), [holdText(r), r.notes.length ? h('div', { class: 'note' }, r.notes.join(' · ')) : null]]);
        }), { nw: [0, 1, 4] }) : h('p', { class: 'note' }, '판독 대기 · 보류인 출하 예정 호기가 없습니다.')),
      h('section', { class: 'card' }, h('h2', null, '식별이 필요한 검사표 (' + idNeed.length + ')'),
        h('p', { class: 'note' }, '파일 이름에 규칙이 없어 모델 · 품번 · 호기는 검사표 내용에서 찾습니다. 그림만 있는 스캔은 ② 판독에서 머리 칸과 호기를 적어 주세요. 「파일 이름으로만」 은 이름에서 짐작한 것이라 검사표와 대조가 필요합니다.'),
        idNeed.length ? table(['검사표', '종류', '상태', '짐작한 것', '모자란 것', ''], idNeed.map(function (i) {
          var g = [i.meta.model, i.meta.part, i.hogis.join(', ')].filter(Boolean).join(' · ');
          return [sheetName(i.file, i.page), kindLabel(i.file, i.page), st(i.status === 'hint' ? 'hold' : 'wait', ID_LABEL[i.status]), g || '-', i.missing.join(', ') || (i.status === 'hint' ? '검사표로 확인' : ''), i.file.hasPdf !== false ? readLink(i.file.id, i.page, '수동 판독 · 태그', true) : ''];
        })) : h('p', { class: 'note' }, '모든 검사표의 모델/품번 · 호기를 찾았습니다.')),
      an.extra.length ? h('section', { class: 'card' }, h('h2', null, '출하일정에 없는 검사 (' + an.extra.length + ')'),
        h('p', { class: 'note' }, '검사표는 있으나 출하일정에 같은 품번(모델) · 호기가 없습니다. 출하일정 기간 밖이거나 호기를 잘못 적었을 수 있습니다.'),
        table(['작성일', '모델', '품번', '공정', '호기', '판정', '검사표', ''], an.extra.map(function (u) { return [u.meta.date || '-', u.meta.model || '', u.meta.part || '', u.meta.process || '', u.hogi, st(u.judge.verdict), sheetName(u.file, u.page), readLink(u.fileId, u.page, '보기')]; }), { nw: [0, 4] })) : null,
      db.ships.length ? h('details', { class: 'block' }, h('summary', null, '출하일정 전체 ' + an.rows.length + '대'),
        table(withU.concat(['상태', '알림', '검사표', '']), an.rows.slice().sort(function (a, b) { return (a.ship.date || '') < (b.ship.date || '') ? -1 : 1; }).map(function (r) {
          return { cls: r.status === 'fail' ? 'row-fail' : r.status === 'none' ? 'row-missing' : null, cells: shipCells(r, [st(r.status), r.notes.join(' · ')]) };
        }), { nw: [0, 1, 4] })) : null
    ];
  }

  /* ── ② 판독 ─────────────────────────────── */
  function idInfoMap(an) { var m = {}; an.ids.forEach(function (i) { m[i.key] = i; }); return m; }
  /* 검사표(쪽) 하나의 판정 — 호기를 못 찾았거나 판독 전이면 'wait' */
  function sheetVerdict(an, sk) {
    var us = an.units.filter(function (u) { return u.sheet === sk; });
    if (!us.length || us.some(function (u) { return !u.done; })) return 'wait';
    if (us.some(function (u) { return u.judge.verdict === 'fail'; })) return 'fail';
    if (us.some(function (u) { return u.judge.verdict === 'hold'; })) return 'hold';
    return 'pass';
  }
  /* 파일 전체 — 쪽 가운데 하나라도 판독 대기면 대기, 아니면 불합격 > 보류 > 합격 */
  function fileVerdict(an, fid) {
    var f = fileById(fid), vs = f ? L.sheetPages(f).map(function (pg) { return sheetVerdict(an, L.sheetKey(fid, pg)); }) : ['wait'];
    return ['wait', 'fail', 'hold'].filter(function (v) { return vs.indexOf(v) >= 0; })[0] || 'pass';
  }
  function sheetTodo(an, ids, f, pg) { var sk = L.sheetKey(f.id, pg), v = sheetVerdict(an, sk), i = ids[sk]; return v === 'wait' || v === 'hold' || (i && i.status !== 'ok'); }
  function pageRead() {
    var head0 = pageHead('②', '판독 — 검사표 보며 값 옮기기');
    var withPdf = db.files.filter(function (f) { return f.hasPdf !== false; });
    if (!withPdf.length) return [head0, emptyNotice('아직 넣은 검사표 PDF 가 없습니다.')];
    var an = analysis(), ids = idInfoMap(an);
    function todo(f) { return L.sheetPages(f).some(function (pg) { return sheetTodo(an, ids, f, pg); }); }
    var groups = groupsOf(); if (ui.grp && ui.grp !== '(없음)' && groups.indexOf(ui.grp) < 0) ui.grp = '';
    var list = withPdf.filter(function (f) {
      if (ui.grp && (ui.grp === '(없음)' ? !!f.group : f.group !== ui.grp)) return false;
      if (ui.fileFilter === 'todo') return todo(f);
      if (ui.fileFilter === 'fail') return L.sheetPages(f).some(function (pg) { return sheetVerdict(an, L.sheetKey(f.id, pg)) === 'fail'; });
      return true;
    });
    list = list.map(function (x, i) { return [x, fileVerdict(an, x.id) === 'wait' ? 0 : 1, i]; }).sort(function (a, b) { return a[1] - b[1] || a[2] - b[2]; }).map(function (x) { return x[0]; });
    if (!ui.cur || !fileById(ui.cur)) { ui.cur = (list[0] || withPdf[0]).id; ui.page = firstTodoPage(fileById(ui.cur)); }
    var f = fileById(ui.cur);
    if (ui.page > L.pagesOf(f)) ui.page = 1;
    var idx = list.indexOf(f);
    function firstTodoPage(x) { var pgs = L.sheetPages(x); return pgs.filter(function (pg) { return sheetTodo(an, ids, x, pg); })[0] || pgs[0] || 1; }
    function pick(n) { ui.cur = n.id; ui.page = firstTodoPage(n); ui.aiText = ''; ui.lowConf = []; render(); }
    function go(d) { var arr = list.length ? list : withPdf, i = arr.indexOf(f); pick(arr[(i + d + arr.length) % arr.length]); }
    var fileSel = h('select', { 'aria-label': '검사표 파일' }, (idx < 0 ? [f].concat(list) : list).map(function (x) {
      var v = fileVerdict(an, x.id), np = L.pagesOf(x);
      var unk = L.sheetPages(x).some(function (pg) { var i = ids[L.sheetKey(x.id, pg)]; return i && i.status !== 'ok'; });
      return h('option', { value: x.id, selected: x.id === f.id }, '[' + L.VERDICT_LABEL[v] + (unk ? ' · 식별 필요' : '') + '] ' + x.name + (np > 1 ? ' (' + np + '쪽)' : ''));
    }));
    fileSel.addEventListener('change', function () { pick(fileById(fileSel.value)); });
    var filt = h('div', { class: 'tabs', role: 'group', 'aria-label': '파일 거르기', style: 'margin:0' }, [['todo', '할 일(대기 · 보류 · 식별)'], ['fail', '불합격'], ['all', '전체 ' + withPdf.length]].map(function (x) {
      return h('button', { type: 'button', 'aria-pressed': ui.fileFilter === x[0] ? 'true' : 'false', onclick: function () { ui.fileFilter = x[0]; ui.cur = ''; render(); } }, x[1]);
    }));
    // 구분(고객사 등)으로 거르기 · 지금 파일의 구분 바꾸기
    function grpBar() {
      var sel = h('select', { class: 'grp-sel', 'aria-label': '구분으로 거르기' }, [['', '구분: 전체']].concat(groups.map(function (g) { return [g, g + ' (' + withPdf.filter(function (x) { return x.group === g; }).length + ')']; })).concat(groups.length ? [['(없음)', '구분 없음']] : []).map(function (o) { return h('option', { value: o[0], selected: (ui.grp || '') === o[0] }, o[1]); }));
      sel.addEventListener('change', function () { ui.grp = sel.value; ui.cur = ''; render(); });
      return h('span', { class: 'grp-bar' }, sel, h('button', { type: 'button', class: 'btn btn-sm', onclick: function () {
        var v = prompt('이 검사표 파일의 구분(고객사 등)을 적어 주세요. 비우면 구분 없음.', f.group || ''); if (v === null) return;
        v = v.trim().slice(0, 40); if (v) f.group = v; else delete f.group; save(); render();
      } }, f.group ? '구분: ' + f.group + ' (바꾸기)' : '이 파일 구분 정하기'));
    }
    var sk = L.sheetKey(f.id, ui.page);
    return [
      pageHead('②', '판독 — 검사표 보며 값 옮기기', '글자 층이 있는 검사표는 이미 읽어 두었고, 스캔 그림은 왼쪽을 보며 오른쪽 칸에 적습니다(호기 열 × 항목 줄). 적는 즉시 호기마다 합불이 나옵니다. 여러 쪽 PDF 는 쪽마다 따로 판독합니다.'),
      ui.busy ? h('p', { class: 'alert info' }, ui.busy) : null,
      h('div', { class: 'read-bar' }, filt, grpBar()),
      h('div', { class: 'read-bar' },
        h('button', { type: 'button', class: 'btn', onclick: function () { go(-1); } }, '이전'),
        fileSel,
        h('button', { type: 'button', class: 'btn', onclick: function () { go(1); } }, '다음'),
        list.length ? h('span', { class: 'note' }, (idx >= 0 ? (idx + 1) + ' / ' : '') + list.length + '개') : h('span', { class: 'note' }, '거른 목록이 비었습니다')),
      h('div', { class: 'read-grid' }, viewerCard(f, sk, an), readForm(f, sk, ids[sk], an))
    ];
  }
  var SCOPE_LABEL = { sheet: '이 파일 · 이 쪽만', process: '이 공정 기본', all: '모든 검사표 기본' };
  /* 고를 범위의 가릴 영역 배열 — create 면 아래 단계(공정 → 모든 검사표)를 복사해 새로 만듦 */
  function scopeArr(scope, sk, proc, create) {
    db.maskSheets = db.maskSheets || {}; db.maskProcs = db.maskProcs || {};
    if (scope === 'all') return db.masks;
    var pk = L.procKey(proc);
    if (scope === 'process') { if (!pk) return null; if (!Array.isArray(db.maskProcs[pk]) && create) db.maskProcs[pk] = clone(L.masksFor(Object.assign({}, db, { maskProcs: {} }), '\u0000', proc).masks); return db.maskProcs[pk] || null; }
    if (!Array.isArray(db.maskSheets[sk]) && create) db.maskSheets[sk] = clone(L.masksFor(db, '\u0000', proc).masks);
    return db.maskSheets[sk] || null;
  }
  /* 공정의 가릴 영역 틀 고르기(2026-11-10 4차 ②) — '' = 공정명으로 자동, 'none' = 틀 안 씀(모든 검사표 기본), 그 밖 = 틀 id */
  function tplSelect(proc) {
    var pk = L.procKey(proc), cur = pk && db.maskTpl && Object.prototype.hasOwnProperty.call(db.maskTpl, pk) ? db.maskTpl[pk] : '';
    var auto = L.tplForProcess(proc), autoT = auto ? L.maskTemplate(auto) : null;
    var sel = h('select', { 'aria-label': '공정 「' + (proc || '') + '」 의 가릴 영역 틀', disabled: !pk },
      h('option', { value: '', selected: cur === '' }, '자동 — ' + (autoT ? autoT.name : '맞는 틀 없음(모든 검사표 기본)')),
      L.MASK_TEMPLATES.map(function (t) { return h('option', { value: t.id, selected: cur === t.id }, t.name); }),
      h('option', { value: 'none', selected: cur === 'none' }, '틀 안 씀 — 모든 검사표 기본'));
    sel.addEventListener('change', function () {
      db.maskTpl = db.maskTpl || {};
      if (sel.value === '') delete db.maskTpl[pk]; else db.maskTpl[pk] = sel.value;
      save(); render(); toast('공정 「' + proc + '」 의 틀을 바꿨습니다.' + (Array.isArray((db.maskProcs || {})[pk]) ? ' 이 공정에 직접 그린 기본 영역이 있어 그것이 먼저입니다.' : ''));
    });
    return sel;
  }
  function tplTable() {
    var procs = [];
    function add(p) { p = L.str(p); if (p && !procs.some(function (x) { return L.procKey(x) === L.procKey(p); })) procs.push(p); }
    Object.keys(db.readings).forEach(function (k) { var r = db.readings[k]; if (r && r.meta) add(r.meta.process); });
    ['가공', 'CHANNEL 평행도 검사', "ASS'Y 검사", '도장', '시트', '각장', '비용접구간'].forEach(add);
    return table(['공정', '쓰는 틀', '네모'], procs.map(function (p) {
      var eff = L.masksFor(db, '\u0000', p);
      return [p, tplSelect(p), eff.src === 'process' ? '직접 그린 공정 기본 ' + eff.masks.length + '곳' : eff.masks.length + '곳(' + eff.masks.map(function (m) { return m.label || '영역'; }).join(' · ') + ')'];
    }));
  }
  function viewerCard(f, sk, an) {
    var page = ui.page, proc = procOf(sk), eff = masksOf(sk, proc);
    var box = h('div', { class: 'viewer' + (ui.zoom === 'big' ? ' big' : '') }, h('div', { class: 'viewer-msg' }, '검사표를 그리는 중…'));
    var pagesNote = h('span', { class: 'note' });
    var avail = Math.max(300, (main.clientWidth || 900) - 70 - (window.innerWidth >= 1360 ? 536 : 0));
    var w = Math.round((ui.zoom === 'big' ? 1800 : avail) * Math.min(window.devicePixelRatio || 1, 2));
    pageCanvas(f, page, w).then(function (c) {
      box.innerHTML = '';
      var pw = h('div', { class: 'page-wrap', style: ui.zoom === 'big' ? 'width:1800px' : '' });
      pw.appendChild(c);
      pw.appendChild(maskLayer(sk, proc, eff));
      box.appendChild(pw);
      var np = +c.dataset.pages || 1;
      if (np !== L.pagesOf(f)) { f.pages = np; save(); }
      pagesNote.textContent = np > 1 ? '쪽 ' + page + ' / ' + np : '1쪽';
      if (np > 1) pagesNote.appendChild(h('span', null, ' ', h('button', { type: 'button', class: 'btn btn-sm', disabled: page <= 1, onclick: function () { ui.page--; ui.aiText = ''; ui.lowConf = []; render(); } }, '이전 쪽'), ' ',
        h('button', { type: 'button', class: 'btn btn-sm', disabled: page >= np, onclick: function () { ui.page++; ui.aiText = ''; ui.lowConf = []; render(); } }, '다음 쪽')));
    }, function (e) {
      box.innerHTML = '';
      box.appendChild(h('div', { class: 'viewer-msg' }, e.message === 'nobytes'
        ? '이 브라우저에 PDF 내용이 없습니다(백업을 다른 PC 에서 불러왔거나 브라우저 저장소가 막힘). 자료 넣기에서 같은 PDF 를 다시 넣어 주세요.'
        : 'PDF 를 열지 못했습니다: ' + e.message));
    });
    // 쪽 방향 — 90° · 180° 돌아간 스캔(2026-11-10 추가 요청). 쪽마다 저장하고 AI 에 보낼 그림도 이 방향으로
    var rot = rotOf(f, page), auto = f.rotAuto && f.rotAuto[page];
    function turn(d) { f.rot = f.rot || {}; f.rot[page] = L.normRot(rot + d); if (f.rotAuto) delete f.rotAuto[page]; save(); render(); }
    var rotBar = h('div', { class: 'btn-row rot-bar', role: 'group', 'aria-label': '쪽 방향' },
      h('span', { class: 'note' }, '방향 ' + (rot ? '시계 방향 ' + rot + '° 돌림' + (auto ? '(글자 방향으로 자동)' : '') : '그대로')),
      h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { turn(-90); } }, '왼쪽으로 90°'),
      h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { turn(90); } }, '오른쪽으로 90°'),
      h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { turn(180); } }, '180°'),
      L.pagesOf(f) > 1 ? h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { f.skip = f.skip || {}; if (f.skip[page]) delete f.skip[page]; else f.skip[page] = true; save(); render(); } }, f.skip && f.skip[page] ? '이 쪽 다시 검사표로' : '이 쪽은 검사표 아님') : null);
    // 가릴 영역 — 범위 고르기
    var scopeSel = h('select', { 'aria-label': '가릴 영역을 적용할 범위' }, ['sheet', 'process', 'all'].map(function (k) {
      return h('option', { value: k, selected: ui.maskScope === k, disabled: k === 'process' && !L.procKey(proc) }, SCOPE_LABEL[k] + (k === 'process' ? (proc ? '(' + proc + ')' : '(공정명을 먼저 적어 주세요)') : ''));
    }));
    scopeSel.addEventListener('change', function () { ui.maskScope = scopeSel.value; render(); });
    var tpl = eff.tpl ? L.maskTemplate(eff.tpl) : null;
    var effText = { sheet: '이 파일 · 이 쪽만의 설정', process: '공정 「' + proc + '」 기본(직접 그림)', all: '모든 검사표 기본',
      template: tpl ? '틀 「' + tpl.name + '」' + (eff.auto ? '(공정명 「' + proc + '」 으로 자동)' : '(설정에서 고름)') : '' }[eff.src];
    // 이 공정의 틀 — 자동(공정명) · 틀 다섯 가지 · 틀 안 씀(모든 검사표 기본). 설정 화면과 같은 값(db.maskTpl)
    var tplSel = tplSelect(proc);
    // 글자 층이 있는 쪽 — 「결재 · 검사자 · 확인」 낱말 자리로 영역 짐작(스캔은 글자가 없어 못 함)
    var guessBtn = (f.pageText || {})[page] !== false ? h('button', { type: 'button', class: 'btn btn-sm', onclick: function () {
      openDoc(f).then(function (pdf) { return pdf.getPage(page); }).then(function (pg) {
        return pg.getTextContent().then(function (tc) {
          var turn = L.normRot((pg.rotate || 0) + rotOf(f, page)), vp = pg.getViewport({ scale: 1, rotation: turn });
          var items = turn ? L.normTextItems(L.viewportTextItems(tc.items, vp.transform, vp.height), vp.height) : L.normTextItems(tc.items, vp.height);
          var ms = L.masksFromText(items, vp.width, vp.height);
          if (!ms.length) { toast('이 쪽 글자 층에서 「입안 · 심사 · 결정」 · 「검사자」 · 「확인」 낱말을 찾지 못했습니다. 틀을 쓰거나 직접 그려 주세요.', true); return; }
          db.maskSheets = db.maskSheets || {}; db.maskSheets[sk] = ms; save(); render();
          toast('글자 자리로 ' + ms.length + '곳을 짐작해 「이 파일 · 이 쪽만」 에 넣었습니다(' + ms.map(function (m) { return m.label; }).join(' · ') + '). 그림과 맞는지 확인해 주세요.');
        });
      }).catch(function (e) { toast('글자 층을 읽지 못했습니다: ' + e.message, true); });
    } }, '글자 자리로 짐작(이 쪽)') : null;
    return h('section', { class: 'card', style: 'padding:14px' },
      h('div', { class: 'btn-row', style: 'margin-bottom:10px;align-items:center' },
        h('strong', { style: 'margin-right:auto;overflow-wrap:anywhere' }, f.name, ' ', h('span', { class: 'tag' }, kindLabel(f, page))),
        pagesNote,
        h('button', { type: 'button', class: 'btn btn-sm', 'aria-pressed': ui.zoom === 'big' ? 'true' : 'false', onclick: function () { ui.zoom = ui.zoom === 'big' ? 'fit' : 'big'; render(); } }, ui.zoom === 'big' ? '폭 맞춤' : '크게 보기'),
        h('button', { type: 'button', class: 'btn btn-sm', 'aria-pressed': ui.drawMask ? 'true' : 'false', onclick: function () { ui.drawMask = !ui.drawMask; render(); } }, ui.drawMask ? '가릴 영역 고치기 끝' : '가릴 영역 고치기')),
      rotBar,
      (f.pageText || {})[page] === false && !rot ? h('p', { class: 'note', style: 'margin:0 0 8px' }, '스캔 그림은 방향을 자동으로 알 수 없습니다. 옆으로 눕거나 거꾸로 된 쪽은 위 단추로 바로 세워 주세요(이 쪽에 저장).') : null,
      f.skip && f.skip[page] ? h('p', { class: 'alert warn' }, '이 쪽은 「검사표 아님」으로 빼 두었습니다. 출하 대조 · 결과에서 셈하지 않습니다.') : null,
      ui.drawMask ? h('div', { class: 'alert info mask-help' },
        h('p', { style: 'margin:0 0 6px' }, '검사표 위를 끌어 네모를 그리면 AI 에 보낼 그림에서 그 자리를 흰색으로 덮습니다(결재 칸 · 검사자 서명). 네모의 × 로 뺍니다. 공정마다 결재 칸 자리가 다르면 「이 공정 기본」, 이 PDF 만 다르면 「이 파일 · 이 쪽만」 으로 그려 주세요.'),
        h('div', { class: 'form-grid' }, fld('적용할 범위', scopeSel), fld('이 공정의 틀', tplSel, L.procKey(proc) ? '공정명으로 자동으로 고릅니다. 어긋나면 틀을 바꾸거나, 이 쪽만 고쳐 주세요.' : '공정명을 적으면 틀을 고를 수 있습니다.')),
        guessBtn ? h('div', { class: 'btn-row', style: 'margin-top:8px' }, guessBtn, h('span', { class: 'note' }, '글자 층이 있는 PDF 만 — 스캔 그림은 틀을 씁니다.')) : null) : null,
      box,
      h('p', { class: 'note', style: 'margin:8px 0 0' }, 'AI 에 보낼 때 가리는 영역 ' + eff.masks.length + '곳 — ' + effText + ' (빨간 점선) ',
        eff.src === 'sheet' ? h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { delete db.maskSheets[sk]; save(); render(); toast('이 파일 · 쪽의 따로 맞춘 영역을 지웠습니다 — 공정(또는 모든 검사표) 기본을 씁니다.'); } }, '이 파일 조정 지우기(기본으로)') : null,
        eff.src === 'process' ? h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { if (!confirm('공정 「' + proc + '」 의 기본 가릴 영역을 지울까요? 같은 공정의 모든 검사표에 적용됩니다.')) return; delete db.maskProcs[L.procKey(proc)]; save(); render(); } }, '공정 기본 지우기') : null,
        eff.src === 'template' ? h('span', { class: 'note' }, '쪽마다 어긋나면 「가릴 영역 고치기」 → 「이 파일 · 이 쪽만」 으로 맞춰 주세요. ') : null,
        eff.src === 'all' && eff.masks.length ? h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { if (!confirm('모든 검사표 기본 가릴 영역을 지울까요?')) return; db.masks = []; save(); render(); } }, '모두 지우기') : null));
  }
  function maskLayer(sk, proc, eff) {
    var layer = h('div', { class: 'mask-layer' + (ui.drawMask ? ' drawing' : '') });
    eff.masks.forEach(function (m, i) {
      var el = h('div', { class: 'mask', style: 'left:' + m.x * 100 + '%;top:' + m.y * 100 + '%;width:' + m.w * 100 + '%;height:' + m.h * 100 + '%', title: 'AI 에 보낼 때 가림' }, h('span', null, (m.label || '가릴 영역 ' + (i + 1))));
      if (ui.drawMask) el.appendChild(h('button', { type: 'button', class: 'mask-x', 'aria-label': '이 영역 빼기', title: '빼기(' + SCOPE_LABEL[ui.maskScope] + ')', onpointerdown: function (e) { e.stopPropagation(); }, onclick: function (e) {
        e.stopPropagation();
        var arr = scopeArr(ui.maskScope, sk, proc, true);
        if (!arr) { toast('공정명을 먼저 적어 주세요.', true); return; }
        var j = arr.findIndex(function (x) { return x.x === m.x && x.y === m.y && x.w === m.w && x.h === m.h; });
        if (j < 0) { toast('이 영역은 「' + SCOPE_LABEL[ui.maskScope] + '」 설정에 없습니다. 범위를 바꿔 주세요.', true); return; }
        arr.splice(j, 1); save(); render(); toast('뺐습니다(' + SCOPE_LABEL[ui.maskScope] + ').');
      } }, '×'));
      layer.appendChild(el);
    });
    if (!ui.drawMask) return layer;
    var start = null, tmp = null;
    function rel(e) { var b = layer.getBoundingClientRect(); return { x: Math.max(0, Math.min(1, (e.clientX - b.left) / b.width)), y: Math.max(0, Math.min(1, (e.clientY - b.top) / b.height)) }; }
    layer.addEventListener('pointerdown', function (e) { e.preventDefault(); start = rel(e); tmp = h('div', { class: 'mask' }); layer.appendChild(tmp); layer.setPointerCapture(e.pointerId); });
    layer.addEventListener('pointermove', function (e) {
      if (!start) return; var p = rel(e);
      tmp.style.left = Math.min(start.x, p.x) * 100 + '%'; tmp.style.top = Math.min(start.y, p.y) * 100 + '%';
      tmp.style.width = Math.abs(p.x - start.x) * 100 + '%'; tmp.style.height = Math.abs(p.y - start.y) * 100 + '%';
    });
    layer.addEventListener('pointerup', function (e) {
      if (!start) return; var p = rel(e);
      var m = { x: L.round(Math.min(start.x, p.x), 4), y: L.round(Math.min(start.y, p.y), 4), w: L.round(Math.abs(p.x - start.x), 4), h: L.round(Math.abs(p.y - start.y), 4) };
      start = null;
      if (m.w > 0.01 && m.h > 0.005) {
        var arr = scopeArr(ui.maskScope, sk, proc, true);
        if (!arr) toast('공정명을 먼저 적어 주세요.', true);
        else { arr.push(m); save(); toast('가릴 영역을 더했습니다(' + SCOPE_LABEL[ui.maskScope] + ').'); }
      }
      render();
    });
    return layer;
  }
  /* 같은 품번(없으면 모델)의 다른 검사표에서 항목 목록을 가져옵니다 — 스캔 검사표의 칸을 미리 채우는 데 씀.
     공정명을 알면 같은 공정의 검사표를 먼저, 없으면 공정 기본 항목(시트 = 풍선 체크), 공정이 다른 검사표의 항목은 가져오지 않습니다. */
  function itemsFromOthers(meta, exceptId) {
    var kp = L.key(meta.part), models = L.splitModels(meta.model).map(L.key), pk = L.procKey(meta.process), best = null, bestP = null;
    function better(r, b0) { return !b0 || (r.done && !b0.done) || String(r.at) > String(b0.at); }
    Object.keys(db.readings).forEach(function (id) {
      if (id === exceptId) return;
      var r = db.readings[id];
      if (!r || !r.items || !r.items.length || !r.meta) return;
      var hit = (kp && L.key(r.meta.part) === kp) || (!kp && models.length && L.splitModels(r.meta.model).some(function (m) { return models.indexOf(L.key(m)) >= 0; }));
      if (!hit) return;
      var rp = L.procKey(r.meta.process);
      if (pk && rp === pk) { if (better(r, bestP)) bestP = r; }
      else if (!pk || !rp) { if (better(r, best)) best = r; }
    });
    if (bestP) return { items: clone(bestP.items), from: '같은 ' + (kp ? '품번' : '모델') + ' · 같은 공정의 다른 검사표' };
    var pre = L.procPreset(meta.process);
    if (pre) return { items: pre, from: '시트 공정 기본 항목(풍선 체크)' };
    if (best) return { items: clone(best.items), from: '같은 ' + (kp ? '품번' : '모델') + '의 다른 검사표' };
    var ov = (db.overrides || []).filter(function (o) { return kp && L.key(o.part) === kp; });
    if (ov.length) return { items: ov.map(function (o, i) { return L.cleanItem({ no: /^\d+$/.test(o.item) ? o.item : String(i + 1), label: /^\d+$/.test(o.item) ? '' : o.item, kind: o.kind }, i); }), from: '스펙 덮어쓰기 표' };
    return null;
  }
  /* 판독 초안 — 저장된 기록이 있으면 그것, 없으면 파일 이름 힌트 + 같은 품번 항목으로 */
  function draftOf(f, sk, info) {
    var rd = db.readings[sk];
    if (rd) {
      var d = clone(rd);
      d.meta = d.meta || {}; d.items = d.items || []; d.hogis = d.hogis && d.hogis.length ? d.hogis : [{ no: '', values: {} }]; d.problems = d.problems || {};
      d.src = {};
      if (info) L.META_IDS.forEach(function (k) { if (!d.meta[k] && info.meta[k] && info.src[k] === 'filename') { d.meta[k] = info.meta[k]; d.src[k] = 'filename'; } });
      if (info && !d.hogis.some(function (x) { return x.no; }) && info.hogiSrc === 'filename') { d.hogis[0].no = info.hogis[0]; d.src.hogi = 'filename'; }
      return d;
    }
    var meta = {}, src = {};
    if (info) L.META_IDS.forEach(function (k) { if (info.meta[k]) { meta[k] = info.meta[k]; src[k] = 'filename'; } });
    var hg = info && info.hogis.length ? info.hogis.map(function (x) { return { no: x, values: {} }; }) : [{ no: '', values: {} }];
    if (info && info.hogis.length) src.hogi = 'filename';
    var pre = itemsFromOthers(meta, sk);
    return { meta: meta, src: src, items: pre ? pre.items : [], itemsFrom: pre ? pre.from : '', hogis: hg, problems: {}, done: false, fresh: true };
  }
  function noValues(d) { return d.hogis.every(function (x) { return !Object.keys(x.values || {}).some(function (k) { return L.str(x.values[k]); }); }); }
  function readForm(f, sk, info, an) {
    var d = draftOf(f, sk, info);
    var rd = db.readings[sk];
    var touched = false;
    var verdict = sheetVerdict(an, sk);
    function store(done) {
      var r0 = db.readings[sk] || {};
      var by;
      if (ui.aiFor === sk) by = 'AI 읽기 + 사람 확인';
      else if (r0.auto === 'text') by = touched ? '글자 층 자동 + 사람 고침' : (r0.by || '글자 층 자동');
      else by = r0.by && r0.by !== '글자 층 자동' ? r0.by : '사람';
      db.readings[sk] = { meta: d.meta, items: d.items, hogis: d.hogis.filter(function (x, i) { return x.no || Object.keys(x.values || {}).some(function (k) { return x.values[k]; }) || i === 0; }), problems: d.problems, specPick: d.specPick || {},
        auto: r0.auto || '', warnings: r0.warnings, by: by, at: nowText(), done: done == null ? !!r0.done : done };
      save();
    }
    // 머리 칸
    var metaGrid = h('div', { class: 'meta-grid six' });
    var dl = h('datalist', { id: 'knownParts' });
    L.knownIds(db.ships, db.readings).parts.forEach(function (p) { dl.appendChild(h('option', { value: p })); });
    var dlm = h('datalist', { id: 'knownModels' });
    L.knownIds(db.ships, db.readings).models.forEach(function (p) { dlm.appendChild(h('option', { value: p })); });
    L.TEMPLATE.header.forEach(function (fd) {
      var inp = h('input', { value: d.meta[fd.id] || '', 'aria-label': fd.label, spellcheck: 'false', list: fd.id === 'part' ? 'knownParts' : fd.id === 'model' ? 'knownModels' : null, placeholder: fd.id === 'date' ? 'YYYY-MM-DD' : '' });
      inp.addEventListener('change', function () {
        var v = inp.value.trim();
        d.meta[fd.id] = fd.id === 'date' ? (L.toDate(v) || v) : v; touched = true;
        delete d.src[fd.id];
        // 품번 · 모델을 적으면 항목을 가져오고, 공정명을 바꾸면(값을 아직 안 적었을 때) 그 공정의 항목으로 바꿉니다(시트 = 풍선 체크)
        var refill = ((fd.id === 'part' || fd.id === 'model') && !d.items.length) || (fd.id === 'process' && noValues(d));
        if (refill) { var pre = itemsFromOthers(d.meta, sk); if (pre) { d.items = pre.items; toast('검사 항목을 ' + pre.from + '에서 가져왔습니다.'); } }
        store(); render();
      });
      metaGrid.appendChild(h('label', { class: 'field' + (d.src[fd.id] === 'filename' ? ' from-name' : '') }, h('span', null, fd.label, d.src[fd.id] === 'filename' ? h('small', { class: 'hint' }, ' 파일 이름에서 짐작') : null), inp));
    });
    // 체크한 모델(2026-11-10 4차 ④) — 모델명 칸에 모델이 여럿이거나 모델별 규격이 둘 이상인 항목이 있을 때. 고르면 그 모델의 규격으로 판정
    d.specPick = d.specPick || {};
    var multiItems = d.items.map(L.cleanItem).filter(function (it) { return L.modelSpecs(it.label).length; });
    var mm = L.modelMarks(d.meta.model), modelChoices = mm.models.slice();
    multiItems.forEach(function (it) { L.modelSpecs(it.label).forEach(function (x) { x.models.forEach(function (m) { if (!modelChoices.some(function (c) { return L.key(c) === L.key(m); })) modelChoices.push(m); }); }); });
    var curModel = L.unitModel(d.meta);
    var checkBox = null;
    if (mm.models.length > 1 || multiItems.length) {
      var ckSel = h('select', { 'aria-label': '체크한 모델' },
        h('option', { value: '', selected: !L.str(d.meta.checkModel) }, mm.checked.length === 1 ? '모델명 칸의 체크 — ' + mm.checked[0] : mm.models.length === 1 ? '모델명 칸 — ' + mm.models[0] : '모름(모델별 규격 항목은 보류)'),
        modelChoices.map(function (m) { return h('option', { value: m, selected: L.key(d.meta.checkModel) === L.key(m) }, m); }));
      ckSel.addEventListener('change', function () { if (ckSel.value) d.meta.checkModel = ckSel.value; else delete d.meta.checkModel; touched = true; store(); render(); });
      checkBox = h('div', { class: 'alert ' + (multiItems.length && !curModel ? 'warn' : 'info') + ' model-pick' },
        h('p', { style: 'margin:0 0 6px' }, '모델명 칸에 체크(✓ · V · ○ · ■)한 모델을 골라 주세요. ', multiItems.length ? '모델별 규격이 둘 이상인 항목 ' + multiItems.length + '개(No ' + multiItems.map(function (x) { return x.no; }).join(', ') + ')는 이 모델의 규격으로 판정하고, 모르면 「보류」 로 둡니다.' : ''),
        fld('체크한 모델', ckSel, curModel ? '지금 판정에 쓰는 모델: ' + curModel : '모델을 모릅니다'));
    }
    // 항목 × 호기 표
    var cellsList = L.itemCells(d.items.map(L.cleanItem));
    var hogiInputs = [], valInputs = [], foot = [];
    var grid = h('table', { class: 'list sheet-grid' });
    var trh = h('tr', null, h('th', { scope: 'col' }, 'No'), h('th', { scope: 'col' }, '검사 항목'), h('th', { scope: 'col' }, '구분'));
    d.hogis.forEach(function (hg, hi) {
      var inp = h('input', { value: hg.no || '', 'aria-label': '호기 ' + (hi + 1), placeholder: '호기', spellcheck: 'false', class: d.src.hogi === 'filename' && hi === 0 ? 'from-name' : null });
      inp.addEventListener('change', function () { hg.no = inp.value.trim().replace(/^#/, ''); delete d.src.hogi; touched = true; store(); render(); });
      hogiInputs.push(inp);
      trh.appendChild(h('th', { scope: 'col', class: 'hcol' }, inp, d.hogis.length > 1 ? h('button', { type: 'button', class: 'x', title: '이 호기 열 빼기', 'aria-label': '호기 열 빼기', onclick: function () { d.hogis.splice(hi, 1); touched = true; store(); render(); } }, '×') : null));
    });
    trh.appendChild(h('th', { scope: 'col' }, '문제점(조치)'));
    grid.appendChild(h('thead', null, trh));
    var tb = h('tbody');
    var lowSet = ui.aiFor === sk ? ui.lowConf : [];
    cellsList.forEach(function (c, ci) {
      var first = ci === 0 || cellsList[ci - 1].item !== c.item;
      var span = (c.item.subs.length || 1);
      var tr = h('tr', { class: first ? 'first' : null });
      if (first) {
        var sp = L.resolveSpec(c.item, d.meta.part, db.overrides, { model: curModel, pick: d.specPick[c.item.no] });
        var pickSel = null;
        if (sp.modelSpecs && sp.src !== 'override') {
          // 이 항목의 규격을 직접 고르기 — 고르지 않으면 체크한 모델로, 그것도 모르면 보류
          pickSel = h('select', { class: 'spec-pick', 'aria-label': 'No ' + c.item.no + ' 규격 고르기' },
            h('option', { value: '', selected: d.specPick[c.item.no] == null }, '체크한 모델로(' + (curModel || '모름 → 보류') + ')'),
            sp.modelSpecs.map(function (x, xi) { return h('option', { value: String(xi), selected: String(d.specPick[c.item.no]) === String(xi) }, x.label + ': ' + L.specText({ kind: 'num', lower: x.lower, upper: x.upper, unit: x.unit })); }));
          (function (no, sel) { sel.addEventListener('change', function () { if (sel.value === '') delete d.specPick[no]; else d.specPick[no] = +sel.value; touched = true; store(); render(); }); })(c.item.no, pickSel);
        }
        tr.appendChild(h('td', { rowspan: span, class: 'no' }, c.item.no));
        tr.appendChild(h('td', { rowspan: span, class: 'lab' }, h('div', null, c.item.label || '(이름 없음)'), h('small', { class: sp.src === 'none' || sp.ambiguous ? 'bad' : null }, (c.item.tool ? c.item.tool + ' · ' : '') + L.specText(sp) + (sp.src === 'override' ? ' (덮어씀)' : '')), pickSel));
      }
      tr.appendChild(h('td', { class: 'sub' }, c.sub || ''));
      d.hogis.forEach(function (hg, hi) {
        var inp = h('input', { value: (hg.values || {})[c.key] || '', inputmode: c.item.kind === 'num' ? 'decimal' : 'text', placeholder: c.item.kind === 'attach' ? 'OK · 23/24' : null, 'aria-label': (hg.no || '호기 ' + (hi + 1)) + ' · ' + c.key, autocomplete: 'off', spellcheck: 'false', 'data-k': c.key, 'data-h': hi, class: lowSet.indexOf(c.key) >= 0 ? 'low' : null });
        inp.addEventListener('input', function () { (hg.values = hg.values || {})[c.key] = inp.value; judgeAll(); });
        inp.addEventListener('change', function () { touched = true; store(); });
        inp.addEventListener('keydown', function (e) {
          if (e.key !== 'Enter') return;
          e.preventDefault();
          var col = valInputs.filter(function (x) { return x.h === hi; }), i = col.map(function (x) { return x.el; }).indexOf(inp);
          if (i < col.length - 1) col[i + 1].el.focus();
          else { var nx = valInputs.filter(function (x) { return x.h === hi + 1; })[0]; if (nx) nx.el.focus(); else saveDone(true); }
        });
        valInputs.push({ el: inp, h: hi, key: c.key });
        tr.appendChild(h('td', { class: 'v' }, inp));
      });
      if (first) {
        var pin = h('input', { value: d.problems[c.item.no] || '', 'aria-label': 'No ' + c.item.no + ' 문제점', spellcheck: 'false' });
        pin.addEventListener('change', function () { if (pin.value.trim()) d.problems[c.item.no] = pin.value.trim(); else delete d.problems[c.item.no]; touched = true; store(); });
        tr.appendChild(h('td', { rowspan: span, class: 'pr' }, pin));
      }
      tb.appendChild(tr);
    });
    grid.appendChild(tb);
    var ftr = h('tr', null, h('th', { colspan: 3, scope: 'row' }, '호기 판정'));
    d.hogis.forEach(function () { var td = h('td', { class: 'jd' }); foot.push(td); ftr.appendChild(td); });
    ftr.appendChild(h('td'));
    grid.appendChild(h('tfoot', null, ftr));
    var detail = h('div', { class: 'jd-detail note' });
    function judgeAll() {
      var lines = [];
      d.hogis.forEach(function (hg, hi) {
        var j = L.judgeUnit({ items: d.items, values: hg.values || {}, meta: d.meta, problems: d.problems, specPick: d.specPick, done: true }, db.overrides, db.settings);
        var by = {};
        j.cells.forEach(function (x) { by[x.key] = x; });
        valInputs.filter(function (x) { return x.h === hi; }).forEach(function (x) {
          var c = by[x.key], s = c ? c.status : '';
          x.el.parentNode.className = 'v ' + (s === 'pass' ? (c.near ? 'near' : c.check ? 'chk' : 'pass') : s === 'fail' ? 'fail' : s === 'unread' || s === 'nospec' || s === 'modelspec' ? 'hold' : s === 'blank' && L.isDone(rd) ? 'empty' : '');
          x.el.title = c ? (c.msg || (s === 'pass' ? '합격' : '')) : '';
        });
        foot[hi].innerHTML = '';
        foot[hi].appendChild(st(j.verdict === 'wait' ? 'hold' : j.verdict));
        if (j.fails.length) lines.push((hg.no || '호기 ' + (hi + 1)) + ' 불합격 — ' + j.fails.map(function (x) { return 'No ' + x.key + ' ' + x.raw + '(' + x.msg + ')'; }).join(', '));
        if (j.unreads.length) lines.push((hg.no || '호기 ' + (hi + 1)) + ' 확인 필요 — ' + j.unreads.map(function (x) { return 'No ' + x.key + '(' + x.msg + ')'; }).join(', '));
        if (j.blanks.length && j.verdict !== 'fail') lines.push((hg.no || '호기 ' + (hi + 1)) + ' 빈 칸 ' + j.blanks.length + '개');
      });
      detail.textContent = lines.join(' / ');
    }
    function saveDone(next) {
      if (!d.meta.part && !d.meta.model) { toast('품번이나 모델명을 적어 주세요(출하일정과 짝지을 때 씀).', true); return; }
      if (!d.hogis.some(function (x) { return x.no; })) { toast('호기 번호를 적어 주세요.', true); return; }
      d.src = {};
      store(true);
      ui.aiText = ''; ui.lowConf = []; ui.aiFor = '';
      toast('저장했습니다 — 호기 ' + d.hogis.filter(function (x) { return x.no; }).length + '대.');
      if (next) {
        // 다음 할 일 — 같은 파일의 다음 쪽부터, 그다음 다른 파일
        var an2 = analysis(), ids2 = idInfoMap(an2), order = [];
        db.files.filter(function (x) { return x.hasPdf !== false; }).forEach(function (x) { L.sheetPages(x).forEach(function (pg) { order.push([x, pg]); }); });
        var at = order.findIndex(function (o) { return o[0] === f && o[1] === ui.page; });
        order = order.slice(at + 1).concat(order.slice(0, Math.max(0, at)));
        var nx = order.filter(function (o) { var v = sheetVerdict(an2, L.sheetKey(o[0].id, o[1])), i = ids2[L.sheetKey(o[0].id, o[1])]; return v === 'wait' || (i && i.status !== 'ok'); })[0];
        if (nx) { ui.cur = nx[0].id; ui.page = nx[1]; }
      }
      render();
    }
    var notice = rd && rd.auto === 'text' ? h('p', { class: 'alert info' }, '글자 층에서 읽었습니다(' + (rd.by || '') + '). 그림과 다른 칸이 있으면 고친 뒤 「판독 완료」를 눌러 주세요.' + (rd.warnings && rd.warnings.length ? ' 알림: ' + rd.warnings.join(' ') : ''))
      : (f.pageText || {})[ui.page] === false || f.text === false ? h('p', { class: 'alert warn' }, '그림만 있는 쪽(스캔)입니다. 왼쪽 그림을 보며 머리 칸 · 호기 · 값을 적거나, 아래 「AI 로 읽기」로 채운 뒤 대조해 주세요.' + (d.itemsFrom ? ' 검사 항목은 ' + d.itemsFrom + '에서 가져왔습니다.' : ''))
      : null;
    /* 수동 판독(2026-11-10 추가 요청) — 판독 대기 · 보류인 검사표를 눌러서 바로 칸에 적기.
       항목이 없으면 같은 공정 · 같은 품번 검사표나 공정 기본(시트 = 풍선 체크)으로, 그것도 없으면 빈 줄 3개로 칸을 만들고 첫 빈 칸에 커서를 둡니다. */
    function manual() {
      if (!d.items.length) {
        var pre = itemsFromOthers(d.meta, sk);
        d.items = pre ? pre.items : [1, 2, 3].map(function (i) { return L.cleanItem({ no: String(i), label: '' }, i - 1); });
        toast(pre ? '검사 항목을 ' + pre.from + '에서 가져왔습니다. 값을 적어 주세요.' : '빈 항목 3줄을 만들었습니다. 「검사 항목 고치기」에서 이름 · 규격을 적고 값을 적어 주세요.');
      }
      if (!d.hogis.length) d.hogis = [{ no: '', values: {} }];
      touched = true; store(); ui.focus = sk; render();
    }
    function reauto() {
      if (rd && rd.done && !confirm('지금 판독값을 지우고 글자 층에서 다시 읽을까요?')) return;
      ui.busy = '글자 층을 다시 읽는 중…'; render();
      openDoc(f).then(function (pdf) { return readPage(f, pdf, ui.page, true); }).then(function (r) {
        ui.busy = ''; save(); render(); toast(r && r.ok && r.items.length ? '글자 층에서 다시 읽었습니다.' : '이 쪽은 글자 층으로 읽을 수 없습니다(스캔 그림). 수동 판독이나 AI 로 읽어 주세요.', !(r && r.ok));
      }, function (e) { ui.busy = ''; render(); toast('다시 읽지 못했습니다: ' + e.message, true); });
    }
    var cfgAi = E.load();
    var manualBar = verdict === 'wait' || verdict === 'hold' || !d.items.length ? h('div', { class: 'alert warn manual-bar' },
      h('p', { style: 'margin:0 0 8px' }, h('b', null, '이 검사표: ' + (verdict === 'hold' ? '보류' : '판독 대기')), verdict === 'hold' ? ' — 빈 칸 · 못 읽은 칸 · 모델을 골라야 하는 규격이 있습니다(색이 칠해진 칸).' : ' — 아직 값이 판정에 들어가지 않았습니다.'),
      h('div', { class: 'btn-row' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: manual }, '수동 판독(직접 입력)'),
        (f.pageText || {})[ui.page] !== false ? h('button', { type: 'button', class: 'btn', onclick: reauto }, '글자 층에서 다시 읽기') : null,
        h('button', { type: 'button', class: 'btn', onclick: function () { var el = document.querySelector('details.ai-read'); if (el) { el.open = true; el.scrollIntoView({ block: 'start' }); } } }, E.isReady(cfgAi) ? 'AI 로 읽기(설정됨)' : 'AI 로 읽기(반자동)'))) : null;
    var form = h('section', { class: 'card read-form' },
      h('h2', null, '값 옮기기' + (L.pagesOf(f) > 1 ? ' — ' + ui.page + '쪽' : '')),
      manualBar,
      notice,
      info && info.status === 'hint' ? h('p', { class: 'alert warn' }, '노란 칸은 파일 이름에서 짐작한 것입니다. 검사표 머리 칸 · 호기와 같은지 확인해 주세요.') : null,
      metaGrid, dl, dlm, checkBox,
      d.items.length ? h('div', { class: 'table-wrap', style: 'margin-top:12px' }, grid) : h('p', { class: 'alert warn' }, '검사 항목이 아직 없습니다. 품번(또는 모델명)을 적으면 같은 품번의 다른 검사표에서 항목을 가져옵니다. 처음 보는 양식이면 아래 「검사 항목 고치기」로 넣거나 AI 로 읽어 주세요.'),
      d.items.length ? detail : null,
      d.items.length ? h('p', { class: 'note', style: 'margin:8px 0' }, '값은 적힌 그대로(단위 · 쉼표도 그대로). 빈 칸은 비워 두고(보류), 알아볼 수 없으면 「?」(확인 필요), 체크만 있으면 「V」, 육안은 OK / NG. Enter 로 아래 칸.') : null,
      h('div', { class: 'btn-row' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { saveDone(true); } }, '판독 완료 · 다음'),
        h('button', { type: 'button', class: 'btn', onclick: function () { saveDone(false); } }, '판독 완료'),
        h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { d.hogis.push({ no: '', values: {} }); touched = true; store(); render(); } }, '호기 열 더하기'),
        L.isDone(rd) ? h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { db.readings[sk].done = false; save(); render(); toast('다시 「판독 대기」로 돌렸습니다.'); } }, '다시 대기로') : null),
      L.isDone(rd) ? h('p', { class: 'note', style: 'margin:8px 0 0' }, '판독: ' + (rd.by || '') + ' · ' + (rd.at || '')) : null,
      itemEditor(sk, d, function () { touched = true; store(); render(); }),
      sheetAiPanel(f, sk, d, store));
    if (d.items.length) setTimeout(judgeAll, 0);
    if (ui.focus === sk) setTimeout(function () {
      ui.focus = '';
      var hin = hogiInputs.filter(function (x) { return !x.value; })[0];
      var empty = valInputs.filter(function (x) { return !x.el.value || /hold|empty/.test(x.el.parentNode.className); })[0];
      var el = hin || (empty ? empty.el : valInputs[0] && valInputs[0].el);
      if (el) { el.scrollIntoView({ block: 'center' }); el.focus(); }
    }, 30);
    return form;
  }
  function itemEditor(sk, d, changed) {
    var t = h('table', { class: 'list spec' });
    t.appendChild(h('thead', null, h('tr', null, ['No', '검사 항목(규격 글자)', '측정기', '구분', '방식', ''].map(function (x) { return h('th', null, x); }))));
    var tb = h('tbody');
    d.items.forEach(function (it, i) {
      function inp(k, v, w) { var e = h('input', { value: v, 'aria-label': k, style: w ? 'width:' + w : null }); e.addEventListener('change', function () { it[k] = k === 'subs' ? e.value.split(/[,\/\s]+/).filter(Boolean).map(function (s) { return s.toUpperCase(); }) : e.value.trim(); if (k === 'label' || k === 'tool') delete it.kind; d.items[i] = L.cleanItem(it, i); changed(); }); return e; }
      var kind = h('select', { 'aria-label': '방식' }, [['num', '수치'], ['okng', '육안 OK/NG'], ['attach', '부착 확인(풍선 체크)']].map(function (o) { return h('option', { value: o[0], selected: L.cleanItem(it, i).kind === o[0] }, o[1]); }));
      kind.addEventListener('change', function () { it.kind = kind.value; changed(); });
      tb.appendChild(h('tr', null, h('td', null, inp('no', it.no, '52px')), h('td', null, inp('label', it.label || '')), h('td', null, inp('tool', it.tool || '')), h('td', null, inp('subs', (it.subs || []).join(','), '80px')), h('td', null, kind),
        h('td', null, h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { d.items.splice(i, 1); changed(); } }, '빼기'))));
    });
    t.appendChild(tb);
    var pre = itemsFromOthers(d.meta, sk);
    return h('details', { class: 'block', style: 'margin-top:14px' },
      h('summary', null, '검사 항목 고치기 (' + d.items.length + '개)'),
      h('p', { class: 'note' }, '검사 항목 칸의 「420±1」 「48 +1.0/-0.5」 「1.5 MAX」 같은 글자에서 하한 · 상한을 읽습니다. 구분은 「LH,RH」 처럼 쉼표로. 고치면 이 검사표에만 적용됩니다(⑤ 덮어쓰기 표는 품번 전체).'),
      d.items.length ? h('div', { class: 'table-wrap' }, t) : null,
      h('div', { class: 'btn-row', style: 'margin-top:8px' },
        h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { d.items.push({ no: String(d.items.length + 1), label: '', tool: '', subs: [] }); changed(); } }, '줄 더하기'),
        pre ? h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { if (d.items.length && !confirm('지금 항목을 ' + pre.from + '의 항목으로 바꿀까요?')) return; d.items = pre.items; changed(); } }, pre.from + '에서 가져오기') : null));
  }
  function sheetAiPanel(f, sk, d, store) {
    var prompt = L.buildSheetPrompt(d.items);
    var ans = h('textarea', { rows: 6, placeholder: 'AI 의 답(JSON)을 여기에 붙여 넣어 주세요.', 'aria-label': 'AI 답' });
    ans.value = ui.aiFor === sk ? ui.aiText || '' : '';
    function apply() {
      var p = L.parseSheetAi(ans.value, d.items);
      ui.aiText = ans.value; ui.aiFor = sk;
      if (!p.ok) { toast(p.error, true); return; }
      Object.keys(p.meta).forEach(function (k) { d.meta[k] = p.meta[k]; });
      // 답에 항목이 있고, 지금 칸과 맞지 않으면(다른 공정 양식) 답의 항목으로 바꿈
      if (p.items && p.items.length && (!d.items.length || p.unknown.length)) { d.items = p.items; p = L.parseSheetAi(ans.value, d.items); }
      if (!d.items.length) { var pre = itemsFromOthers(d.meta, sk); if (pre) d.items = pre.items; }
      if (p.hogis.length) {
        var keep = d.hogis.filter(function (x) { return x.no && !p.hogis.some(function (y) { return L.key(y.no) === L.key(x.no); }); });
        d.hogis = p.hogis.map(function (y) {
          var old = d.hogis.filter(function (x) { return x.no && L.key(x.no) === L.key(y.no); })[0];
          return { no: y.no, values: Object.assign({}, old ? old.values : {}, y.values) };
        }).concat(keep.filter(function (x) { return Object.keys(x.values || {}).length; }));
      }
      Object.keys(p.problems).forEach(function (k) { d.problems[k] = p.problems[k]; });
      ui.lowConf = p.lowConf;
      d.src = {};
      store(false); render();
      toast('AI 가 읽은 값을 칸에 넣었습니다. 노란 칸(확신 낮음)부터 검사표와 대조한 뒤 「판독 완료」를 눌러 주세요.' + (p.unknown.length ? ' 모르는 칸: ' + p.unknown.join(', ') : ''));
    }
    function saveImage() {
      maskedImage(f, ui.page, d.meta.process).then(function (c) {
        c.toBlob(function (b) { download(f.name.replace(/\.pdf$/i, '') + (L.pagesOf(f) > 1 ? '_' + ui.page + '쪽' : '') + '_가림.jpg', b); toast('그림을 받았습니다(가릴 영역은 흰색). AI 대화창에 이 그림을 첨부해 주세요.'); }, 'image/jpeg', 0.88);
      }, function (e) { toast('그림을 만들지 못했습니다: ' + e.message, true); });
    }
    function copyImage() {
      if (!(navigator.clipboard && window.ClipboardItem && window.isSecureContext)) { toast('이 브라우저에서는 그림 복사가 안 됩니다. 「그림 받기」를 써 주세요.', true); return; }
      maskedImage(f, ui.page, d.meta.process).then(function (c) {
        c.toBlob(function (b) { navigator.clipboard.write([new window.ClipboardItem({ 'image/png': b })]).then(function () { toast('그림을 복사했습니다. AI 대화창에 붙여 넣어 주세요.'); }, function () { toast('그림을 복사하지 못했습니다. 「그림 받기」를 써 주세요.', true); }); }, 'image/png');
      });
    }
    var cfg = E.load(), ready = E.isReady(cfg);
    var sendBtn = h('button', { type: 'button', class: 'btn', disabled: !ready, title: ready ? E.describe(cfg) : 'AI 연결 설정이 필요합니다' }, '그림과 함께 AI 서버로 보내기(선택)');
    sendBtn.addEventListener('click', function () {
      var c = E.load(), v = E.validateConfig(c, location.protocol);
      if (v.errors.length) { toast(v.errors[0], true); return; }
      sendBtn.disabled = true; sendBtn.textContent = '보내는 중…';
      maskedImage(f, ui.page, d.meta.process).then(function (cv) {
        var msgs = [{ role: 'system', content: '너는 손으로 쓴 제품 검사표의 값을 옮겨 적는 보조야. 그림에 없는 값을 만들지 마.' },
          { role: 'user', content: [{ type: 'text', text: prompt }, { type: 'image_url', image_url: { url: cv.toDataURL('image/jpeg', 0.88) } }] }];
        return E.callChat(c, msgs, { maxTokens: 2000 });
      }).then(function (text) { ans.value = text; apply(); }, function (e) { toast(e.message, true); sendBtn.disabled = false; sendBtn.textContent = '그림과 함께 AI 서버로 보내기(선택)'; });
    });
    var truth = f.sample && SM.truth[f.name + (ui.page > 1 ? '#' + ui.page : '')];
    var demo = truth && (f.pageText || {})[ui.page] === false ? h('button', { type: 'button', class: 'btn btn-sm', onclick: function () {
      var m = truth.meta, low = [];
      var hogis = truth.hogis.map(function (x) { return { 호기: x.no, 값: x.values }; });
      var firstNum = L.itemCells(truth.items.map(L.cleanItem)).filter(function (c) { return c.item.kind === 'num'; })[0];
      if (firstNum) low.push(firstNum.key);
      var o = { 모델명: m.model.split(' / ')[0], 품번: m.part, 품명: m.name, 관리번호: m.mgmt, 작성일: m.date, 제작공정명: m.process };
      o['항목'] = truth.items.map(function (it) { return { No: it.no, 검사항목: it.label, 측정기: it.tool, 구분: it.subs }; });
      o['호기'] = hogis; o['문제점'] = truth.problems; o['확신낮음'] = low;
      ans.value = '```json\n' + JSON.stringify(o, null, 2) + '\n```';
      ui.aiText = ans.value; ui.aiFor = sk; toast('가상 「AI 답」 예시를 넣었습니다(샘플 전용). 「답 읽어 칸 채우기」를 눌러 보세요.');
    } }, '가상 AI 답 예시 넣기(샘플 전용)') : null;
    return h('details', { class: 'block ai-read', style: 'margin-top:10px', open: !!(ui.aiFor === sk && ui.aiText) || (f.pageText || {})[ui.page] === false },
      h('summary', null, 'AI 로 읽기 (반자동 · 선택)'),
      h('ol', { class: 'steps' },
        h('li', null, '검사표 그림을 받거나 복사합니다. 가릴 영역(결재 칸 · 검사자 서명)은 흰색으로 덮이고, 돌려 둔 방향 그대로 나갑니다.'),
        h('li', null, '회사가 허용한 AI 대화창에 그림을 첨부하고 아래 프롬프트를 붙여 넣습니다. 이 도구는 기본으로 아무 데도 보내지 않습니다.'),
        h('li', null, 'AI 의 답(JSON)을 붙여 넣고 「답 읽어 칸 채우기」 → 검사표와 대조해 고친 뒤 「판독 완료」.')),
      h('div', { class: 'btn-row' },
        h('button', { type: 'button', class: 'btn', onclick: saveImage }, '그림 받기(JPG)'),
        h('button', { type: 'button', class: 'btn', onclick: copyImage }, '그림 복사'),
        h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { copyText(prompt); } }, '프롬프트 복사')),
      h('details', { style: 'margin:8px 0' }, h('summary', null, '프롬프트 보기'), h('pre', { class: 'mono prompt' }, prompt)),
      ans,
      h('div', { class: 'btn-row', style: 'margin-top:8px' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: apply }, '답 읽어 칸 채우기'),
        h('span', { class: 'ai-send' }, sendBtn, ready ? null : h('small', { class: 'note', style: 'margin-left:6px' }, '자동 보내기는 ', h('a', { href: '#/settings' }, 'AI 연결 설정'), ' 후에(그림 입력을 받는 모델).')),
        demo),
      h('p', { class: 'note', style: 'margin:8px 0 0' }, '외부 AI 에 검사표 그림을 보내도 되는지는 회사 보안 기준을 먼저 확인해 주세요. AI 가 읽은 값은 그대로 믿지 말고 반드시 검사표와 대조합니다.'));
  }

  /* ── ③ 검사 결과 ─────────────────────────── */
  function pageResult() {
    var head0 = pageHead('③', '검사 결과 — 불합격 리스트 · 호기별 결과');
    if (!db.files.length) return [head0, emptyNotice()];
    var an = analysis(), c = an.counts, parts = L.byPart(an), items = L.byItem(an);
    var units = an.units.slice().sort(function (a, b) { var o = { fail: 0, hold: 1, wait: 2, pass: 3 }; return (o[a.judge.verdict] - o[b.judge.verdict]) || ((a.meta.date || '') < (b.meta.date || '') ? -1 : (a.meta.date || '') > (b.meta.date || '') ? 1 : 0); });
    var shipOf = {};
    an.rows.forEach(function (r) { r.units.forEach(function (u) { shipOf[u.id] = r.ship; }); });
    var judged = an.units.filter(function (u) { return u.judge.verdict === 'pass' || u.judge.verdict === 'fail'; }).length;
    return [
      pageHead('③', '검사 결과 — 불합격 리스트 · 호기별 결과', '검사표의 호기마다 판정합니다. 한 칸이라도 공차 밖이면 불합격, 공차와 같은 값은 합격(설정에서 바꿈).',
        h('div', { class: 'btn-row' }, h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { exportResult(an); } }, '엑셀로 내려받기'))),
      h('div', { class: 'tiles' },
        tile('불합격 호기', c.unitFail + '대', '판정 ' + judged + '대 중 ' + pct(judged ? c.unitFail / judged : null), c.unitFail ? 'over' : 'ok'),
        tile('호기 기록', c.units + '대', '검사표 ' + c.files + '장'),
        tile('보류 · 판독 대기', an.units.filter(function (u) { return u.judge.verdict === 'hold'; }).length + ' · ' + an.units.filter(function (u) { return u.judge.verdict === 'wait'; }).length, '빈 칸 · 못 읽은 칸 · 아직 안 옮김'),
        tile('식별 필요 파일', (c.unknown + c.hint) + '장', '모델/품번 · 호기를 못 찾음', c.unknown ? 'over' : null)),
      h('section', { class: 'card' }, h('h2', null, '불합격 리스트 (' + c.unitFail + '대 · ' + an.fails.length + '칸)'),
        an.fails.length ? table(['출하일', '남은 날', '모델', '품번', '공정', '호기', 'No', '검사 항목', '기준', '적힌 값', '내용', '문제점', '검사표'], an.fails.map(function (x) {
          return { cls: 'row-fail', cells: [x.shipDate || '(출하일정 없음)', ddayCell(x.days), x.model, x.part, x.process, h('b', null, x.hogi), x.no + (x.sub ? ' ' + x.sub : ''), x.label, x.spec, x.raw, x.msg, x.problem, x.file] };
        }), { nw: [0, 1, 5, 6, 8] }) : h('p', { class: 'note' }, '판정한 호기 가운데 불합격이 없습니다.')),
      h('section', { class: 'card' }, h('h2', null, '호기별 결과 (' + an.units.length + ')'),
        units.length ? table(['작성일', '모델', '품번', '공정', '호기', '판정', '벗어난 칸', '빈 칸 · 확인 필요', '체크만', '출하일', '검사표', '판독'], units.map(function (u) {
          var j = u.judge, s = shipOf[u.id];
          return { cls: j.verdict === 'fail' ? 'row-fail' : null, cells: [u.meta.date || '-', u.meta.model || '', u.meta.part || '', u.meta.process || '', h('b', null, u.hogi), st(j.verdict),
            failText(j), [j.blanks.length ? '빈 칸 ' + j.blanks.length : '', j.unreads.length ? '확인 ' + j.unreads.length : ''].filter(Boolean).join(' · '), j.checks.length || '',
            s ? s.date : '-', sheetName(u.file, u.page), u.done ? u.by : readLink(u.fileId, u.page, '수동 판독', true)] };
        }), { nw: [0, 4, 9] }) : h('p', { class: 'note' }, '호기를 찾은 검사표가 없습니다.')),
      h('div', { class: 'grid-2' },
        h('section', { class: 'card' }, h('h2', null, '품번별'),
          table(['품번(모델)', '호기', '판정', '불합격', '불합격률', '보류', '대기'], parts.map(function (p) { return [p.part, p.units, p.judged, p.fail, pct(p.failRate), p.hold, p.wait]; }), { num: [0, 1, 1, 1, 1, 1, 1] })),
        h('section', { class: 'card' }, h('h2', null, '칸별 (불합격 · 경계 근처 많은 순)'),
          items.length ? table(['품번', '칸', '검사 항목', '기준', '판정', '불합격', '하한 밑', '상한 위', '경계 근처'], items.slice(0, 40).map(function (x) { return [x.part, x.key, x.label, x.spec, x.n, x.fail, x.low, x.high, x.near]; }), { num: [0, 0, 0, 0, 1, 1, 1, 1, 1] }) : h('p', { class: 'note' }, '판독한 값이 없습니다.'),
          h('p', { class: 'note', style: 'margin-top:8px' }, '경계 근처 = 합격이지만 공차 폭의 ' + settings().nearPct + '% 안쪽(설정에서 바꿈).')))
    ];
  }
  function exportResult(an) {
    var wb = XLSX.utils.book_new();
    function sheet(name, head, rows) { XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([head].concat(rows)), name); }
    var shipOf = {};
    an.rows.forEach(function (r) { r.units.forEach(function (u) { shipOf[u.id] = r.ship; }); });
    sheet('불합격_리스트', ['출하일', '남은 날', '모델', '품번', '품명', '호기', 'No', '구분', '검사 항목', '기준', '적힌 값', '내용', '문제점(조치)', '작성일', '검사표', '공정'],
      an.fails.map(function (x) { return [x.shipDate, L.dday(x.days), x.model, x.part, x.name, x.hogi, x.no, x.sub, x.label, x.spec, x.raw, x.msg, x.problem, x.sheetDate, x.file, x.process]; }));
    sheet('검사_결과', ['작성일', '모델', '품번', '품명', '관리 번호', '호기', '판정', '벗어난 칸', '빈 칸', '확인 필요', '체크만', '출하일', '검사표', '판독', '판독 시각', '공정'], an.units.map(function (u) {
      var j = u.judge, s = shipOf[u.id];
      return [u.meta.date || '', u.meta.model || '', u.meta.part || '', u.meta.name || '', u.meta.mgmt || '', u.hogi, L.VERDICT_LABEL[j.verdict], failText(j), j.blanks.map(function (x) { return x.key; }).join(', '), j.unreads.map(function (x) { return x.key; }).join(', '), j.checks.length, s ? s.date : '', sheetName(u.file, u.page), u.done ? u.by : '', u.done ? u.at : '', u.meta.process || ''];
    }));
    var long = [];
    an.units.forEach(function (u) { u.judge.cells.forEach(function (c) { long.push([u.meta.part || '', u.meta.model || '', u.hogi, c.no, c.sub, c.label, c.spec, c.raw, c.value == null ? '' : c.value, ({ pass: '합격', fail: '불합격', blank: '빈 칸', unread: '확인 필요', nospec: '규격 없음', modelspec: '모델별 규격 — 고르기 필요' })[c.status], c.msg, sheetName(u.file, u.page)]); }); });
    sheet('측정값_전체', ['품번', '모델', '호기', 'No', '구분', '검사 항목', '기준', '적힌 값', '읽은 값', '판정', '내용', '검사표'], long);
    var ship = function (statuses) { return L.shipList(an, statuses).map(function (r) { return [r.ship.date, L.dday(r.days), r.ship.model, r.ship.part, r.ship.hogi, r.ship.qty, L.VERDICT_LABEL[r.status], r.unit ? failText(r.unit.judge) || holdText(r) : '', r.notes.join(' · '), r.unit ? sheetName(r.unit.file, r.unit.page) : '']; }); };
    var sh = ['출하일', '남은 날', '모델', '품번', '호기', '수량', '상태', '내용', '알림', '검사표'];
    sheet('출하_대조', sh, ship(['none', 'fail', 'wait', 'hold', 'pass']));
    sheet('검사표_없음', sh, ship(['none']));
    sheet('미판정', sh, ship(['wait', 'hold']));
    sheet('식별_필요', ['파일', '상태', '짐작한 모델', '짐작한 품번', '짐작한 호기', '모자란 것'], an.ids.filter(function (i) { return i.status !== 'ok'; }).map(function (i) { return [sheetName(i.file, i.page), ID_LABEL[i.status], i.meta.model || '', i.meta.part || '', i.hogis.join(', '), i.missing.join(', ')]; }));
    sheet('출하일정에_없는_검사', ['작성일', '모델', '품번', '호기', '판정', '검사표'], an.extra.map(function (u) { return [u.meta.date || '', u.meta.model || '', u.meta.part || '', u.hogi, L.VERDICT_LABEL[u.judge.verdict], sheetName(u.file, u.page)]; }));
    XLSX.writeFile(wb, '검사결과_' + stamp() + '.xlsx');
  }

  /* ── ④ AI 요약 ───────────────────────────── */
  function pageAi() {
    var head0 = pageHead('④', 'AI 요약 프롬프트');
    if (!db.files.length && !db.ships.length) return [head0, emptyNotice()];
    var s = settings(), an = analysis();
    var built = L.buildSummaryPrompt(an, { mask: s.maskAlias });
    var mask = h('input', { type: 'checkbox' }); mask.checked = s.maskAlias;
    mask.addEventListener('change', function () { db.settings.maskAlias = mask.checked; save(); render(); });
    var answer = h('textarea', { rows: 10, 'aria-label': 'AI 답' }); answer.value = db.aiNote || '';
    answer.addEventListener('change', function () { db.aiNote = L.unmask(answer.value, built.map); save(); });
    var mapRows = Object.keys(built.map).map(function (k) { return [built.map[k], k]; });
    return [
      pageHead('④', 'AI 요약 프롬프트', '계산된 숫자만 담은 프롬프트입니다. 회사가 허용한 AI 에 붙여 넣어 보고용 요약과 먼저 처리할 순서를 받습니다.'),
      h('section', { class: 'card' },
        h('label', { class: 'opt' }, mask, h('span', null, h('span', { class: 't' }, '품번 · 모델을 「제품1」 「모델1」 같은 별칭으로 가리기(권장)'), h('span', { class: 's' }, '대응표는 이 PC 에만 있고, 답의 별칭은 원래 이름으로 되돌려 저장합니다.'))),
        mapRows.length && s.maskAlias ? h('details', { style: 'margin-top:8px' }, h('summary', null, '별칭 대응표 ' + mapRows.length + '개'), table(['별칭', '원래 이름'], mapRows)) : null,
        h('pre', { class: 'mono prompt', style: 'margin-top:12px;max-height:420px;overflow:auto' }, built.text),
        h('div', { class: 'btn-row', style: 'margin:10px 0' },
          h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { copyText(built.text); } }, '프롬프트 복사'),
          AIP.sendButton('설정한 AI 서버로 보내기(선택)', function () { return built.text; }, function (text) { text = L.unmask(text, built.map); answer.value = text; db.aiNote = text; save(); toast('AI 답을 받았습니다(별칭은 원래 이름으로 되돌림).'); },
            { toast: toast, settingsHref: '#/settings', system: '너는 제조 품질 담당자를 돕는 보조야. 주어진 사실 밖의 숫자를 만들지 마.' })),
        fld('AI 답 (붙여 넣으면 이 브라우저에 저장)', answer))
    ];
  }

  /* ── ⑤ 스펙 ──────────────────────────────── */
  function pageSpec() {
    // 품번마다 가장 나중에 판독한 검사표의 항목
    var latest = {};
    Object.keys(db.readings).forEach(function (id) {
      var r = db.readings[id]; if (!r || !r.items || !r.items.length || !r.meta) return;
      var k = r.meta.part || r.meta.model; if (!k) return;
      if (r.meta.process) k += ' · ' + r.meta.process;      // 공정마다 양식이 달라(제작 · 시트 · 도장 …) 품번 + 공정으로 따로
      if (!latest[k] || String(r.at) > String(latest[k].at)) latest[k] = r;
    });
    var parts = Object.keys(latest).sort();
    var rows = db.overrides;
    function upd(i, k, v) {
      var s0 = Object.assign({}, rows[i]); s0[k] = v;
      var c = L.cleanSpec({ part: s0.part, item: s0.item, lower: s0.lower, upper: s0.upper, unit: s0.unit, kind: s0.kind === 'okng' ? 'OK/NG' : s0.kind === 'attach' ? '부착' : '수치' });
      if (c.error && !/바꿔 넣/.test(c.error)) toast(c.error, true);
      delete c.error; rows[i] = c; save(); render();
    }
    var t = h('table', { class: 'list spec' });
    t.appendChild(h('thead', null, h('tr', null, ['품번', '검사 항목(또는 No)', '하한', '상한', '단위', '방식', ''].map(function (x) { return h('th', null, x); }))));
    var tb = h('tbody');
    rows.forEach(function (s0, i) {
      function inp(k, num) { var e = h('input', { value: s0[k] == null ? '' : s0[k], inputmode: num ? 'decimal' : null, 'aria-label': k }); e.addEventListener('change', function () { upd(i, k, e.value); }); return e; }
      var kind = h('select', { 'aria-label': '방식' }, [['num', '수치'], ['okng', '육안 OK/NG'], ['attach', '부착 확인(풍선)']].map(function (o) { return h('option', { value: o[0], selected: s0.kind === o[0] }, o[1]); }));
      kind.addEventListener('change', function () { upd(i, 'kind', kind.value); });
      tb.appendChild(h('tr', null, h('td', null, inp('part')), h('td', null, inp('item')), h('td', null, s0.kind !== 'num' ? '-' : inp('lower', true)), h('td', null, s0.kind !== 'num' ? '-' : inp('upper', true)), h('td', null, inp('unit')), h('td', null, kind),
        h('td', null, h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { rows.splice(i, 1); save(); render(); } }, '빼기'))));
    });
    t.appendChild(tb);
    return [
      pageHead('⑤', '스펙 — 검사표에서 읽은 규격 · 덮어쓰기', '규격은 검사표 「검사 항목」 칸의 글자(420±1 · 48 +1.0/-0.5 · 1.5 MAX · MIN 40 · 8.00~8.10)에서 읽습니다. 다른 기준을 쓰려면 아래 덮어쓰기 표에 품번 · 항목(또는 No) · 하한 · 상한을 넣습니다.',
        h('div', { class: 'btn-row' },
          h('button', { type: 'button', class: 'btn', onclick: function () { rows.push({ part: parts[0] ? (latest[parts[0]].meta.part || latest[parts[0]].meta.model) : '', item: '', lower: null, upper: null, unit: '', kind: 'num' }); save(); render(); } }, '덮어쓰기 줄 더하기'),
          h('button', { type: 'button', class: 'btn', disabled: !rows.length, onclick: function () {
            var wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['품번', '검사 항목', '하한', '상한', '단위', '판정 방식']].concat(rows.map(function (s0) { return [s0.part, s0.item, s0.lower, s0.upper, s0.unit, s0.kind === 'okng' ? 'OK/NG' : s0.kind === 'attach' ? '부착' : '수치']; }))), '스펙덮어쓰기');
            XLSX.writeFile(wb, '스펙덮어쓰기_' + stamp() + '.xlsx');
          } }, '덮어쓰기 표 엑셀로'))),
      parts.length ? parts.map(function (p) {
        var r = latest[p];
        return h('section', { class: 'card' }, h('h2', null, p + (r.meta.model && r.meta.part ? ' · ' + r.meta.model : '')),
          table(['No', '검사 항목', '측정기', '구분', '읽은 규격', '쓰는 기준', ''], r.items.map(L.cleanItem).map(function (it) {
            var own = L.resolveSpec(it, r.meta.part, []), use = L.resolveSpec(it, r.meta.part, db.overrides);
            return [it.no, it.label, it.tool, it.subs.join(', '), own.src === 'none' ? h('span', { class: 'st hold' }, '못 읽음') : L.specText(own), use.src === 'override' ? h('b', null, L.specText(use) + ' (덮어씀)') : L.specText(use),
              h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { rows.push({ part: r.meta.part || r.meta.model, item: it.label || it.no, lower: own.lower, upper: own.upper, unit: own.unit || '', kind: own.kind }); save(); render(); toast('덮어쓰기 표에 넣었습니다 — 하한 · 상한을 고쳐 주세요.'); } }, '덮어쓰기')];
          }), { nw: [0, 3] }));
      }) : h('section', { class: 'card' }, h('p', { class: 'note' }, '아직 항목을 읽은 검사표가 없습니다.')),
      h('section', { class: 'card' }, h('h2', null, '덮어쓰기 표 (' + rows.length + '줄)' + (db.overrideFile ? ' — ' + db.overrideFile : '')),
        rows.length ? h('div', { class: 'table-wrap' }, t) : h('p', { class: 'note' }, '없음 — 검사표의 규격 글자를 그대로 씁니다.')),
      h('section', { class: 'card' }, h('h2', null, '판정 규칙 (설정에서 바꿈)'),
        h('ul', null,
          h('li', null, '호기마다, 한 칸이라도 하한 · 상한을 벗어나면 그 호기는 불합격입니다(빈 칸이 함께 있어도).'),
          h('li', null, '하한 · 상한과 같은 값은 ' + (settings().inclusive ? '합격' : '불합격') + '으로 봅니다(420±1 이면 419 · 421 은 ' + (settings().inclusive ? '합격' : '불합격') + ').'),
          h('li', null, '육안 항목은 OK / NG(○ · × 도 읽음). NG 는 불합격.'),
          h('li', null, '빈 칸이 있으면 ' + (settings().blankAs === 'hold' ? '보류(미판정)' : '적힌 칸만으로 판정') + ', 못 읽은 칸(「?」) · 규격을 못 읽은 항목은 보류입니다.'),
          h('li', null, '측정값 칸에는 실측값 또는 OK 를 적습니다. 수치 항목에 「OK」 가 적혀 있으면 검사자가 공차 안을 확인한 것으로 보고 합격, 「NG」 는 불합격입니다.'),
          h('li', null, '값 없이 체크(V · ✓)만 한 수치 칸은 ' + (settings().checkAs === 'pass' ? '확인한 것으로 보고 합격' : '값이 없어 보류') + '입니다.'),
          h('li', null, '시트 공정(부착 확인)은 모델링 뷰의 풍선 체크로 판정합니다 — 「OK」(모두 체크) 또는 「체크 수/풍선 수」(23/24 면 1개 빠져 불합격), 「누락」 · 「NG」 는 불합격.'),
          h('li', null, '같은 호기의 공정별 검사표(제작 · 시트 · 도장 …)는 모두 봅니다. 하나라도 불합격이면 그 호기는 불합격, 하나라도 판독 전이면 미판정입니다. 같은 공정이 둘이면 나중에 판독한 것을 씁니다.')))
    ];
  }

  /* ── 설정 ───────────────────────────────── */
  function pageSettings() {
    var s = settings();
    var inc = h('select', { 'aria-label': '경계값' }, [['1', '같으면 합격(포함)'], ['0', '같으면 불합격(미포함)']].map(function (o) { return h('option', { value: o[0], selected: (s.inclusive ? '1' : '0') === o[0] }, o[1]); }));
    var blank = h('select', { 'aria-label': '빈 칸' }, [['hold', '보류 — 미판정으로 셈'], ['ignore', '무시 — 적힌 칸만으로 판정']].map(function (o) { return h('option', { value: o[0], selected: s.blankAs === o[0] }, o[1]); }));
    var chk = h('select', { 'aria-label': '체크만' }, [['pass', '합격 — 확인한 것으로 봄'], ['hold', '보류 — 값이 없으니 미판정']].map(function (o) { return h('option', { value: o[0], selected: s.checkAs === o[0] }, o[1]); }));
    var near = h('input', { type: 'number', min: 0, max: 50, step: 1, value: s.nearPct });
    var backup = h('input', { type: 'file', accept: '.json,application/json' });
    backup.addEventListener('change', function () {
      var f = backup.files[0]; if (!f) return;
      f.text().then(function (t) {
        var o = JSON.parse(t); if (!o || o.v !== 2) throw new Error('이 도구(2차)의 백업 파일이 아닙니다.');
        db = L.restoreDb(o); save(); toast('백업을 불러왔습니다. PDF 내용은 백업에 없으니 그림을 보려면 PDF 를 다시 넣어 주세요.'); render();
      }).catch(function (e) { toast('불러오지 못했습니다: ' + e.message, true); });
    });
    return [
      pageHead('설정', '판정 규칙 · 가릴 영역 · AI 연결 · 백업'),
      h('section', { class: 'card' }, h('h2', null, '판정 규칙'),
        h('div', { class: 'form-grid' }, fld('하한 · 상한과 같은 값', inc, '10장 ④: 공차 안이면 합격'), fld('빈 칸이 있을 때', blank, '불합격 칸이 있으면 빈 칸과 관계없이 불합격'),
          fld('값 없이 체크(V · ✓)만 한 수치 칸', chk, '양식이 체크만 하는 곳이면 「합격」'), fld('「경계 근처」 — 공차 폭의 몇 % 안', near, '0 이면 표시하지 않음')),
        h('div', { class: 'btn-row', style: 'margin-top:12px' }, h('button', { type: 'button', class: 'btn btn-primary', onclick: function () {
          db.settings = L.cleanSettings(Object.assign({}, db.settings, { inclusive: inc.value === '1', blankAs: blank.value, checkAs: chk.value, nearPct: near.value })); save(); render(); toast('저장했습니다.');
        } }, '저장'))),
      h('section', { class: 'card' }, h('h2', null, 'AI 에 보낼 그림에서 가리는 영역'),
        h('p', { class: 'note' }, '② 판독 화면의 「가릴 영역 고치기」로 정합니다. 쪽 크기 대비 비율입니다. 세 겹으로 두고 좁은 것이 먼저입니다 — 이 파일 · 쪽만 > 공정 기본 > 모든 검사표 기본. 공정마다 결재 칸 자리가 다르면 공정 기본을, 한 PDF 만 다르면 그 파일에서 따로 맞춰 주세요.'),
        h('h3', { style: 'margin-top:12px' }, '공정별 틀'),
        h('p', { class: 'note' }, '공정명으로 틀을 자동으로 고릅니다 — 평행도 · ASSY · 가공 · 제작 · 도장은 ① 제작 · 가공형(결재 · 검사자 · 개정 확인란), 시트 · 각장 · 비용접구간은 ② 상단 결재란 1칸형. 공정마다 바꾸거나 「틀 안 씀」 으로 둘 수 있습니다. 판독 화면에서 그 공정을 직접 그리면(「이 공정 기본」) 그것이 틀보다 먼저입니다.'),
        tplTable(),
        h('h3', { style: 'margin-top:12px' }, '설정한 범위'),
        table(['범위', '설정', ''], [
          ['모든 검사표 기본', db.masks.length + '곳', db.masks.length ? h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { if (confirm('모든 검사표 기본 영역을 지울까요?')) { db.masks = []; save(); render(); } } }, '지우기') : '']
        ].concat(Object.keys(db.maskProcs || {}).map(function (k) {
          return ['공정 기본 — ' + k, db.maskProcs[k].length + '곳', h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { delete db.maskProcs[k]; save(); render(); } }, '지우기')];
        })).concat(Object.keys(db.maskSheets || {}).length ? [['파일 · 쪽별', Object.keys(db.maskSheets).length + '개 검사표', h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { if (confirm('파일 · 쪽별로 맞춘 영역을 모두 지울까요?')) { db.maskSheets = {}; save(); render(); } } }, '모두 지우기')]] : []))),
      AIP.settingsCard({ toast: toast, onChange: render }),
      h('section', { class: 'card' }, h('h2', null, '백업'),
        h('p', { class: 'note' }, '출하일정 · 파일 목록 · 판독값 · 덮어쓰기 표는 이 브라우저에만 있습니다. 다른 PC 로 옮기거나 브라우저 정리 전에 JSON 으로 받아 두세요. PDF 내용은 백업에 넣지 않습니다(용량).'),
        h('div', { class: 'btn-row' },
          h('button', { type: 'button', class: 'btn', onclick: function () { download('불량제품파악_백업_' + stamp() + '.json', new Blob([JSON.stringify(db)], { type: 'application/json' })); } }, 'JSON 백업 내려받기'),
          h('label', { class: 'btn' }, backup, 'JSON 백업 불러오기'),
          h('button', { type: 'button', class: 'btn btn-danger', onclick: function () { if (confirm('모든 자료(출하일정 · 파일 · 판독값 · 덮어쓰기 표)를 지울까요?')) { startFresh(); save(); render(); toast('모두 지웠습니다.'); } } }, '모든 자료 지우기')))
    ];
  }

  /* ── 라우터 ─────────────────────────────── */
  var PAGES = [
    { id: 'import', no: '0', t: '자료 넣기', s: '출하일정 · 검사표 PDF', fn: pageImport },
    { id: 'ship', no: '①', t: '출하 대조', s: '검사표 없음 · 불합격 · 미판정', fn: pageShip },
    { id: 'read', no: '②', t: '판독', s: '검사표 보며 값 옮기기 · AI', fn: pageRead },
    { id: 'result', no: '③', t: '검사 결과', s: '불합격 리스트 · 엑셀', fn: pageResult },
    { id: 'ai', no: '④', t: 'AI 요약', s: '숫자만 · 품번 가림', fn: pageAi },
    { id: 'spec', no: '⑤', t: '스펙', s: '읽은 규격 · 덮어쓰기', fn: pageSpec },
    { id: 'settings', no: '설', t: '설정', s: '판정 규칙 · AI · 백업', fn: pageSettings }
  ];
  function route() { var id = (location.hash || '').replace(/^#\//, '').split('?')[0]; if (id === 'check') id = 'ship'; return PAGES.filter(function (p) { return p.id === id; })[0] || (db.ships.length || db.files.length ? PAGES[1] : PAGES[0]); }
  function render() {
    var p = route();
    var nav = document.getElementById('nav'); nav.innerHTML = '';
    PAGES.forEach(function (x) {
      nav.appendChild(h('a', { href: '#/' + x.id, 'aria-current': x === p ? 'page' : null }, h('span', { class: 'no' }, x.no), h('span', { class: 't' }, x.t), h('span', { class: 's' }, x.s)));
    });
    var c = analysis().counts;
    document.getElementById('dataChip').textContent = db.ships.length || db.files.length ? '출하 ' + c.ships + '대 · 검사표 없음 ' + c.none + ' · 불합격 ' + c.fail + ' · 미판정 ' + c.undecided : '자료 없음';
    var b = document.getElementById('sampleBanner');
    b.hidden = !db.sample;
    b.textContent = '지금 보시는 것은 가상 샘플입니다(가짜 모델 MODEL-A100 · B200, 품번 QA-…, 검사표 PDF 7개 — 글자 층 4 · 스캔 그림 2 · 합본 1(1쪽 자주 검사 + 2쪽 시트, 2쪽은 누운 스캔), 기준일 ' + (db.refDate || '') + '). 실제 자료를 넣기 전에 설정의 「모든 자료 지우기」를 해 주세요.';
    var keepY = window.scrollY, rf0 = main.querySelector('.read-form'), vw0 = main.querySelector('.viewer');
    var keep = { form: rf0 ? rf0.scrollTop : 0, vt: vw0 ? vw0.scrollTop : 0, vl: vw0 ? vw0.scrollLeft : 0, id: p.id };
    main.innerHTML = '';
    add(main, p.fn());
    // 단추를 누를 때마다 화면이 맨 위로 올라가지 않게(11-11 요청) — 같은 화면을 다시 그릴 때만. 화면을 옮길 때는 hashchange 가 맨 위로 올림
    if (render._last === p.id) {
      var rf1 = main.querySelector('.read-form'), vw1 = main.querySelector('.viewer');
      if (rf1) rf1.scrollTop = keep.form; if (vw1) { vw1.scrollTop = keep.vt; vw1.scrollLeft = keep.vl; }
      window.scrollTo(0, keepY);
    }
    render._last = p.id;
    document.title = p.t + ' — 불량 제품 파악';
    save();
  }
  window.addEventListener('hashchange', function () { ui.drawMask = false; render(); main.focus(); window.scrollTo(0, 0); });
  var lastW = window.innerWidth;
  window.addEventListener('resize', function () { clearTimeout(render._r); render._r = setTimeout(function () { if (Math.abs(window.innerWidth - lastW) > 60 && route().id === 'read') { lastW = window.innerWidth; render(); } }, 300); });
  // 처음 열면 바로 체험할 수 있게 — 아무 자료도 없으면 샘플을 불러옵니다(1단계 자료는 형식이 달라 새 샘플로)
  if (db.legacy) setTimeout(function () { toast('2차에서 자료 형식(출하일정 · 호기)이 바뀌어 1단계 자료 대신 새 샘플로 시작합니다.'); }, 300);
  if (!db.ships.length && !db.files.length && !db.cleared) { loadSample(); }
  else {
    render();
    // 글자 층을 아직 확인하지 않은 파일(백업을 불러온 경우 등)
    // 여러 쪽 PDF 는 2026-11-10 전까지 1쪽만 읽었으니 쪽별 기록(pageText)이 없으면 나머지 쪽도 읽음(1쪽 판독값은 그대로 둠)
    var unchecked = db.files.filter(function (f) { return f.hasPdf !== false && ((f.text == null && !db.readings[f.id]) || (L.pagesOf(f) > 1 && !f.pageText)); });
    if (unchecked.length) autoReadAll(unchecked);
  }
})();
