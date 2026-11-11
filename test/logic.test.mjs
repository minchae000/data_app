// 실행: node test/logic.test.mjs   (의존성 없음 — vendor 의 SheetJS · pdf.js 만 씀)
// 2차(2026-10-28): 수강생이 보낸 「자주 검사 CHECK SHEET」 의 칸 배치만 본뜬 가상 샘플(scripts/make-samples.py)로 검사합니다.
// 실제 검사표 그림 · 값은 리포에 두지 않습니다.
import assert from 'node:assert/strict';
import { readFileSync, readdirSync, statSync } from 'node:fs';
import { createRequire } from 'node:module';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
const require = createRequire(import.meta.url);
const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const L = require('../js/logic.js');
const SM = require('../js/sample-data.js');
const Pdfs = require('../js/sample-pdfs.js');
const XLSX = require('../vendor/xlsx.full.min.js');
const E = require('../js/ai-endpoint.js');
// pdf.js 를 node 에서 — 그림 그리기용 canvas 가 없다는 경고는 글자 층 읽기와 관계없어 잠시 끕니다
const cw = console.warn, cl = console.log; console.warn = () => {}; console.log = () => {};
const pdfjs = require('../vendor/pdfjs/pdf.min.js'); console.warn = cw; console.log = cl;
pdfjs.GlobalWorkerOptions.workerSrc = path.join(ROOT, 'vendor/pdfjs/pdf.worker.min.js');

let passed = 0;
async function test(name, fn) {
  try { await fn(); passed++; console.log('  ok  ' + name); }
  catch (e) { console.error('  FAIL ' + name + '\n       ' + e.message); process.exitCode = 1; }
}
const aoaOf = (file) => {
  const wb = XLSX.read(readFileSync(path.join(ROOT, 'samples', file)), { type: 'buffer', cellDates: true });
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: null });
};
async function textOf(name) {
  const doc = await pdfjs.getDocument({ data: new Uint8Array(Buffer.from(Pdfs[name], 'base64')), isEvalSupported: false, verbosity: 0 }).promise;
  const pg = await doc.getPage(1), vp = pg.getViewport({ scale: 1 }), tc = await pg.getTextContent();
  return L.normTextItems(tc.items, vp.height);
}
const S0 = L.cleanSettings({});
const num = (lower, upper) => ({ kind: 'num', lower, upper, unit: '' });
const verdictOf = (spec, raw, s) => L.judgeItem(spec, raw, s || S0).status;
// 샘플 파일을 앱과 같은 모양으로
// 여러 쪽 샘플(kind 'multi')은 쪽마다 종류가 다름(pages: ['text', 'scan']) — 정답 열쇠는 1쪽 「이름」, 2쪽부터 「이름#2」
const pageKinds = (f) => f.pages || [f.kind];
const sampleFiles = () => SM.files.map(f => ({ id: L.fileId(f.name, f.size), name: f.name, size: f.size, text: pageKinds(f)[0] === 'text', pages: pageKinds(f).length }));
const readingOf = (name) => { const t = JSON.parse(JSON.stringify(SM.truth[name])); return { meta: t.meta, items: t.items, hogis: t.hogis, problems: t.problems, done: true, by: '시험', at: '2026-10-28 09:00' }; };
const readingsFor = (kinds) => { const o = {}; SM.files.forEach(f => { pageKinds(f).forEach((k, i) => { if (kinds.includes(k)) o[L.sheetKey(L.fileId(f.name, f.size), i + 1)] = readingOf(f.name + (i ? '#' + (i + 1) : '')); }); }); return o; };

console.log('규격 글자 읽기');
await test('대칭 공차 — 420±1 · 2740 ± 6 mm · 85+/-2 · 앞의 항목 이름 · 지름 기호, 단위는 기억만', () => {
  const a = L.parseSpecString('420±1');
  assert.deepEqual([a.ok, a.lower, a.upper, a.nominal, a.unit], [true, 419, 421, 420, '']);
  const b = L.parseSpecString('길이 2740 ± 6 mm');
  assert.deepEqual([b.lower, b.upper, b.unit], [2734, 2746, 'mm']);
  assert.deepEqual([L.parseSpecString('85+/-2').lower, L.parseSpecString('85 +- 2').upper], [83, 87]);
  assert.deepEqual([L.parseSpecString('Φ12.5±0.1').lower, L.parseSpecString('Φ12.5±0.1').upper], [12.4, 12.6]);
  assert.deepEqual([L.parseSpecString('73,5±2').lower], [71.5]);              // 쉼표 소수점
});
await test('비대칭 공차 · MAX · MIN · 범위 — 48 +1.0/-0.5 · 12(-0.2, +0) · 1.5 MAX · MAX 0.03 · MIN 40 · 40 이상 · 8.00~8.10', () => {
  const pick = (s) => { const r = L.parseSpecString(s); return [r.ok, r.lower, r.upper]; };
  assert.deepEqual(pick('48 +1.0/-0.5'), [true, 47.5, 49]);
  assert.deepEqual(pick('48 -0.5/+1.0'), [true, 47.5, 49]);
  assert.deepEqual(pick('12(+0/-0.2)'), [true, 11.8, 12]);
  assert.deepEqual(pick('평탄도 1.5 MAX'), [true, null, 1.5]);
  assert.deepEqual(pick('MAX 0.03'), [true, null, 0.03]);
  assert.deepEqual(pick('최대 2'), [true, null, 2]);
  assert.deepEqual(pick('MIN 40'), [true, 40, null]);
  assert.deepEqual(pick('40 이상'), [true, 40, null]);
  assert.deepEqual(pick('8.00~8.10'), [true, 8, 8.1]);
  assert.deepEqual(pick('8.10 ~ 8.00 mm'), [true, 8, 8.1]);
});
await test('규격이 아닌 글자는 읽지 않음 — 「가이드 홀 방향」 · 「외 관」 · 빈 칸', () => {
  ['가이드 홀 방향(전·후) 확인할 것', '외 관', '', 'Drum Slot Hole'].forEach(s => assert.equal(L.parseSpecString(s).ok, false, s));
});

console.log('판정');
await test('경계값 포함(10장 ④) — 420±1 이면 419 · 421 합격, 418.99 · 421.01 불합격 · 「미포함」 설정이면 421 불합격', () => {
  const sp = L.resolveSpec({ no: '2', label: '420±1', kind: 'num' }, 'P', []);
  assert.equal(verdictOf(sp, '419'), 'pass');
  assert.equal(verdictOf(sp, '421'), 'pass');
  assert.equal(verdictOf(sp, '420'), 'pass');
  assert.equal(verdictOf(sp, '418.99'), 'fail');
  assert.equal(verdictOf(sp, '421.01'), 'fail');
  assert.equal(L.judgeItem(sp, '421.01', S0).side, 'high');
  assert.equal(verdictOf(sp, '421', L.cleanSettings({ inclusive: false })), 'fail');
  // 소수 오차 — 0.1 + 0.2 처럼 만들어진 기준도 경계값은 합격
  assert.equal(verdictOf(num(0.1 + 0.2, 0.7), '0.3'), 'pass');
});
await test('비대칭 · MAX · MIN 경계 — 47.5 · 49 합격, 49.3 불합격 · 1.5 합격, 1.51 불합격 · 40 합격, 39.9 불합격', () => {
  const asym = L.resolveSpec({ label: '48 +1.0/-0.5' }, '', []), mx = L.resolveSpec({ label: '1.5 MAX' }, '', []), mn = L.resolveSpec({ label: 'MIN 40' }, '', []);
  assert.deepEqual(['47.5', '49', '49.3', '47.4'].map(v => verdictOf(asym, v)), ['pass', 'pass', 'fail', 'fail']);
  assert.deepEqual(['1.5', '1.51', '0'].map(v => verdictOf(mx, v)), ['pass', 'fail', 'pass']);
  assert.deepEqual(['40', '39.9', '1000'].map(v => verdictOf(mn, v)), ['pass', 'fail', 'pass']);
});
await test('값 읽기 — 단위를 붙여도 · 쉼표 소수점 · 육안 OK/NG/○/× · 체크(V · ✓) · 빈 칸 · 「?」 · 규격 없음', () => {
  const sp = num(419, 421), eye = { kind: 'okng' };
  assert.equal(verdictOf(sp, '420.5mm'), 'pass');
  assert.equal(verdictOf(sp, '420,5'), 'pass');
  assert.equal(verdictOf(num(0.4, 0.6), '0.5cm'), 'pass');                     // 기준 단위가 없으면 숫자끼리(단위는 따지지 않음)
  assert.deepEqual(['OK', '○', 'NG', '×', '', '?'].map(v => verdictOf(eye, v)), ['pass', 'pass', 'fail', 'fail', 'blank', 'unread']);
  assert.equal(verdictOf(sp, 'V'), 'pass');
  assert.equal(verdictOf(sp, '✓', L.cleanSettings({ checkAs: 'hold' })), 'unread');
  assert.equal(verdictOf(sp, 'NG'), 'fail');                                   // 수치 칸에 NG 를 적으면 불합격
  assert.equal(verdictOf(sp, '42?.3'), 'unread');
  assert.equal(verdictOf(num(null, null), '12'), 'nospec');
});
await test('호기 판정 — LH/RH 줄이 따로, 한 줄이라도 벗어나면 그 호기 불합격(빈 칸이 함께 있어도) · 칸 이름 「2-rh」 도 읽음', () => {
  const items = [{ no: '1', label: '방향 확인', tool: '육안' }, { no: '2', label: '420±1', tool: '줄자', subs: ['LH', 'RH'] }, { no: '3', label: '85±2', subs: ['LH', 'RH'] }];
  const u = (values, done = true) => L.judgeUnit({ items, values, meta: { part: 'P' }, done }, [], S0);
  const ok = u({ '1': 'OK', '2 LH': '420.2', '2 RH': '419.8', '3 LH': '85', '3 RH': '86' });
  assert.equal(ok.verdict, 'pass'); assert.equal(ok.cells.length, 5);
  const rh = u({ '1': 'OK', '2 LH': '420.2', '2-rh': '421.4', '3 LH': '85', '3 RH': '86' });
  assert.equal(rh.verdict, 'fail'); assert.deepEqual(rh.fails.map(x => x.key), ['2 RH']);
  assert.equal(u({ '1': 'OK', '2 LH': '420.2', '2 RH': '', '3 LH': '89', '3 RH': '86' }).verdict, 'fail');
  assert.equal(u({ '1': 'OK', '2 LH': '420.2', '2 RH': '', '3 LH': '85', '3 RH': '86' }).verdict, 'hold');
  assert.equal(L.judgeUnit({ items, values: { '1': 'OK', '2 LH': '420.2', '2 RH': '', '3 LH': '85', '3 RH': '86' }, meta: {}, done: true }, [], L.cleanSettings({ blankAs: 'ignore' })).verdict, 'pass');
  assert.equal(u({ '1': 'NG', '2 LH': '420', '2 RH': '420', '3 LH': '85', '3 RH': '85' }).verdict, 'fail');
  assert.equal(u({}, false).verdict, 'wait');                                 // 판독 끝나지 않았으면 판정하지 않음
});
await test('스펙 덮어쓰기 표가 검사표 글자보다 먼저 — 품번 + 항목 이름 또는 No', () => {
  const it = { no: '2', label: '420±1' };
  assert.equal(L.resolveSpec(it, 'QA-1', [{ part: 'QA-1', item: '420±1', lower: 419.5, upper: 420.5 }]).upper, 420.5);
  assert.equal(L.resolveSpec(it, 'QA-1', [{ part: 'QA-1', item: '2', lower: 400, upper: 440 }]).lower, 400);
  assert.equal(L.resolveSpec(it, 'QA-1', [{ part: 'QA-9', item: '2', lower: 400, upper: 440 }]).src, 'sheet');   // 다른 품번은 무시
  const ov = L.parseSpec([['품번', '검사 항목', '기준값', '공차'], ['QA-1', '420±1', 420, 0.5], [null, '외관', null, null]]);
  assert.deepEqual([ov.rows[0].lower, ov.rows[0].upper], [419.5, 420.5]);
  assert.equal(ov.warnings.length, 1);                                        // 외관 — 기준이 없음(판정 방식을 적으라고 알림)
});

console.log('검사표 글자 층');
await test('글자 층 샘플 4장 — 머리 칸 · 항목(규격 · 측정기 · LH/RH) · 호기 열 · 값 · 문제점을 정답대로 읽음', async () => {
  const texts = SM.files.filter(f => f.kind === 'text');
  assert.equal(texts.length, 4);
  for (const f of texts) {
    const r = L.parseSheetText(await textOf(f.name)), t = SM.truth[f.name];
    assert.ok(r.ok, f.name + ' ' + r.reason);
    assert.equal(r.template, 'self-check');
    assert.deepEqual(r.meta, t.meta, f.name + ' 머리 칸');
    assert.deepEqual(r.items.map(i => [i.no, i.label.replace(/\s+/g, ''), i.tool, i.subs]), t.items.map(i => [i.no, i.label.replace(/\s+/g, ''), i.tool, i.subs]), f.name + ' 항목');
    assert.deepEqual(r.hogis.map(x => x.no), t.hogis.map(x => x.no), f.name + ' 호기');
    t.hogis.forEach((th, i) => {
      const want = Object.fromEntries(Object.entries(th.values).filter(([, v]) => v !== ''));
      assert.deepEqual(r.hogis[i].values, want, f.name + ' 호기 ' + th.no + ' 값');
    });
    assert.deepEqual(r.problems, t.problems, f.name + ' 문제점');
    assert.deepEqual(r.warnings, []);
    assert.deepEqual(r.items.map(i => i.kind), t.items.map(i => (i.tool === '육안' ? 'okng' : 'num')));
  }
});
await test('스캔 샘플 2장 — 글자 층이 없어 「notext」, PDF 안에 호기 번호가 글자로 없음(사람 · AI 판독 대상)', async () => {
  const scans = SM.files.filter(f => f.kind === 'scan');
  assert.equal(scans.length, 2);
  for (const f of scans) {
    const items = await textOf(f.name);
    assert.equal(items.length, 0, f.name);
    assert.equal(L.parseSheetText(items).reason, 'notext');
    const raw = Buffer.from(Pdfs[f.name], 'base64').toString('latin1');
    assert.ok(!raw.includes(SM.truth[f.name].hogis[0].no));
  }
});
await test('표 왼쪽 도면 글자(LH · RH · 치수)와 검사자 줄(서명 · 이름)은 값으로 읽지 않음', () => {
  const T = [
    ['자주 검사 CHECK SHEET', 40, 40], ['모델명', 40, 100], ['MODEL-Z9', 120, 100], ['품번', 260, 100], ['ZZ-1', 320, 100],
    ['No.', 430, 230], ['검사 항목', 470, 230], ['측정기', 540, 236], ['호기수', 590, 224], ['901', 610, 230], ['902', 646, 230], ['문제점', 780, 226],
    ['1', 438, 270], ['420±1', 478, 270], ['줄자', 545, 270], ['LH', 580, 262], ['RH', 580, 280], ['420.1', 605, 262], ['421.5', 641, 280], ['419.9', 605, 280],
    ['LH', 200, 270], ['RH', 200, 280], ['(2) 420', 300, 262],                                   // 도면 글자(표 왼쪽)
    ['2', 438, 305], ['외 관', 478, 305], ['육안', 545, 305], ['OK', 608, 305], ['NG', 644, 305], ['찍힘', 785, 305],
    ['검사자', 470, 320], ['홍길동', 605, 320], ['김철수', 641, 320]                                // 서명 줄(마지막 줄 바로 아래)
  ].map(([s, x, y]) => ({ s, x, y, w: s.length * 5, h: 8, cx: x + s.length * 2.5 }));
  const r = L.parseSheetText(T);
  assert.ok(r.ok);
  assert.deepEqual(r.meta, { model: 'MODEL-Z9', part: 'ZZ-1' });
  assert.deepEqual(r.items.map(i => [i.no, i.label, i.subs.join('/'), i.kind]), [['1', '420±1', 'LH/RH', 'num'], ['2', '외 관', '', 'okng']]);
  assert.deepEqual(r.hogis, [{ no: '901', values: { '1 LH': '420.1', '1 RH': '419.9', '2': 'OK' } }, { no: '902', values: { '1 RH': '421.5', '2': 'NG' } }]);
  assert.deepEqual(r.problems, { '2': '찍힘' });
  assert.ok(!JSON.stringify(r).includes('홍길동'));
});

console.log('식별(모델 · 품번 · 호기)');
const known = L.knownIds(SM.ships, {});
await test('파일 이름은 약한 힌트 — 「모델-호기」 · 「img00001」 · 「모델 상부 #호기」 · 「모델 하부센터 (날짜)」 · 「날짜~날짜 작업」', () => {
  const g = (n) => { const x = L.hintFromName(n, known); return [x.model, x.part, x.hogi]; };
  assert.deepEqual(g('MODEL-A100-314501.pdf'), ['MODEL-A100', '', '314501']);
  assert.deepEqual(g('img00001.pdf'), ['', '', '']);
  assert.deepEqual(g('MODEL-A100 상부 #314504.pdf'), ['MODEL-A100', '', '314504']);
  assert.deepEqual(g('MODEL-B200 하부센터 (2026.10.12).pdf'), ['MODEL-B200', '', '']);   // 날짜 숫자를 호기로 읽지 않음
  assert.deepEqual(g('1013~1015 작업.pdf'), ['', '', '']);
  assert.deepEqual(g('QA-80022_315011.pdf'), ['', 'QA-80022', '315011']);
  assert.deepEqual(g('MODEL-A100 20261015.pdf'), ['MODEL-A100', '', '']);            // 모르는 숫자(날짜)는 # 이 없으면 호기 아님
});
await test('검사표 내용이 파일 이름보다 먼저 — 이름이 다른 모델 · 호기를 가리켜도 내용으로 식별, 내용이 없으면 「파일 이름으로만」 · 「식별 안 됨」', () => {
  const files = [{ id: 'a', name: 'MODEL-B200-315012.pdf' }, { id: 'b', name: 'MODEL-A100 상부 #314504.pdf' }, { id: 'c', name: 'img00009.pdf' }];
  const readings = { a: { meta: { model: 'MODEL-A100 / MODEL-A100E', part: 'QA-70011' }, items: [], hogis: [{ no: '314501', values: {} }], done: true } };
  const b = L.buildUnits(files, readings, known);
  assert.deepEqual(b.ids.map(i => i.status), ['ok', 'hint', 'unknown']);
  assert.deepEqual([b.units[0].meta.part, b.units[0].meta.model, b.units[0].hogi], ['QA-70011', 'MODEL-A100 / MODEL-A100E', '314501']);
  assert.equal(b.ids[0].src.model, 'content');
  assert.deepEqual([b.ids[1].src.model, b.ids[1].hogiSrc, b.units[1].hogi, b.units[1].done], ['filename', 'filename', '314504', false]);
  assert.deepEqual(b.ids[2].missing, ['품번 · 모델', '호기']);
  assert.equal(b.units.length, 2);
});

console.log('출하일정');
await test('출하일정 샘플 엑셀 — 제목 줄 아래 머리행 · 엑셀 날짜 · 숫자 호기 · 호기 범위 펼침 · 품번 빈 줄은 모델로', () => {
  const r = L.parseShip(aoaOf('출하일정_샘플.xlsx'));
  assert.equal(r.det.headerRow, 2);
  assert.deepEqual(r.rows.map(x => { const o = Object.assign({}, x); delete o.row; return o; }), SM.ships);
  assert.equal(r.rows.length, 13);
  assert.deepEqual(r.rows.slice(-2).map(x => [x.hogi, x.qty]), [['315015', 1], ['315016', 1]]);
  assert.deepEqual(r.warnings, []);
});
await test('호기 칸 나누기 — 범위 · 쉼표 · 숫자 · # · 「호기」 꼬리 · 100대 넘는 범위는 펼치지 않음', () => {
  assert.deepEqual(L.splitHogi('314498~314500'), ['314498', '314499', '314500']);
  assert.deepEqual(L.splitHogi('A-0098~A-0101'), ['A-0098', 'A-0099', 'A-0100', 'A-0101']);
  assert.deepEqual(L.splitHogi('314498, 314502 / #314510'), ['314498', '314502', '314510']);
  assert.deepEqual(L.splitHogi(314498), ['314498']);
  assert.deepEqual(L.splitHogi('314498호기'), ['314498']);
  assert.deepEqual(L.splitHogi('1~500'), ['1~500']);
  assert.deepEqual(L.splitHogi(''), []);
});
await test('출하 대조(샘플, 글자 층만 읽은 상태) — 검사표 없음 4 · 불합격 1 · 미판정 3(대기 2 · 보류 1) · 합격 5 · 출하일정에 없는 검사 1', () => {
  const an = L.analyze(SM.ships, sampleFiles(), readingsFor(['text']), [], S0, SM.refDate);
  const c = an.counts;
  assert.deepEqual([c.ships, c.none, c.fail, c.wait, c.hold, c.pass, c.extra, c.unknown, c.hint], [13, 4, 1, 2, 1, 5, 1, 1, 2]);
  const st = Object.fromEntries(an.rows.map(r => [r.ship.hogi, r.status]));
  assert.deepEqual([st['315010'], st['315012'], st['315013'], st['314503'], st['314504'], st['314505']], ['fail', 'pass', 'hold', 'pass', 'wait', 'none']);
  assert.deepEqual(an.fails.map(x => [x.hogi, x.no + ' ' + x.sub, x.raw, x.problem]), [['315010', '3 LH', '49.3', '재가공 후 재측정']]);
  assert.equal(an.extra[0].hogi, '314499');
  // 출하일 순 · 남은 날
  const none = L.shipList(an, ['none']);
  assert.deepEqual(none.map(r => [r.ship.hogi, L.dday(r.days)]), [['314505', 'D-5'], ['315014', 'D-7'], ['315015', 'D-8'], ['315016', 'D-8']]);
  // 합본(1쪽 자주 검사 합격 + 2쪽 시트 판독 전) — 공정 둘을 모두 봐서 「판독 대기」
  const r506 = an.rows.find(r => r.ship.hogi === '314506');
  assert.deepEqual([r506.status, r506.procs.length, r506.procs.map(x => x.verdict)], ['wait', 2, ['pass', 'wait']]);
  assert.match(an.rows.find(r => r.ship.hogi === '314504').notes.join(), /파일 이름으로만/);
});
await test('스캔 2장까지 판독하면 — 품번 없는 출하 줄은 모델+호기로 짝지어 외관 NG 불합격, 상한과 같은 421 은 합격', () => {
  const an = L.analyze(SM.ships, sampleFiles(), readingsFor(['text', 'scan']), [], S0, SM.refDate);
  const c = an.counts;
  assert.deepEqual([c.none, c.fail, c.wait, c.hold, c.pass, c.unknown, c.hint], [3, 3, 0, 1, 6, 0, 0]);
  // 합본 2쪽 시트 — 풍선 24개 중 23개 체크 → 그 호기는 불합격(1쪽 자주 검사는 합격이어도)
  const r506 = an.rows.find(r => r.ship.hogi === '314506');
  assert.deepEqual([r506.status, r506.unit.page, r506.unit.meta.process], ['fail', 2, '시트 부착']);
  assert.ok(an.fails.some(x => x.hogi === '314506' && x.raw === '23/24' && /2쪽/.test(x.file) && x.process === '시트 부착'));
  const r505 = an.rows.find(r => r.ship.hogi === '314505');
  assert.deepEqual([r505.status, r505.how], ['fail', 'model']);
  assert.equal(an.rows.find(r => r.ship.hogi === '314504').status, 'pass');
});
await test('짝 짓기 규칙 — 품번+호기가 먼저, 품번이 다르면 모델+호기로 짝짓고 알림, 같은 호기 검사표가 둘이면 나중에 판독한 것', () => {
  const unit = (id, part, model, hogi, v, at) => ({ id, name: id + '.pdf' });
  const files = [unit('x'), unit('y'), unit('z')];
  const items = [{ no: '1', label: '10±1' }];
  const rd = (part, model, hogi, v, at) => ({ meta: { part, model }, items, hogis: [{ no: hogi, values: { '1': v } }], done: true, at });
  const readings = { x: rd('P-1', 'M-1', '100', '10', '2026-10-01 10:00'), y: rd('P-OLD', 'M-2', '200', '10', '2026-10-01 10:00'), z: rd('P-1', 'M-1', '100', '12', '2026-10-02 10:00') };
  const ships = [{ date: '2026-10-05', part: 'P-1', model: 'M-1', hogi: '100' }, { date: '2026-10-06', part: 'P-2', model: 'M-2', hogi: '200' }, { date: '2026-10-07', part: '', model: 'M-1', hogi: '999' }];
  const an = L.analyze(ships, files, readings, [], S0, '2026-10-04');
  assert.deepEqual(an.rows.map(r => [r.status, r.how]), [['fail', 'part'], ['pass', 'model'], ['none', '']]);
  assert.match(an.rows[0].notes.join(), /같은 호기 검사표 2개/);
  assert.match(an.rows[1].notes.join(), /모델\+호기/);
});

console.log('AI 반자동 판독');
await test('판독 프롬프트 — 칸 이름(No + 구분)과 규격만, 회사 정보 없음, 서명은 읽지 말라고 함, 모르는 양식이면 항목도 읽어 달라고 함', () => {
  const t = SM.truth['MODEL-A100 상부 #314504.pdf'];
  const p = L.buildSheetPrompt(t.items);
  assert.match(p, /"2 LH"/); assert.match(p, /"5 RE"/); assert.match(p, /420±1/); assert.match(p, /서명은 읽지도 적지도 마/);
  assert.ok(!p.includes('QA-70011') && !p.includes('MODEL-A100'));
  assert.match(L.buildSheetPrompt([]), /"항목"/);
});
await test('AI 답 읽기 — 코드 묶음 · 끝 쉼표 · 「2-lh」 같은 칸 이름 · 모르는 칸 · 머리 칸 · 문제점 · 확신 낮음 · 항목까지 읽어 온 답', () => {
  const t = SM.truth['1013~1015 작업.pdf'];
  const ans = '```json\n{"모델명":"MODEL-A100","품번":"QA-70011","작성일":"2026.10.15","호기":[{"호기":"#314505","값":{"1":"OK","2-lh":"420.5","2 RH":"419.7","9 XX":"1",}}],"문제점":{"6":"찍힘"},"확신낮음":["2 lh"],}\n```';
  const p = L.parseSheetAi(ans, t.items);
  assert.ok(p.ok, p.error);
  assert.deepEqual(p.meta, { model: 'MODEL-A100', part: 'QA-70011', date: '2026-10-15' });
  assert.deepEqual(p.hogis, [{ no: '314505', values: { '1': 'OK', '2 LH': '420.5', '2 RH': '419.7' } }]);
  assert.deepEqual(p.unknown, ['9 XX']); assert.deepEqual(p.problems, { '6': '찍힘' }); assert.deepEqual(p.lowConf, ['2 LH']);
  const q = L.parseSheetAi(JSON.stringify({ 항목: [{ No: '1', 검사항목: '30±1', 구분: ['LH', 'RH'] }], 호기: [{ 호기: '7', 값: { '1 LH': '30.2', '1 RH': '31.5' } }] }), []);
  assert.deepEqual([q.items.length, q.items[0].subs, q.hogis[0].values['1 RH']], [1, ['LH', 'RH'], '31.5']);
  assert.equal(L.parseSheetAi('답이 없습니다', []).ok, false);
});
await test('요약 프롬프트 — 품번 · 모델을 별칭으로 가리고, 답의 별칭을 원래 이름으로 되돌림', () => {
  const an = L.analyze(SM.ships, sampleFiles(), readingsFor(['text']), [], S0, SM.refDate);
  const b = L.buildSummaryPrompt(an, { mask: true });
  assert.ok(!/QA-\d|MODEL-[AB]/.test(b.text), b.text.match(/QA-\d+|MODEL-\w+/));
  assert.match(b.text, /제품1/); assert.match(b.text, /7일 안에 출하/);
  assert.equal(L.unmask('제품1 확인', b.map), Object.keys(b.map).find(k => b.map[k] === '제품1') + ' 확인');
  assert.match(L.buildSummaryPrompt(an, { mask: false }).text, /QA-80022/);
});
await test('그림을 함께 보내는 요청 — 이미지 칸이 든 messages 를 그대로 보냄', () => {
  const msgs = [{ role: 'user', content: [{ type: 'text', text: '읽어줘' }, { type: 'image_url', image_url: { url: 'data:image/png;base64,AAAA' } }] }];
  const req = E.buildChatRequest({ baseUrl: 'https://api.openai.com/v1/', model: 'm', apiKey: 'k' }, msgs, { maxTokens: 800 });
  const body = JSON.parse(req.init.body);
  assert.equal(req.url, 'https://api.openai.com/v1/chat/completions');
  assert.equal(body.messages[0].content[1].image_url.url.slice(0, 10), 'data:image');
  assert.equal(E.STORE_KEY, 'data10-13.ai');
});

console.log('2026-11-10 요청 반영');
await test('측정값 칸은 실측값 또는 OK(답 ①) — 수치 항목에 OK 는 합격(설정과 관계없이), NG 는 불합격, V 만 한 칸은 종전대로 설정을 따름', () => {
  const sp = num(419, 421);
  const ok = L.judgeItem(sp, 'OK', L.cleanSettings({ checkAs: 'hold' }));
  assert.deepEqual([ok.status, ok.okMark, ok.check], ['pass', true, false]);
  assert.equal(verdictOf(sp, 'ok'), 'pass');
  assert.equal(verdictOf(sp, 'NG'), 'fail');
  assert.equal(verdictOf(sp, 'V', L.cleanSettings({ checkAs: 'hold' })), 'unread');
  const j = L.judgeUnit({ items: [{ no: '1', label: '420±1' }, { no: '2', label: '85±2' }], values: { '1': 'OK', '2': '85.5' }, meta: {}, done: true }, [], S0);
  assert.deepEqual([j.verdict, j.checks.length], ['pass', 0]);
});
await test('시트 공정(답 ①) — 풍선 체크: OK · V · 「24/24」 합격, 「23/24」 · 누락 · NG 불합격, 「25/24」 는 확인 필요 · 공정명에 「시트」 면 이 항목으로 시작', () => {
  const at = { kind: 'attach' };
  assert.deepEqual(['OK', 'V', '✓', '24/24', '전부'].map(v => verdictOf(at, v)), ['pass', 'pass', 'pass', 'pass', 'pass']);
  assert.deepEqual(['23/24', '누락', 'NG', '미부착'].map(v => verdictOf(at, v)), ['fail', 'fail', 'fail', 'fail']);
  assert.match(L.judgeItem(at, '23 / 24', S0).msg, /24개 중 1개 체크 없음/);
  assert.equal(verdictOf(at, '25/24'), 'unread');
  assert.equal(verdictOf(at, ''), 'blank');
  assert.equal(L.cleanItem({ no: '1', label: '시트 부착 확인(모델링 뷰 풍선 체크)' }, 0).kind, 'attach');
  assert.equal(L.specText(L.resolveSpec(L.cleanItem({ no: '1', label: '풍선 체크' }, 0), '', [])), '부착 확인(풍선 체크)');
  assert.equal(L.procPreset('시트 부착')[0].kind, 'attach');
  assert.equal(L.procPreset('SEAT')[0].no, '1');
  assert.equal(L.procPreset('ASSY 검사'), null);
  assert.match(L.buildSheetPrompt(L.procPreset('시트 부착')), /체크 수\/풍선 수/);
});
await test('다른 공정 양식의 규격 글자 — 「80㎛↑ 관리」 는 하한, 「진직도2↓」 는 상한 · 머리 칸 MODEL · P/No. · PART NAME · C/No. · 검사일', () => {
  const a = L.parseSpecString('도막측정기 (80㎛↑ 관리)'), b = L.parseSpecString('진직도2↓');
  assert.deepEqual([a.ok, a.lower, a.upper, a.unit], [true, 80, null, '㎛']);
  assert.deepEqual([b.ok, b.lower, b.upper], [true, null, 2]);
  assert.equal(verdictOf(L.resolveSpec({ no: '3', label: '도막 두께 80㎛↑', kind: 'num' }, '', []), '79'), 'fail');
  const t = (s, x, y) => ({ s, x, y, w: s.length * 6, h: 8, cx: x + s.length * 3 });
  const r = L.parseSheetText([t('MODEL', 10, 10), t('M-1', 60, 10), t('P/No.', 120, 10), t('QA-1', 160, 10), t('PART NAME', 10, 30), t('프레임', 80, 30), t('C/No.', 140, 30), t('QF-1', 180, 30), t('검사일', 10, 50), t('2026.11.02', 60, 50)]);
  assert.deepEqual([r.meta.model, r.meta.part, r.meta.name, r.meta.mgmt, r.meta.date], ['M-1', 'QA-1', '프레임', 'QF-1', '2026-11-02']);
});
await test('초대형 출하일정 형식(답 ③) — 제목 줄 · 두 줄 머리 · 「모   델」 · 품번 칸 없음 · 「11/10」 은 기준 연도로 · 호기 범위 펼침', () => {
  const r = L.parseShip(aoaOf('출하일정_초대형형식_샘플.xlsx'), null, { year: 2026 });
  assert.equal(r.det.headerRow, 2);
  assert.deepEqual([r.det.map.model, r.det.map.hogi, r.det.map.date, r.det.map.part], [4, 5, 8, undefined]);
  assert.deepEqual(r.rows.map(x => [x.date, x.model, x.hogi]), [['2026-11-10', 'MODEL-C300L', '520101'], ['2026-11-10', 'MODEL-D500', '610907'], ['2026-11-12', 'MODEL-A100L', '314501'], ['2026-11-14', 'MODEL-C300L', '520102'], ['2026-11-14', 'MODEL-C300L', '520103']]);
  assert.deepEqual(r.warnings.map(w => w.msg), ['연도 없이 「월/일」로 적힌 출하일 4줄은 2026년으로 봤습니다(기준일의 연도).']);   // Center · Side 둘째 머리 줄은 조용히 건너뜀
  assert.match(L.parseShip(aoaOf('출하일정_초대형형식_샘플.xlsx')).warnings.map(w => w.msg).join(), /출하일을 읽지 못했습니다/);   // 연도를 안 주면 읽지 않고 알림
});
await test('출하일정 합치기(소형 · 초대형 · 버켓 파일) · 모델 이름이 조금 달라도(「MODEL-A100L」 ↔ 「MODEL-A100」) 같은 호기면 짝', () => {
  const big = L.parseShip(aoaOf('출하일정_초대형형식_샘플.xlsx'), null, { year: 2026 }).rows;
  const all = L.mergeShips(SM.ships, big);
  assert.equal(all.length, 13 + 5);
  assert.equal(L.mergeShips(all, big).length, 18);                       // 같은 파일을 다시 더해도 늘지 않음
  const an = L.analyze(big, sampleFiles(), readingsFor(['text']), [], S0, '2026-11-10');
  const r = an.rows.find(x => x.ship.hogi === '314501');
  assert.deepEqual([r.status, r.how], ['pass', 'model']);
  assert.match(r.notes.join(), /모델 이름이 조금 다름/);
  assert.deepEqual([L.nearModelName('MODEL-A100L', 'MODEL-A100'), L.nearModelName('A1', 'A12'), L.nearModelName('MODEL-A100', 'MODEL-B200')], [true, false, false]);
});
await test('한 PDF 에 여러 쪽(답 ④) — 쪽마다 검사표 · 열쇠는 1쪽 「파일id」 · 2쪽 「파일id#2」, 「검사표 아님」 으로 뺀 쪽은 셈하지 않음', () => {
  assert.deepEqual([L.sheetKey('f', 1), L.sheetKey('f', 2), L.sheetKey('f')], ['f', 'f#2', 'f']);
  assert.deepEqual(L.sheetPages({ pages: 3, skip: { 2: true } }), [1, 3]);
  const files = [{ id: 'f', name: '9월08일~9월23일 작업.pdf', pages: 3 }];
  const items = [{ no: '1', label: '10±1' }];
  const rd = (hogi, v) => ({ meta: { model: 'M-1', process: '가공' }, items, hogis: [{ no: hogi, values: { '1': v } }], done: true, at: 'x' });
  const b = L.buildUnits(files, { f: rd('100', '10'), 'f#2': rd('101', '10'), 'f#3': rd('102', '13') }, {});
  assert.deepEqual(b.units.map(u => [u.sheet, u.page, u.hogi]), [['f', 1, '100'], ['f#2', 2, '101'], ['f#3', 3, '102']]);
  const an = L.analyze([{ date: '2026-11-11', model: 'M-1', hogi: '102' }], files, { f: rd('100', '10'), 'f#3': rd('102', '13') }, [], S0, '2026-11-10');
  assert.deepEqual([an.rows[0].status, an.fails[0].file], ['fail', '9월08일~9월23일 작업.pdf · 3쪽']);
  assert.equal(L.buildUnits([{ id: 'f', name: 'a.pdf', pages: 3, skip: { 2: true, 3: true } }], {}, {}).ids.length, 1);
});
await test('같은 호기 · 다른 공정 검사표는 모두 봄 — 하나라도 불합격이면 불합격, 하나라도 판독 전이면 미판정, 같은 공정은 나중 것', () => {
  const items = [{ no: '1', label: '10±1' }];
  const rd = (proc, v, at, done) => ({ meta: { model: 'M-1', process: proc }, items, hogis: [{ no: '7', values: { '1': v } }], done: done !== false, at });
  const files = ['a', 'b', 'c'].map(id => ({ id, name: id + '.pdf' }));
  const ship = [{ date: '2026-11-11', model: 'M-1', hogi: '7' }];
  const run = (rs) => L.analyze(ship, files, rs, [], S0, '2026-11-10').rows[0];
  assert.equal(run({ a: rd('제작', '10', '1'), b: rd('도장', '10', '1') }).status, 'pass');
  assert.equal(run({ a: rd('제작', '10', '1'), b: rd('도장', '13', '1') }).status, 'fail');
  assert.equal(run({ a: rd('제작', '10', '1'), b: rd('도장', '10', '1', false) }).status, 'wait');
  assert.equal(run({ a: rd('제작', '13', '1'), b: rd('제작', '10', '2'), c: rd('도장', '10', '1') }).status, 'pass');   // 제작 재검사(나중 것)가 합격
  assert.match(run({ a: rd('제작', '10', '1'), b: rd('도장', '13', '1') }).notes.join(), /공정 2개 검사표를 모두 봄/);
});
await test('가릴 영역(추가 요청) — 파일(쪽)별 > 공정별 > 모든 검사표, 빈 배열도 「가리지 않음」 으로 존중 · 옛 자료는 빈 칸으로 읽힘', () => {
  const all = [{ x: 0.1, y: 0.1, w: 0.1, h: 0.1 }], proc = [{ x: 0.5, y: 0.1, w: 0.2, h: 0.1 }], own = [{ x: 0.7, y: 0.7, w: 0.1, h: 0.1 }];
  const d = { masks: all, maskProcs: { [L.procKey('시트 부착')]: proc }, maskSheets: { 'f#2': own, g: [] } };
  // 「ASSY 검사」 는 4차부터 공정 틀(ASSY 양식)이 자동으로 붙음 — 모르는 공정명이라야 「모든 검사표 기본」
  assert.deepEqual([L.masksFor(d, 'f', 'ASSY 검사').src, L.masksFor(d, 'f', '하부 프레임 사상').src, L.masksFor(d, 'f', '시트  부착').src, L.masksFor(d, 'f#2', '시트 부착').src, L.masksFor(d, 'g', '').src], ['template', 'all', 'process', 'sheet', 'sheet']);
  assert.equal(L.masksFor(d, 'f', '시트 부착').masks, proc);
  assert.equal(L.masksFor(d, 'g', '').masks.length, 0);
  const r = L.restoreDb({ v: 2, masks: all, maskProcs: { a: [{ x: 1, y: 1, w: 1, h: 1 }, { x: 'z' }], b: 'x' }, maskSheets: { s: own } });
  assert.deepEqual([r.maskProcs.a.length, 'b' in r.maskProcs, r.maskSheets.s.length], [1, false, 1]);
  const old = L.restoreDb({ v: 2, masks: all, ships: [], files: [] });     // 2차(10-28) 자료 — 새 칸이 없음
  assert.deepEqual([old.masks.length, old.maskProcs, old.maskSheets, L.masksFor(old, 'f', '시트 부착').src, L.masksFor(old, 'f', '하부 프레임 사상').src], [1, {}, {}, 'template', 'all']);   // 4차 — 「시트」 는 틀이 자동
});
await test('돌아간 쪽(추가 요청) — 글자 방향으로 돌릴 각도를 알아내고(0 · 90 · 180 · 270), 돌려 옮긴 글자 층을 똑같이 읽음', async () => {
  const T = (a, b, c, d) => ({ str: '가나다', transform: [a, b, c, d, 0, 0] });
  assert.deepEqual([L.textRotation([T(9, 0, 0, 9)]), L.textRotation([T(0, 9, -9, 0)]), L.textRotation([T(-9, 0, 0, -9)]), L.textRotation([T(0, -9, 9, 0)]), L.textRotation([])], [0, 90, 180, 270, 0]);
  assert.deepEqual([L.normRot(-90), L.normRot(450), L.normRot(180)], [270, 90, 180]);
  // 글자 층 샘플을 시계 반대 90° 로 누인 쪽(세로)으로 만들고 → 90° 돌려 읽으면 원래와 같은 결과
  const name = SM.files.find(f => f.kind === 'text').name;
  const doc = await pdfjs.getDocument({ data: new Uint8Array(Buffer.from(Pdfs[name], 'base64')), isEvalSupported: false, verbosity: 0 }).promise;
  const pg = await doc.getPage(1), vp = pg.getViewport({ scale: 1 }), raw = (await pg.getTextContent()).items;
  const Hh = vp.height;
  const lying = raw.map(t => { const m = t.transform; return Object.assign({}, t, { transform: [-m[1], m[0], -m[3], m[2], Hh - m[5], m[4]] }); });
  assert.equal(L.textRotation(lying), 90);
  const vp90 = pg.getViewport({ scale: 1, rotation: 90 });
  assert.deepEqual(Array.from(vp90.transform).map(x => Math.round(x * 1000) / 1000), [0, 1, 1, 0, 0, 0]);   // 누운 세로 쪽을 90° 돌린 화면 행렬과 같은 모양
  const want = L.parseSheetText(L.normTextItems(raw, Hh));
  const got = L.parseSheetText(L.normTextItems(L.viewportTextItems(lying, [0, 1, 1, 0, 0, 0], Hh), Hh));
  assert.deepEqual([got.meta, got.items, got.hogis], [want.meta, want.items, want.hogis]);
  const same = L.parseSheetText(L.normTextItems(L.viewportTextItems(raw, vp.transform, Hh), Hh));   // 돌리지 않은 쪽은 그대로
  assert.deepEqual(same.hogis, want.hogis);
});

console.log('2026-11-10 4차 요청 반영');
await test('가릴 영역 틀(4차 ②) — 공정명으로 자동(평행도 · ASSY · 가공 · 제작 · 도장 → ①, 시트 · 각장 · 비용접 → ②), 설정이 자동보다 먼저 · 세 겹은 그대로', () => {
  const P = ['CHANNEL 평행도 검사', "ASS'Y 검사", '가공', '제작', '도장', '시트', '각장', '비용접구간', '시트 부착', '하부 프레임 사상', ''];
  assert.deepEqual(P.map(L.tplForProcess), ['flat4', 'assy4', 'make4', 'make4', 'paint2', 'top1', 'top1', 'top1', 'top1', '', '']);
  const g = (id) => L.maskTemplate(id);
  assert.deepEqual(['make4', 'flat4', 'assy4'].map(id => g(id).masks.map(m => m.label).sort()), Array(3).fill(['개정 확인란', '검사자', '결재(아래)', '결재(위)']));   // ① = 기존 3칸 + 확인란
  assert.deepEqual([g('paint2').masks.length, g('top1').masks.length, g('top1').group, g('make4').group], [2, 1, 2, 1]);
  L.MASK_TEMPLATES.forEach(t => t.masks.forEach(m => assert.ok(m.x >= 0 && m.y >= 0 && m.x + m.w <= 1 && m.y + m.h <= 1 && m.w > 0.02 && m.h > 0.01, t.id + ' ' + m.label)));
  // 시트 틀은 위쪽 10% 안 한 줄(상단 결재란) · 가공 틀의 결재(아래)는 아래쪽
  assert.ok(g('top1').masks[0].y < 0.1 && g('make4').masks.find(m => m.label === '결재(아래)').y > 0.8);
  const all = [{ x: 0.1, y: 0.1, w: 0.1, h: 0.1 }], own = [{ x: 0.7, y: 0.7, w: 0.1, h: 0.1 }], drawn = [{ x: 0.2, y: 0.2, w: 0.2, h: 0.1 }];
  const d = { masks: all, maskProcs: { [L.procKey('도장')]: drawn }, maskSheets: { s: own }, maskTpl: { [L.procKey('가공')]: 'top1', [L.procKey('시트')]: 'none' } };
  const f = (sk, p) => { const r = L.masksFor(d, sk, p); return [r.src, r.tpl || '', r.auto || false, r.masks.length]; };
  assert.deepEqual(f('x', 'CHANNEL 평행도 검사'), ['template', 'flat4', true, 4]);      // 자동
  assert.deepEqual(f('x', '가공'), ['template', 'top1', false, 1]);                    // 설정에서 바꿈
  assert.deepEqual(f('x', '시트'), ['all', '', false, 1]);                            // 「틀 안 씀」 → 모든 검사표 기본
  assert.deepEqual(f('x', '도장'), ['process', '', false, 1]);                        // 직접 그린 공정 기본이 틀보다 먼저
  assert.deepEqual(f('s', '가공'), ['sheet', '', false, 1]);                          // 이 파일 · 이 쪽만이 가장 먼저
  const m1 = L.masksFor(d, 'x', '제작').masks; m1[0].x = 0.99;                         // 틀은 복사본 — 고쳐도 틀이 바뀌지 않음
  assert.equal(L.maskTemplate('make4').masks[0].x, 0.83);
  const r = L.restoreDb({ v: 2, maskTpl: { a: 'top1', b: 'none', c: 'zzz', '': 'make4' } });
  assert.deepEqual(r.maskTpl, { a: 'top1', b: 'none' });
  assert.deepEqual(L.restoreDb({ v: 2 }).maskTpl, {});                                // 옛 자료 — 모두 자동
});
await test('글자 층이 있는 쪽은 「입안 · 심사 · 결정」 · 「검사자」 · 「개정 … 확인」 낱말 자리로 가릴 영역을 짐작(샘플 결재 칸과 겹침 50% 이상) · 스캔 쪽은 짐작하지 않음', async () => {
  const iou = (a, b) => { const x = Math.max(0, Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x)), y = Math.max(0, Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y)); const i = x * y; return i / (a.w * a.h + b.w * b.h - i); };
  for (const f of SM.files.filter(x => x.kind === 'text')) {
    const doc = await pdfjs.getDocument({ data: new Uint8Array(Buffer.from(Pdfs[f.name], 'base64')), isEvalSupported: false, verbosity: 0 }).promise;
    const pg = await doc.getPage(1), vp = pg.getViewport({ scale: 1 });
    const est = L.masksFromText(L.normTextItems((await pg.getTextContent()).items, vp.height), vp.width, vp.height);
    assert.deepEqual(est.map(m => m.label).sort(), ['검사자', '결재', '결재'], f.name);
    SM.masks.forEach(m => assert.ok(Math.max(...est.map(e => iou(e, m))) >= 0.5, f.name + ' ' + m.label));
  }
  // 「개정 내용 … 확인」 열 — 개정일 줄이 끝나는 곳까지
  const T = (s, x, y, w) => ({ s, x, y, w: w || s.length * 7, h: 8, cx: x + (w || s.length * 7) / 2 });
  const est2 = L.masksFromText([T('개정내용', 100, 500), T('확인', 300, 500), T('2024.03.02', 20, 515), T('2025.01.10', 20, 530), T('1', 5, 530, 6)], 800, 600);
  assert.equal(est2.length, 1);
  assert.equal(est2[0].label, '개정 확인란');
  assert.ok(est2[0].y * 600 > 500 && (est2[0].y + est2[0].h) * 600 >= 535 && est2[0].x * 800 < 300 && (est2[0].x + est2[0].w) * 800 > 314);
  assert.deepEqual(L.masksFromText([T('확인', 300, 500)], 800, 600), []);           // 개정 표 머리가 없는 「확인」 은 아님
  assert.deepEqual(L.masksFromText([], 800, 600), []);                               // 스캔(글자 없음)
});
await test('출하일정은 「모델 · 호기 · 출하예정」 세 칸만(4차 ③) — 소형 · 대형 · 초대형 형식 모두, 나머지 칸(상부 · 하부 · Line · 품번 · 수량 · BOOM)은 읽지도 알리지도 않음', () => {
  const small = L.parseShip(aoaOf('출하일정_소형형식_샘플.xlsx'), null, { year: 2026, only3: true });
  assert.deepEqual([small.det.headerRow, small.det.map], [0, { model: 1, hogi: 2, date: 3 }]);
  assert.deepEqual(small.rows.map(x => [x.date, x.model, x.hogi, x.part, x.qty]), [['2026-11-12', 'MODEL-A100', '314501', '', null], ['2026-11-12', 'MODEL-A100', '314502', '', null], ['2026-11-13', 'MODEL-B200', '315010', '', null], ['', 'MODEL-B200', '315011', '', null]]);
  assert.deepEqual(small.warnings.map(w => w.msg), ['출하예정이 빈 줄 1개는 날짜 없이 넣었습니다(호기는 대조함).']);   // 빈 날짜는 한 줄로만
  const big = L.parseShip(aoaOf('출하일정_대형형식_샘플.xlsx'), null, { year: 2026, only3: true });
  assert.deepEqual([big.det.headerRow, big.det.map], [2, { model: 1, hogi: 3, date: 6 }]);   // 「호기수(참고)」 보다 이름이 똑같은 「호  기」
  assert.deepEqual(big.rows.map(x => [x.date, x.model, x.hogi, x.part]), [['2026-11-10', 'MODEL-A100', '314503', ''], ['2026-11-12', 'MODEL-B200', '315012', ''], ['2026-11-12', 'MODEL-B200', '315013', ''], ['2026-11-14', 'MODEL-C300', '520104', '']]);
  assert.deepEqual(big.warnings.map(w => w.msg), ['연도 없이 「월/일」로 적힌 출하일 3줄은 2026년으로 봤습니다(기준일의 연도).']);   // 「11/10(화)」 · 「11월 12일」
  const ultra = L.parseShip(aoaOf('출하일정_초대형형식_샘플.xlsx'), null, { year: 2026, only3: true });
  assert.deepEqual([ultra.det.map, ultra.rows.length, ultra.warnings.length], [{ model: 4, hogi: 5, date: 8 }, 5, 1]);
  // 끄면(세 칸만 읽기 해제) 품번 · 수량도 읽음 — 예전 방식
  const full = L.parseShip(aoaOf('출하일정_대형형식_샘플.xlsx'), null, { year: 2026 });
  assert.deepEqual([full.det.map.part, full.rows[0].part, full.rows[0].qty], [4, 'QA-70011', 1]);
  // 모델 칸이 없으면 세 칸 방식에서는 필수 칸이 빠졌다고 알림
  assert.deepEqual(L.parseShip([['품번', '호기', '출하예정'], ['P-1', '1', '2026-11-10']], null, { only3: true }).det.missing, ['모델']);
  // 세 칸만 쓰므로 출하일정에 품번이 없어도 검사표와 모델 + 호기로 짝
  const an = L.analyze(big.rows, sampleFiles(), readingsFor(['text']), [], S0, '2026-11-10');
  assert.deepEqual(an.rows.filter(r => r.ship.hogi === '314503' || r.ship.hogi === '315012').map(r => [r.ship.hogi, r.status, r.how]), [['314503', 'pass', 'model'], ['315012', 'pass', 'model']]);
});
await test('모델명 칸의 체크 표시 읽기(4차 ④) — ✓ · ■ · ○ · (V) · 띄어 쓴 V, 빈 상자 □ 는 체크 아님, 이름 속 V 는 그대로 · 한 칸의 모델별 규격 나누기', () => {
  const mk = (s) => { const r = L.modelMarks(s); return [r.models, r.checked]; };
  assert.deepEqual(mk('MODEL-A100 / MODEL-B200✓ / MODEL-C300'), [['MODEL-A100', 'MODEL-B200', 'MODEL-C300'], ['MODEL-B200']]);
  assert.deepEqual(mk('□A100 ■B200 □C300'), [['A100', 'B200', 'C300'], ['B200']]);
  assert.deepEqual(mk('○A100/B200'), [['A100', 'B200'], ['A100']]);
  assert.deepEqual(mk('A100(V)/B200'), [['A100', 'B200'], ['A100']]);
  assert.deepEqual(mk('V B200 / EV100 / VX-2'), [['B200', 'EV100', 'VX-2'], ['B200']]);
  assert.deepEqual(mk('MODEL-A100 / MODEL-A100E'), [['MODEL-A100', 'MODEL-A100E'], []]);
  assert.deepEqual(L.splitModels('MODEL-A100✓/MODEL-B200'), ['MODEL-A100', 'MODEL-B200']);      // 짝 지을 때는 체크 표시를 뗀 이름
  const ms = L.modelSpecs('LH 상면 높이 MODEL-A100 SPEC:366±2 MODEL-B200 SPEC:363±2');
  assert.deepEqual(ms.map(x => [x.models, x.lower, x.upper]), [[['MODEL-A100'], 364, 368], [['MODEL-B200'], 361, 365]]);
  const ms2 = L.modelSpecs('높이(H/G) A100 BR:571±2 / A200E/EHR:468±2');                       // 「E/EHR」 → 「A200E · A200EHR」
  assert.deepEqual(ms2.map(x => x.models), [['A100 BR'], ['A200E', 'A200EHR']]);
  assert.deepEqual([L.modelSpecs('420±1'), L.modelSpecs('LH 420±1 RH 420±2'), L.modelSpecs('A100: 10±1 / B200: 10±1'), L.modelSpecs('48 +1.0/-0.5')], [[], [], [], []]);   // 규격 하나 · 구분 줄 · 같은 규격은 모델별이 아님
});
await test('한 칸에 모델별 규격이 둘(4차 ④) — 체크한 모델의 규격으로 판정, 모르면 그 칸 「보류」(첫 규격만 읽지 않음) · 사람이 모델이나 규격을 고르면 판정', () => {
  const items = [{ no: '1', label: 'LH 상면 높이 MODEL-A100 SPEC:366±2 MODEL-B200 SPEC:363±2' }, { no: '2', label: '10±1' }];
  const u = (model, v, extra) => Object.assign({ items, meta: { model }, values: { '1': v, '2': '10' }, done: true }, extra || {});
  const J = (x) => { const j = L.judgeUnit(x, [], S0); return [j.verdict, j.cells[0].status]; };
  assert.deepEqual(J(u('MODEL-A100 ✓ / MODEL-B200', '367')), ['pass', 'pass']);       // A100 = 364~368
  assert.deepEqual(J(u('MODEL-A100 / ✓MODEL-B200', '367')), ['fail', 'fail']);        // B200 = 361~365 → 상한 초과
  assert.deepEqual(J(u('MODEL-B200', '364')), ['pass', 'pass']);                       // 모델이 하나뿐이면 그것
  assert.deepEqual(J(u('MODEL-A100 / MODEL-B200', '367')), ['hold', 'modelspec']);     // 체크를 못 읽음 → 보류
  assert.deepEqual(J(u('MODEL-A100 ✓ / MODEL-B200 ✓', '367')), ['hold', 'modelspec']); // 체크가 둘 → 보류
  assert.deepEqual(J(u('MODEL-C300', '367')), ['hold', 'modelspec']);                  // 어느 규격에도 없는 모델 → 보류
  assert.match(L.judgeUnit(u('MODEL-A100 / MODEL-B200', '367'), [], S0).cells[0].msg, /모델별 규격이 2개/);
  // 사람이 고름 — 「체크한 모델」 또는 이 항목의 규격 번호
  assert.deepEqual(J(u('MODEL-A100 / MODEL-B200', '367', { meta: { model: 'MODEL-A100 / MODEL-B200', checkModel: 'MODEL-B200' } })), ['fail', 'fail']);
  assert.deepEqual(J(u('MODEL-A100 / MODEL-B200', '367', { specPick: { '1': 0 } })), ['pass', 'pass']);
  assert.deepEqual(J(u('MODEL-A100 ✓ / MODEL-B200', '367', { specPick: { '1': 1 } })), ['fail', 'fail']);   // 고른 규격이 체크보다 먼저
  assert.deepEqual(J(u('MODEL-A100 / MODEL-B200', 'OK')), ['pass', 'pass']);           // OK 는 규격과 관계없이 합격
  assert.deepEqual(J(u('MODEL-A100 / MODEL-B200', '?', { meta: { model: 'MODEL-A100 / MODEL-B200', checkModel: '?' } })), ['hold', 'unread']);
  // 가까운 이름(출하일정 · 검사표의 꼬리 글자 차이)도 그 모델의 규격으로
  assert.deepEqual(J(u('MODEL-A100L', '367')), ['pass', 'pass']);
  assert.equal(L.specText(L.resolveSpec(L.cleanItem(items[0], 0), '', [], { model: '' })), '모델별 규격 2개(MODEL-A100 364 ~ 368 / MODEL-B200 361 ~ 365) — 모델을 골라 주세요');
  // 저장 → 출하 대조까지: 보류는 미판정, 체크한 모델을 적으면 판정
  const files = [{ id: 'f', name: 'a.pdf' }], ship = [{ date: '2026-11-11', model: 'MODEL-A100', hogi: '7' }];
  const rd = (meta) => ({ f: { meta, items, hogis: [{ no: '7', values: { '1': '367', '2': '10' } }], done: true, at: '1' } });
  assert.equal(L.analyze(ship, files, rd({ model: 'MODEL-A100 / MODEL-B200' }), [], S0, '2026-11-10').rows[0].status, 'hold');
  assert.equal(L.analyze(ship, files, rd({ model: 'MODEL-A100 / MODEL-B200', checkModel: 'MODEL-A100' }), [], S0, '2026-11-10').rows[0].status, 'pass');
  // AI 답의 「체크모델」 도 같은 칸으로 · 판독 프롬프트가 체크모델을 묻고, 모델별 규격 항목이 있으면 고르지 말라고 함
  const p = L.parseSheetAi(JSON.stringify({ 모델명: 'MODEL-A100 / MODEL-B200', 체크모델: 'MODEL-B200', 호기: [{ 호기: '7', 값: { '1': '367' } }] }), items);
  assert.deepEqual([p.meta.model, p.meta.checkModel], ['MODEL-A100 / MODEL-B200', 'MODEL-B200']);
  const pr = L.buildSheetPrompt(items);
  assert.match(pr, /"체크모델"/); assert.match(pr, /어느 규격을 쓸지는 고르지 마/);
  assert.doesNotMatch(L.buildSheetPrompt([{ no: '1', label: '10±1' }]), /어느 규격을 쓸지/);
});

console.log('샘플 · 저장 · 리포');
await test('샘플 PDF 7장 — base64 가 PDF 이고 크기 · 내용이 samples/ 파일과 같음', () => {
  assert.equal(SM.files.length, 7);
  SM.files.forEach(f => {
    const b = Buffer.from(Pdfs[f.name], 'base64');
    assert.equal(b.slice(0, 5).toString(), '%PDF-', f.name);
    assert.equal(b.length, f.size, f.name);
    assert.deepEqual(readFileSync(path.join(ROOT, 'samples/공유폴더_샘플', f.name)), b, f.name);
  });
});
await test('저장 데이터 되살리기 — 1단계(v1) 자료는 legacy 로 비우고, 2차는 이상한 값만 버림 · 기준일 유지', () => {
  const v1 = L.restoreDb({ plan: [{ part: 'A' }], files: [{ id: '1', name: 'a.pdf' }] });
  assert.equal(v1.legacy, true); assert.equal(v1.files.length, 0);
  const d = L.restoreDb({ v: 2, ships: [{ part: 'A', hogi: '1' }, { part: 'B' }, null], files: [{ id: '1', name: 'a.pdf' }, { name: 'x' }], settings: { nearPct: 99, checkAs: 'hold' }, masks: [{ x: 0.1, y: 0.1, w: 0.2, h: 0.1 }, { x: 'a' }], refDate: '2026.10.28' });
  assert.deepEqual([d.ships.length, d.files.length, d.masks.length, d.settings.nearPct, d.settings.checkAs, d.refDate], [1, 1, 1, 10, 'hold', '2026-10-28']);
  assert.deepEqual(L.restoreDb('x'), L.emptyDb());
});
await test('리포에 회사 · 지명 낱말과 첨부 검사표의 실제 모델 · 품번 · 관리 번호가 없음(js · 문서 · 스크립트 · 샘플 PDF 글자 층)', async () => {
  // 낱말을 리포에 그대로 두지 않으려고 글자 번호로 적습니다
  const banned = [[0xac74, 0xd654], [0xcc3d, 0xc6d0], [0xd604, 0xb300], [0xb450, 0xc0b0], [0xd55c, 0xc6b8], [0x45, 0x43, 0x37, 0x35, 0x30], [0x31, 0x34, 0x37, 0x30, 0x37, 0x31, 0x33, 0x34], [0x4b, 0x48, 0x51, 0x2d]].map(a => String.fromCharCode(...a));
  const files = [];
  (function walk(d) { for (const n of readdirSync(d)) { if (['.git', 'vendor', 'node_modules'].includes(n)) continue; const p = path.join(d, n); if (statSync(p).isDirectory()) walk(p); else if (/\.(js|mjs|md|py|html|css|json)$/.test(n)) files.push(p); } })(ROOT);
  // 회사 이름은 2026-11-10 애본님 지시로 홈페이지 머리글에 일부러 넣는다(js/auth-config.js · docs/plan.json 의 landing) — 그 두 곳만 빼고 본다
  let txt = files.filter(p => !p.endsWith(path.join('js', 'auth-config.js'))).map(p => {
    const t = readFileSync(p, 'utf8');
    if (p.endsWith(path.join('docs', 'plan.json'))) { const j = JSON.parse(t); delete j.landing; return JSON.stringify(j); }
    return t;
  }).join('\n');
  for (const f of SM.files.filter(x => x.kind === 'text')) txt += (await textOf(f.name)).map(t => t.s).join(' ');
  banned.forEach(w => assert.ok(!txt.includes(w), '낱말 ' + w.length + '글자'));
});

const { versions } = await import('../scripts/bump-version.mjs');
await test('판 맞추기 — index.html ?v= · logic.js · app.js 가 같은 판 (scripts/bump-version.mjs)', () => {
  const v = versions();
  assert.equal(v.logic, L.VERSION, 'logic.js VERSION'); assert.equal(v.app, L.VERSION, 'app.js APP_VERSION');
  assert.ok(v.html.length >= 8 && v.html.every(x => x === L.VERSION), 'index.html 의 모든 js·css 에 ?v=' + L.VERSION + ' — ' + JSON.stringify(v.html));
});

console.log('\n' + passed + '개 통과' + (process.exitCode ? ' — 실패 있음' : ''));
