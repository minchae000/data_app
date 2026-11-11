// 헤드리스 브라우저로 주요 흐름 점검 (선택 — playwright-core 가 있는 PC 에서만)
//   PLAYWRIGHT_CORE=/경로/node_modules/playwright-core node scripts/check-browser.mjs
// file:// 로 index.html 을 열어(2차): 처음 열면 샘플 → 글자 층 4장 자동 읽기 → ① 출하 대조 숫자 → ② 스캔 검사표(파일 이름 힌트 · 같은 모델 항목 미리 채움 · 가상 AI 답)
// → 식별 안 된 스캔(품번 · 호기를 직접 적고 값 하나를 벗어나게 → 바로 불합격) → ① 숫자가 바뀜 → 가릴 영역 · 가린 그림 받기
// → ③ 엑셀 다시 읽기 → ④ 프롬프트에 품번 · 모델 없음 → ⑤ 덮어쓰기 → 폭별 가로 넘침 → 모두 지우기 → 출하일정 엑셀 · PDF 를 파일 고르기로 넣기.
// JS 오류가 하나라도 있으면 실패. BREAK=1 이면 일부러 넓은 상자를 넣어 넘침 점검기가 잡는지 봅니다. SHOTS=<폴더> 면 화면을 찍어 둡니다.
import { createRequire } from 'node:module';
import path from 'node:path';
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
const require = createRequire(import.meta.url);
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const { chromium } = require(process.env.PLAYWRIGHT_CORE || 'playwright-core');
const XLSX = require(path.join(root, 'vendor/xlsx.full.min.js'));
const url = 'file://' + path.join(root, 'index.html');
const shots = process.env.SHOTS;

const errors = [], problems = [];
const browser = await chromium.launch();
const ctx = await browser.newContext({ acceptDownloads: true, viewport: { width: 1440, height: 1000 } });
const page = await ctx.newPage();
page.on('pageerror', e => errors.push('pageerror: ' + e.message));
page.on('console', m => { if (m.type() === 'error') errors.push('console: ' + m.text()); });
const outside = [];
page.on('request', r => { const u = r.url(); if (!/^(file|data|blob):/.test(u)) outside.push(u); });
page.on('dialog', d => d.accept());
function ok(cond, msg) { if (!cond) { problems.push(msg); console.log('  FAIL ' + msg); } else console.log('  ok  ' + msg); }
async function go(hash) { await page.evaluate(h => { location.hash = h; }, hash); await page.waitForTimeout(300); }
async function click(sel, text) { await page.locator(sel, text ? { hasText: text } : undefined).first().click(); await page.waitForTimeout(300); }
async function shot(name) { if (shots) await page.screenshot({ path: path.join(shots, name + '.png'), fullPage: false }); }
async function dl(fn) {
  const [d] = await Promise.all([page.waitForEvent('download'), fn()]);
  const f = path.join(root, '.check-' + Date.now() + path.extname(d.suggestedFilename()));
  await d.saveAs(f); const b = fs.readFileSync(f); fs.unlinkSync(f);
  return { name: d.suggestedFilename(), bytes: b };
}
const db = () => page.evaluate(() => JSON.parse(localStorage.getItem('data10-13.db')));
const tileText = () => page.evaluate(() => [...document.querySelectorAll('.tile')].map(t => t.innerText.replace(/\s+/g, ' ')));
const doneCount = (d) => Object.values(d.readings).filter(r => r.done === true).length;
async function waitIdle() { await page.waitForFunction(() => !document.querySelector('.alert.info') || !/읽는 중/.test(document.querySelector('.alert.info').textContent), null, { timeout: 30000 }); await page.waitForTimeout(300); }
async function pickFile(name) { await page.selectOption('.read-bar select', { label: [...(await page.$$eval('.read-bar select option', o => o.map(x => x.textContent)))].find(t => t.endsWith(name)) }); await page.waitForTimeout(400); await page.waitForSelector('.page-wrap canvas', { timeout: 20000 }); }
async function tilesAre(re, msg) { const t = await tileText(); ok(re.every((r, i) => r.test(t[i])), msg + ' — ' + t.join(' | ')); }

await page.goto(url);
await page.waitForTimeout(800);
await waitIdle();
let d = await db();
ok(d && d.sample && d.ships.length === 13 && d.files.length === 7, '처음 열면 가상 샘플(출하 13대 · 검사표 PDF 7개)');
ok(doneCount(d) === 5 && d.files.filter(f => f.text === true).length === 5 && d.files.filter(f => f.text === false).length === 2, '글자 층 5쪽은 바로 읽고, 스캔 2장은 「그림만」 으로 표시');
const multi = d.files.find(f => f.pages === 2);
ok(multi && multi.pageText && multi.pageText[1] === true && multi.pageText[2] === false, '합본 PDF 는 쪽마다 — 1쪽 글자 층 · 2쪽 그림만');
ok(page.url().endsWith('#/ship'), '샘플을 넣으면 ① 출하 대조로');
await tilesAre([/13대/, /검사표 없음 4대/, /불합격 1대/, /미판정 3대/], '① 검사표 없음 4 · 불합격 1 · 미판정 3(합본 2쪽 판독 전)');
const lists = await page.evaluate(() => [...document.querySelectorAll('.card h2')].map(h => h.textContent));
ok(lists.some(t => /식별이 필요한 검사표 \(3\)/.test(t)) && lists.some(t => /출하일정에 없는 검사 \(1\)/.test(t)), '식별 필요 2장 · 출하일정에 없는 검사 1 — ' + lists.filter(t => /식별|없는 검사/.test(t)).join(', '));
ok(await page.locator('.dd.soon').count() > 0, '출하가 가까운(3일 안) 남은 날을 강조');
ok(await page.locator('a.btn', { hasText: '수동 판독' }).count() >= 2, '미판정 · 식별 필요 줄에 「수동 판독」 단추');
await shot('ship');

// ② 스캔 — 파일 이름 힌트 · 같은 모델 항목 미리 채움 · 가상 AI 답
await go('#/read');
await page.waitForSelector('.page-wrap canvas', { timeout: 20000 });
const cw = await page.evaluate(() => { const c = document.querySelector('.page-wrap canvas'); return [c.width, c.height]; });
ok(cw[0] > 300 && cw[0] > cw[1], '검사표 그림을 그림(가로 양식 ' + cw.join('×') + ')');
await pickFile('MODEL-A100 상부 #314504.pdf');
const hint = await page.evaluate(() => ({ yellow: document.querySelectorAll('.from-name').length, hogi: document.querySelector('th.hcol input').value, cells: document.querySelectorAll('td.v input').length }));
ok(hint.yellow >= 2 && hint.hogi === '314504', '파일 이름에서 짐작한 모델 · 호기를 노란 칸으로 — ' + JSON.stringify(hint));
ok(hint.cells === 10, '같은 모델 검사표에서 항목 6개(LH/RH 포함 10칸)를 미리 채움 — ' + hint.cells);
await click('button', '가상 AI 답 예시 넣기');
await click('button', '답 읽어 칸 채우기');
const vals = await page.evaluate(() => [...document.querySelectorAll('td.v input')].map(i => i.value));
ok(vals.length === 10 && vals.every(v => v !== ''), 'AI 답으로 칸이 채워짐 — ' + vals.join(', '));
ok(await page.locator('td.v input.low').count() === 1, '확신 낮은 칸 하나가 노랗게');
ok(/합격/.test(await page.locator('tfoot td.jd').first().innerText()), '호기 판정이 바로 보임(상한과 같은 421 은 합격)');
await shot('read');
await click('button', '판독 완료 · 다음');
d = await db();
const r504 = Object.entries(d.readings).find(([k]) => k.includes('314504'));
ok(r504 && r504[1].done === true && /AI/.test(r504[1].by) && r504[1].meta.part === 'QA-70011', '판독 완료 저장(AI 읽기 + 사람 확인 · 품번까지)');
const next = await page.evaluate(() => document.querySelector('.read-bar select').selectedOptions[0].textContent);
ok(/1013~1015/.test(next), '다음 할 일(식별 안 된 스캔)로 넘어감 — ' + next);

// ② 식별 안 된 스캔 — 품번 · 호기를 직접 적고 값 하나를 벗어나게
await page.waitForSelector('.page-wrap canvas', { timeout: 20000 });
ok(await page.locator('td.v input').count() === 0, '처음에는 항목이 없음(모델 · 품번을 모름)');
await page.fill('input[aria-label="품번"]', 'QA-70011'); await page.press('input[aria-label="품번"]', 'Tab'); await page.waitForTimeout(400);
ok(await page.locator('td.v input').count() === 10, '품번을 적으면 같은 품번 검사표의 항목을 가져옴');
await page.fill('th.hcol input', '314505'); await page.press('th.hcol input', 'Tab'); await page.waitForTimeout(400);
const firstNum = page.locator('td.v input[data-k="2 LH"]');
await firstNum.fill('430'); await page.waitForTimeout(150);
ok(await page.evaluate(() => document.querySelector('td.v input[data-k="2 LH"]').parentNode.className.includes('fail')) && /불합격/.test(await page.locator('tfoot td.jd').first().innerText()), '공차 밖 값(420±1 에 430)을 적으면 바로 불합격');
await firstNum.fill('420.5'); await firstNum.press('Enter'); await page.waitForTimeout(150);
ok(await page.evaluate(() => document.activeElement && document.activeElement.dataset.k === '2 RH'), 'Enter 로 같은 호기의 아래 칸(2 RH)으로');
await click('button', '가상 AI 답 예시 넣기');
await click('button', '답 읽어 칸 채우기');
await click('button', '판독 완료 · 다음');
const nextMulti = await page.evaluate(() => document.querySelector('.read-bar select').selectedOptions[0].textContent);
ok(/314506\.pdf \(2쪽\)/.test(nextMulti) && /2쪽/.test(await page.locator('.read-form h2').first().innerText()), '다음 할 일은 합본의 2쪽(시트) — ' + nextMulti);

// ② 합본 2쪽 — 누운 스캔 돌리기 · 공정명 「시트 부착」 → 풍선 체크 항목 · 공정 기본 가릴 영역 · 이 파일만 따로
await page.waitForSelector('.page-wrap canvas', { timeout: 20000 });
const c0 = await page.evaluate(() => { const c = document.querySelector('.page-wrap canvas'); return c.width > c.height; });
await click('button', '오른쪽으로 90°');
await page.waitForSelector('.page-wrap canvas', { timeout: 20000 }); await page.waitForTimeout(300);
const c1 = await page.evaluate(() => { const c = document.querySelector('.page-wrap canvas'); return c.width > c.height; });
d = await db();
ok(!c0 && c1 && d.files.find(f => f.pages === 2).rot['2'] === 90, '누운 스캔(세로)을 오른쪽으로 90° 돌리면 가로로 바로 섬 · 쪽에 저장');
await page.fill('input[aria-label="제작 공정명"]', '시트 부착'); await page.press('input[aria-label="제작 공정명"]', 'Tab'); await page.waitForTimeout(400);
ok(await page.locator('td.v input').count() === 1 && /풍선/.test(await page.locator('td.lab').first().innerText()), '공정명 「시트 부착」 → 풍선 체크 항목 한 줄로 바뀜');
ok(/공정 「시트 부착」 기본/.test(await page.locator('main').innerText()), '시트 공정은 결재 칸 자리가 달라 공정 기본 가릴 영역을 씀');
await click('button', '가릴 영역 고치기');
await page.locator('.mask-layer').scrollIntoViewIfNeeded();
const box2 = await page.locator('.mask-layer').boundingBox();
await page.mouse.move(box2.x + box2.width * 0.05, box2.y + box2.height * 0.85); await page.mouse.down();
await page.mouse.move(box2.x + box2.width * 0.25, box2.y + box2.height * 0.95, { steps: 5 }); await page.mouse.up(); await page.waitForTimeout(300);
d = await db();
const own = Object.entries(d.maskSheets || {}).find(([k]) => k.endsWith('#2'));
ok(own && own[1].length === 2 && d.maskProcs['시트부착'].length === 1, '「이 파일 · 이 쪽만」 으로 그리면 공정 기본을 복사해 이 파일만 따로(공정 기본은 그대로)');
await click('button', '가릴 영역 고치기 끝');
await page.waitForSelector('.page-wrap canvas', { timeout: 20000 });
await click('button', '가상 AI 답 예시 넣기');
await click('button', '답 읽어 칸 채우기');
ok(await page.locator('td.v input').first().inputValue() === '23/24' && /불합격/.test(await page.locator('tfoot td.jd').first().innerText()), '풍선 24개 중 23개 체크 → 바로 불합격');
await click('button', '판독 완료 · 다음');
await go('#/ship');
await tilesAre([/13대/, /검사표 없음 3대/, /불합격 3대/, /미판정 1대/], '스캔 2장 + 합본 2쪽을 판독하면 ① 검사표 없음 3 · 불합격 3(외관 NG · 시트 미체크) · 미판정 1');
// 수동 판독 — 보류(빈 칸) 줄의 단추를 누르면 그 검사표의 빈 칸에 커서
await page.locator('a.btn', { hasText: '수동 판독' }).first().click(); await page.waitForTimeout(500);
await page.waitForSelector('.page-wrap canvas', { timeout: 20000 }); await page.waitForTimeout(200);
ok(await page.evaluate(() => { const a = document.activeElement; return !!(a && a.closest('td.v') && a.value === ''); }) && await page.locator('.manual-bar').count() === 1, '「수동 판독」 → 보류 검사표의 빈 칸에 커서 · 수동 판독 안내 줄');

// 가릴 영역 · 가린 그림
await go('#/read');
await page.waitForSelector('.page-wrap canvas', { timeout: 20000 });
const masks0 = (await db()).masks.length;
const own0 = Object.keys((await db()).maskSheets || {}).length;
await click('button', '가릴 영역 고치기');
const box = await page.locator('.mask-layer').boundingBox();
await page.mouse.move(box.x + box.width * 0.1, box.y + box.height * 0.3);
await page.mouse.down();
await page.mouse.move(box.x + box.width * 0.3, box.y + box.height * 0.4, { steps: 5 });
await page.mouse.up();
await page.waitForTimeout(300);
d = await db();
ok(d.masks.length === masks0 && Object.keys(d.maskSheets).length === own0 + 1 && Object.values(d.maskSheets).some(m => m.length === masks0 + 1), '끌어서 가릴 영역을 더함 — 기본(결재 · 서명 3곳)을 복사해 이 파일만 4곳, 기본은 그대로');
await page.waitForSelector('.page-wrap canvas', { timeout: 20000 });
const aiOpen = await page.locator('details.block[open] summary', { hasText: 'AI 로 읽기' }).count();
if (!aiOpen) await click('summary', 'AI 로 읽기');
const img = await dl(() => page.locator('button', { hasText: '그림 받기' }).first().click());
ok(img.bytes.length > 20000 && img.bytes[0] === 0xff && img.bytes[1] === 0xd8, '가린 그림을 JPG 로 받음 — ' + img.name + ' ' + Math.round(img.bytes.length / 1024) + 'KB');

// ③ 검사 결과 엑셀
await go('#/result');
const x = await dl(() => page.locator('button', { hasText: '엑셀로 내려받기' }).first().click());
const wb = XLSX.read(x.bytes, { type: 'buffer' });
ok(['불합격_리스트', '검사_결과', '측정값_전체', '출하_대조', '검사표_없음', '미판정', '식별_필요'].every(s => wb.SheetNames.includes(s)), '엑셀 시트 — ' + wb.SheetNames.join(', '));
const fails = XLSX.utils.sheet_to_json(wb.Sheets['불합격_리스트']);
ok(fails.length === 3 && fails.some(r => String(r['호기']) === '315010' && r['No'] === '3' && r['구분'] === 'LH' && String(r['적힌 값']) === '49.3') && fails.some(r => String(r['호기']) === '314505' && r['적힌 값'] === 'NG') && fails.some(r => String(r['호기']) === '314506' && r['적힌 값'] === '23/24' && r['공정'] === '시트 부착'), '불합격 리스트 — 315010 No 3 LH 49.3 · 314505 외관 NG · 314506 시트 23/24 (' + fails.length + '줄)');
ok(XLSX.utils.sheet_to_json(wb.Sheets['검사_결과']).length === 12, '검사 결과(호기별) 12(검사표 8쪽의 호기)');
ok(XLSX.utils.sheet_to_json(wb.Sheets['검사표_없음']).length === 3, '검사표 없음 3대');
await shot('result');

// ④ AI 요약
await go('#/ai');
const prompt = await page.locator('pre').first().innerText();
ok(!/QA-\d|MODEL-[AB]/.test(prompt) && /제품1/.test(prompt), '요약 프롬프트에 품번 · 모델 없음(제품1 · 모델1 …)');

// ⑤ 스펙 — 읽은 규격 · 덮어쓰기
await go('#/spec');
const specTxt = await page.locator('main').innerText();
ok(/419 ~ 421/.test(specTxt) && /47\.5 ~ 49/.test(specTxt) && /1\.5 이하/.test(specTxt), '⑤ 검사표 글자에서 읽은 규격(419~421 · 47.5~49 · 1.5 이하)');
await page.locator('button', { hasText: '덮어쓰기' }).nth(2).click(); await page.waitForTimeout(300);
ok((await db()).overrides.length === 1, '읽은 규격을 덮어쓰기 표로 옮김');
for (const hsh of ['#/settings', '#/import']) await go(hsh);
ok(true, '설정 · 자료 넣기 화면 열림');

// 4차(2026-11-10) — 모델별 규격(체크한 모델 · 규격 고르기) · 가릴 영역 틀(공정명 자동 · 바꾸기 · 글자 자리 짐작) · 출하일정 세 칸
await go('#/read');
await click('.read-bar .tabs button', '전체');
await pickFile('MODEL-A100-314501.pdf');
await page.locator('details.block summary', { hasText: '검사 항목 고치기' }).click();
const lab2 = page.locator('details.block table.spec tbody tr').nth(1).locator('input').nth(1);
await lab2.fill('MODEL-A100 SPEC:420±1 MODEL-A100E SPEC:418±1'); await lab2.dispatchEvent('change'); await page.waitForTimeout(500);
ok(await page.locator('.model-pick').count() === 1 && await page.locator('select.spec-pick').count() === 1, '한 칸에 모델별 규격 둘 — 「체크한 모델」 · 「규격 고르기」 가 나옴');
ok(await page.locator('table.sheet-grid td.v.hold').count() >= 3, '모델을 모르면 그 칸은 보류(노랑) — 첫 규격만 읽지 않음');
await page.locator('.model-pick select').selectOption('MODEL-A100'); await page.waitForTimeout(500);
ok(await page.locator('table.sheet-grid td.v.hold').count() === 0 && /모델 MODEL-A100/.test(await page.locator('table.sheet-grid td.lab').nth(1).innerText()), '「MODEL-A100」 을 고르면 그 규격(419~421)으로 판정');
await page.locator('select.spec-pick').selectOption('1'); await page.waitForTimeout(500);
ok(await page.locator('table.sheet-grid td.v.fail').count() >= 1, '규격을 직접 「MODEL-A100E 417~419」 로 고르면 420 근처 값은 불합격');
await page.locator('select.spec-pick').selectOption(''); await page.waitForTimeout(400);
// 가릴 영역 — 공정명 「가공」 → ① 4칸 틀 자동
const procIn = page.locator('.meta-grid input[aria-label="제작 공정명"]');
await procIn.fill('가공'); await procIn.dispatchEvent('change'); await page.waitForTimeout(800);
ok(await page.locator('.page-wrap .mask').count() === 4 && /가공 양식.*자동/.test(await page.locator('.card p.note', { hasText: 'AI 에 보낼 때 가리는 영역' }).innerText()), '공정명 「가공」 → ① 제작 · 가공형 4칸 틀(자동)');
await click('button', '가릴 영역 고치기');
await page.locator('.mask-help select[aria-label*="가릴 영역 틀"]').selectOption('top1'); await page.waitForTimeout(800);
ok(await page.locator('.page-wrap .mask').count() === 1, '틀을 ② 상단 결재란 1칸형으로 바꾸면 1칸');
await click('button', '글자 자리로 짐작(이 쪽)'); await page.waitForTimeout(800);
ok(await page.locator('.page-wrap .mask').count() === 3 && /이 파일 · 이 쪽만/.test(await page.locator('.card p.note', { hasText: 'AI 에 보낼 때 가리는 영역' }).innerText()), '글자 층 쪽 — 「입안 · 심사 · 결정」 · 「검사자」 자리로 3곳 짐작(이 쪽만)');
await shot('read_4th');
await go('#/settings');
ok(await page.locator('table select[aria-label*="가공"]').count() >= 1 && await page.locator('table select[aria-label*="가공"]').first().inputValue() === 'top1', '설정 — 공정별 틀 표(가공 = 고른 틀)');
await page.locator('table select[aria-label*="가공"]').first().selectOption(''); await page.waitForTimeout(300);
ok(!('가공' in ((await db()).maskTpl || {})), '설정에서 「자동」 으로 되돌림');
await go('#/import');
await page.locator('input[type=file]').nth(0).setInputFiles(path.join(root, 'samples/출하일정_대형형식_샘플.xlsx'));
await page.waitForTimeout(600);
const pend = await page.locator('main').innerText();
ok(await page.locator('label.opt input[type=checkbox]').first().isChecked() && /호기 4개/.test(pend) && /출하예정\s+모델\s+호기/.test(pend) && /알림 1건/.test(pend), '대형 형식 출하일정 — 세 칸만(품번 · 수량 · BOOM 무시) · 호기 4개 · 알림 1건(연도)');
await click('button', '취소');

// 폭별 가로 넘침
async function overflow() {
  return page.evaluate(() => {
    const W = document.documentElement.clientWidth, bad = [];
    for (const el of document.querySelectorAll('body *')) {
      const r = el.getBoundingClientRect();
      if (!r.width || !r.height) continue;
      let p = el.parentElement, inner = false;
      while (p) { const ox = getComputedStyle(p).overflowX; if (ox === 'auto' || ox === 'scroll') { inner = true; break; } p = p.parentElement; }
      if (!inner && r.right > W + 1) bad.push(el.tagName.toLowerCase() + '.' + el.className + ' ' + Math.round(r.right) + '>' + W);
    }
    return bad.slice(0, 3);
  });
}
const addBrk = () => page.evaluate(() => { if (document.getElementById('brk')) return; const b = document.createElement('div'); b.style.width = '2000px'; b.style.height = '5px'; b.id = 'brk'; document.body.appendChild(b); });
for (const w of [360, 390, 768, 1280]) {
  await page.setViewportSize({ width: w, height: 900 });
  for (const hsh of ['#/import', '#/ship', '#/read', '#/result', '#/ai', '#/spec', '#/settings']) {
    await go(hsh);
    if (hsh === '#/read') await page.waitForSelector('.page-wrap canvas', { timeout: 20000 });
    if (process.env.BREAK) await addBrk();
    const bad = await overflow();
    if (bad.length) problems.push(w + 'px ' + hsh + ' 가로 넘침: ' + bad.join(' | '));
    if (shots && (w === 390 || w === 1280) && (hsh === '#/read' || hsh === '#/ship')) await shot(w + hsh.replace('#/', '_'));
  }
  console.log('  ..  ' + w + 'px 7화면 넘침 점검');
}
await page.setViewportSize({ width: 1440, height: 1000 });

// 모든 자료 지우기 → 파일 고르기로 넣기
await go('#/settings');
await click('button', '모든 자료 지우기');
d = await db();
ok(d.ships.length === 0 && d.files.length === 0 && d.cleared, '모든 자료 지움');
await page.reload(); await page.waitForTimeout(600);
ok((await db()).files.length === 0, '지운 뒤 다시 열어도 샘플이 저절로 들어오지 않음');
await go('#/import');
await page.locator('input[type=file]').nth(0).setInputFiles(path.join(root, 'samples/출하일정_샘플.xlsx'));
await page.waitForTimeout(600);
const prev = await page.locator('main').innerText();
ok(/호기 13개/.test(prev), '출하일정 열 연결 · 미리 보기(호기 13개, 범위 펼침)');
await click('button', '이 내용으로 바꾸기');
const pdfDir = path.join(root, 'samples/공유폴더_샘플');
const pdfs = fs.readdirSync(pdfDir).filter(f => f.endsWith('.pdf')).map(f => path.join(pdfDir, f));
await page.locator('input[type=file][multiple]').first().setInputFiles(pdfs);
await page.waitForTimeout(800); await waitIdle();
d = await db();
ok(d.ships.length === 13 && d.files.length === 7 && !d.sample && doneCount(d) === 5, '엑셀 + PDF 7개를 파일 고르기로 넣음 → 글자 층 5쪽 자동 읽기');
await page.locator('input[type=file][multiple]').first().setInputFiles(pdfs.slice(0, 2));
await page.waitForTimeout(800); await waitIdle();
ok((await db()).files.length === 7, '같은 PDF 를 다시 넣으면 건너뜀');
await go('#/ship');
await tilesAre([/13대/, /검사표 없음 4대/, /불합격 1대/, /미판정 3대/], '내 파일로 넣어도 같은 결과');
// 초대형 형식 출하일정을 「더하기」 — 제목 줄 · 두 줄 머리 · 「11/10」
await go('#/import');
await page.locator('input[type=file]').nth(0).setInputFiles(path.join(root, 'samples/출하일정_초대형형식_샘플.xlsx'));
await page.waitForTimeout(600);
ok(/호기 5개/.test(await page.locator('main').innerText()), '초대형 형식 출하일정 — 호기 5개(범위 펼침) · 「11/10」 을 기준 연도로');
await click('button', '지금 출하일정에 더하기');
ok((await db()).ships.length === 18, '출하일정에 더함 13 + 5 = 18');
await go('#/read');
await page.waitForSelector('.page-wrap canvas', { timeout: 20000 });
ok(true, '넣은 PDF(IndexedDB)를 판독 화면에서 그림');

await browser.close();
ok(outside.length === 0, '바깥 주소 요청 0건' + (outside.length ? ' — ' + outside.slice(0, 3).join(', ') : ''));
if (errors.length) problems.push('JS 오류: ' + errors.slice(0, 5).join(' | '));
console.log(problems.length ? '\n실패 ' + problems.length + '건\n- ' + problems.join('\n- ') : '\n모두 통과');
process.exitCode = problems.length ? 1 : 0;
