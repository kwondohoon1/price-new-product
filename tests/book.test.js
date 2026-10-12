// 참조 옮기기 규칙 시험. 실제 가격세팅 파일 시험은 PNP_FIXTURE(레포 밖 경로)가 있을 때만 돈다.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { createRequire } from 'node:module';
import { shiftFormulaForInsert, shiftSqref, shiftFormulaRows, mapRefs, openBook, saveBook, recalcBook, literalValue, numToCol } from '../src/book.js';
import { locate, addProducts, guessBrand, guessCategory, defaultAnchor, outputName } from '../src/newproduct.js';

const require = createRequire(import.meta.url);
const JSZip = require('jszip');

test('같은 시트 참조: 끼운 행 이상만 내려간다(절대 참조 포함)', () => {
  assert.equal(shiftFormulaForInsert('=A21+A22+$D$22+$D$3', 'S', 'S', 22, 1), '=A21+A23+$D$23+$D$3');
});

test('다른 시트를 가리키는 참조는 그 시트에 끼울 때만 바뀐다', () => {
  const f = '=IF(정산가세팅!L22="","",ROUNDUP(정산가세팅!L22*(1-F7),0))';
  assert.equal(shiftFormulaForInsert(f, '마켓별상품코드', '정산가세팅', 22, 1), '=IF(정산가세팅!L23="","",ROUNDUP(정산가세팅!L23*(1-F7),0))');
  assert.equal(shiftFormulaForInsert(f, '마켓별상품코드', '마켓별상품코드', 7, 1), '=IF(정산가세팅!L22="","",ROUNDUP(정산가세팅!L22*(1-F8),0))');
});

test('따옴표 시트명·문자열 속 글자·함수 이름은 건드리지 않는다', () => {
  assert.equal(shiftFormulaForInsert("='G마켓 시트'!B30&\"A30\"", 'X', 'G마켓 시트', 30, 1), "='G마켓 시트'!B31&\"A30\"");
  assert.equal(shiftFormulaForInsert('=LOG10(A30)', 'S', 'S', 5, 1), '=LOG10(A31)');
  assert.equal(shiftFormulaForInsert('=SUM(A:A)', 'S', 'S', 5, 1), '=SUM(A:A)');
});

test('범위: 시작이 아래면 통째로, 걸치면 끝만 늘고, 바로 아래에 끼우면 그대로', () => {
  assert.equal(shiftFormulaForInsert('=SUM(B17:B38)', 'S', 'S', 10, 1), '=SUM(B18:B39)');
  assert.equal(shiftFormulaForInsert('=SUM(B17:B38)', 'S', 'S', 22, 1), '=SUM(B17:B39)');
  assert.equal(shiftFormulaForInsert('=SUM(B17:B38)', 'S', 'S', 39, 1), '=SUM(B17:B38)');
  assert.equal(shiftSqref('D17:D38 D42:D63', 22, 1), 'D17:D39 D43:D64');
  assert.equal(shiftSqref('A1:Q23', 24, 1), 'A1:Q23');
});

test('행 복제용 이동은 상대 행만 옮긴다', () => {
  assert.equal(shiftFormulaRows('=마켓별상품코드!A6+$D$3+I21+$I22', 1), '=마켓별상품코드!A7+$D$3+I22+$I23');
});

test('mapRefs는 참조가 아닌 토큰을 그대로 둔다', () => {
  const f = '=IF(F25="Y","11%",IF(F25="N","11%","ERROR"))*1.5%';
  assert.equal(mapRefs(f, 'S', () => null), f);
});

test('추정: 브랜드·카테고리·기본 자리·파일명', () => {
  const info = { products: [
    { name: 'PNY 지포스 RTX 5080', cat: 'VGA', brand: 'PNY' },
    { name: '에이서 PREDATOR 라데온 RX 9070 XT', cat: 'VGA', brand: 'Acer' },
    { name: 'PNY 지포스 RTX 5060', cat: 'VGA', brand: 'PNY' },
    { name: 'Predator GM9000 2TB', cat: 'SSD', brand: 'Acer' },
  ] };
  assert.equal(guessBrand('에이서 NITRO 라데온 RX 9070 XT OC D6 16GB', info), 'Acer');
  assert.equal(guessCategory('에이서 NITRO 라데온 RX 9070 XT OC D6 16GB', info), 'VGA');
  assert.equal(guessCategory('KingBank KP270 4TB', info), 'SSD');
  assert.equal(guessCategory('ACER D5-6000 32GB', info), 'RAM');
  assert.equal(defaultAnchor(info, 'VGA', 'Acer'), 1);
  assert.equal(defaultAnchor(info, 'VGA', 'MSI'), 2);
  assert.equal(defaultAnchor(info, 'HDD', 'WD'), 3);
  assert.equal(outputName('261012_가격세팅.xlsx', new Date(2026, 9, 13)), '261013_가격세팅.xlsx');
  assert.equal(outputName('가격세팅.xlsx', new Date(2026, 9, 13)), '261013_가격세팅.xlsx');
});

// ---------------------------------------------------------------- 실제 파일(레포 밖)

const FIX = process.env.PNP_FIXTURE;
test('실제 가격세팅 파일: 넣은 뒤에도 기존 값이 그대로이고 표가 맞물린다', { skip: !FIX && 'PNP_FIXTURE 없음' }, async () => {
  const buf = fs.readFileSync(FIX);
  const before = await openBook(JSZip, buf);
  recalcBook(before);
  const info0 = locate(before);
  assert.deepEqual(info0.problems, []);
  const N = info0.products.length;

  const book = await openBook(JSZip, buf);
  const res = addProducts(book, [
    { name: '시험 신제품 A', cost: 1000000, cat: info0.products[0].cat, brand: info0.products[0].brand, price: 1300000, codes: {} },
    { name: '시험 신제품 B', cost: 500000, cat: info0.products[N - 1].cat, brand: '없는브랜드', price: null, codes: {}, afterName: '' },
  ]);
  assert.deepEqual(res.failed, []);
  const info = locate(book);
  assert.deepEqual(info.problems, []);
  assert.equal(info.products.length, N + 2);
  assert.equal(info.products[0].name, '시험 신제품 B');

  // 다시 열어도 같은 결과(저장·읽기 왕복)
  const again = await openBook(JSZip, await saveBook(book));
  recalcBook(again);
  assert.equal(locate(again).products.length, N + 2);

  // 기존 상품의 정산가세팅 수식 값은 그대로(자리만 이동)
  const ss = info0.settle.sheet;
  const sheetOf = (b) => b.sheets.find((s) => s.name === ss);
  for (const [k, b0] of info0.settle.blocks.entries()) {
    const b1 = info.settle.blocks[k];
    for (let i = 0; i < N; i++) {
      const j = info.products.findIndex((p) => p.name === info0.products[i].name);
      for (const c of [b0.cols.settle, b0.cols.margin, b0.cols.rate]) {
        const v0 = literalValue(sheetOf(before).model.rows.get(b0.start + i).cells.get(c), before.sst);
        const v1 = literalValue(sheetOf(again).model.rows.get(b1.start + j).cells.get(c), again.sst);
        assert.equal(v1, v0, `${b0.label} ${info0.products[i].name} ${numToCol(c)}열`);
      }
    }
  }
});
