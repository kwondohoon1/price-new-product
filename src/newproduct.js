// 가격세팅 워크북에 신제품 행을 넣는 업무 규칙.
// 네 시트의 상품 표는 같은 순서로 맞물려 있다(마켓별상품코드 n번째 = 정산가세팅 두 표의 n번째 =
// 최저가비교 n번째 = G마켓_프로모션 n번째). 그래서 '어느 상품 다음'을 하나 정하면 모든 표의 같은 자리에 넣는다.

import {
  sheetByName, insertRow, fillRowFrom, extendRangesToRow, recalcBook, literalValue, setCellValue, KEEP,
  colToNum, numToCol,
} from './book.js';

export const SHEET = { codes: '마켓별상품코드', settle: '정산가세팅', lowest: '최저가비교', promo: 'G마켓_프로모션' };

export const norm = (s) => String(s ?? '').replace(/\s+/g, ' ').trim();
const key = (s) => norm(s).toUpperCase();

function val(book, sh, r, c) {
  const row = sh.model.rows.get(r);
  return literalValue(row && row.cells.get(c), book.sst);
}
const text = (v) => (v == null || typeof v === 'object' ? '' : norm(v));
function hasContent(sh, r, c) {
  const row = sh.model.rows.get(r);
  const cell = row && row.cells.get(c);
  return !!cell && (cell.f != null || cell.v != null || !!cell.rest);
}
function rowTexts(book, sh, r) {
  const row = sh.model.rows.get(r);
  const out = new Map();
  if (!row) return out;
  for (const c of row.cells.keys()) out.set(c, text(val(book, sh, r, c)));
  return out;
}
const findCol = (texts, label, from = 1) => {
  for (const [c, t] of [...texts.entries()].sort((a, b) => a[0] - b[0])) if (c >= from && t === label) return c;
  return 0;
};
const maxRow = (sh) => Math.max(0, ...sh.model.rows.keys());

// ---------------------------------------------------------------- 표 찾기

/**
 * 워크북에서 네 표의 위치를 찾는다. 수식 캐시 값이 있어야 하므로 recalcBook 뒤에 부른다.
 * @returns {{ products, codes, settle, lowest, promo, problems: string[], warnings: string[] }}
 */
export function locate(book) {
  const problems = [];
  const warnings = [];
  const out = { products: [], codes: null, settle: null, lowest: null, promo: null, problems, warnings };

  // 마켓별상품코드 — 상품 목록의 기준
  const cs = sheetByName(book, SHEET.codes);
  if (!cs) { problems.push(`'${SHEET.codes}' 시트가 없습니다.`); return out; }
  const h = rowTexts(book, cs, 1);
  const cols = { cat: findCol(h, '카테고리'), brand: findCol(h, '브랜드'), name: findCol(h, '모델명'), rate: findCol(h, '등재가(%)') };
  if (!cols.cat || !cols.brand || !cols.name) { problems.push(`'${SHEET.codes}' 1행에서 카테고리·브랜드·모델명 머리글을 찾지 못했습니다.`); return out; }
  const afterCol = Math.max(findCol(h, '마진율'), findCol(h, '정산가'), cols.name);
  const codeCols = [...h.entries()].filter(([c, t]) => c > afterCol && t).sort((a, b) => a[0] - b[0]).map(([col, label]) => ({ col, label }));
  let r = 2;
  while (text(val(book, cs, r, cols.name))) {
    const codes = {};
    for (const { col, label } of codeCols) codes[label] = text(val(book, cs, r, col));
    out.products.push({ idx: r - 2, name: text(val(book, cs, r, cols.name)), cat: text(val(book, cs, r, cols.cat)), brand: text(val(book, cs, r, cols.brand)), codes });
    r++;
  }
  const N = out.products.length;
  if (!N) { problems.push(`'${SHEET.codes}'에 상품이 없습니다.`); return out; }
  out.codes = { sheet: cs.name, start: 2, cols, codeCols };
  const names = out.products.map((p) => key(p.name));

  // 정산가세팅 — '카테고리 … 모델명' 머리글 행마다 표 하나
  const ss = sheetByName(book, SHEET.settle);
  if (!ss) { problems.push(`'${SHEET.settle}' 시트가 없습니다.`); return out; }
  const blocks = [];
  for (let hr = 1; hr <= maxRow(ss); hr++) {
    const t = rowTexts(book, ss, hr);
    if (!findCol(t, '카테고리') || !findCol(t, '모델명')) continue;   // 앞에 '상태' 열이 있어도 됨
    const label = [...rowTexts(book, ss, hr - 1).entries()].sort((a, b) => a[0] - b[0]).map(([, v]) => v).find(Boolean) || `표 ${blocks.length + 1}`;
    const nameCol = findCol(t, '모델명');
    const bc = {
      name: nameCol, base: findCol(t, '기준'), input: findCol(t, '입력가'), cost: findCol(t, '원가'),
      settle: findCol(t, '정산가'), margin: findCol(t, '마진'), rate: findCol(t, '마진율'), maxBenefit: findCol(t, '최대혜택가'),
      shown: [...t.entries()].find(([, v]) => v.startsWith('노출가') && v.includes('2차'))?.[0] || 0,
    };
    // 마켓별 묶음: '등재가(세팅)' 열마다 바로 위 행에 마켓 이름
    const top = rowTexts(book, ss, hr - 1);
    const markets = [...t.entries()].filter(([, v]) => v === '등재가(세팅)').sort((a, b) => a[0] - b[0])
      .map(([c]) => {
        // 묶음: 등재가(세팅) · 노출가 · (최대혜택가) · 정산액 · 차이 — 정산액 자리는 머리글로
        let settle = c + 2;
        for (let k = c + 1; k <= c + 4; k++) if (t.get(k) === '정산액') { settle = k; break; }
        return { label: top.get(c) || numToCol(c), list: c, shown: c + 1, settle };
      });
    let n = 0;
    while (hasContent(ss, hr + 1 + n, nameCol) && text(val(book, ss, hr + 1 + n, nameCol))) n++;
    blocks.push({ label, header: hr, start: hr + 1, count: n, cols: bc, markets });
  }
  if (!blocks.length) { problems.push(`'${SHEET.settle}'에서 '카테고리 · 모델명' 머리글 행을 찾지 못했습니다.`); return out; }
  for (const b of blocks) {
    if (b.count !== N) { problems.push(`'${SHEET.settle}' ${b.label} 표의 상품 수(${b.count})가 '${SHEET.codes}'(${N})와 다릅니다.`); continue; }
    for (let i = 0; i < N; i++) {
      const nm = key(val(book, ss, b.start + i, b.cols.name));
      if (nm !== names[i]) { problems.push(`'${SHEET.settle}' ${b.label} ${b.start + i}행 상품명이 '${SHEET.codes}' ${i + 2}행과 다릅니다.`); break; }
    }
    for (const k of ['input', 'cost', 'base']) if (!b.cols[k]) problems.push(`'${SHEET.settle}' ${b.label} 표에 '${{ input: '입력가', cost: '원가', base: '기준' }[k]}' 머리글이 없습니다.`);
  }
  out.settle = { sheet: ss.name, blocks };

  // 최저가비교 — 다나와 동일스펙 사이트에서 받아 붙인 표. 순서가 맞을 때만 손댄다
  const ls = sheetByName(book, SHEET.lowest);
  if (ls) {
    let hr = 0;
    for (let i = 1; i <= Math.min(10, maxRow(ls)); i++) { const t = rowTexts(book, ls, i); if (t.get(1) === '다나와코드' && t.get(2) === '상품명') { hr = i; break; } }
    if (!hr) warnings.push(`'${SHEET.lowest}' 머리글(다나와코드·상품명)을 찾지 못해 이 시트는 건드리지 않습니다.`);
    else {
      const t = rowTexts(book, ls, hr);
      const all = (label) => [...t.entries()].filter(([, v]) => v === label).map(([c]) => c).sort((a, b) => a - b);
      const lc = {
        code: 1, name: 2, shown: all('노출가')[0] || 0, maxBenefit: all('최대혜택가')[0] || 0, cost: all('원가')[0] || 0,
        margin: all('마진')[0] || 0, rate: all('마진율')[0] || 0, propose: all('노출가')[1] || 0,
        filterUrl: findCol(t, '필터 URL'), verdict: findCol(t, '판정'),
      };
      let n = 0;
      while (text(val(book, ls, hr + 1 + n, 2))) n++;
      const ok = n === N && out.products.every((p, i) => key(val(book, ls, hr + 1 + i, 2)) === names[i]);
      if (ok) out.lowest = { sheet: ls.name, start: hr + 1, cols: lc };
      else warnings.push(`'${SHEET.lowest}' 상품 순서가 '${SHEET.codes}'와 달라 이 시트는 건드리지 않습니다. 동일스펙 사이트에서 다시 받아 붙이세요.`);
    }
  }

  // G마켓_프로모션 — 행사 취합표. 행 자리만 맞춘다
  const ps = sheetByName(book, SHEET.promo);
  if (ps) {
    let hr = 0;
    for (let i = 1; i <= Math.min(60, maxRow(ps)); i++) { const t = rowTexts(book, ps, i); if (t.get(1) === '업체명' && findCol(t, '모델명')) { hr = i; break; } }
    if (!hr) warnings.push(`'${SHEET.promo}' 머리글(업체명·모델명)을 찾지 못해 이 시트는 건드리지 않습니다.`);
    else {
      const t = rowTexts(book, ps, hr);
      const pc = { company: 1, brand: findCol(t, '브랜드'), name: findCol(t, '모델명'), shown: findCol(t, '노출가') };
      const start = hr + 1;
      let ok = true;
      for (let i = 0; i < N && ok; i++) { const nm = key(val(book, ps, start + i, pc.name)); if (nm && nm !== names[i]) ok = false; }
      for (let i = start + N; i <= maxRow(ps) && ok; i++) if (text(val(book, ps, i, pc.name))) ok = false;
      if (!key(val(book, ps, start + N - 1, pc.name)) && !key(val(book, ps, start, pc.name))) ok = false;
      if (ok) out.promo = { sheet: ps.name, start, cols: pc };
      else warnings.push(`'${SHEET.promo}' 상품 순서가 '${SHEET.codes}'와 달라 이 시트는 건드리지 않습니다.`);
    }
  }
  return out;
}

// ---------------------------------------------------------------- 추정·검사

const BRAND_ALIAS = [
  [/에이서|ACER|PREDATOR|프레데터|NITRO|니트로/i, 'Acer'],
  [/삼성/, '삼성전자'],
  [/킹뱅크|KINGBANK/i, 'KingBank'],
];

export function guessBrand(name, info) {
  const nm = key(name);
  const brands = [...new Set(info.products.map((p) => p.brand).filter(Boolean))];
  for (const b of brands.sort((a, c) => c.length - a.length)) if (nm.includes(key(b))) return b;
  for (const [re, b] of BRAND_ALIAS) if (re.test(name)) return brands.find((x) => key(x) === key(b)) || b;
  return norm(name).split(' ')[0] || '';
}

export function guessCategory(name, info) {
  const s = String(name || '');
  const cats = [...new Set(info.products.map((p) => p.cat).filter(Boolean))];
  const pick = (c) => cats.find((x) => key(x) === c) || c;
  if (/RTX|GTX|\bRX\s?\d|라데온|지포스|GeForce|Radeon|\bARC\b/i.test(s)) return pick('VGA');
  if (/DDR\d|UDIMM|SODIMM|\bRAM\b|D\d-\d{4}/i.test(s)) return pick('RAM');
  if (/SSD|NVMe|M\.2|GM\d{4}|\d\s?TB\b/i.test(s)) return pick('SSD');
  return '';
}

/** 넣을 자리 기본값: 같은 카테고리·브랜드의 마지막 상품 다음 → 없으면 같은 카테고리 마지막 → 없으면 맨 끝 */
export function defaultAnchor(info, cat, brand) {
  const ps = info.products;
  let i = -1;
  ps.forEach((p, k) => { if (key(p.cat) === key(cat) && key(p.brand) === key(brand)) i = k; });
  if (i < 0) ps.forEach((p, k) => { if (key(p.cat) === key(cat)) i = k; });
  return i < 0 ? ps.length - 1 : i;
}

export function validate(info, p, others = []) {
  const errs = [];
  if (!norm(p.name)) errs.push('상품명을 넣으세요.');
  else if (info.products.some((x) => key(x.name) === key(p.name))) errs.push('같은 상품명이 이미 있습니다.');
  else if (others.some((x) => x !== p && key(x.name) === key(p.name))) errs.push('같은 상품명을 두 번 넣었습니다.');
  if (!(p.cost > 0)) errs.push('원가를 숫자로 넣으세요.');
  if (!norm(p.cat)) errs.push('카테고리를 고르세요.');
  if (!norm(p.brand)) errs.push('브랜드를 넣으세요.');
  if (p.price != null && !(p.price > 0)) errs.push('노출가는 비우거나 숫자로 넣으세요.');
  if (p.jasaPrice != null && !(p.jasaPrice > 0)) errs.push('자사몰 노출가는 비우거나 숫자로 넣으세요.');
  for (const [label, v] of Object.entries(p.codes || {})) {
    if (!norm(v)) continue;
    const dup = info.products.find((x) => norm(x.codes[label]) === norm(v));
    if (dup) errs.push(`${label} 코드 ${norm(v)}는 이미 '${dup.name}'에 있습니다.`);
  }
  return errs;
}

// ---------------------------------------------------------------- 넣기

/**
 * 상품 하나를 모든 표의 같은 자리에 넣는다. book은 recalcBook·locate가 끝난 상태여야 한다.
 * @param p { name, cost, cat, brand, price?, jasaPrice?, codes?: {마켓: 코드}, after: 앞 상품 인덱스(-1이면 맨 위) }
 * @returns 바꾼 위치 목록 [{ sheet, row, note }]
 */
export function insertProduct(book, info, p) {
  const pos = p.after + 1;
  const changes = [];
  const { codes, settle, lowest, promo } = info;

  // 1) 빈 행 끼우기 — 같은 시트 안에서는 아래 표부터
  insertRow(book, codes.sheet, codes.start + pos);
  const blocks = [...settle.blocks].sort((a, b) => b.start - a.start);
  for (const b of blocks) insertRow(book, settle.sheet, b.start + pos);
  if (lowest) insertRow(book, lowest.sheet, lowest.start + pos);
  if (promo) insertRow(book, promo.sheet, promo.start + pos);

  const tplOf = (row) => (pos > 0 ? row - 1 : row + 1);

  // 2) 마켓별상품코드
  {
    const row = codes.start + pos;
    const v = { [codes.cols.cat]: norm(p.cat), [codes.cols.brand]: norm(p.brand), [codes.cols.name]: norm(p.name) };
    if (codes.cols.rate) v[codes.cols.rate] = KEEP;
    for (const { col, label } of codes.codeCols) v[col] = norm(p.codes?.[label]) || null;
    fillRowFrom(book, codes.sheet, tplOf(row), row, v);
    extendRangesToRow(book, codes.sheet, tplOf(row), row);
    changes.push({ sheet: codes.sheet, row, note: '상품명·카테고리·브랜드·마켓 코드, 수식은 윗행에서 이어받음' });
  }

  // 3) 정산가세팅 — 위 표부터 순서대로. 위 표에 한 행 넣었으니 아래 표 시작은 k행씩 밀림
  const settleRows = [];
  const sorted = [...settle.blocks].sort((a, b) => a.start - b.start);
  sorted.forEach((b, k) => {
    const row = b.start + k + pos;
    const tpl = tplOf(row);
    const tplCost = book.sheets.find((s) => s.name === settle.sheet).model.rows.get(tpl)?.cells.get(b.cols.cost);
    const price = k === 0 ? p.price : (p.jasaPrice ?? p.price);
    const v = { [b.cols.base]: KEEP, [b.cols.input]: price ?? null };
    if (!(tplCost && tplCost.f)) v[b.cols.cost] = p.cost;
    fillRowFrom(book, settle.sheet, tpl, row, v);
    extendRangesToRow(book, settle.sheet, tpl, row);
    settleRows.push({ block: b, row });
    changes.push({ sheet: settle.sheet, row, note: `${b.label} — 입력가 ${price != null ? price.toLocaleString('ko-KR') : '비움'}, 원가 ${tplCost && tplCost.f ? '위 표 연결' : p.cost.toLocaleString('ko-KR')}` });
  });

  // 4) 최저가비교·프로모션은 정산가세팅 계산 결과가 필요하다 → 우선 자리만 만들고 값은 계산 뒤에
  if (lowest) {
    const row = lowest.start + pos;
    const lc = lowest.cols;
    const v = { [lc.code]: null, [lc.name]: norm(p.name), [lc.cost]: p.cost };
    if (lc.verdict) v[lc.verdict] = '비교 불가';
    if (lc.filterUrl) v[lc.filterUrl] = categoryListUrl(book, info, p.cat, lowest) || null;
    fillRowFrom(book, lowest.sheet, tplOf(row), row, v);
    extendRangesToRow(book, lowest.sheet, tplOf(row), row);
    changes.push({ sheet: lowest.sheet, row, note: '현재세팅 값·제안=현재, 다나와코드는 비움(다음에 사이트에서 받으면 채워짐)' });
  }
  if (promo) {
    const row = promo.start + pos;
    const pc = promo.cols;
    const v = { [pc.company]: KEEP, [pc.name]: norm(p.name) };
    if (pc.brand) v[pc.brand] = norm(p.brand);
    fillRowFrom(book, promo.sheet, tplOf(row), row, v, { copyFormulas: false });
    extendRangesToRow(book, promo.sheet, tplOf(row), row);
    changes.push({ sheet: promo.sheet, row, note: `행 자리만 맞춤(행사 신청 대상 아님) — 상품명${p.price != null ? '·노출가' : ''}만 적음` });
  }

  // 5) 계산해서 최저가비교·프로모션 값 채우기
  const calc = recalcBook(book);
  const g = settleRows[0];
  const ssName = settle.sheet;
  const num = (c) => { if (!c) return null; const x = calc.get(ssName, c, g.row); return typeof x === 'number' ? x : null; };
  if (lowest) {
    const row = lowest.start + pos;
    const lc = lowest.cols;
    const shown = num(g.block.cols.shown);
    const set = (c, x) => { if (c) setCellValue(book, lowest.sheet, row, c, x); };
    set(lc.shown, shown);
    set(lc.maxBenefit, num(g.block.cols.maxBenefit));
    set(lc.margin, p.price != null ? num(g.block.cols.margin) : null);
    set(lc.rate, p.price != null ? num(g.block.cols.rate) : null);
    set(lc.propose, shown);
  }
  if (promo && promo.cols.shown) {
    const gm = g.block.markets[0];
    setCellValue(book, promo.sheet, promo.start + pos, promo.cols.shown, gm ? num(gm.shown) : null);
  }
  return changes;
}

// 같은 카테고리 상품의 다나와 필터 URL에서 목록 주소(…/list/?cate=번호)만 뽑는다
function categoryListUrl(book, info, cat, lowest) {
  const ls = sheetByName(book, lowest.sheet);
  for (let i = info.products.length - 1; i >= 0; i--) {
    if (key(info.products[i].cat) !== key(cat)) continue;
    const u = text(val(book, ls, lowest.start + i, lowest.cols.filterUrl));
    const m = /^https?:\/\/prod\.danawa\.com\/list\/\?cate=\d+/.exec(u);
    if (m) return m[0];
  }
  return '';
}

/**
 * 여러 상품을 차례로 넣는다. 각 상품의 after는 '앞 상품 이름'(없으면 기본 자리).
 * @returns {{ changes, preview: Array<{name, blocks:[...]}>, failed }}
 */
export function addProducts(book, list) {
  let calc = recalcBook(book);
  let info = locate(book);
  if (info.problems.length) throw new Error(info.problems.join('\n'));
  const changes = [];
  for (const p of list) {
    const errs = validate(info, p);
    if (errs.length) throw new Error(`${norm(p.name) || '상품'}: ${errs.join(' ')}`);
    let after;
    if (p.afterName === '') after = -1;
    else if (p.afterName) {
      after = info.products.findIndex((x) => key(x.name) === key(p.afterName));
      if (after < 0) throw new Error(`넣을 자리의 앞 상품 '${p.afterName}'을 찾지 못했습니다.`);
    } else after = defaultAnchor(info, p.cat, p.brand);
    changes.push(...insertProduct(book, info, { ...p, after }));
    calc = recalcBook(book);
    info = locate(book);
    if (info.problems.length) throw new Error(`넣은 뒤 표가 어긋났습니다: ${info.problems.join(' ')}`);
  }
  const preview = list.map((p) => previewOf(book, info, calc, p.name));
  return { changes, preview, failed: calc.failed, info };
}

/** 넣은 상품의 정산가세팅 계산 결과 */
export function previewOf(book, info, calc, name) {
  const i = info.products.findIndex((x) => key(x.name) === key(name));
  if (i < 0) return null;
  const ss = info.settle.sheet;
  const blocks = info.settle.blocks.map((b) => {
    const r = b.start + i;
    const g = (c) => (c ? calc.get(ss, c, r) : null);
    return {
      label: b.label, row: r,
      input: g(b.cols.input), shown: g(b.cols.shown), maxBenefit: g(b.cols.maxBenefit),
      settle: g(b.cols.settle), cost: g(b.cols.cost), margin: g(b.cols.margin), rate: g(b.cols.rate),
      markets: b.markets.map((m) => ({ label: m.label, list: g(m.list), shown: g(m.shown), settle: g(m.settle) })),
    };
  });
  return { name, index: i, blocks };
}

/** 내려받을 파일 이름: 앞 6자리 날짜(YYMMDD_)를 오늘로 바꾼다 */
export function outputName(original, today = new Date()) {
  const ymd = String(today.getFullYear()).slice(2) + String(today.getMonth() + 1).padStart(2, '0') + String(today.getDate()).padStart(2, '0');
  const base = String(original || '가격세팅.xlsx').replace(/\.xlsx$/i, '');
  if (/^\d{6}_/.test(base)) return `${ymd}${base.slice(6)}.xlsx`;
  return `${ymd}_${base}.xlsx`;
}

export { colToNum, numToCol };
