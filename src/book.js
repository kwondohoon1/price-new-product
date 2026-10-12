// 가격세팅 워크북을 XML 그대로 열고, 행을 끼우고, 다시 저장한다.
// 원본 파트는 통째로 다시 쓰지 않는다. 고친 시트·워크북·메모만 바꿔 서식·메모·하이퍼링크를 보존한다.

import {
  parseSheet, serializeSheet, expandSharedFormulas, parseSharedStrings,
  parseAttrs, getAttr, setAttr, attrsToStr, colToNum, numToCol, splitRef, esc, unesc,
} from './xml.js';
import { parse, evaluate, isErr, isUnsupported } from './formula.js';

const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

// ---------------------------------------------------------------- 열기·저장

export async function openBook(JSZip, data) {
  const zip = await JSZip.loadAsync(data);
  const read = async (p) => { const f = zip.file(p); return f ? f.async('string') : null; };
  const workbookXml = await read('xl/workbook.xml');
  if (!workbookXml) throw new Error('엑셀(xlsx) 파일이 아닙니다.');
  const wbRels = await read('xl/_rels/workbook.xml.rels');
  const relTargets = relMap(wbRels);
  const sheets = [];
  const sheetRe = /<sheet\b([^>]*?)\/>/g;
  let m;
  while ((m = sheetRe.exec(workbookXml))) {
    const a = parseAttrs(m[1]);
    const name = unesc(getAttr(a, 'name'));
    const rid = getAttr(a, 'r:id');
    const path = resolvePath('xl/', relTargets.get(rid));
    const xml = await read(path);
    const relsPath = path.replace(/([^/]+)$/, '_rels/$1.rels');
    const rels = await read(relsPath);
    const model = parseSheet(xml);
    expandSharedFormulas(model);
    const sheet = { name, path, relsPath, model, comments: null };
    if (rels) {
      for (const [, r] of relEntries(rels)) {
        if (r.type.endsWith('/comments')) sheet.commentsPath = resolvePath(dirOf(path), r.target);
        if (r.type.endsWith('/vmlDrawing')) sheet.vmlPath = resolvePath(dirOf(path), r.target);
      }
    }
    if (sheet.commentsPath) sheet.commentsXml = await read(sheet.commentsPath);
    if (sheet.vmlPath) sheet.vmlXml = await read(sheet.vmlPath);
    sheets.push(sheet);
  }
  const sst = parseSharedStrings(await read('xl/sharedStrings.xml'));
  return { zip, workbookXml, sheets, sst };
}

export async function saveBook(book) {
  const { zip } = book;
  for (const sh of book.sheets) {
    sh.model.head = fixDimension(sh.model);
    zip.file(sh.path, serializeSheet(sh.model));
    if (sh.commentsXml != null) zip.file(sh.commentsPath, sh.commentsXml);
    if (sh.vmlXml != null) zip.file(sh.vmlPath, sh.vmlXml);
  }
  // 계산 체인은 셀 위치가 바뀌면 틀어지므로 지운다. Excel이 열 때 다시 만든다.
  let wb = book.workbookXml;
  if (zip.file('xl/calcChain.xml')) {
    zip.remove('xl/calcChain.xml');
    const relsP = 'xl/_rels/workbook.xml.rels';
    const rels = await zip.file(relsP).async('string');
    zip.file(relsP, rels.replace(/<Relationship\b[^>]*calcChain[^>]*\/>/g, ''));
    const ct = await zip.file('[Content_Types].xml').async('string');
    zip.file('[Content_Types].xml', ct.replace(/<Override\b[^>]*calcChain[^>]*\/>/g, ''));
  }
  // 열 때 전체 다시 계산하게 한다(우리가 넣은 캐시 값과 Excel 계산이 어긋나도 Excel 쪽이 이긴다)
  wb = /<calcPr\b/.test(wb)
    ? wb.replace(/<calcPr\b([^>]*?)\/>/, (s, a) => {
      const at = parseAttrs(a); setAttr(at, 'fullCalcOnLoad', '1'); return `<calcPr${attrsToStr(at)}/>`;
    })
    : wb.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
  zip.file('xl/workbook.xml', wb);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
}

function relMap(xml) {
  const out = new Map();
  for (const [id, r] of relEntries(xml)) out.set(id, r.target);
  return out;
}
function relEntries(xml) {
  const out = [];
  const re = /<Relationship\b([^>]*?)\/>/g;
  let m;
  while ((m = re.exec(xml || ''))) {
    const a = parseAttrs(m[1]);
    out.push([getAttr(a, 'Id'), { type: getAttr(a, 'Type'), target: getAttr(a, 'Target') }]);
  }
  return out;
}
const dirOf = (p) => p.replace(/[^/]+$/, '');
function resolvePath(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = (base + target).split('/');
  const out = [];
  for (const p of parts) { if (p === '..') out.pop(); else if (p !== '.' && p !== '') out.push(p); }
  return out.join('/');
}
void REL_NS;

export function sheetByName(book, name) {
  const k = String(name).toUpperCase();
  return book.sheets.find((s) => s.name.toUpperCase() === k) || null;
}

// ---------------------------------------------------------------- 참조 훑기

// 수식 안의 셀/범위 참조를 하나씩 fn에 넘기고, fn이 돌려준 행 번호로 바꿔 쓴다.
// 문자열 리터럴은 건드리지 않는다. fn(sheet, a, b) → { r1, r2 } | null  (b는 범위 끝, 단일 참조면 null)
const BREAK = new Set([...'+-*/^&=<>,;:%(){} \n\r\t!"\'']);
const REF_TOKEN = /^(\$?)([A-Z]{1,3})(\$?)(\d+)$/;

export function mapRefs(formula, homeSheet, fn) {
  const s = formula;
  let out = '';
  let i = 0;
  while (i < s.length) {
    const ch = s[i];
    if (ch === '"') {
      let j = i + 1;
      while (j < s.length) { if (s[j] === '"') { if (s[j + 1] === '"') { j += 2; continue; } break; } j++; }
      out += s.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    let sheet = homeSheet;
    let prefix = '';
    let j = i;
    if (ch === "'") {
      let k = i + 1;
      while (k < s.length) { if (s[k] === "'") { if (s[k + 1] === "'") { k += 2; continue; } break; } k++; }
      if (s[k + 1] !== '!') { out += s.slice(i, k + 1); i = k + 1; continue; }
      sheet = s.slice(i + 1, k).replace(/''/g, "'");
      prefix = s.slice(i, k + 2);
      j = k + 2;
    } else if (BREAK.has(ch)) {
      out += ch; i++; continue;
    } else {
      let k = i;
      while (k < s.length && !BREAK.has(s[k])) k++;
      if (s[k] === '!') { sheet = s.slice(i, k); prefix = s.slice(i, k + 1); j = k + 1; }
    }
    // j부터 참조 토큰
    let k = j;
    while (k < s.length && !BREAK.has(s[k])) k++;
    const tok = s.slice(j, k);
    const a = REF_TOKEN.exec(tok);
    if (!a || s[k] === '(') { out += s.slice(i, k); i = Math.max(k, i + 1); continue; }
    let end = k;
    let b = null;
    if (s[k] === ':') {
      let k2 = k + 1;
      while (k2 < s.length && !BREAK.has(s[k2])) k2++;
      const bm = REF_TOKEN.exec(s.slice(k + 1, k2));
      if (bm) { b = bm; end = k2; }
    }
    const ra = { absC: !!a[1], col: a[2], absR: !!a[3], row: +a[4] };
    const rb = b ? { absC: !!b[1], col: b[2], absR: !!b[3], row: +b[4] } : null;
    const res = fn(sheet, ra, rb);
    if (!res) { out += s.slice(i, end); i = end; continue; }
    const w = (r, row) => `${r.absC ? '$' : ''}${r.col}${r.absR ? '$' : ''}${row}`;
    out += prefix + w(ra, res.r1) + (rb ? ':' + w(rb, res.r2) : '');
    i = end;
  }
  return out;
}

// 행 at 위에 n행을 끼울 때 Excel과 같은 규칙으로 행 번호를 옮긴다.
// 단일 참조: at 이상이면 +n. 범위: 시작이 at 이상이면 통째로, 시작<at≤끝이면 끝만 늘린다.
export function shiftRange(r1, r2, at, n) {
  if (r2 == null) return r1 >= at ? r1 + n : r1;
  const lo = Math.min(r1, r2), hi = Math.max(r1, r2);
  let nlo = lo, nhi = hi;
  if (lo >= at) { nlo += n; nhi += n; } else if (hi >= at) nhi += n;
  return r1 <= r2 ? [nlo, nhi] : [nhi, nlo];
}

export function shiftFormulaForInsert(formula, homeSheet, target, at, n) {
  const T = target.toUpperCase();
  return mapRefs(formula, homeSheet, (sheet, a, b) => {
    if (!sheet || sheet.toUpperCase() !== T) return null;
    if (!b) { const r = shiftRange(a.row, null, at, n); return r === a.row ? null : { r1: r }; }
    const [r1, r2] = shiftRange(a.row, b.row, at, n);
    return r1 === a.row && r2 === b.row ? null : { r1, r2 };
  });
}

// "A1:B2 C5" 같은 sqref 문자열을 옮긴다
export function shiftSqref(sqref, at, n) {
  return sqref.split(/\s+/).filter(Boolean).map((part) => {
    const [x, y] = part.split(':');
    const a = splitRef(x.replace(/\$/g, ''));
    if (!y) return numToCol(a.col) + shiftRange(a.row, null, at, n);
    const b = splitRef(y.replace(/\$/g, ''));
    const [r1, r2] = shiftRange(a.row, b.row, at, n);
    return `${numToCol(a.col)}${r1}:${numToCol(b.col)}${r2}`;
  }).join(' ');
}

// ---------------------------------------------------------------- 행 끼우기

/**
 * sheetName의 at행 위에 빈 행 하나를 끼운다(at행부터 아래가 한 칸씩 내려감).
 * 모든 시트의 수식, 이 시트의 병합·유효성 검사·조건부 서식·하이퍼링크·필터·창 고정·선택,
 * 워크북 이름 정의, 메모 위치를 함께 옮긴다.
 */
export function insertRow(book, sheetName, at) {
  const target = sheetByName(book, sheetName);
  if (!target) throw new Error(`시트가 없습니다: ${sheetName}`);
  const n = 1;
  for (const sh of book.sheets) {
    for (const row of sh.model.rows.values()) {
      for (const cell of row.cells.values()) {
        if (!cell.f || cell.f.text == null) continue;
        cell.f.text = shiftFormulaForInsert(cell.f.text, sh.name, target.name, at, n);
        const ref = getAttr(cell.f.attrs, 'ref');
        if (ref && sh === target) setAttr(cell.f.attrs, 'ref', shiftSqref(ref, at, n));
      }
    }
  }
  // 행 번호
  const m = target.model;
  const rows = new Map();
  for (const [r, row] of [...m.rows.entries()].sort((x, y) => x[0] - y[0])) {
    const nr = r >= at ? r + n : r;
    if (nr !== r) {
      setAttr(row.attrs, 'r', String(nr));
      for (const [c, cell] of row.cells) setAttr(cell.attrs, 'r', numToCol(c) + nr);
    }
    rows.set(nr, row);
  }
  m.rows = rows;
  m.head = shiftSheetXmlPart(m.head, target.name, at, n);
  m.tail = shiftSheetXmlPart(m.tail, target.name, at, n);
  // 이름 정의
  book.workbookXml = book.workbookXml.replace(/(<definedName\b[^>]*>)([\s\S]*?)(<\/definedName>)/g, (s, o, body, c) =>
    o + esc(shiftFormulaForInsert(unesc(body), null, target.name, at, n)) + c);
  // 메모
  if (target.commentsXml) {
    target.commentsXml = target.commentsXml.replace(/(<comment\b[^>]*\bref=")([A-Z]+\d+)(")/g, (s, a, ref, z) => a + shiftSqref(ref, at, n) + z);
  }
  if (target.vmlXml) {
    // x:Row는 0부터 센다. 메모가 붙은 행이 at 이상이면 내린다
    target.vmlXml = target.vmlXml.replace(/<x:ClientData\b[\s\S]*?<\/x:ClientData>/g, (blk) => {
      const rm = /<x:Row>(\d+)<\/x:Row>/.exec(blk);
      if (!rm || +rm[1] + 1 < at) return blk;
      return blk
        .replace(/<x:Row>(\d+)<\/x:Row>/, (s, v) => `<x:Row>${+v + n}</x:Row>`)
        .replace(/<x:Anchor>([^<]*)<\/x:Anchor>/, (s, v) => {
          const p = v.split(',').map((x) => x.trim());
          if (p.length === 8) { p[2] = String(+p[2] + n); p[6] = String(+p[6] + n); }
          return `<x:Anchor>${p.join(', ')}</x:Anchor>`;
        });
    });
  }
}

// 시트 XML의 sheetData 바깥 부분(머리·꼬리)에서 셀 위치를 담은 속성을 옮긴다
function shiftSheetXmlPart(xml, sheetName, at, n) {
  return xml
    .replace(/(<(?:mergeCell|hyperlink|autoFilter|conditionalFormatting|dataValidation|selection|protectedRange|ignoredError)\b[^>]*?\b(?:ref|sqref)=")([^"]+)(")/g,
      (s, a, v, z) => a + shiftSqref(v, at, n) + z)
    .replace(/(<selection\b[^>]*?\bactiveCell=")([^"]+)(")/g, (s, a, v, z) => a + shiftSqref(v, at, n) + z)
    // 창 고정의 topLeftCell은 스크롤 위치라서 Excel도 옮기지 않는다
    // 조건부 서식·유효성 검사 수식 안의 참조
    .replace(/(<(formula|formula1|formula2)>)([\s\S]*?)(<\/\2>)/g, (s, o, tag, body, c) =>
      o + esc(shiftFormulaForInsert(unesc(body), sheetName, sheetName, at, n)) + c);
}

// 끼운 행이 표의 맨 아래(또는 맨 위)에 붙으면 Excel 규칙상 범위가 늘지 않는다.
// 바로 옆 본보기 행을 덮는 유효성 검사·조건부 서식·필터 범위를 새 행까지 늘린다.
export function extendRangesToRow(book, sheetName, tplRow, newRow) {
  const sh = sheetByName(book, sheetName);
  const grow = (sqref) => sqref.split(/\s+/).filter(Boolean).map((part) => {
    const [x, y] = part.split(':');
    const a = splitRef(x.replace(/\$/g, ''));
    const b = y ? splitRef(y.replace(/\$/g, '')) : a;
    let lo = Math.min(a.row, b.row), hi = Math.max(a.row, b.row);
    if (tplRow >= lo && tplRow <= hi && (newRow === hi + 1 || newRow === lo - 1)) {
      lo = Math.min(lo, newRow); hi = Math.max(hi, newRow);
    }
    const ca = numToCol(a.col), cb = numToCol(b.col);
    return lo === hi && ca === cb ? `${ca}${lo}` : `${ca}${lo}:${cb}${hi}`;
  }).join(' ');
  sh.model.tail = sh.model.tail.replace(/(<(?:autoFilter|conditionalFormatting|dataValidation)\b[^>]*?\b(?:ref|sqref)=")([^"]+)(")/g,
    (s, a, v, z) => a + grow(v) + z);
  const q = sh.name.replace(/'/g, "''");
  book.workbookXml = book.workbookXml.replace(/(<definedName\b[^>]*name="_xlnm\._FilterDatabase"[^>]*>)([\s\S]*?)(<\/definedName>)/g, (s, o, body, c) => {
    const t = unesc(body);
    const mm = /^(?:'?(.+?)'?)!(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)$/.exec(t);
    if (!mm || mm[1].replace(/''/g, "'").toUpperCase() !== sh.name.toUpperCase()) return s;
    const g = grow(mm[2].replace(/\$/g, '')).split(':').map((r) => { const p = splitRef(r); return `$${numToCol(p.col)}$${p.row}`; }).join(':');
    void q;
    return o + esc(`${mm[1]}!${g}`) + c;
  });
}

function fixDimension(model) {
  let minR = Infinity, maxR = 0, minC = Infinity, maxC = 0;
  for (const [r, row] of model.rows) {
    for (const c of row.cells.keys()) {
      minR = Math.min(minR, r); maxR = Math.max(maxR, r); minC = Math.min(minC, c); maxC = Math.max(maxC, c);
    }
  }
  if (!maxR) return model.head;
  const ref = `${numToCol(minC)}${minR}:${numToCol(maxC)}${maxR}`;
  return model.head.replace(/<dimension ref="[^"]*"\/>/, `<dimension ref="${ref}"/>`);
}

// ---------------------------------------------------------------- 행 복제

/**
 * 본보기 행(tplRow)의 서식으로 newRow를 채운다.
 * - 본보기의 수식은 행 차이만큼 옮겨 넣는다(copyFormulas가 false면 넣지 않음)
 * - values[열번호]가 있으면 그 값을 쓴다: number | string | null(비움) | KEEP(본보기 값 그대로)
 * - 그 밖의 본보기 값(숫자·글자)은 비운다
 */
export const KEEP = Symbol('keep');
export function fillRowFrom(book, sheetName, tplRow, newRow, values = {}, { copyFormulas = true } = {}) {
  const sh = sheetByName(book, sheetName);
  const m = sh.model;
  const tpl = m.rows.get(tplRow);
  const row = { attrs: [['r', String(newRow)]], cells: new Map() };
  if (tpl) for (const [k, v] of tpl.attrs) if (k !== 'r') row.attrs.push([k, v]);
  m.rows.set(newRow, row);
  const cols = new Set([...(tpl ? tpl.cells.keys() : []), ...Object.keys(values).map(Number)]);
  for (const c of [...cols].sort((a, b) => a - b)) {
    const t = tpl && tpl.cells.get(c);
    const cell = { attrs: [['r', numToCol(c) + newRow]], f: null, v: null, rest: '' };
    const s = t && getAttr(t.attrs, 's');
    if (s) cell.attrs.push(['s', s]);
    const has = Object.prototype.hasOwnProperty.call(values, c);
    const val = has ? values[c] : undefined;
    if (has && val === KEEP && t) {
      for (const [k, v] of t.attrs) if (k === 't') cell.attrs.push([k, v]);
      cell.v = t.v; cell.rest = t.rest;
      if (t.f && copyFormulas) cell.f = { attrs: [], text: shiftFormulaRows(t.f.text, newRow - tplRow) };
    } else if (has && val != null && val !== KEEP) {
      if (typeof val === 'number') cell.v = String(val);
      else {
        const txt = String(val);
        const sp = /^\s|\s$/.test(txt) ? ' xml:space="preserve"' : '';
        cell.rest = `<is><t${sp}>${esc(txt)}</t></is>`;
        cell.attrs.push(['t', 'inlineStr']);
      }
    } else if (!has && t && t.f && t.f.text != null && copyFormulas) {
      cell.f = { attrs: [], text: shiftFormulaRows(t.f.text, newRow - tplRow) };
    }
    row.cells.set(c, cell);
  }
  return row;
}

// 상대 행 참조만 dr만큼 옮긴다(절대 행·열은 그대로)
export function shiftFormulaRows(formula, dr) {
  return mapRefs(formula, null, (sheet, a, b) => {
    const r1 = a.absR ? a.row : a.row + dr;
    if (!b) return { r1 };
    return { r1, r2: b.absR ? b.row : b.row + dr };
  });
}

export function setCellValue(book, sheetName, row, col, val) {
  const sh = sheetByName(book, sheetName);
  const r = sh.model.rows.get(row);
  const cell = r && r.cells.get(col);
  if (!cell) return;
  cell.f = null; cell.v = null; cell.rest = '';
  setAttr(cell.attrs, 't', null);
  if (val == null) return;
  if (typeof val === 'number') cell.v = String(val);
  else { cell.rest = `<is><t>${esc(String(val))}</t></is>`; setAttr(cell.attrs, 't', 'inlineStr'); }
}

// ---------------------------------------------------------------- 값 읽기·다시 계산

export function literalValue(cell, sst) {
  if (!cell) return null;
  const t = getAttr(cell.attrs, 't');
  if (t === 's') return cell.v == null ? null : sst[+cell.v];
  if (t === 'inlineStr') { const mm = /<t\b[^>]*>([\s\S]*?)<\/t>/.exec(cell.rest); return mm ? unesc(mm[1]) : ''; }
  if (t === 'str') return cell.v == null ? '' : unesc(cell.v);
  if (t === 'b') return cell.v === '1';
  if (t === 'e') return { err: cell.v };
  if (cell.v == null) return null;
  const x = Number(cell.v);
  return Number.isFinite(x) ? x : cell.v;
}

/**
 * 모든 수식을 계산해 캐시 값(<v>)을 새로 쓴다. 계산하지 못한 셀은 캐시를 비워 Excel에 맡긴다.
 * @returns {{ get(sheet, col, row), failed: Array<{sheet, addr, why}> }}
 */
export function recalcBook(book) {
  const memo = new Map();
  const visiting = new Set();
  const astCache = new Map();
  const failed = [];
  const byName = new Map(book.sheets.map((s) => [s.name.toUpperCase(), s]));
  const astOf = (text) => {
    if (!astCache.has(text)) { try { astCache.set(text, parse(text)); } catch (e) { astCache.set(text, { parseError: e.message }); } }
    return astCache.get(text);
  };
  function get(sheet, col, row) {
    const sh = byName.get(String(sheet).toUpperCase());
    if (!sh) return { err: '#REF!' };
    const r = sh.model.rows.get(row);
    const cell = r && r.cells.get(col);
    if (!cell) return null;
    if (!cell.f || cell.f.text == null) return literalValue(cell, book.sst);
    const k = `${sh.name}!${col},${row}`;
    if (memo.has(k)) return memo.get(k);
    if (visiting.has(k)) return { err: '#CYCLE!' };
    visiting.add(k);
    const ast = astOf(cell.f.text);
    let v = ast.parseError ? { parseError: ast.parseError } : evaluate(ast, get, sh.name);
    if (v && typeof v === 'object' && 'range' in v) v = v.range[0] ?? null;
    visiting.delete(k);
    memo.set(k, v);
    return v;
  }
  for (const sh of book.sheets) {
    for (const [r, row] of sh.model.rows) {
      for (const [c, cell] of row.cells) {
        if (!cell.f || cell.f.text == null) continue;
        const v = get(sh.name, c, r);
        writeCache(cell, v);
        if (v == null || typeof v !== 'object') continue;
        if (isUnsupported(v) || v.parseError || (isErr(v) && v.err === '#CYCLE!')) {
          failed.push({ sheet: sh.name, addr: numToCol(c) + r, why: v.parseError || v.fn || v.err || '지원하지 않는 함수' });
        }
      }
    }
  }
  return { get, failed };
}

function writeCache(cell, v) {
  cell.rest = '';
  setAttr(cell.attrs, 't', null);
  if (v == null) { cell.v = null; return; }
  if (typeof v === 'number') { cell.v = Object.is(v, -0) ? '0' : String(v); return; }
  if (typeof v === 'string') { setAttr(cell.attrs, 't', 'str'); cell.v = esc(v); return; }
  if (typeof v === 'boolean') { setAttr(cell.attrs, 't', 'b'); cell.v = v ? '1' : '0'; return; }
  if (isErr(v) && v.err !== '#CYCLE!') { setAttr(cell.attrs, 't', 'e'); cell.v = v.err; return; }
  cell.v = null; // 계산 못 함 → Excel이 계산
}

export { colToNum, numToCol };
