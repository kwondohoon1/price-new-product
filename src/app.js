// 화면 동작. 엑셀 처리는 book.js·newproduct.js가 하고, 여기서는 입력을 모으고 결과를 보여준다.
// 아무것도 저장하지 않는다(가격·원가는 민감 정보).
import { openBook, saveBook, recalcBook } from './book.js';
import { locate, addProducts, validate, guessBrand, guessCategory, defaultAnchor, outputName, norm } from './newproduct.js';

const JSZip = window.JSZip;
const $ = (s, root = document) => root.querySelector(s);

const state = {
  file: null,          // { name, buf }
  info: null,          // locate() 결과(원본 기준)
  items: [],           // 입력 카드 [{ el, touched:Set }]
  downloadUrl: null,
  previewSeq: 0,
};

const fmt = (n) => (typeof n === 'number' ? Math.round(n).toLocaleString('ko-KR') : '');
const pct = (n) => (typeof n === 'number' ? (n * 100).toFixed(2) + '%' : '');
const parseNum = (s) => {
  const t = String(s ?? '').replace(/[^\d.]/g, '');
  return t ? Number(t) : null;
};
const el = (tag, attrs = {}, ...kids) => {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === 'class') e.className = v;
    else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
    else if (v === true) e.setAttribute(k, '');
    else if (v !== false && v != null) e.setAttribute(k, v);
  }
  for (const k of kids.flat()) if (k != null) e.append(k.nodeType ? k : document.createTextNode(String(k)));
  return e;
};
function setMsg(text, cls = '') { const m = $('#file-msg'); m.textContent = text; m.className = 'msg ' + cls; }

// ---------------------------------------------------------------- 파일

async function takeFile(f) {
  if (!f) return;
  if (!/\.xlsx$/i.test(f.name)) { setMsg(`${f.name}: xlsx 파일만 됩니다.`, 'bad'); return; }
  setMsg('파일을 읽는 중…', 'busy');
  try {
    const buf = new Uint8Array(await f.arrayBuffer());
    const book = await openBook(JSZip, buf);
    recalcBook(book);
    const info = locate(book);
    state.file = { name: f.name, buf };
    state.info = info;
    renderFileInfo();
    if (info.problems.length) {
      setMsg('이 파일에는 신제품을 넣을 수 없습니다. 아래 내용을 확인하세요.', 'bad');
      $('#step-input').hidden = true; $('#step-build').hidden = true;
      return;
    }
    setMsg(`${f.name} — 상품 ${info.products.length}개를 읽었습니다.`);
    fillLists();
    if (!state.items.length) addItem(); else state.items.forEach((it) => refreshAnchor(it));
    $('#step-input').hidden = false; $('#step-build').hidden = false;
    $('#name-preview').textContent = outputName(f.name);
    resetDownload();
    schedulePreview();
  } catch (e) {
    console.error(e);
    setMsg(`읽지 못했습니다: ${e.message}`, 'bad');
  }
}

function renderFileInfo() {
  const { info } = state;
  $('#file-info').hidden = false;
  const stat = (k, v, s, long) => el('div', { class: 'stat' }, el('div', { class: 'k' }, k), el('div', { class: 'v' + (long ? ' long' : '') }, v), s ? el('div', { class: 's' }, s) : null);
  const cats = [...new Set(info.products.map((p) => p.cat))];
  $('#file-stats').replaceChildren(
    stat('상품', String(info.products.length), cats.map((c) => `${c} ${info.products.filter((p) => p.cat === c).length}`).join(' · ')),
    stat('정산가세팅', info.settle ? info.settle.blocks.map((b) => b.label).join(' · ') : '없음', '표마다 같은 자리에 넣음', true),
    stat('최저가비교', info.lowest ? '함께 넣음' : '건드리지 않음', info.lowest ? '현재세팅 값 · 판정 "비교 불가"' : '순서가 맞지 않거나 시트 없음', true),
    stat('G마켓_프로모션', info.promo ? '자리만 맞춤' : '건드리지 않음', info.promo ? '상품명·노출가만, 행사 수식은 넣지 않음' : '순서가 맞지 않거나 시트 없음', true),
  );
  const notes = [];
  if (info.problems.length) notes.push(el('div', { class: 'note bad' }, '넣을 수 없음', el('ul', {}, info.problems.map((t) => el('li', {}, t)))));
  if (info.warnings.length) notes.push(el('div', { class: 'note warn' }, '참고', el('ul', {}, info.warnings.map((t) => el('li', {}, t)))));
  $('#file-notes').replaceChildren(...notes);
  const t = $('#product-table');
  t.replaceChildren(
    el('thead', {}, el('tr', {}, el('th', { class: 'n' }, '#'), el('th', {}, '카테고리'), el('th', {}, '브랜드'), el('th', {}, '상품명'), el('th', {}, '자사몰 코드'))),
    el('tbody', {}, info.products.map((p, i) => el('tr', {}, el('td', { class: 'n muted' }, i + 1), el('td', {}, p.cat), el('td', {}, p.brand), el('td', {}, p.name), el('td', { class: 'muted' }, p.codes['자사몰'] || '')))),
  );
}

function fillLists() {
  const uniq = (k) => [...new Set(state.info.products.map((p) => p[k]).filter(Boolean))];
  $('#cat-list').replaceChildren(...uniq('cat').map((v) => el('option', { value: v })));
  $('#brand-list').replaceChildren(...uniq('brand').map((v) => el('option', { value: v })));
}

// ---------------------------------------------------------------- 입력 카드

function addItem() {
  const node = $('#item-tpl').content.firstElementChild.cloneNode(true);
  const it = { el: node, touched: new Set() };
  state.items.push(it);
  // 다른 마켓 코드(자사몰 말고) + 자사몰 노출가
  const more = $('.more-fields', node);
  for (const { label } of state.info.codes.codeCols) {
    if (label === '자사몰') continue;
    more.append(el('label', { class: 'f' }, el('span', {}, `${label} 코드`), el('input', { type: 'text', 'data-k': `code:${label}`, autocomplete: 'off' })));
  }
  if (state.info.settle.blocks.length > 1) {
    more.append(el('label', { class: 'f' }, el('span', {}, `${state.info.settle.blocks[1].label} 노출가`), el('input', { type: 'text', 'data-k': 'jasaPrice', inputmode: 'numeric', autocomplete: 'off', placeholder: '비우면 노출가와 같게' })));
  }
  if (!state.info.codes.codeCols.some((c) => c.label === '자사몰')) $('[data-k="code:자사몰"]', node).closest('label').remove();
  node.addEventListener('input', (e) => onInput(it, e));
  node.addEventListener('change', (e) => onInput(it, e));
  $('.item-del', node).addEventListener('click', () => {
    state.items = state.items.filter((x) => x !== it);
    node.remove();
    renumber();
    schedulePreview();
  });
  $('#items').append(node);
  renumber();
  refreshAnchor(it);
  $('[data-k="name"]', node).focus();
}

function renumber() {
  state.items.forEach((it, i) => {
    $('.item-no', it.el).textContent = state.items.length > 1 ? `신제품 ${i + 1}` : '신제품';
    $('.item-del', it.el).hidden = state.items.length < 2;
    refreshAnchor(it);
  });
}

function onInput(it, e) {
  const k = e.target.dataset.k;
  if (!k) return;
  if (e.type === 'input') it.touched.add(k);
  if (k === 'cost' || k === 'price' || k === 'jasaPrice') {
    if (e.type === 'change') { const n = parseNum(e.target.value); e.target.value = n == null ? '' : n.toLocaleString('ko-KR'); }
  }
  if (k === 'name') {
    const name = e.target.value;
    // 직접 고치지 않은 칸만 상품명으로 추정해서 채운다
    if (!it.touched.has('brand') || !$('[data-k="brand"]', it.el).value) $('[data-k="brand"]', it.el).value = name.trim() ? guessBrand(name, state.info) : '';
    if (!it.touched.has('cat') || !$('[data-k="cat"]', it.el).value) $('[data-k="cat"]', it.el).value = name.trim() ? guessCategory(name, state.info) : '';
  }
  if (k === 'name' || k === 'cat' || k === 'brand') state.items.forEach((x) => refreshAnchor(x));
  resetDownload();
  schedulePreview();
}

function readItem(it) {
  const g = (k) => { const x = $(`[data-k="${k}"]`, it.el); return x ? x.value : ''; };
  const codes = {};
  for (const inp of it.el.querySelectorAll('[data-k^="code:"]')) if (norm(inp.value)) codes[inp.dataset.k.slice(5)] = norm(inp.value);
  const after = g('after');
  return {
    name: norm(g('name')), cost: parseNum(g('cost')), cat: norm(g('cat')), brand: norm(g('brand')),
    price: parseNum(g('price')), jasaPrice: parseNum(g('jasaPrice')), codes,
    afterName: after === '__auto' ? undefined : after === '__top' ? '' : after,
  };
}

// 넣을 자리 목록: 자동(같은 카테고리·브랜드 끝) · 맨 위 · 각 상품 다음 · 앞 카드의 신제품 다음
function refreshAnchor(it) {
  if (!state.info) return;
  const sel = $('[data-k="after"]', it.el);
  const prev = sel.value;
  const me = readItem(it);
  const idx = state.items.indexOf(it);
  const earlier = state.items.slice(0, idx).map(readItem).filter((p) => p.name);
  // 앞 카드의 신제품까지 넣은 목록에서 기본 자리를 계산
  const list = [...state.info.products];
  for (const p of earlier) {
    const at = p.afterName === '' ? -1 : p.afterName ? list.findIndex((x) => x.name === p.afterName) : defaultAnchor({ products: list }, p.cat, p.brand);
    list.splice(at + 1, 0, { name: p.name, cat: p.cat, brand: p.brand, fresh: true });
  }
  const auto = me.cat ? list[defaultAnchor({ products: list }, me.cat, me.brand)] : null;
  sel.replaceChildren(
    el('option', { value: '__auto' }, auto ? `자동 — ${auto.name} 다음` : '자동 — 카테고리를 넣으면 정해집니다'),
    el('option', { value: '__top' }, '맨 위'),
    ...list.map((p, i) => el('option', { value: p.name }, `${i + 1}. ${p.name} 다음${p.fresh ? ' (신제품)' : ''}`)),
  );
  sel.value = [...sel.options].some((o) => o.value === prev) ? prev : '__auto';
}

// ---------------------------------------------------------------- 미리보기(실제로 넣어 보고 계산)

let timer = 0;
function schedulePreview() {
  clearTimeout(timer);
  timer = setTimeout(runPreview, 250);
}

async function runPreview() {
  if (!state.file || !state.info) return;
  const seq = ++state.previewSeq;
  const items = state.items.map(readItem);
  const ready = [];
  const others = [];
  state.items.forEach((it, i) => {
    const p = items[i];
    const errs = p.name || p.cost ? validate(state.info, p, items) : [];
    const box = $('.item-errors', it.el);
    box.replaceChildren(...(errs.length ? [el('div', { class: 'note warn' }, el('ul', {}, errs.map((t) => el('li', {}, t))))] : []));
    if (p.name && !errs.length) ready.push({ p, it }); else others.push(it);
  });
  for (const it of others) $('.item-preview', it.el).replaceChildren();
  $('#build').disabled = !ready.length || ready.length !== state.items.length;
  if (!ready.length) return;
  try {
    const book = await openBook(JSZip, state.file.buf);
    const res = addProducts(book, ready.map((x) => x.p));
    if (seq !== state.previewSeq) return;
    ready.forEach(({ it }, i) => renderPreview(it, res.preview[i], ready[i].p));
  } catch (e) {
    if (seq !== state.previewSeq) return;
    for (const { it } of ready) $('.item-preview', it.el).replaceChildren(el('div', { class: 'note bad' }, e.message));
    $('#build').disabled = true;
  }
}

function renderPreview(it, pv, p) {
  const box = $('.item-preview', it.el);
  if (!pv) { box.replaceChildren(); return; }
  const head = el('p', { class: 'hint' }, `${pv.index + 1}번째 자리에 들어갑니다.`);
  if (p.price == null) {
    box.replaceChildren(head, el('p', { class: 'hint' }, '노출가를 비워 두면 정산가세팅 입력가(E열)가 빈칸으로 들어가고 가격 칸은 비어 있습니다. 엑셀에서 입력가만 넣으면 나머지가 계산됩니다. 여기에 노출가를 넣으면 계산 결과를 바로 볼 수 있습니다.'));
    return;
  }
  const rows = pv.blocks.map((b) => el('tr', {},
    el('td', { class: 'strong' }, b.label), el('td', { class: 'n' }, fmt(b.input)), el('td', { class: 'n' }, fmt(b.shown)), el('td', { class: 'n' }, fmt(b.maxBenefit)),
    el('td', { class: 'n' }, fmt(b.settle)), el('td', { class: 'n' }, fmt(b.cost)), el('td', { class: 'n' }, fmt(b.margin)), el('td', { class: 'n strong' }, pct(b.rate))));
  const t1 = el('table', { class: 'grid' },
    el('thead', {}, el('tr', {}, el('th', {}, '기준'), el('th', { class: 'n' }, '입력가'), el('th', { class: 'n' }, '노출가(2차혜택가)'), el('th', { class: 'n' }, '최대혜택가'), el('th', { class: 'n' }, '정산가'), el('th', { class: 'n' }, '원가'), el('th', { class: 'n' }, '마진'), el('th', { class: 'n' }, '마진율'))),
    el('tbody', {}, rows));
  const g = pv.blocks[0];
  const t2 = el('table', { class: 'grid' },
    el('thead', {}, el('tr', {}, el('th', {}, `마켓 (${g.label} 정산가로 맞춘 값)`), el('th', { class: 'n' }, '등재가(세팅)'), el('th', { class: 'n' }, '노출가'), el('th', { class: 'n' }, '정산액'))),
    el('tbody', {}, g.markets.map((m) => el('tr', {}, el('td', {}, m.label), el('td', { class: 'n strong' }, fmt(m.list)), el('td', { class: 'n' }, fmt(m.shown)), el('td', { class: 'n' }, fmt(m.settle))))));
  box.replaceChildren(head, el('div', { class: 'table-wrap' }, t1), el('h3', {}, '마켓별 등재가'), el('div', { class: 'table-wrap' }, t2));
}

// ---------------------------------------------------------------- 만들기·내려받기

function resetDownload() {
  if (state.downloadUrl) URL.revokeObjectURL(state.downloadUrl);
  state.downloadUrl = null;
  $('#download').hidden = true;
  $('#report').replaceChildren();
}

async function build() {
  resetDownload();
  const btn = $('#build');
  btn.disabled = true;
  try {
    const list = state.items.map(readItem);
    const book = await openBook(JSZip, state.file.buf);
    const res = addProducts(book, list);
    const bytes = await saveBook(book);
    const name = outputName(state.file.name);
    state.downloadUrl = URL.createObjectURL(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    const a = $('#download');
    a.href = state.downloadUrl; a.download = name; a.textContent = `내려받기 · ${name}`; a.hidden = false;
    const items = res.changes.map((c) => el('li', {}, `${c.sheet} ${c.row}행 — ${c.note}`));
    const notes = [el('h3', {}, '넣은 곳'), el('ul', {}, items)];
    if (res.failed.length) notes.push(el('div', { class: 'note warn' }, `미리 계산하지 못한 셀 ${res.failed.length}개는 Excel에서 열 때 계산됩니다.`));
    notes.push(el('p', { class: 'hint' }, '엑셀을 열면 전체를 다시 계산합니다. 다나와 최저가 비교는 동일스펙 사이트에 이 파일을 다시 올려 받으면 채워집니다.'));
    $('#report').replaceChildren(...notes);
  } catch (e) {
    console.error(e);
    $('#report').replaceChildren(el('div', { class: 'note bad' }, e.message));
  } finally {
    btn.disabled = false;
  }
}

// ---------------------------------------------------------------- 연결

const drop = $('#drop');
$('#picker').addEventListener('change', (e) => { takeFile(e.target.files[0]); e.target.value = ''; });
drop.addEventListener('dragover', (e) => { e.preventDefault(); drop.classList.add('over'); });
drop.addEventListener('dragleave', () => drop.classList.remove('over'));
drop.addEventListener('drop', (e) => { e.preventDefault(); drop.classList.remove('over'); takeFile(e.dataTransfer.files[0]); });
$('#add-item').addEventListener('click', () => { addItem(); schedulePreview(); });
$('#build').addEventListener('click', build);
