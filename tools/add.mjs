// 명령줄 실행기(검증용). 웹 화면과 같은 src/ 모듈을 쓴다.
//   node tools/add.mjs <가격세팅.xlsx> <결과.xlsx> <상품.json>
//   상품.json: [{ "name": "...", "cost": 1000, "cat": "VGA", "brand": "Acer", "price": null,
//                "codes": { "자사몰": "123" }, "afterName": "앞 상품명(생략하면 기본 자리)" }]
// 입력·결과 파일은 레포 밖에 둔다. 레포에는 어떤 엑셀도 커밋하지 않는다.
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { openBook, saveBook } from '../src/book.js';
import { addProducts } from '../src/newproduct.js';

const require = createRequire(import.meta.url);
const JSZip = require('jszip');

const [input, output, spec] = process.argv.slice(2);
if (!input || !output || !spec) { console.error('usage: node tools/add.mjs <in.xlsx> <out.xlsx> <products.json>'); process.exit(2); }
const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
if (path.resolve(output).startsWith(repo)) { console.error('결과 파일을 레포 안에 쓰지 않습니다.'); process.exit(2); }

const list = JSON.parse(fs.readFileSync(spec, 'utf8'));
const book = await openBook(JSZip, fs.readFileSync(input));
const res = addProducts(book, list);
for (const c of res.changes) console.log(`${c.sheet} ${c.row}행 — ${c.note}`);
for (const p of res.preview) {
  console.log(`\n[${p.name}] ${p.index + 1}번째`);
  for (const b of p.blocks) console.log(`  ${b.label} ${b.row}행: 입력가 ${b.input} 노출가 ${b.shown} 최대혜택가 ${b.maxBenefit} 정산가 ${b.settle} 원가 ${b.cost} 마진율 ${b.rate}`);
}
if (res.failed.length) console.log('계산 못 한 셀(Excel이 열 때 계산):', res.failed.map((f) => `${f.sheet}!${f.addr}`).join(', '));
fs.writeFileSync(output, await saveBook(book));
console.log('\n저장:', output);
