// 驗收：1) 逐(基金|二級) 的 !isSubtotal 金額總和 before/after 必須不變
//      2) 抽查拆出的子列內容
// 用法：node verify.mjs <before.html> <after.html> <pdf>
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.js';
import { readFile } from 'fs/promises';
import vm from 'node:vm';

async function load(htmlPath, pdfData) {
    const html = await readFile(new URL(htmlPath, import.meta.url), 'utf8');
    const js = html.split('<script>').slice(-1)[0].split('</script>')[0]
        .replace(/pdfjsLib\.GlobalWorkerOptions[^\n]*\n/, '');
    const els = {};
    const mk = id => els[id] || (els[id] = {
        id, value: '', textContent: '', innerHTML: '', style: {}, disabled: false,
        files: { length: 0 }, options: [], addEventListener() { },
        querySelectorAll: () => [], getElementsByTagName: () => [], tBodies: [],
    });
    const ctx = {
        console, document: { getElementById: mk, querySelectorAll: () => [], createElement: () => mk('_new') },
        window: {}, XLSX: {}, pdfjsLib: { GlobalWorkerOptions: {} },
        URL: { createObjectURL: () => '', revokeObjectURL() { } }, Blob: function () { },
    };
    ctx.globalThis = ctx;
    vm.createContext(ctx);
    const pdf = await getDocument({ data: pdfData }).promise;
    ctx.__pdf = pdf;
    const analysis = `
(async () => {
    const raw = await _extractNarrativeText(__pdf);
    const cleaned = _autoCleanText(raw);
    const { rows } = parseNarrativeText(cleaned);
    const sums = new Map();
    for (const r of rows) {
        if (r.isSubtotal || !r.amount) continue;
        const k = r.fund + '|' + (r.l2Item || '(L1)');
        sums.set(k, (sums.get(k) || 0) + r.amount);
    }
    const derived = rows.filter(r => r.derived).map(r => ({
        fund: r.fund, l1: r.l1Item, l2: r.l2Item, l3: r.l3Item,
        amt_k: r.amount / 1000, isSub: r.isSubtotal, filled: r.accountFilled,
        desc: (r.description || '').slice(0, 30),
    }));
    return JSON.stringify({
        rows: rows.length,
        isSubtotal: rows.filter(r => r.isSubtotal).length,
        derived: derived.length,
        sums: [...sums.entries()].sort().map(([k, v]) => k + '=' + v),
        sampleDerived: derived.slice(0, 8),
    });
})()`;
    const res = JSON.parse(await vm.runInContext(js + '\n' + analysis, ctx));
    await pdf.destroy();
    return res;
}

const [before, after, pdfPath] = process.argv.slice(2);
const data = new Uint8Array(await readFile(new URL(pdfPath, import.meta.url)));
const b = await load(before, data);
const a = await load(after, data);
console.log('before:', b.rows, 'rows,', b.isSubtotal, 'isSubtotal,', b.derived, 'derived');
console.log('after :', a.rows, 'rows,', a.isSubtotal, 'isSubtotal,', a.derived, 'derived');
console.log('新增 derived 子列樣本:');
for (const s of a.sampleDerived.slice(0, 10)) console.log('  ', JSON.stringify(s));
// 金額總和不變檢查
const bm = new Map(b.sums.map(s => [s.split('=')[0], parseInt(s.split('=')[1], 10)]));
const am = new Map(a.sums.map(s => [s.split('=')[0], parseInt(s.split('=')[1], 10)]));
const keys = new Set([...bm.keys(), ...am.keys()]);
let changed = 0;
for (const k of keys) {
    if ((bm.get(k) || 0) !== (am.get(k) || 0)) { changed++; console.log(`  ✗ ${k}: before ${bm.get(k)} after ${am.get(k)}`); }
}
console.log(changed === 0 ? `✓ 逐(基金|二級)金額總和 ${keys.size} 組全部不變` : `✗ ${changed} 組金額總和改變！`);
