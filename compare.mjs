// before/after 對照：跑真實 runReconcile（經 vm 注入真 pdfjsLib 與 fake file input）
// 用法：node compare.mjs <index.html 路徑> <pdf 路徑>
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.js';
import { readFile } from 'fs/promises';
import vm from 'node:vm';

const [htmlPath, pdfPath] = process.argv.slice(2);
const html = await readFile(new URL(htmlPath, import.meta.url), 'utf8');
const js = html.split('<script>').slice(-1)[0].split('</script>')[0]
    .replace(/pdfjsLib\.GlobalWorkerOptions[^\n]*\n/, '');

const els = {};
const mk = id => els[id] || (els[id] = {
    id, value: '', textContent: '', innerHTML: '', style: {}, disabled: false,
    files: { length: 0 }, options: [], addEventListener() { },
    querySelectorAll: () => [], getElementsByTagName: () => [], tBodies: [],
    classList: { add() { }, remove() { } },
});
const ctx = {
    console, document: { getElementById: mk, querySelectorAll: () => [], createElement: () => mk('_new') },
    window: {}, XLSX: {},
    pdfjsLib: { GlobalWorkerOptions: {} },
    URL: { createObjectURL: () => '', revokeObjectURL() { } }, Blob: function () { },
};
// 注入真實 pdfjs：runReconcile 會呼叫 pdfjsLib.getDocument(...).promise
ctx.pdfjsLib.getDocument = (params) => getDocument(params);
ctx.globalThis = ctx;
vm.createContext(ctx);

const data = new Uint8Array(await readFile(new URL(pdfPath, import.meta.url)));
mk('fileInputR');   // 先建立 stub 再注入 fake file
els.fileInputR.files = [{ name: pdfPath.split('/').pop(), arrayBuffer: async () => data }];
mk('fileInputD').files = { length: 0 };

const analysis = `
(async () => {
    try {
        await runReconcile();
    } catch (e) {
        console.error('runReconcile 失敗:', e);
    }
    const { diff, onlyOne, ok } = _reconcileData;
    const sumAbs = diff.reduce((s, r) => s + Math.abs(r.delta || 0), 0);
    const big = diff.filter(r => Math.abs(r.delta || 0) >= 1000 * 1000);   // 差額 ≥ 1,000 千元
    const out = {
        rows: parseNarrativeText(_autoCleanText(await _extractNarrativeText(__pdf))).rows.length,
        diff: diff.length, onlyOne: onlyOne.length, ok: ok.length,
        sumAbsDelta_k: Math.round(sumAbs / 1000),
        big: big.map(r => ({ fund: r.fund, plan: r.plan, l1: r.l1, l2: r.l2, delta_k: Math.round((r.delta || 0) / 1000) })),
    };
    return JSON.stringify(out);
})()
`;
const pdf = await getDocument({ data }).promise;
ctx.__pdf = pdf;
const res = JSON.parse(await vm.runInContext(js + '\n' + analysis, ctx));
await pdf.destroy();
console.log(JSON.stringify(res, null, 1));
