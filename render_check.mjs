// 檢查對帳新呈現（雙欄對齊）的 HTML 結構與資料一致性
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.js';
import { readFile } from 'fs/promises';
import vm from 'node:vm';

const html = await readFile(new URL('./index.html', import.meta.url), 'utf8');
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
    window: {}, XLSX: {},
    pdfjsLib: { GlobalWorkerOptions: {} },
    URL: { createObjectURL: () => '', revokeObjectURL() { } }, Blob: function () { },
};
ctx.pdfjsLib.getDocument = (params) => getDocument(params);
ctx.globalThis = ctx;
vm.createContext(ctx);

const data = new Uint8Array(await readFile('examples/moa-fund-115.pdf'));
mk('fileInputR');
els.fileInputR.files = [{ name: 'moa.pdf', arrayBuffer: async () => data }];
mk('fileInputD').files = { length: 0 };

const analysis = `
(async () => {
    await runReconcile();
    const { diff, onlyOne, ok } = _reconcileData;
    const cnt = id => (contHTML.match(new RegExp(id, 'g')) || []).length;
    const contHTML = document.getElementById('tableContainerR').innerHTML;
    // 每筆 diff：驗證 nRows/dRows 加總 === n/d（資料一致性）
    const sumIssues = [];
    for (const r of [...diff, ...onlyOne]) {
        const sn = (r.nRows || []).reduce((s, x) => s + x.amount, 0);
        const sd = (r.dRows || []).reduce((s, x) => s + x.amount, 0);
        if (r.n != null && sn !== r.n) sumIssues.push('n 不一致: ' + r.fund + '/' + r.l2 + ' n=' + r.n + ' sumRows=' + sn);
        if (r.d != null && sd !== r.d) sumIssues.push('d 不一致: ' + r.fund + '/' + r.l2 + ' d=' + r.d + ' sumRows=' + sd);
    }
    // 第一個 diff 的展開內容樣本
    const first = diff[0];
    const pairRows = (first.nRows || []).length + (first.dRows || []).length;
    return JSON.stringify({
        diff: diff.length, onlyOne: onlyOne.length, ok: ok.length,
        reconItems: cnt('class="recon-item"'),
        pairTables: cnt('class="recon-pair"'),
        openItems: cnt('recon-item" open'),
        sumIssues,
        first: {
            fund: first.fund, plan: first.plan, l1: first.l1, l2: first.l2,
            n_k: first.n / 1000, d_k: first.d / 1000, delta_k: first.delta / 1000,
            rollup: first.rollup || null,
            nRows: (first.nRows || []).map(x => (x.l3 || x.l2) + '=' + (x.amount / 1000) + 'k'),
            dRows: (first.dRows || []).map(x => (x.l3 || x.l2) + '=' + (x.amount / 1000) + 'k'),
            pairRows,
        },
    });
})()
`;
const pdf = await getDocument({ data }).promise;
ctx.__pdf = pdf;
const res = JSON.parse(await vm.runInContext(js + '\n' + analysis, ctx));
await pdf.destroy();
console.log(JSON.stringify(res, null, 1));
