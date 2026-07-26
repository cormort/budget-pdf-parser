// 回歸測試：以版控中的 PDF 驗證「編列說明」解析管線沒有退化。
//
// 兩個關鍵設計（皆源自 unit-budget-parser 的教訓）：
//  1. 直接載入 index.html 內的純函式（_extractNarrativeText → _autoCleanText →
//     parseNarrativeText）來跑，不自行複寫規則。自行複寫的驗證腳本會與實際行為漂移。
//  2. 除了正向指標，另做「反向驗證」：檢查清理階段刪掉的東西有沒有誤刪正文。
//     正向指標只能證明「數字對」，證明不了「文字完整」——unit-budget-parser 曾因此
//     讓 36 筆被折行截斷的名稱長期無人發現（四層加總驗算全程 0 不符）。
//
//   npm install && npm test
//
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.js';
import { readFile } from 'fs/promises';
import vm from 'node:vm';

const FIXTURE = '農業基金用途明細114年度預算案 (1).pdf';

// ── 期望值：任何規則改動若動到既有結果，這裡就會失敗 ──
const EXPECT = {
    rows: 169,          // 解析出的資料列
    unmatched: 0,       // 未匹配到任何科目的行（非資料的殘行）
    withL2: 147,        // 有二級科目者
    inferred: 40,       // 推論補上、以黃底標示者
    funds: ['農業發展基金'],
};

function loadTool(html) {
    // 取最後一個 <script>（前面兩個是 CDN 的 pdf.js 與 SheetJS）
    const js = html.split('<script>').slice(-1)[0].split('</script>')[0]
        .replace(/pdfjsLib\.GlobalWorkerOptions[^\n]*\n/, '');
    const els = {};
    const mk = id => els[id] || (els[id] = {
        id, value: '', textContent: '', innerHTML: '', style: {},
        files: { length: 0 }, options: [], addEventListener() { },
        querySelectorAll: () => [], getElementsByTagName: () => [],
    });
    const ctx = {
        console, document: { getElementById: mk, querySelectorAll: () => [], createElement: () => mk('_new') },
        window: {}, XLSX: {}, pdfjsLib: { GlobalWorkerOptions: {} },
        URL: { createObjectURL: () => '', revokeObjectURL() { } }, Blob: function () { },
    };
    ctx.globalThis = ctx;
    vm.createContext(ctx);
    vm.runInContext(js, ctx);
    return ctx;
}

// 反向驗證：stripPageNumbers 會把「重複出現 ≥3 次且長度 ≥8」的頁首字串
// 從句子中間剝除。若這些字串同時是正文用字（如長基金名出現在敘述裡），
// 剝除就會破壞內容且不會有任何徵兆。此檢查確保沒有誤刪。
function headerStripSafety(rawText) {
    const lines = rawText.split('\n');
    const pureCjk = /^[一-鿿]{2,20}$/;
    const counts = {};
    for (const line of lines) {
        const c = line.trim().replace(/[\s　]/g, '');
        if (pureCjk.test(c)) counts[c] = (counts[c] || 0) + 1;
    }
    const stripped = Object.keys(counts).filter(c => counts[c] >= 3 && c.length >= 8);
    const bodyLines = lines.filter(l => {
        const c = l.trim().replace(/[\s　]/g, '');
        return !(pureCjk.test(c) && counts[c] >= 3);
    });
    const bad = [];
    for (const h of stripped) {
        const hits = bodyLines.filter(l => l.replace(/[\s　]/g, '').includes(h));
        if (hits.length) bad.push(`「${h}」在正文出現 ${hits.length} 行，從句中剝除會破壞內容：${hits[0].trim().slice(0, 40)}…`);
    }
    return { stripped, bad };
}

const html = await readFile(new URL('./index.html', import.meta.url), 'utf8');
const ctx = loadTool(html);
const data = new Uint8Array(await readFile(new URL(`./${FIXTURE}`, import.meta.url)));
const task = getDocument({ data });
const pdf = await task.promise;

const raw = await ctx._extractNarrativeText(pdf);
const cleaned = ctx._autoCleanText(raw);
const { rows, unmatchedLines } = ctx.parseNarrativeText(cleaned);
await task.destroy();

const got = {
    rows: rows.length,
    unmatched: unmatchedLines.length,
    withL2: rows.filter(r => r.l2Item).length,
    inferred: rows.filter(r => r.accountFilled).length,
    funds: [...new Set(rows.map(r => r.fund).filter(Boolean))],
};

const errs = [];
for (const [k, want] of Object.entries(EXPECT)) {
    const a = JSON.stringify(got[k]), b = JSON.stringify(want);
    if (a !== b) errs.push(`${k}: 期望 ${b}，實際 ${a}`);
}
const safety = headerStripSafety(raw);
errs.push(...safety.bad.map(m => '頁首剝除誤傷正文 → ' + m));

if (errs.length) {
    console.error(`✗ ${FIXTURE}`);
    errs.forEach(e => console.error('    ' + e));
    console.error('\n若為刻意調整規則，請一併更新 test.mjs 的 EXPECT 與 README 數字。');
    process.exit(1);
}
console.log(`✓ ${FIXTURE}`);
console.log(`    ${got.funds.join('、')}｜${got.rows} 列｜有二級科目 ${got.withL2}（其中推論 ${got.inferred}）｜未匹配 ${got.unmatched} 行`);
console.log(`    反向驗證：句中剝除字串 ${safety.stripped.length} 種（${safety.stripped.join('、') || '無'}），皆未誤傷正文`);
console.log('\n全部通過。');
