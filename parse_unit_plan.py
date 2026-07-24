#!/usr/bin/env python3
"""解析單位預算案的「歲出計畫提要及分支計畫概況表」。

層級：工作計畫(10碼) → 分支計畫(2碼) → 用途別一級(X000) → 用途別二級(其餘4碼)。
右側敘述(計畫內容/預期成果/說明)不納入結構表；只留左側表格 + 承辦單位。
用字座標切左表：代碼與名稱 x0<160、金額右靠 x0∈[200,290)、承辦單位 x0∈[290,360)。
"""
import re, sys
import pdfplumber

TITLE = "歲出計畫提要及分支計畫概況表"
NUM = re.compile(r"^-?[\d,]+")  # 金額欄有時與承辦單位黏成一詞，只取開頭數字


def _lines(page):
    """把左表區的字組成 (name, amount, unit) 列。依 top 鄰近分群，容忍同列微小高低差。"""
    ws = sorted(page.extract_words(), key=lambda w: w["top"])
    rows = []
    for w in ws:
        if not rows or w["top"] - rows[-1]["_top"] > 6:
            rows.append({"_top": w["top"], "name": [], "amt": [], "unit": []})
        r = rows[-1]
        x0, t = w["x0"], w["text"]
        if x0 < 160:
            r["name"].append((x0, t))
        elif 200 <= x0 < 292 and NUM.match(t):
            m = NUM.match(t)
            r["amt"].append(m.group(0))
            rest = t[m.end():]                    # 黏在金額後的承辦單位
            if rest:
                r["unit"].append((x0, rest))
        elif 292 <= x0 < 360:
            r["unit"].append((x0, t))
    out = []
    for r in rows:
        name = "".join(t for _, t in sorted(r["name"]))
        amt = "".join(r["amt"])
        unit = "".join(t for _, t in sorted(r["unit"]))
        out.append((name, amt, unit))
    return out


def parse(pdf_path):
    pdf = pdfplumber.open(pdf_path)
    result = []
    cur_plan = None            # (code, name, budget)
    ctx = {"branch": None, "l1": None}
    for page in pdf.pages:
        text = page.extract_text() or ""
        if TITLE not in text:
            continue
        # 工作計畫表頭：工作計畫名稱及編號 <10碼> <名稱> 預算金額 <金額>
        m = re.search(r"計畫名稱及編號\s+(\d{6,})\s+(.+?)\s+預算金額\s+([\d,]+)", text)
        if m:
            plan = (m.group(1), m.group(2).strip(), m.group(3))
            if plan[0] != (cur_plan[0] if cur_plan else None):
                cur_plan = plan
                ctx = {"branch": None, "l1": None}
        for name, amt, unit in _lines(page):
            m = re.match(r"^(\d{2,4})\s*(.+)$", name)
            if not m or not amt:
                continue
            code, cname = m.group(1), m.group(2).strip()
            if len(code) == 2:                    # 分支計畫
                ctx = {"branch": (code, cname), "l1": None}
                level, l1c, l1n, l2c, l2n = "分支計畫", "", "", "", ""
            elif len(code) == 4 and code.endswith("000"):   # 用途別一級
                ctx["l1"] = (code, cname)
                level, l1c, l1n, l2c, l2n = "用途別一級", code, cname, "", ""
            elif len(code) == 4:                  # 用途別二級
                level, l1c, l1n, l2c, l2n = "用途別二級", \
                    (ctx["l1"][0] if ctx["l1"] else ""), \
                    (ctx["l1"][1] if ctx["l1"] else ""), code, cname
            else:
                continue
            b = ctx["branch"] or ("", "")
            result.append({
                "工作計畫代碼": cur_plan[0], "工作計畫名稱": cur_plan[1],
                "工作計畫預算": cur_plan[2],
                "分支計畫代碼": b[0], "分支計畫名稱": b[1],
                "用途別一級代碼": l1c, "用途別一級名稱": l1n,
                "用途別二級代碼": l2c, "用途別二級名稱": l2n,
                "層級": level, "金額": amt.replace(",", ""),
                "承辦單位": unit if level == "分支計畫" else "",
            })
    return result


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else "115年度行政院主計總處單位預算案.pdf"
    rows = parse(src)
    try:
        import openpyxl
        wb = openpyxl.Workbook(); ws = wb.active
        cols = list(rows[0].keys())
        ws.append(cols)
        for r in rows:
            ws.append([r[c] for c in cols])
        out = src.rsplit(".", 1)[0] + "_歲出計畫概況表.xlsx"
        wb.save(out)
        print(f"{len(rows)} 列 -> {out}")
    except ImportError:
        import csv
        w = csv.DictWriter(sys.stdout, fieldnames=list(rows[0].keys()))
        w.writeheader(); w.writerows(rows)


if __name__ == "__main__":
    main()
