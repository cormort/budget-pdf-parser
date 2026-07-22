"""解析「基金用途明細表」數字表（非說明版），輸出 計畫/科目 × 各年度金額。

兩種來源版面：
  預算書細表：前年度決算數 | 業務計畫及用途別科目 | 本年度預算數 | 上年度預算數
  決算表    ：業務計畫及用途別科目 | 預算數 | 決算數 | 比較增減 | 備註

以頁首「前年度/本年度/上年度」或「預算數/決算數」標題的 x 座標動態決定欄位帶，
金額 token 依 x 距離歸到最近欄位；名稱層級用計畫標號(壹/一/（一）)與科目表判定。
"""

import re
import sys
from pathlib import Path

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from parse_budget import (
    CODE_TO_NAME,
    SECTION_MARKER,
    match_account,
    _fund_of_page,
)

DETAIL_MARKER = "基金用途明細表"
NUM_RE = re.compile(r"^[\d,]+$")
PLAN_L1 = re.compile(r"^[壹貳參肆伍陸柒捌玖拾]、")
PLAN_L2 = re.compile(r"^[一二三四五六七八九十]、")
PLAN_L3 = re.compile(r"^[（(][一二三四五六七八九十]+[）)]")


def _to_int(s: str):
    s = s.replace(",", "")
    if s in ("", "-", "--"):
        return None
    try:
        return int(s)
    except ValueError:
        return None


def _rows_by_line(page, y_tol=3):
    """把 words 依 y 分行，回傳 [(top, [(x0,text),...]), ...] 已按 x 排序。"""
    lines = {}
    for w in page.extract_words():
        key = None
        for k in lines:
            if abs(k - w["top"]) <= y_tol:
                key = k
                break
        lines.setdefault(key if key is not None else w["top"], []).append(w)
    out = []
    for top in sorted(lines):
        cells = sorted(lines[top], key=lambda w: w["x0"])
        out.append((top, [(w["x0"], w["text"]) for w in cells]))
    return out


def _find_columns(rows):
    """從表頭決定金額欄中心與名稱右界。回傳 (amount_cols:dict, right_cut:float)。

    預算書細表：前年度決算 / 本年度預算 / 上年度預算
    決算表    ：預算數(114) / 決算數(114)（忽略比較增減、%）
    """
    pos = {}
    for _top, cells in rows[:14]:
        for x0, text in cells:
            pos.setdefault(text.strip(), x0)
    right_cut = pos.get("計畫內容說明") or pos.get("備註") or pos.get("備") or 9999
    amount_cols = {}
    if "前年度" in pos:  # 預算書細表：前年度決算 | 本年度預算 | 上年度預算
        amount_cols = {
            "前年度決算": pos["前年度"] + 20,
            "本年度預算": pos["本年度"] + 20,
            "上年度預算": pos["上年度"] + 20,
        }
    elif "預" in pos and "決" in pos:  # 決算表(表頭字距大被拆字)：預算數 | 決算數
        amount_cols = {
            "預算數": pos["預"] + 13,
            "決算數": pos["決"] + 13,
        }
    return amount_cols, right_cut


def parse_detail_pdf(pdf_path: Path):
    """generator：逐列產出 row = {fund, level, plan, l1/l2/l3, name, amounts}"""
    with pdfplumber.open(pdf_path) as pdf:
        page_texts = [p.extract_text() or "" for p in pdf.pages]
        for pi, page in enumerate(pdf.pages):
            head = [ln.replace(" ", "") for ln in page_texts[pi].split("\n")[:6]]
            # 只取細表頁：頁首含「基金用途明細表」(非說明) 且頁內有科目
            if not any(ln == DETAIL_MARKER for ln in head):
                continue
            if not any(k in page_texts[pi] for k in ("用人費用", "服務費用", "材料及用品費")):
                continue  # 略過只有計畫層的彙總頁
            fund = _fund_of_page(page_texts[pi], SECTION_MARKER) or _fund_of_page(
                page_texts[pi], DETAIL_MARKER
            )
            rows = _rows_by_line(page)
            cols, right_cut = _find_columns(rows)
            if not cols:
                continue
            yield from _parse_page(rows, cols, right_cut, fund)


def _assign(cells, cols, right_cut):
    """把一行切成 (name, {欄名:int})。名稱=右界左側的 CJK；金額=最近欄(≤45px)。"""
    name_parts, amounts = [], {}
    for x0, text in cells:
        t = text.strip()
        if not t:
            continue
        if NUM_RE.match(t.replace("-", "")) and any(c.isdigit() for c in t):
            best = min(cols.items(), key=lambda kv: abs(kv[1] - x0))
            if abs(best[1] - x0) <= 45:  # 太遠的數字（比較增減/%）不歸欄
                amounts.setdefault(best[0], _to_int(t))
        elif x0 < right_cut - 8:  # 邊距避免說明欄第一字被吃進名稱
            name_parts.append((x0, t))
    name_parts.sort()
    return "".join(p[1] for p in name_parts), amounts


def _parse_page(rows, cols, right_cut, fund):
    plan1 = plan2 = plan3 = ""
    l1 = l2 = l3 = ""
    for _top, cells in rows:
        name, amounts = _assign(cells, cols, right_cut)
        if not name:
            continue
        if name in ("業務計畫及用途別科目", "計畫內容說明", "基金用途"):
            continue
        if PLAN_L1.match(name):
            plan1 = PLAN_L1.sub("", name).strip()
            plan2 = plan3 = l1 = l2 = l3 = ""
            level = "計畫"
        elif PLAN_L2.match(name):
            plan2 = PLAN_L2.sub("", name).strip()
            plan3 = l1 = l2 = l3 = ""
            level = "計畫"
        elif PLAN_L3.match(name):
            plan3 = PLAN_L3.sub("", name).strip()
            l1 = l2 = l3 = ""
            level = "計畫"
        else:
            # 科目：用科目表判層級
            m = match_account(name)
            if not m:
                continue
            clen = len(m["code"])
            if clen == 1:
                l1, l2, l3 = m["name"], "", ""
            elif clen == 2:
                l1 = CODE_TO_NAME.get(m["code"][:1], l1)
                l2, l3 = m["name"], ""
            else:
                l1 = CODE_TO_NAME.get(m["code"][:1], l1)
                l2 = CODE_TO_NAME.get(m["code"][:2], l2)
                l3 = m["name"]
            level = "科目"
        yield {
            "fund": fund or "",
            "level": level,
            "plan": " - ".join(p for p in (plan1, plan2, plan3) if p),
            "l1": l1 if level == "科目" else "",
            "l2": l2 if level == "科目" else "",
            "l3": l3 if level == "科目" else "",
            "name": name,
            "amounts": amounts,
        }


if __name__ == "__main__":
    p = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("115年度預算案 (2).pdf")
    rows = list(parse_detail_pdf(p))
    print(f"rows={len(rows)}")
    for r in rows[:20]:
        print(r["fund"][:6], r["level"], "|", r["plan"][:20], "|", r["name"][:16], r["amounts"])
