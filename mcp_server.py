"""MCP server exposing the budget-PDF pipeline as two tools:

  1. extract_and_clean : PDF -> cleaned text (頁碼/標題移除、依標號重新分段)
  2. build_table       : cleaned text -> structured rows (+ optional xlsx)

Register:
  claude mcp add budget-parser -- <venv-python> mcp_server.py
"""

from pathlib import Path

from mcp.server.fastmcp import FastMCP

import parse_budget as pb

mcp = FastMCP("budget-parser")


@mcp.tool()
def extract_and_clean(pdf_path: str) -> str:
    """從預算書 PDF 提取文字並做一鍵清理（去頁碼、去頁首標題、依項目符號重新分段）。
    回傳清理後的文字，可先人工檢查或再做 regex 取代，之後交給 build_table。"""
    p = Path(pdf_path).expanduser()
    if not p.is_file():
        raise FileNotFoundError(f"找不到 PDF：{p}")
    return pb.auto_clean(pb.extract_pdf_text(p))


@mcp.tool()
def build_table(cleaned_text: str, out_xlsx: str = "") -> dict:
    """將清理後的文字解析為表格列（計畫/一二三級科目/金額/編列說明）。
    傳入 out_xlsx 路徑時同時輸出 xlsx（含摘要與清理後文字工作表）。"""
    rows, unmatched = pb.parse(cleaned_text)
    result = {
        "rows": rows,
        "unmatched": unmatched,
        "row_count": len(rows),
        "unmatched_count": len(unmatched),
    }
    if out_xlsx:
        out = Path(out_xlsx).expanduser()
        pb.write_xlsx(rows, unmatched, out, cleaned_text)
        result["xlsx"] = str(out)
    return result


if __name__ == "__main__":
    mcp.run()
