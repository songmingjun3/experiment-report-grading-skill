#!/usr/bin/env python
"""Write manually reviewed grades/comments to Word reports and an Excel gradebook.

Input must be a manual decision CSV. This script refuses to run unless the CSV
contains an explicit manual-basis column. It does not calculate scores.
"""

from __future__ import annotations

import argparse
import csv
import re
import shutil
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill
from docx import Document
from docx.shared import Pt


REQUIRED_COLUMNS = [
    "实验",
    "学号",
    "姓名",
    "源文件",
    "最终成绩",
    "最终评语",
    "人工判定依据",
]


def clean_stem(path: Path) -> str:
    return re.sub(r"[（(].*?批阅.*?[）)]", "", path.stem).strip()


def clear_existing_score(paragraph) -> None:
    seen = False
    for run in paragraph.runs:
        text = run.text or ""
        if "成绩评定" in text or "得分" in text:
            seen = True
            continue
        if seen:
            run.text = re.sub(r"\d{1,3}", "", text).replace("分", "")


def write_docx(src: Path, dst: Path, score: str, comment: str) -> None:
    doc = Document(src)
    if not doc.tables or len(doc.tables[0].rows) <= 5:
        raise ValueError("cannot locate grading cell in first table")
    cell = doc.tables[0].rows[5].cells[0]
    while len(cell.paragraphs) <= 4:
        cell.add_paragraph()
    p = cell.paragraphs[2]
    p.clear()
    run = p.add_run(comment)
    run.font.size = Pt(12)
    score_p = cell.paragraphs[4]
    clear_existing_score(score_p)
    replaced = False
    for run in score_p.runs:
        if run.font.underline and run.text.strip() not in {"成绩评定", "得分"}:
            run.text = f"  {score}  "
            replaced = True
            break
    if not replaced:
        score_p.add_run(f"  {score}  ")
    doc.save(dst)


def validate_rows(rows: list[dict[str, str]]) -> None:
    if not rows:
        raise SystemExit("manual decision CSV is empty")
    missing = [c for c in REQUIRED_COLUMNS if c not in rows[0]]
    if missing:
        raise SystemExit(f"manual decision CSV missing required columns: {', '.join(missing)}")
    for i, row in enumerate(rows, 2):
        if not row.get("人工判定依据", "").strip():
            raise SystemExit(f"row {i} missing 人工判定依据; refuse to write final results")


def update_gradebook(template: Path, output: Path, rows: list[dict[str, str]]) -> int:
    shutil.copy2(template, output)
    wb = openpyxl.load_workbook(output)
    ws = wb.active
    headers = {str(ws.cell(1, c).value or "").strip(): c for c in range(1, ws.max_column + 1)}
    sid_col = headers.get("学号")
    if not sid_col:
        raise ValueError("gradebook missing 学号 column")
    by_key = {(r["实验"], r["学号"]): r for r in rows}
    red_font = Font(color="FFFF0000")
    red_fill = PatternFill(fill_type="solid", fgColor="FFFFE5E5")
    blue_font = Font(color="FF0000FF")
    blue_fill = PatternFill(fill_type="solid", fgColor="FFE5F0FF")
    updated = 0
    for rr in range(2, ws.max_row + 1):
        sid = str(ws.cell(rr, sid_col).value or "").strip()
        if sid.endswith(".0"):
            sid = sid[:-2]
        for exp in range(1, 50):
            col = headers.get(f"实验{exp}")
            if not col:
                continue
            row = by_key.get((str(exp), sid))
            if not row:
                continue
            cell = ws.cell(rr, col)
            cell.value = int(float(row["最终成绩"]))
            flag = row.get("标记", "")
            if "抄袭" in flag:
                cell.font = blue_font
                cell.fill = blue_fill
            elif row.get("是否补交") == "是" or "补交" in flag:
                cell.font = red_font
                cell.fill = red_fill
            updated += 1
    wb.save(output)
    wb.close()
    return updated


def main() -> None:
    parser = argparse.ArgumentParser(description="Write manual experiment report grading results.")
    parser.add_argument("manual_csv", help="Manual decision CSV with required columns")
    parser.add_argument("--out-dir", required=True, help="Output directory for reviewed reports and summaries")
    parser.add_argument("--gradebook", help="Optional gradebook template .xlsx")
    parser.add_argument("--gradebook-out", default="final_gradebook.xlsx")
    args = parser.parse_args()

    with open(args.manual_csv, "r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    validate_rows(rows)
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    ok = 0
    for row in rows:
        src = Path(row["源文件"])
        exp_dir = out_dir / f"实验{row['实验']}最终批阅版"
        exp_dir.mkdir(exist_ok=True)
        dst = exp_dir / f"{clean_stem(src)}（最终批阅版）.docx"
        write_docx(src, dst, row["最终成绩"], row["最终评语"])
        row["最终报告"] = str(dst.resolve())
        ok += 1
    summary_csv = out_dir / "最终人工评分明细.csv"
    fields = list(rows[0].keys())
    if "最终报告" not in fields:
        fields.append("最终报告")
    with summary_csv.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)
    updated = 0
    if args.gradebook:
        updated = update_gradebook(Path(args.gradebook), out_dir / args.gradebook_out, rows)
    print(f"wrote {ok} reviewed reports")
    print(f"wrote {summary_csv}")
    if args.gradebook:
        print(f"updated {updated} gradebook cells")


if __name__ == "__main__":
    main()
