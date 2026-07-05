#!/usr/bin/env python
"""Extract report text, image counts, and optional OCR evidence.

This script is intentionally evidence-only. It does not decide completeness,
deductions, comments, final scores, plagiarism, or formatting penalties.
"""

from __future__ import annotations

import argparse
import csv
import re
import zipfile
from pathlib import Path

from docx import Document


SECTION_LABELS = [
    "实验内容",
    "实验要求",
    "实验任务",
    "实验步骤",
    "实验结果",
    "实验步骤、实验结果及分析",
    "实验总结",
    "教师评语",
]


def normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def text_from_docx(path: Path) -> str:
    doc = Document(path)
    parts: list[str] = []
    for p in doc.paragraphs:
        if p.text.strip():
            parts.append(p.text)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                cell_text = "\n".join(p.text for p in cell.paragraphs if p.text.strip())
                if cell_text.strip():
                    parts.append(cell_text)
    return "\n".join(parts)


def section(text: str, starts: list[str], ends: list[str]) -> str:
    positions = [(text.find(s), s) for s in starts if text.find(s) >= 0]
    if not positions:
        return ""
    pos, label = min(positions, key=lambda x: x[0])
    end = len(text)
    for e in ends:
        p = text.find(e, pos + len(label))
        if p >= 0:
            end = min(end, p)
    return normalize(text[pos:end])


def image_count(path: Path) -> int:
    try:
        with zipfile.ZipFile(path) as zf:
            return sum(1 for n in zf.namelist() if n.startswith("word/media/"))
    except Exception:
        return 0


def parse_identity(path: Path) -> tuple[str, str]:
    stem = path.stem
    stem = re.sub(r"[（(].*?[）)]", "", stem)
    m = re.match(r"(?P<sid>\d{8,})[-_－]?(?P<name>.*)", stem)
    if not m:
        return "", stem
    return m.group("sid"), m.group("name").strip("-_－ ")


def report_rows(root: Path) -> list[dict[str, str]]:
    rows = []
    for path in sorted(root.rglob("*.docx")):
        if path.name.startswith("~$"):
            continue
        if "最终批阅版" in path.name:
            continue
        sid, name = parse_identity(path)
        exp_match = re.search(r"实验\s*(\d+)", str(path))
        exp = exp_match.group(1) if exp_match else ""
        late = "是" if any(x in path.name for x in ["补交", "迟交", "逾期"]) else "否"
        try:
            text = text_from_docx(path)
            content = section(text, ["实验内容", "实验要求", "实验任务"], ["实验设备", "实验步骤", "实验结果", "实验总结", "教师评语"])
            process = section(text, ["实验步骤、实验结果及分析", "实验步骤", "实验结果"], ["实验总结", "教师评语", "成绩评定"])
            summary = section(text, ["实验总结"], ["教师评语", "成绩评定"])
            error = ""
        except Exception as exc:
            text = content = process = summary = ""
            error = str(exc)
        rows.append(
            {
                "实验": exp,
                "学号": sid,
                "姓名": name,
                "文件": str(path.resolve()),
                "是否补交": late,
                "图片数": str(image_count(path)),
                "实验内容长度": str(len(content)),
                "实验过程长度": str(len(process)),
                "总结长度": str(len(summary)),
                "实验内容摘录": content[:500],
                "实验过程摘录": process[:1000],
                "总结摘录": summary[:800],
                "读取错误": error,
                "脚本结论": "证据提取，仅供人工复核",
            }
        )
    return rows


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract evidence from experiment report docx files.")
    parser.add_argument("root", help="Course/report root directory")
    parser.add_argument("--out", default="report_evidence.csv", help="Output CSV path")
    args = parser.parse_args()
    rows = report_rows(Path(args.root))
    fields = list(rows[0].keys()) if rows else ["实验", "学号", "姓名", "文件", "脚本结论"]
    with open(args.out, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)
    print(f"wrote {len(rows)} rows to {args.out}")


if __name__ == "__main__":
    main()
