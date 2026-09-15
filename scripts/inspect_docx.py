#!/usr/bin/env python3
"""Export a read-only DOCX structure and effective-format inspection as JSON."""

from __future__ import annotations

import argparse
import hashlib
import json
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

from docx_model import DocxPackage
from effective_style import EffectiveStyleResolver
from structure_classifier import StructureClassifier


def build_inspection(path: str | Path, include_runs: bool = False, max_paragraphs: int | None = None) -> dict:
    path = Path(path)
    package = DocxPackage(path)
    resolver = EffectiveStyleResolver(package)
    paragraphs = list(package.iter_paragraphs())
    classified = StructureClassifier(package, resolver).classify_document(paragraphs)
    if max_paragraphs is not None:
        classified_for_output = classified[: max(0, max_paragraphs)]
    else:
        classified_for_output = classified
    paragraph_data = []
    for item in classified_for_output:
        paragraph = item.paragraph
        entry = {
            "index": paragraph.index,
            "location": paragraph.location,
            "text": paragraph.text,
            "style_id": paragraph.style_id,
            "section_index": paragraph.section_index,
            "in_table": paragraph.in_table,
            "in_content_control": paragraph.in_content_control,
            "in_revision": paragraph.in_revision,
            "role": item.classification.role,
            "role_confidence": item.classification.confidence,
            "role_evidence": item.classification.evidence,
            "paragraph_format": resolver.get_effective_paragraph_format(paragraph),
        }
        if include_runs:
            entry["runs"] = [
                {
                    "index": run.index,
                    "text": run.text,
                    "style_id": run.style_id,
                    "effective_format": resolver.get_effective_run_format(run, paragraph),
                }
                for run in package.iter_runs(paragraph)
            ]
        paragraph_data.append(entry)
    role_counts = Counter(item.classification.role for item in classified)
    return {
        "schema_version": "1.0.0-phase3",
        "source": {
            "path": str(path.resolve()),
            "filename": path.name,
            "size_bytes": path.stat().st_size,
            "sha256": hashlib.sha256(path.read_bytes()).hexdigest().upper(),
        },
        "created_at_utc": datetime.now(timezone.utc).isoformat(),
        "capabilities": {
            "read_only": True,
            "effective_run_format": True,
            "effective_paragraph_format": True,
            "structure_classification": True,
            "final_page_numbers_available": False,
            "render_dependent_layout_checked": False,
            "rule_evaluation_implemented": False,
            "comment_writing_implemented": False,
            "report_writing_implemented": False,
        },
        "inventory": package.inventory(),
        "sections": package.sections(),
        "settings": package.settings(),
        "numbering": package.numbering_summary(),
        "styles": package.styles_summary(),
        "relationships": package.relationship_inventory(),
        "role_counts": dict(sorted(role_counts.items())),
        "paragraphs": paragraph_data,
        "limitations": [
            "OOXML 本身不能可靠给出最终页码、自动分页、视觉重叠和大面积留白。",
            "表格条件样式当前只合并表格样式主体和 wholeTable 层，复杂首行、末行及条带条件保留给后续阶段。",
            "结构角色是带置信度的候选分类，低置信度结果需要后续规则引擎或人工确认。",
        ],
    }


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="只读解析 DOCX 结构和最终有效格式，并输出 JSON。")
    parser.add_argument("docx", help="待解析的 DOCX 文件")
    parser.add_argument("--output", "-o", help="输出 JSON 文件；省略时写到标准输出")
    parser.add_argument("--include-runs", action="store_true", help="包含逐文字片段的有效格式详情")
    parser.add_argument("--max-paragraphs", type=int, help="仅限制 JSON 中输出的段落数量，不影响统计")
    args = parser.parse_args(argv)
    result = build_inspection(args.docx, include_runs=args.include_runs, max_paragraphs=args.max_paragraphs)
    text = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        output = Path(args.output)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(text + "\n", encoding="utf-8")
    else:
        print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
