#!/usr/bin/env python3
"""Run all Phase 7 checks and create comments plus the Word report."""

from __future__ import annotations

import argparse
import json
from collections import Counter
from pathlib import Path
from zipfile import ZipFile

from lxml import etree

from comment_writer import default_output_path, write_comments
from docx_model import DocxPackage
from report_writer import DEFAULT_TEMPLATE_NAME, default_report_path, write_report
from review_rules import DEFAULT_RULES, build_review


PRESERVED_INVENTORY_KEYS = (
    "paragraphs_body", "tables", "sections", "content_controls",
    "tracked_insertions", "tracked_deletions", "drawings_inline",
    "drawings_anchor", "vml_shapes", "omml_objects", "fields", "bookmarks",
    "hyperlinks", "footnotes", "endnotes", "hidden_runs",
    "manual_page_breaks", "manual_column_breaks", "paragraph_section_breaks",
    "ole_objects", "text_boxes", "smartart_relationship_markers",
)


def _verify_commented_copy(source: Path, output: Path, comments: dict) -> dict:
    before = DocxPackage(source)
    after = DocxPackage(output)
    before_inventory = before.inventory()
    after_inventory = after.inventory()
    differences = {
        key: {"before": before_inventory.get(key), "after": after_inventory.get(key)}
        for key in PRESERVED_INVENTORY_KEYS
        if before_inventory.get(key) != after_inventory.get(key)
    }
    expected_comments = before_inventory.get("comments", 0) + comments.get("comment_count", 0)
    if after_inventory.get("comments") != expected_comments:
        differences["comments"] = {"expected": expected_comments, "after": after_inventory.get("comments")}

    allowed_changed = {
        "[Content_Types].xml", "word/document.xml", "word/_rels/document.xml.rels",
        "word/comments.xml",
    }
    with ZipFile(source) as before_zip, ZipFile(output) as after_zip:
        corrupt_part = after_zip.testzip()
        xml_error = None
        try:
            for name in after_zip.namelist():
                if name.endswith(".xml") or name.endswith(".rels"):
                    etree.fromstring(after_zip.read(name))
        except Exception as exc:  # pragma: no cover - defensive package validation
            xml_error = str(exc)
        byte_changed = []
        for name in before_zip.namelist():
            if name in allowed_changed:
                continue
            if name not in after_zip.namelist() or before_zip.read(name) != after_zip.read(name):
                byte_changed.append(name)
    if byte_changed:
        differences["unchanged_parts"] = byte_changed
    if corrupt_part:
        differences["zip_crc"] = corrupt_part
    if xml_error:
        differences["xml_parse"] = xml_error
    return {
        "passed": not differences,
        "checked_inventory_keys": list(PRESERVED_INVENTORY_KEYS),
        "expected_comment_count": expected_comments,
        "actual_comment_count": after_inventory.get("comments"),
        "differences": differences,
    }


def _apply_pkg002_verification(review: dict, comments: dict, verification: dict) -> None:
    pkg_findings = [item for item in review["findings"] if item.get("rule_id") == "PKG-002"]
    if verification["passed"]:
        review["findings"] = [item for item in review["findings"] if item.get("rule_id") != "PKG-002"]
        allowed = {item["finding_id"] for item in review["findings"]}
        comments["findings"] = [item for item in comments.get("findings", []) if item.get("finding_id") in allowed]
    elif pkg_findings:
        pkg_findings[0].update(
            severity="ERROR",
            problem="审查副本无损性验证失败",
            actual=json.dumps(verification["differences"], ensure_ascii=False),
            suggestion="停止交付该副本，检查发生变化的对象或部件后重新生成。",
            evidence=verification,
        )
    for result in review["rule_results"]:
        if result.get("rule_id") == "PKG-002":
            result.update(
                status="PASS" if verification["passed"] else "ERROR",
                finding_count=0 if verification["passed"] else 1,
                evaluated_count=len(PRESERVED_INVENTORY_KEYS),
            )
            break
    severity = Counter(item["severity"] for item in review["findings"])
    review["statistics"]["total_findings"] = len(review["findings"])
    review["statistics"]["findings_by_severity"] = {
        name: severity.get(name, 0) for name in ("ERROR", "WARNING", "MANUAL_REVIEW")
    }
    statuses = Counter(item["status"] for item in review["rule_results"])
    review["statistics"]["rules_by_status"] = {
        name: statuses.get(name, 0) for name in ("PASS", "ERROR", "WARNING", "MANUAL_REVIEW")
    }
    comments["commented_finding_count"] = sum(bool(item.get("commented")) for item in comments.get("findings", []))
    comments["uncommented_finding_count"] = len(comments.get("findings", [])) - comments["commented_finding_count"]


def run_review(
    docx_path: str | Path,
    *,
    output: str | Path | None = None,
    report_output: str | Path | None = None,
    manifest_path: str | Path | None = None,
    rules_path: str | Path = DEFAULT_RULES,
    author: str = "HUST格式审查",
    initials: str = "HUST",
    template_name: str = DEFAULT_TEMPLATE_NAME,
) -> dict:
    source = Path(docx_path)
    output = Path(output) if output is not None else default_output_path(source)
    report_output = Path(report_output) if report_output is not None else default_report_path(source)
    if output.resolve() == report_output.resolve():
        raise ValueError("批注版与审查报告必须使用不同的输出路径")
    review = build_review(source, rules_path)
    comments = write_comments(
        source,
        review["findings"],
        output,
        author=author,
        initials=initials,
    )
    package_verification = _verify_commented_copy(source, output, comments)
    _apply_pkg002_verification(review, comments, package_verification)
    report = write_report(
        source,
        review,
        comments,
        output=report_output,
        template_name=template_name,
    )
    result = {
        "schema_version": "1.0.0-phase7",
        "capabilities": {
            "writes_word_comments": True,
            "writes_word_report": True,
            "comments_only_reliable_main_document_anchors": True,
        },
        "review": review,
        "comments": comments,
        "report": report,
        "package_verification": package_verification,
        "limitations": [
            "仅 ERROR/WARNING 且能可靠定位到主文档段落的发现写入批注。",
            "MANUAL_REVIEW、页/节级问题及页眉页脚部件问题不强行锚定。",
            "图、表、公式、目录、参考文献和脚注均已进入专项检查；语义与视觉事项仍明确标为人工复核。",
            "最终页码和其他渲染依赖版面现象无法仅凭 OOXML 稳定判断。",
        ],
    }
    if manifest_path is not None:
        manifest_path = Path(manifest_path)
        manifest_path.parent.mkdir(parents=True, exist_ok=True)
        manifest_path.write_text(
            json.dumps(result, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
    return result


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="执行 HUST 论文阶段7完整审查并生成批注副本与 Word 审查报告")
    parser.add_argument("docx", type=Path, help="待审查的源 DOCX")
    parser.add_argument("--rules", type=Path, default=DEFAULT_RULES, help="运行时 JSON 规则配置")
    parser.add_argument("--output", type=Path, help="批注版 DOCX 输出路径")
    parser.add_argument("--report-output", type=Path, help="审查报告 DOCX 输出路径")
    parser.add_argument("--manifest", type=Path, help="可选的完整审查与批注映射 JSON")
    parser.add_argument("--author", default="HUST格式审查", help="Word 批注作者")
    parser.add_argument("--initials", default="HUST", help="Word 批注作者缩写")
    parser.add_argument("--template-name", default=DEFAULT_TEMPLATE_NAME, help="报告中显示的标准模板名称")
    args = parser.parse_args(argv)
    result = run_review(
        args.docx,
        output=args.output,
        report_output=args.report_output,
        manifest_path=args.manifest,
        rules_path=args.rules,
        author=args.author,
        initials=args.initials,
        template_name=args.template_name,
    )
    if args.manifest is None:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
