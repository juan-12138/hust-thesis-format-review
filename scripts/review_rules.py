#!/usr/bin/env python3
"""Run the Phase 4 HUST thesis checks and export traceable JSON."""

from __future__ import annotations

import argparse
import hashlib
import json
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

from advanced_rule_engine import ADVANCED_RULE_IDS, AdvancedRuleEngine
from docx_model import DocxPackage
from rule_engine import PHASE4_RULE_IDS, RuleEngine, load_rule_config


SCRIPT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_RULES = SCRIPT_ROOT / "rules" / "hust_master_thesis_rules.json"


def build_review(docx_path: str | Path, rules_path: str | Path = DEFAULT_RULES) -> dict:
    docx_path = Path(docx_path)
    rules_path = Path(rules_path)
    source_bytes = docx_path.read_bytes()
    source_hash = hashlib.sha256(source_bytes).hexdigest().upper()
    config = load_rule_config(rules_path)
    package = DocxPackage(docx_path)
    base_review = RuleEngine(package, rules_path=rules_path).review()
    advanced_review = AdvancedRuleEngine(package, rules_path=rules_path).review()
    finding_objects = list(base_review.findings) + list(advanced_review.findings)
    result_by_id = {result.rule_id: result for result in (*base_review.rule_results, *advanced_review.rule_results)}
    findings = [finding.to_dict() for finding in finding_objects]
    rule_results = [result_by_id[rule["rule_id"]].to_dict() for rule in config["rules"]]
    finding_counts = Counter(finding["severity"] for finding in findings)
    result_counts = Counter(result["status"] for result in rule_results)
    return {
        "schema_version": "1.0.0-phase7",
        "source": {
            "path": str(docx_path.resolve()),
            "filename": docx_path.name,
            "size_bytes": len(source_bytes),
            "sha256": source_hash,
        },
        "created_at_utc": datetime.now(timezone.utc).isoformat(),
        "rule_set": {
            "rule_set_id": config.get("rule_set_id"),
            "schema_version": config.get("schema_version"),
            "runtime_rules_path": str(rules_path.resolve()),
            "runtime_mirror": config.get("runtime_mirror"),
        },
        "capabilities": {
            "read_only": True,
            "page_section_checks": True,
            "header_footer_checks": True,
            "body_font_color_checks": True,
            "body_paragraph_checks": True,
            "heading_format_numbering_checks": True,
            "abstract_keyword_checks": True,
            "toc_checks": True,
            "figure_checks": True,
            "table_checks": True,
            "equation_checks": True,
            "footnote_checks": True,
            "bibliography_checks": True,
            "special_object_checks": True,
            "language_structure_checks": True,
            "writes_word_comments": False,
            "writes_word_report": False,
            "render_dependent_layout_checked": False,
        },
        "coverage": {
            "implemented_rule_ids": [rule["rule_id"] for rule in config["rules"]],
            "implemented_rule_count": len(PHASE4_RULE_IDS) + len(ADVANCED_RULE_IDS),
            "total_rule_count": len(config["rules"]),
            "deferred_to_later_phases": [
                "依赖 Microsoft Word 或兼容渲染器的最终分页与视觉检查",
                "参考文献元数据真实性和学科语义判断",
            ],
        },
        "statistics": {
            "total_findings": len(findings),
            "findings_by_severity": {name: finding_counts.get(name, 0) for name in ("ERROR", "WARNING", "MANUAL_REVIEW")},
            "rules_by_status": {name: result_counts.get(name, 0) for name in ("PASS", "ERROR", "WARNING", "MANUAL_REVIEW")},
        },
        "rule_results": rule_results,
        "findings": findings,
        "limitations": [
            "OOXML 不能可靠给出最终页码、自动分页、空白页和视觉碰撞。",
            "PAGE-004、PARA-003、语义、事实真实性和复杂对象视觉效果只提供候选或人工复核项。",
            "PASS 且 evaluated_count 为 0 表示未发现适用对象，不等同于已证明该类内容合规。",
            "本 JSON 检查本身不修改源 DOCX；一键流程会另存批注版和 Word 审查报告。",
        ],
    }


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="执行 HUST 论文阶段4基础格式检查并导出 JSON")
    parser.add_argument("docx", type=Path, help="待审查的 DOCX 文件")
    parser.add_argument("--rules", type=Path, default=DEFAULT_RULES, help="运行时 JSON 规则配置")
    parser.add_argument("--output", type=Path, help="JSON 输出路径；省略时写入标准输出")
    args = parser.parse_args(argv)
    payload = build_review(args.docx, args.rules)
    text = json.dumps(payload, ensure_ascii=False, indent=2) + "\n"
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(text, encoding="utf-8")
    else:
        print(text, end="")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
