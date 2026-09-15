#!/usr/bin/env python3
"""Write Phase 4 findings as native Word comments with minimal OOXML edits."""

from __future__ import annotations

import argparse
import copy
import json
import os
import tempfile
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
COMMENTS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
COMMENTS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
NS = {"w": W_NS, "r": R_NS, "ct": CT_NS}
SEVERITY_ZH = {"ERROR": "错误", "WARNING": "警告", "MANUAL_REVIEW": "人工复核"}


def qn(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def default_output_path(source: str | Path) -> Path:
    source = Path(source)
    return source.with_name(f"{source.stem}_格式审查批注版.docx")


def _normalized_text(paragraph) -> str:
    pieces = []
    text_tags = {
        qn(W_NS, "t"),
        qn(W_NS, "delText"),
        qn(W_NS, "instrText"),
        qn(W_NS, "delInstrText"),
    }
    for node in paragraph.iter():
        if node.tag in text_tags:
            pieces.append(node.text or "")
        elif node.tag == qn(W_NS, "tab"):
            pieces.append("\t")
        elif node.tag == qn(W_NS, "br"):
            pieces.append("\n")
    return " ".join("".join(pieces).split())


def _skip_reason(finding: dict) -> str | None:
    if finding.get("severity") == "MANUAL_REVIEW":
        return "manual_review"
    if finding.get("severity") not in {"ERROR", "WARNING"}:
        return "unsupported_severity"
    anchor = finding.get("anchor")
    if not isinstance(anchor, dict):
        return "missing_anchor"
    if anchor.get("part_name") != "word/document.xml":
        return "unsupported_story_part"
    if not isinstance(anchor.get("paragraph_index"), int):
        return "invalid_paragraph_index"
    return None


def _format_finding(finding: dict) -> str:
    severity = finding.get("severity", "")
    lines = [
        f"【格式审查：{finding.get('rule_id', '')}】",
        f"问题：{finding.get('problem', '')}",
        f"当前：{finding.get('actual', '')}",
        f"要求：{finding.get('expected', '')}",
        f"依据：{finding.get('source', '')}",
        f"建议：{finding.get('suggestion', '')}",
        f"严重程度：{SEVERITY_ZH.get(severity, severity)}",
    ]
    return "\n".join(lines)


def _append_comment(comments_root, comment_id: int, findings: list[dict], *, author: str, initials: str, created_at: str) -> None:
    comment = etree.SubElement(
        comments_root,
        qn(W_NS, "comment"),
        {
            qn(W_NS, "id"): str(comment_id),
            qn(W_NS, "author"): author,
            qn(W_NS, "initials"): initials,
            qn(W_NS, "date"): created_at,
        },
    )
    for index, finding in enumerate(findings):
        paragraph = etree.SubElement(comment, qn(W_NS, "p"))
        run = etree.SubElement(paragraph, qn(W_NS, "r"))
        for line_index, line in enumerate(_format_finding(finding).splitlines()):
            if line_index:
                etree.SubElement(run, qn(W_NS, "br"))
            text = etree.SubElement(run, qn(W_NS, "t"))
            text.text = line
        if index != len(findings) - 1:
            etree.SubElement(comment, qn(W_NS, "p"))


def _comment_reference_run(comment_id: int):
    run = etree.Element(qn(W_NS, "r"))
    properties = etree.SubElement(run, qn(W_NS, "rPr"))
    etree.SubElement(properties, qn(W_NS, "rStyle"), {qn(W_NS, "val"): "CommentReference"})
    etree.SubElement(run, qn(W_NS, "commentReference"), {qn(W_NS, "id"): str(comment_id)})
    return run


def _anchor_paragraph(paragraph, comment_id: int, findings: list[dict]) -> str:
    ranges = {
        (anchor.get("run_start"), anchor.get("run_end"))
        for finding in findings
        if isinstance((anchor := finding.get("anchor")), dict)
        and isinstance(anchor.get("run_start"), int)
        and isinstance(anchor.get("run_end"), int)
    }
    has_paragraph_scope = any(
        not isinstance(finding.get("anchor", {}).get("run_start"), int)
        for finding in findings
    )
    if len(ranges) == 1 and not has_paragraph_scope:
        run_range = next(iter(ranges))
        runs = paragraph.xpath(".//w:r", namespaces=NS)
        start_index, end_index = run_range
        if 0 <= start_index <= end_index < len(runs):
            selected = runs[start_index : end_index + 1]
            if selected and all(run.getparent() is paragraph for run in selected):
                first_position = paragraph.index(selected[0])
                last_position = paragraph.index(selected[-1])
                paragraph.insert(first_position, etree.Element(qn(W_NS, "commentRangeStart"), {qn(W_NS, "id"): str(comment_id)}))
                end = etree.Element(qn(W_NS, "commentRangeEnd"), {qn(W_NS, "id"): str(comment_id)})
                paragraph.insert(last_position + 2, end)
                paragraph.insert(paragraph.index(end) + 1, _comment_reference_run(comment_id))
                return "run_range"

    start = etree.Element(qn(W_NS, "commentRangeStart"), {qn(W_NS, "id"): str(comment_id)})
    start_position = 1 if len(paragraph) and paragraph[0].tag == qn(W_NS, "pPr") else 0
    paragraph.insert(start_position, start)
    end = etree.Element(qn(W_NS, "commentRangeEnd"), {qn(W_NS, "id"): str(comment_id)})
    paragraph.append(end)
    paragraph.append(_comment_reference_run(comment_id))
    return "paragraph"


def _next_comment_id(document_root, comments_root) -> int:
    ids = []
    values = document_root.xpath(
        ".//w:commentRangeStart/@w:id | .//w:commentRangeEnd/@w:id | .//w:commentReference/@w:id",
        namespaces=NS,
    )
    values.extend(comments_root.xpath("./w:comment/@w:id", namespaces=NS))
    for value in values:
        try:
            ids.append(int(value))
        except (TypeError, ValueError):
            continue
    return max(ids, default=-1) + 1


def _ensure_comments_relationship(rels_root) -> None:
    existing = rels_root.xpath("./r:Relationship[@Type=$type]", namespaces=NS, type=COMMENTS_REL_TYPE)
    if existing:
        return
    used = {node.get("Id") for node in rels_root.xpath("./r:Relationship", namespaces=NS)}
    number = 1
    while f"rId{number}" in used:
        number += 1
    etree.SubElement(
        rels_root,
        qn(R_NS, "Relationship"),
        {"Id": f"rId{number}", "Type": COMMENTS_REL_TYPE, "Target": "comments.xml"},
    )


def _ensure_comments_content_type(content_types_root) -> None:
    existing = content_types_root.xpath("./ct:Override[@PartName='/word/comments.xml']", namespaces=NS)
    if not existing:
        etree.SubElement(
            content_types_root,
            qn(CT_NS, "Override"),
            {"PartName": "/word/comments.xml", "ContentType": COMMENTS_CONTENT_TYPE},
        )


def _xml_bytes(root) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _new_zip_info(name: str, template: ZipInfo | None = None) -> ZipInfo:
    if template is None:
        info = ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
        info.compress_type = ZIP_DEFLATED
        return info
    info = copy.copy(template)
    info.filename = name
    info.orig_filename = name
    return info


def _write_package(source: Path, output: Path, replacements: dict[str, bytes]) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(prefix=f".{output.stem}.", suffix=".tmp", dir=output.parent, delete=False) as handle:
        temporary = Path(handle.name)
    try:
        with ZipFile(source, "r") as before, ZipFile(temporary, "w") as after:
            names = set()
            template = None
            for info in before.infolist():
                names.add(info.filename)
                if info.filename == "word/document.xml":
                    template = info
                after.writestr(info, replacements.get(info.filename, before.read(info.filename)))
            for name, data in replacements.items():
                if name not in names:
                    after.writestr(_new_zip_info(name, template), data)
        os.replace(temporary, output)
    except Exception:
        temporary.unlink(missing_ok=True)
        raise


def write_comments(
    source: str | Path,
    findings: list[dict],
    output: str | Path | None = None,
    *,
    author: str = "HUST格式审查",
    initials: str = "HUST",
    created_at: str | None = None,
) -> dict:
    """Create a commented DOCX copy and return a finding-to-comment manifest."""
    source = Path(source).resolve()
    output = Path(output or default_output_path(source)).resolve()
    if source == output:
        raise ValueError("输出文件不能覆盖源 DOCX")
    if source.suffix.lower() != ".docx" or not source.is_file():
        raise ValueError("源文件必须是存在的 DOCX 文件")
    created_at = created_at or datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")

    with ZipFile(source, "r") as package:
        names = set(package.namelist())
        required = {"word/document.xml", "word/_rels/document.xml.rels", "[Content_Types].xml"}
        missing = sorted(required - names)
        if missing:
            raise ValueError(f"DOCX 缺少必要部件：{', '.join(missing)}")
        document_root = etree.fromstring(package.read("word/document.xml"))
        rels_root = etree.fromstring(package.read("word/_rels/document.xml.rels"))
        content_types_root = etree.fromstring(package.read("[Content_Types].xml"))
        if "word/comments.xml" in names:
            comments_root = etree.fromstring(package.read("word/comments.xml"))
        else:
            comments_root = etree.Element(qn(W_NS, "comments"), nsmap={"w": W_NS})

    paragraphs = document_root.xpath(".//w:p", namespaces=NS)
    records = []
    groups: dict[int, list[tuple[int, dict]]] = defaultdict(list)
    for index, finding in enumerate(findings):
        record = {
            "finding_id": finding.get("finding_id"),
            "rule_id": finding.get("rule_id"),
            "commented": False,
            "comment_id": None,
            "anchor_mode": None,
            "skip_reason": None,
        }
        reason = _skip_reason(finding)
        anchor = finding.get("anchor") or {}
        paragraph_index = anchor.get("paragraph_index")
        if reason is None and not 0 <= paragraph_index < len(paragraphs):
            reason = "paragraph_not_found"
        if reason is None:
            snippet = " ".join(str(anchor.get("snippet") or "").split())
            if snippet and snippet not in _normalized_text(paragraphs[paragraph_index]):
                reason = "snippet_mismatch"
        if reason is None:
            groups[paragraph_index].append((index, finding))
        else:
            record["skip_reason"] = reason
        records.append(record)

    next_id = _next_comment_id(document_root, comments_root)
    for paragraph_index in sorted(groups):
        indexed_findings = groups[paragraph_index]
        comment_id = next_id
        next_id += 1
        grouped_findings = [finding for _, finding in indexed_findings]
        anchor_mode = _anchor_paragraph(paragraphs[paragraph_index], comment_id, grouped_findings)
        _append_comment(comments_root, comment_id, grouped_findings, author=author, initials=initials, created_at=created_at)
        for finding_index, _ in indexed_findings:
            records[finding_index].update(
                commented=True,
                comment_id=comment_id,
                anchor_mode=anchor_mode,
                skip_reason=None,
            )

    if groups:
        _ensure_comments_relationship(rels_root)
        _ensure_comments_content_type(content_types_root)
        replacements = {
            "word/document.xml": _xml_bytes(document_root),
            "word/comments.xml": _xml_bytes(comments_root),
            "word/_rels/document.xml.rels": _xml_bytes(rels_root),
            "[Content_Types].xml": _xml_bytes(content_types_root),
        }
        modified_parts = sorted(replacements)
    else:
        replacements = {}
        modified_parts = []
    _write_package(source, output, replacements)

    return {
        "schema_version": "1.0.0-phase5",
        "source": str(source),
        "output": str(output),
        "comment_count": len(groups),
        "commented_finding_count": sum(record["commented"] for record in records),
        "uncommented_finding_count": sum(not record["commented"] for record in records),
        "modified_parts": modified_parts,
        "findings": records,
    }


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="将阶段4审查发现写为 Word 原生批注")
    parser.add_argument("docx", type=Path, help="待批注的源 DOCX")
    parser.add_argument("--findings", type=Path, required=True, help="阶段4审查 JSON")
    parser.add_argument("--output", type=Path, help="输出 DOCX；省略时使用默认批注版文件名")
    parser.add_argument("--manifest", type=Path, help="可选的批注映射 JSON")
    args = parser.parse_args(argv)
    payload = json.loads(args.findings.read_text(encoding="utf-8"))
    findings = payload.get("findings") if isinstance(payload, dict) else payload
    if not isinstance(findings, list):
        raise ValueError("输入 JSON 必须包含 findings 列表")
    manifest = write_comments(args.docx, findings, args.output)
    text = json.dumps(manifest, ensure_ascii=False, indent=2) + "\n"
    if args.manifest:
        args.manifest.parent.mkdir(parents=True, exist_ok=True)
        args.manifest.write_text(text, encoding="utf-8")
    else:
        print(text, end="")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
