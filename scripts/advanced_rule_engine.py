"""Phase 7 advanced-object checks for HUST master-thesis DOCX files.

The checks deliberately separate structural evidence from render- or
meaning-dependent judgements.  A MANUAL_REVIEW finding is a first-class result,
not a guessed failure or an implicit pass.
"""

from __future__ import annotations

import json
import re
from pathlib import Path

from docx_model import DocxPackage, NS, ParagraphRecord, qn
from effective_style import EffectiveStyleResolver
from rule_engine import Finding, ReviewResult, RuleResult, load_rule_config
from structure_classifier import StructureClassifier


ADVANCED_RULE_IDS = (
    "PKG-001", "PKG-002", "COVER-001", "COVER-002", "COVER-003", "FRONT-001",
    "ABS-001", "ABS-002", "ABS-003", "ABS-004", "ABS-005", "ABS-006",
    "TOC-001", "TOC-002", "TOC-003",
    "FIG-001", "FIG-002", "FIG-003", "FIG-004", "FIG-005",
    "TAB-001", "TAB-002", "TAB-003", "TAB-004", "TAB-005", "TAB-006",
    "EQ-001", "EQ-002",
    "NOTE-001", "NOTE-002", "NOTE-003", "NOTE-004",
    "REF-001", "REF-002", "REF-003", "REF-004", "REF-005", "REF-006",
    "REF-007", "REF-008", "REF-009", "REF-010", "REF-011",
    "OBJ-001", "OBJ-002", "OBJ-003",
    "LANG-001", "LANG-002", "LANG-003",
    "STRUCT-001", "STRUCT-002", "APP-001", "APP-002", "APP-003",
    "RENDER-001", "RENDER-002", "RENDER-003",
)

CAPTION_PUNCTUATION = "。！？；;,.，：:、"
REF_FINAL_PUNCTUATION = "。！？；;,.，：:"


class AdvancedRuleEngine:
    """Evaluate all rules not handled by the Phase 4 foundation engine."""

    def __init__(self, package: DocxPackage, *, rules_path: str | Path):
        self.package = package
        config = load_rule_config(rules_path)
        self.rules = {rule["rule_id"]: rule for rule in config["rules"]}
        self.resolver = EffectiveStyleResolver(package)
        self.classifier = StructureClassifier(package, self.resolver)
        self._sequence = 0
        self._checked_counts: dict[str, int] = {}

    def review(self) -> ReviewResult:
        paragraphs = list(self.package.iter_paragraphs())
        findings: list[Finding] = []
        findings.extend(self._check_package())
        findings.extend(self._check_cover_and_front(paragraphs))
        findings.extend(self._check_abstracts(paragraphs))
        findings.extend(self._check_toc(paragraphs))
        findings.extend(self._check_figures(paragraphs))
        findings.extend(self._check_tables(paragraphs))
        findings.extend(self._check_equations(paragraphs))
        findings.extend(self._check_notes())
        findings.extend(self._check_references(paragraphs))
        findings.extend(self._check_objects(paragraphs))
        findings.extend(self._check_language_and_structure(paragraphs))
        findings.extend(self._check_appendices(paragraphs))
        findings.extend(self._render_boundaries())
        return ReviewResult(tuple(self._summarize(findings)), tuple(findings))

    def _mark(self, rule_id: str, count: int = 1):
        self._checked_counts[rule_id] = self._checked_counts.get(rule_id, 0) + count

    def _summarize(self, findings: list[Finding]) -> list[RuleResult]:
        priority = {"PASS": 0, "MANUAL_REVIEW": 1, "WARNING": 2, "ERROR": 3}
        results = []
        for rule_id in ADVANCED_RULE_IDS:
            matched = [item for item in findings if item.rule_id == rule_id]
            status = max((item.severity for item in matched), key=priority.get) if matched else "PASS"
            results.append(RuleResult(rule_id, status, len(matched), self._checked_counts.get(rule_id, 0)))
        return results

    def _finding(self, rule_id: str, *, problem: str, actual: str, location: dict,
                 evidence: dict, suggestion: str, anchor: dict | None = None,
                 severity: str | None = None) -> Finding:
        rule = self.rules[rule_id]
        self._sequence += 1
        return Finding(
            finding_id=f"{rule_id}-A{self._sequence:04d}",
            rule_id=rule_id,
            category=rule.get("category_zh", rule["category"]),
            object_type=rule.get("object_type_zh", rule["object_type"]),
            severity=severity or rule["severity"],
            problem=problem,
            actual=actual,
            expected=rule.get("expected_value_zh", json.dumps(rule["expected_value"], ensure_ascii=False)),
            source=rule.get("source_zh", rule["source"]),
            suggestion=suggestion,
            location=location,
            anchor=anchor,
            evidence=evidence,
        )

    @staticmethod
    def _location(paragraph: ParagraphRecord | None = None, *, scope="document", label="全文") -> dict:
        if paragraph is None:
            return {"scope": scope, "label": label}
        return {
            "scope": "paragraph",
            "part_name": paragraph.part_name,
            "section_index": paragraph.section_index,
            "paragraph_index": paragraph.index,
            "label": paragraph.location,
            "snippet": " ".join(paragraph.text.split())[:40],
        }

    @staticmethod
    def _anchor(paragraph: ParagraphRecord | None) -> dict | None:
        if paragraph is None or paragraph.part_name != DocxPackage.MAIN_PART:
            return None
        return {
            "part_name": paragraph.part_name,
            "paragraph_index": paragraph.index,
            "snippet": " ".join(paragraph.text.split())[:40],
        }

    def _manual(self, rule_id: str, *, problem: str, actual: str, suggestion: str,
                paragraph: ParagraphRecord | None = None, evidence: dict | None = None) -> Finding:
        self._mark(rule_id)
        return self._finding(
            rule_id, severity="MANUAL_REVIEW", problem=problem, actual=actual,
            location=self._location(paragraph), evidence=evidence or {"determinism": "manual"},
            suggestion=suggestion, anchor=self._anchor(paragraph),
        )

    @staticmethod
    def _paragraph_map(paragraphs: list[ParagraphRecord]):
        return {id(paragraph.element): paragraph for paragraph in paragraphs}

    def _main_matter_bounds(self, paragraphs: list[ParagraphRecord]) -> tuple[int, int]:
        """Return paragraph-index bounds for the numbered thesis chapters.

        Numbering is often stored in Word list definitions rather than visible
        text, so the boundary must use classified heading roles instead of a
        regular expression that expects a literal chapter number.
        """
        classified = self.classifier.classify_document(paragraphs)
        last_toc = max(
            (item.paragraph.index for item in classified if item.classification.role == "toc_entry"),
            default=-1,
        )
        start = next(
            (
                item.paragraph.index
                for item in classified
                if item.paragraph.index > last_toc and item.classification.role == "heading_1"
            ),
            0,
        )
        end_roles = {"references_title", "acknowledgements_title", "appendix_heading"}
        end = next(
            (
                item.paragraph.index
                for item in classified
                if item.paragraph.index > start and item.classification.role in end_roles
            ),
            len(paragraphs),
        )
        return start, end

    @staticmethod
    def _sibling_paragraph(element, mapping, *, preceding=False):
        sibling = element.getprevious() if preceding else element.getnext()
        while sibling is not None:
            if sibling.tag == qn("w", "p"):
                return mapping.get(id(sibling))
            if sibling.tag == qn("w", "tbl"):
                return None
            sibling = sibling.getprevious() if preceding else sibling.getnext()
        return None

    def _check_package(self) -> list[Finding]:
        self._mark("PKG-001", len(self.package.part_names))
        return [self._manual(
            "PKG-002",
            problem="审查副本的无损性需在写入批注后执行前后清单比较",
            actual="当前仅完成源 DOCX 的只读结构清点；输出副本将在一键流程末端验证",
            suggestion="保留生成清单并核对除批注相关部件外的所有关系和对象计数。",
            evidence={"source_part_count": len(self.package.part_names), "stage": "pre-write"},
        )]

    def _check_cover_and_front(self, paragraphs: list[ParagraphRecord]) -> list[Finding]:
        findings = []
        early = [p for p in paragraphs[:80] if p.text.strip()]
        for rule_id in ("COVER-001", "COVER-002", "COVER-003"):
            self._mark(rule_id, len(early))
        chinese_candidates = [p for p in early if 8 <= len(re.findall(r"[\u4e00-\u9fff]", p.text)) <= 45]
        english_candidates = [p for p in early if len(re.findall(r"[A-Za-z]+", p.text)) >= 4]
        if chinese_candidates:
            p = chinese_candidates[0]
            count = len(re.findall(r"[\u4e00-\u9fff]", p.text))
            if count > 30:
                findings.append(self._finding("COVER-001", problem="中文封面题目超过建议字数上限",
                    actual=f"检测到 {count} 个汉字", location=self._location(p), evidence={"character_count": count},
                    suggestion="精简题目至 30 个汉字以内，或确认院系批准的例外。", anchor=self._anchor(p)))
        if english_candidates:
            p = english_candidates[0]
            findings.append(self._manual("COVER-002", problem="英文题目的实词首字母大写需结合题意复核",
                actual=p.text[:120], suggestion="核对专有名词、冠词和介词的大小写。", paragraph=p))
        findings.append(self._manual("COVER-003", problem="英文页作者姓名顺序依赖作者身份",
            actual="已定位封面候选区域，无法仅凭文本确认姓与名", suggestion="核对中文姓名是否姓在前、名在后，并按要求大写姓氏。"))
        required = ("独创性声明", "学位论文版权使用授权书")
        missing = [name for name in required if not any(name in p.text for p in early)]
        self._mark("FRONT-001", len(early))
        if missing:
            findings.append(self._finding("FRONT-001", problem="前置声明页可能不完整",
                actual="未定位到：" + "、".join(missing), location=self._location(scope="front_matter", label="前置部分"),
                evidence={"missing_headings": missing}, suggestion="核对并补齐模板要求的声明、授权、签名和日期区域。"))
        return findings

    @staticmethod
    def _section_between(paragraphs, start_pattern, stop_patterns):
        start = next((i for i, p in enumerate(paragraphs) if re.fullmatch(start_pattern, p.text.strip(), re.I)), None)
        if start is None:
            return [], None
        end = len(paragraphs)
        for index in range(start + 1, len(paragraphs)):
            if any(re.fullmatch(pattern, paragraphs[index].text.strip(), re.I) for pattern in stop_patterns):
                end = index
                break
        return paragraphs[start + 1:end], paragraphs[start]

    def _check_abstracts(self, paragraphs: list[ParagraphRecord]) -> list[Finding]:
        findings = []
        # Real theses commonly insert spaces in Chinese display headings and
        # append Word field-update instructions to the TOC title.  Those are
        # still section boundaries and must not be absorbed into the abstract.
        zh_region, zh_heading = self._section_between(
            paragraphs,
            r"\s*摘\s*要\s*",
            (r"\s*Abstract\s*", r"\s*目\s*录.*"),
        )
        en_region, en_heading = self._section_between(
            paragraphs,
            r"\s*Abstract\s*",
            (r"\s*Contents?.*", r"\s*目\s*录.*", r"\s*第?\s*1\s*[章.．、\s].*", r"\s*1(?:\.0)?\s+.*"),
        )
        zh_text = "".join(p.text for p in zh_region if not re.match(r"^关键词\s*[：:]", p.text.strip()))
        self._mark("ABS-001", len(zh_region))
        if zh_heading and not 500 <= len(re.findall(r"[\u4e00-\u9fff]", zh_text)) <= 600:
            count = len(re.findall(r"[\u4e00-\u9fff]", zh_text))
            findings.append(self._finding("ABS-001", problem="中文摘要长度不在建议范围",
                actual=f"约 {count} 个汉字", location=self._location(zh_heading), evidence={"han_character_count": count},
                suggestion="在不削弱完整性的前提下调整至约 500–600 个汉字。", anchor=self._anchor(zh_heading)))
        for rule_id, region, expected_multiple in (("ABS-002", zh_region, None), ("ABS-003", en_region, 1.5)):
            self._mark(rule_id, len(region))
            for p in region:
                if not p.text.strip():
                    continue
                pformat = self.resolver.get_effective_paragraph_format(p)
                bad = []
                for run in self.package.iter_runs(p):
                    if not run.text.strip():
                        continue
                    fmt = self.resolver.get_effective_run_format(run, p)
                    if fmt.get("size_pt") is not None and abs(fmt["size_pt"] - 12) > 0.2:
                        bad.append(f"字号 {fmt['size_pt']:g} 磅")
                if expected_multiple and pformat.get("line_spacing_multiple") is not None and abs(pformat["line_spacing_multiple"] - expected_multiple) > 0.05:
                    bad.append(f"行距 {pformat['line_spacing_multiple']:.2f} 倍")
                if bad:
                    findings.append(self._finding(rule_id, problem="摘要文字格式不符合规则",
                        actual="、".join(dict.fromkeys(bad)), location=self._location(p),
                        evidence={"paragraph_format": pformat}, suggestion="按摘要对应语言设置字体、12 磅和规定行距。",
                        anchor=self._anchor(p)))
        keyword = next((p for p in paragraphs if re.match(r"^\s*(关键词|Key\s*words?)\s*[：:]", p.text, re.I)), None)
        self._mark("ABS-004", 1 if keyword else 0)
        if keyword:
            value = re.split(r"[：:]", keyword.text, maxsplit=1)[-1].strip()
            tokens = [x.strip() for x in re.split(r"[；;]", value.rstrip(CAPTION_PUNCTUATION)) if x.strip()]
            problems = []
            if not 3 <= len(tokens) <= 8:
                problems.append(f"关键词数量为 {len(tokens)}")
            if ";" in value:
                problems.append("使用了英文分号")
            if value and value[-1] in CAPTION_PUNCTUATION:
                problems.append("末尾带标点")
            if problems:
                findings.append(self._finding("ABS-004", problem="关键词结构不符合要求", actual="、".join(problems),
                    location=self._location(keyword), evidence={"tokens": tokens, "raw": value},
                    suggestion="保留 3–8 个关键词，用中文分号分隔，末尾不加标点。", anchor=self._anchor(keyword)))
        findings.append(self._manual("ABS-005", problem="中英文摘要与关键词的语义对应关系无法仅凭格式稳定判定",
            actual=f"中文摘要={'已定位' if zh_heading else '未定位'}；英文摘要={'已定位' if en_heading else '未定位'}",
            suggestion="逐项核对题意、术语和关键词对应关系。", paragraph=en_heading or zh_heading))
        self._mark("ABS-006", len(zh_region) + len(en_region))
        for p in zh_region + en_region:
            candidates = re.findall(r"\b(?:我们|本文作者|I|we|our)\b", p.text, re.I)
            has_objects = bool(p.element.xpath(".//w:drawing | .//w:pict | .//m:oMath", namespaces=NS))
            if candidates or has_objects:
                findings.append(self._finding("ABS-006", problem="摘要中发现需复核的内容形式",
                    actual=("第一人称：" + ", ".join(candidates) if candidates else "") + ("；含图形或公式对象" if has_objects else ""),
                    location=self._location(p), evidence={"first_person": candidates, "has_objects": has_objects},
                    suggestion="改为客观表述，并确认摘要具有自含性且不依赖图表公式。", anchor=self._anchor(p)))
        return findings

    def _check_toc(self, paragraphs: list[ParagraphRecord]) -> list[Finding]:
        findings = []
        root = self.package.xml(DocxPackage.MAIN_PART)
        field_codes = " ".join(root.xpath(".//w:instrText/text() | .//w:fldSimple/@w:instr", namespaces=NS))
        toc_entries = [p for p in paragraphs if (p.style_id or "").upper().startswith("TOC")]
        self._mark("TOC-001", 1 + len(toc_entries))
        if not re.search(r"\bTOC\b", field_codes, re.I):
            p = toc_entries[0] if toc_entries else next((x for x in paragraphs if x.text.strip() == "目录"), None)
            findings.append(self._finding("TOC-001", problem="未检测到 Word 自动目录 TOC 域",
                actual=f"目录样式段落 {len(toc_entries)} 个，但无 TOC 域代码", location=self._location(p),
                evidence={"field_codes": field_codes[:300], "toc_entry_count": len(toc_entries)},
                suggestion="使用 Word 的自动目录功能，并将收录层级限制为二级标题。", anchor=self._anchor(p)))
        self._mark("TOC-002", len(toc_entries))
        for p in toc_entries:
            sizes = []
            for run in self.package.iter_runs(p):
                fmt = self.resolver.get_effective_run_format(run, p)
                if run.text.strip() and fmt.get("size_pt") is not None:
                    sizes.append(fmt["size_pt"])
            pformat = self.resolver.get_effective_paragraph_format(p)
            if any(abs(size - 14) > 0.2 for size in sizes) or (pformat.get("line_spacing_pt") is not None and abs(pformat["line_spacing_pt"] - 25) > 0.5):
                findings.append(self._finding("TOC-002", problem="目录条目字号或行距不符合文字规则",
                    actual=f"字号={sorted(set(sizes)) or '无法解析'}，行距={pformat.get('line_spacing_pt') or pformat.get('line_spacing_multiple')}",
                    location=self._location(p), evidence={"sizes_pt": sizes, "paragraph_format": pformat},
                    suggestion="设置四号（14 磅）和固定值 25 磅行距。", anchor=self._anchor(p)))
        if toc_entries or "TOC" in field_codes.upper():
            findings.append(self._manual("TOC-003", problem="目录最终页码及域结果需要更新后复核",
                actual="已检测目录结构，但 OOXML 不提供可靠的最终分页结果", suggestion="在 Word 中更新全部域后逐项核对目录文字、书签和页码。",
                paragraph=toc_entries[0] if toc_entries else None))
        else:
            self._mark("TOC-003", 0)
        return findings

    def _check_figures(self, paragraphs: list[ParagraphRecord]) -> list[Finding]:
        findings = []
        start, end = self._main_matter_bounds(paragraphs)
        figures = [
            p for p in paragraphs
            if start <= p.index < end
            and not p.in_table
            and p.element.xpath(".//w:drawing | .//w:pict", namespaces=NS)
        ]
        caption_re = re.compile(r"^\s*图\s*(\d+)\s*[-－.]\s*(\d+)\s*.*")
        for rule_id in ("FIG-001", "FIG-002", "FIG-003", "FIG-004", "FIG-005"):
            self._mark(rule_id, len(figures))
        for fig in figures:
            position = paragraphs.index(fig)
            before = paragraphs[position - 1] if position else None
            after = paragraphs[position + 1] if position + 1 < len(paragraphs) else None
            before_match = caption_re.match(before.text) if before else None
            after_match = caption_re.match(after.text) if after else None
            if before_match and not after_match:
                findings.append(self._finding("FIG-002", problem="图题位于图片上方",
                    actual=f"检测到上方图题：{before.text}", location=self._location(before),
                    evidence={"figure_paragraph_index": fig.index, "caption_position": "before"},
                    suggestion="将图题移至图片下方并与图片相邻。", anchor=self._anchor(before)))
            elif not after_match:
                findings.append(self._finding("FIG-002", problem="图片下方未检测到规范图题",
                    actual="相邻下方段落不是“图章号-序号”格式", location=self._location(fig),
                    evidence={"next_text": after.text[:100] if after else None},
                    suggestion="在图片下方添加居中的章编号图题。", anchor=self._anchor(fig)))
            caption = after if after_match else before if before_match else None
            if caption:
                fmt = self.resolver.get_effective_paragraph_format(caption)
                if fmt.get("alignment") != "center" or caption.text.rstrip().endswith(tuple(CAPTION_PUNCTUATION)):
                    findings.append(self._finding("FIG-002", problem="图题对齐或末尾标点不符合要求",
                        actual=f"对齐={fmt.get('alignment')}；末字符={caption.text.rstrip()[-1:]}", location=self._location(caption),
                        evidence={"paragraph_format": fmt}, suggestion="图题居中，末尾不加标点。", anchor=self._anchor(caption)))
                number = f"{caption_re.match(caption.text).group(1)}-{caption_re.match(caption.text).group(2)}"
                body_refs = [p for p in paragraphs if p is not caption and re.search(rf"图\s*{re.escape(number)}\b", p.text)]
                if not body_refs:
                    findings.append(self._finding("FIG-004", problem="未在正文中定位到该图编号的引用",
                        actual=f"图 {number} 仅出现在图题候选中", location=self._location(caption), evidence={"number": number},
                        suggestion="在相关正文中引用该图，并核对引用图片的来源标注。", anchor=self._anchor(caption)))
            drawing_kind = "浮动" if fig.element.xpath(".//wp:anchor", namespaces=NS) else "嵌入"
            if drawing_kind == "浮动":
                findings.append(self._manual("FIG-001", problem="浮动图片的最终位置和越界需渲染复核",
                    actual="检测到浮动 DrawingML 对象", suggestion="渲染后确认居中、未超出版芯且未遮挡正文。", paragraph=fig))
            findings.append(self._manual("FIG-003", problem="栅格图内文字无法从 OOXML 可靠读取",
                actual="图片对象已清点，图内文字需查看原始像素", suggestion="人工确认图中文字不大于五号且缩放后清晰。", paragraph=fig))
            findings.append(self._manual("FIG-005", problem="图片可读性与科学性需要视觉和专业判断",
                actual="已定位图片对象", suggestion="检查清晰度、自明性、比例尺、色标与风格一致性。", paragraph=fig))
        return findings

    def _check_tables(self, paragraphs: list[ParagraphRecord]) -> list[Finding]:
        findings = []
        root = self.package.xml(DocxPackage.MAIN_PART)
        mapping = self._paragraph_map(paragraphs)
        start, end = self._main_matter_bounds(paragraphs)
        tables = []
        for table in root.xpath(".//w:tbl", namespaces=NS):
            first = next((mapping.get(id(node)) for node in table.xpath(".//w:p", namespaces=NS) if mapping.get(id(node))), None)
            if first is not None and start <= first.index < end:
                tables.append(table)
        cap_re = re.compile(r"^\s*表\s*(\d+)\s*[-－.]\s*(\d+)\s*.*")
        for rule_id in ("TAB-001", "TAB-002", "TAB-003", "TAB-004", "TAB-005", "TAB-006"):
            self._mark(rule_id, len(tables))
        for index, table in enumerate(tables, 1):
            before = self._sibling_paragraph(table, mapping, preceding=True)
            after = self._sibling_paragraph(table, mapping)
            cap = before if before and cap_re.match(before.text) else None
            if cap is None:
                misplaced = after if after and cap_re.match(after.text) else None
                findings.append(self._finding("TAB-002", problem="表题缺失或位于表格下方",
                    actual=(f"检测到下方表题：{misplaced.text}" if misplaced else "表格上方相邻段落不是规范表题"),
                    location=self._location(misplaced or before), evidence={"table_index": index, "caption_position": "after" if misplaced else "missing"},
                    suggestion="将按章编号的表题放在表格上方并居中，末尾不加标点。", anchor=self._anchor(misplaced or before)))
            elif cap.text.rstrip().endswith(tuple(CAPTION_PUNCTUATION)) or self.resolver.get_effective_paragraph_format(cap).get("alignment") != "center":
                findings.append(self._finding("TAB-002", problem="表题对齐或末尾标点不符合要求", actual=cap.text,
                    location=self._location(cap), evidence={"table_index": index}, suggestion="表题居中且末尾不加标点。", anchor=self._anchor(cap)))
            border_nodes = table.xpath("./w:tblPr/w:tblBorders/* | .//w:tcPr/w:tcBorders/*", namespaces=NS)
            borders = [(node.tag.rsplit("}", 1)[-1], node.get(qn("w", "val"), "single")) for node in border_nodes]
            active_vertical = [name for name, val in borders if name in {"left", "right", "insideV"} and val not in {"nil", "none"}]
            header_bottom = table.xpath("./w:tr[1]/w:tc/w:tcPr/w:tcBorders/w:bottom[not(@w:val='nil' or @w:val='none')]", namespaces=NS)
            top = any(name == "top" and val not in {"nil", "none"} for name, val in borders)
            bottom = any(name == "bottom" and val not in {"nil", "none"} for name, val in borders)
            problems = []
            if active_vertical:
                problems.append("检测到竖线：" + ",".join(sorted(set(active_vertical))))
            if not (top and bottom and header_bottom):
                problems.append("缺少完整的表顶线、表头下分隔线或表底线")
            if problems:
                anchor = cap or next((mapping.get(id(p)) for p in table.xpath(".//w:p", namespaces=NS) if mapping.get(id(p))), None)
                findings.append(self._finding("TAB-001", problem="表格不符合三线表结构",
                    actual="；".join(problems), location=self._location(anchor, scope="table", label=f"第 {index} 个表格"),
                    evidence={"borders": borders, "header_bottom_count": len(header_bottom)},
                    suggestion="保留表顶线、表头下分隔线和表底线，删除通常不需要的竖线和内部横线。", anchor=self._anchor(anchor)))
            cell_paragraphs = [mapping[id(p)] for p in table.xpath(".//w:p", namespaces=NS) if id(p) in mapping]
            text_format_issues = []
            for p in cell_paragraphs:
                for run in self.package.iter_runs(p):
                    if not run.text.strip():
                        continue
                    fmt = self.resolver.get_effective_run_format(run, p)
                    if fmt.get("size_pt") is not None and abs(fmt["size_pt"] - 10.5) > 0.2:
                        text_format_issues.append({"paragraph": p, "size_pt": fmt["size_pt"], "text": run.text[:30], "format": fmt})
                        break
            if text_format_issues:
                samples = "；".join(
                    f"{item['size_pt']:g} 磅：{item['text']}" for item in text_format_issues[:4]
                )
                anchor = cap or text_format_issues[0]["paragraph"]
                findings.append(self._finding("TAB-003", problem="表中文字字号不符合要求",
                    actual=f"{len(text_format_issues)} 个单元格段落不符；示例：{samples}",
                    location=self._location(anchor),
                    evidence={"table_index": index, "issue_count": len(text_format_issues),
                              "examples": [{"size_pt": x["size_pt"], "text": x["text"]} for x in text_format_issues[:10]]},
                    suggestion="将表中文字设置为五号（10.5 磅），并分别核对中西文字体。", anchor=self._anchor(anchor)))
            first_row = table.find("w:tr", NS)
            if first_row is not None:
                header_texts = ["".join(cell.xpath(".//w:t/text()", namespaces=NS)).strip() for cell in first_row.findall("w:tc", NS)]
                if any(text.endswith(tuple(CAPTION_PUNCTUATION)) for text in header_texts):
                    p = next((mapping.get(id(x)) for x in first_row.xpath(".//w:p", namespaces=NS) if mapping.get(id(x))), None)
                    findings.append(self._finding("TAB-003", problem="表头文字末尾带标点",
                        actual="；".join(header_texts), location=self._location(p), evidence={"headers": header_texts},
                        suggestion="删除表头文字末尾标点。", anchor=self._anchor(p)))
            row_count = len(table.findall("w:tr", NS))
            if row_count >= 25 and (first_row is None or first_row.find("w:trPr/w:tblHeader", NS) is None):
                findings.append(self._finding("TAB-004", problem="长表格未设置重复标题行",
                    actual=f"表格有 {row_count} 行，首行未见 tblHeader", location=self._location(cap, scope="table", label=f"第 {index} 个表格"),
                    evidence={"row_count": row_count}, suggestion="设置跨页重复标题行，并在续页按要求标注续表。", anchor=self._anchor(cap)))
            note = after if after and re.match(r"^\s*(注|Note)\s*[：:]", after.text, re.I) else None
            if note:
                for run in self.package.iter_runs(note):
                    fmt = self.resolver.get_effective_run_format(run, note)
                    if run.text.strip() and fmt.get("size_pt") is not None and abs(fmt["size_pt"] - 10.5) > 0.2:
                        findings.append(self._finding("TAB-005", problem="表下注释字号不符合要求",
                            actual=f"{fmt['size_pt']:g} 磅", location=self._location(note), evidence={"format": fmt},
                            suggestion="将表下注释设置为宋体五号（10.5 磅）。", anchor=self._anchor(note)))
                        break
            findings.append(self._manual("TAB-006", problem="表格内容质量与数据准确性需要人工复核",
                actual=f"已定位第 {index} 个表格", suggestion="核对数据、单位、术语、自明性和专业准确性。", paragraph=cap))
        return findings

    def _check_equations(self, paragraphs: list[ParagraphRecord]) -> list[Finding]:
        findings = []
        start, end = self._main_matter_bounds(paragraphs)
        equation_paragraphs = [
            p for p in paragraphs
            if start <= p.index < end and p.element.xpath(".//m:oMath | .//m:oMathPara", namespaces=NS)
        ]
        for rule_id in ("EQ-001", "EQ-002"):
            self._mark(rule_id, len(equation_paragraphs))
        number_re = re.compile(r"\((\d+)\s*[-.]\s*(\d+)\)\s*$")
        loose_re = re.compile(r"(?<!\d)(\d+)\s*[-.]\s*(\d+)\s*$")
        for p in equation_paragraphs:
            pformat = self.resolver.get_effective_paragraph_format(p)
            if pformat.get("alignment") not in {"center", None} and not pformat.get("tabs"):
                findings.append(self._finding("EQ-001", problem="公式段落未见居中或制表位对齐结构",
                    actual=f"对齐={pformat.get('alignment')}，制表位={pformat.get('tabs')}", location=self._location(p),
                    evidence={"paragraph_format": pformat}, suggestion="使用可编辑公式，并设置公式居中、编号右对齐。", anchor=self._anchor(p)))
            text = p.text.strip()
            if not number_re.search(text):
                loose = loose_re.search(text)
                findings.append(self._finding("EQ-002", problem="公式编号未使用规范圆括号形式或未检测到编号",
                    actual=loose.group(0) if loose else text[-30:] or "无可见编号", location=self._location(p),
                    evidence={"visible_text": text}, suggestion="使用按章连续编号形式，例如“(2-3)”。", anchor=self._anchor(p)))
        return findings

    def _check_notes(self) -> list[Finding]:
        findings = []
        if not self.package.has_part("word/footnotes.xml"):
            for rule_id in ("NOTE-001", "NOTE-002", "NOTE-003", "NOTE-004"):
                self._mark(rule_id, 0)
            return findings
        paragraphs = [p for p in self.package.iter_paragraphs("word/footnotes.xml") if p.text.strip()]
        for rule_id in ("NOTE-001", "NOTE-003", "NOTE-004"):
            self._mark(rule_id, len(paragraphs))
        for p in paragraphs:
            pformat = self.resolver.get_effective_paragraph_format(p)
            issues = []
            sizes = []
            for run in self.package.iter_runs(p):
                if not run.text.strip():
                    continue
                fmt = self.resolver.get_effective_run_format(run, p)
                if fmt.get("size_pt") is not None:
                    sizes.append(fmt["size_pt"])
                east = fmt.get("fonts", {}).get("eastAsia")
                if re.search(r"[\u4e00-\u9fff]", run.text) and east and east not in {"宋体", "SimSun"}:
                    issues.append(f"中文字体 {east}")
            if any(abs(size - 9) > 0.2 for size in sizes):
                issues.append("字号 " + "/".join(f"{x:g}" for x in sorted(set(sizes))) + " 磅")
            multiple = pformat.get("line_spacing_multiple")
            if multiple is not None and abs(multiple - 1) > 0.05:
                issues.append(f"行距 {multiple:.2f} 倍")
            if any(abs(pformat.get(key, 0) or 0) > 0.1 for key in ("left_indent_pt", "first_line_indent_pt")):
                issues.append("存在左缩进或首行缩进")
            if issues:
                findings.append(self._finding("NOTE-001", problem="脚注文字或段落格式不符合要求",
                    actual="、".join(issues), location=self._location(p), evidence={"sizes_pt": sizes, "paragraph_format": pformat},
                    suggestion="设置中文宋体、英文 Times New Roman、9 磅、单倍行距并取消缩进。"))
            if re.search(r"https?://|www\.", p.text, re.I):
                findings.append(self._manual("NOTE-004", problem="网页来源脚注的引用语境需要复核", actual=p.text[:120],
                    suggestion="确认该网页来源位于对应引用页且信息完整。", paragraph=p))
            if re.search(r"\bpp?\.?\s*\d+[-–]\d+\b", p.text, re.I) or re.search(r"第\s*\d+[-–]\d+\s*页", p.text):
                findings.append(self._manual("NOTE-003", problem="脚注页码写法需要按语种复核", actual=p.text[:120],
                    suggestion="核对中英文页码范围和对应语种标点。", paragraph=p))
        root = self.package.xml("word/footnotes.xml")
        settings = self.package.xml("word/settings.xml") if self.package.has_part("word/settings.xml") else None
        restart_nodes = [] if settings is None else settings.xpath(".//w:footnotePr/w:numRestart", namespaces=NS)
        self._mark("NOTE-002", 1)
        if not restart_nodes or restart_nodes[0].get(qn("w", "val")) != "eachPage":
            findings.append(self._finding("NOTE-002", problem="未检测到脚注每页重新编号设置",
                actual="settings.xml 中无 w:numRestart='eachPage'", location=self._location(scope="footnotes", label="脚注设置"),
                evidence={"footnote_count": len(root.xpath(".//w:footnote", namespaces=NS))},
                suggestion="在 Word 脚注设置中选择每页重新编号，并渲染确认。", severity="WARNING"))
        return findings

    def _reference_region(self, paragraphs):
        start = next((i for i, p in enumerate(paragraphs) if re.fullmatch(r"\s*参考文献\s*", p.text)), None)
        if start is None:
            return [], None
        entries = []
        for p in paragraphs[start + 1:]:
            if re.match(r"^\s*(致谢|附录|攻读学位期间)", p.text):
                break
            if re.match(r"^\s*\[\d+\]", p.text):
                entries.append(p)
        return entries, paragraphs[start]

    def _check_references(self, paragraphs: list[ParagraphRecord]) -> list[Finding]:
        findings = []
        entries, heading = self._reference_region(paragraphs)
        numbers = [int(re.match(r"^\s*\[(\d+)\]", p.text).group(1)) for p in entries]
        for rule_id in ("REF-001", "REF-002", "REF-003", "REF-004", "REF-005", "REF-006", "REF-007", "REF-008", "REF-010", "REF-011"):
            self._mark(rule_id, len(entries))
        if entries and numbers != list(range(1, len(numbers) + 1)):
            findings.append(self._finding("REF-001", problem="参考文献编号存在跳号、缺失或顺序异常",
                actual=str(numbers), location=self._location(entries[0]), evidence={"numbers": numbers},
                suggestion="按正文首次出现顺序连续编号，并检查重复与缺失。", anchor=self._anchor(entries[0])))
        citation_numbers = []
        for p in paragraphs[: paragraphs.index(heading) if heading else len(paragraphs)]:
            citation_numbers.extend(int(x) for x in re.findall(r"\[(\d+)\]", p.text))
        missing_entries = sorted(set(citation_numbers) - set(numbers))
        if missing_entries:
            p = next((p for p in paragraphs if any(f"[{n}]" in p.text for n in missing_entries)), heading)
            findings.append(self._finding("REF-001", problem="正文引用编号在参考文献表中无对应条目",
                actual=str(missing_entries), location=self._location(p), evidence={"citations": citation_numbers, "entries": numbers},
                suggestion="补齐对应条目或修正正文引用编号。", anchor=self._anchor(p)))
        for p in entries:
            text = p.text.strip()
            content = re.sub(r"^\s*\[\d+\]\s*", "", text)
            if content.endswith(tuple(REF_FINAL_PUNCTUATION)):
                findings.append(self._finding("REF-003", problem="参考文献末尾存在格式追加标点候选",
                    actual=f"末尾字符“{content[-1]}”", location=self._location(p), evidence={"entry": content},
                    suggestion="删除条目末尾由格式追加的标点；若为字段固有句点，请人工确认后保留。", anchor=self._anchor(p)))
            author_segment = re.split(r"\.\s+", content, maxsplit=1)[0]
            authors = [x.strip() for x in re.split(r"[,，]", author_segment) if x.strip()]
            if len(authors) > 6 and not re.search(r"(?:等|et\s+al\.?)\s*$", author_segment, re.I):
                findings.append(self._finding("REF-002", problem="作者超过 6 人但未使用规定省略形式",
                    actual=author_segment, location=self._location(p), evidence={"parsed_author_count": len(authors)},
                    suggestion="列前 6 人后，中文加“等”，英文加“et al.”。", anchor=self._anchor(p)))
            if re.search(r"[A-Za-z]", author_segment) and re.search(r"(?:^|,\s*)[A-Z](?:\s+[A-Z])?\s+[A-Z][a-z]+", author_segment):
                findings.append(self._finding("REF-010", problem="英文作者名字首字母可能缺少句点",
                    actual=author_segment, location=self._location(p), evidence={"author_segment": author_segment},
                    suggestion="英文姓名使用带句点的名首字母在前、姓氏全称在后。", anchor=self._anchor(p)))
            lower = content.lower()
            if "[博士" in content or "[硕士" in content or "学位论文]" in content:
                rule_id = "REF-008"
                ok = bool(re.search(r":\s*\[(?:博士|硕士)学位论文\].*:\s*.*[,]，\s*\d{4}", content))
            elif re.search(r"\bZL\s*\d+|专利|patent", content, re.I):
                rule_id = "REF-007"
                ok = bool(re.search(r"(?:中国|CN|US|EP|WO).*(?:ZL|CN|US|EP|WO)?\s*\d+.*\d{4}", content, re.I))
            elif " in:" in lower or " 见" in content or "会议" in content or "conference" in lower:
                rule_id = "REF-006"
                ok = bool(re.search(r"(?:见|\bin)\s*[：:]", content, re.I) and re.search(r"\d{4}", content))
            elif re.search(r",\s*\d{4}\s*,\s*\d+", content) and re.search(r":\s*[A-Za-z]?\d+", content):
                rule_id = "REF-005"
                ok = bool(re.search(r",\s*\d{4}\s*,\s*\d+(?:\s*\(\d+\))?\s*:\s*\S+", content))
            else:
                rule_id = "REF-004"
                ok = bool(re.search(r"[.:：]\s*[^,，]+[,]，\s*\d{4}", content))
            if not ok:
                findings.append(self._finding(rule_id, problem="参考文献条目的字段顺序或必填字段需要复核",
                    actual=content[:180], location=self._location(p), evidence={"classification": rule_id},
                    suggestion="按该文献类型的 HUST 著录结构核对作者、题名、出处、地点和年份。", anchor=self._anchor(p)))
        self._mark("REF-009", len(entries))
        if heading and len(entries) < 40:
            findings.append(self._finding("REF-009", problem="参考文献数量低于建议值",
                actual=f"检测到 {len(entries)} 条", location=self._location(heading), evidence={"entry_count": len(entries)},
                suggestion="结合研究需要补充高质量文献；不少于 40 条属于建议性指标。", anchor=self._anchor(heading)))
        if heading:
            findings.append(self._manual("REF-011", problem="参考文献事实准确性不能仅凭文档内部结构确认",
                actual=f"已完成 {len(entries)} 条条目的内部编号和格式候选检查",
                suggestion="使用权威数据库逐条核对作者、题名、年份、卷期页码及重复项。", paragraph=heading))
        return findings

    def _check_objects(self, paragraphs: list[ParagraphRecord]) -> list[Finding]:
        findings = []
        root = self.package.xml(DocxPackage.MAIN_PART)
        shapes = root.xpath(".//wp:anchor | .//wp:inline | .//w:pict | .//v:shape | .//w:txbxContent", namespaces=NS)
        self._mark("OBJ-001", len(shapes))
        if shapes:
            findings.append(self._manual("OBJ-001", problem="形状、文本框或浮动对象的重叠与裁切需要渲染复核",
                actual=f"检测到 {len(shapes)} 个相关 OOXML 对象", suggestion="逐页检查锚定、环绕、颜色、文字和边界。",
                evidence={"object_count": len(shapes)}))
        bookmark_names = {node.get(qn("w", "name")) for node in root.xpath(".//w:bookmarkStart", namespaces=NS)}
        codes = root.xpath(".//w:instrText/text() | .//w:fldSimple/@w:instr", namespaces=NS)
        refs = []
        for code in codes:
            match = re.search(r"\b(?:REF|PAGEREF)\s+([^\s\\]+)", code, re.I)
            if match:
                refs.append(match.group(1))
        broken = sorted(set(refs) - bookmark_names)
        self._mark("OBJ-002", len(codes) + len(bookmark_names))
        if broken:
            findings.append(self._finding("OBJ-002", problem="交叉引用指向不存在的书签",
                actual="、".join(broken), location=self._location(scope="fields", label="域与书签"),
                evidence={"bookmarks": sorted(x for x in bookmark_names if x), "reference_targets": refs},
                suggestion="修复或重新插入交叉引用，然后更新全部域。"))
        revisions = len(root.xpath(".//w:ins | .//w:del | .//w:moveFrom | .//w:moveTo", namespaces=NS))
        controls = len(root.xpath(".//w:sdt", namespaces=NS))
        comment_ranges = len(root.xpath(".//w:commentRangeStart", namespaces=NS))
        self._mark("OBJ-003", revisions + controls + comment_ranges)
        if revisions or controls or comment_ranges:
            findings.append(self._finding("OBJ-003", problem="文档包含必须在审查副本中保留的复杂对象",
                actual=f"修订 {revisions}、内容控件 {controls}、原批注范围 {comment_ranges}",
                location=self._location(scope="objects", label="全文复杂对象清单"),
                evidence={"revisions": revisions, "content_controls": controls, "comment_ranges": comment_ranges},
                suggestion="生成副本后进行对象计数与关系比对，不自动接受修订或删除内容控件。", severity="WARNING"))
        return findings

    def _check_language_and_structure(self, paragraphs: list[ParagraphRecord]) -> list[Finding]:
        findings = [self._manual("LANG-001", problem="术语、符号和计量单位的一致性需要专业人工复核",
            actual="已执行缩写、空格和单位模式候选扫描", suggestion="由作者按学科现行标准统一术语、量、单位、符号和标点。")]
        start, end = self._main_matter_bounds(paragraphs)
        main_paragraphs = [p for p in paragraphs if start <= p.index < end and not p.in_table]
        all_text = "\n".join(p.text for p in main_paragraphs)
        acronyms: dict[str, list[ParagraphRecord]] = {}
        for p in main_paragraphs:
            for token in re.findall(r"\b[A-Z][A-Z0-9-]{2,}\b", p.text):
                acronyms.setdefault(token, []).append(p)
        self._mark("LANG-002", len(acronyms))
        for token, locations in list(acronyms.items())[:30]:
            first = locations[0]
            if not re.search(rf"(?:[A-Za-z][A-Za-z\s-]{{3,}}|[\u4e00-\u9fff]{{2,}})[（(]{re.escape(token)}[）)]", first.text):
                findings.append(self._finding("LANG-002", problem="缩写首次出现位置未检测到全称解释候选",
                    actual=f"{token}：{first.text[:100]}", location=self._location(first), evidence={"token": token, "occurrences": len(locations)},
                    suggestion="首次出现时给出中文或英文全称；公知缩写可人工豁免。", anchor=self._anchor(first)))
        unit_pattern = re.compile(r"(?<!\w)(\d+(?:\.\d+)?)(mm|cm|m|km|mg|kg|g|kPa|MPa|Pa|Hz|kHz|MHz|°C)(?!\w)")
        self._mark("LANG-003", len(main_paragraphs))
        for p in main_paragraphs:
            matches = [m.group(0) for m in unit_pattern.finditer(p.text)]
            if matches:
                findings.append(self._finding("LANG-003", problem="数字与单位之间缺少空格候选",
                    actual="、".join(matches[:8]), location=self._location(p), evidence={"matches": matches},
                    suggestion="结合量和单位规范确认是否应保留一个半角空格。", anchor=self._anchor(p)))
        classified = self.classifier.classify_document(paragraphs)
        chapters = [
            item for item in classified
            if start <= item.paragraph.index < end and item.classification.role == "heading_1"
        ]
        visible_count = len(re.sub(r"\s+", "", all_text))
        self._mark("STRUCT-001", len(chapters))
        if len(chapters) < 5 or visible_count < 25000:
            anchor = chapters[0].paragraph if chapters else None
            findings.append(self._finding("STRUCT-001", problem="章节数或可见字数低于建议指标",
                actual=f"检测到 {len(chapters)} 个编号一级标题，约 {visible_count} 个非空白字符；最终页数未知",
                location=self._location(anchor), evidence={"chapter_count": len(chapters), "visible_character_count": visible_count},
                suggestion="结合论文内容确认至少 5 章的建议结构和约 2.5 万字要求；页数需渲染确认。", anchor=self._anchor(anchor), severity="WARNING"))
        summaries = [p for p in paragraphs if re.search(r"本章小结", p.text)]
        self._mark("STRUCT-002", len(chapters))
        expected_summary_count = max(0, len(chapters) - 2)
        if len(summaries) < expected_summary_count:
            findings.append(self._finding("STRUCT-002", problem="研究章节的“本章小结”数量可能不足",
                actual=f"研究章节候选 {expected_summary_count} 个，检测到本章小结 {len(summaries)} 个",
                location=self._location(chapters[0].paragraph if chapters else None),
                evidence={"chapter_count": len(chapters), "summary_count": len(summaries)},
                suggestion="核对各研究章节末尾是否有简洁的本章小结；实际篇幅需人工复核。", severity="WARNING"))
        return findings

    def _check_appendices(self, paragraphs: list[ParagraphRecord]) -> list[Finding]:
        findings = []
        achievement = next((p for p in paragraphs if re.search(r"攻读.*学位.*(成果|论文|专利)", p.text)), None)
        appendix = next((p for p in paragraphs if re.match(r"^\s*附录", p.text)), None)
        for rule_id in ("APP-001", "APP-002"):
            self._mark(rule_id, 1 if achievement else 0)
        if achievement:
            findings.append(self._manual("APP-001", problem="成果附录作者身份、姓名加粗和第一单位需要人工核对",
                actual="已定位成果附录候选", suggestion="列出全部作者，核对本人加粗及第一作者单位说明。", paragraph=achievement))
            findings.append(self._manual("APP-002", problem="成果发表状态及页码/链接需要人工核对",
                actual="已定位成果附录候选", suggestion="逐条标明已发表、在线、接收、修改、投稿或拟投状态。", paragraph=achievement))
        self._mark("APP-003", 1 if appendix else 0)
        if appendix:
            findings.append(self._manual("APP-003", problem="附录内容是否属于适宜的补充材料需要人工判断",
                actual="已定位附录候选", suggestion="确认仅放置详细推导、实验数据、程序或原始资料等补充内容。", paragraph=appendix))
        return findings

    def _render_boundaries(self) -> list[Finding]:
        messages = {
            "RENDER-001": ("最终页面版式尚未执行可靠渲染检查", "OOXML 无法确认裁切、重叠、空白页和大面积异常留白", "使用 Microsoft Word 或兼容渲染器逐页检查。"),
            "RENDER-002": ("页眉页脚视觉效果尚未逐页渲染复核", "结构和颜色可检查，但碰撞、双线效果和各页一致性不能由 XML 保证", "逐页核对页眉文字、红色双线、页码位置及正文间距。"),
            "RENDER-003": ("图表公式最终尺寸下的可读性需要人工视觉复核", "对象结构已清点，最终缩放和页面构图未知", "以打印等效尺寸检查清晰度、可读性和页面均衡。"),
        }
        return [self._manual(rule_id, problem=problem, actual=actual, suggestion=suggestion)
                for rule_id, (problem, actual, suggestion) in messages.items()]


__all__ = ["ADVANCED_RULE_IDS", "AdvancedRuleEngine"]
