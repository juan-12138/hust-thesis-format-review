"""Phase 4 rule evaluation for HUST master-thesis DOCX files.

The engine is read-only. Rule values and severities come from the confirmed
rule configuration; this module supplies routing, normalization, comparison,
and traceable result construction.
"""

from __future__ import annotations

import json
import re
import hashlib
from dataclasses import asdict, dataclass
from pathlib import Path

from docx_model import DocxPackage, NS, ParagraphRecord
from effective_style import EffectiveStyleResolver
from structure_classifier import ClassifiedParagraph, StructureClassifier


PHASE4_RULE_IDS = (
    "PAGE-001",
    "PAGE-002",
    "PAGE-003",
    "PAGE-004",
    "HF-001",
    "HF-002",
    "HF-003",
    "HF-004",
    "FONT-001",
    "FONT-002",
    "FONT-003",
    "PARA-001",
    "PARA-002",
    "PARA-003",
    "HEAD-001",
    "HEAD-002",
    "HEAD-003",
    "HEAD-004",
    "HEAD-005",
)


def load_rule_config(path: str | Path) -> dict:
    """Load and validate the standard-library JSON runtime rule mirror."""
    path = Path(path)
    with path.open("r", encoding="utf-8") as handle:
        config = json.load(handle)
    mirror = config.get("runtime_mirror") or {}
    source_file = mirror.get("source_file")
    expected_source_hash = mirror.get("source_sha256")
    if source_file and expected_source_hash:
        source_path = path.with_name(source_file)
        if source_path.is_file():
            actual_source_hash = hashlib.sha256(source_path.read_bytes()).hexdigest().upper()
            if actual_source_hash != expected_source_hash.upper():
                raise ValueError("运行时 JSON 规则镜像与源 YAML 不一致，请重新生成镜像")
    rules = config.get("rules")
    if not isinstance(rules, list):
        raise ValueError("规则配置缺少 rules 列表")
    ids = [rule.get("rule_id") for rule in rules if isinstance(rule, dict)]
    if len(ids) != len(rules) or any(not rule_id for rule_id in ids):
        raise ValueError("每条规则都必须包含非空 rule_id")
    duplicates = sorted({rule_id for rule_id in ids if ids.count(rule_id) > 1})
    if duplicates:
        raise ValueError(f"规则 ID 重复：{', '.join(duplicates)}")
    missing = [rule_id for rule_id in PHASE4_RULE_IDS if rule_id not in ids]
    if missing:
        raise ValueError(f"阶段4规则缺失：{', '.join(missing)}")
    return config


@dataclass(frozen=True)
class Finding:
    finding_id: str
    rule_id: str
    category: str
    object_type: str
    severity: str
    problem: str
    actual: str
    expected: str
    source: str
    suggestion: str
    location: dict
    anchor: dict | None
    evidence: dict

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass(frozen=True)
class RuleResult:
    rule_id: str
    status: str
    finding_count: int
    evaluated_count: int

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass(frozen=True)
class ReviewResult:
    rule_results: tuple[RuleResult, ...]
    findings: tuple[Finding, ...]

    def to_dict(self) -> dict:
        return {
            "rule_results": [result.to_dict() for result in self.rule_results],
            "findings": [finding.to_dict() for finding in self.findings],
        }


class RuleEngine:
    """Evaluate the Phase 4 rule subset without modifying the package."""

    def __init__(self, package: DocxPackage, *, rules_path: str | Path):
        self.package = package
        config = load_rule_config(rules_path)
        self.rules = {rule["rule_id"]: rule for rule in config["rules"]}
        self.resolver = EffectiveStyleResolver(package)
        self.classifier = StructureClassifier(package, self.resolver)
        self._sequence = 0
        self._checked_counts: dict[str, int] = {}

    def review(self) -> ReviewResult:
        self._sequence = 0
        self._checked_counts = {}
        paragraphs = list(self.package.iter_paragraphs())
        classified = self.classifier.classify_document(paragraphs)
        contexts = self._main_matter_contexts(classified)
        main_section_indices = sorted({
            item.paragraph.section_index
            for item in classified
            if contexts.get(item.paragraph.index) and item.paragraph.section_index is not None
        })
        findings: list[Finding] = []
        findings.extend(self._check_pages(main_section_indices))
        findings.extend(self._check_header_footer(main_section_indices))
        findings.extend(self._check_body_rules(classified, contexts))
        findings.extend(self._check_spacing_rules(paragraphs, contexts))
        findings.extend(self._check_heading_rules(classified, contexts))
        results = self._summarize(findings, PHASE4_RULE_IDS)
        return ReviewResult(tuple(results), tuple(findings))

    def _summarize(self, findings: list[Finding], rule_ids) -> list[RuleResult]:
        priority = {"PASS": 0, "MANUAL_REVIEW": 1, "WARNING": 2, "ERROR": 3}
        results = []
        for rule_id in rule_ids:
            matched = [finding for finding in findings if finding.rule_id == rule_id]
            status = max((finding.severity for finding in matched), key=priority.get) if matched else "PASS"
            results.append(RuleResult(rule_id, status, len(matched), self._checked_counts.get(rule_id, 0)))
        return results

    def _mark_checked(self, rule_id: str, count: int = 1):
        self._checked_counts[rule_id] = self._checked_counts.get(rule_id, 0) + count

    def _finding(self, rule_id: str, *, severity: str | None = None, problem: str, actual: str, location: dict, evidence: dict, suggestion: str, anchor: dict | None = None) -> Finding:
        rule = self.rules[rule_id]
        self._sequence += 1
        return Finding(
            finding_id=f"{rule_id}-{self._sequence:04d}",
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

    def _main_matter_contexts(self, classified: list[ClassifiedParagraph]) -> dict[int, bool]:
        active = False
        contexts: dict[int, bool] = {}
        end_roles = {"references_title", "acknowledgements_title", "appendix_heading"}
        for item in classified:
            role = item.classification.role
            if role == "heading_1":
                paragraph_format = self.resolver.get_effective_paragraph_format(item.paragraph)
                if re.match(r"^\s*\d+", item.paragraph.text) or paragraph_format.get("numbering_id") is not None:
                    active = True
            if role in end_roles:
                active = False
            contexts[item.paragraph.index] = active
        return contexts

    @staticmethod
    def _location(paragraph: ParagraphRecord) -> dict:
        snippet = " ".join(paragraph.text.split())[:40]
        return {
            "scope": "paragraph",
            "part_name": paragraph.part_name,
            "section_index": paragraph.section_index,
            "paragraph_index": paragraph.index,
            "label": paragraph.location,
            "snippet": snippet,
        }

    @staticmethod
    def _anchor(paragraph: ParagraphRecord, run_indices: list[int] | None = None) -> dict:
        anchor = {
            "part_name": paragraph.part_name,
            "paragraph_index": paragraph.index,
            "snippet": " ".join(paragraph.text.split())[:40],
        }
        if run_indices:
            anchor["run_start"] = min(run_indices)
            anchor["run_end"] = max(run_indices)
        return anchor

    @staticmethod
    def _font_key(value: str | None) -> str | None:
        if value is None:
            return None
        compact = re.sub(r"[\s_-]+", "", value).casefold()
        aliases = {
            "simsun": "宋体",
            "宋体": "宋体",
            "simhei": "黑体",
            "黑体": "黑体",
            "timesnewroman": "timesnewroman",
        }
        return aliases.get(compact, compact)

    @staticmethod
    def _has_cjk(text: str) -> bool:
        return bool(re.search(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]", text))

    @staticmethod
    def _has_latin_or_digit(text: str) -> bool:
        return bool(re.search(r"[A-Za-z0-9]", text))

    @staticmethod
    def _run_excluded_from_color(run) -> bool:
        return bool(
            run.paragraph.in_revision
            or run.element.xpath("ancestor::w:hyperlink | ancestor::w:fldSimple | ancestor::w:ins | ancestor::w:del", namespaces=NS)
            or run.element.find("w:rPr/w:vanish", NS) is not None
            or run.element.find("w:fldChar", NS) is not None
            or run.element.find("w:instrText", NS) is not None
        )

    def _check_body_rules(self, classified: list[ClassifiedParagraph], contexts: dict[int, bool]) -> list[Finding]:
        findings: list[Finding] = []
        font_rule = self.rules["FONT-001"]
        expected = font_rule["expected_value"]
        expected_east = self._font_key(expected["east_asian_font"])
        expected_latin = self._font_key(expected["latin_font"])
        size_tolerance = font_rule["tolerance"]["size_pt"]

        for item in classified:
            paragraph = item.paragraph
            if not contexts.get(paragraph.index) or item.classification.role != "body" or paragraph.in_table or not paragraph.text.strip():
                continue
            runs = [run for run in self.package.iter_runs(paragraph) if run.text.strip()]
            for rule_id in ("FONT-001", "FONT-002", "PARA-001"):
                self._mark_checked(rule_id)
            mismatches = []
            mismatch_runs: list[int] = []
            mixed_mismatches = []
            color_mismatches = []
            color_runs: list[int] = []
            for run in runs:
                fmt = self.resolver.get_effective_run_format(run, paragraph)
                fonts = fmt.get("fonts", {})
                east = fonts.get("eastAsia")
                latin = fonts.get("hAnsi") or fonts.get("ascii")
                cjk = self._has_cjk(run.text)
                latin_text = self._has_latin_or_digit(run.text)
                current = {
                    "run_index": run.index,
                    "text": run.text[:30],
                    "east_asian_font": east,
                    "latin_font": latin,
                    "size_pt": fmt.get("size_pt"),
                }
                wrong = []
                unresolved = []
                if cjk:
                    if east is None:
                        unresolved.append("east_asian_font")
                    elif self._font_key(east) != expected_east:
                        wrong.append("east_asian_font")
                if latin_text:
                    if latin is None:
                        unresolved.append("latin_font")
                    elif self._font_key(latin) != expected_latin:
                        wrong.append("latin_font")
                size = fmt.get("size_pt")
                if size is None:
                    unresolved.append("size_pt")
                elif abs(size - expected["size_pt"]) > size_tolerance:
                    wrong.append("size_pt")
                if wrong or unresolved:
                    current["mismatched_properties"] = wrong
                    current["unresolved_properties"] = unresolved
                    mismatches.append(current)
                    mismatch_runs.append(run.index)
                    if cjk and latin_text and any(name in wrong + unresolved for name in ("east_asian_font", "latin_font")):
                        mixed_mismatches.append(current)
                if cjk and latin_text:
                    self._mark_checked("FONT-003")

                if not self._run_excluded_from_color(run):
                    color = fmt.get("color_hex")
                    if color not in {None, "auto", "000000"}:
                        color_mismatches.append({"run_index": run.index, "text": run.text[:30], "color_hex": color})
                        color_runs.append(run.index)

            if mismatches:
                findings.append(self._finding(
                    "FONT-001",
                    severity="MANUAL_REVIEW" if not any(entry["mismatched_properties"] for entry in mismatches) else None,
                    problem="正文有效字体或字号不符合规则",
                    actual=json.dumps(mismatches, ensure_ascii=False, sort_keys=True),
                    location=self._location(paragraph),
                    anchor=self._anchor(paragraph, mismatch_runs),
                    evidence={"classification": asdict(item.classification), "effective_formats_compared": len(runs)},
                    suggestion="将中文设为宋体、英文和数字设为 Times New Roman，并统一为 12 磅。",
                ))
            if color_mismatches:
                findings.append(self._finding(
                    "FONT-002",
                    problem="普通正文存在非黑色文字",
                    actual=json.dumps(color_mismatches, ensure_ascii=False, sort_keys=True),
                    location=self._location(paragraph),
                    anchor=self._anchor(paragraph, color_runs),
                    evidence={"excluded_contexts": ["hyperlink", "field", "revision", "hidden text"]},
                    suggestion="确认不是合法强调后，将正文文字颜色改为自动黑色或 #000000。",
                ))
            if mixed_mismatches:
                findings.append(self._finding(
                    "FONT-003",
                    severity="MANUAL_REVIEW" if not any(entry["mismatched_properties"] for entry in mixed_mismatches) else None,
                    problem="同一文字片段中的中文与英文/数字未分别使用规定字体",
                    actual=json.dumps(mixed_mismatches, ensure_ascii=False, sort_keys=True),
                    location=self._location(paragraph),
                    anchor=self._anchor(paragraph, [item["run_index"] for item in mixed_mismatches]),
                    evidence={"script_detection": "CJK plus Latin/digit in the same run"},
                    suggestion="分别设置 w:eastAsia 与 w:ascii/w:hAnsi 字体。",
                ))

            paragraph_format = self.resolver.get_effective_paragraph_format(paragraph)
            alignment = paragraph_format.get("alignment")
            multiple = paragraph_format.get("line_spacing_multiple")
            line_tolerance = self.rules["PARA-001"]["tolerance"]["line_spacing_twips"] / 240
            wrong_paragraph = {}
            unresolved_paragraph = []
            if alignment is None:
                unresolved_paragraph.append("alignment")
            elif alignment not in {"both", "justified"}:
                wrong_paragraph["alignment"] = alignment
            if multiple is None:
                unresolved_paragraph.append("line_spacing")
            elif abs(multiple - 1.5) > line_tolerance:
                wrong_paragraph["line_spacing_multiple"] = multiple
                if "line_spacing_pt" in paragraph_format:
                    wrong_paragraph["line_spacing_pt"] = paragraph_format["line_spacing_pt"]
            if wrong_paragraph or unresolved_paragraph:
                paragraph_actual = dict(wrong_paragraph)
                if unresolved_paragraph:
                    paragraph_actual["unresolved_properties"] = unresolved_paragraph
                findings.append(self._finding(
                    "PARA-001",
                    severity="MANUAL_REVIEW" if not wrong_paragraph else None,
                    problem="正文对齐方式或行距不符合规则",
                    actual=json.dumps(paragraph_actual, ensure_ascii=False, sort_keys=True),
                    location=self._location(paragraph),
                    anchor=self._anchor(paragraph),
                    evidence={"provenance": paragraph_format.get("provenance", {})},
                    suggestion="将正文设置为两端对齐和 1.5 倍行距。",
                ))
        return findings

    def _check_spacing_rules(self, paragraphs: list[ParagraphRecord], contexts: dict[int, bool]) -> list[Finding]:
        findings: list[Finding] = []
        empty_run: list[ParagraphRecord] = []
        self._mark_checked("PARA-002", len(paragraphs))
        self._mark_checked("PARA-003", len(paragraphs))

        def flush_empty_run():
            if len(empty_run) >= 2:
                first = empty_run[0]
                findings.append(self._finding(
                    "PARA-002",
                    problem="正文区域存在连续空段，可能以空段代替段落间距",
                    actual=f"连续空段 {len(empty_run)} 个（第{first.index + 1}至第{empty_run[-1].index + 1}段）",
                    location=self._location(first),
                    anchor=self._anchor(first),
                    evidence={"paragraph_indices": [paragraph.index for paragraph in empty_run]},
                    suggestion="删除非必要空段，使用段前、段后或分页属性控制版面。",
                ))
            empty_run.clear()

        for paragraph in paragraphs:
            in_main = contexts.get(paragraph.index, False)
            if in_main and not paragraph.in_table and not paragraph.text.strip():
                empty_run.append(paragraph)
            else:
                flush_empty_run()

            if in_main and paragraph.text.strip() and (re.search(r" {2,}", paragraph.text) or "\u3000" in paragraph.text):
                findings.append(self._finding(
                    "PARA-002",
                    problem="正文段落包含连续空格或全角空格",
                    actual=repr(paragraph.text[:80]),
                    location=self._location(paragraph),
                    anchor=self._anchor(paragraph),
                    evidence={"consecutive_spaces": bool(re.search(r" {2,}", paragraph.text)), "fullwidth_space": "\u3000" in paragraph.text},
                    suggestion="删除用于排版的人工空格，改用缩进、制表位或段落属性。",
                ))

            break_types = paragraph.element.xpath('.//w:br[@w:type="page" or @w:type="column"]/@w:type', namespaces=NS)
            if break_types:
                findings.append(self._finding(
                    "PARA-003",
                    severity="MANUAL_REVIEW",
                    problem="检测到人工分页或分栏，但 OOXML 无法判断其最终分页影响",
                    actual=json.dumps(break_types, ensure_ascii=False),
                    location=self._location(paragraph),
                    anchor=self._anchor(paragraph),
                    evidence={"break_types": break_types, "render_required": True},
                    suggestion="在 Word 最终分页视图中检查是否造成空白页、孤行或题注分离。",
                ))
        flush_empty_run()
        return findings

    @staticmethod
    def _heading_number(text: str) -> tuple[int, ...] | None:
        match = re.match(r"^\s*(\d+(?:[.．]\d+)*)", text)
        if not match:
            return None
        try:
            return tuple(int(part) for part in re.split(r"[.．]", match.group(1)))
        except ValueError:
            return None

    def _check_heading_rules(self, classified: list[ClassifiedParagraph], contexts: dict[int, bool]) -> list[Finding]:
        findings: list[Finding] = []
        headings: list[tuple[int, ClassifiedParagraph, dict, tuple[int, ...] | None, str]] = []
        rule_by_level = {1: "HEAD-001", 2: "HEAD-002", 3: "HEAD-003"}

        for item in classified:
            role = item.classification.role
            if role not in {"heading_1", "heading_2", "heading_3"} or not contexts.get(item.paragraph.index):
                continue
            level = int(role.rsplit("_", 1)[1])
            rule_id = rule_by_level[level]
            self._mark_checked(rule_id)
            paragraph = item.paragraph
            paragraph_format = self.resolver.get_effective_paragraph_format(paragraph)
            number = self._heading_number(paragraph.text)
            number_mode = "automatic" if paragraph_format.get("numbering_id") is not None else ("manual" if number else "missing")
            headings.append((level, item, paragraph_format, number, number_mode))

            rule = self.rules[rule_id]
            expected = rule["expected_value"]
            size_tolerance = rule["tolerance"]["size_pt"]
            mismatches = []
            mismatch_runs: list[int] = []
            runs = [run for run in self.package.iter_runs(paragraph) if run.text.strip()]
            for run in runs:
                fmt = self.resolver.get_effective_run_format(run, paragraph)
                current = {
                    "run_index": run.index,
                    "text": run.text[:30],
                    "east_asian_font": fmt.get("fonts", {}).get("eastAsia"),
                    "size_pt": fmt.get("size_pt"),
                    "bold": fmt.get("bold"),
                }
                wrong = []
                if self._has_cjk(run.text) and self._font_key(current["east_asian_font"]) != self._font_key(expected["east_asian_font"]):
                    wrong.append("east_asian_font")
                if current["size_pt"] is None or abs(current["size_pt"] - expected["size_pt"]) > size_tolerance:
                    wrong.append("size_pt")
                if current["bold"] is not True:
                    wrong.append("bold")
                if wrong:
                    current["mismatched_properties"] = wrong
                    mismatches.append(current)
                    mismatch_runs.append(run.index)

            paragraph_mismatches = {}
            if level == 1 and paragraph_format.get("alignment") != "center":
                paragraph_mismatches["alignment"] = paragraph_format.get("alignment")
            if number_mode == "missing" or (number is not None and len(number) != level):
                paragraph_mismatches["numbering"] = number or "missing"
            if mismatches or paragraph_mismatches:
                actual = {"runs": mismatches, "paragraph": paragraph_mismatches}
                findings.append(self._finding(
                    rule_id,
                    problem=f"{level}级标题的字体、字号、粗体、对齐或编号格式不符合规则",
                    actual=json.dumps(actual, ensure_ascii=False, sort_keys=True),
                    location=self._location(paragraph),
                    anchor=self._anchor(paragraph, mismatch_runs),
                    evidence={"classification": asdict(item.classification), "paragraph_format_provenance": paragraph_format.get("provenance", {})},
                    suggestion=f"按规则统一设置{level}级标题格式并使用对应层级编号。",
                ))

        self._mark_checked("HEAD-004", len(headings))
        self._mark_checked("HEAD-005", sum(level in {2, 3} for level, *_ in headings))
        previous_level = None
        last_number_by_level: dict[int, tuple[int, ...]] = {}
        modes = {mode for _, _, _, _, mode in headings if mode in {"automatic", "manual"}}
        mixed_mode_reported = False
        generic_titles = {"研究内容", "相关工作", "概述", "其他", "小结", "本章小结"}

        for level, item, paragraph_format, number, number_mode in headings:
            paragraph = item.paragraph
            if previous_level is not None and level > previous_level + 1:
                findings.append(self._finding(
                    "HEAD-004",
                    problem=f"标题层级跳级：从{previous_level}级直接进入{level}级",
                    actual=paragraph.text,
                    location=self._location(paragraph),
                    anchor=self._anchor(paragraph),
                    evidence={"previous_level": previous_level, "current_level": level},
                    suggestion="补齐中间标题层级或调整当前标题样式和大纲级别。",
                ))
            previous_level = level

            if number and len(number) == level:
                previous_number = last_number_by_level.get(level)
                discontinuous = False
                if previous_number:
                    same_parent = number[:-1] == previous_number[:-1]
                    if same_parent and number[-1] != previous_number[-1] + 1:
                        discontinuous = True
                    elif not same_parent and number[-1] != 1:
                        discontinuous = True
                elif number[-1] != 1:
                    discontinuous = True
                if discontinuous:
                    findings.append(self._finding(
                        "HEAD-004",
                        problem=f"{level}级标题编号不连续",
                        actual=".".join(map(str, number)),
                        location=self._location(paragraph),
                        anchor=self._anchor(paragraph),
                        evidence={"previous_number": previous_number, "current_number": number},
                        suggestion="核对标题编号字段、列表定义及前一同级标题，恢复连续编号。",
                    ))
                last_number_by_level[level] = number
                for deeper in range(level + 1, 10):
                    last_number_by_level.pop(deeper, None)

            if len(modes) > 1 and not mixed_mode_reported:
                findings.append(self._finding(
                    "HEAD-004",
                    problem="标题体系混用了自动编号和手工编号",
                    actual=json.dumps(sorted(modes), ensure_ascii=False),
                    location=self._location(paragraph),
                    anchor=self._anchor(paragraph),
                    evidence={"numbering_modes": sorted(modes)},
                    suggestion="统一使用同一套多级列表与标题样式。",
                ))
                mixed_mode_reported = True

            if level in {2, 3}:
                title = re.sub(r"^\s*\d+(?:[.．]\d+)*\s*", "", paragraph.text).strip(" ：:、。.．")
                if title in generic_titles:
                    findings.append(self._finding(
                        "HEAD-005",
                        problem="标题表述较笼统，需要确认是否可改为能反映具体研究内容的标题",
                        actual=title,
                        location=self._location(paragraph),
                        anchor=self._anchor(paragraph),
                        evidence={"heuristic": "generic-title candidate", "semantic_confirmation_required": True},
                        suggestion="结合本节实际内容人工判断，并在合适时改为更具体的标题。",
                    ))
        return findings

    def _check_pages(self, main_section_indices: list[int]) -> list[Finding]:
        findings: list[Finding] = []
        sections = self.package.sections()
        for rule_id in ("PAGE-001", "PAGE-002", "PAGE-003"):
            self._mark_checked(rule_id, len(sections))
        self._mark_checked("PAGE-004")
        for section in sections:
            index = section["section_index"]
            location = {"scope": "section", "part_name": DocxPackage.MAIN_PART, "section_index": index, "label": f"第{index + 1}节"}

            page_rule = self.rules["PAGE-001"]
            page_expected = page_rule["expected_value"]
            page_tolerance = page_rule["tolerance"]["twips"]
            expected_width = page_expected["width_inches"] * 1440
            expected_height = page_expected["height_inches"] * 1440
            page_values = {
                "page_width_twips": (section["page_width_twips"], expected_width),
                "page_height_twips": (section["page_height_twips"], expected_height),
            }
            missing_page = [name for name, (actual, _) in page_values.items() if actual is None]
            bad_page = [name for name, (actual, expected) in page_values.items() if actual is not None and abs(actual - expected) > page_tolerance]
            if section["orientation"] is None:
                missing_page.append("orientation")
            elif section["orientation"] != page_expected["orientation"]:
                bad_page.append("orientation")
            page_bad = bool(missing_page or bad_page)
            if page_bad:
                actual = {
                    "page_width_twips": section["page_width_twips"],
                    "page_height_twips": section["page_height_twips"],
                    "orientation": section["orientation"],
                }
                findings.append(self._finding(
                    "PAGE-001",
                    severity="MANUAL_REVIEW" if not bad_page else None,
                    problem="纸张尺寸或页面方向不符合规则" if bad_page else "纸张尺寸或页面方向属性缺失，无法确定最终页面设置",
                    actual=json.dumps(actual, ensure_ascii=False, sort_keys=True),
                    location=location,
                    evidence={"section": section, "missing_properties": missing_page, "mismatched_properties": bad_page},
                    suggestion="将本节纸张设置为 A4 纵向。",
                ))

            margin_rule = self.rules["PAGE-002"]
            margin_tolerance = margin_rule["tolerance"]["twips"]
            margin_expected = margin_rule["expected_value"]
            margin_pairs = {
                "top": (section["margin_top_twips"], margin_expected["top_inches"] * 1440),
                "bottom": (section["margin_bottom_twips"], margin_expected["bottom_inches"] * 1440),
                "left": (section["margin_left_twips"], margin_expected["left_inches"] * 1440),
                "right": (section["margin_right_twips"], margin_expected["right_inches"] * 1440),
            }
            missing_margins = [name for name, (actual, _) in margin_pairs.items() if actual is None]
            bad_margins = {name: actual for name, (actual, expected) in margin_pairs.items() if actual is not None and abs(actual - expected) > margin_tolerance}
            if bad_margins or missing_margins:
                findings.append(self._finding(
                    "PAGE-002",
                    severity="MANUAL_REVIEW" if not bad_margins else None,
                    problem="页边距不符合规则" if bad_margins else "页边距属性缺失，无法确定最终页边距",
                    actual=json.dumps({"twips": bad_margins, "missing": missing_margins}, ensure_ascii=False, sort_keys=True),
                    location=location,
                    evidence={"compared_values": margin_pairs, "tolerance_twips": margin_tolerance},
                    suggestion="按规则值调整本节上下左右页边距。",
                ))

            distance_rule = self.rules["PAGE-003"]
            distance_tolerance = distance_rule["tolerance"]["twips"]
            distance_expected = distance_rule["expected_value"]
            distance_pairs = {
                "header": (section["header_distance_twips"], distance_expected["header_distance_pt"] * 20),
                "footer": (section["footer_distance_twips"], distance_expected["footer_distance_pt"] * 20),
            }
            missing_distances = [name for name, (actual, _) in distance_pairs.items() if actual is None]
            bad_distances = {name: actual for name, (actual, expected) in distance_pairs.items() if actual is not None and abs(actual - expected) > distance_tolerance}
            if bad_distances or missing_distances:
                findings.append(self._finding(
                    "PAGE-003",
                    severity="MANUAL_REVIEW" if not bad_distances else None,
                    problem="页眉或页脚距边界与模板不一致" if bad_distances else "页眉或页脚距边界属性缺失，无法确定最终距离",
                    actual=json.dumps({"twips": bad_distances, "missing": missing_distances}, ensure_ascii=False, sort_keys=True),
                    location=location,
                    evidence={"compared_values": distance_pairs, "tolerance_twips": distance_tolerance},
                    suggestion="核对并调整本节页眉、页脚距边界。",
                ))

        section_summary = [
            {
                "section_index": section["section_index"],
                "section_start": section["section_start"],
                "page_number_format": section["page_number_format"],
                "page_number_start": section["page_number_start"],
            }
            for section in sections
        ]
        findings.append(self._finding(
            "PAGE-004",
            severity="MANUAL_REVIEW",
            problem="OOXML 无法确定分节符是否造成非预期空白页或实际分页异常",
            actual=json.dumps(section_summary, ensure_ascii=False, sort_keys=True),
            location={"scope": "document", "part_name": DocxPackage.MAIN_PART, "section_index": None, "label": "整篇文档的分节结构"},
            evidence={"render_required": True, "section_count": len(sections)},
            suggestion="在 Word 最终分页视图中人工复核分节、空白页和页码切换。",
        ))
        if main_section_indices:
            first_main = sections[main_section_indices[0]]
            number_format = first_main.get("page_number_format")
            number_start = first_main.get("page_number_start")
            if number_format not in {None, "decimal"} or number_start not in {None, 1}:
                findings.append(self._finding(
                    "PAGE-004",
                    severity="WARNING",
                    problem="正文起始节的页码体系疑似不正确",
                    actual=json.dumps({"page_number_format": number_format, "page_number_start": number_start}, ensure_ascii=False, sort_keys=True),
                    location={"scope": "section", "part_name": DocxPackage.MAIN_PART, "section_index": first_main["section_index"], "label": f"正文起始节（第{first_main['section_index'] + 1}节）"},
                    evidence={"main_section_indices": main_section_indices},
                    suggestion="正文通常使用从 1 开始的阿拉伯数字页码；核对前置部分与正文的分节和页码格式。",
                ))
        return findings

    def _header_footer_location(self, part_name: str, paragraph: ParagraphRecord | None = None, section_index: int | None = None) -> dict:
        if paragraph is not None:
            location = self._location(paragraph)
            location["scope"] = "header_footer_paragraph"
            return location
        return {
            "scope": "section",
            "part_name": part_name,
            "section_index": section_index,
            "label": f"第{section_index + 1}节" if section_index is not None else part_name,
        }

    def _check_header_footer(self, main_section_indices: list[int]) -> list[Finding]:
        findings: list[Finding] = []
        sections = self.package.sections()
        main_sections = [sections[index] for index in main_section_indices if 0 <= index < len(sections)]
        settings = self.package.settings()

        def active_parts(reference_key: str) -> list[str]:
            parts = set()
            for section in main_sections:
                references = section[reference_key]
                if references.get("default"):
                    parts.add(references["default"])
                if settings.get("even_and_odd_headers") and references.get("even"):
                    parts.add(references["even"])
                if section.get("different_first_page") and references.get("first"):
                    parts.add(references["first"])
            return sorted(parts)

        header_parts = active_parts("effective_header_references")
        footer_parts = active_parts("effective_footer_references")
        self._mark_checked("HF-001", max(1, len(header_parts)) if main_sections else 0)
        self._mark_checked("HF-002", max(1, len(header_parts)) if main_sections else 0)
        self._mark_checked("HF-003", max(1, len(footer_parts)) if main_sections else 0)
        self._mark_checked("HF-004", len(main_sections))

        if main_sections and not header_parts:
            first = main_sections[0]
            location = self._header_footer_location(DocxPackage.MAIN_PART, section_index=first["section_index"])
            findings.append(self._finding(
                "HF-001",
                problem="正文节未找到有效页眉引用",
                actual="无有效页眉部件",
                location=location,
                evidence={"section": first},
                suggestion="为正文节设置规定的论文页眉，并核对与前一节的链接状态。",
            ))
            findings.append(self._finding(
                "HF-002",
                problem="缺少页眉，无法确认红色文字与双线外观",
                actual="无有效页眉部件",
                location=location,
                evidence={"section": first},
                suggestion="按模板设置页眉外观，并在 Word 中复核双线位置。",
            ))

        expected_header_text = self.rules["HF-001"]["expected_value"]["text"]
        for part_name in header_parts:
            paragraphs = list(self.package.iter_paragraphs(part_name)) if self.package.has_part(part_name) else []
            visible = [paragraph for paragraph in paragraphs if paragraph.text.strip()]
            expected_compact = "".join(expected_header_text.split())
            paragraph = next((candidate for candidate in visible if "".join(candidate.text.split()) == expected_compact), visible[0] if visible else None)
            text_value = "".join(paragraph.text.split()) if paragraph else ""
            format_mismatches = []
            run_indices = []
            if paragraph:
                for run in self.package.iter_runs(paragraph):
                    if not run.text.strip():
                        continue
                    fmt = self.resolver.get_effective_run_format(run, paragraph)
                    wrong = []
                    if self._has_cjk(run.text) and self._font_key(fmt.get("fonts", {}).get("eastAsia")) != self._font_key(self.rules["HF-001"]["expected_value"]["east_asian_font"]):
                        wrong.append("east_asian_font")
                    if fmt.get("size_pt") is None or abs(fmt["size_pt"] - self.rules["HF-001"]["expected_value"]["size_pt"]) > 0.25:
                        wrong.append("size_pt")
                    if wrong:
                        format_mismatches.append({"run_index": run.index, "text": run.text[:30], "font": fmt.get("fonts", {}).get("eastAsia"), "size_pt": fmt.get("size_pt"), "mismatched_properties": wrong})
                        run_indices.append(run.index)
            if text_value != expected_compact or format_mismatches:
                findings.append(self._finding(
                    "HF-001",
                    problem="正文页眉文字、字体或字号不符合规则",
                    actual=json.dumps({"text": text_value, "format_mismatches": format_mismatches}, ensure_ascii=False, sort_keys=True),
                    location=self._header_footer_location(part_name, paragraph),
                    anchor=self._anchor(paragraph, run_indices) if paragraph else None,
                    evidence={"header_part": part_name},
                    suggestion="将页眉改为规定文字，并按规则设置楷体 18 磅。",
                ))

            root = self.package.xml(part_name) if self.package.has_part(part_name) else None
            colors = []
            if paragraph:
                for run in self.package.iter_runs(paragraph):
                    if run.text.strip():
                        colors.append(self.resolver.get_effective_run_format(run, paragraph).get("color_hex"))
            red_shapes = 0
            if root is not None:
                for shape in root.xpath(".//v:line | .//v:shape", namespaces=NS):
                    stroke = (shape.get("strokecolor") or "").strip().casefold().lstrip("#")
                    if stroke in {"ff0000", "red"}:
                        red_shapes += 1
            if not colors or any(color != "FF0000" for color in colors) or red_shapes < self.rules["HF-002"]["expected_value"]["rule_count"]:
                findings.append(self._finding(
                    "HF-002",
                    problem="页眉红色文字或双线外观与模板不一致",
                    actual=json.dumps({"text_colors": colors, "red_rule_count": red_shapes}, ensure_ascii=False, sort_keys=True),
                    location=self._header_footer_location(part_name, paragraph),
                    anchor=self._anchor(paragraph) if paragraph else None,
                    evidence={"header_part": part_name, "visual_confirmation_required": True},
                    suggestion="核对页眉文字颜色和两条红线，并在 Word 渲染视图中确认位置与粗细。",
                ))

        if main_sections and not footer_parts:
            first = main_sections[0]
            findings.append(self._finding(
                "HF-003",
                problem="正文节未找到有效页脚页码",
                actual="无有效页脚部件",
                location=self._header_footer_location(DocxPackage.MAIN_PART, section_index=first["section_index"]),
                evidence={"section": first},
                suggestion="在正文页脚中插入 PAGE 域并设置视觉居中。",
            ))

        for part_name in footer_parts:
            paragraphs = list(self.package.iter_paragraphs(part_name)) if self.package.has_part(part_name) else []
            field_paragraph = None
            centered = False
            page_field = False
            for paragraph in paragraphs:
                instructions = paragraph.element.xpath(".//w:fldSimple/@w:instr | .//w:instrText/text()", namespaces=NS)
                if any(re.search(r"\bPAGE\b", instruction, re.I) for instruction in instructions):
                    page_field = True
                    field_paragraph = paragraph
                    fmt = self.resolver.get_effective_paragraph_format(paragraph)
                    centered = fmt.get("alignment") == "center" or any(tab.get("alignment") == "center" for tab in fmt.get("tabs", []))
                    if centered:
                        break
            if not page_field or not centered:
                findings.append(self._finding(
                    "HF-003",
                    problem="页脚缺少 PAGE 域或页码未能确认视觉居中",
                    actual=json.dumps({"page_field": page_field, "centered_by_alignment_or_tab": centered}, ensure_ascii=False, sort_keys=True),
                    location=self._header_footer_location(part_name, field_paragraph),
                    anchor=self._anchor(field_paragraph) if field_paragraph else None,
                    evidence={"footer_part": part_name, "tab_aware_alignment": True},
                    suggestion="使用 PAGE 域生成页码，并通过居中对齐或居中制表位实现视觉居中。",
                ))

        if main_sections:
            first_main = main_sections[0]
            broken_inheritance = (
                first_main.get("header_linked_to_previous") and not first_main.get("effective_header_references")
            ) or (
                first_main.get("footer_linked_to_previous") and not first_main.get("effective_footer_references")
            )
            if broken_inheritance:
                findings.append(self._finding(
                    "HF-004",
                    problem="正文起始节链接到前一节，但没有可解析的有效页眉或页脚",
                    actual=json.dumps({"header_linked_to_previous": first_main.get("header_linked_to_previous"), "footer_linked_to_previous": first_main.get("footer_linked_to_previous"), "effective_headers": first_main.get("effective_header_references"), "effective_footers": first_main.get("effective_footer_references")}, ensure_ascii=False, sort_keys=True),
                    location=self._header_footer_location(DocxPackage.MAIN_PART, section_index=first_main["section_index"]),
                    evidence={"section": first_main},
                    suggestion="核对前一节的页眉页脚部件和“链接到前一节”状态，修复失效继承。",
                ))
        return findings


__all__ = ["Finding", "PHASE4_RULE_IDS", "ReviewResult", "RuleEngine", "RuleResult", "load_rule_config"]
