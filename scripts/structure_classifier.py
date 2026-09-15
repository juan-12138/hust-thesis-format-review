"""Deterministic, evidence-carrying paragraph role classification."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Iterable

from docx_model import DocxPackage, NS, ParagraphRecord, qn
from effective_style import EffectiveStyleResolver


@dataclass(frozen=True)
class Classification:
    role: str
    confidence: float
    evidence: str


@dataclass(frozen=True)
class ClassifiedParagraph:
    paragraph: ParagraphRecord
    classification: Classification


def _compact(text: str) -> str:
    return re.sub(r"[\s\u3000]+", "", text or "")


class StructureClassifier:
    """Classify paragraphs without pretending to understand page layout."""

    FIGURE_RE = re.compile(r"^图\s*\d+\s*[-－—.]\s*\d+")
    TABLE_RE = re.compile(r"^(?:续\s*)?表\s*\d+\s*[-－—.]\s*\d+")
    REFERENCE_RE = re.compile(r"^\s*[\[［【]\s*\d+\s*[\]］】]")
    MAIN_HEADING_RE = re.compile(r"^\s*\d+(?:\s+|[、.．])")

    def __init__(self, package: DocxPackage, resolver: EffectiveStyleResolver | None = None):
        self.package = package
        self.resolver = resolver or EffectiveStyleResolver(package)

    def _style_name(self, paragraph: ParagraphRecord) -> str:
        if not paragraph.style_id:
            return ""
        definition = self.resolver.styles.get(("paragraph", paragraph.style_id))
        if definition is None:
            return paragraph.style_id
        name = definition.element.find("w:name", NS)
        return name.get(qn("w", "val"), paragraph.style_id) if name is not None else paragraph.style_id

    def _heading_level(self, paragraph: ParagraphRecord) -> int | None:
        style_name = self._style_name(paragraph).lower().replace(" ", "")
        style_id = (paragraph.style_id or "").lower().replace(" ", "")
        for level in (1, 2, 3, 4, 5, 6, 7, 8, 9):
            aliases = {f"heading{level}", f"标题{level}"}
            if style_name in aliases or style_id in aliases:
                return level
        outline = self.resolver.get_effective_paragraph_format(paragraph).get("outline_level")
        if isinstance(outline, int) and 0 <= outline <= 8:
            return outline + 1
        return None

    def classify_document(self, paragraphs: Iterable[ParagraphRecord] | None = None) -> list[ClassifiedParagraph]:
        paragraphs = list(paragraphs if paragraphs is not None else self.package.iter_paragraphs())
        context = "front"
        results: list[ClassifiedParagraph] = []
        for paragraph in paragraphs:
            text = paragraph.text.strip()
            compact = _compact(text)
            style_name = self._style_name(paragraph)
            style_key = style_name.lower().replace(" ", "")

            if not text:
                classification = Classification("empty", 1.0, "段落无可见文字")
            elif compact == "摘要":
                context = "abstract_cn"
                classification = Classification("abstract_title_cn", 1.0, "标题文字为“摘要”")
            elif compact.lower() == "abstract":
                context = "abstract_en"
                classification = Classification("abstract_title_en", 1.0, "标题文字为“Abstract”")
            elif re.match(r"^关键词\s*[:：]", text):
                classification = Classification("keywords_cn", 1.0, "段落以“关键词：”开头")
            elif re.match(r"^key\s*words?\s*[:：]", text, re.I):
                classification = Classification("keywords_en", 1.0, "段落以“Key words:”开头")
            elif compact == "目录":
                context = "toc"
                classification = Classification("toc_title", 1.0, "标题文字为“目录”")
            elif style_key.startswith("toc") or style_key.startswith("目录"):
                classification = Classification("toc_entry", 0.99, f"目录样式：{style_name}")
            elif compact == "参考文献":
                context = "references"
                classification = Classification("references_title", 1.0, "标题文字为“参考文献”")
            elif compact == "致谢":
                context = "acknowledgements"
                classification = Classification("acknowledgements_title", 1.0, "标题文字为“致谢”")
            elif re.match(r"^附录\s*[A-Za-z0-9一二三四五六七八九十]*", text):
                context = "appendix"
                classification = Classification("appendix_heading", 0.99, "段落以“附录”开头")
            elif self.FIGURE_RE.match(text):
                classification = Classification("figure_caption", 0.98, "文字符合按章编号的图题模式")
            elif self.TABLE_RE.match(text):
                classification = Classification("table_caption", 0.98, "文字符合按章编号的表题模式")
            elif self.REFERENCE_RE.match(text) or context == "references":
                classification = Classification("reference_entry", 0.96 if self.REFERENCE_RE.match(text) else 0.72, "参考文献编号模式或位于参考文献区域")
            else:
                heading_level = self._heading_level(paragraph)
                if heading_level is not None:
                    context = "main" if self.MAIN_HEADING_RE.match(text) or context in {"toc", "main"} else context
                    classification = Classification(f"heading_{heading_level}", 0.97, f"标题样式或大纲级别：{style_name or paragraph.style_id}")
                elif paragraph.element.xpath(".//m:oMath | .//m:oMathPara", namespaces=NS):
                    classification = Classification("equation", 0.95, "段落中包含 OMML 公式")
                elif context == "abstract_cn":
                    classification = Classification("abstract_body_cn", 0.90, "位于中文摘要标题之后")
                elif context == "abstract_en":
                    classification = Classification("abstract_body_en", 0.90, "位于英文摘要标题之后")
                elif context == "toc":
                    classification = Classification("toc_related_text", 0.65, "位于目录区域但未使用目录条目样式")
                elif context == "acknowledgements":
                    classification = Classification("acknowledgements_body", 0.85, "位于致谢标题之后")
                elif context == "appendix":
                    classification = Classification("appendix_body", 0.80, "位于附录标题之后")
                else:
                    classification = Classification("body", 0.75, "未命中更具体的结构规则")
            results.append(ClassifiedParagraph(paragraph, classification))
        return results
