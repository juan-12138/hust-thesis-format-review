"""Read-only OOXML structure model for Word DOCX files.

The module deliberately works below python-docx so it can see fields, drawings,
revisions, text boxes, content controls, relationships, and section properties
without reserializing the source package.
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
import re
from typing import Iterator
from zipfile import BadZipFile, ZipFile

try:
    from lxml import etree
except ModuleNotFoundError as exc:  # pragma: no cover - depends on caller runtime
    raise ModuleNotFoundError(
        "阶段3 DOCX 解析器需要 lxml。请先加载 Codex 工作区依赖，"
        "并使用返回的 Python executable 运行本脚本。"
    ) from exc


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "v": "urn:schemas-microsoft-com:vml",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "o": "urn:schemas-microsoft-com:office:office",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
}


def qn(prefix: str, local: str) -> str:
    return f"{{{NS[prefix]}}}{local}"


def _int_attr(element, name: str, default=None):
    if element is None:
        return default
    value = element.get(qn("w", name))
    try:
        return int(value) if value is not None else default
    except ValueError:
        return default


def _bool_element(parent, local: str, default=False):
    element = parent.find(f"w:{local}", NS) if parent is not None else None
    if element is None:
        return default
    value = element.get(qn("w", "val"))
    return value not in {"0", "false", "off", "no"}


def relationship_part_name(source_part: str) -> str:
    source = PurePosixPath(source_part)
    return str(source.parent / "_rels" / f"{source.name}.rels")


def resolve_target_part(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    source_dir = PurePosixPath(source_part).parent
    combined = source_dir / target
    normalized: list[str] = []
    for item in combined.parts:
        if item in {"", "."}:
            continue
        if item == "..":
            if normalized:
                normalized.pop()
        else:
            normalized.append(item)
    return "/".join(normalized)


@dataclass(frozen=True)
class Relationship:
    relationship_id: str
    relationship_type: str
    target: str
    target_part: str | None
    external: bool


@dataclass
class ParagraphRecord:
    part_name: str
    index: int
    element: etree._Element
    text: str
    style_id: str | None
    section_index: int | None
    in_table: bool
    in_content_control: bool
    in_revision: bool

    @property
    def location(self) -> str:
        snippet = " ".join(self.text.split())[:40]
        return f"{self.part_name} 第{self.index + 1}段" + (f"：{snippet}" if snippet else "")


@dataclass
class RunRecord:
    part_name: str
    paragraph_index: int
    index: int
    element: etree._Element
    paragraph: ParagraphRecord
    text: str
    style_id: str | None


class DocxPackage:
    """Read-only facade over an OPC/DOCX ZIP package."""

    MAIN_PART = "word/document.xml"

    def __init__(self, path: str | Path):
        self.path = Path(path)
        if not self.path.is_file():
            raise FileNotFoundError(self.path)
        try:
            with ZipFile(self.path) as zf:
                self._parts = {name: zf.read(name) for name in zf.namelist()}
        except BadZipFile as exc:
            raise ValueError(f"不是有效的 DOCX/ZIP 文件：{self.path}") from exc
        if "[Content_Types].xml" not in self._parts or self.MAIN_PART not in self._parts:
            raise ValueError("DOCX 缺少 [Content_Types].xml 或 word/document.xml")
        self._xml_cache: dict[str, etree._Element] = {}
        self._rels_cache: dict[str, dict[str, Relationship]] = {}

    @property
    def part_names(self) -> tuple[str, ...]:
        return tuple(self._parts)

    def has_part(self, part_name: str) -> bool:
        return part_name in self._parts

    def part_bytes(self, part_name: str) -> bytes:
        return self._parts[part_name]

    def xml(self, part_name: str) -> etree._Element:
        if part_name not in self._xml_cache:
            raw = self._parts[part_name]
            parser = etree.XMLParser(resolve_entities=False, no_network=True, recover=False, huge_tree=True)
            self._xml_cache[part_name] = etree.fromstring(raw, parser=parser)
        return self._xml_cache[part_name]

    def relationships(self, source_part: str) -> dict[str, Relationship]:
        if source_part in self._rels_cache:
            return self._rels_cache[source_part]
        rel_part = relationship_part_name(source_part)
        relationships: dict[str, Relationship] = {}
        if rel_part in self._parts:
            root = self.xml(rel_part)
            for rel in root.findall("rel:Relationship", NS):
                rel_id = rel.get("Id", "")
                target = rel.get("Target", "")
                external = rel.get("TargetMode") == "External"
                relationships[rel_id] = Relationship(
                    relationship_id=rel_id,
                    relationship_type=rel.get("Type", ""),
                    target=target,
                    target_part=None if external else resolve_target_part(source_part, target),
                    external=external,
                )
        self._rels_cache[source_part] = relationships
        return relationships

    def related_parts(self, source_part: str, relationship_suffix: str) -> list[str]:
        return [
            rel.target_part
            for rel in self.relationships(source_part).values()
            if rel.relationship_type.endswith("/" + relationship_suffix) and rel.target_part
        ]

    def _paragraph_parts(self) -> list[str]:
        preferred = [self.MAIN_PART]
        rels = self.relationships(self.MAIN_PART)
        for rel in rels.values():
            if rel.target_part and any(rel.relationship_type.endswith("/" + kind) for kind in ("header", "footer")):
                preferred.append(rel.target_part)
        for part in ("word/footnotes.xml", "word/endnotes.xml", "word/comments.xml"):
            if part in self._parts:
                preferred.append(part)
        seen = set()
        return [part for part in preferred if not (part in seen or seen.add(part))]

    @staticmethod
    def paragraph_text(paragraph: etree._Element) -> str:
        pieces: list[str] = []
        for node in paragraph.iter():
            if node.tag in {qn("w", "t"), qn("w", "delText"), qn("w", "instrText"), qn("w", "delInstrText")}:
                pieces.append(node.text or "")
            elif node.tag == qn("w", "tab"):
                pieces.append("\t")
            elif node.tag == qn("w", "br"):
                pieces.append("\n")
        return "".join(pieces)

    def iter_paragraphs(self, part_name: str = MAIN_PART) -> Iterator[ParagraphRecord]:
        root = self.xml(part_name)
        paragraphs = root.xpath(".//w:p", namespaces=NS)
        section_indices: list[int | None] = []
        current_section = 0
        if part_name == self.MAIN_PART:
            for paragraph in paragraphs:
                section_indices.append(current_section)
                if paragraph.find("w:pPr/w:sectPr", NS) is not None:
                    current_section += 1
        else:
            section_indices = [None] * len(paragraphs)
        for index, paragraph in enumerate(paragraphs):
            style = paragraph.find("w:pPr/w:pStyle", NS)
            yield ParagraphRecord(
                part_name=part_name,
                index=index,
                element=paragraph,
                text=self.paragraph_text(paragraph),
                style_id=style.get(qn("w", "val")) if style is not None else None,
                section_index=section_indices[index],
                in_table=bool(paragraph.xpath("ancestor::w:tc", namespaces=NS)),
                in_content_control=bool(paragraph.xpath("ancestor::w:sdt", namespaces=NS)),
                in_revision=bool(paragraph.xpath("ancestor::w:ins or ancestor::w:del", namespaces=NS)),
            )

    def iter_all_paragraphs(self) -> Iterator[ParagraphRecord]:
        for part_name in self._paragraph_parts():
            yield from self.iter_paragraphs(part_name)

    def iter_runs(self, paragraph: ParagraphRecord) -> Iterator[RunRecord]:
        for index, run in enumerate(paragraph.element.xpath(".//w:r", namespaces=NS)):
            style = run.find("w:rPr/w:rStyle", NS)
            text = "".join(run.xpath(".//w:t/text() | .//w:delText/text()", namespaces=NS))
            yield RunRecord(
                part_name=paragraph.part_name,
                paragraph_index=paragraph.index,
                index=index,
                element=run,
                paragraph=paragraph,
                text=text,
                style_id=style.get(qn("w", "val")) if style is not None else None,
            )

    def sections(self) -> list[dict]:
        root = self.xml(self.MAIN_PART)
        rels = self.relationships(self.MAIN_PART)
        sections = []
        for index, sect in enumerate(root.xpath(".//w:sectPr", namespaces=NS)):
            page_size = sect.find("w:pgSz", NS)
            margins = sect.find("w:pgMar", NS)
            page_num = sect.find("w:pgNumType", NS)
            section_type = sect.find("w:type", NS)
            result = {
                "section_index": index,
                "page_width_twips": _int_attr(page_size, "w"),
                "page_height_twips": _int_attr(page_size, "h"),
                "orientation": page_size.get(qn("w", "orient"), "portrait") if page_size is not None else None,
                "margin_top_twips": _int_attr(margins, "top"),
                "margin_right_twips": _int_attr(margins, "right"),
                "margin_bottom_twips": _int_attr(margins, "bottom"),
                "margin_left_twips": _int_attr(margins, "left"),
                "gutter_twips": _int_attr(margins, "gutter", 0),
                "header_distance_twips": _int_attr(margins, "header"),
                "footer_distance_twips": _int_attr(margins, "footer"),
                "section_start": section_type.get(qn("w", "val"), "nextPage") if section_type is not None else "nextPage",
                "different_first_page": sect.find("w:titlePg", NS) is not None,
                "page_number_format": page_num.get(qn("w", "fmt")) if page_num is not None else None,
                "page_number_start": _int_attr(page_num, "start"),
                "header_references": {},
                "footer_references": {},
            }
            for kind in ("header", "footer"):
                for ref in sect.findall(f"w:{kind}Reference", NS):
                    rel_id = ref.get(qn("r", "id"))
                    rel = rels.get(rel_id)
                    if rel and rel.target_part:
                        result[f"{kind}_references"][ref.get(qn("w", "type"), "default")] = rel.target_part
            sections.append(result)
        previous_headers: dict[str, str] = {}
        previous_footers: dict[str, str] = {}
        for index, section in enumerate(sections):
            explicit_headers = section["header_references"]
            explicit_footers = section["footer_references"]
            effective_headers = dict(previous_headers) if index else {}
            effective_footers = dict(previous_footers) if index else {}
            effective_headers.update(explicit_headers)
            effective_footers.update(explicit_footers)
            section["header_linked_to_previous"] = index > 0 and "default" not in explicit_headers
            section["footer_linked_to_previous"] = index > 0 and "default" not in explicit_footers
            section["effective_header_references"] = effective_headers
            section["effective_footer_references"] = effective_footers
            previous_headers = effective_headers
            previous_footers = effective_footers
        return sections

    def settings(self) -> dict:
        if "word/settings.xml" not in self._parts:
            return {
                "default_tab_stop_twips": None,
                "even_and_odd_headers": False,
                "track_revisions": False,
                "update_fields_on_open": False,
                "mirror_margins": False,
            }
        root = self.xml("word/settings.xml")
        default_tab = root.find("w:defaultTabStop", NS)
        update_fields = root.find("w:updateFields", NS)
        return {
            "default_tab_stop_twips": _int_attr(default_tab, "val"),
            "even_and_odd_headers": root.find("w:evenAndOddHeaders", NS) is not None,
            "track_revisions": root.find("w:trackRevisions", NS) is not None,
            "update_fields_on_open": _bool_element(root, "updateFields"),
            "mirror_margins": root.find("w:mirrorMargins", NS) is not None,
            "compatibility_mode": self._compatibility_mode(root),
        }

    @staticmethod
    def _compatibility_mode(root: etree._Element):
        values = root.xpath('.//w:compatSetting[@w:name="compatibilityMode"]/@w:val', namespaces=NS)
        try:
            return int(values[-1]) if values else None
        except ValueError:
            return values[-1] if values else None

    def numbering_summary(self) -> dict:
        if "word/numbering.xml" not in self._parts:
            return {"abstract_numbering_definitions": 0, "numbering_instances": 0, "levels": []}
        root = self.xml("word/numbering.xml")
        levels = []
        for abstract in root.findall("w:abstractNum", NS):
            abstract_id = _int_attr(abstract, "abstractNumId")
            for level in abstract.findall("w:lvl", NS):
                fmt = level.find("w:numFmt", NS)
                text = level.find("w:lvlText", NS)
                start = level.find("w:start", NS)
                levels.append(
                    {
                        "abstract_numbering_id": abstract_id,
                        "level": _int_attr(level, "ilvl"),
                        "format": fmt.get(qn("w", "val")) if fmt is not None else None,
                        "text": text.get(qn("w", "val")) if text is not None else None,
                        "start": _int_attr(start, "val"),
                    }
                )
        return {
            "abstract_numbering_definitions": len(root.findall("w:abstractNum", NS)),
            "numbering_instances": len(root.findall("w:num", NS)),
            "levels": levels,
        }

    def styles_summary(self) -> dict:
        if "word/styles.xml" not in self._parts:
            return {"paragraph_styles": 0, "character_styles": 0, "table_styles": 0, "numbering_styles": 0, "styles": []}
        root = self.xml("word/styles.xml")
        counter = Counter()
        styles = []
        for element in root.findall("w:style", NS):
            style_type = element.get(qn("w", "type"), "unknown")
            counter[style_type] += 1
            name = element.find("w:name", NS)
            based = element.find("w:basedOn", NS)
            styles.append(
                {
                    "style_id": element.get(qn("w", "styleId")),
                    "type": style_type,
                    "name": name.get(qn("w", "val")) if name is not None else None,
                    "based_on": based.get(qn("w", "val")) if based is not None else None,
                    "default": element.get(qn("w", "default")) in {"1", "true", "on"},
                }
            )
        return {
            "paragraph_styles": counter["paragraph"],
            "character_styles": counter["character"],
            "table_styles": counter["table"],
            "numbering_styles": counter["numbering"],
            "styles": styles,
        }

    def relationship_inventory(self) -> dict:
        result = {}
        source_parts = [self.MAIN_PART]
        source_parts.extend(part for part in self._paragraph_parts() if part != self.MAIN_PART)
        for source in source_parts:
            relationships = self.relationships(source)
            if relationships:
                result[source] = [
                    {
                        "id": rel.relationship_id,
                        "type": rel.relationship_type.rsplit("/", 1)[-1],
                        "target": rel.target,
                        "target_part": rel.target_part,
                        "external": rel.external,
                    }
                    for rel in relationships.values()
                ]
        return result

    def _field_counter(self) -> Counter:
        counter: Counter = Counter()
        for part_name in self._paragraph_parts():
            root = self.xml(part_name)
            for instruction in root.xpath(".//w:fldSimple/@w:instr", namespaces=NS):
                self._count_field_instruction(counter, instruction)
            instruction_stack: list[list[str]] = []
            for node in root.iter():
                if node.tag == qn("w", "fldChar"):
                    field_type = node.get(qn("w", "fldCharType"))
                    if field_type == "begin":
                        instruction_stack.append([])
                    elif field_type == "end" and instruction_stack:
                        instruction = "".join(instruction_stack.pop())
                        self._count_field_instruction(counter, instruction)
                elif node.tag in {qn("w", "instrText"), qn("w", "delInstrText")} and instruction_stack:
                    instruction_stack[-1].append(node.text or "")
        return counter

    @staticmethod
    def _count_field_instruction(counter: Counter, instruction: str):
        command = (instruction or "").strip().split()
        if command:
            counter[command[0].upper()] += 1

    def inventory(self) -> dict:
        main = self.xml(self.MAIN_PART)
        all_roots = [self.xml(part) for part in self._paragraph_parts()]
        field_counter = self._field_counter()
        text_values = [text for root in all_roots for text in root.xpath(".//w:t/text() | .//w:delText/text()", namespaces=NS)]
        all_paragraphs = list(self.iter_all_paragraphs())
        return {
            "parts": len(self._parts),
            "paragraphs_body": len(list(self.iter_paragraphs(self.MAIN_PART))),
            "paragraphs_all": len(list(self.iter_all_paragraphs())),
            "tables": len(main.xpath(".//w:tbl", namespaces=NS)),
            "sections": len(main.xpath(".//w:sectPr", namespaces=NS)),
            "content_controls": sum(len(root.xpath(".//w:sdt", namespaces=NS)) for root in all_roots),
            "tracked_insertions": sum(len(root.xpath(".//w:ins", namespaces=NS)) for root in all_roots),
            "tracked_deletions": sum(len(root.xpath(".//w:del", namespaces=NS)) for root in all_roots),
            "drawings_inline": sum(len(root.xpath(".//wp:inline", namespaces=NS)) for root in all_roots),
            "drawings_anchor": sum(len(root.xpath(".//wp:anchor", namespaces=NS)) for root in all_roots),
            "vml_shapes": sum(len(root.xpath(".//v:shape", namespaces=NS)) for root in all_roots),
            "omml_objects": sum(len(root.xpath(".//m:oMath | .//m:oMathPara", namespaces=NS)) for root in all_roots),
            "fields": dict(sorted(field_counter.items())),
            "bookmarks": sum(len(root.xpath(".//w:bookmarkStart", namespaces=NS)) for root in all_roots),
            "hyperlinks": sum(len(root.xpath(".//w:hyperlink", namespaces=NS)) for root in all_roots),
            "footnotes": self._note_count("word/footnotes.xml", "footnote"),
            "endnotes": self._note_count("word/endnotes.xml", "endnote"),
            "comments": self._note_count("word/comments.xml", "comment", include_nonpositive=True),
            "hidden_runs": sum(len(root.xpath(".//w:r[w:rPr/w:vanish]", namespaces=NS)) for root in all_roots),
            "soft_line_breaks": sum(len(root.xpath('.//w:br[not(@w:type) or @w:type="textWrapping"]', namespaces=NS)) for root in all_roots),
            "manual_page_breaks": sum(len(root.xpath('.//w:br[@w:type="page"]', namespaces=NS)) for root in all_roots),
            "manual_column_breaks": sum(len(root.xpath('.//w:br[@w:type="column"]', namespaces=NS)) for root in all_roots),
            "paragraph_section_breaks": len(main.xpath(".//w:pPr/w:sectPr", namespaces=NS)),
            "ole_objects": sum(len(root.xpath(".//o:OLEObject", namespaces=NS)) for root in all_roots),
            "text_boxes": sum(len(root.xpath(".//w:txbxContent", namespaces=NS)) for root in all_roots),
            "smartart_relationship_markers": sum(len(root.xpath(".//dgm:relIds", namespaces=NS)) for root in all_roots),
            "nonbreaking_spaces": sum(text.count("\u00a0") for text in text_values),
            "fullwidth_spaces": sum(text.count("\u3000") for text in text_values),
            "consecutive_space_sequences": sum(len(re.findall(r" {2,}", text)) for text in text_values),
            "empty_paragraphs": sum(not paragraph.text.strip() for paragraph in all_paragraphs),
        }

    def _note_count(self, part_name: str, local: str, include_nonpositive=False) -> int:
        if part_name not in self._parts:
            return 0
        count = 0
        for element in self.xml(part_name).xpath(f".//w:{local}", namespaces=NS):
            value = _int_attr(element, "id")
            if include_nonpositive or value is None or value > 0:
                count += 1
        return count


__all__ = ["DocxPackage", "NS", "ParagraphRecord", "Relationship", "RunRecord", "qn"]
