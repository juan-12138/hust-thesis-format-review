"""Effective Word formatting resolution with provenance.

Resolution order follows the skill rule model: document defaults, table style,
paragraph style inheritance, character style inheritance, then direct formatting.
The resolver never mutates XML.
"""

from __future__ import annotations

from dataclasses import dataclass

from lxml import etree

from docx_model import DocxPackage, NS, ParagraphRecord, RunRecord, qn


RUN_BOOLEAN_PROPERTIES = {
    "b": "bold",
    "bCs": "bold_complex_script",
    "i": "italic",
    "iCs": "italic_complex_script",
    "strike": "strike",
    "dstrike": "double_strike",
    "outline": "outline",
    "shadow": "shadow",
    "emboss": "emboss",
    "imprint": "imprint",
    "smallCaps": "small_caps",
    "caps": "caps",
    "vanish": "hidden",
    "webHidden": "web_hidden",
    "noProof": "no_proofing",
}

PARAGRAPH_BOOLEAN_PROPERTIES = {
    "keepNext": "keep_with_next",
    "keepLines": "keep_lines_together",
    "widowControl": "widow_control",
    "pageBreakBefore": "page_break_before",
    "contextualSpacing": "contextual_spacing",
    "suppressLineNumbers": "suppress_line_numbers",
    "bidi": "bidi",
}


def _on_off(element: etree._Element) -> bool:
    value = element.get(qn("w", "val"))
    return value not in {"0", "false", "off", "no"}


def _number(element: etree._Element | None, attr: str, divisor=1.0):
    if element is None:
        return None
    value = element.get(qn("w", attr))
    try:
        return int(value) / divisor if value is not None else None
    except (TypeError, ValueError):
        return None


def _hex_byte(value: str | None) -> int | None:
    try:
        return int(value, 16) if value is not None else None
    except ValueError:
        return None


def apply_tint_shade(rgb: str, tint: str | None = None, shade: str | None = None) -> str:
    """Apply OOXML tint/shade bytes to an sRGB color."""
    rgb = rgb.upper().lstrip("#")
    if len(rgb) != 6:
        return rgb
    channels = [int(rgb[i : i + 2], 16) for i in (0, 2, 4)]
    shade_value = _hex_byte(shade)
    tint_value = _hex_byte(tint)
    if shade_value is not None:
        channels = [(channel * shade_value + 127) // 255 for channel in channels]
    if tint_value is not None:
        channels = [channel + ((255 - channel) * tint_value + 127) // 255 for channel in channels]
    return "".join(f"{max(0, min(255, channel)):02X}" for channel in channels)


@dataclass(frozen=True)
class StyleDefinition:
    style_id: str
    style_type: str
    based_on: str | None
    element: etree._Element


class ThemeResolver:
    def __init__(self, package: DocxPackage):
        self.fonts: dict[str, dict[str, str | None]] = {"major": {}, "minor": {}}
        self.colors: dict[str, str] = {}
        theme_parts = package.related_parts(DocxPackage.MAIN_PART, "theme")
        if not theme_parts and package.has_part("word/theme/theme1.xml"):
            theme_parts = ["word/theme/theme1.xml"]
        if theme_parts:
            self._parse(package.xml(theme_parts[0]))

    def _parse(self, root: etree._Element):
        scheme = root.find(".//a:fontScheme", NS)
        if scheme is not None:
            for kind, element_name in (("major", "majorFont"), ("minor", "minorFont")):
                group = scheme.find(f"a:{element_name}", NS)
                if group is None:
                    continue
                latin = group.find("a:latin", NS)
                east_asian = group.find("a:ea", NS)
                complex_script = group.find("a:cs", NS)
                hans = group.xpath('a:font[@script="Hans"]/@typeface', namespaces=NS)
                self.fonts[kind] = {
                    "latin": latin.get("typeface") if latin is not None else None,
                    "eastAsia": (east_asian.get("typeface") if east_asian is not None else None) or (hans[0] if hans else None),
                    "cs": complex_script.get("typeface") if complex_script is not None else None,
                }
        color_scheme = root.find(".//a:clrScheme", NS)
        if color_scheme is not None:
            for child in color_scheme:
                local_name = etree.QName(child).localname
                color = child.find("a:srgbClr", NS)
                if color is not None and color.get("val"):
                    self.colors[local_name] = color.get("val").upper()
                    continue
                system = child.find("a:sysClr", NS)
                if system is not None:
                    value = system.get("lastClr") or system.get("val")
                    if value:
                        self.colors[local_name] = value.upper()

    def font(self, token: str | None, slot: str) -> str | None:
        if not token:
            return None
        lowered = token.lower()
        kind = "major" if lowered.startswith("major") else "minor"
        if "eastasia" in lowered:
            family = "eastAsia"
        elif "bidi" in lowered:
            family = "cs"
        else:
            family = "eastAsia" if slot == "eastAsia" else ("cs" if slot == "cs" else "latin")
        return self.fonts.get(kind, {}).get(family)

    def color(self, token: str | None, tint: str | None = None, shade: str | None = None) -> str | None:
        if not token:
            return None
        aliases = {"text1": "dk1", "text2": "dk2", "background1": "lt1", "background2": "lt2", "followedHyperlink": "folHlink"}
        base = self.colors.get(aliases.get(token, token))
        return apply_tint_shade(base, tint, shade) if base else None


class EffectiveStyleResolver:
    def __init__(self, package: DocxPackage):
        self.package = package
        self.theme = ThemeResolver(package)
        self.styles: dict[tuple[str, str], StyleDefinition] = {}
        self.default_style_ids: dict[str, str] = {}
        self.default_run_properties = None
        self.default_paragraph_properties = None
        self._load_styles()

    def _load_styles(self):
        style_parts = self.package.related_parts(DocxPackage.MAIN_PART, "styles")
        if not style_parts and self.package.has_part("word/styles.xml"):
            style_parts = ["word/styles.xml"]
        if not style_parts:
            return
        root = self.package.xml(style_parts[0])
        self.default_run_properties = root.find("w:docDefaults/w:rPrDefault/w:rPr", NS)
        self.default_paragraph_properties = root.find("w:docDefaults/w:pPrDefault/w:pPr", NS)
        for style in root.findall("w:style", NS):
            style_id = style.get(qn("w", "styleId"))
            style_type = style.get(qn("w", "type"))
            if not style_id or not style_type:
                continue
            based = style.find("w:basedOn", NS)
            definition = StyleDefinition(
                style_id=style_id,
                style_type=style_type,
                based_on=based.get(qn("w", "val")) if based is not None else None,
                element=style,
            )
            self.styles[(style_type, style_id)] = definition
            is_default = style.get(qn("w", "default")) in {"1", "true", "on"}
            if is_default:
                self.default_style_ids[style_type] = style_id

    def style_chain(self, style_id: str | None, style_type: str) -> list[str]:
        if not style_id:
            style_id = self.default_style_ids.get(style_type)
        chain: list[str] = []
        visited: set[str] = set()
        current = style_id
        while current and current not in visited:
            definition = self.styles.get((style_type, current))
            if definition is None:
                break
            visited.add(current)
            chain.append(current)
            current = definition.based_on
        chain.reverse()
        return chain

    def _table_style_id(self, paragraph: ParagraphRecord) -> str | None:
        table = paragraph.element.xpath("ancestor::w:tbl[1]", namespaces=NS)
        if not table:
            return None
        style = table[0].find("w:tblPr/w:tblStyle", NS)
        return style.get(qn("w", "val")) if style is not None else None

    def _run_layers(self, run: RunRecord, paragraph: ParagraphRecord):
        if self.default_run_properties is not None:
            yield self.default_run_properties, "document-defaults", False
        table_style_id = self._table_style_id(paragraph)
        for style_id in self.style_chain(table_style_id, "table"):
            style = self.styles[("table", style_id)].element
            rpr = style.find("w:rPr", NS)
            if rpr is not None:
                yield rpr, f"table-style:{style_id}", True
            whole = style.find('w:tblStylePr[@w:type="wholeTable"]/w:rPr', NS)
            if whole is not None:
                yield whole, f"table-style:{style_id}:wholeTable", True
        for style_id in self.style_chain(paragraph.style_id, "paragraph"):
            rpr = self.styles[("paragraph", style_id)].element.find("w:rPr", NS)
            if rpr is not None:
                yield rpr, f"paragraph-style:{style_id}", True
        for style_id in self.style_chain(run.style_id, "character"):
            rpr = self.styles[("character", style_id)].element.find("w:rPr", NS)
            if rpr is not None:
                yield rpr, f"character-style:{style_id}", True
        direct = run.element.find("w:rPr", NS)
        if direct is not None:
            yield direct, "direct-run", False

    def get_effective_run_format(self, run: RunRecord, paragraph: ParagraphRecord | None = None) -> dict:
        paragraph = paragraph or run.paragraph
        result = {"fonts": {"ascii": None, "hAnsi": None, "eastAsia": None, "cs": None}, "provenance": {}}
        for rpr, source, style_toggle in self._run_layers(run, paragraph):
            self._merge_run_properties(result, rpr, source, style_toggle)
        return result

    def _set(self, result: dict, key: str, value, source: str):
        if value is None:
            return
        result[key] = value
        result["provenance"][key] = source

    def _merge_run_properties(self, result: dict, rpr: etree._Element, source: str, style_toggle: bool):
        fonts = rpr.find("w:rFonts", NS)
        if fonts is not None:
            for slot in ("ascii", "hAnsi", "eastAsia", "cs"):
                direct = fonts.get(qn("w", slot))
                theme_token = fonts.get(qn("w", slot + "Theme"))
                value = direct or self.theme.font(theme_token, slot)
                if value:
                    result["fonts"][slot] = value
                    result["provenance"][f"fonts.{slot}"] = source
        for xml_name, output_name in RUN_BOOLEAN_PROPERTIES.items():
            element = rpr.find(f"w:{xml_name}", NS)
            if element is None:
                continue
            value = _on_off(element)
            if style_toggle and value and output_name in result:
                value = not bool(result[output_name])
            self._set(result, output_name, value, source)
        size = rpr.find("w:sz", NS)
        self._set(result, "size_pt", _number(size, "val", 2), source)
        size_cs = rpr.find("w:szCs", NS)
        self._set(result, "size_complex_script_pt", _number(size_cs, "val", 2), source)
        underline = rpr.find("w:u", NS)
        if underline is not None:
            self._set(result, "underline", underline.get(qn("w", "val"), "single"), source)
        vertical = rpr.find("w:vertAlign", NS)
        if vertical is not None:
            self._set(result, "vertical_align", vertical.get(qn("w", "val")), source)
        spacing = rpr.find("w:spacing", NS)
        self._set(result, "character_spacing_pt", _number(spacing, "val", 20), source)
        scale = rpr.find("w:w", NS)
        scale_value = _number(scale, "val", 1)
        self._set(result, "character_scale_percent", int(scale_value) if scale_value is not None else None, source)
        position = rpr.find("w:position", NS)
        self._set(result, "position_pt", _number(position, "val", 2), source)
        highlight = rpr.find("w:highlight", NS)
        if highlight is not None:
            self._set(result, "highlight", highlight.get(qn("w", "val")), source)
        color = rpr.find("w:color", NS)
        if color is not None:
            direct = color.get(qn("w", "val"))
            if direct and direct.lower() != "auto":
                resolved = direct.upper()
            else:
                resolved = self.theme.color(
                    color.get(qn("w", "themeColor")),
                    color.get(qn("w", "themeTint")),
                    color.get(qn("w", "themeShade")),
                ) or ("auto" if direct and direct.lower() == "auto" else None)
            self._set(result, "color_hex", resolved, source)

    def _paragraph_layers(self, paragraph: ParagraphRecord):
        if self.default_paragraph_properties is not None:
            yield self.default_paragraph_properties, "document-defaults"
        table_style_id = self._table_style_id(paragraph)
        for style_id in self.style_chain(table_style_id, "table"):
            style = self.styles[("table", style_id)].element
            ppr = style.find("w:pPr", NS)
            if ppr is not None:
                yield ppr, f"table-style:{style_id}"
            whole = style.find('w:tblStylePr[@w:type="wholeTable"]/w:pPr', NS)
            if whole is not None:
                yield whole, f"table-style:{style_id}:wholeTable"
        for style_id in self.style_chain(paragraph.style_id, "paragraph"):
            ppr = self.styles[("paragraph", style_id)].element.find("w:pPr", NS)
            if ppr is not None:
                yield ppr, f"paragraph-style:{style_id}"
        direct = paragraph.element.find("w:pPr", NS)
        if direct is not None:
            yield direct, "direct-paragraph"

    def get_effective_paragraph_format(self, paragraph: ParagraphRecord) -> dict:
        result = {"tabs": [], "provenance": {}}
        for ppr, source in self._paragraph_layers(paragraph):
            self._merge_paragraph_properties(result, ppr, source)
        return result

    def _merge_paragraph_properties(self, result: dict, ppr: etree._Element, source: str):
        alignment = ppr.find("w:jc", NS)
        if alignment is not None:
            self._set(result, "alignment", alignment.get(qn("w", "val")), source)
        spacing = ppr.find("w:spacing", NS)
        if spacing is not None:
            self._set(result, "space_before_pt", _number(spacing, "before", 20), source)
            self._set(result, "space_after_pt", _number(spacing, "after", 20), source)
            line = _number(spacing, "line", 1)
            line_rule = spacing.get(qn("w", "lineRule"), "auto") if line is not None else None
            self._set(result, "line_rule", line_rule, source)
            if line is not None:
                if line_rule == "auto":
                    self._set(result, "line_spacing_multiple", line / 240.0, source)
                    result.pop("line_spacing_pt", None)
                    result["provenance"].pop("line_spacing_pt", None)
                else:
                    self._set(result, "line_spacing_pt", line / 20.0, source)
                    result.pop("line_spacing_multiple", None)
                    result["provenance"].pop("line_spacing_multiple", None)
        indentation = ppr.find("w:ind", NS)
        if indentation is not None:
            for attr, output in (("left", "left_indent_pt"), ("right", "right_indent_pt"), ("firstLine", "first_line_indent_pt"), ("hanging", "hanging_indent_pt")):
                self._set(result, output, _number(indentation, attr, 20), source)
        for xml_name, output_name in PARAGRAPH_BOOLEAN_PROPERTIES.items():
            element = ppr.find(f"w:{xml_name}", NS)
            if element is not None:
                self._set(result, output_name, _on_off(element), source)
        outline = ppr.find("w:outlineLvl", NS)
        value = _number(outline, "val", 1)
        self._set(result, "outline_level", int(value) if value is not None else None, source)
        num_pr = ppr.find("w:numPr", NS)
        if num_pr is not None:
            ilvl = _number(num_pr.find("w:ilvl", NS), "val", 1)
            num_id = _number(num_pr.find("w:numId", NS), "val", 1)
            self._set(result, "numbering_level", int(ilvl) if ilvl is not None else None, source)
            self._set(result, "numbering_id", int(num_id) if num_id is not None else None, source)
        tabs = ppr.find("w:tabs", NS)
        if tabs is not None:
            values = []
            for tab in tabs.findall("w:tab", NS):
                values.append(
                    {
                        "alignment": tab.get(qn("w", "val")),
                        "position_pt": _number(tab, "pos", 20),
                        "leader": tab.get(qn("w", "leader")),
                    }
                )
            result["tabs"] = values
            result["provenance"]["tabs"] = source
        text_direction = ppr.find("w:textDirection", NS)
        if text_direction is not None:
            self._set(result, "text_direction", text_direction.get(qn("w", "val")), source)


__all__ = ["EffectiveStyleResolver", "StyleDefinition", "ThemeResolver", "apply_tint_shade"]
