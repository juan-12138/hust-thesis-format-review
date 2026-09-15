import hashlib
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from .helpers import MINIMAL_STYLES, MINIMAL_THEME, make_docx


DOCUMENT = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:o="urn:schemas-microsoft-com:office:office">
 <w:body>
  <w:p w:rsidR="001"><w:bookmarkStart w:id="0" w:name="intro"/><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>第一章</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>
  <w:sdt><w:sdtContent><w:p><w:r><w:t>内容控件文字</w:t></w:r></w:p></w:sdtContent></w:sdt>
  <w:p><w:hyperlink r:id="rIdLink"><w:r><w:t>链接</w:t></w:r></w:hyperlink><w:ins w:id="1"><w:r><w:t>新增</w:t></w:r></w:ins><w:del w:id="2"><w:r><w:delText>删除</w:delText></w:r></w:del><w:r><w:drawing><wp:inline/></w:drawing></w:r></w:p>
  <w:tbl><w:tr><w:tc><w:p><w:r><w:t>单元格</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
  <w:p><w:fldSimple w:instr=" REF intro "><w:r><w:t>第一章</w:t></w:r></w:fldSimple></w:p>
  <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="2549" w:right="1584" w:bottom="1584" w:left="1584" w:header="851" w:footer="964" w:gutter="0"/><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/><w:pgNumType w:fmt="decimal" w:start="1"/><w:titlePg/></w:sectPr>
 </w:body>
</w:document>"""


class DocxModelTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.path = make_docx(Path(self.temp.name) / "fixture.docx", DOCUMENT, MINIMAL_STYLES, MINIMAL_THEME)

    def tearDown(self):
        self.temp.cleanup()

    def test_parses_body_and_auxiliary_structures_without_modifying_input(self):
        from docx_model import DocxPackage

        before = hashlib.sha256(self.path.read_bytes()).hexdigest()
        package = DocxPackage(self.path)
        body = list(package.iter_paragraphs("word/document.xml"))
        all_paragraphs = list(package.iter_all_paragraphs())
        inventory = package.inventory()

        self.assertEqual([p.text for p in body], ["第一章", "内容控件文字", "链接新增删除", "单元格", "第一章"])
        self.assertEqual(len(all_paragraphs), 9)
        self.assertEqual(inventory["tables"], 1)
        self.assertEqual(inventory["content_controls"], 1)
        self.assertEqual(inventory["tracked_insertions"], 1)
        self.assertEqual(inventory["tracked_deletions"], 1)
        self.assertEqual(inventory["drawings_inline"], 1)
        self.assertEqual(inventory["fields"]["REF"], 1)
        self.assertEqual(inventory["bookmarks"], 1)
        self.assertEqual(inventory["hyperlinks"], 1)
        self.assertEqual(hashlib.sha256(self.path.read_bytes()).hexdigest(), before)

    def test_resolves_relationships_and_section_properties(self):
        from docx_model import DocxPackage

        package = DocxPackage(self.path)
        rels = package.relationships("word/document.xml")
        section = package.sections()[0]

        self.assertEqual(rels["rIdHeader"].target_part, "word/header1.xml")
        self.assertTrue(rels["rIdLink"].external)
        self.assertEqual(section["page_width_twips"], 11906)
        self.assertEqual(section["margin_top_twips"], 2549)
        self.assertEqual(section["page_number_format"], "decimal")
        self.assertEqual(section["page_number_start"], 1)
        self.assertTrue(section["different_first_page"])
        self.assertEqual(section["header_references"]["default"], "word/header1.xml")

    def test_parses_settings_numbering_styles_and_extended_object_inventory(self):
        from docx_model import DocxPackage

        document = DOCUMENT.replace(
            "<w:sectPr>",
            '<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t xml:space="preserve">隐藏  连续\u3000空格\u00a0</w:t><w:br/><w:br w:type="page"/></w:r><w:object><o:OLEObject/></w:object><w:txbxContent><w:p><w:r><w:t>文本框</w:t></w:r></w:p></w:txbxContent></w:p><w:sectPr>',
        )
        settings = """<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:defaultTabStop w:val="420"/><w:evenAndOddHeaders/><w:trackRevisions/><w:updateFields w:val="1"/></w:settings>"""
        numbering = """<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"""
        comments = """<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0"><w:p><w:r><w:t>原批注</w:t></w:r></w:p></w:comment></w:comments>"""
        rich_path = make_docx(
            Path(self.temp.name) / "rich.docx",
            document,
            MINIMAL_STYLES,
            MINIMAL_THEME,
            {"word/settings.xml": settings, "word/numbering.xml": numbering, "word/comments.xml": comments},
        )
        package = DocxPackage(rich_path)
        inventory = package.inventory()

        self.assertEqual(package.settings()["default_tab_stop_twips"], 420)
        self.assertTrue(package.settings()["even_and_odd_headers"])
        self.assertTrue(package.settings()["track_revisions"])
        self.assertTrue(package.settings()["update_fields_on_open"])
        self.assertEqual(package.numbering_summary()["abstract_numbering_definitions"], 1)
        self.assertEqual(package.numbering_summary()["numbering_instances"], 1)
        self.assertEqual(package.numbering_summary()["levels"][0]["text"], "%1")
        self.assertGreaterEqual(package.styles_summary()["paragraph_styles"], 3)
        self.assertEqual(inventory["comments"], 1)
        self.assertEqual(inventory["hidden_runs"], 1)
        self.assertEqual(inventory["soft_line_breaks"], 1)
        self.assertEqual(inventory["manual_page_breaks"], 1)
        self.assertEqual(inventory["ole_objects"], 1)
        self.assertEqual(inventory["text_boxes"], 1)
        self.assertEqual(inventory["nonbreaking_spaces"], 1)
        self.assertEqual(inventory["fullwidth_spaces"], 1)
        self.assertEqual(inventory["consecutive_space_sequences"], 1)

    def test_section_header_footer_references_are_resolved_through_link_to_previous(self):
        from docx_model import DocxPackage

        two_sections = """<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
          <w:p><w:r><w:t>第一节</w:t></w:r><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr></w:pPr></w:p>
          <w:p><w:r><w:t>第二节</w:t></w:r></w:p>
          <w:sectPr/>
        </w:body></w:document>"""
        path = make_docx(Path(self.temp.name) / "sections.docx", two_sections, MINIMAL_STYLES, MINIMAL_THEME)
        package = DocxPackage(path)
        sections = package.sections()
        paragraph_sections = [paragraph.section_index for paragraph in package.iter_paragraphs()]

        self.assertFalse(sections[0]["header_linked_to_previous"])
        self.assertEqual(sections[0]["effective_header_references"]["default"], "word/header1.xml")
        self.assertTrue(sections[1]["header_linked_to_previous"])
        self.assertTrue(sections[1]["footer_linked_to_previous"])
        self.assertEqual(sections[1]["effective_header_references"]["default"], "word/header1.xml")
        self.assertEqual(sections[1]["effective_footer_references"]["default"], "word/footer1.xml")
        self.assertEqual(paragraph_sections, [0, 1])

    def test_fragmented_complex_field_instruction_is_counted_once(self):
        from docx_model import DocxPackage

        split_field = DOCUMENT.replace(
            "<w:sectPr>",
            '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGEREF </w:instrText></w:r><w:r><w:instrText> bookmark \\h </w:instrText></w:r><w:r><w:instrText> \\* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p><w:sectPr>',
        )
        path = make_docx(Path(self.temp.name) / "fields.docx", split_field, MINIMAL_STYLES, MINIMAL_THEME)
        fields = DocxPackage(path).inventory()["fields"]

        self.assertEqual(fields, {"PAGE": 1, "PAGEREF": 1, "REF": 1})


if __name__ == "__main__":
    unittest.main()
