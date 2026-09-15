import sys
import tempfile
import unittest
import hashlib
from pathlib import Path
from zipfile import ZipFile

from lxml import etree


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


class NativeCommentWriterTests(unittest.TestCase):
    def test_groups_findings_by_paragraph_and_keeps_visible_text_unchanged(self):
        from comment_writer import write_comments
        from review_rules import build_review
        from tests.test_rule_engine import build_content_docx

        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            source = build_content_docx(
                temp / "论文.docx", invalid_body=True, add_spacing_issues=True
            )
            output = temp / "论文_格式审查批注版.docx"
            review = build_review(source)

            manifest = write_comments(
                source,
                review["findings"],
                output,
                created_at="2026-09-14T00:00:00Z",
            )

            with ZipFile(source) as before, ZipFile(output) as after:
                before_document = etree.fromstring(before.read("word/document.xml"))
                after_document = etree.fromstring(after.read("word/document.xml"))
                before_text = "".join(before_document.xpath(".//w:t/text()", namespaces={"w": W_NS}))
                after_text = "".join(after_document.xpath(".//w:t/text()", namespaces={"w": W_NS}))
                comments = etree.fromstring(after.read("word/comments.xml"))
                rels = etree.fromstring(after.read("word/_rels/document.xml.rels"))
                content_types = etree.fromstring(after.read("[Content_Types].xml"))

            self.assertEqual(before_text, after_text)
            comment_count = len(comments.xpath("./w:comment", namespaces={"w": W_NS}))
            self.assertGreaterEqual(comment_count, 2)
            comment_texts = [
                "".join(node.xpath(".//w:t/text()", namespaces={"w": W_NS}))
                for node in comments.xpath("./w:comment", namespaces={"w": W_NS})
            ]
            self.assertTrue(any("【格式审查：FONT-001】" in text for text in comment_texts))
            self.assertTrue(any("【格式审查：PARA-002】" in text for text in comment_texts))
            self.assertTrue(all("问题：" in text and "严重程度：" in text for text in comment_texts))
            self.assertEqual(
                len(rels.xpath("./r:Relationship[contains(@Type, '/comments')]", namespaces={"r": R_NS})),
                1,
            )
            self.assertEqual(
                len(content_types.xpath("./ct:Override[@PartName='/word/comments.xml']", namespaces={"ct": CT_NS})),
                1,
            )
            starts = after_document.xpath(".//w:commentRangeStart", namespaces={"w": W_NS})
            ends = after_document.xpath(".//w:commentRangeEnd", namespaces={"w": W_NS})
            refs = after_document.xpath(".//w:commentReference", namespaces={"w": W_NS})
            self.assertEqual(len(starts), comment_count)
            self.assertEqual(len(ends), comment_count)
            self.assertEqual(len(refs), comment_count)
            self.assertEqual(
                {node.get(f"{{{W_NS}}}id") for node in starts},
                {node.get(f"{{{W_NS}}}id") for node in refs},
            )

            by_rule = {item["rule_id"]: item for item in manifest["findings"]}
            self.assertTrue(by_rule["FONT-001"]["commented"])
            self.assertEqual(by_rule["FONT-001"]["comment_id"], by_rule["PARA-001"]["comment_id"])
            self.assertFalse(by_rule["PAGE-004"]["commented"])
            self.assertEqual(by_rule["PAGE-004"]["skip_reason"], "manual_review")

    def test_existing_comment_ids_are_preserved_and_ignore_unrelated_word_ids(self):
        from comment_writer import write_comments
        from tests.helpers import MINIMAL_STYLES, MINIMAL_THEME, make_docx

        document_xml = f'''<w:document xmlns:w="{W_NS}"><w:body>
          <w:p><w:bookmarkStart w:id="99" w:name="keep"/><w:r><w:t>待检查文字</w:t></w:r><w:bookmarkEnd w:id="99"/></w:p>
          <w:p><w:commentRangeStart w:id="5"/><w:r><w:t>原批注文字</w:t></w:r><w:commentRangeEnd w:id="5"/><w:r><w:commentReference w:id="5"/></w:r></w:p>
          <w:sectPr/>
        </w:body></w:document>'''
        comments_xml = f'''<w:comments xmlns:w="{W_NS}"><w:comment w:id="5" w:author="原作者"><w:p><w:r><w:t>原有批注</w:t></w:r></w:p></w:comment></w:comments>'''
        finding = {
            "finding_id": "FONT-001-0001",
            "rule_id": "FONT-001",
            "severity": "ERROR",
            "problem": "字体不符",
            "actual": "微软雅黑",
            "expected": "宋体",
            "source": "模板正文要求",
            "suggestion": "改用宋体",
            "anchor": {"part_name": "word/document.xml", "paragraph_index": 0, "snippet": "待检查文字", "run_start": 0, "run_end": 0},
        }

        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            source = make_docx(
                temp / "existing.docx",
                document_xml,
                MINIMAL_STYLES,
                MINIMAL_THEME,
                {"word/comments.xml": comments_xml},
            )
            output = temp / "result.docx"
            manifest = write_comments(source, [finding], output, created_at="2026-09-14T00:00:00Z")
            with ZipFile(output) as package:
                comments = etree.fromstring(package.read("word/comments.xml"))

        ids = [node.get(f"{{{W_NS}}}id") for node in comments.xpath("./w:comment", namespaces={"w": W_NS})]
        texts = ["".join(node.xpath(".//w:t/text()", namespaces={"w": W_NS})) for node in comments.xpath("./w:comment", namespaces={"w": W_NS})]
        self.assertEqual(ids, ["5", "6"])
        self.assertEqual(texts[0], "原有批注")
        self.assertEqual(manifest["findings"][0]["comment_id"], 6)
        self.assertEqual(manifest["findings"][0]["anchor_mode"], "run_range")

    def test_unmodified_package_parts_remain_byte_identical(self):
        from comment_writer import write_comments
        from review_rules import build_review
        from tests.test_rule_engine import build_content_docx

        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            source = build_content_docx(temp / "source.docx", invalid_body=True)
            output = temp / "commented.docx"
            write_comments(source, build_review(source)["findings"], output)

            with ZipFile(source) as before, ZipFile(output) as after:
                before_names = set(before.namelist())
                after_names = set(after.namelist())
                allowed = {
                    "word/document.xml",
                    "word/comments.xml",
                    "word/_rels/document.xml.rels",
                    "[Content_Types].xml",
                }
                for name in sorted(before_names - allowed):
                    self.assertEqual(
                        hashlib.sha256(before.read(name)).digest(),
                        hashlib.sha256(after.read(name)).digest(),
                        name,
                    )

            self.assertEqual(after_names - before_names, {"word/comments.xml"})

    def test_uncommentable_findings_copy_document_without_comment_plumbing(self):
        from comment_writer import write_comments
        from tests.test_rule_engine import build_content_docx

        findings = [
            {
                "finding_id": "PAGE-004-0001",
                "rule_id": "PAGE-004",
                "severity": "MANUAL_REVIEW",
                "anchor": None,
            },
            {
                "finding_id": "HF-001-0002",
                "rule_id": "HF-001",
                "severity": "ERROR",
                "anchor": {"part_name": "word/header1.xml", "paragraph_index": 0},
            },
        ]
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            source = build_content_docx(temp / "source.docx")
            output = temp / "copy.docx"
            manifest = write_comments(source, findings, output)
            with ZipFile(source) as before, ZipFile(output) as after:
                self.assertEqual(before.namelist(), after.namelist())
                for name in before.namelist():
                    self.assertEqual(before.read(name), after.read(name), name)

        self.assertEqual(manifest["comment_count"], 0)
        self.assertEqual(manifest["modified_parts"], [])
        self.assertEqual(manifest["findings"][0]["skip_reason"], "manual_review")
        self.assertEqual(manifest["findings"][1]["skip_reason"], "unsupported_story_part")

    def test_snippet_mismatch_is_not_forced_onto_wrong_paragraph(self):
        from comment_writer import write_comments
        from tests.test_rule_engine import build_content_docx

        finding = {
            "finding_id": "FONT-001-0001",
            "rule_id": "FONT-001",
            "severity": "ERROR",
            "anchor": {
                "part_name": "word/document.xml",
                "paragraph_index": 1,
                "snippet": "这段文字并不存在",
            },
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            source = build_content_docx(temp / "source.docx")
            output = temp / "result.docx"
            manifest = write_comments(source, [finding], output)

        self.assertFalse(manifest["findings"][0]["commented"])
        self.assertEqual(manifest["findings"][0]["skip_reason"], "snippet_mismatch")

    def test_field_instructions_use_the_same_anchor_text_model_as_phase4(self):
        from comment_writer import write_comments
        from tests.helpers import MINIMAL_STYLES, MINIMAL_THEME, make_docx

        document_xml = f'''<w:document xmlns:w="{W_NS}"><w:body>
          <w:p><w:r><w:t>见</w:t></w:r><w:r><w:instrText> REF _Ref123 \\h </w:instrText></w:r><w:r><w:t>图1.1</w:t></w:r></w:p>
          <w:sectPr/>
        </w:body></w:document>'''
        finding = {
            "finding_id": "PARA-002-0001",
            "rule_id": "PARA-002",
            "severity": "WARNING",
            "problem": "空段",
            "actual": "存在",
            "expected": "无多余空段",
            "source": "模板",
            "suggestion": "检查",
            "anchor": {
                "part_name": "word/document.xml",
                "paragraph_index": 0,
                "snippet": "见 REF _Ref123 \\h 图1.1",
            },
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            source = make_docx(temp / "field.docx", document_xml, MINIMAL_STYLES, MINIMAL_THEME)
            output = temp / "field-commented.docx"
            manifest = write_comments(source, [finding], output)

        self.assertTrue(manifest["findings"][0]["commented"])
        self.assertIsNone(manifest["findings"][0]["skip_reason"])

    def test_source_file_cannot_be_overwritten(self):
        from comment_writer import write_comments
        from tests.test_rule_engine import build_content_docx

        with tempfile.TemporaryDirectory() as temp_dir:
            source = build_content_docx(Path(temp_dir) / "source.docx")
            original = source.read_bytes()
            with self.assertRaisesRegex(ValueError, "不能覆盖"):
                write_comments(source, [], source)
            self.assertEqual(source.read_bytes(), original)

    def test_structural_inventory_is_unchanged_except_for_new_comments(self):
        from comment_writer import write_comments
        from docx_model import DocxPackage
        from tests.helpers import MINIMAL_STYLES, MINIMAL_THEME, make_docx

        document_xml = f'''<w:document xmlns:w="{W_NS}"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
          xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
          <w:body>
            <w:p><w:bookmarkStart w:id="8" w:name="anchor"/><w:r><w:t>检查这一段</w:t></w:r><w:bookmarkEnd w:id="8"/></w:p>
            <w:p><w:hyperlink r:id="rIdLink"><w:r><w:t>链接</w:t></w:r></w:hyperlink></w:p>
            <w:p><w:r><w:fldChar w:fldCharType="begin"/><w:instrText> PAGE </w:instrText><w:fldChar w:fldCharType="end"/></w:r></w:p>
            <w:p><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></w:p>
            <w:p><w:r><w:drawing><wp:inline/></w:drawing></w:r></w:p>
            <w:sdt><w:sdtContent><w:p><w:r><w:t>内容控件</w:t></w:r></w:p></w:sdtContent></w:sdt>
            <w:sectPr><w:pgSz w:w="11907" w:h="16840"/></w:sectPr>
          </w:body>
        </w:document>'''
        finding = {
            "finding_id": "PARA-001-0001",
            "rule_id": "PARA-001",
            "severity": "ERROR",
            "problem": "段落格式不符",
            "actual": "左对齐",
            "expected": "两端对齐",
            "source": "模板",
            "suggestion": "调整段落",
            "anchor": {"part_name": "word/document.xml", "paragraph_index": 0, "snippet": "检查这一段"},
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            source = make_docx(temp / "objects.docx", document_xml, MINIMAL_STYLES, MINIMAL_THEME)
            output = temp / "objects-commented.docx"
            before = DocxPackage(source).inventory()
            write_comments(source, [finding], output)
            after = DocxPackage(output).inventory()

        preserved_keys = {
            "paragraphs_body",
            "tables",
            "sections",
            "content_controls",
            "tracked_insertions",
            "tracked_deletions",
            "drawings_inline",
            "drawings_anchor",
            "vml_shapes",
            "omml_objects",
            "fields",
            "bookmarks",
            "hyperlinks",
            "footnotes",
            "endnotes",
            "hidden_runs",
            "manual_page_breaks",
            "manual_column_breaks",
            "paragraph_section_breaks",
            "ole_objects",
            "text_boxes",
            "smartart_relationship_markers",
        }
        for key in preserved_keys:
            self.assertEqual(after[key], before[key], key)
        self.assertGreater(after["paragraphs_all"], before["paragraphs_all"])
        self.assertEqual(after["comments"], before["comments"] + 1)
        self.assertEqual(after["parts"], before["parts"] + 1)


if __name__ == "__main__":
    unittest.main()
