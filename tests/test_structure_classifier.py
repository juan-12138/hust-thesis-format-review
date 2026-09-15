import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from .helpers import MINIMAL_STYLES, MINIMAL_THEME, make_docx


DOCUMENT = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:body>
  <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>摘  要</w:t></w:r></w:p>
  <w:p><w:r><w:t>摘要正文</w:t></w:r></w:p>
  <w:p><w:r><w:t>关键词：结构；格式；论文</w:t></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>目  录</w:t></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="toc 1"/></w:pPr><w:r><w:t>1 绪论</w:t></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>1 绪论</w:t></w:r></w:p>
  <w:p><w:r><w:t>正文内容</w:t></w:r></w:p>
  <w:p><w:r><w:t>图2-1 结构示意图</w:t></w:r></w:p>
  <w:p><w:r><w:t>表2-1 参数</w:t></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>参考文献</w:t></w:r></w:p>
  <w:p><w:r><w:t>[1] 作者. 题名</w:t></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>附录1 研究成果</w:t></w:r></w:p>
  <w:sectPr/>
 </w:body>
</w:document>"""


class StructureClassifierTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.path = make_docx(Path(self.temp.name) / "roles.docx", DOCUMENT, MINIMAL_STYLES, MINIMAL_THEME)

    def tearDown(self):
        self.temp.cleanup()

    def test_classifies_structural_roles_with_evidence(self):
        from docx_model import DocxPackage
        from effective_style import EffectiveStyleResolver
        from structure_classifier import StructureClassifier

        package = DocxPackage(self.path)
        classifier = StructureClassifier(package, EffectiveStyleResolver(package))
        results = classifier.classify_document()

        roles = [item.classification.role for item in results]
        self.assertEqual(
            roles,
            [
                "abstract_title_cn",
                "abstract_body_cn",
                "keywords_cn",
                "toc_title",
                "toc_entry",
                "heading_1",
                "body",
                "figure_caption",
                "table_caption",
                "references_title",
                "reference_entry",
                "appendix_heading",
            ],
        )
        self.assertGreaterEqual(results[0].classification.confidence, 0.95)
        self.assertIn("标题文字", results[0].classification.evidence)

    def test_cli_inspection_composes_parser_styles_and_roles(self):
        from inspect_docx import build_inspection

        inspection = build_inspection(self.path, include_runs=True)

        self.assertEqual(inspection["inventory"]["paragraphs_body"], 12)
        self.assertEqual(inspection["role_counts"]["reference_entry"], 1)
        self.assertEqual(inspection["paragraphs"][5]["paragraph_format"]["alignment"], "center")
        self.assertEqual(inspection["paragraphs"][5]["runs"][0]["effective_format"]["fonts"]["eastAsia"], "黑体")
        self.assertFalse(inspection["capabilities"]["final_page_numbers_available"])
        self.assertFalse(inspection["settings"]["even_and_odd_headers"])
        self.assertEqual(inspection["numbering"]["numbering_instances"], 0)
        self.assertGreaterEqual(inspection["styles"]["paragraph_styles"], 3)
        self.assertIn("word/document.xml", inspection["relationships"])


if __name__ == "__main__":
    unittest.main()
