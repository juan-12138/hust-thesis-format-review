import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from review_rules import build_review
from tests.helpers import MINIMAL_STYLES, MINIMAL_THEME, make_docx


ADVANCED_DOCUMENT = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
 <w:body>
  <w:p><w:r><w:t>摘要</w:t></w:r></w:p>
  <w:p><w:r><w:t>这是一个过短的摘要，我们提出一种方法。</w:t></w:r></w:p>
  <w:p><w:r><w:t>关键词：格式；审查。</w:t></w:r></w:p>
  <w:p><w:r><w:t>目  录 （此处为目录域）</w:t></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:t>1 绪论 1</w:t></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>1 绪论</w:t></w:r></w:p>
  <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>图1-1 错在图上方</w:t></w:r></w:p>
  <w:p><w:r><w:drawing><wp:inline><wp:extent cx="1000000" cy="1000000"/><a:graphic/></wp:inline></w:drawing></w:r></w:p>
  <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>表1-1 示例表</w:t></w:r></w:p>
  <w:tbl>
   <w:tblPr><w:jc w:val="center"/><w:tblBorders><w:top w:val="single"/><w:bottom w:val="single"/><w:insideV w:val="single"/></w:tblBorders></w:tblPr>
   <w:tr><w:tc><w:p><w:r><w:t>列1。</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>列2</w:t></w:r></w:p></w:tc></w:tr>
   <w:tr><w:tc><w:p><w:r><w:t>1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>2</w:t></w:r></w:p></w:tc></w:tr>
  </w:tbl>
  <w:p><m:oMathPara><m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath></m:oMathPara><w:r><w:t>2-1</w:t></w:r></w:p>
  <w:p><w:r><w:t>正文引用[1]和[3]。</w:t></w:r></w:p>
  <w:p><w:r><w:t>参考文献</w:t></w:r></w:p>
  <w:p><w:r><w:t>[1] 张三. 示例题名. 北京: 出版社, 2020.</w:t></w:r></w:p>
  <w:p><w:r><w:t>[3] J Smith. Example title. Journal, 2021, 1(1): 1-2.</w:t></w:r></w:p>
  <w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>
  <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="2041" w:right="1644" w:bottom="1786" w:left="1701" w:header="737" w:footer="737"/><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr>
 </w:body>
</w:document>"""


BAD_FOOTNOTE = """<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:footnote w:id="1"><w:p><w:pPr><w:ind w:firstLine="420"/></w:pPr><w:r><w:rPr><w:rFonts w:eastAsia="黑体"/><w:sz w:val="24"/></w:rPr><w:t>脚注内容</w:t></w:r></w:p></w:footnote>
</w:footnotes>"""


def _review(tmp_path: Path):
    path = make_docx(
        tmp_path / "advanced.docx",
        ADVANCED_DOCUMENT,
        MINIMAL_STYLES.replace("</w:styles>", '<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/></w:style></w:styles>'),
        MINIMAL_THEME,
        {"word/footnotes.xml": BAD_FOOTNOTE},
    )
    return build_review(path)


class AdvancedRuleEngineTests(unittest.TestCase):
    def setUp(self):
        self.tempdir = tempfile.TemporaryDirectory()
        self.review = _review(Path(self.tempdir.name))

    def tearDown(self):
        self.tempdir.cleanup()

    def test_phase7_returns_all_76_rule_results(self):
        ids = {item["rule_id"] for item in self.review["rule_results"]}
        self.assertEqual(len(ids), 76)
        self.assertEqual(self.review["coverage"]["implemented_rule_count"], 76)
        self.assertEqual(self.review["schema_version"], "1.0.0-phase7")

    def test_detects_manufactured_advanced_errors(self):
        by_rule = {finding["rule_id"]: finding for finding in self.review["findings"]}
        self.assertIn("上方", by_rule["FIG-002"]["actual"])
        self.assertIn("竖线", by_rule["TAB-001"]["actual"])
        self.assertIn("2-1", by_rule["EQ-002"]["actual"])
        self.assertIn("[1, 3]", by_rule["REF-001"]["actual"])
        self.assertIn("REF-003", by_rule)
        self.assertIn("12", by_rule["NOTE-001"]["actual"])

    def test_manual_boundaries_are_explicit(self):
        manual_ids = {
            finding["rule_id"]
            for finding in self.review["findings"]
            if finding["severity"] == "MANUAL_REVIEW"
        }
        self.assertTrue(
            {"ABS-005", "FIG-005", "TAB-006", "REF-011", "LANG-001", "RENDER-001", "RENDER-002", "RENDER-003"}
            <= manual_ids
        )


if __name__ == "__main__":
    unittest.main()
