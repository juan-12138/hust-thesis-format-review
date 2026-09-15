import sys
import tempfile
import unittest
import hashlib
import json
import shutil
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
RULES_JSON = ROOT / "rules" / "hust_master_thesis_rules.json"
TEMPLATE = ROOT / "references" / "华中科技大学硕士学位论文参考模板.docx"


def build_page_docx(path: Path, *, width=11907, height=16840, orient=None, top=2549, bottom=1584, left=1584, right=1584, header=851, footer=964):
    from tests.helpers import MINIMAL_STYLES, MINIMAL_THEME, make_docx

    orient_attr = f' w:orient="{orient}"' if orient else ""
    document_xml = f'''<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>封面</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="{width}" w:h="{height}"{orient_attr}/>
      <w:pgMar w:top="{top}" w:right="{right}" w:bottom="{bottom}" w:left="{left}" w:header="{header}" w:footer="{footer}" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>'''
    return make_docx(path, document_xml, MINIMAL_STYLES, MINIMAL_THEME)


def build_content_docx(path: Path, *, invalid_body=False, add_spacing_issues=False, revised_red=False):
    from tests.helpers import MINIMAL_THEME, make_docx

    styles_xml = '''<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="24"/><w:color w:val="000000"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:jc w:val="both"/><w:spacing w:line="360" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/><w:jc w:val="center"/></w:pPr><w:rPr><w:rFonts w:eastAsia="黑体" w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="32"/><w:b/></w:rPr></w:style>
</w:styles>'''
    if invalid_body:
        body_ppr = '<w:pPr><w:jc w:val="left"/><w:spacing w:line="240" w:lineRule="auto"/></w:pPr>'
        body_rpr = '<w:rPr><w:rFonts w:eastAsia="微软雅黑" w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/><w:color w:val="FF0000"/></w:rPr>'
    else:
        body_ppr = ""
        body_rpr = ""
    spacing = ""
    if add_spacing_issues:
        spacing = '<w:p/><w:p/><w:p><w:r><w:br w:type="page"/><w:t>分页后文字</w:t></w:r></w:p>'
    revision = ""
    if revised_red:
        revision = '<w:p><w:ins w:id="1"><w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>修订文字</w:t></w:r></w:ins></w:p>'
    document_xml = f'''<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>1 绪论</w:t></w:r></w:p>
    <w:p>{body_ppr}<w:r>{body_rpr}<w:t>中文ABC123</w:t></w:r><w:r>{body_rpr}<w:t>继续DEF</w:t></w:r></w:p>
    <w:p><w:hyperlink r:id="rIdLink"><w:r><w:rPr><w:color w:val="0563C1"/></w:rPr><w:t>https://example.com</w:t></w:r></w:hyperlink></w:p>
    {revision}
    {spacing}
    <w:sectPr><w:pgSz w:w="11907" w:h="16840"/><w:pgMar w:top="2549" w:right="1584" w:bottom="1584" w:left="1584" w:header="851" w:footer="964"/></w:sectPr>
  </w:body>
</w:document>'''
    return make_docx(path, document_xml, styles_xml, MINIMAL_THEME)


def build_heading_docx(path: Path, *, invalid=False):
    from tests.helpers import MINIMAL_THEME, make_docx

    heading1 = '<w:rFonts w:eastAsia="黑体"/><w:sz w:val="32"/><w:b/>'
    heading2 = '<w:rFonts w:eastAsia="黑体"/><w:sz w:val="28"/><w:b/>'
    heading3 = '<w:rFonts w:eastAsia="黑体"/><w:sz w:val="24"/><w:b/>'
    if invalid:
        heading1 = '<w:rFonts w:eastAsia="宋体"/><w:sz w:val="28"/><w:b w:val="0"/>'
        heading2 = '<w:rFonts w:eastAsia="宋体"/><w:sz w:val="24"/><w:b w:val="0"/>'
        heading3 = '<w:rFonts w:eastAsia="宋体"/><w:sz w:val="20"/><w:b w:val="0"/>'
    h1_alignment = "left" if invalid else "center"
    styles_xml = f'''<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="24"/><w:color w:val="000000"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:jc w:val="both"/><w:spacing w:line="360" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/><w:jc w:val="{h1_alignment}"/></w:pPr><w:rPr>{heading1}</w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:rPr>{heading2}</w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:pPr><w:outlineLvl w:val="2"/></w:pPr><w:rPr>{heading3}</w:rPr></w:style>
</w:styles>'''
    if invalid:
        headings = '''
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>1 绪论</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t>1.1.1 具体方法</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>1.1 研究内容</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>3 试验结果</w:t></w:r></w:p>'''
    else:
        headings = '''
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>1 绪论</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>1.1 干栏建筑形态演化</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t>1.1.1 木构连接特征</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>2 试验结果</w:t></w:r></w:p>'''
    document_xml = f'''<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>{headings}
    <w:sectPr><w:pgSz w:w="11907" w:h="16840"/><w:pgMar w:top="2549" w:right="1584" w:bottom="1584" w:left="1584" w:header="851" w:footer="964"/></w:sectPr>
  </w:body>
</w:document>'''
    return make_docx(path, document_xml, styles_xml, MINIMAL_THEME)


def build_header_footer_docx(path: Path, *, invalid=False):
    from tests.helpers import MINIMAL_STYLES, MINIMAL_THEME, make_docx

    if invalid:
        header_xml = '''<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:rPr><w:rFonts w:eastAsia="宋体"/><w:sz w:val="20"/><w:color w:val="0563C1"/></w:rPr><w:t>错误页眉</w:t></w:r></w:p></w:hdr>'''
        footer_xml = '''<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:fldSimple w:instr=" PAGE "><w:r><w:t>3</w:t></w:r></w:fldSimple></w:p></w:ftr>'''
        page_num = '<w:pgNumType w:fmt="upperRoman" w:start="3"/>'
    else:
        header_xml = '''<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml"><w:p><w:r><w:rPr><w:rFonts w:eastAsia="楷体"/><w:sz w:val="36"/><w:color w:val="FF0000"/></w:rPr><w:t>华中科技大学硕士学位论文</w:t></w:r><v:line strokecolor="#FF0000"/><v:line strokecolor="red"/></w:p></w:hdr>'''
        footer_xml = '''<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>'''
        page_num = '<w:pgNumType w:fmt="decimal" w:start="1"/>'
    document_xml = f'''<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>1 绪论</w:t></w:r></w:p>
    <w:p><w:r><w:t>正文ABC</w:t></w:r></w:p>
    <w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/><w:pgSz w:w="11907" w:h="16840"/><w:pgMar w:top="2549" w:right="1584" w:bottom="1584" w:left="1584" w:header="851" w:footer="964"/>{page_num}</w:sectPr>
  </w:body>
</w:document>'''
    return make_docx(path, document_xml, MINIMAL_STYLES, MINIMAL_THEME, {"word/header1.xml": header_xml, "word/footer1.xml": footer_xml})


def build_unknown_format_docx(path: Path):
    from tests.helpers import MINIMAL_THEME, make_docx

    styles_xml = '''<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/><w:jc w:val="center"/></w:pPr></w:style></w:styles>'''
    document_xml = '''<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>1 绪论</w:t></w:r></w:p><w:p><w:r><w:t>中文ABC</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11907" w:h="16840"/><w:pgMar w:top="2549" w:right="1584" w:bottom="1584" w:left="1584" w:header="851" w:footer="964"/></w:sectPr></w:body></w:document>'''
    return make_docx(path, document_xml, styles_xml, MINIMAL_THEME)


def build_implicit_black_docx(path: Path):
    from tests.helpers import MINIMAL_THEME, make_docx

    styles_xml = '''<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:jc w:val="both"/><w:spacing w:line="360" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/><w:jc w:val="center"/></w:pPr><w:rPr><w:rFonts w:eastAsia="黑体"/><w:sz w:val="32"/><w:b/></w:rPr></w:style></w:styles>'''
    document_xml = '''<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>1 绪论</w:t></w:r></w:p><w:p><w:r><w:t>中文ABC</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11907" w:h="16840"/><w:pgMar w:top="2549" w:right="1584" w:bottom="1584" w:left="1584" w:header="851" w:footer="964"/></w:sectPr></w:body></w:document>'''
    return make_docx(path, document_xml, styles_xml, MINIMAL_THEME)


class RuleConfigTests(unittest.TestCase):
    def test_runtime_rule_config_contains_all_rules_and_phase4_subset(self):
        from rule_engine import PHASE4_RULE_IDS, load_rule_config

        config = load_rule_config(RULES_JSON)
        ids = [rule["rule_id"] for rule in config["rules"]]

        self.assertEqual(len(ids), 76)
        self.assertEqual(len(ids), len(set(ids)))
        self.assertEqual(len(PHASE4_RULE_IDS), 19)
        self.assertTrue(set(PHASE4_RULE_IDS).issubset(ids))

    def test_runtime_rule_mirror_rejects_source_yaml_drift(self):
        from rule_engine import load_rule_config

        with tempfile.TemporaryDirectory() as temp_dir:
            copied_json = Path(temp_dir) / RULES_JSON.name
            copied_yaml = copied_json.with_suffix(".yaml")
            shutil.copyfile(RULES_JSON, copied_json)
            copied_yaml.write_text("changed: true\n", encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "镜像"):
                load_rule_config(copied_json)


class PageRuleTests(unittest.TestCase):
    def _review(self, path: Path):
        from docx_model import DocxPackage
        from rule_engine import RuleEngine

        return RuleEngine(DocxPackage(path), rules_path=RULES_JSON).review()

    def test_valid_page_geometry_passes_deterministic_rules(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = build_page_docx(Path(temp_dir) / "valid.docx")
            review = self._review(path)

        statuses = {result.rule_id: result.status for result in review.rule_results}
        self.assertEqual(statuses["PAGE-001"], "PASS")
        self.assertEqual(statuses["PAGE-002"], "PASS")
        self.assertEqual(statuses["PAGE-003"], "PASS")
        self.assertEqual(statuses["PAGE-004"], "MANUAL_REVIEW")

    def test_invalid_page_geometry_reports_each_rule_at_the_section(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = build_page_docx(
                Path(temp_dir) / "invalid.docx",
                width=12240,
                height=15840,
                orient="landscape",
                top=1440,
                bottom=1440,
                left=1800,
                right=1800,
                header=720,
                footer=720,
            )
            review = self._review(path)

        findings = {finding.rule_id: finding for finding in review.findings}
        self.assertEqual(findings["PAGE-001"].severity, "ERROR")
        self.assertEqual(findings["PAGE-001"].location["section_index"], 0)
        self.assertIn("12240", findings["PAGE-001"].actual)
        self.assertEqual(findings["PAGE-002"].severity, "ERROR")
        self.assertEqual(findings["PAGE-003"].severity, "WARNING")
        self.assertEqual(findings["PAGE-004"].severity, "MANUAL_REVIEW")

    def test_missing_page_properties_are_manual_review_not_definite_errors(self):
        from tests.helpers import MINIMAL_STYLES, MINIMAL_THEME, make_docx

        document_xml = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sectPr/></w:body></w:document>'
        with tempfile.TemporaryDirectory() as temp_dir:
            path = make_docx(Path(temp_dir) / "missing-page.docx", document_xml, MINIMAL_STYLES, MINIMAL_THEME)
            review = self._review(path)

        results = {result.rule_id: result for result in review.rule_results}
        for rule_id in ("PAGE-001", "PAGE-002", "PAGE-003"):
            self.assertEqual(results[rule_id].status, "MANUAL_REVIEW", rule_id)


class BodyRuleTests(unittest.TestCase):
    def _review(self, path: Path):
        from docx_model import DocxPackage
        from rule_engine import RuleEngine

        return RuleEngine(DocxPackage(path), rules_path=RULES_JSON).review()

    def test_wrong_body_font_size_color_alignment_and_spacing_are_reported_once_per_paragraph(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = build_content_docx(Path(temp_dir) / "invalid-body.docx", invalid_body=True)
            review = self._review(path)

        by_rule = {}
        for finding in review.findings:
            by_rule.setdefault(finding.rule_id, []).append(finding)
        self.assertEqual(len(by_rule["FONT-001"]), 1)
        self.assertEqual(len(by_rule["FONT-002"]), 1)
        self.assertEqual(len(by_rule["FONT-003"]), 1)
        self.assertEqual(len(by_rule["PARA-001"]), 1)
        self.assertIn("微软雅黑", by_rule["FONT-001"][0].actual)
        self.assertIn("FF0000", by_rule["FONT-002"][0].actual)
        self.assertEqual(by_rule["FONT-001"][0].anchor["paragraph_index"], 1)

    def test_valid_body_and_colored_hyperlink_pass_body_rules(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = build_content_docx(Path(temp_dir) / "valid-body.docx")
            review = self._review(path)

        statuses = {result.rule_id: result.status for result in review.rule_results}
        self.assertEqual(statuses["FONT-001"], "PASS")
        self.assertEqual(statuses["FONT-002"], "PASS")
        self.assertEqual(statuses["FONT-003"], "PASS")
        self.assertEqual(statuses["PARA-001"], "PASS")

    def test_colored_text_inside_tracked_revision_is_excluded_from_ordinary_body_color(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = build_content_docx(Path(temp_dir) / "revision.docx", revised_red=True)
            review = self._review(path)

        color_findings = [finding for finding in review.findings if finding.rule_id == "FONT-002"]
        self.assertEqual(color_findings, [])

    def test_spacing_and_manual_break_candidates_are_located_without_claiming_rendered_effect(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = build_content_docx(Path(temp_dir) / "spacing.docx", add_spacing_issues=True)
            review = self._review(path)

        spacing_findings = [finding for finding in review.findings if finding.rule_id == "PARA-002"]
        break_findings = [finding for finding in review.findings if finding.rule_id == "PARA-003"]
        self.assertEqual(len(spacing_findings), 1)
        self.assertEqual(spacing_findings[0].severity, "WARNING")
        self.assertEqual(len(break_findings), 1)
        self.assertEqual(break_findings[0].severity, "MANUAL_REVIEW")
        self.assertTrue(break_findings[0].evidence["render_required"])

    def test_unresolved_effective_format_is_manual_review_not_a_definite_error(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            review = self._review(build_unknown_format_docx(Path(temp_dir) / "unknown.docx"))

        relevant = {finding.rule_id: finding for finding in review.findings if finding.rule_id in {"FONT-001", "FONT-003", "PARA-001"}}
        self.assertEqual(relevant["FONT-001"].severity, "MANUAL_REVIEW")
        self.assertEqual(relevant["FONT-003"].severity, "MANUAL_REVIEW")
        self.assertEqual(relevant["PARA-001"].severity, "MANUAL_REVIEW")

    def test_omitted_color_property_uses_word_automatic_black_default(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            review = self._review(build_implicit_black_docx(Path(temp_dir) / "implicit-black.docx"))

        result = next(result for result in review.rule_results if result.rule_id == "FONT-002")
        self.assertEqual(result.status, "PASS")
        self.assertEqual(result.finding_count, 0)


class HeadingRuleTests(unittest.TestCase):
    def _review(self, path: Path):
        from docx_model import DocxPackage
        from rule_engine import RuleEngine

        return RuleEngine(DocxPackage(path), rules_path=RULES_JSON).review()

    def test_valid_heading_levels_and_numbering_pass(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            review = self._review(build_heading_docx(Path(temp_dir) / "valid-headings.docx"))

        statuses = {result.rule_id: result.status for result in review.rule_results}
        checked = {result.rule_id: result.evaluated_count for result in review.rule_results}
        for rule_id in ("HEAD-001", "HEAD-002", "HEAD-003", "HEAD-004", "HEAD-005"):
            self.assertEqual(statuses[rule_id], "PASS", rule_id)
            self.assertGreater(checked[rule_id], 0, rule_id)

    def test_heading_format_hierarchy_sequence_and_generic_title_are_reported(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            review = self._review(build_heading_docx(Path(temp_dir) / "invalid-headings.docx", invalid=True))

        by_rule = {}
        for finding in review.findings:
            by_rule.setdefault(finding.rule_id, []).append(finding)
        self.assertTrue(by_rule["HEAD-001"])
        self.assertTrue(by_rule["HEAD-002"])
        self.assertTrue(by_rule["HEAD-003"])
        self.assertTrue(by_rule["HEAD-004"])
        self.assertTrue(any("跳级" in finding.problem for finding in by_rule["HEAD-004"]))
        self.assertTrue(any("不连续" in finding.problem for finding in by_rule["HEAD-004"]))
        self.assertEqual(by_rule["HEAD-005"][0].severity, "WARNING")


class Phase4CliTests(unittest.TestCase):
    def test_build_review_is_serializable_traceable_and_read_only(self):
        from review_rules import build_review

        with tempfile.TemporaryDirectory() as temp_dir:
            path = build_content_docx(Path(temp_dir) / "cli.docx", invalid_body=True, add_spacing_issues=True)
            before = hashlib.sha256(path.read_bytes()).hexdigest().upper()
            payload = build_review(path, RULES_JSON)
            encoded = json.dumps(payload, ensure_ascii=False)
            after = hashlib.sha256(path.read_bytes()).hexdigest().upper()

        self.assertEqual(payload["schema_version"], "1.0.0-phase7")
        self.assertEqual(payload["source"]["sha256"], before)
        self.assertEqual(before, after)
        self.assertEqual(len(payload["rule_results"]), 76)
        self.assertEqual(payload["statistics"]["total_findings"], len(payload["findings"]))
        self.assertIn("FONT-001", payload["coverage"]["implemented_rule_ids"])
        self.assertTrue(payload["capabilities"]["header_footer_checks"])
        self.assertFalse(payload["capabilities"]["writes_word_comments"])
        self.assertIn("微软雅黑", encoded)

    def test_retained_template_exercises_main_matter_and_header_footer_checks(self):
        from docx_model import DocxPackage
        from rule_engine import RuleEngine

        review = RuleEngine(DocxPackage(TEMPLATE), rules_path=RULES_JSON).review()
        results = {result.rule_id: result for result in review.rule_results}

        self.assertGreater(results["HEAD-001"].evaluated_count, 0)
        self.assertGreater(results["FONT-001"].evaluated_count, 0)
        self.assertEqual(results["HF-001"].evaluated_count, 1)
        self.assertEqual(results["HF-003"].status, "PASS")
        self.assertEqual(results["HF-004"].status, "PASS")


class HeaderFooterRuleTests(unittest.TestCase):
    def _review(self, path: Path):
        from docx_model import DocxPackage
        from rule_engine import RuleEngine

        return RuleEngine(DocxPackage(path), rules_path=RULES_JSON).review()

    def test_valid_header_footer_and_main_page_number_system_pass(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            review = self._review(build_header_footer_docx(Path(temp_dir) / "valid-hf.docx"))

        results = {result.rule_id: result for result in review.rule_results}
        for rule_id in ("HF-001", "HF-002", "HF-003", "HF-004"):
            self.assertEqual(results[rule_id].status, "PASS", rule_id)
            self.assertGreater(results[rule_id].evaluated_count, 0, rule_id)

    def test_invalid_header_footer_and_page_number_section_are_reported(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            review = self._review(build_header_footer_docx(Path(temp_dir) / "invalid-hf.docx", invalid=True))

        by_rule = {}
        for finding in review.findings:
            by_rule.setdefault(finding.rule_id, []).append(finding)
        self.assertEqual(by_rule["HF-001"][0].severity, "ERROR")
        self.assertEqual(by_rule["HF-002"][0].severity, "WARNING")
        self.assertEqual(by_rule["HF-003"][0].severity, "ERROR")
        self.assertTrue(any("页码体系" in finding.problem for finding in by_rule["PAGE-004"]))


if __name__ == "__main__":
    unittest.main()
