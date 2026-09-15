import hashlib
import json
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))


class Phase7WorkflowTests(unittest.TestCase):
    def test_one_step_review_creates_both_docx_outputs_and_manifest_without_touching_source(self):
        from review_thesis import run_review
        from tests.test_rule_engine import build_content_docx

        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            source = build_content_docx(temp / "学位论文.docx", invalid_body=True)
            output = temp / "指定批注版.docx"
            report_output = temp / "指定审查报告.docx"
            manifest_path = temp / "指定批注版.json"
            source_hash = hashlib.sha256(source.read_bytes()).hexdigest()

            result = run_review(
                source,
                output=output,
                report_output=report_output,
                manifest_path=manifest_path,
            )

            self.assertTrue(output.is_file())
            self.assertTrue(report_output.is_file())
            self.assertEqual(hashlib.sha256(source.read_bytes()).hexdigest(), source_hash)
            self.assertEqual(result["schema_version"], "1.0.0-phase7")
            self.assertTrue(result["capabilities"]["writes_word_comments"])
            self.assertTrue(result["capabilities"]["writes_word_report"])
            self.assertGreater(result["review"]["statistics"]["total_findings"], 0)
            self.assertGreater(result["comments"]["commented_finding_count"], 0)
            self.assertEqual(result["report"]["finding_count"], result["review"]["statistics"]["total_findings"])
            self.assertTrue(result["package_verification"]["passed"])
            pkg_result = next(item for item in result["review"]["rule_results"] if item["rule_id"] == "PKG-002")
            self.assertEqual(pkg_result["status"], "PASS")
            self.assertEqual(Path(result["report"]["output"]), report_output.resolve())
            self.assertEqual(json.loads(manifest_path.read_text(encoding="utf-8")), result)

    def test_default_output_uses_required_suffix(self):
        from review_thesis import run_review
        from tests.test_rule_engine import build_content_docx

        with tempfile.TemporaryDirectory() as temp_dir:
            source = build_content_docx(Path(temp_dir) / "论文初稿.docx", invalid_body=True)
            result = run_review(source)
            expected_comments = source.with_name("论文初稿_格式审查批注版.docx")
            expected_report = source.with_name("论文初稿_格式审查报告.docx")

            self.assertEqual(Path(result["comments"]["output"]), expected_comments.resolve())
            self.assertEqual(Path(result["report"]["output"]), expected_report.resolve())
            self.assertTrue(expected_comments.is_file())
            self.assertTrue(expected_report.is_file())

    def test_comment_and_report_outputs_must_be_different_files(self):
        from review_thesis import run_review
        from tests.test_rule_engine import build_content_docx

        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            source = build_content_docx(temp / "论文.docx", invalid_body=True)
            same_output = temp / "冲突输出.docx"
            with self.assertRaisesRegex(ValueError, "不同"):
                run_review(source, output=same_output, report_output=same_output)
            self.assertFalse(same_output.exists())


if __name__ == "__main__":
    unittest.main()
