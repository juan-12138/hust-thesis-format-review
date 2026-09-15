import hashlib
import sys
import unittest
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
TEMPLATE = ROOT / "references" / "华中科技大学硕士学位论文参考模板.docx"


class TemplateIntegrationTests(unittest.TestCase):
    def test_retained_template_parses_with_known_structural_evidence_and_is_read_only(self):
        from docx_model import DocxPackage
        from effective_style import EffectiveStyleResolver
        from structure_classifier import StructureClassifier

        before = hashlib.sha256(TEMPLATE.read_bytes()).hexdigest().upper()
        package = DocxPackage(TEMPLATE)
        paragraphs = list(package.iter_paragraphs())
        roles = Counter(item.classification.role for item in StructureClassifier(package, EffectiveStyleResolver(package)).classify_document(paragraphs))
        inventory = package.inventory()

        self.assertEqual(before, "43DB10A8AFB05D346F84B9A7C08A7E811667D27660CC0018B7C047115674B9C1")
        # This low-level parser intentionally includes paragraphs nested in
        # tables/content controls/revisions; python-docx's Document.paragraphs
        # view used during Phase 1 reported only 344 top-level paragraphs.
        self.assertEqual(len(paragraphs), 416)
        sections = package.sections()
        self.assertEqual(len(sections), 3)
        self.assertFalse(sections[0]["header_linked_to_previous"])
        self.assertFalse(sections[1]["header_linked_to_previous"])
        self.assertTrue(sections[2]["header_linked_to_previous"])
        self.assertTrue(sections[2]["footer_linked_to_previous"])
        self.assertEqual(inventory["content_controls"], 20)
        self.assertEqual(inventory["drawings_inline"], 2)
        self.assertEqual(inventory["drawings_anchor"], 3)
        self.assertEqual(inventory["fields"]["PAGEREF"], 32)
        self.assertEqual(inventory["fields"]["PAGE"], 2)
        self.assertEqual(roles["heading_2"], 26)
        self.assertEqual(roles["heading_3"], 9)
        self.assertEqual(hashlib.sha256(TEMPLATE.read_bytes()).hexdigest().upper(), before)


if __name__ == "__main__":
    unittest.main()
