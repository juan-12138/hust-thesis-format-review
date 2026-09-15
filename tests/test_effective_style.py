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
  <w:p>
   <w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="right"/><w:spacing w:line="480" w:lineRule="exact"/><w:ind w:left="240"/></w:pPr>
   <w:r><w:rPr><w:rStyle w:val="Emphasis"/></w:rPr><w:t>继承格式</w:t></w:r>
   <w:r><w:rPr><w:rStyle w:val="Emphasis"/><w:rFonts w:ascii="Arial"/><w:b w:val="0"/><w:sz w:val="28"/><w:color w:val="595959"/><w:u w:val="single"/><w:highlight w:val="yellow"/><w:spacing w:val="20"/><w:w w:val="90"/><w:position w:val="2"/></w:rPr><w:t>直接格式</w:t></w:r>
  </w:p>
  <w:sectPr/>
 </w:body>
</w:document>"""


class EffectiveStyleTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.path = make_docx(Path(self.temp.name) / "style.docx", DOCUMENT, MINIMAL_STYLES, MINIMAL_THEME)

    def tearDown(self):
        self.temp.cleanup()

    def test_resolves_defaults_based_on_paragraph_character_and_theme_layers(self):
        from docx_model import DocxPackage
        from effective_style import EffectiveStyleResolver

        package = DocxPackage(self.path)
        paragraph = next(package.iter_paragraphs())
        first_run = next(package.iter_runs(paragraph))
        fmt = EffectiveStyleResolver(package).get_effective_run_format(first_run, paragraph)

        self.assertEqual(fmt["fonts"]["ascii"], "Times New Roman")
        self.assertEqual(fmt["fonts"]["eastAsia"], "黑体")
        self.assertEqual(fmt["size_pt"], 16.0)
        self.assertTrue(fmt["bold"])
        self.assertTrue(fmt["italic"])
        self.assertEqual(fmt["color_hex"], "99B3CC")
        self.assertEqual(fmt["provenance"]["fonts.eastAsia"], "paragraph-style:Heading1")
        self.assertEqual(fmt["provenance"]["italic"], "character-style:Emphasis")

    def test_direct_run_formatting_overrides_inherited_values(self):
        from docx_model import DocxPackage
        from effective_style import EffectiveStyleResolver

        package = DocxPackage(self.path)
        paragraph = next(package.iter_paragraphs())
        run = list(package.iter_runs(paragraph))[1]
        fmt = EffectiveStyleResolver(package).get_effective_run_format(run, paragraph)

        self.assertEqual(fmt["fonts"]["ascii"], "Arial")
        self.assertEqual(fmt["fonts"]["eastAsia"], "黑体")
        self.assertEqual(fmt["size_pt"], 14.0)
        self.assertFalse(fmt["bold"])
        self.assertTrue(fmt["italic"])
        self.assertEqual(fmt["color_hex"], "595959")
        self.assertEqual(fmt["underline"], "single")
        self.assertEqual(fmt["highlight"], "yellow")
        self.assertEqual(fmt["character_spacing_pt"], 1.0)
        self.assertEqual(fmt["character_scale_percent"], 90)
        self.assertEqual(fmt["position_pt"], 1.0)
        self.assertEqual(fmt["provenance"]["bold"], "direct-run")

    def test_resolves_effective_paragraph_properties(self):
        from docx_model import DocxPackage
        from effective_style import EffectiveStyleResolver

        package = DocxPackage(self.path)
        paragraph = next(package.iter_paragraphs())
        fmt = EffectiveStyleResolver(package).get_effective_paragraph_format(paragraph)

        self.assertEqual(fmt["alignment"], "right")
        self.assertEqual(fmt["space_after_pt"], 6.0)
        self.assertEqual(fmt["line_rule"], "exact")
        self.assertEqual(fmt["line_spacing_pt"], 24.0)
        self.assertEqual(fmt["left_indent_pt"], 12.0)
        self.assertEqual(fmt["outline_level"], 0)
        self.assertEqual(fmt["provenance"]["space_after_pt"], "paragraph-style:Base")
        self.assertEqual(fmt["provenance"]["alignment"], "direct-paragraph")

    def test_style_chain_is_cycle_safe(self):
        from docx_model import DocxPackage
        from effective_style import EffectiveStyleResolver

        cycle_styles = MINIMAL_STYLES.replace(
            "</w:styles>",
            '<w:style w:type="paragraph" w:styleId="CycleA"><w:basedOn w:val="CycleB"/></w:style>'
            '<w:style w:type="paragraph" w:styleId="CycleB"><w:basedOn w:val="CycleA"/></w:style>'
            "</w:styles>",
        )
        cycle_path = make_docx(Path(self.temp.name) / "cycle.docx", DOCUMENT, cycle_styles, MINIMAL_THEME)
        resolver = EffectiveStyleResolver(DocxPackage(cycle_path))

        chain = resolver.style_chain("CycleA", "paragraph")
        self.assertEqual(len(chain), 2)
        self.assertEqual(set(chain), {"CycleA", "CycleB"})


if __name__ == "__main__":
    unittest.main()
