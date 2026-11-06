from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

try:
    import normalize_word_inline_math_layout as normalizer
except ImportError:  # lxml is an optional dependency of the checker only.
    normalizer = None

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
DOCUMENT_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{ns}">
  <w:body>
    <w:p>
      <w:pPr><w:ind w:firstLine="480" w:firstLineChars="200"/></w:pPr>
      <w:r><w:t>\u5176\u4e2d\uff0c</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:spacing w:line="360" w:lineRule="exact"/>
      </w:pPr>
      <w:r><w:t>\u901a\u8fc7</w:t></w:r>
      <w:r>
        <w:rPr><w:position w:val="3"/></w:rPr>
        <w:object><w:embed/></w:object>
      </w:r>
      <w:r><w:t>\u6c42\u89e3\u3002</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="MTDisplayEquation"/></w:pPr>
      <w:r><w:object><w:embed/></w:object></w:r>
    </w:p>
  </w:body>
</w:document>
""".format(ns=W_NS)
STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{ns}">
  <w:style w:type="paragraph" w:styleId="BodyText">
    <w:pPr><w:spacing w:line="400" w:lineRule="atLeast"/></w:pPr>
  </w:style>
</w:styles>
""".format(ns=W_NS)


def build_docx(path: Path) -> None:
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", DOCUMENT_TEMPLATE)
        archive.writestr("word/styles.xml", STYLES_XML)
        archive.writestr("word/embeddings/oleObject1.bin", b"MathType-OLE-bytes")
        archive.writestr("[Content_Types].xml", "<Types/>")


@unittest.skipUnless(normalizer is not None, "lxml is not installed")
class NormalizeInlineMathLayoutTests(unittest.TestCase):
    def test_check_reports_unnormalized_paragraphs(self) -> None:
        result = normalizer.inspect_document(
            DOCUMENT_TEMPLATE.encode("utf-8"), STYLES_XML.encode("utf-8")
        )
        self.assertFalse(result["ok"])
        self.assertEqual(result["where_paragraphs_with_indent"], [1])
        self.assertEqual(result["inline_paragraphs_not_center"], [2])
        self.assertEqual(result["inline_paragraphs_with_nonstandard_position"], [2])
        self.assertEqual(result["inline_paragraphs_without_minimum_line_spacing"], [2])

    def test_patch_normalizes_and_keeps_display_equations_untouched(self) -> None:
        patched, report = normalizer.patch_document(
            DOCUMENT_TEMPLATE.encode("utf-8"), STYLES_XML.encode("utf-8")
        )
        validation = normalizer.inspect_document(patched, STYLES_XML.encode("utf-8"))
        self.assertTrue(validation["ok"])
        self.assertEqual(report["where_paragraph_count"], 1)
        self.assertEqual(report["inline_math_paragraph_count"], 1)
        self.assertEqual(report["inline_object_run_positions_normalized"], 1)
        self.assertIn(b'w:firstLine="0"', patched)

    def test_rewrite_preserves_embedded_ole_parts(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.docx"
            output = Path(directory) / "output.docx"
            build_docx(source)
            report, validation, preservation = normalizer.rewrite_docx(source, output)
            self.assertTrue(validation["ok"])
            self.assertTrue(preservation["ole_parts_unchanged"])
            self.assertEqual(preservation["ole_count"], 1)
            self.assertGreater(report["inline_math_paragraph_count"], 0)
            with ZipFile(output) as archive:
                self.assertEqual(
                    archive.read("word/embeddings/oleObject1.bin"),
                    b"MathType-OLE-bytes",
                )


if __name__ == "__main__":
    unittest.main()
