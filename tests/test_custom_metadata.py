import tempfile
import unittest
from pathlib import Path

from metadata_editor.engine import load, write_markdown, write_pdf


class CustomMetadataTests(unittest.TestCase):
    def test_markdown_scalar_edit_preserves_nested_yaml_and_body(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "sample.md"
            path.write_text("---\ntitle: Old\nproject: alpha\n# Keep comment\nsettings:\n  owner: Alice\n---\n# Body\n", encoding="utf-8")
            meta = load(str(path))
            write_markdown(meta, {"title": "New"}, custom={**meta.custom, "project": "beta"})
            text = path.read_text(encoding="utf-8")
            self.assertIn("project: beta", text)
            self.assertIn("# Keep comment\nsettings:\n  owner: Alice", text)
            self.assertTrue(text.endswith("# Body\n"))

    def test_pdf_custom_info_edit_and_delete(self):
        from pypdf import PdfReader, PdfWriter

        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "sample.pdf"
            writer = PdfWriter()
            writer.add_blank_page(width=72, height=72)
            writer.add_metadata({"/Title": "Original", "/Project": "Alpha"})
            with path.open("wb") as output:
                writer.write(output)
            meta = load(str(path))
            self.assertEqual(meta.custom["/Project"], "Alpha")
            write_pdf(meta, {**meta.values, "title": "Changed"}, False, {"/Project": "Beta"})
            self.assertEqual(load(str(path)).custom["/Project"], "Beta")
            write_pdf(load(str(path)), load(str(path)).values, False, {})
            self.assertNotIn("/Project", load(str(path)).custom)
            self.assertEqual(len(PdfReader(path).pages), 1)


if __name__ == "__main__":
    unittest.main()
