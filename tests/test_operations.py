import tempfile
import unittest
import zipfile
from pathlib import Path

from pypdf import PdfWriter

from metadata_editor.engine import load, verify_embedded_clear, clear_embedded_properties
from metadata_editor.operations import changed_values, create_output_copy, output_copy_path, restore_from_backup, selected_fields, write_selected


class SelectiveOperationsTest(unittest.TestCase):
    def test_copy_destination_collision_and_restore(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "input.md"
            output = root / "copies"
            output.mkdir()
            source.write_text("original", encoding="utf-8")
            first = create_output_copy(source, str(output), "{name}.已处理")
            second = create_output_copy(source, str(output), "{name}.已处理")
            self.assertNotEqual(first, second)
            self.assertEqual(first.read_text(encoding="utf-8"), "original")
            self.assertEqual(second.read_text(encoding="utf-8"), "original")
            first.write_text("edited", encoding="utf-8")
            restore_from_backup(first, source)
            self.assertEqual(first.read_text(encoding="utf-8"), "original")
            self.assertEqual(source.read_text(encoding="utf-8"), "original")
            with self.assertRaises(ValueError):
                output_copy_path(source, str(output), "{name.__class__}")

    def test_markdown_selective_clear_and_batch_title(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "report.md"
            path.write_text("---\ntitle: Old\nauthor: Alice\ntags: secret\n---\n# Body\n", encoding="utf-8")
            before = load(str(path))
            keys = {key for key, _ in selected_fields(before, {"身份与来源"})}
            self.assertIn("author", keys)
            self.assertNotIn("title", keys)
            after = write_selected(before, {key: "" for key in keys})
            self.assertEqual(after.values.get("title"), "Old")
            self.assertFalse(after.values.get("author"))
            self.assertIn("# Body", path.read_text(encoding="utf-8"))
            self.assertIn(("作者", "Alice", ""), changed_values(before, after, keys))
            final = write_selected(after, {"title": "report"})
            self.assertEqual(final.values.get("title"), "report")
            self.assertEqual(final.values.get("tags"), "secret")

    def test_pdf_title_clear_keeps_author(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "sample.pdf"
            writer = PdfWriter()
            writer.add_blank_page(width=100, height=100)
            writer.add_metadata({"/Title": "Old", "/Author": "Alice"})
            with path.open("wb") as stream:
                writer.write(stream)
            before = load(str(path))
            after = write_selected(before, {"title": ""})
            self.assertFalse(after.values.get("title"))
            self.assertEqual(after.values.get("author"), "Alice")
            clear_embedded_properties(after)
            _, remaining = verify_embedded_clear(path)
            self.assertFalse(remaining)

    def test_archive_comment_clear_preserves_entries(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "sample.zip"
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr("data.txt", "payload")
                archive.comment = b"note"
            after = write_selected(load(str(path)), {"comment": ""})
            self.assertFalse(after.values.get("comment"))
            with zipfile.ZipFile(path) as archive:
                self.assertEqual(archive.read("data.txt"), b"payload")

    def test_office_selective_clear_preserves_author_and_custom(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "sample.docx"
            core = b'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Old</dc:title><dc:creator>Alice</dc:creator></cp:coreProperties>'
            custom = b'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property name="Project" pid="2" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"><vt:lpwstr>Alpha</vt:lpwstr></property></Properties>'
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr("docProps/core.xml", core)
                archive.writestr("docProps/custom.xml", custom)
                archive.writestr("word/document.xml", "<document/>")
            before = load(str(path))
            after = write_selected(before, {"title": ""})
            self.assertFalse(after.values.get("title"))
            self.assertEqual(after.values.get("creator"), "Alice")
            self.assertEqual(after.custom.get("Project"), "Alpha")


if __name__ == "__main__":
    unittest.main()
