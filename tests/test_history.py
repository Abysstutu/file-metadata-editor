import hashlib
import json
import os
import tempfile
import unittest
import zipfile
from pathlib import Path

from metadata_editor.engine import Metadata, clear_embedded_properties, load, write_markdown
from metadata_editor.history import (FileChangedError, append_history, delete_history_records,
                                     read_history, restore_recorded_metadata)
from metadata_editor.operations import create_output_copy


class HistoryTests(unittest.TestCase):
    def test_multiple_edits_append_to_one_file_with_result_hash(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "原文件.md"
            result = root / "副本.md"
            history = root / "history" / "edit-history.jsonl"
            source.write_text("original", encoding="utf-8")
            for title in ("第一版", "第二版"):
                result.write_text(title, encoding="utf-8")
                meta = Metadata(result, "Markdown", {"title": title}, {"extra": "值"})
                append_history(source, meta, "save", before=meta, destination=history)
            rows = [json.loads(line) for line in history.read_text(encoding="utf-8").splitlines()]
            self.assertEqual([row["metadata_after"]["values"]["title"] for row in rows], ["第一版", "第二版"])
            self.assertEqual(rows[1]["source_name"], source.name)
            self.assertEqual(rows[1]["result_name"], result.name)
            self.assertEqual(rows[1]["sha256"], hashlib.sha256(result.read_bytes()).hexdigest())
            self.assertEqual(rows[1]["metadata_after"]["custom"], {"extra": "值"})
            self.assertEqual(list(root.glob("*.metadata-backup*")), [])

    def test_markdown_write_then_history_contains_read_back(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "report.md"
            history = Path(folder) / "edit-history.jsonl"
            path.write_text("---\ntitle: Old\n---\nBody\n", encoding="utf-8")
            before = load(str(path))
            write_markdown(before, {"title": "New"})
            after = load(str(path))
            append_history(path, after, "save", before=before, destination=history)
            row = json.loads(history.read_text(encoding="utf-8"))
            self.assertEqual(row["metadata_before"]["values"]["title"], "Old")
            self.assertEqual(row["metadata_after"]["values"]["title"], "New")
            self.assertEqual(row["sha256"], hashlib.sha256(path.read_bytes()).hexdigest())

    def test_reader_skips_damaged_lines_and_keeps_valid_records(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "history.jsonl"
            path.write_text('{"result_path": "A.md", "source_name": "A.md"}\n{bad\n{"other": 1}\n', encoding="utf-8")
            records, skipped = read_history(path)
            self.assertEqual(len(records), 1)
            self.assertEqual(records[0]["result_path"], "A.md")
            self.assertEqual(skipped, 2)

    def test_delete_selected_history_preserves_other_rows_and_file(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            document = root / "report.md"
            history = root / "edit-history.jsonl"
            document.write_text("Keep this file", encoding="utf-8")
            rows = [{"result_path": f"{index}.md", "source_name": f"{index}.md"} for index in range(3)]
            history.write_text("".join(json.dumps(row) + "\n" for row in rows) + "{bad\n", encoding="utf-8")
            self.assertEqual(delete_history_records({1: rows[1]}, history), 1)
            remaining, skipped = read_history(history)
            self.assertEqual(remaining, [rows[0], rows[2]])
            self.assertEqual(skipped, 1)
            self.assertEqual(document.read_text(encoding="utf-8"), "Keep this file")

    def test_delete_selected_refuses_changed_history(self):
        with tempfile.TemporaryDirectory() as folder:
            history = Path(folder) / "edit-history.jsonl"
            current = {"result_path": "new.md"}
            history.write_text(json.dumps(current) + "\n", encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "已变化"):
                delete_history_records({0: {"result_path": "old.md"}}, history)
            self.assertEqual(read_history(history)[0], [current])

    def test_clear_all_history_then_append_again(self):
        with tempfile.TemporaryDirectory() as folder:
            history = Path(folder) / "edit-history.jsonl"
            history.write_text('{"result_path":"a.md"}\n{bad\n', encoding="utf-8")
            self.assertEqual(delete_history_records(path=history), 1)
            self.assertEqual(history.read_text(encoding="utf-8"), "")
            document = Path(folder) / "new.md"
            document.write_text("Body", encoding="utf-8")
            meta = Metadata(document, "Markdown", {"title": "New"}, {})
            append_history(document, meta, "save", before=meta, destination=history)
            self.assertEqual(len(read_history(history)[0]), 1)

    def test_restore_refuses_changed_file_without_writing(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "report.md"
            log = Path(folder) / "history.jsonl"
            path.write_text("---\ntitle: First\n---\nBody\n", encoding="utf-8")
            before = load(str(path))
            append_history(path, before, "clear", before=before, destination=log)
            record = read_history(log)[0][0]
            path.write_text("---\ntitle: Changed\n---\nBody\n", encoding="utf-8")
            before = path.read_bytes()
            with self.assertRaises(FileChangedError):
                restore_recorded_metadata(record)
            self.assertEqual(path.read_bytes(), before)

    def test_restore_can_reset_file_time_when_content_hash_matches(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "report.md"
            log = Path(folder) / "history.jsonl"
            path.write_text("---\ntitle: Original\n---\nBody\n", encoding="utf-8")
            before = load(str(path))
            append_history(path, before, "clear", before=before, destination=log)
            record = read_history(log)[0][0]
            original_hash = record["sha256"]
            os.utime(path, (path.stat().st_atime, path.stat().st_mtime - 120))
            _, after, changed = restore_recorded_metadata(record)
            self.assertTrue(changed)
            self.assertEqual(after.values["filesystem_modified"], record["metadata_before"]["values"]["filesystem_modified"])
            self.assertEqual(hashlib.sha256(path.read_bytes()).hexdigest(), original_hash)
            _, _, changed_again = restore_recorded_metadata(record)
            self.assertFalse(changed_again)

    def test_restore_writes_supported_metadata_when_hash_matches(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "report.md"
            log = Path(folder) / "history.jsonl"
            path.write_text("---\ntitle: Current\n---\nBody\n", encoding="utf-8")
            current = load(str(path))
            saved = Metadata(path, current.kind, dict(current.values), dict(current.custom))
            saved.values["title"] = "Recorded"
            append_history(path, current, "clear", before=saved, destination=log)
            record = read_history(log)[0][0]
            _, after, changed = restore_recorded_metadata(record)
            self.assertTrue(changed)
            self.assertEqual(after.values["title"], "Recorded")
            self.assertIn("Body", path.read_text(encoding="utf-8"))

    def test_clear_then_restore_markdown_from_pre_edit_snapshot(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "report.md"
            log = Path(folder) / "history.jsonl"
            path.write_text("---\ntitle: Private\nauthor: Alice\nextra: Keep\n---\nBody\n", encoding="utf-8")
            before = load(str(path))
            clear_embedded_properties(before)
            cleared = load(str(path))
            self.assertFalse(cleared.values.get("title"))
            append_history(path, cleared, "clear", before=before, verification="passed", destination=log)
            record = read_history(log)[0][0]
            self.assertEqual(record["metadata_before"]["values"]["title"], "Private")
            self.assertFalse(record["metadata_after"]["values"].get("title"))
            self.assertEqual(record["sha256"], hashlib.sha256(path.read_bytes()).hexdigest())
            _, restored, changed = restore_recorded_metadata(record)
            self.assertTrue(changed)
            self.assertEqual(restored.values["title"], "Private")
            self.assertEqual(restored.values["author"], "Alice")
            self.assertEqual(restored.custom["extra"], "Keep")
            self.assertIn("Body", path.read_text(encoding="utf-8"))

    def test_legacy_history_without_pre_edit_snapshot_cannot_restore(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "report.md"
            path.write_text("Body", encoding="utf-8")
            record = {"operation": "clear", "result_path": str(path), "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
                      "metadata_after": {"values": {"title": ""}, "custom": {}}}
            with self.assertRaisesRegex(ValueError, "没有保存操作前"):
                restore_recorded_metadata(record)

    def test_save_record_is_not_offered_as_clear_recovery(self):
        with self.assertRaisesRegex(ValueError, "只有清除记录"):
            restore_recorded_metadata({"operation": "save"})

    def test_clear_then_restore_pdf_title_and_custom_info(self):
        from pypdf import PdfWriter

        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "report.pdf"
            log = Path(folder) / "history.jsonl"
            writer = PdfWriter()
            writer.add_blank_page(width=200, height=200)
            writer.add_metadata({"/Title": "Private", "/Author": "Alice", "/Extra": "Keep"})
            with path.open("wb") as stream:
                writer.write(stream)
            before = load(str(path))
            clear_embedded_properties(before)
            after = load(str(path))
            append_history(path, after, "clear", before=before, destination=log)
            _, restored, changed = restore_recorded_metadata(read_history(log)[0][0])
            self.assertTrue(changed)
            self.assertEqual(restored.values["title"], "Private")
            self.assertEqual(restored.values["author"], "Alice")
            self.assertEqual(restored.custom["/Extra"], "Keep")

    def test_clear_then_restore_word_title(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "report.docx"
            log = Path(folder) / "history.jsonl"
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
                archive.writestr("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')
                archive.writestr("docProps/core.xml", '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Private</dc:title></cp:coreProperties>')
                archive.writestr("docProps/app.xml", '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"/>')
                archive.writestr("word/document.xml", "Body stays intact")
            before = load(str(path))
            clear_embedded_properties(before)
            append_history(path, load(str(path)), "clear", before=before, destination=log)
            _, restored, changed = restore_recorded_metadata(read_history(log)[0][0])
            self.assertTrue(changed)
            self.assertEqual(restored.values["title"], "Private")
            with zipfile.ZipFile(path) as archive:
                self.assertEqual(archive.read("word/document.xml"), b"Body stays intact")

    def test_copy_clear_restores_copy_without_touching_original(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "source.md"
            log = root / "history.jsonl"
            source.write_text("---\ntitle: Private\n---\nBody\n", encoding="utf-8")
            before = load(str(source))
            copy = create_output_copy(source, str(root), "{name}.copy")
            clear_embedded_properties(load(str(copy)))
            append_history(source, load(str(copy)), "clear", before=before, destination=log)
            record = read_history(log)[0][0]
            self.assertEqual(record["source_path"], str(source.resolve()))
            self.assertEqual(record["result_path"], str(copy.resolve()))
            _, restored, changed = restore_recorded_metadata(record)
            self.assertTrue(changed)
            self.assertEqual(restored.values["title"], "Private")
            self.assertEqual(load(str(source)).values["title"], "Private")


if __name__ == "__main__":
    unittest.main()
