import struct
import tempfile
import unittest
import wave
from pathlib import Path
from unittest.mock import patch

from PIL import Image

from metadata_editor.engine import Metadata, clear_embedded_properties, load, verify_embedded_clear, write_audio
from metadata_editor import engine


class MediaMetadataTest(unittest.TestCase):
    def test_wav_info_read_edit_and_clear_preserve_audio(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "sample.wav"
            with wave.open(str(path), "wb") as output:
                output.setnchannels(1)
                output.setsampwidth(2)
                output.setframerate(8000)
                output.writeframes(b"\0\0" * 800)
            info = b"INFO"
            for key, value in ((b"INAM", "报告 (v2)"), (b"IART", "测试作者"), (b"ISFT", "Test Writer")):
                data = value.encode("utf-8") + b"\0"
                info += key + struct.pack("<I", len(data)) + data + (b"\0" if len(data) % 2 else b"")
            with path.open("ab") as output:
                output.write(b"LIST" + struct.pack("<I", len(info)) + info)
            with path.open("r+b") as output:
                output.seek(4)
                output.write(struct.pack("<I", path.stat().st_size - 8))
            with wave.open(str(path), "rb") as source:
                original_audio = source.readframes(source.getnframes())

            before = load(str(path))
            self.assertEqual(before.values["title"], "报告 (v2)")
            self.assertEqual(before.values["author"], "测试作者")
            self.assertIn("ISFT", before.values["_all_tags"])

            updated = dict(before.values, title="新标题", author="新作者")
            write_audio(before, updated)
            self.assertEqual(load(str(path)).values["title"], "新标题")
            self.assertIn("ISFT", load(str(path)).values["_all_tags"])

            clear_embedded_properties(load(str(path)))
            _, remaining = verify_embedded_clear(path)
            self.assertFalse(remaining)
            with wave.open(str(path), "rb") as source:
                self.assertEqual(source.readframes(source.getnframes()), original_audio)

    def test_tiff_structural_tags_are_not_metadata_residue(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "sample.tif"
            Image.new("RGB", (20, 10), "blue").save(path)
            self.assertFalse(verify_embedded_clear(path)[1])

    def test_legacy_zero_edit_time_is_not_metadata_residue(self):
        record = Metadata(Path("sample.doc"), "Legacy Office", values={"title": "", "totalTime": "0"})
        with patch.object(engine, "load", return_value=record):
            self.assertFalse(verify_embedded_clear(record.path)[1])

    def test_bundled_media_tools_precede_path(self):
        with tempfile.TemporaryDirectory() as folder:
            tool = Path(folder) / "media-tools" / "bin" / "ffprobe.exe"
            tool.parent.mkdir(parents=True)
            tool.touch()
            with patch.object(engine.sys, "_MEIPASS", folder, create=True), patch.dict(engine.os.environ, {}, clear=True):
                self.assertEqual(engine._ffprobe(), str(tool))


if __name__ == "__main__":
    unittest.main()
