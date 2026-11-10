import datetime as dt
import hashlib
import io
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from metadata_editor.updates import (DownloadCancelled, Release, UpdateError,
                                     download_installer, is_newer, latest_release,
                                     read_state, should_auto_check, version_tuple,
                                     write_state)


class Response(io.BytesIO):
    def __init__(self, content: bytes, url: str):
        super().__init__(content)
        self.url = url

    def geturl(self):
        return self.url


class UpdateTests(unittest.TestCase):
    def test_version_comparison_and_daily_check(self):
        release = Release("v4.3", (4, 3, 0), "setup.exe", "https://github.com/a", 1, "0" * 64, "")
        self.assertTrue(is_newer(release, "4.2.0"))
        self.assertFalse(is_newer(release, "4.3.0-dev"))
        self.assertEqual(version_tuple("v4.3"), (4, 3, 0))
        now = dt.datetime(2026, 11, 10, tzinfo=dt.timezone.utc)
        self.assertFalse(should_auto_check(now, {"last_checked_utc": (now - dt.timedelta(hours=12)).isoformat()}))
        self.assertTrue(should_auto_check(now, {"last_checked_utc": (now - dt.timedelta(days=2)).isoformat()}))

    def test_release_requires_matching_installer_and_sha256(self):
        name = "FileMetadataEditor-v4.3-Windows-x64-Setup.exe"
        digest = hashlib.sha256(b"abc").hexdigest()
        info = {"tag_name": "v4.3", "html_url": "https://github.com/Abysstutu/file-metadata-editor/releases/tag/v4.3",
                "assets": [{"name": name, "browser_download_url":
                            f"https://github.com/Abysstutu/file-metadata-editor/releases/download/v4.3/{name}",
                            "size": 3, "digest": f"sha256:{digest}"}]}
        with patch("metadata_editor.updates._request", return_value=Response(json.dumps(info).encode(), "https://api.github.com/release")):
            release = latest_release()
        self.assertEqual(release.version, (4, 3, 0))
        self.assertEqual(release.sha256, digest)
        info["assets"][0]["digest"] = None
        with patch("metadata_editor.updates._request", return_value=Response(json.dumps(info).encode(), "https://api.github.com/release")):
            with self.assertRaises(UpdateError):
                latest_release()

    def test_repository_manifest_is_a_valid_release(self):
        manifest = Path(__file__).resolve().parent.parent / "updates" / "latest.json"
        with patch("metadata_editor.updates._request", return_value=Response(manifest.read_bytes(),
                                                                            "https://raw.githubusercontent.com/Abysstutu/file-metadata-editor/main/updates/latest.json")):
            release = latest_release()
        self.assertEqual(release.tag, "v4.1")
        self.assertEqual(release.asset_size, 151299107)

    def test_download_verifies_bytes_and_can_reuse_verified_file(self):
        name = "FileMetadataEditor-v4.3-Windows-x64-Setup.exe"
        release = Release("v4.3", (4, 3, 0), name, "https://github.com/download", 3,
                          hashlib.sha256(b"abc").hexdigest(), "")
        with tempfile.TemporaryDirectory() as folder:
            with patch("metadata_editor.updates._request", return_value=Response(b"abc", "https://github.com/download")) as request:
                path = download_installer(release, folder=Path(folder))
            self.assertEqual(path.read_bytes(), b"abc")
            self.assertEqual(request.call_count, 1)
            with patch("metadata_editor.updates._request", side_effect=AssertionError("Should reuse cache")):
                self.assertEqual(download_installer(release, folder=Path(folder)), path)

    def test_bad_download_and_cancel_leave_no_partial_file(self):
        name = "FileMetadataEditor-v4.3-Windows-x64-Setup.exe"
        release = Release("v4.3", (4, 3, 0), name, "https://github.com/download", 3,
                          hashlib.sha256(b"abc").hexdigest(), "")
        with tempfile.TemporaryDirectory() as folder:
            directory = Path(folder)
            with patch("metadata_editor.updates._request", return_value=Response(b"abd", "https://github.com/download")):
                with self.assertRaisesRegex(UpdateError, "校验失败"):
                    download_installer(release, folder=directory)
            self.assertEqual(list(directory.iterdir()), [])
            with patch("metadata_editor.updates._request", return_value=Response(b"abc", "https://github.com/download")):
                with self.assertRaises(DownloadCancelled):
                    download_installer(release, cancelled=lambda: True, folder=directory)
            self.assertEqual(list(directory.iterdir()), [])

    def test_update_state_stays_in_its_own_file(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "update-state.json"
            self.assertEqual(read_state(path), {})
            write_state({"dismissed_tag": "v4.3"}, path)
            self.assertEqual(read_state(path)["dismissed_tag"], "v4.3")


if __name__ == "__main__":
    unittest.main()
