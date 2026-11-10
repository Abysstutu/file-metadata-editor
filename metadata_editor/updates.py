"""Check the public GitHub Release and download its verified Windows installer."""

from __future__ import annotations

import datetime as dt
import hashlib
import json
import os
import re
import ssl
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

import truststore

from .version import APP_VERSION

LATEST_URL = "https://raw.githubusercontent.com/Abysstutu/file-metadata-editor/main/updates/latest.json"
RELEASES_URL = "https://github.com/Abysstutu/file-metadata-editor/releases"
_MAX_RELEASE_JSON = 2 * 1024 * 1024
_MAX_INSTALLER_SIZE = 1024 * 1024 * 1024
_VERSION = re.compile(r"^v?(\d+)\.(\d+)(?:\.(\d+))?(?:[-+].*)?$", re.IGNORECASE)
_DIGEST = re.compile(r"^sha256:([0-9a-fA-F]{64})$")


class UpdateError(RuntimeError):
    pass


class DownloadCancelled(UpdateError):
    pass


@dataclass(frozen=True)
class Release:
    tag: str
    version: tuple[int, int, int]
    asset_name: str
    asset_url: str
    asset_size: int
    sha256: str
    page_url: str


def version_tuple(text: str) -> tuple[int, int, int]:
    match = _VERSION.fullmatch(text.strip())
    if not match:
        raise ValueError(f"无法识别版本号：{text}")
    return tuple(int(part or 0) for part in match.groups())


def update_folder() -> Path:
    base = os.environ.get("LOCALAPPDATA")
    return (Path(base) / "FileMetadataEditor" if base else Path.home() / ".file-metadata-editor") / "updates"


def state_path() -> Path:
    return update_folder().parent / "update-state.json"


def read_state(path: Path | None = None) -> dict:
    target = path or state_path()
    try:
        value = json.loads(target.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    return value if isinstance(value, dict) else {}


def write_state(state: dict, path: Path | None = None) -> None:
    target = path or state_path()
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_suffix(".tmp")
    temporary.write_text(json.dumps(state, ensure_ascii=False), encoding="utf-8")
    os.replace(temporary, target)


def should_auto_check(now: dt.datetime | None = None, state: dict | None = None) -> bool:
    state = read_state() if state is None else state
    now = now or dt.datetime.now(dt.timezone.utc)
    try:
        last = dt.datetime.fromisoformat(state["last_checked_utc"])
        return now - last.astimezone(dt.timezone.utc) >= dt.timedelta(hours=24)
    except (KeyError, TypeError, ValueError):
        return True


def _request(url: str, timeout: int):
    request = urllib.request.Request(url, headers={
        "Accept": "application/json" if url == LATEST_URL else "application/octet-stream",
        "User-Agent": "FileMetadataEditor-Updater",
        "Cache-Control": "no-cache",
    })
    return urllib.request.urlopen(request, timeout=timeout,
                                  context=truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT))


def latest_release() -> Release:
    with _request(LATEST_URL, 10) as response:
        if urllib.parse.urlsplit(response.geturl()).scheme != "https":
            raise UpdateError("更新信息使用了不安全的连接。")
        raw = response.read(_MAX_RELEASE_JSON + 1)
    if len(raw) > _MAX_RELEASE_JSON:
        raise UpdateError("更新信息过大。")
    try:
        info = json.loads(raw)
        tag = info["tag_name"]
        version = version_tuple(tag)
        expected_name = f"FileMetadataEditor-{tag}-Windows-x64-Setup.exe"
        asset = next(item for item in info["assets"] if item.get("name") == expected_name)
        url = asset["browser_download_url"]
        digest_match = _DIGEST.fullmatch(asset["digest"])
        size = asset["size"]
    except (KeyError, StopIteration, TypeError, ValueError) as exc:
        raise UpdateError("最新版本没有可校验的 Windows 安装程序。") from exc
    if not digest_match or not isinstance(size, int) or not 0 < size <= _MAX_INSTALLER_SIZE:
        raise UpdateError("安装程序缺少有效的 SHA-256 或文件大小。")
    parsed = urllib.parse.urlsplit(url)
    prefix = f"/Abysstutu/file-metadata-editor/releases/download/{tag}/"
    if parsed.scheme != "https" or parsed.netloc.lower() != "github.com" or not parsed.path.startswith(prefix):
        raise UpdateError("安装程序下载地址无效。")
    return Release(tag, version, expected_name, url, size, digest_match.group(1).lower(),
                   info.get("html_url") or RELEASES_URL)


def is_newer(release: Release, current: str = APP_VERSION) -> bool:
    return release.version > version_tuple(current)


def _matching_file(path: Path, release: Release) -> bool:
    if not path.is_file() or path.stat().st_size != release.asset_size:
        return False
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest() == release.sha256


def download_installer(release: Release, progress: Callable[[int, int], None] | None = None,
                       cancelled: Callable[[], bool] | None = None,
                       folder: Path | None = None) -> Path:
    target_folder = folder or update_folder()
    target_folder.mkdir(parents=True, exist_ok=True)
    target = target_folder / release.asset_name
    if _matching_file(target, release):
        if progress:
            progress(release.asset_size, release.asset_size)
        return target
    partial = target.with_suffix(".download")
    digest = hashlib.sha256()
    received = 0
    try:
        with _request(release.asset_url, 30) as response, partial.open("wb") as output:
            if urllib.parse.urlsplit(response.geturl()).scheme != "https":
                raise UpdateError("下载使用了不安全的连接。")
            while True:
                if cancelled and cancelled():
                    raise DownloadCancelled("下载已取消。")
                chunk = response.read(1024 * 1024)
                if not chunk:
                    break
                received += len(chunk)
                if received > release.asset_size:
                    raise UpdateError("下载的安装程序大小不符。")
                digest.update(chunk)
                output.write(chunk)
                if progress:
                    progress(received, release.asset_size)
            output.flush()
            os.fsync(output.fileno())
        if received != release.asset_size or digest.hexdigest() != release.sha256:
            raise UpdateError("安装程序校验失败，请稍后重试。")
        os.replace(partial, target)
        return target
    finally:
        if partial.exists():
            partial.unlink()
