from __future__ import annotations

import datetime as dt
import base64
import ctypes
import os
import re
import shutil
import stat
import struct
import subprocess
import sys
import tempfile
import zipfile
import json
from fractions import Fraction
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET

Image = None
ExifTags = None
MutagenFile = None


def _load_pillow():
    global Image, ExifTags
    if Image is None:
        try:
            from PIL import Image as pillow_image, ExifTags as pillow_exif_tags
        except ImportError as exc:
            raise RuntimeError("图片功能需要安装 Pillow。") from exc
        Image, ExifTags = pillow_image, pillow_exif_tags
    return Image, ExifTags


def _load_mutagen():
    global MutagenFile
    if MutagenFile is None:
        try:
            from mutagen import File as mutagen_file
        except ImportError as exc:
            raise RuntimeError("音频功能需要安装 mutagen：python -m pip install mutagen") from exc
        MutagenFile = mutagen_file
    return MutagenFile

CORE = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC = "http://purl.org/dc/elements/1.1/"
DCTERMS = "http://purl.org/dc/terms/"
CP = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
APP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
NS = {"cp": CORE, "dc": DC, "dcterms": DCTERMS, "app": APP}
for prefix, uri in {"cp": CORE, "dc": DC, "dcterms": DCTERMS, "vt": VT}.items():
    ET.register_namespace(prefix, uri)

OFFICE_SUFFIXES = {".docx", ".xlsx", ".pptx", ".docm", ".xlsm", ".pptm", ".dotx", ".xltx", ".potx", ".dotm", ".xltm", ".potm", ".ppsx", ".ppsm"}
LEGACY_SUFFIXES = {".doc", ".xls", ".ppt"}
MARKDOWN_SUFFIXES = {".md", ".markdown", ".mdown", ".mkd"}
IMAGE_SUFFIXES = {".jpg", ".jpeg", ".tif", ".tiff", ".webp"}
ZIP_SUFFIXES = {".zip"}
AUDIO_SUFFIXES = {".mp3", ".flac", ".m4a", ".wav"}
VIDEO_SUFFIXES = {".mp4", ".mov", ".mkv"}

FIELDS = [
    ("title", "标题"), ("subject", "主题"), ("creator", "作者"), ("lastModifiedBy", "最后保存者"),
    ("keywords", "标记 / 关键词"), ("category", "类别"), ("description", "备注"), ("contentStatus", "状态"),
    ("revision", "修订号"), ("created", "创建内容时间"), ("modified", "最后一次保存时间"),
    ("lastPrinted", "最后打印时间"), ("totalTime", "总编辑时长（分钟）"), ("application", "程序名称"),
    ("company", "公司"), ("manager", "管理者"), ("hyperlinkBase", "超链接基址"), ("template", "文档模板"),
]
PDF_FIELDS = [("title", "标题"), ("author", "作者"), ("subject", "主题"), ("keywords", "关键词"), ("creator", "Creator"), ("producer", "Producer"), ("created", "创建内容时间"), ("modified", "修改内容时间")]
MARKDOWN_FIELDS = [("title", "标题"), ("author", "作者"), ("date", "日期"), ("description", "描述"), ("tags", "标签"), ("category", "分类")]
IMAGE_FIELDS = [("title", "标题"), ("author", "作者"), ("description", "描述"), ("created", "拍摄时间"), ("gps", "GPS")]
ZIP_FIELDS = [("comment", "压缩包注释"), ("file_count", "文件数量")]
MEDIA_FIELDS = [("title", "标题"), ("author", "作者"), ("album", "专辑"), ("date", "日期"), ("genre", "类型"), ("description", "描述")]
WAV_INFO_FIELDS = {"title": b"INAM", "author": b"IART", "album": b"IPRD", "date": b"ICRD", "genre": b"IGNR", "description": b"ICMT"}
TIFF_STRUCTURAL_TAGS = {256, 257, 258, 259, 262, 273, 277, 278, 279, 282, 283, 284, 296, 339}


@dataclass
class Metadata:
    path: Path
    kind: str
    values: dict[str, str] = field(default_factory=dict)
    custom: dict[str, str] = field(default_factory=dict)
    readonly: dict[str, str] = field(default_factory=dict)


def iso_display(value: str) -> str:
    return value.replace("T", " ").replace("Z", "")[:19] if value else ""


def iso_store(value: str) -> str:
    if not value:
        return ""
    value = value.strip().replace(" ", "T")
    if "T" in value and not value.endswith("Z") and "+" not in value[10:]:
        return value + "Z"
    return value


def filesystem_values(path: Path) -> dict[str, str]:
    stat = path.stat()
    return {
        "filesystem_created": dt.datetime.fromtimestamp(stat.st_ctime).strftime("%Y-%m-%d %H:%M:%S"),
        "filesystem_modified": dt.datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
    }


def _markdown_parts(path: Path) -> tuple[str, str, str]:
    text = path.read_text(encoding="utf-8-sig")
    lines = text.splitlines(keepends=True)
    if not lines or lines[0].strip() != "---":
        return "", text, ""
    for index in range(1, len(lines)):
        if lines[index].strip() in {"---", "..."}:
            return "".join(lines[1:index]), "".join(lines[index + 1:]), "".join(lines[:index + 1])
    return "", text, ""


def read_markdown(path: Path) -> Metadata:
    result = Metadata(path, "Markdown")
    result.values.update(filesystem_values(path))
    frontmatter, _, _ = _markdown_parts(path)
    known = {key for key, _ in MARKDOWN_FIELDS}
    for line in frontmatter.splitlines():
        if not line.strip() or line[0].isspace() or line.lstrip().startswith("#") or ":" not in line:
            continue
        key, value = line.split(":", 1)
        key, value = key.strip(), value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in {"'", '"'}:
            value = value[1:-1]
        result.values[key] = value
        if key not in known:
            result.custom[key] = value
    return result


def write_markdown(meta: Metadata, values: dict[str, str], clear: bool = False,
                   custom: dict[str, str] | None = None) -> None:
    frontmatter, body, _ = _markdown_parts(meta.path)
    if clear:
        with meta.path.open("w", encoding="utf-8", newline="") as output:
            output.write(body)
        return
    fields = [(key, values.get(key, "").strip()) for key, _ in MARKDOWN_FIELDS]
    known = {key for key, _ in MARKDOWN_FIELDS}
    requested = meta.custom if custom is None else custom
    for key, value in requested.items():
        if not key or key in known or ":" in key or "\n" in key:
            raise ValueError(f"无效的 Markdown 自定义属性名称：{key}")
        if "\n" in value:
            raise ValueError(f"暂不支持修改多行 YAML 属性：{key}")
        fields.append((key, value))
    desired = dict(fields)
    newline = "\r\n" if "\r\n" in frontmatter else "\n"
    original_lines = frontmatter.splitlines(keepends=True)
    kept = []
    seen = set()
    for index, line in enumerate(original_lines):
        if not line.strip() or line[0].isspace() or line.lstrip().startswith("#") or ":" not in line:
            kept.append(line)
            continue
        key, original = line.split(":", 1)
        key = key.strip()
        if not key or (key not in desired and key not in meta.custom):
            kept.append(line)
            continue
        if key not in desired:
            nested = index + 1 < len(original_lines) and original_lines[index + 1][:1].isspace()
            if nested or original.strip().startswith(("|", ">", "[", "{")):
                raise ValueError(f"暂不支持逐项删除嵌套或多行 YAML 属性：{key}")
            continue
        seen.add(key)
        wanted = desired[key]
        previous = meta.values.get(key, "") if key in known else meta.custom.get(key, "")
        if wanted == previous:
            kept.append(line)
            continue
        nested = index + 1 < len(original_lines) and original_lines[index + 1][:1].isspace()
        if nested or original.strip().startswith(("|", ">", "[", "{")):
            raise ValueError(f"暂不支持逐项修改嵌套或多行 YAML 属性：{key}")
        if wanted:
            scalar = json.dumps(wanted, ensure_ascii=False) if any(char in wanted for char in ":#") else wanted
            kept.append(f"{key}: {scalar}{newline}")
    for key, value in fields:
        if key in seen or not value:
            continue
        scalar = json.dumps(value, ensure_ascii=False) if any(char in value for char in ":#") else value
        kept.append(f"{key}: {scalar}{newline}")
    with meta.path.open("w", encoding="utf-8", newline="") as output:
        output.write(f"---{newline}" + "".join(kept) + f"---{newline}" + body)


def read_image(path: Path) -> Metadata:
    result = Metadata(path, "Image")
    result.values.update(filesystem_values(path))
    image_api, exif_tags = _load_pillow()
    with image_api.open(path) as image:
        exif = image.getexif()
        names = {exif_tags.TAGS.get(key, str(key)): value for key, value in exif.items()} if exif_tags else {}
        result.values.update({"title": _xp_text(exif.get(40091)) or str(names.get("ImageDescription", "")),
                              "author": _xp_text(exif.get(40093)) or str(names.get("Artist", "")),
                              "description": _xp_text(exif.get(40092)) or str(names.get("UserComment", "")),
                              "created": str(names.get("DateTimeOriginal", "")),
                              "gps": _gps_text(exif.get_ifd(34853))})
        result.values["_all_exif"] = "\\n".join(
            f"{exif_tags.TAGS.get(key, key) if exif_tags else key}: {value}"
            for key, value in exif.items()
            if str(value).strip() and not (path.suffix.lower() in {".tif", ".tiff"} and key in TIFF_STRUCTURAL_TAGS)
        )
        extra_metadata = [key for key in ("xmp", "photoshop", "iptc", "comment") if image.info.get(key)]
        result.values["_extra_image_metadata"] = ", ".join(extra_metadata)
    return result


def _xp_text(value: object) -> str:
    if isinstance(value, (bytes, bytearray)):
        return bytes(value).decode("utf-16le", "replace").rstrip("\0")
    return ""


def _gps_text(gps: dict) -> str:
    try:
        def coordinate(tag: int, ref: int) -> float:
            degrees, minutes, seconds = (float(part) for part in gps[tag])
            value = degrees + minutes / 60 + seconds / 3600
            return -value if gps.get(ref, "").upper() in {"S", "W"} else value
        latitude, longitude = coordinate(2, 1), coordinate(4, 3)
        return ", ".join(f"{value:.6f}".rstrip("0").rstrip(".") for value in (latitude, longitude))
    except (KeyError, TypeError, ValueError, ZeroDivisionError):
        return ""


def _gps_ifd(value: str) -> dict:
    from PIL.TiffImagePlugin import IFDRational

    parts = [part.strip() for part in value.split(",")]
    if len(parts) != 2:
        raise ValueError("GPS 格式应为“纬度, 经度”，例如 39.90, 116.40。")
    try:
        latitude, longitude = (float(part) for part in parts)
    except ValueError as exc:
        raise ValueError("GPS 坐标必须是数字。") from exc
    if not -90 <= latitude <= 90 or not -180 <= longitude <= 180:
        raise ValueError("GPS 纬度范围为 -90~90，经度范围为 -180~180。")

    def dms(number: float):
        absolute = abs(number)
        degrees = int(absolute)
        minutes = int((absolute - degrees) * 60)
        seconds = Fraction((absolute - degrees - minutes / 60) * 3600).limit_denominator(1000000)
        return (IFDRational(degrees, 1), IFDRational(minutes, 1),
                IFDRational(seconds.numerator, seconds.denominator))

    return {1: "S" if latitude < 0 else "N", 2: dms(latitude),
            3: "W" if longitude < 0 else "E", 4: dms(longitude)}


def write_image(meta: Metadata, values: dict[str, str], clear: bool = False) -> None:
    image_api, _ = _load_pillow()
    with image_api.open(meta.path) as image:
        exif = image.getexif()
        if clear:
            for tag in list(exif):
                del exif[tag]
        else:
            tags = {"title": (40091, 270), "author": (40093, 315), "description": (40092, 37510)}
            for key, (xp_tag, ascii_tag) in tags.items():
                value = values.get(key, "").strip()
                if value:
                    exif[xp_tag] = value.encode("utf-16le") + b"\0\0"
                    if value.isascii() and key != "description": exif[ascii_tag] = value
                    elif ascii_tag in exif: del exif[ascii_tag]
                else:
                    for tag in (xp_tag, ascii_tag):
                        if tag in exif: del exif[tag]
            created = values.get("created", "").strip()
            if created: exif[36867] = created
            elif 36867 in exif: del exif[36867]
            gps = values.get("gps", "").strip()
            if gps: exif[34853] = _gps_ifd(gps)
            elif 34853 in exif: del exif[34853]
        tmp = meta.path.with_name(f".{meta.path.name}.metadata-tmp")
        image.save(tmp, exif=exif.tobytes(), format=image.format)
    os.replace(tmp, meta.path)


def read_zip(path: Path) -> Metadata:
    result = Metadata(path, "Archive")
    result.values.update(filesystem_values(path))
    with zipfile.ZipFile(path) as archive:
        result.values.update({"comment": archive.comment.decode("utf-8", "replace"), "file_count": str(len(archive.infolist()))})
    return result


def write_zip(meta: Metadata, values: dict[str, str], clear: bool = False) -> None:
    tmp = meta.path.with_name(f".{meta.path.name}.metadata-tmp")
    with zipfile.ZipFile(meta.path) as source, zipfile.ZipFile(tmp, "w") as target:
        target.comment = b"" if clear else values.get("comment", "").encode("utf-8")
        for item in source.infolist():
            target.writestr(item, source.read(item.filename))
    os.replace(tmp, meta.path)


def _wav_info(path: Path) -> dict[bytes, str]:
    """Read RIFF LIST/INFO fields, which Mutagen's WAVE reader does not expose."""
    fields = {}
    with path.open("rb") as stream:
        header = stream.read(12)
        if header[:4] != b"RIFF" or header[8:] != b"WAVE":
            return fields
        while chunk := stream.read(8):
            if len(chunk) != 8:
                break
            kind, size = chunk[:4], struct.unpack("<I", chunk[4:])[0]
            if kind == b"LIST" and size >= 4:
                list_type = stream.read(4)
                payload = stream.read(size - 4)
                if list_type == b"INFO":
                    offset = 0
                    while offset + 8 <= len(payload):
                        tag = payload[offset:offset + 4]
                        length = struct.unpack("<I", payload[offset + 4:offset + 8])[0]
                        raw = payload[offset + 8:offset + 8 + length].rstrip(b"\0")
                        try:
                            value = raw.decode("utf-8-sig")
                        except UnicodeDecodeError:
                            value = raw.decode("mbcs" if os.name == "nt" else "latin-1", "replace")
                        fields[tag] = value
                        offset += 8 + length + (length % 2)
            else:
                stream.seek(size, os.SEEK_CUR)
            if size % 2:
                stream.seek(1, os.SEEK_CUR)
    return fields


def _write_wav_info(path: Path, fields: dict[bytes, str]) -> None:
    """Replace INFO chunks while streaming audio data unchanged."""
    descriptor, tmp_name = tempfile.mkstemp(suffix=".wav", dir=path.parent)
    os.close(descriptor)
    tmp = Path(tmp_name)
    try:
        with path.open("rb") as source, tmp.open("wb") as target:
            header = source.read(12)
            if header[:4] != b"RIFF" or header[8:] != b"WAVE":
                raise ValueError("仅支持标准 RIFF/WAVE 文件的 INFO 属性。")
            target.write(header[:4] + b"\0\0\0\0" + header[8:])
            while chunk := source.read(8):
                if len(chunk) != 8:
                    raise ValueError("WAV 数据块不完整。")
                kind, size = chunk[:4], struct.unpack("<I", chunk[4:])[0]
                if kind == b"LIST" and size >= 4:
                    list_type = source.read(4)
                    if len(list_type) != 4:
                        raise ValueError("WAV LIST 数据块不完整。")
                    if list_type == b"INFO":
                        source.seek(size - 4 + (size % 2), os.SEEK_CUR)
                        continue
                    target.write(chunk + list_type)
                    remaining = size - 4 + (size % 2)
                else:
                    target.write(chunk)
                    remaining = size + (size % 2)
                while remaining:
                    data = source.read(min(1024 * 1024, remaining))
                    if not data:
                        raise ValueError("WAV 数据块不完整。")
                    target.write(data)
                    remaining -= len(data)
            if fields:
                payload = bytearray(b"INFO")
                for tag, value in fields.items():
                    raw = value.encode("utf-8") + b"\0"
                    payload.extend(tag + struct.pack("<I", len(raw)) + raw)
                    if len(raw) % 2:
                        payload.append(0)
                target.write(b"LIST" + struct.pack("<I", len(payload)) + payload)
                if len(payload) % 2:
                    target.write(b"\0")
            riff_size = target.tell() - 8
            if riff_size > 0xFFFFFFFF:
                raise ValueError("WAV 超过标准 RIFF 的 4 GB 大小限制。")
            target.seek(4)
            target.write(struct.pack("<I", riff_size))
        os.replace(tmp, path)
    finally:
        tmp.unlink(missing_ok=True)


def read_audio(path: Path) -> Metadata:
    mutagen_file = _load_mutagen()
    result = Metadata(path, "Audio"); result.values.update(filesystem_values(path))
    audio = mutagen_file(path, easy=True)
    if audio is not None:
        for key, target in {"title": "title", "artist": "author", "album": "album", "date": "date", "genre": "genre", "comment": "description"}.items():
            value = audio.get(key, [""])
            result.values[target] = str(value[0]) if value else ""
        result.values["_all_tags"] = "\\n".join(
            f"{key}: {value}"
            for key, value in (audio.tags or {}).items()
            if str(value).strip()
        )
    if path.suffix.lower() == ".mp3":
        from mutagen.id3 import ID3, ID3NoHeaderError

        try:
            comments = ID3(path).getall("COMM")
        except ID3NoHeaderError:
            comments = []
        if comments:
            result.values["description"] = str(comments[0].text[0]) if comments[0].text else ""
            result.values["_all_tags"] = "\n".join(filter(None, [result.values.get("_all_tags", ""),
                *(f"COMM ({frame.lang}/{frame.desc}): {', '.join(frame.text)}" for frame in comments)]))
    if path.suffix.lower() == ".wav":
        info = _wav_info(path)
        for key, tag in WAV_INFO_FIELDS.items():
            if info.get(tag):
                result.values[key] = info[tag]
        info_lines = [f"{tag.decode('ascii', 'replace')}: {value}" for tag, value in info.items() if value.strip()]
        result.values["_all_tags"] = "\n".join(filter(None, [result.values.get("_all_tags", ""), *info_lines]))
    return result


def write_audio(meta: Metadata, values: dict[str, str], clear: bool = False) -> None:
    mutagen_file = _load_mutagen()
    audio = mutagen_file(meta.path, easy=True)
    if audio is None: raise ValueError("无法识别音频容器。")
    if meta.path.suffix.lower() == ".wav":
        if clear:
            if audio.tags:
                audio.delete()
            _write_wav_info(meta.path, {})
        else:
            info = _wav_info(meta.path)
            for key, tag in WAV_INFO_FIELDS.items():
                value = values.get(key, "").strip()
                if value:
                    info[tag] = value
                else:
                    info.pop(tag, None)
            _write_wav_info(meta.path, info)
        return
    if clear:
        audio.delete()
        return
    if audio.tags is None: audio.add_tags()
    mapping = {"title": "title", "author": "artist", "album": "album", "date": "date", "genre": "genre"}
    if meta.path.suffix.lower() != ".mp3": mapping["description"] = "comment"
    for source, target in mapping.items():
        value = "" if clear else values.get(source, "").strip()
        if value: audio[target] = [value]
        elif target in audio: del audio[target]
    audio.save()
    if meta.path.suffix.lower() == ".mp3":
        from mutagen.id3 import COMM, ID3

        description = values.get("description", "").strip()
        if description != meta.values.get("description", ""):
            tags = ID3(meta.path)
            tags.delall("COMM")
            if description: tags.add(COMM(encoding=3, lang="eng", desc="", text=description))
            tags.save(meta.path)


def _media_executable(name: str) -> str:
    configured = os.environ.get(f"FILE_METADATA_{name.upper()}")
    if configured:
        return configured
    package_root = Path(__file__).resolve().parent.parent
    bundled_root = Path(getattr(sys, "_MEIPASS", package_root))
    executable = bundled_root / "media-tools" / "bin" / (f"{name}.exe" if os.name == "nt" else name)
    return str(executable) if executable.exists() else name


def _ffprobe() -> str:
    return _media_executable("ffprobe")


def _ffmpeg() -> str:
    return _media_executable("ffmpeg")


def _video_structural_tag(scope: str, key: str, value: object) -> bool:
    name = key.lower()
    if scope == "容器" and name in {"major_brand", "minor_version", "compatible_brands", "encoder"}:
        return True
    if scope.startswith("音视频流"):
        if name in {"duration", "encoder", "bps", "number_of_frames"}:
            return True
        if name == "language" and str(value).lower() == "und":
            return True
        if name == "handler_name" and str(value) in {"VideoHandler", "SoundHandler", "SubtitleHandler"}:
            return True
        if name == "vendor_id":
            return True
    return False


def read_video(path: Path) -> Metadata:
    result = Metadata(path, "Video"); result.values.update(filesystem_values(path))
    try:
        completed = subprocess.run(
            [_ffprobe(), "-v", "quiet", "-print_format", "json", "-show_format", "-show_streams", "-show_chapters", str(path)],
            text=True, encoding="utf-8", errors="replace", capture_output=True, check=True,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        probe = json.loads(completed.stdout)
        format_info = probe.get("format", {})
        tags = format_info.get("tags", {})
        normalized_tags = {str(key).lower(): value for key, value in tags.items()}
        for source, target in {"title": "title", "artist": "author", "comment": "description", "date": "date",
                               "album": "album", "genre": "genre"}.items():
            result.values[target] = str(normalized_tags.get(source, ""))
        all_tags = []
        scopes = [("容器", tags)]
        scopes.extend((f"音视频流 {item.get('index', index)}", item.get("tags", {})) for index, item in enumerate(probe.get("streams", [])))
        scopes.extend((f"章节 {index + 1}", item.get("tags", {})) for index, item in enumerate(probe.get("chapters", [])))
        for scope, scoped_tags in scopes:
            all_tags.extend(f"{scope} / {key}: {value}" for key, value in scoped_tags.items()
                            if str(value).strip() and not _video_structural_tag(scope, key, value))
        result.values["_all_tags"] = "\n".join(all_tags)
    except FileNotFoundError as exc:
        raise RuntimeError("视频功能需要 ffprobe；请安装 FFmpeg，或设置 FILE_METADATA_FFPROBE。") from exc
    return result


def write_video(meta: Metadata, values: dict[str, str], clear: bool = False) -> None:
    tmp = meta.path.with_name(f".{meta.path.name}.metadata-tmp{meta.path.suffix}")
    command = [_ffmpeg(), "-y", "-i", str(meta.path), "-map", "0", "-c", "copy"]
    if clear: command += ["-map_metadata", "-1", "-map_metadata:s", "-1", "-map_chapters", "-1"]
    else:
        for key, tag in (("title", "title"), ("author", "artist"), ("description", "comment"),
                         ("date", "date"), ("album", "album"), ("genre", "genre")):
            if key in values:
                command += ["-metadata", f"{tag}={values[key].strip()}"]
    command += [str(tmp)]
    try:
        subprocess.run(command, check=True, capture_output=True, text=True, encoding="utf-8", errors="replace",
                       creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except FileNotFoundError as exc:
        raise RuntimeError("视频写入需要 FFmpeg；请安装 FFmpeg，或将其加入 PATH。") from exc
    except subprocess.CalledProcessError as exc:
        raise RuntimeError(f"视频元数据写入失败：{(exc.stderr or '').strip()[-500:]}") from exc
    os.replace(tmp, meta.path)


def set_filesystem_times(path: Path, created: str, modified: str) -> None:
    """Set Windows creation and last-write timestamps without touching document data."""
    try:
        created_dt = dt.datetime.strptime(created, "%Y-%m-%d %H:%M:%S") if created else None
        modified_dt = dt.datetime.strptime(modified, "%Y-%m-%d %H:%M:%S") if modified else None
    except ValueError as exc:
        raise ValueError("文件时间格式必须为 YYYY-MM-DD HH:MM:SS。") from exc
    if os.name != "nt":
        if modified_dt: os.utime(path, (path.stat().st_atime, modified_dt.timestamp()))
        return
    # Windows FILETIME counts 100-nanosecond intervals since 1601-01-01 UTC.
    class FILETIME(ctypes.Structure):
        _fields_ = [("dwLowDateTime", ctypes.c_ulong), ("dwHighDateTime", ctypes.c_ulong)]
    def as_filetime(value: dt.datetime | None):
        if value is None: return None
        ticks = int((value.astimezone(dt.timezone.utc) - dt.datetime(1601, 1, 1, tzinfo=dt.timezone.utc)).total_seconds() * 10_000_000)
        return FILETIME(ticks & 0xFFFFFFFF, ticks >> 32)
    creation, write = as_filetime(created_dt), as_filetime(modified_dt)
    handle = ctypes.windll.kernel32.CreateFileW(str(path), 0x100, 0, None, 3, 0x80, None)
    if handle == ctypes.c_void_p(-1).value: raise OSError("无法打开文件以更新文件系统时间。")
    try:
        if not ctypes.windll.kernel32.SetFileTime(handle, ctypes.byref(creation) if creation else None, None, ctypes.byref(write) if write else None):
            raise OSError("更新文件系统时间失败；请确认文件未被其他程序占用。")
    finally:
        ctypes.windll.kernel32.CloseHandle(handle)


def inplace_write_block_reason(path: Path) -> str | None:
    """Return a user-facing reason when replacing this file is unsafe."""
    attributes = getattr(path.stat(), "st_file_attributes", 0)
    if os.name == "nt" and attributes & stat.FILE_ATTRIBUTE_READONLY:
        return "文件标记为只读"
    try:
        with path.open("r+b"):
            pass
    except PermissionError:
        return "文件无写入权限，或正被 Word、WPS、微信、同步软件等程序占用"
    except OSError as exc:
        return f"无法写入源文件：{exc}"
    try:
        descriptor, probe_name = tempfile.mkstemp(prefix=".metadata-write-test-", dir=path.parent)
        os.close(descriptor)
        Path(probe_name).unlink(missing_ok=True)
    except PermissionError:
        return "当前文件夹没有创建新文件的权限"
    except OSError as exc:
        return f"当前文件夹不可写：{exc}"
    return None


def make_writable(path: Path) -> None:
    """Remove the read-only bit from a program-created backup or copy."""
    try:
        os.chmod(path, path.stat().st_mode | stat.S_IWRITE)
        if os.name == "nt":
            attributes = ctypes.windll.kernel32.GetFileAttributesW(str(path))
            if attributes != 0xFFFFFFFF:
                # FILE_ATTRIBUTE_READONLY = 0x1
                ctypes.windll.kernel32.SetFileAttributesW(str(path), attributes & ~0x1)
    except OSError as exc:
        raise PermissionError(f"已生成副本，但无法解除其只读属性：{exc}") from exc


def suggested_copy_path(path: Path, label: str = "已清除属性") -> Path:
    """Choose a non-destructive, recognizable output name beside the source."""
    candidate = path.with_name(f"{path.stem}.{label}{path.suffix}")
    index = 2
    while candidate.exists():
        candidate = path.with_name(f"{path.stem}.{label} ({index}){path.suffix}")
        index += 1
    return candidate


def _text(root: ET.Element | None, xpath: str) -> str:
    node = root.find(xpath, NS) if root is not None else None
    return node.text or "" if node is not None else ""


def _xml_from_zip(zf: zipfile.ZipFile, name: str) -> ET.Element | None:
    try:
        return ET.fromstring(zf.read(name))
    except KeyError:
        return None


def read_office(path: Path) -> Metadata:
    result = Metadata(path, "Office")
    result.values.update(filesystem_values(path))
    with zipfile.ZipFile(path) as zf:
        core, app, custom = (_xml_from_zip(zf, "docProps/core.xml"), _xml_from_zip(zf, "docProps/app.xml"), _xml_from_zip(zf, "docProps/custom.xml"))
    core_map = {"title": "dc:title", "subject": "dc:subject", "creator": "dc:creator", "lastModifiedBy": "cp:lastModifiedBy", "keywords": "cp:keywords", "category": "cp:category", "description": "dc:description", "contentStatus": "cp:contentStatus", "revision": "cp:revision", "created": "dcterms:created", "modified": "dcterms:modified"}
    app_map = {"application": "app:Application", "company": "app:Company", "manager": "app:Manager", "hyperlinkBase": "app:HyperlinkBase", "template": "app:Template", "lastPrinted": "app:LastPrinted", "totalTime": "app:TotalTime"}
    for key, xp in core_map.items(): result.values[key] = iso_display(_text(core, xp)) if key in {"created", "modified"} else _text(core, xp)
    for key, xp in app_map.items(): result.values[key] = iso_display(_text(app, xp)) if key == "lastPrinted" else _text(app, xp)
    if custom is not None:
        for prop in custom.findall(f"{{{CP}}}property"):
            value = next(iter(prop), None)
            result.custom[prop.get("name", "未命名属性")] = value.text or "" if value is not None else ""
    return result


def _set(root: ET.Element, tag: str, value: str) -> None:
    node = root.find(tag, NS)
    if node is None:
        node = ET.SubElement(root, tag.replace("cp:", f"{{{CORE}}}").replace("dc:", f"{{{DC}}}").replace("dcterms:", f"{{{DCTERMS}}}").replace("app:", f"{{{APP}}}"))
    node.text = value


def write_office(meta: Metadata, new_values: dict[str, str], custom: dict[str, str], clear: bool) -> None:
    with zipfile.ZipFile(meta.path) as original:
        core, app, custom_root = (_xml_from_zip(original, "docProps/core.xml"), _xml_from_zip(original, "docProps/app.xml"), _xml_from_zip(original, "docProps/custom.xml"))
        if core is None: raise ValueError("此文件缺少 Office 核心属性包，无法安全写入。")
        if app is None: app = ET.Element(f"{{{APP}}}Properties")
        if custom_root is None: custom_root = ET.Element(f"{{{CP}}}Properties")
        core_map = {"title": "dc:title", "subject": "dc:subject", "creator": "dc:creator", "lastModifiedBy": "cp:lastModifiedBy", "keywords": "cp:keywords", "category": "cp:category", "description": "dc:description", "contentStatus": "cp:contentStatus", "revision": "cp:revision", "created": "dcterms:created", "modified": "dcterms:modified"}
        app_map = {"application": "app:Application", "company": "app:Company", "manager": "app:Manager", "hyperlinkBase": "app:HyperlinkBase", "template": "app:Template", "lastPrinted": "app:LastPrinted", "totalTime": "app:TotalTime"}
        for key, tag in core_map.items():
            value = "" if clear else new_values.get(key, "")
            _set(core, tag, iso_store(value) if key in {"created", "modified"} else value)
        for key, tag in app_map.items():
            value = "" if clear else new_values.get(key, "")
            _set(app, tag, iso_store(value) if key == "lastPrinted" else value)
        existing = {prop.get("name"): prop for prop in custom_root.findall(f"{{{CP}}}property")}
        for name, prop in existing.items():
            if clear or name not in custom:
                custom_root.remove(prop)
            else:
                child = next(iter(prop), None)
                if child is None:
                    ET.SubElement(prop, f"{{{VT}}}lpwstr").text = custom[name]
                elif (child.text or "") != custom[name]:
                    # Keep untouched typed values and IDs; an edited value becomes text.
                    prop.remove(child)
                    ET.SubElement(prop, f"{{{VT}}}lpwstr").text = custom[name]
        used_pids = {int(prop.get("pid", "0")) for prop in custom_root.findall(f"{{{CP}}}property")}
        for name, value in custom.items() if not clear else ():
            if name in existing: continue
            pid = 2
            while pid in used_pids: pid += 1
            used_pids.add(pid)
            prop = ET.SubElement(custom_root, f"{{{CP}}}property", {"fmtid": "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}", "pid": str(pid), "name": name})
            ET.SubElement(prop, f"{{{VT}}}lpwstr").text = value
        # Read entries before closing the source archive. Windows won't replace a
        # path while its ZipFile handle is open.
        entries = [(item, original.read(item.filename)) for item in original.infolist()
                   if item.filename not in {"docProps/core.xml", "docProps/app.xml", "docProps/custom.xml"}]
        had_custom = "docProps/custom.xml" in original.namelist()
        if custom and not had_custom:
            content_type = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
            relation_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
            patched = []
            for item, contents in entries:
                if item.filename == "[Content_Types].xml":
                    root = ET.fromstring(contents)
                    namespace = "http://schemas.openxmlformats.org/package/2006/content-types"
                    ET.SubElement(root, f"{{{namespace}}}Override", {"PartName": "/docProps/custom.xml", "ContentType": content_type})
                    contents = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                elif item.filename == "_rels/.rels":
                    root = ET.fromstring(contents)
                    namespace = "http://schemas.openxmlformats.org/package/2006/relationships"
                    ids = {rel.get("Id") for rel in root}
                    index = 1
                    while f"rId{index}" in ids: index += 1
                    ET.SubElement(root, f"{{{namespace}}}Relationship", {"Id": f"rId{index}", "Type": relation_type, "Target": "docProps/custom.xml"})
                    contents = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                patched.append((item, contents))
            entries = patched
    descriptor, tmp_name = tempfile.mkstemp(suffix=meta.path.suffix, dir=meta.path.parent)
    os.close(descriptor)
    tmp = Path(tmp_name)
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
            for item, contents in entries: out.writestr(item, contents)
            out.writestr("docProps/core.xml", ET.tostring(core, encoding="utf-8", xml_declaration=True))
            out.writestr("docProps/app.xml", ET.tostring(app, encoding="utf-8", xml_declaration=True))
            if custom or had_custom: out.writestr("docProps/custom.xml", ET.tostring(custom_root, encoding="utf-8", xml_declaration=True))
        os.replace(tmp, meta.path)
    finally:
        if tmp.exists(): tmp.unlink(missing_ok=True)


def _pdf_unescape(value: bytes) -> str:
    return re.sub(rb"\\([\\()])", rb"\1", value).decode("utf-8", "replace")


def _pdf_escape(value: str) -> str:
    return value.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def read_pdf(path: Path) -> Metadata:
    from pypdf import PdfReader

    reader = PdfReader(path)
    if reader.is_encrypted:
        raise ValueError("加密 PDF 不支持读取或写入。")
    result = Metadata(path, "PDF")
    result.values.update(filesystem_values(path))
    mapping = {"title": "/Title", "author": "/Author", "subject": "/Subject", "keywords": "/Keywords", "creator": "/Creator", "producer": "/Producer", "created": "/CreationDate", "modified": "/ModDate"}
    for key, pdf_key in mapping.items():
        value = str((reader.metadata or {}).get(pdf_key, ""))
        result.values[key] = value.replace("D:", "") if key in {"created", "modified"} else value
    known_keys = set(mapping.values())
    result.custom.update({str(key): str(value) for key, value in (reader.metadata or {}).items()
                          if str(key) not in known_keys})
    # Word/WPS exports often retain their visible metadata in the XMP packet.
    try:
        xmp = reader.xmp_metadata
    except Exception:
        xmp = None
    if xmp is not None:
        title = (xmp.dc_title or {}).get("x-default") or next(iter((xmp.dc_title or {}).values()), "")
        subject = (xmp.dc_description or {}).get("x-default") or next(iter((xmp.dc_description or {}).values()), "")
        result.values["title"] = title or result.values["title"]
        result.values["author"] = ", ".join(xmp.dc_creator or []) or result.values["author"]
        result.values["subject"] = subject or result.values["subject"]
        result.values["keywords"] = ", ".join(xmp.dc_subject or []) or xmp.pdf_keywords or result.values["keywords"]
        result.values["creator"] = xmp.xmp_creator_tool or result.values["creator"]
        result.values["producer"] = xmp.pdf_producer or result.values["producer"]
    return result


def _pdf_date(value: str) -> str:
    digits = re.sub(r"\D", "", value)
    return "D:" + (digits[:14] if digits else dt.datetime.now().strftime("%Y%m%d%H%M%S"))


def write_pdf(meta: Metadata, values: dict[str, str], clear: bool, custom: dict[str, str] | None = None) -> None:
    from pypdf import PdfReader, PdfWriter
    from pypdf.generic import DictionaryObject
    from pypdf.xmp import XmpInformation

    # Rewriting a signed PDF invalidates its signature. Refuse it explicitly so
    # single-file and batch operations both report a safe, actionable result.
    if b"/ByteRange" in meta.path.read_bytes():
        raise ValueError("检测到数字签名 PDF，已跳过以避免签名失效。")
    reader = PdfReader(meta.path)
    if reader.is_encrypted:
        raise ValueError("加密 PDF 不支持写入。")
    writer = PdfWriter()
    writer.clone_document_from_reader(reader)

    pairs = {"/Title": "title", "/Author": "author", "/Subject": "subject", "/Keywords": "keywords", "/Creator": "creator", "/Producer": "producer", "/CreationDate": "created", "/ModDate": "modified"}
    # clone_document_from_reader keeps the source /Info object. Replacing it is
    # essential: add_metadata alone cannot remove old keys from that object.
    info = {} if clear else {str(key): str(value) for key, value in (reader.metadata or {}).items()}
    if not clear:
        requested = meta.custom if custom is None else custom
        for key in list(info):
            if key not in pairs and key not in requested: info.pop(key)
        for key, value in requested.items():
            if not key.startswith("/") or key in pairs:
                raise ValueError(f"PDF 自定义属性名称必须以 / 开头，且不能覆盖预设字段：{key}")
            info[key] = value
    for pdf_key, key in pairs.items():
        value = "" if clear else values.get(key, "")
        if value:
            info[pdf_key] = _pdf_date(value) if key in {"created", "modified"} else value
        else:
            info.pop(pdf_key, None)
    writer._info = writer._add_object(DictionaryObject())
    writer.add_metadata(info)

    if clear:
        # Remove the entire XMP packet so unrecognized or vendor-specific fields
        # cannot survive a "clear all properties" operation.
        writer.xmp_metadata = None
    else:
        try:
            xmp = reader.xmp_metadata or XmpInformation.create()
        except Exception:
            xmp = XmpInformation.create()
        xmp.dc_title = {"x-default": values.get("title", "")}
        xmp.dc_creator = [values["author"]] if values.get("author") else []
        xmp.dc_description = {"x-default": values.get("subject", "")}
        xmp.dc_subject = [item.strip() for item in values.get("keywords", "").split(",") if item.strip()]
        xmp.pdf_keywords = values.get("keywords", "")
        xmp.xmp_creator_tool = values.get("creator", "")
        xmp.pdf_producer = values.get("producer", "")
        writer.xmp_metadata = xmp

    descriptor, tmp_name = tempfile.mkstemp(suffix=meta.path.suffix, dir=meta.path.parent)
    os.close(descriptor)
    tmp = Path(tmp_name)
    try:
        with tmp.open("wb") as output:
            writer.write(output)
        os.replace(tmp, meta.path)
    finally:
        if tmp.exists():
            tmp.unlink(missing_ok=True)


def clear_embedded_properties(meta: Metadata) -> None:
    """Clear document-internal metadata without changing filesystem timestamps."""
    if meta.kind == "Office":
        write_office(meta, {key: "" for key, _ in FIELDS}, {}, True)
    elif meta.kind == "Legacy Office":
        write_legacy(meta, {}, True)
    elif meta.kind == "PDF":
        write_pdf(meta, {}, True)
    elif meta.kind == "Markdown":
        write_markdown(meta, {}, True)
    elif meta.kind == "Image":
        write_image(meta, {}, True)
    elif meta.kind == "Archive":
        write_zip(meta, {}, True)
    elif meta.kind == "Audio":
        write_audio(meta, {}, True)
    elif meta.kind == "Video":
        write_video(meta, {}, True)
    else:
        raise ValueError("不支持清除该文件类型的属性。")


def _legacy_helper_jar() -> Path:
    package_root = Path(__file__).resolve().parent.parent
    bundled_root = Path(getattr(sys, "_MEIPASS", package_root))
    candidates = [
        bundled_root / "legacy_helper" / "legacy-metadata-helper.jar",
        package_root / "legacy_helper" / "legacy-metadata-helper.jar",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    raise RuntimeError(
        "旧版 Office 兼容组件尚未安装。请构建 legacy_helper 并将 "
        "legacy-metadata-helper.jar 放入 legacy_helper 文件夹。"
    )


def _java_executable() -> str:
    """Resolve Java without requiring a system-wide installation."""
    configured = os.environ.get("FILE_METADATA_JAVA")
    if configured:
        return configured
    package_root = Path(__file__).resolve().parent.parent
    bundled_root = Path(getattr(sys, "_MEIPASS", package_root))
    bundled_java = bundled_root / "runtime" / "bin" / ("java.exe" if os.name == "nt" else "java")
    if bundled_java.exists():
        return str(bundled_java)
    return "java"


def _run_legacy_helper(*args: str, values: dict[str, str] | None = None) -> str:
    java = _java_executable()
    payload = ""
    if values is not None:
        payload = "\n".join(
            f"{key}\t{base64.b64encode(value.encode('utf-8')).decode('ascii')}"
            for key, value in values.items()
        ) + "\n"
    try:
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0
        completed = subprocess.run(
            [java, "-jar", str(_legacy_helper_jar()), *args],
            input=payload, text=True, encoding="utf-8", capture_output=True, check=False,
            creationflags=creationflags,
        )
    except FileNotFoundError as exc:
        raise RuntimeError("未找到 Java 运行环境，无法处理旧版 .doc/.xls/.ppt 文件。") from exc
    if completed.returncode != 0:
        detail = (completed.stderr or completed.stdout).strip()
        raise RuntimeError(f"旧版 Office 属性处理失败：{detail or 'Java 辅助程序异常'}")
    return completed.stdout


def read_legacy(path: Path) -> Metadata:
    result = Metadata(path, "Legacy Office")
    result.values.update(filesystem_values(path))
    output = _run_legacy_helper("read", str(path))
    for line in output.splitlines():
        try:
            key, encoded = line.split("\t", 1)
            value = base64.b64decode(encoded).decode("utf-8")
            if key.startswith("custom:"):
                name = base64.urlsafe_b64decode(key[7:] + "=" * (-len(key[7:]) % 4)).decode("utf-8")
                result.custom[name] = value
            else:
                result.values[key] = value
        except (ValueError, UnicodeDecodeError):
            continue
    return result


def write_legacy(meta: Metadata, values: dict[str, str], clear: bool, custom: dict[str, str] | None = None) -> None:
    descriptor, tmp_name = tempfile.mkstemp(suffix=meta.path.suffix, dir=meta.path.parent)
    os.close(descriptor)
    tmp = Path(tmp_name)
    try:
        payload = dict(values)
        for name, value in (meta.custom if custom is None else custom).items():
            encoded = base64.urlsafe_b64encode(name.encode("utf-8")).decode("ascii").rstrip("=")
            payload[f"custom:{encoded}"] = value
        _run_legacy_helper("write", str(meta.path), str(tmp), str(clear).lower(), values=payload)
        os.replace(tmp, meta.path)
    finally:
        if tmp.exists():
            tmp.unlink(missing_ok=True)


def load(path: str) -> Metadata:
    item = Path(path)
    suffix = item.suffix.lower()
    if suffix in OFFICE_SUFFIXES: return read_office(item)
    if suffix == ".pdf": return read_pdf(item)
    if suffix in LEGACY_SUFFIXES: return read_legacy(item)
    if suffix in MARKDOWN_SUFFIXES: return read_markdown(item)
    if suffix in IMAGE_SUFFIXES: return read_image(item)
    if suffix in ZIP_SUFFIXES: return read_zip(item)
    if suffix in AUDIO_SUFFIXES: return read_audio(item)
    if suffix in VIDEO_SUFFIXES: return read_video(item)
    raise ValueError("仅支持 Office、PDF 和 Markdown 文件。")



def verify_embedded_clear(path: str | Path) -> tuple[Metadata, list[str]]:
    """Re-read a file and report non-empty embedded properties after clearing.

    Filesystem timestamps and structural values such as archive entry counts are
    intentionally excluded; they are not embedded descriptive metadata.
    """
    result = load(str(path))
    ignored = {"filesystem_created", "filesystem_modified", "file_count"}
    labels = {
        key: label
        for fields in (FIELDS, PDF_FIELDS, MARKDOWN_FIELDS, IMAGE_FIELDS, ZIP_FIELDS, MEDIA_FIELDS)
        for key, label in fields
    }
    labels["_all_tags"] = "音频/视频容器、流、章节或标签"
    labels["_all_exif"] = "其他 EXIF 属性"
    labels["_extra_image_metadata"] = "图片 XMP/IPTC/注释"
    remaining = [
        labels.get(key, key)
        for key, value in result.values.items()
        if key not in ignored and str(value).strip()
        and not (result.kind == "Legacy Office" and key == "totalTime" and str(value).strip() == "0")
    ]
    remaining.extend(
        f"自定义属性：{name}"
        for name, value in result.custom.items()
        if str(value).strip()
    )
    return result, remaining

def backup(path: Path) -> Path:
    # Never overwrite an earlier backup: network-sync folders and document apps
    # often keep the old .metadata-backup file open or read-only.
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    # Keep the original extension so Windows recognizes the backup as a PDF or
    # Office document and it can be opened directly.
    target = path.with_name(f"{path.stem}.{stamp}.metadata-backup{path.suffix}")
    try:
        shutil.copy2(path, target)
        make_writable(target)
    except PermissionError as exc:
        raise PermissionError(
            "无法在当前文件夹创建备份。请确认该文件夹可写、文件未被同步软件或 Office/WPS 占用，"
            "或者在界面中将“保存前备份”设为“否”。"
        ) from exc
    return target
