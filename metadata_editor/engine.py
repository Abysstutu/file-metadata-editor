from __future__ import annotations

import datetime as dt
import ctypes
import os
import re
import shutil
import stat
import tempfile
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET

from pypdf import PdfReader, PdfWriter
from pypdf.generic import DictionaryObject
from pypdf.xmp import XmpInformation

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

FIELDS = [
    ("title", "标题"), ("subject", "主题"), ("creator", "作者"), ("lastModifiedBy", "最后保存者"),
    ("keywords", "标记 / 关键词"), ("category", "类别"), ("description", "备注"), ("contentStatus", "状态"),
    ("revision", "修订号"), ("created", "创建内容时间"), ("modified", "最后一次保存时间"),
    ("lastPrinted", "最后打印时间"), ("totalTime", "总编辑时长（分钟）"), ("application", "程序名称"),
    ("company", "公司"), ("manager", "管理者"), ("hyperlinkBase", "超链接基址"), ("template", "文档模板"),
]
PDF_FIELDS = [("title", "标题"), ("author", "作者"), ("subject", "主题"), ("keywords", "关键词"), ("creator", "Creator"), ("producer", "Producer"), ("created", "创建内容时间"), ("modified", "修改内容时间")]


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
        for child in list(custom_root): custom_root.remove(child)
        if not clear:
            for pid, (name, value) in enumerate(custom.items(), start=2):
                prop = ET.SubElement(custom_root, f"{{{CP}}}property", {"fmtid": "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}", "pid": str(pid), "name": name})
                ET.SubElement(prop, f"{{{VT}}}lpwstr").text = value
        # Read entries before closing the source archive. Windows won't replace a
        # path while its ZipFile handle is open.
        entries = [(item, original.read(item.filename)) for item in original.infolist()
                   if item.filename not in {"docProps/core.xml", "docProps/app.xml", "docProps/custom.xml"}]
        had_custom = "docProps/custom.xml" in original.namelist()
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
    reader = PdfReader(path)
    if reader.is_encrypted:
        raise ValueError("加密 PDF 不支持读取或写入。")
    result = Metadata(path, "PDF")
    result.values.update(filesystem_values(path))
    mapping = {"title": "/Title", "author": "/Author", "subject": "/Subject", "keywords": "/Keywords", "creator": "/Creator", "producer": "/Producer", "created": "/CreationDate", "modified": "/ModDate"}
    for key, pdf_key in mapping.items():
        value = str((reader.metadata or {}).get(pdf_key, ""))
        result.values[key] = value.replace("D:", "") if key in {"created", "modified"} else value
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


def write_pdf(meta: Metadata, values: dict[str, str], clear: bool) -> None:
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
    for pdf_key, key in pairs.items():
        value = "" if clear else values.get(key, "")
        if value:
            info[pdf_key] = _pdf_date(value) if key in {"created", "modified"} else value
        else:
            info.pop(pdf_key, None)
    writer._info = writer._add_object(DictionaryObject())
    writer.add_metadata(info)

    try:
        xmp = reader.xmp_metadata or XmpInformation.create()
    except Exception:
        xmp = XmpInformation.create()
    if clear:
        xmp.dc_title, xmp.dc_creator, xmp.dc_description, xmp.dc_subject = {}, [], {}, []
        xmp.pdf_keywords = xmp.xmp_creator_tool = xmp.pdf_producer = ""
    else:
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
    elif meta.kind == "PDF":
        write_pdf(meta, {}, True)
    else:
        raise ValueError("不支持清除该文件类型的属性。")


def load(path: str) -> Metadata:
    item = Path(path)
    suffix = item.suffix.lower()
    if suffix in OFFICE_SUFFIXES: return read_office(item)
    if suffix == ".pdf": return read_pdf(item)
    if suffix in LEGACY_SUFFIXES: raise ValueError("旧版 Office 二进制文件暂不支持。请在 Word、Excel 或 WPS 中另存为 DOCX/XLSX/PPTX 后再编辑。")
    raise ValueError("仅支持新版 Office 文档和 PDF。")


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
