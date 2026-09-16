from __future__ import annotations

import datetime as dt
import ctypes
import os
import re
import shutil
import tempfile
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET

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
    data = path.read_bytes()
    if b"/Encrypt" in data: raise ValueError("加密 PDF 不支持读取或写入。")
    result = Metadata(path, "PDF")
    result.values.update(filesystem_values(path))
    mapping = {"title": b"Title", "author": b"Author", "subject": b"Subject", "keywords": b"Keywords", "creator": b"Creator", "producer": b"Producer", "created": b"CreationDate", "modified": b"ModDate"}
    for key, pdf_key in mapping.items():
        matches = list(re.finditer(rb"/" + pdf_key + rb"\s*\((.*?)\)", data, re.S))
        value = _pdf_unescape(matches[-1].group(1)) if matches else ""
        result.values[key] = value.replace("D:", "") if key in {"created", "modified"} else value
    return result


def _pdf_date(value: str) -> str:
    digits = re.sub(r"\D", "", value)
    return "D:" + (digits[:14] if digits else dt.datetime.now().strftime("%Y%m%d%H%M%S"))


def write_pdf(meta: Metadata, values: dict[str, str], clear: bool) -> None:
    data = meta.path.read_bytes()
    if b"/Encrypt" in data: raise ValueError("加密 PDF 不支持写入。")
    starts = list(re.finditer(rb"startxref\s*(\d+)", data))
    trailers = list(re.finditer(rb"trailer\s*<<(.*?)>>", data, re.S))
    if not starts or not trailers: raise ValueError("此 PDF 结构不支持安全的增量属性写入。")
    trailer, previous = trailers[-1].group(1), int(starts[-1].group(1))
    root = re.search(rb"/Root\s+(\d+\s+\d+\s+R)", trailer)
    size = re.search(rb"/Size\s+(\d+)", trailer)
    if not root or not size: raise ValueError("未找到 PDF 根对象，无法写入。")
    obj = int(size.group(1))
    pairs = [("Title", "title"), ("Author", "author"), ("Subject", "subject"), ("Keywords", "keywords"), ("Creator", "creator"), ("Producer", "producer"), ("CreationDate", "created"), ("ModDate", "modified")]
    body = [f"{obj} 0 obj", "<<"]
    for pdf_key, key in pairs:
        value = "" if clear else values.get(key, "")
        if value:
            value = _pdf_date(value) if key in {"created", "modified"} else value
            body.append(f"/{pdf_key} ({_pdf_escape(value)})")
    body += [">>", "endobj"]
    appendix = ("\n".join(body) + "\n").encode("utf-8")
    offset = len(data) + len(appendix)
    xref = f"xref\n{obj} 1\n{len(data):010d} 00000 n \ntrailer\n<< /Size {obj + 1} /Root {root.group(1).decode()} /Info {obj} 0 R /Prev {previous} >>\nstartxref\n{offset}\n%%EOF\n".encode()
    meta.path.write_bytes(data + appendix + xref)


def load(path: str) -> Metadata:
    item = Path(path)
    suffix = item.suffix.lower()
    if suffix in OFFICE_SUFFIXES: return read_office(item)
    if suffix == ".pdf": return read_pdf(item)
    if suffix in LEGACY_SUFFIXES: raise ValueError("旧版 Office 二进制文件暂不支持。请在 Word、Excel 或 WPS 中另存为 DOCX/XLSX/PPTX 后再编辑。")
    raise ValueError("仅支持新版 Office 文档和 PDF。")


def backup(path: Path) -> Path:
    target = path.with_name(path.name + ".metadata-backup")
    shutil.copy2(path, target)
    return target
