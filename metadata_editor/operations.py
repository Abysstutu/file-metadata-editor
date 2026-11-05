"""Shared field selection and write rules for previews and batch operations."""

from __future__ import annotations

import os
import shutil
import tempfile
from pathlib import Path

from .engine import (
    FIELDS, PDF_FIELDS, MARKDOWN_FIELDS, IMAGE_FIELDS, ZIP_FIELDS, MEDIA_FIELDS,
    Metadata, load, write_office, write_legacy, write_pdf, write_markdown,
    write_image, write_zip, write_audio, write_video, make_writable,
)


def output_copy_path(source: Path, directory: str, template: str) -> Path:
    """Resolve a new output path without overwriting a source or earlier copy."""
    folder = Path(directory).expanduser() if directory.strip() else source.parent
    if not folder.is_dir():
        raise ValueError(f"副本输出文件夹不存在：{folder}")
    if not template.strip() or any(char in template for char in "\\/:*?\"<>|"):
        raise ValueError("命名规则不能为空，且不能包含路径或 Windows 禁用字符。")
    literal = template.replace("{name}", "")
    if "{" in literal or "}" in literal:
        raise ValueError("命名规则只支持 {name} 占位符。")
    stem = template.replace("{name}", source.stem).strip().rstrip(". ")
    if not stem or stem in {".", ".."} or any(char in stem for char in "\\/:*?\"<>|"):
        raise ValueError("命名规则没有生成有效文件名。")
    candidate = folder / f"{stem}{source.suffix}"
    index = 2
    while candidate.exists() or candidate.resolve() == source.resolve():
        candidate = folder / f"{stem} ({index}){source.suffix}"
        index += 1
    return candidate


def create_output_copy(source: Path, directory: str, template: str) -> Path:
    while True:
        target = output_copy_path(source, directory, template)
        try:
            with source.open("rb") as original, target.open("xb") as output:
                shutil.copyfileobj(original, output)
            break
        except FileExistsError:
            continue
        except Exception:
            target.unlink(missing_ok=True)
            raise
    shutil.copystat(source, target)
    make_writable(target)
    return target


def restore_from_backup(target: Path, saved: Path) -> None:
    """Replace target using a temporary copy; keep the selected backup intact."""
    if target.suffix.lower() != saved.suffix.lower() or not saved.is_file():
        raise ValueError("备份文件不存在或类型与目标文件不同。")
    descriptor, tmp_name = tempfile.mkstemp(suffix=target.suffix, dir=target.parent)
    os.close(descriptor)
    temporary = Path(tmp_name)
    try:
        shutil.copy2(saved, temporary)
        make_writable(temporary)
        os.replace(temporary, target)
    finally:
        temporary.unlink(missing_ok=True)

FIELD_SETS = {
    "Office": FIELDS, "Legacy Office": FIELDS, "PDF": PDF_FIELDS,
    "Markdown": MARKDOWN_FIELDS, "Image": IMAGE_FIELDS,
    "Archive": ZIP_FIELDS, "Audio": MEDIA_FIELDS, "Video": MEDIA_FIELDS,
}
READ_ONLY = {"Archive": {"file_count"}, "Video": {"album", "genre"}}
CATEGORIES = {
    "身份与来源": {"author", "creator", "lastModifiedBy", "application", "company", "manager", "producer"},
    "标题与描述": {"title", "subject", "keywords", "category", "description", "contentStatus", "album", "genre", "tags", "comment"},
    "内部时间": {"created", "modified", "lastPrinted", "date", "totalTime"},
    "位置与设备": {"gps", "template", "hyperlinkBase"},
    "其他属性": {"revision"},
}


def editable_fields(meta: Metadata) -> list[tuple[str, str]]:
    return [(key, label) for key, label in FIELD_SETS[meta.kind]
            if key not in READ_ONLY.get(meta.kind, set())]


def selected_fields(meta: Metadata, categories: set[str]) -> list[tuple[str, str]]:
    keys = set().union(*(CATEGORIES[name] for name in categories)) if categories else set()
    return [(key, label) for key, label in editable_fields(meta) if key in keys]


def changed_values(before: Metadata, after: Metadata, keys: set[str]) -> list[tuple[str, str, str]]:
    labels = dict(FIELD_SETS[before.kind])
    changes = [(labels.get(key, key), before.values.get(key, ""), after.values.get(key, ""))
               for key in keys if before.values.get(key, "") != after.values.get(key, "")]
    if before.kind == "Office" and "__custom__" in keys:
        for key in sorted(set(before.custom) | set(after.custom)):
            old, new = before.custom.get(key, ""), after.custom.get(key, "")
            if old != new:
                changes.append((f"自定义：{key}", old, new))
    if "__all__" in keys:
        for key in ("_all_exif", "_all_tags", "_extra_image_metadata"):
            old, new = before.values.get(key, ""), after.values.get(key, "")
            if old != new:
                changes.append(("其他嵌入属性", old, new))
    return changes


def write_selected(meta: Metadata, updates: dict[str, str]) -> Metadata:
    """Write selected embedded fields, keeping other supported fields intact."""
    allowed = {key for key, _ in editable_fields(meta)}
    invalid = set(updates) - allowed
    if invalid:
        raise ValueError(f"不支持修改这些属性：{', '.join(sorted(invalid))}")
    values = {key: meta.values.get(key, "") for key, _ in editable_fields(meta)}
    values.update(updates)
    if meta.kind == "Office":
        write_office(meta, values, meta.custom, False)
    elif meta.kind == "Legacy Office":
        write_legacy(meta, values, False)
    elif meta.kind == "PDF":
        write_pdf(meta, values, False)
    elif meta.kind == "Markdown":
        write_markdown(meta, values)
    elif meta.kind == "Image":
        write_image(meta, values)
    elif meta.kind == "Archive":
        write_zip(meta, values)
    elif meta.kind == "Audio":
        write_audio(meta, values)
    elif meta.kind == "Video":
        write_video(meta, values)
    else:
        raise ValueError("不支持此文件类型。")
    return load(str(meta.path))
