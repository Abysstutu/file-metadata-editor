"""Shared field selection and write rules for previews and batch operations."""

from __future__ import annotations

from .engine import (
    FIELDS, PDF_FIELDS, MARKDOWN_FIELDS, IMAGE_FIELDS, ZIP_FIELDS, MEDIA_FIELDS,
    Metadata, load, write_office, write_legacy, write_pdf, write_markdown,
    write_image, write_zip, write_audio, write_video,
)

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
