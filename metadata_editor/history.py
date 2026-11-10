"""Append-only history of completed metadata edits; never stores file contents."""

from __future__ import annotations

import datetime as dt
import hashlib
import hmac
import json
import os
import re
from pathlib import Path

from .engine import (Metadata, inplace_write_block_reason, load, set_filesystem_times,
                     write_legacy, write_markdown, write_office, write_pdf)
from .operations import editable_fields, write_selected


class FileChangedError(ValueError):
    """The target file no longer has the hash saved with the history entry."""


def history_path() -> Path:
    base = os.environ.get("LOCALAPPDATA")
    folder = Path(base) / "FileMetadataEditor" if base else Path.home() / ".file-metadata-editor"
    return folder / "edit-history.jsonl"


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read_history(path: Path | None = None) -> tuple[list[dict], int]:
    """Read valid records, skipping incomplete or damaged JSONL lines."""
    target = path or history_path()
    if not target.exists():
        return [], 0
    records: list[dict] = []
    skipped = 0
    with target.open("r", encoding="utf-8-sig", errors="replace") as stream:
        for line in stream:
            if not line.strip():
                continue
            try:
                record = json.loads(line)
            except json.JSONDecodeError:
                skipped += 1
                continue
            if not isinstance(record, dict) or not record.get("result_path"):
                skipped += 1
                continue
            records.append(record)
    return records, skipped


def restore_recorded_metadata(record: dict) -> tuple[Metadata, Metadata, bool]:
    """Restore pre-clear metadata only when the cleared file hash still matches."""
    if record.get("operation") not in {"clear", "batch_clear"}:
        raise ValueError("只有清除记录可以恢复到清除前状态。")
    path_text = record.get("result_path")
    expected_hash = record.get("sha256")
    snapshot = record.get("metadata_before")
    if not isinstance(path_text, str) or not path_text or not isinstance(expected_hash, str) or not re.fullmatch(r"[0-9a-fA-F]{64}", expected_hash):
        raise ValueError("这条记录缺少有效的文件路径或 SHA-256，无法恢复。")
    if not isinstance(snapshot, dict) or not isinstance(snapshot.get("values"), dict) or not isinstance(snapshot.get("custom"), dict):
        raise ValueError("这条旧记录没有保存操作前的元数据，无法恢复。请使用更新后的程序生成新记录。")
    path = Path(path_text)
    if not path.is_file():
        raise FileNotFoundError(f"记录中的文件不存在：{path}")
    if not hmac.compare_digest(file_sha256(path).lower(), expected_hash.lower()):
        raise FileChangedError(f"文件已变动，当前 SHA-256 与历史记录不一致：\n{path}")

    current = load(str(path))
    recorded_kind = record.get("file_type")
    if recorded_kind and recorded_kind != current.kind:
        raise ValueError("记录中的文件类型与当前文件不一致。")
    saved_values = snapshot["values"]
    saved_custom = snapshot["custom"]
    if not all(isinstance(key, str) and isinstance(value, str) for key, value in saved_values.items()):
        raise ValueError("记录中的预设属性格式无效。")
    if not all(isinstance(key, str) and isinstance(value, str) for key, value in saved_custom.items()):
        raise ValueError("记录中的自定义属性格式无效。")

    updates = {key: saved_values[key] for key, _ in editable_fields(current) if key in saved_values}
    supports_custom = current.kind in {"Office", "Legacy Office", "PDF", "Markdown"}
    custom_changed = supports_custom and current.custom != saved_custom
    embedded_changed = custom_changed or any(current.values.get(key, "") != value for key, value in updates.items())
    created = saved_values.get("filesystem_created", "")
    modified = saved_values.get("filesystem_modified", "")
    times_changed = any(saved_values.get(key) is not None and current.values.get(key, "") != saved_values[key]
                        for key in ("filesystem_created", "filesystem_modified"))
    if not embedded_changed and not times_changed:
        return current, current, False

    reason = inplace_write_block_reason(path)
    if reason:
        raise PermissionError(f"无法写回记录中的文件：{reason}")
    # A second check narrows the window in which another process could change it.
    if not hmac.compare_digest(file_sha256(path).lower(), expected_hash.lower()):
        raise FileChangedError(f"文件已变动，当前 SHA-256 与历史记录不一致：\n{path}")
    if embedded_changed:
        if supports_custom:
            values = {key: current.values.get(key, "") for key, _ in editable_fields(current)}
            values.update(updates)
            if current.kind == "Office":
                write_office(current, values, saved_custom, False)
            elif current.kind == "Legacy Office":
                write_legacy(current, values, False, saved_custom)
            elif current.kind == "PDF":
                write_pdf(current, values, False, saved_custom)
            else:
                write_markdown(current, values, custom=saved_custom)
        else:
            write_selected(current, updates)
    if (times_changed or embedded_changed) and (created or modified):
        set_filesystem_times(path, created, modified)
    after = load(str(path))
    remaining = [key for key, value in updates.items() if after.values.get(key, "") != value]
    if supports_custom and after.custom != saved_custom:
        remaining.append("自定义属性")
    remaining.extend(key for key in ("filesystem_created", "filesystem_modified")
                     if saved_values.get(key) and after.values.get(key) != saved_values[key])
    if remaining:
        raise RuntimeError("写回已执行，但以下属性未通过读回验证：" + "、".join(remaining[:10]))
    return current, after, True


def append_history(source: Path, result: Metadata | Path, operation: str, *,
                   before: Metadata, verification: str = "read_back",
                   destination: Path | None = None) -> Path:
    """Record pre-edit metadata and post-edit hash in one local JSONL file."""
    target = destination or history_path()
    result_path = result.path if isinstance(result, Metadata) else result
    record = {
        "time_utc": dt.datetime.now(dt.timezone.utc).isoformat(),
        "operation": operation,
        "source_name": source.name,
        "source_path": str(source.resolve()),
        "result_name": result_path.name,
        "result_path": str(result_path.resolve()),
        "sha256": file_sha256(result_path),
        "file_type": result.kind if isinstance(result, Metadata) else None,
        "verification": verification,
        "metadata_before": {"values": dict(before.values), "custom": dict(before.custom),
                            "readonly": dict(before.readonly)},
        "metadata_after": ({"values": result.values, "custom": result.custom,
                            "readonly": result.readonly} if isinstance(result, Metadata) else None),
    }
    target.parent.mkdir(parents=True, exist_ok=True)
    with target.open("a", encoding="utf-8", newline="\n") as stream:
        stream.write(json.dumps(record, ensure_ascii=False, sort_keys=True) + "\n")
        stream.flush()
        os.fsync(stream.fileno())
    return target
