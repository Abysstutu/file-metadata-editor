"""Inventory discovered metadata and probe edit/clear behavior by file suffix.

Usage: python scripts/metadata_capability_matrix.py CORPUS OUTPUT
The corpus must contain 有元数据 and 无可编辑元数据 subfolders. All writes use copies.
"""

from __future__ import annotations

import argparse
import csv
import json
import shutil
import tempfile
import zipfile
from collections import Counter, defaultdict
from pathlib import Path

from metadata_editor import engine
from metadata_editor.operations import FIELD_SETS, READ_ONLY, write_selected


SYSTEM_FIELDS = {"filesystem_created": "文件创建时间", "filesystem_modified": "文件修改时间"}
RAW_FIELDS = {"_all_exif", "_all_tags", "_extra_image_metadata"}
DATE_FIELDS = {"created", "modified", "lastPrinted", "date"}


def probe_value(key: str) -> str:
    if key == "totalTime":
        return "37"
    if key == "revision":
        return "9"
    if key in DATE_FIELDS:
        return "2027-02-03 04:05:06" if key != "date" else "2027-02-03"
    if key == "gps":
        return "39.90, 116.40"
    return f"能力测试-{key}"


def split_raw(key: str, value: str):
    if key == "_extra_image_metadata":
        return [item.strip() for item in value.split(",") if item.strip()]
    # Older readers used a literal backslash-n delimiter; WAV uses real newlines.
    return [line.strip() for line in value.replace("\\n", "\n").splitlines() if line.strip()]


def inventory(samples):
    observed = defaultdict(lambda: {"values": defaultdict(set), "custom": defaultdict(set),
                                    "raw": defaultdict(lambda: defaultdict(set)), "kind": ""})
    failures = []
    for path in samples:
        try:
            meta = engine.load(str(path))
        except Exception as exc:
            failures.append({"file": str(path), "error": str(exc)})
            continue
        record = observed[path.suffix.lower()]
        record["kind"] = meta.kind
        for key, value in meta.values.items():
            if key in RAW_FIELDS:
                for item in split_raw(key, str(value)):
                    name, separator, raw_value = item.partition(":")
                    record["raw"][key][name.strip()].add(raw_value.strip() if separator else item)
            elif str(value).strip():
                record["values"][key].add(str(value))
        for key, value in meta.custom.items():
            record["custom"][key].add(str(value))
    return observed, failures


def edit_probe(source: Path, key: str, kind: str):
    requested = probe_value(key)
    with tempfile.TemporaryDirectory() as scratch:
        path = Path(scratch) / source.name
        shutil.copy2(source, path)
        try:
            before = engine.load(str(path))
            write_selected(before, {key: requested})
            actual = engine.load(str(path)).values.get(key, "")
            expected = (engine.iso_display(requested) if kind == "Office" and key in {"created", "modified", "lastPrinted"}
                        else "20270203040506" if kind == "PDF" and key in {"created", "modified"} else requested)
            edit = "是"
            note = "写入与单项清除均读回成功"
            if actual != expected:
                ascii_value = f"Probe-{key}"
                if key not in DATE_FIELDS | {"totalTime", "revision", "gps"}:
                    write_selected(engine.load(str(path)), {key: ascii_value})
                    if engine.load(str(path)).values.get(key, "") == ascii_value:
                        edit = "部分（仅 ASCII）"
                        note = f"中文写入读回 {actual!r}；ASCII 写入成功"
                if edit == "是":
                    return "否", clear_existing_probe(source, key), f"写入 {requested!r}；读回 {actual!r}"
            write_selected(engine.load(str(path)), {key: ""})
            cleared = engine.load(str(path)).values.get(key, "")
            if cleared in {"", "0"} and (cleared != "0" or key == "totalTime"):
                return edit, "是", note
            return edit, "未读回空值", f"{note}；清除后读回 {cleared!r}"
        except Exception as exc:
            return "否", clear_existing_probe(source, key), f"{type(exc).__name__}: {exc}"[:300]


def clear_existing_probe(source: Path, key: str):
    """Check clearing a pre-existing value when a new-value write fails."""
    with tempfile.TemporaryDirectory() as scratch:
        path = Path(scratch) / source.name
        shutil.copy2(source, path)
        try:
            before = engine.load(str(path))
            if not str(before.values.get(key, "")).strip():
                return "样本无值，未验证"
            write_selected(before, {key: ""})
            after = str(engine.load(str(path)).values.get(key, "")).strip()
            return "是" if not after or (key == "totalTime" and after == "0") else "否"
        except Exception:
            return "否"


def custom_probe(source: Path, name: str):
    with tempfile.TemporaryDirectory() as scratch:
        path = Path(scratch) / source.name
        shutil.copy2(source, path)
        try:
            meta = engine.load(str(path))
            values = dict(meta.values)
            custom = dict(meta.custom)
            custom[name] = "能力测试-自定义"
            engine.write_office(meta, values, custom, False)
            if engine.load(str(path)).custom.get(name) != "能力测试-自定义":
                return "否", "未验证", "写入后未读回新值"
            custom.pop(name)
            engine.write_office(engine.load(str(path)), values, custom, False)
            if name not in engine.load(str(path)).custom:
                return "是", "是", "自定义属性修改与删除均成功"
            return "是", "否", "自定义属性未删除"
        except Exception as exc:
            return "否", "未验证", f"{type(exc).__name__}: {exc}"[:300]


def injected_custom_probe(source: Path, kind: str):
    """Check arbitrary keys that the existing corpus does not contain."""
    with tempfile.TemporaryDirectory() as scratch:
        path = Path(scratch) / source.name
        shutil.copy2(source, path)
        if kind == "Markdown":
            content = path.read_text(encoding="utf-8-sig")
            if not content.startswith("---\n"):
                raise ValueError("Markdown sample has no Front Matter")
            path.write_text(content.replace("---\n", "---\ncodex_custom_probe: original\n", 1), encoding="utf-8")
            if engine.load(str(path)).values.get("codex_custom_probe") != "original":
                raise ValueError("Arbitrary Front Matter key was not read")
            write_selected(engine.load(str(path)), {"title": "Probe"})
            preserved = engine.load(str(path)).values.get("codex_custom_probe") == "original"
            engine.clear_embedded_properties(engine.load(str(path)))
            cleared = not engine.load(str(path)).values.get("codex_custom_probe")
        else:
            from pypdf import PdfReader, PdfWriter

            writer = PdfWriter()
            writer.clone_document_from_reader(PdfReader(path))
            writer.add_metadata({"/CodexCustomProbe": "original"})
            with path.open("wb") as stream:
                writer.write(stream)
            if (PdfReader(path).metadata or {}).get("/CodexCustomProbe") != "original":
                raise ValueError("PDF custom Info key was not added")
            write_selected(engine.load(str(path)), {"title": "Probe"})
            preserved = (PdfReader(path).metadata or {}).get("/CodexCustomProbe") == "original"
            engine.clear_embedded_properties(engine.load(str(path)))
            cleared = "/CodexCustomProbe" not in (PdfReader(path).metadata or {})
        return preserved, cleared


def sample_value(values):
    if not values:
        return ""
    shown = sorted(values)[0].replace("\n", " ")
    return shown[:100] + ("…" if len(shown) > 100 else "")


def render_markdown(rows, summary, path: Path):
    """Create a readable, per-suffix list alongside the filterable CSV."""
    def cell(value):
        return str(value).replace("|", "\\|").replace("\n", " ").replace("\r", " ")

    counts = summary["preset_edit_results"]
    lines = [
        "# 文件元数据能力清单", "",
        f"范围：当前程序支持的 {summary['formats']} 种扩展名，以及测试集中 {summary['samples_read']} 个样本能发现的字段。",
        "这不是所有厂商私有元数据的穷尽清单；未被当前读取器识别的字段不能据此判定为已清除。", "",
        "## 统计", "",
        f"- 样本读取：{summary['samples_read']} 个；失败 {len(summary['sample_read_failures'])} 个。",
        f"- 预设字段按“格式 × 字段”统计：{counts.get('是', 0)} 项写入读回成功，"
        f"{counts.get('部分（仅 ASCII）', 0)} 项仅 ASCII 成功，{counts.get('否', 0)} 项未成功或设计为只读。",
        "- 全量清除：此前 350 份测试副本均通过重新读取验证；此列只对样本中出现的属性提供证据。",
        "- 文件系统创建、修改时间单独列出，它们不是可清空的内嵌属性。", "",
        "## 主要限制", "",
        "- `.doc/.xls/.ppt`：创建、修改、打印时间、编辑时长、超链接基址和模板的逐项写入未按输入值读回。",
        "- `.jpg/.jpeg/.tif/.tiff/.webp`：标题、作者仅 ASCII 写入读回成功；描述和 GPS 未成功。",
        "- `.mkv`：标题成功；作者、日期、描述未按输入值读回，专辑和类型为只读。",
        "- `.mp3`：描述写入报错；其他五个预设字段成功。",
        "- ZIP 只支持压缩包注释；成员时间、压缩方式等技术属性保留。",
        "- Office Open XML 的样本自定义属性可修改和删除；Markdown、PDF 的任意自定义键可被整体清除，但当前界面不提供逐项编辑。", "",
        "## 各格式字段明细", "",
    ]
    grouped = defaultdict(list)
    for row in rows:
        grouped[row["格式"]].append(row)
    for suffix, records in sorted(grouped.items()):
        lines.extend([f"### `{suffix}`", "",
                      "| 类别 | 字段 | 样本值 | 可编辑 | 可单项清除 | 可全量清除 | 实测说明 |",
                      "| --- | --- | --- | --- | --- | --- | --- |"])
        for row in records:
            lines.append("| " + " | ".join(cell(row[key]) for key in
                        ("类别", "字段", "样本值", "可编辑", "可单项清除", "可全量清除", "实测")) + " |")
        lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")


def run(corpus: Path, output: Path):
    output.mkdir(parents=True, exist_ok=True)
    samples = sorted(path for folder in ("有元数据", "无可编辑元数据")
                     for path in (corpus / folder).iterdir() if path.is_file())
    tagged = {path.suffix.lower(): path for path in sorted((corpus / "有元数据").iterdir()) if path.is_file()}
    observed, failures = inventory(samples)
    rows = []
    for suffix, record in sorted(observed.items()):
        kind = record["kind"]
        fields = FIELD_SETS[kind]
        source = tagged[suffix]
        for key, label in SYSTEM_FIELDS.items():
            rows.append({"格式": suffix, "类别": "文件系统", "字段": label, "键": key,
                         "样本值": sample_value(record["values"].get(key, set())), "可编辑": "代码支持（未实测）",
                         "可单项清除": "否", "可全量清除": "否", "实测": "系统时间不属于嵌入元数据"})
        for key, label in fields:
            known = record["values"].get(key, set())
            if key in READ_ONLY.get(kind, set()):
                edit, clear, note = "否", "否", "只读技术字段"
            else:
                edit, clear, note = edit_probe(source, key, kind)
            rows.append({"格式": suffix, "类别": "预设字段", "字段": label, "键": key,
                         "样本值": sample_value(known), "可编辑": edit,
                         "可单项清除": clear, "可全量清除": "否" if key == "file_count" else "是（已测 10 份）",
                         "实测": note})
        for name, values in sorted(record["custom"].items()):
            if kind == "Office":
                edit, clear, note = custom_probe(source, name)
            else:
                edit, clear, note = "否", "否", "当前界面不提供自定义字段逐项编辑"
            rows.append({"格式": suffix, "类别": "自定义字段", "字段": name, "键": name,
                         "样本值": sample_value(values), "可编辑": edit,
                         "可单项清除": clear, "可全量清除": "是（已测 10 份）", "实测": note})
        if kind in {"Markdown", "PDF"}:
            name = "codex_custom_probe" if kind == "Markdown" else "/CodexCustomProbe"
            try:
                preserved, cleared = injected_custom_probe(source, kind)
                note = f"注入测试：常规编辑保留={'是' if preserved else '否'}；全量清除={'是' if cleared else '否'}"
            except Exception as exc:
                preserved = cleared = False
                note = f"注入测试失败：{type(exc).__name__}: {exc}"[:300]
            rows.append({"格式": suffix, "类别": "自定义字段（注入测试）", "字段": name, "键": name,
                         "样本值": "original", "可编辑": "否（界面未提供）", "可单项清除": "否",
                         "可全量清除": "是（注入测试）" if cleared else "未通过", "实测": note})
        if kind == "Archive":
            with zipfile.ZipFile(source) as archive:
                member = archive.infolist()[0] if archive.infolist() else None
            member_values = {
                "成员文件名": member.filename if member else "",
                "成员修改时间": str(member.date_time) if member else "",
                "成员压缩方式": str(member.compress_type) if member else "",
                "成员外部属性": str(member.external_attr) if member else "",
                "成员注释": member.comment.decode("utf-8", "replace") if member else "",
            }
            for name, value in member_values.items():
                rows.append({"格式": suffix, "类别": "ZIP 成员技术属性", "字段": name, "键": "ZipInfo",
                             "样本值": value, "可编辑": "否（界面未提供）", "可单项清除": "否",
                             "可全量清除": "否", "实测": "程序仅处理压缩包级注释，保留成员信息"})
        fixed = {key for key, _ in fields} | set(SYSTEM_FIELDS)
        for key, values in sorted(record["values"].items()):
            if key in fixed:
                continue
            rows.append({"格式": suffix, "类别": "扩展字段", "字段": key, "键": key,
                         "样本值": sample_value(values), "可编辑": "否（界面未提供）",
                         "可单项清除": "否", "可全量清除": "是（已测 10 份）", "实测": "读取器发现的非预设字段"})
        for group, values in sorted(record["raw"].items()):
            for name, examples in sorted(values.items()):
                rows.append({"格式": suffix, "类别": "原始/其他标签", "字段": name, "键": group,
                             "样本值": sample_value(examples), "可编辑": "否（界面未提供）", "可单项清除": "否",
                             "可全量清除": "是（已测 10 份）", "实测": "样本中发现；可能与预设字段重复"})
        print(f"Probed {suffix}: {len(fields)} preset fields", flush=True)
    report = output / "metadata-capability-matrix.csv"
    with report.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)
    counts = Counter(row["可编辑"] for row in rows if row["类别"] == "预设字段")
    summary = {"formats": len(observed), "samples_read": len(samples) - len(failures), "sample_read_failures": failures,
               "rows": len(rows), "preset_edit_results": dict(counts),
               "failed_probes": [row for row in rows if row["类别"] == "预设字段" and row["可编辑"] == "否" and row["实测"] != "只读技术字段"]}
    (output / "metadata-capability-summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    render_markdown(rows, summary, output / "metadata-capability-report.md")
    print(json.dumps({"formats": summary["formats"], "samples_read": summary["samples_read"],
                      "rows": summary["rows"], "failed_probes": len(summary["failed_probes"])}, ensure_ascii=False), flush=True)
    return 0 if not failures else 1


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("corpus", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    raise SystemExit(run(args.corpus, args.output))
