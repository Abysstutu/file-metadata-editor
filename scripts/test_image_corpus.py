"""Check image metadata round trips and decoded pixels on corpus copies."""

import argparse
import json
import shutil
import tempfile
from pathlib import Path

from metadata_editor.engine import clear_embedded_properties, load, write_image
from scripts.regression_corpus import content_fingerprint


SUFFIXES = {".jpg", ".jpeg", ".tif", ".tiff", ".webp"}


def run(corpus: Path, report: Path) -> int:
    sources = sorted(path for directory in ("有元数据", "无可编辑元数据")
                     for path in (corpus / directory).iterdir() if path.suffix.lower() in SUFFIXES)
    rows = []
    for source in sources:
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / source.name
            try:
                shutil.copy2(source, target)
                fingerprint = content_fingerprint(target)
                meta = load(str(target))
                values = {**meta.values, "title": "图片标题", "author": "测试作者",
                          "description": "中文描述", "gps": "39.90, 116.40"}
                write_image(meta, values)
                changed = load(str(target))
                assert all(changed.values[key] == values[key] for key in ("title", "author", "description"))
                assert changed.values["gps"] == "39.9, 116.4"
                assert content_fingerprint(target) == fingerprint, "图像像素在写入时发生变化"
                clear_embedded_properties(changed)
                cleared = load(str(target))
                assert all(not cleared.values.get(key) for key in ("title", "author", "description", "gps"))
                assert content_fingerprint(target) == fingerprint, "图像像素在清除时发生变化"
                rows.append({"file": source.name, "status": "passed"})
            except Exception as exc:
                rows.append({"file": source.name, "status": "failed", "reason": str(exc)})
        print(f"{rows[-1]['status']}: {source.name}", flush=True)
    summary = {"samples": len(rows), "passed": sum(row["status"] == "passed" for row in rows), "results": rows}
    report.parent.mkdir(parents=True, exist_ok=True)
    report.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Images: {summary['passed']}/{len(rows)} passed", flush=True)
    return 0 if rows and summary["passed"] == len(rows) else 1


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("corpus", type=Path)
    parser.add_argument("report", type=Path)
    args = parser.parse_args()
    raise SystemExit(run(args.corpus, args.report))
