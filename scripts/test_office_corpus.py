"""Verify Office edit/clear round trips and untouched document content.

Usage: python -m scripts.test_office_corpus CORPUS REPORT [--suffixes .doc .docx]
"""

from __future__ import annotations

import argparse
import json
import shutil
import tempfile
from pathlib import Path

from metadata_editor.engine import FIELDS, clear_embedded_properties, load, write_legacy, write_office
from scripts.regression_corpus import content_fingerprint


DEFAULT_SUFFIXES = {".doc", ".docx", ".docm", ".dotx", ".dotm"}


def run(corpus: Path, report: Path, suffixes: set[str] = DEFAULT_SUFFIXES) -> int:
    sources = sorted(path for directory in ("有元数据", "无可编辑元数据")
                     for path in (corpus / directory).iterdir() if path.suffix.lower() in suffixes)
    results = []
    for source in sources:
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / source.name
            try:
                shutil.copy2(source, target)
                fingerprint = content_fingerprint(target)
                before = load(str(target))
                values = dict(before.values)
                values.update({"title": "Word 回归测试", "created": "2027-02-03 04:05:06", "totalTime": "37"})
                custom = {**before.custom, "回归测试": "已写入"}
                if before.kind == "Legacy Office": write_legacy(before, values, False, custom)
                else: write_office(before, values, custom, False)
                after = load(str(target))
                assert after.values["title"] == "Word 回归测试"
                assert after.values["created"] == "2027-02-03 04:05:06"
                assert after.values["totalTime"] == "37"
                assert after.custom["回归测试"] == "已写入"
                assert content_fingerprint(target) == fingerprint
                clear_embedded_properties(after)
                cleared = load(str(target))
                assert all(not cleared.values.get(key, "") for key, _ in FIELDS)
                assert not cleared.custom
                assert content_fingerprint(target) == fingerprint
                results.append({"file": source.name, "status": "passed"})
            except Exception as exc:
                results.append({"file": source.name, "status": "failed", "reason": str(exc)})
        print(f"{results[-1]['status']}: {source.name}", flush=True)
    summary = {"samples": len(sources), "passed": sum(row["status"] == "passed" for row in results), "results": results}
    report.parent.mkdir(parents=True, exist_ok=True)
    report.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Office samples: {summary['passed']}/{summary['samples']} passed", flush=True)
    return 0 if sources and summary["passed"] == len(sources) else 1


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("corpus", type=Path)
    parser.add_argument("report", type=Path)
    parser.add_argument("--suffixes", nargs="+", default=sorted(DEFAULT_SUFFIXES))
    args = parser.parse_args()
    raise SystemExit(run(args.corpus, args.report, set(args.suffixes)))
