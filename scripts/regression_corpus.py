"""Exercise every supported sample on a copy and compare its actual content.

Usage: python scripts/regression_corpus.py CORPUS OUTPUT
Requires the app dependencies plus olefile; video cases require FFmpeg.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import shutil
import subprocess
import tempfile
import time
import zipfile
from collections import Counter
from pathlib import Path

from metadata_editor import engine


def digest(parts):
    checksum = hashlib.sha256()
    for part in parts:
        if isinstance(part, str):
            part = part.encode("utf-8")
        checksum.update(len(part).to_bytes(8, "big"))
        checksum.update(part)
    return checksum.hexdigest()


def zip_payload(path, office=False):
    with zipfile.ZipFile(path) as archive:
        bad = archive.testzip()
        if bad:
            raise ValueError(f"ZIP member failed CRC: {bad}")
        names = archive.namelist()
        if office and "[Content_Types].xml" not in names:
            raise ValueError("Office content types are missing")
        parts = []
        for member in archive.infolist():
            if office and member.filename.startswith("docProps/"):
                continue
            parts.extend((member.filename, archive.read(member)))
        return digest(parts)


def legacy_payload(path):
    import olefile

    required = {".doc": "WordDocument", ".xls": "Workbook", ".ppt": "PowerPoint Document"}[path.suffix.lower()]
    with olefile.OleFileIO(str(path)) as compound:
        streams = compound.listdir(streams=True, storages=False)
        names = ["/".join(parts) for parts in streams]
        if required not in names:
            raise ValueError(f"Missing OLE content stream: {required}")
        parts = []
        for name in streams:
            if name[-1] in {"\x05SummaryInformation", "\x05DocumentSummaryInformation"}:
                continue
            parts.extend(("/".join(name), compound.openstream(name).read()))
        return digest(parts)


def pdf_payload(path):
    from pypdf import PdfReader

    reader = PdfReader(path)
    if reader.is_encrypted:
        raise ValueError("Encrypted PDF cannot be inspected")
    parts = [str(len(reader.pages))]
    for page in reader.pages:
        contents = page.get_contents()
        parts.extend((str(tuple(page.mediabox)), contents.get_data() if contents else b"", page.extract_text() or ""))
    return digest(parts)


def image_payload(path):
    from PIL import Image

    parts = []
    with Image.open(path) as image:
        for frame in range(getattr(image, "n_frames", 1)):
            image.seek(frame)
            pixels = image.convert("RGBA")
            parts.extend((str(pixels.size), pixels.tobytes()))
    return digest(parts)


def media_payload(path):
    command = [engine._ffmpeg(), "-v", "error", "-i", str(path), "-map", "0", "-f", "streamhash", "-hash", "SHA256", "-"]
    run = subprocess.run(command, capture_output=True, text=True, encoding="utf-8", errors="replace",
                         creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    if run.returncode:
        raise ValueError(f"Media decode failed: {run.stderr.strip()[-300:]}")
    return digest([run.stdout])


def content_fingerprint(path):
    suffix = path.suffix.lower()
    if suffix in engine.OFFICE_SUFFIXES:
        return zip_payload(path, office=True)
    if suffix in engine.LEGACY_SUFFIXES:
        return legacy_payload(path)
    if suffix == ".pdf":
        return pdf_payload(path)
    if suffix in engine.MARKDOWN_SUFFIXES:
        return digest([engine._markdown_parts(path)[1]])
    if suffix in engine.IMAGE_SUFFIXES:
        return image_payload(path)
    if suffix in engine.ZIP_SUFFIXES:
        return zip_payload(path)
    if suffix in engine.AUDIO_SUFFIXES | engine.VIDEO_SUFFIXES:
        return media_payload(path)
    raise ValueError(f"Unsupported suffix: {suffix}")


def run_corpus(corpus: Path, output: Path):
    output.mkdir(parents=True, exist_ok=True)
    samples = sorted(path for category in ("有元数据", "无可编辑元数据")
                     for path in (corpus / category).iterdir() if path.is_file())
    supported = (engine.OFFICE_SUFFIXES | engine.LEGACY_SUFFIXES | engine.MARKDOWN_SUFFIXES |
                 engine.IMAGE_SUFFIXES | engine.ZIP_SUFFIXES | engine.AUDIO_SUFFIXES |
                 engine.VIDEO_SUFFIXES | {".pdf"})
    if not samples or any(path.suffix.lower() not in supported for path in samples):
        raise ValueError("Corpus is empty or includes unsupported formats")
    rows = []
    for index, source in enumerate(samples, 1):
        start = time.monotonic()
        row = {"file": str(source.relative_to(corpus)), "format": source.suffix.lower(),
               "read": "", "clear": "", "verification": "", "content": "", "detail": ""}
        try:
            before = engine.load(str(source))
            original_content = content_fingerprint(source)
            row["read"] = "pass"
            with tempfile.TemporaryDirectory(dir=output) as scratch:
                target = Path(scratch) / source.name
                shutil.copy2(source, target)
                engine.clear_embedded_properties(engine.load(str(target)))
                row["clear"] = "pass"
                _, residual = engine.verify_embedded_clear(target)
                row["verification"] = "pass" if not residual else "fail"
                if residual:
                    row["detail"] = "Remaining metadata: " + ", ".join(residual)
                after_content = content_fingerprint(target)
                row["content"] = "pass" if original_content == after_content else "fail"
                if row["content"] == "fail":
                    row["detail"] = (row["detail"] + "; " if row["detail"] else "") + "Content fingerprint changed"
        except Exception as exc:
            row["detail"] = (row["detail"] + "; " if row["detail"] else "") + f"{type(exc).__name__}: {exc}"
        row["seconds"] = round(time.monotonic() - start, 3)
        rows.append(row)
        if index % 25 == 0 or index == len(samples):
            print(f"Checked {index}/{len(samples)}", flush=True)
    report = output / "regression-results.csv"
    with report.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)
    summary = {"total": len(rows), "by_check": {key: dict(Counter(row[key] for row in rows))
                                                 for key in ("read", "clear", "verification", "content")},
               "failures": [row for row in rows if any(row[key] != "pass" for key in ("read", "clear", "verification", "content"))]}
    (output / "regression-summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"total": summary["total"], "by_check": summary["by_check"],
                      "failed_files": len(summary["failures"])}, ensure_ascii=False), flush=True)
    return 0 if not summary["failures"] else 1


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("corpus", type=Path)
    parser.add_argument("output", type=Path)
    arguments = parser.parse_args()
    raise SystemExit(run_corpus(arguments.corpus, arguments.output))
