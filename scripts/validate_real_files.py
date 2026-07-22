"""Convert a local real-PPT corpus and optionally render every output."""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys

from ppt2pptx import convert
from ppt2pptx.errors import EncryptedPresentationError, Ppt2PptxError, UnsupportedPptVersionError

CFB_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()

def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("input", type=Path)
    parser.add_argument("-o", "--output", type=Path, required=True)
    parser.add_argument("--report", type=Path)
    parser.add_argument("--render", action="store_true")
    parser.add_argument("--password", help="password applied to encrypted corpus files")
    args = parser.parse_args(argv)
    renderer = shutil.which("soffice") if args.render else None
    if args.render and renderer is None:
        parser.error("--render requires LibreOffice/soffice")
    args.output.mkdir(parents=True, exist_ok=True)
    inputs = [path for path in sorted(args.input.rglob("*")) if path.is_file() and path.suffix.casefold() == ".ppt" and path.read_bytes()[:8] == CFB_SIGNATURE]
    results: list[dict[str, object]] = []
    converted = rendered = unsupported = failed = 0
    for source in inputs:
        destination = args.output / source.relative_to(args.input).with_suffix(".pptx")
        item: dict[str, object] = {"source": str(source), "sha256": _sha256(source), "input_bytes": source.stat().st_size}
        try:
            result = convert(source, destination, password=args.password)
            item.update(status="converted", output=str(destination), slide_count=result.slide_count,
                        text_box_count=sum(len(slide.text_boxes) for slide in result.presentation.slides),
                        picture_count=sum(len(slide.pictures) for slide in result.presentation.slides),
                        shape_count=sum(len(slide.shapes) for slide in result.presentation.slides),
                        comment_count=sum(len(slide.comments) for slide in result.presentation.slides),
                        note_count=sum(len(slide.notes) for slide in result.presentation.slides),
                        warning_count=len(result.report.warnings))
            converted += 1
            if renderer:
                render_dir = args.output / "rendered"
                render_dir.mkdir(parents=True, exist_ok=True)
                process = subprocess.run([renderer, "--headless", "--convert-to", "pdf", "--outdir", str(render_dir), str(destination)], capture_output=True, text=True, timeout=120)
                pdf = render_dir / destination.with_suffix(".pdf").name
                item["rendered"] = process.returncode == 0 and pdf.is_file()
                if item["rendered"]:
                    rendered += 1
                else:
                    item["render_error"] = (process.stderr or process.stdout).strip()
        except (EncryptedPresentationError, UnsupportedPptVersionError) as exc:
            item.update(status="unsupported", error=str(exc)); unsupported += 1
        except (Ppt2PptxError, OSError, subprocess.SubprocessError) as exc:
            item.update(status="failed", error=f"{type(exc).__name__}: {exc}"); failed += 1
        results.append(item)
        print(f"{item['status']:11} {source.name}")
    summary = {"input": str(args.input), "file_count": len(inputs), "converted": converted, "rendered": rendered, "unsupported": unsupported, "failed": failed, "results": results}
    report = args.report or args.output / "report.json"
    report.parent.mkdir(parents=True, exist_ok=True)
    report.write_text(json.dumps(summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({key: summary[key] for key in ("file_count", "converted", "rendered", "unsupported", "failed")}, indent=2))
    return 1 if failed else 0

if __name__ == "__main__":
    raise SystemExit(main())
