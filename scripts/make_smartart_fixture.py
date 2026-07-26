"""Create a minimal legacy PPT containing one PowerPoint-saved SmartArt object."""
from __future__ import annotations

import argparse
import base64
import csv
import hashlib
import io
import json
from pathlib import Path
import subprocess
import tempfile
import time
from urllib.request import urlopen

try:
    import pythoncom
    import win32com.client
except ImportError as exc:  # pragma: no cover
    raise SystemExit(f"pywin32 is required: {exc}") from exc


_SEED_BLOB = "7b7731d78f27352ccea5713400d4994058bcc381"
_SEED_SHA256 = "b97e4c6d2ee1dd4094f50f9043820610268dd452d7717c537f5456636adcc353"


def _powerpoint_pids() -> set[int]:
    completed = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/FO", "CSV", "/NH"],
        capture_output=True,
        text=True,
        errors="replace",
        check=False,
    )
    pids: set[int] = set()
    for row in csv.reader(io.StringIO(completed.stdout)):
        if len(row) >= 2 and row[0].casefold() == "powerpnt.exe":
            try:
                pids.add(int(row[1]))
            except ValueError:
                pass
    return pids


def _download_seed() -> bytes:
    url = f"https://api.github.com/repos/apache/poi/git/blobs/{_SEED_BLOB}"
    with urlopen(url, timeout=60) as response:
        payload = json.load(response)
    content = base64.b64decode(payload["content"])
    digest = hashlib.sha256(content).hexdigest()
    if digest != _SEED_SHA256:
        raise RuntimeError(f"unexpected Apache POI SmartArt seed SHA-256: {digest}")
    return content


def make_smartart_fixture(destination: Path) -> dict[str, object]:
    destination = destination.resolve()
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        destination.unlink()

    pythoncom.CoInitialize()
    app = None
    presentation = None
    owned_pid = None
    try:
        before = _powerpoint_pids()
        app = win32com.client.DispatchEx("PowerPoint.Application")
        try:
            app.Visible = True
        except Exception:
            pass
        app.DisplayAlerts = 1
        try:
            app.AutomationSecurity = 3
        except Exception:
            pass
        for _ in range(20):
            created = _powerpoint_pids() - before
            if len(created) == 1:
                owned_pid = created.pop()
                break
            if len(created) > 1:
                break
            time.sleep(0.1)
        if owned_pid is None:
            raise RuntimeError("unable to record isolated PowerPoint instance PID")
        with tempfile.TemporaryDirectory(prefix="ppt2pptx-smartart-fixture-") as directory:
            seed = Path(directory) / "apache-poi-smartart.pptx"
            seed.write_bytes(_download_seed())
            presentation = app.Presentations.Open(
                str(seed),
                ReadOnly=True,
                Untitled=True,
                WithWindow=False,
            )
            target_slide = None
            smartart = None
            for slide_index in range(1, presentation.Slides.Count + 1):
                slide = presentation.Slides(slide_index)
                for shape_index in range(1, slide.Shapes.Count + 1):
                    shape = slide.Shapes(shape_index)
                    if shape.HasSmartArt:
                        target_slide = slide
                        smartart = shape
                        break
                if smartart is not None:
                    break
            if target_slide is None or smartart is None:
                raise RuntimeError("Apache POI seed contains no SmartArt shape")

            for slide_index in range(presentation.Slides.Count, 0, -1):
                if presentation.Slides(slide_index).SlideID != target_slide.SlideID:
                    presentation.Slides(slide_index).Delete()
            for shape_index in range(target_slide.Shapes.Count, 0, -1):
                if target_slide.Shapes(shape_index).Id != smartart.Id:
                    target_slide.Shapes(shape_index).Delete()
            presentation.SaveAs(str(destination), 1)
            return {
                "path": str(destination),
                "powerpoint_version": str(app.Version),
                "owned_pid": owned_pid,
                "slide_count": int(presentation.Slides.Count),
                "shape_count": int(target_slide.Shapes.Count),
                "seed_blob": _SEED_BLOB,
            }
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        presentation = None
        app = None
        pythoncom.CoUninitialize()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("tests/fixtures/visual_smartart.ppt"),
    )
    args = parser.parse_args()
    print(make_smartart_fixture(args.output))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
