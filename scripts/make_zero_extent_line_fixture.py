"""Create a one-slide legacy PPT retaining zero-axis OfficeArt master lines."""
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


_SEED_BLOB = "d157798671bf2499d3add9a2cf152049975820e6"
_SEED_SHA256 = "676554e0ed9503673ecd3800f6a8e9509949dd02553d2e3c847e5c917e5b4510"
PP_ALERTS_NONE = 1
PP_LAYOUT_BLANK = 12
PP_SAVE_AS_PRESENTATION = 1
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
MSO_GROUP = 6
MSO_LINE = 9


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


def _launch_powerpoint():
    last_error: Exception | None = None
    for attempt in range(3):
        before = _powerpoint_pids()
        app = None
        try:
            app = win32com.client.DispatchEx("PowerPoint.Application")
            try:
                app.Visible = True
            except Exception:
                pass
            app.DisplayAlerts = PP_ALERTS_NONE
            try:
                app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
            except Exception:
                pass
            for _ in range(20):
                created = _powerpoint_pids() - before
                if len(created) == 1:
                    return app, created.pop()
                if len(created) > 1:
                    break
                time.sleep(0.1)
            raise RuntimeError("unable to record isolated PowerPoint instance PID")
        except Exception as exc:
            last_error = exc
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass
            app = None
            time.sleep(0.5 * (attempt + 1))
    raise RuntimeError("unable to launch stable isolated PowerPoint") from last_error


def _download_seed() -> bytes:
    url = f"https://api.github.com/repos/apache/poi/git/blobs/{_SEED_BLOB}"
    with urlopen(url, timeout=60) as response:
        payload = json.load(response)
    content = base64.b64decode(payload["content"])
    digest = hashlib.sha256(content).hexdigest()
    if digest != _SEED_SHA256:
        raise RuntimeError(f"unexpected Apache POI seed SHA-256: {digest}")
    return content


def make_zero_extent_line_fixture(destination: Path) -> dict[str, object]:
    destination = destination.resolve()
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        destination.unlink()

    pythoncom.CoInitialize()
    app = None
    source_presentation = None
    presentation = None
    owned_pid = None
    try:
        app, owned_pid = _launch_powerpoint()

        with tempfile.TemporaryDirectory(
            prefix="ppt2pptx-zero-extent-line-fixture-"
        ) as directory:
            seed = Path(directory) / "apache-poi-37625.ppt"
            seed.write_bytes(_download_seed())
            source_presentation = app.Presentations.Open(
                str(seed),
                ReadOnly=True,
                Untitled=True,
                WithWindow=False,
            )
            source_master = source_presentation.SlideMaster
            source_grid = next(
                (
                    source_master.Shapes(index)
                    for index in range(1, source_master.Shapes.Count + 1)
                    if int(source_master.Shapes(index).Type) == MSO_GROUP
                    and int(source_master.Shapes(index).GroupItems.Count) >= 10
                ),
                None,
            )
            if source_grid is None:
                raise RuntimeError("Apache POI seed contains no master grid group")
            presentation = app.Presentations.Add(True)
            presentation.PageSetup.SlideWidth = 720
            presentation.PageSetup.SlideHeight = 540
            presentation.Slides.Add(1, PP_LAYOUT_BLANK)
            master = presentation.SlideMaster
            for shape_index in range(master.Shapes.Count, 0, -1):
                master.Shapes(shape_index).Delete()
            line_count = 0
            weights: set[float] = set()
            for shape_index in range(1, source_grid.GroupItems.Count + 1):
                source_shape = source_grid.GroupItems(shape_index)
                if int(source_shape.Type) != MSO_LINE:
                    continue
                line = master.Shapes.AddLine(
                    float(source_grid.Left + source_shape.Left),
                    float(source_grid.Top + source_shape.Top),
                    float(source_grid.Left + source_shape.Left + source_shape.Width),
                    float(source_grid.Top + source_shape.Top + source_shape.Height),
                )
                line.Line.Weight = float(source_shape.Line.Weight)
                line.Line.ForeColor.RGB = int(source_shape.Line.ForeColor.RGB)
                line_count += 1
                weights.add(float(line.Line.Weight))
            if line_count < 10:
                raise RuntimeError(
                    f"Apache POI seed retained only {line_count} master lines"
                )
            presentation.SaveAs(str(destination), PP_SAVE_AS_PRESENTATION)
            return {
                "path": str(destination),
                "powerpoint_version": str(app.Version),
                "owned_pid": owned_pid,
                "slide_count": int(presentation.Slides.Count),
                "slide_width_pt": float(presentation.PageSetup.SlideWidth),
                "slide_height_pt": float(presentation.PageSetup.SlideHeight),
                "master_line_count": line_count,
                "line_weight_pt": sorted(weights),
                "seed_blob": _SEED_BLOB,
            }
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        if source_presentation is not None:
            try:
                source_presentation.Close()
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        presentation = None
        source_presentation = None
        app = None
        pythoncom.CoUninitialize()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("tests/fixtures/visual_zero_extent_lines.ppt"),
    )
    args = parser.parse_args()
    print(make_zero_extent_line_fixture(args.output))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
