"""Create a one-slide PPT that disables inherited master objects."""
from __future__ import annotations

import argparse
import csv
import io
from pathlib import Path
import subprocess
import time

try:
    import pythoncom
    import win32com.client
except ImportError as exc:  # pragma: no cover
    raise SystemExit(f"pywin32 is required: {exc}") from exc


PP_ALERTS_NONE = 1
PP_LAYOUT_BLANK = 12
PP_SAVE_AS_PRESENTATION = 1
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
MSO_SHAPE_RECTANGLE = 1
MSO_FALSE = 0
MSO_TRUE = -1


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


def make_master_objects_fixture(destination: Path) -> dict[str, object]:
    destination = destination.resolve()
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        destination.unlink()

    pythoncom.CoInitialize()
    app = None
    presentation = None
    owned_pid = None
    try:
        app, owned_pid = _launch_powerpoint()
        presentation = app.Presentations.Add(True)
        presentation.PageSetup.SlideWidth = 720
        presentation.PageSetup.SlideHeight = 540
        slide = presentation.Slides.Add(1, PP_LAYOUT_BLANK)
        master = presentation.SlideMaster
        for shape_index in range(master.Shapes.Count, 0, -1):
            master.Shapes(shape_index).Delete()

        master_shape = master.Shapes.AddShape(
            MSO_SHAPE_RECTANGLE, 60, 60, 600, 420
        )
        master_shape.Fill.ForeColor.RGB = 0x0000FF
        master_shape.Fill.Solid()
        master_shape.Line.Visible = MSO_FALSE

        slide_shape = slide.Shapes.AddShape(
            MSO_SHAPE_RECTANGLE, 180, 150, 360, 240
        )
        slide_shape.Fill.ForeColor.RGB = 0xFF0000
        slide_shape.Fill.Solid()
        slide_shape.Line.Visible = MSO_FALSE
        slide.DisplayMasterShapes = MSO_FALSE
        presentation.SaveAs(str(destination), PP_SAVE_AS_PRESENTATION)
        return {
            "path": str(destination),
            "powerpoint_version": str(app.Version),
            "owned_pid": owned_pid,
            "slide_count": int(presentation.Slides.Count),
            "slide_width_pt": float(presentation.PageSetup.SlideWidth),
            "slide_height_pt": float(presentation.PageSetup.SlideHeight),
            "display_master_shapes": bool(slide.DisplayMasterShapes),
            "slide_shape_count": int(slide.Shapes.Count),
            "master_shape_count": int(master.Shapes.Count),
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
        default=Path("tests/fixtures/visual_master_objects_disabled.ppt"),
    )
    args = parser.parse_args()
    print(make_master_objects_fixture(args.output))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
