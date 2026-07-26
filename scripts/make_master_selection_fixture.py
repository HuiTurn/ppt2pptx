"""Create a one-slide PPT whose layout master owns the slide background."""
from __future__ import annotations

import argparse
from pathlib import Path

try:
    import pythoncom
except ImportError as exc:  # pragma: no cover
    raise SystemExit(f"pywin32 is required: {exc}") from exc

from make_master_objects_fixture import _launch_powerpoint


PP_LAYOUT_BLANK = 12
PP_SAVE_AS_PRESENTATION = 1
MSO_FALSE = 0
MSO_TRUE = -1


def make_master_selection_fixture(destination: Path) -> dict[str, object]:
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

        master = presentation.SlideMaster
        for shape_index in range(master.Shapes.Count, 0, -1):
            master.Shapes(shape_index).Delete()

        slide = presentation.Slides.Add(1, PP_LAYOUT_BLANK)
        layout = slide.CustomLayout
        for shape_index in range(layout.Shapes.Count, 0, -1):
            layout.Shapes(shape_index).Delete()
        layout.FollowMasterBackground = MSO_FALSE
        layout.Background.Fill.ForeColor.RGB = 0xFF0000
        layout.Background.Fill.Solid()
        slide.DisplayMasterShapes = MSO_TRUE

        presentation.SaveAs(str(destination), PP_SAVE_AS_PRESENTATION)
        return {
            "path": str(destination),
            "powerpoint_version": str(app.Version),
            "owned_pid": owned_pid,
            "slide_count": int(presentation.Slides.Count),
            "slide_width_pt": float(presentation.PageSetup.SlideWidth),
            "slide_height_pt": float(presentation.PageSetup.SlideHeight),
            "slide_shape_count": int(slide.Shapes.Count),
            "layout_shape_count": int(layout.Shapes.Count),
            "layout_follows_master_background": bool(
                layout.FollowMasterBackground
            ),
            "layout_background_rgb": int(
                layout.Background.Fill.ForeColor.RGB
            ),
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
        default=Path("tests/fixtures/visual_master_selection.ppt"),
    )
    args = parser.parse_args()
    print(make_master_selection_fixture(args.output))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
