"""Create a one-slide PPT retaining a legacy OfficeArt patterned line."""
from __future__ import annotations

import argparse
from pathlib import Path
import tempfile

try:
    import pythoncom
except ImportError as exc:  # pragma: no cover
    raise SystemExit(f"pywin32 is required: {exc}") from exc

from make_zero_extent_line_fixture import (
    MSO_GROUP,
    MSO_LINE,
    PP_LAYOUT_BLANK,
    PP_SAVE_AS_PRESENTATION,
    _download_seed,
    _launch_powerpoint,
    _SEED_BLOB,
)


def make_pattern_line_fixture(destination: Path) -> dict[str, object]:
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
            prefix="ppt2pptx-pattern-line-fixture-"
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
            source_line = next(
                (
                    source_grid.GroupItems(index)
                    for index in range(1, source_grid.GroupItems.Count + 1)
                    if int(source_grid.GroupItems(index).Type) == MSO_LINE
                    and float(source_grid.GroupItems(index).Width)
                    > float(source_grid.GroupItems(index).Height)
                ),
                None,
            )
            if source_line is None:
                raise RuntimeError("Apache POI seed contains no horizontal grid line")

            presentation = app.Presentations.Add(True)
            presentation.PageSetup.SlideWidth = 720
            presentation.PageSetup.SlideHeight = 540
            slide = presentation.Slides.Add(1, PP_LAYOUT_BLANK)
            master = presentation.SlideMaster
            for shape_index in range(master.Shapes.Count, 0, -1):
                master.Shapes(shape_index).Delete()

            source_line.Copy()
            pasted = slide.Shapes.Paste()
            line = pasted(1)
            line.Left = 0
            line.Top = 270
            line.Width = 720
            presentation.SaveAs(str(destination), PP_SAVE_AS_PRESENTATION)
            return {
                "path": str(destination),
                "powerpoint_version": str(app.Version),
                "owned_pid": owned_pid,
                "slide_count": int(presentation.Slides.Count),
                "slide_width_pt": float(presentation.PageSetup.SlideWidth),
                "slide_height_pt": float(presentation.PageSetup.SlideHeight),
                "slide_shape_count": int(slide.Shapes.Count),
                "line_left_pt": float(line.Left),
                "line_top_pt": float(line.Top),
                "line_width_pt": float(line.Width),
                "line_height_pt": float(line.Height),
                "line_weight_pt": float(line.Line.Weight),
                "line_fore_color_rgb": int(line.Line.ForeColor.RGB),
                "line_back_color_rgb": int(line.Line.BackColor.RGB),
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
        default=Path("tests/fixtures/visual_pattern_line.ppt"),
    )
    args = parser.parse_args()
    print(make_pattern_line_fixture(args.output))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
