"""Create a one-slide PPT containing one embedded MathType equation."""
from __future__ import annotations

import argparse
from pathlib import Path
import tempfile

import pythoncom

from make_zero_extent_line_fixture import (
    PP_LAYOUT_BLANK,
    PP_SAVE_AS_PRESENTATION,
    _download_seed,
    _launch_powerpoint,
)


MSO_EMBEDDED_OLE_OBJECT = 7


def make_equation_ole_fixture(destination: Path) -> dict[str, object]:
    destination = destination.resolve()
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        destination.unlink()

    pythoncom.CoInitialize()
    app = None
    source_presentation = None
    presentation = None
    owned_pid = None
    result: dict[str, object] | None = None
    try:
        app, owned_pid = _launch_powerpoint()
        try:
            app.AskToUpdateLinks = False
        except Exception:
            pass
        with tempfile.TemporaryDirectory(
            prefix="ppt2pptx-equation-ole-fixture-"
        ) as directory:
            seed = Path(directory) / "apache-poi-37625.ppt"
            seed.write_bytes(_download_seed())
            source_presentation = app.Presentations.Open(
                str(seed),
                ReadOnly=True,
                Untitled=True,
                WithWindow=False,
            )
            equation = None
            for slide_index in range(
                1, source_presentation.Slides.Count + 1
            ):
                slide = source_presentation.Slides(slide_index)
                for shape_index in range(1, slide.Shapes.Count + 1):
                    shape = slide.Shapes(shape_index)
                    if int(shape.Type) != MSO_EMBEDDED_OLE_OBJECT:
                        continue
                    if str(shape.OLEFormat.ProgID) == "Equation.DSMT4":
                        equation = shape
                        break
                if equation is not None:
                    break
            if equation is None:
                raise RuntimeError(
                    "Apache POI seed contains no Equation.DSMT4 object"
                )

            presentation = app.Presentations.Add(True)
            slide = presentation.Slides.Add(1, PP_LAYOUT_BLANK)
            equation.Copy()
            pasted = slide.Shapes.Paste()
            if pasted.Count != 1:
                raise RuntimeError("expected one pasted equation object")
            output_equation = pasted(1)
            output_equation.Left = 180
            output_equation.Top = 180
            presentation.SaveAs(
                str(destination),
                PP_SAVE_AS_PRESENTATION,
            )
            result = {
                "path": str(destination),
                "powerpoint_version": str(app.Version),
                "owned_pid": owned_pid,
                "slide_count": int(presentation.Slides.Count),
                "ole_count": 1,
                "progid": str(output_equation.OLEFormat.ProgID),
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
    if result is None:
        raise RuntimeError("fixture generation did not complete")
    return result


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("tests/fixtures/visual_equation_ole.ppt"),
    )
    args = parser.parse_args()
    print(make_equation_ole_fixture(args.output))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
