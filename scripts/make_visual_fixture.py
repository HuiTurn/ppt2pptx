"""Create a minimal PowerPoint 97–2003 .ppt fixture via PowerPoint COM.

The fixture is used only by visual-regression tooling. Conversion never depends
on COM. Generated files are deterministic enough for local baselines.
"""
from __future__ import annotations

import argparse
from pathlib import Path
import sys

try:
    import pythoncom
    import win32com.client
except ImportError as exc:  # pragma: no cover
    raise SystemExit(f"pywin32 is required: {exc}") from exc

PP_ALERTS_NONE = 1
PP_LAYOUT_BLANK = 12
PP_SAVE_AS_PRESENTATION = 1


def make_fixture(destination: Path) -> dict[str, object]:
    destination = destination.resolve()
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        destination.unlink()

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("PowerPoint.Application")
    presentation = None
    try:
        # Recent PowerPoint builds reject Visible=False for some automation hosts.
        try:
            app.Visible = True
        except Exception:
            pass
        try:
            app.DisplayAlerts = PP_ALERTS_NONE
        except Exception:
            pass
        try:
            app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        except Exception:
            pass
        try:
            app.AskToUpdateLinks = False
        except Exception:
            pass
        presentation = app.Presentations.Add()
        presentation.PageSetup.SlideWidth = 720
        presentation.PageSetup.SlideHeight = 540

        slide1 = presentation.Slides.Add(1, PP_LAYOUT_BLANK)
        box = slide1.Shapes.AddTextbox(1, 72, 72, 500, 80)
        text = box.TextFrame.TextRange
        text.Text = "ppt2pptx visual fixture"
        text.Font.Name = "Arial"
        text.Font.Size = 32
        shape = slide1.Shapes.AddShape(1, 72, 200, 180, 100)
        try:
            shape.Fill.ForeColor.RGB = 0x00A5FF
            shape.Line.ForeColor.RGB = 0x000000
        except Exception:
            # Some hosts reject early-bound color assignment; geometry alone is enough.
            pass

        slide2 = presentation.Slides.Add(2, PP_LAYOUT_BLANK)
        box2 = slide2.Shapes.AddTextbox(1, 72, 200, 400, 60)
        text2 = box2.TextFrame.TextRange
        text2.Text = "second slide"
        text2.Font.Name = "Arial"
        text2.Font.Size = 28
        slide2.SlideShowTransition.Hidden = False

        presentation.SaveAs(str(destination), PP_SAVE_AS_PRESENTATION)
        version = str(getattr(app, "Version", "unknown"))
        return {
            "path": str(destination),
            "powerpoint_version": version,
            "slide_count": int(presentation.Slides.Count),
            "slide_width_pt": float(presentation.PageSetup.SlideWidth),
            "slide_height_pt": float(presentation.PageSetup.SlideHeight),
            "hidden": [False, False],
        }
    finally:
        try:
            if presentation is not None:
                presentation.Close()
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("tests/fixtures/visual_minimal.ppt"),
        help="destination .ppt path",
    )
    args = parser.parse_args(argv)
    info = make_fixture(args.output)
    print(info)
    return 0 if Path(info["path"]).is_file() else 1


if __name__ == "__main__":
    raise SystemExit(main())
