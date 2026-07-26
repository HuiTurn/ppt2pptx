"""Create a minimal PowerPoint 97-2003 fixture with one shape animation."""
from __future__ import annotations

import argparse
from pathlib import Path

try:
    import pythoncom
    import win32com.client
except ImportError as exc:  # pragma: no cover
    raise SystemExit(f"pywin32 is required: {exc}") from exc


PP_ALERTS_NONE = 1
PP_LAYOUT_BLANK = 12
PP_SAVE_AS_PRESENTATION = 1
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
MSO_TEXT_ORIENTATION_HORIZONTAL = 1
PP_EFFECT_APPEAR = 3844


def make_animation_object_fixture(destination: Path) -> dict[str, object]:
    destination = destination.resolve()
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        destination.unlink()

    pythoncom.CoInitialize()
    app = None
    presentation = None
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        app.DisplayAlerts = PP_ALERTS_NONE
        try:
            app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
        except Exception:
            pass
        presentation = app.Presentations.Add(True)
        presentation.PageSetup.SlideWidth = 720
        presentation.PageSetup.SlideHeight = 540
        slide = presentation.Slides.Add(1, PP_LAYOUT_BLANK)
        box = slide.Shapes.AddTextbox(
            MSO_TEXT_ORIENTATION_HORIZONTAL,
            72,
            72,
            480,
            80,
        )
        text = box.TextFrame.TextRange
        text.Text = "animated fixture"
        text.Font.Name = "Arial"
        text.Font.Size = 28
        box.AnimationSettings.Animate = True
        box.AnimationSettings.EntryEffect = PP_EFFECT_APPEAR
        presentation.SaveAs(str(destination), PP_SAVE_AS_PRESENTATION)
        return {
            "path": str(destination),
            "powerpoint_version": str(app.Version),
            "slide_count": int(presentation.Slides.Count),
            "animated_shape_count": 1,
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
        default=Path("tests/fixtures/visual_animation_object.ppt"),
    )
    args = parser.parse_args()
    print(make_animation_object_fixture(args.output))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
