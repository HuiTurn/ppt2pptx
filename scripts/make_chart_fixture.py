"""Create a minimal PowerPoint 97-2003 fixture with one legacy MS Graph chart."""
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


def make_chart_fixture(destination: Path) -> dict[str, object]:
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
        slide = presentation.Slides.Add(1, PP_LAYOUT_BLANK)
        chart = slide.Shapes.AddOLEObject(
            120,
            90,
            480,
            300,
            ClassName="MSGraph.Chart.8",
            Link=False,
            DisplayAsIcon=False,
        )
        chart.Name = "Minimal MS Graph chart"
        presentation.SaveAs(str(destination), PP_SAVE_AS_PRESENTATION)
        return {
            "path": str(destination),
            "powerpoint_version": str(app.Version),
            "slide_count": int(presentation.Slides.Count),
            "chart_object_count": 1,
            "progid": str(chart.OLEFormat.ProgID),
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
        default=Path("tests/fixtures/visual_chart.ppt"),
    )
    args = parser.parse_args()
    print(make_chart_fixture(args.output))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
