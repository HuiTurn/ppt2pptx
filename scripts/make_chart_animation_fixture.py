"""Create a minimal legacy MS Graph chart with a by-series animation."""
from __future__ import annotations

import argparse
from pathlib import Path

try:
    import pythoncom
    import win32com.client
except ImportError as exc:  # pragma: no cover
    raise SystemExit(f"pywin32 is required: {exc}") from exc


def make_chart_animation_fixture(destination: Path) -> dict[str, object]:
    destination = destination.resolve()
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        destination.unlink()

    pythoncom.CoInitialize()
    app = None
    presentation = None
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        app.DisplayAlerts = 1
        try:
            app.AutomationSecurity = 3
        except Exception:
            pass
        presentation = app.Presentations.Add(True)
        presentation.PageSetup.SlideWidth = 720
        presentation.PageSetup.SlideHeight = 540
        slide = presentation.Slides.Add(1, 12)
        chart = slide.Shapes.AddOLEObject(
            120,
            90,
            480,
            300,
            ClassName="MSGraph.Chart.8",
            Link=False,
            DisplayAsIcon=False,
        )
        settings = chart.AnimationSettings
        settings.Animate = True
        settings.EntryEffect = 3844
        settings.ChartUnitEffect = 1
        presentation.SaveAs(str(destination), 1)
        return {
            "path": str(destination),
            "powerpoint_version": str(app.Version),
            "slide_count": int(presentation.Slides.Count),
            "chart_unit_effect": int(settings.ChartUnitEffect),
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
        default=Path("tests/fixtures/visual_chart_animation.ppt"),
    )
    args = parser.parse_args()
    print(make_chart_animation_fixture(args.output))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
