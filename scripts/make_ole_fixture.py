"""Create a minimal PowerPoint 97-2003 fixture with one controlled Package OLE object."""
from __future__ import annotations

import argparse
from pathlib import Path
import tempfile

try:
    import pythoncom
    import win32com.client
except ImportError as exc:  # pragma: no cover
    raise SystemExit(f"pywin32 is required: {exc}") from exc


PP_ALERTS_NONE = 1
PP_LAYOUT_BLANK = 12
PP_SAVE_AS_PRESENTATION = 1
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3


def make_ole_fixture(destination: Path) -> dict[str, object]:
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
        with tempfile.TemporaryDirectory(prefix="ppt2pptx-ole-fixture-") as directory:
            payload = Path(directory) / "fixture.txt"
            payload.write_text(
                "ppt2pptx controlled embedded Package fixture\n",
                encoding="utf-8",
            )
            shape = slide.Shapes.AddOLEObject(
                120,
                90,
                320,
                180,
                FileName=str(payload),
                Link=False,
                DisplayAsIcon=True,
                IconLabel="fixture.txt",
            )
            shape.Name = "Controlled embedded Package"
            presentation.SaveAs(str(destination), PP_SAVE_AS_PRESENTATION)
        return {
            "path": str(destination),
            "powerpoint_version": str(app.Version),
            "slide_count": int(presentation.Slides.Count),
            "ole_object_count": 1,
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
        default=Path("tests/fixtures/visual_ole.ppt"),
    )
    args = parser.parse_args()
    print(make_ole_fixture(args.output))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
