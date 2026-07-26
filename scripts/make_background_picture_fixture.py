"""Create a minimal legacy PPT with an embedded slide-background picture."""
from __future__ import annotations

import argparse
from pathlib import Path
import struct
import tempfile
import zlib

try:
    import pythoncom
    import win32com.client
except ImportError as exc:  # pragma: no cover
    raise SystemExit(f"pywin32 is required: {exc}") from exc


PP_ALERTS_NONE = 1
PP_LAYOUT_BLANK = 12
PP_SAVE_AS_PRESENTATION = 1
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3


def _write_background_png(path: Path, width: int = 240, height: int = 180) -> None:
    """Write a deterministic RGB PNG without adding a runtime dependency."""
    rows = bytearray()
    for y in range(height):
        rows.append(0)
        for x in range(width):
            if x < width // 2 and y < height // 2:
                red, green, blue = 24, 76, 132
            elif x >= width // 2 and y < height // 2:
                red, green, blue = 224, 111, 48
            elif x < width // 2:
                red, green, blue = 112, 173, 71
            else:
                red, green, blue = 91, 155, 213
            if abs(x * height - y * width) < width * 2:
                red, green, blue = 255, 255, 255
            rows.extend((red, green, blue))

    def chunk(kind: bytes, data: bytes) -> bytes:
        body = kind + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body))

    path.write_bytes(
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(bytes(rows), 9))
        + chunk(b"IEND", b"")
    )


def make_background_picture_fixture(destination: Path) -> dict[str, object]:
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
        slide.FollowMasterBackground = False
        with tempfile.TemporaryDirectory(prefix="ppt2pptx-background-fixture-") as directory:
            image = Path(directory) / "background.png"
            _write_background_png(image)
            slide.Background.Fill.UserPicture(str(image))
            title = slide.Shapes.AddTextbox(1, 90, 225, 540, 90)
            title.Name = "Editable foreground text"
            title.TextFrame.TextRange.Text = "EDITABLE BACKGROUND TEST"
            title.TextFrame.TextRange.Font.Name = "Arial"
            title.TextFrame.TextRange.Font.Size = 26
            title.TextFrame.TextRange.Font.Bold = True
            title.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF
            title.TextFrame.TextRange.ParagraphFormat.Alignment = 2
            presentation.SaveAs(str(destination), PP_SAVE_AS_PRESENTATION)
        return {
            "path": str(destination),
            "powerpoint_version": str(app.Version),
            "slide_count": int(presentation.Slides.Count),
            "slide_width_pt": float(presentation.PageSetup.SlideWidth),
            "slide_height_pt": float(presentation.PageSetup.SlideHeight),
            "background_type": int(slide.Background.Fill.Type),
            "foreground_text_box_count": 1,
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
        default=Path("tests/fixtures/visual_background_picture.ppt"),
    )
    args = parser.parse_args()
    print(make_background_picture_fixture(args.output))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
