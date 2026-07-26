"""Create a one-slide PPT with Wingdings-backed editable bullets."""
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
MSO_TEXT_ORIENTATION_HORIZONTAL = 1
MSO_FALSE = 0
MSO_TRUE = -1


def make_bullet_font_fixture(destination: Path) -> dict[str, object]:
    destination = destination.resolve()
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        destination.unlink()

    pythoncom.CoInitialize()
    app = None
    presentation = None
    try:
        app, owned_pid = _launch_powerpoint()
        presentation = app.Presentations.Add(True)
        presentation.PageSetup.SlideWidth = 720
        presentation.PageSetup.SlideHeight = 540
        slide = presentation.Slides.Add(1, PP_LAYOUT_BLANK)
        for shape_index in range(
            presentation.SlideMaster.Shapes.Count, 0, -1
        ):
            presentation.SlideMaster.Shapes(shape_index).Delete()

        shape = slide.Shapes.AddTextbox(
            MSO_TEXT_ORIENTATION_HORIZONTAL, 72, 72, 576, 360
        )
        shape.Fill.Visible = MSO_FALSE
        shape.Line.Visible = MSO_FALSE
        text = shape.TextFrame.TextRange
        text.Text = "Major item\rMinor item\rSecond major"
        text.Font.Name = "Arial"
        text.Font.Size = 30

        expected = (
            ("w", "Wingdings", 1.1, 0xF7896F),
            ("n", "Wingdings", 0.6, 0x8C4540),
            ("w", "Wingdings", 1.1, 0xF7896F),
        )
        for index, (
            character, typeface, relative_size, color
        ) in enumerate(expected, 1):
            paragraph = text.Paragraphs(index)
            paragraph.ParagraphFormat.Bullet.Visible = MSO_TRUE
            paragraph.ParagraphFormat.Bullet.Character = ord(character)
            paragraph.ParagraphFormat.Bullet.Font.Name = typeface
            paragraph.ParagraphFormat.Bullet.RelativeSize = relative_size
            paragraph.ParagraphFormat.Bullet.Font.Color.RGB = color

        presentation.SaveAs(str(destination), PP_SAVE_AS_PRESENTATION)
        return {
            "path": str(destination),
            "powerpoint_version": str(app.Version),
            "owned_pid": owned_pid,
            "slide_count": int(presentation.Slides.Count),
            "slide_width_pt": float(presentation.PageSetup.SlideWidth),
            "slide_height_pt": float(presentation.PageSetup.SlideHeight),
            "paragraphs": [
                {
                    "character": chr(
                        text.Paragraphs(index).ParagraphFormat.Bullet.Character
                    ),
                    "typeface": str(
                        text.Paragraphs(index).ParagraphFormat.Bullet.Font.Name
                    ),
                    "relative_size": float(
                        text.Paragraphs(
                            index
                        ).ParagraphFormat.Bullet.RelativeSize
                    ),
                    "color": int(
                        text.Paragraphs(
                            index
                        ).ParagraphFormat.Bullet.Font.Color.RGB
                    ),
                }
                for index in range(1, 4)
            ],
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
        default=Path("tests/fixtures/visual_bullet_font.ppt"),
    )
    args = parser.parse_args()
    print(make_bullet_font_fixture(args.output))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
