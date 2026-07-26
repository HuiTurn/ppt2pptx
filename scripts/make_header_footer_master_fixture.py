"""Create a one-slide PPT with a styled inherited master date field."""
from __future__ import annotations

import argparse
from pathlib import Path
import struct

try:
    import pythoncom
except ImportError as exc:  # pragma: no cover
    raise SystemExit(f"pywin32 is required: {exc}") from exc

from make_master_objects_fixture import _launch_powerpoint
from ppt2pptx.cfb import CompoundFile
from ppt2pptx.ppt import CONTAINER_VERSION, RT_DOCUMENT, records


PP_ALIGN_LEFT = 1
PP_ALIGN_CENTER = 2
PP_LAYOUT_BLANK = 12
PP_PLACEHOLDER_SLIDE_NUMBER = 13
PP_PLACEHOLDER_FOOTER = 15
PP_PLACEHOLDER_DATE = 16
PP_SAVE_AS_PRESENTATION = 1
MSO_ANCHOR_BOTTOM = 4
MSO_FALSE = 0
RT_HEADERS_FOOTERS = 4057
RT_HEADERS_FOOTERS_ATOM = 4058
RT_CSTRING = 4026


def _placeholder_type(shape) -> int | None:
    try:
        return int(shape.PlaceholderFormat.Type)
    except Exception:
        return None


def _placeholder(master, placeholder_type: int):
    for index in range(1, master.Shapes.Count + 1):
        shape = master.Shapes(index)
        if _placeholder_type(shape) == placeholder_type:
            return shape
    raise RuntimeError(f"missing master placeholder type {placeholder_type}")


def _write_regular_stream(raw: bytearray, compound: CompoundFile, data: bytes) -> None:
    kind, first_sector, stream_size = compound.entries["PowerPoint Document"]
    if kind != 2 or stream_size < compound.mini_cutoff or len(data) != stream_size:
        raise RuntimeError("PowerPoint Document is not a fixed-size FAT stream")
    chain = compound._chain(first_sector, compound.fat)
    for stream_offset in range(0, len(data), compound.sector_size):
        sector_number = stream_offset // compound.sector_size
        physical_offset = (chain[sector_number] + 1) * compound.sector_size
        chunk = data[stream_offset:stream_offset + compound.sector_size]
        raw[physical_offset:physical_offset + len(chunk)] = chunk


def _inject_legacy_date(path: Path) -> dict[str, object]:
    raw = bytearray(path.read_bytes())
    compound = CompoundFile(bytes(raw))
    stream = bytearray(compound.open_stream("PowerPoint Document"))
    roots = list(records(stream))
    document = next(
        record
        for record in reversed(roots)
        if record.type == RT_DOCUMENT
        and record.version == CONTAINER_VERSION
    )
    children = list(
        records(
            stream,
            document.offset + 8,
            document.offset + 8 + len(document.payload),
        )
    )
    slide_fields = next(
        record
        for record in children
        if record.type == RT_HEADERS_FOOTERS and record.instance == 3
    )
    notes_fields = next(
        record
        for record in children
        if record.type == RT_HEADERS_FOOTERS and record.instance == 4
    )
    if (
        slide_fields.offset + 8 + len(slide_fields.payload) != notes_fields.offset
        or notes_fields.offset + 8 + len(notes_fields.payload)
        - slide_fields.offset != 40
    ):
        raise RuntimeError("fixture header/footer records are not adjacent 20-byte records")

    date = "DATE42".encode("utf-16le")
    replacement = (
        struct.pack(
            "<HHI",
            (3 << 4) | CONTAINER_VERSION,
            RT_HEADERS_FOOTERS,
            32,
        )
        + struct.pack("<HHIHH", 0, RT_HEADERS_FOOTERS_ATOM, 4, 0, 0x05)
        + struct.pack("<HHI", 0, RT_CSTRING, len(date))
        + date
    )
    if len(replacement) != 40:
        raise RuntimeError("legacy footer replacement must be exactly 40 bytes")
    stream[slide_fields.offset:slide_fields.offset + 40] = replacement
    _write_regular_stream(raw, compound, bytes(stream))
    path.write_bytes(raw)
    return {
        "stream_offset": slide_fields.offset,
        "date_text": "DATE42",
        "replacement_length": len(replacement),
    }


def make_header_footer_master_fixture(destination: Path) -> dict[str, object]:
    destination = destination.resolve()
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        destination.unlink()

    pythoncom.CoInitialize()
    app = None
    presentation = None
    owned_pid = None
    result: dict[str, object] = {}
    try:
        app, owned_pid = _launch_powerpoint()
        presentation = app.Presentations.Add(True)
        presentation.PageSetup.SlideWidth = 720
        presentation.PageSetup.SlideHeight = 540
        slide = presentation.Slides.Add(1, PP_LAYOUT_BLANK)
        master = presentation.SlideMaster
        master.HeadersFooters.DateAndTime.Visible = True
        master.HeadersFooters.DateAndTime.UseFormat = False
        master.HeadersFooters.DateAndTime.Text = "D"
        master.HeadersFooters.Footer.Visible = True
        master.HeadersFooters.Footer.Text = "F"
        retained_types = {
            PP_PLACEHOLDER_DATE,
            PP_PLACEHOLDER_FOOTER,
            PP_PLACEHOLDER_SLIDE_NUMBER,
        }
        for index in range(master.Shapes.Count, 0, -1):
            if _placeholder_type(master.Shapes(index)) not in retained_types:
                master.Shapes(index).Delete()

        date = _placeholder(master, PP_PLACEHOLDER_DATE)
        date.Left = 6
        date.Top = 510
        date.Width = 150
        date.Height = 36
        date.Fill.Visible = MSO_FALSE
        date.Line.Visible = MSO_FALSE
        date.TextFrame.VerticalAnchor = MSO_ANCHOR_BOTTOM
        date.TextFrame.TextRange.ParagraphFormat.Alignment = PP_ALIGN_LEFT
        date.TextFrame.TextRange.Font.Name = "Tahoma"
        date.TextFrame.TextRange.Font.Size = 14
        date.TextFrame.TextRange.Font.Color.RGB = 0x8C4540

        footer = _placeholder(master, PP_PLACEHOLDER_FOOTER)
        footer.Left = 228
        footer.Top = 510
        footer.Width = 228
        footer.Height = 36
        footer.Fill.Visible = MSO_FALSE
        footer.Line.Visible = MSO_FALSE
        footer.TextFrame.VerticalAnchor = MSO_ANCHOR_BOTTOM
        footer.TextFrame.TextRange.ParagraphFormat.Alignment = PP_ALIGN_CENTER
        footer.TextFrame.TextRange.Font.Name = "Tahoma"
        footer.TextFrame.TextRange.Font.Size = 14
        footer.TextFrame.TextRange.Font.Color.RGB = 0x8C4540

        presentation.SaveAs(str(destination), PP_SAVE_AS_PRESENTATION)
        result = {
            "path": str(destination),
            "powerpoint_version": str(app.Version),
            "owned_pid": owned_pid,
            "slide_count": int(presentation.Slides.Count),
            "slide_width_pt": float(presentation.PageSetup.SlideWidth),
            "slide_height_pt": float(presentation.PageSetup.SlideHeight),
            "slide_shape_count_before_patch": int(slide.Shapes.Count),
            "master_shape_count": int(master.Shapes.Count),
            "footer": [
                float(footer.Left),
                float(footer.Top),
                float(footer.Width),
                float(footer.Height),
            ],
            "date": [
                float(date.Left),
                float(date.Top),
                float(date.Width),
                float(date.Height),
            ],
        }
        presentation.Close()
        presentation = None
        result["binary_patch"] = _inject_legacy_date(destination)
        return result
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
        default=Path("tests/fixtures/visual_header_footer_master.ppt"),
    )
    args = parser.parse_args()
    print(make_header_footer_master_fixture(args.output))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
