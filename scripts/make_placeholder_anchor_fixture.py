"""Create a one-slide PPT whose title inherits bottom anchoring."""
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
from ppt2pptx.ppt import (
    CONTAINER_VERSION,
    RT_DOCUMENT,
    RT_OFFICEART_FOPT,
    _iter_all_records,
    _slide_entries,
    persist_directory,
    records,
)


PP_SAVE_AS_PRESENTATION = 1
PP_LAYOUT_TITLE_ONLY = 11
MSO_ANCHOR_BOTTOM = 4
MSO_FALSE = 0


def _clear_slide_anchor_override(path: Path) -> dict[str, int]:
    raw = bytearray(path.read_bytes())
    compound = CompoundFile(bytes(raw))
    stream = compound.open_stream("PowerPoint Document")
    roots = list(records(stream))
    document = next(
        record
        for record in reversed(roots)
        if record.type == RT_DOCUMENT
        and record.version == CONTAINER_VERSION
    )
    slide_reference, _slide_id = _slide_entries(document)[0]
    slide_start = persist_directory(stream)[slide_reference]
    slide_record = next(records(stream, slide_start))
    slide_end = slide_start + 8 + len(slide_record.payload)

    all_records = tuple(_iter_all_records(stream))
    property_offset = None
    target_fopt = None
    for record in all_records:
        if not (
            record.type == RT_OFFICEART_FOPT
            and slide_start < record.offset < slide_end
        ):
            continue
        for index in range(min(record.instance, len(record.payload) // 6)):
            opid = struct.unpack_from("<H", record.payload, index * 6)[0]
            if opid & 0x3FFF == 135:
                if property_offset is not None:
                    raise RuntimeError("slide contains multiple anchorText overrides")
                property_offset = record.offset + 8 + index * 6
                target_fopt = record
    if property_offset is None or target_fopt is None:
        raise RuntimeError("slide title has no anchorText override to clear")

    complex_index = None
    complex_length = None
    complex_property_id = None
    for index in range(min(target_fopt.instance, len(target_fopt.payload) // 6)):
        opid, value = struct.unpack_from("<HI", target_fopt.payload, index * 6)
        if opid & 0x8000:
            if complex_index is not None:
                raise RuntimeError(
                    "slide title has multiple complex padding targets"
                )
            complex_index = index
            complex_length = value
            complex_property_id = opid & 0x3FFF
    if complex_index is None or complex_length is None:
        raise RuntimeError("slide title has no complex property padding target")

    stream_bytes = bytearray(stream)
    del stream_bytes[property_offset:property_offset + 6]
    fopt_payload_end = target_fopt.offset + 8 + len(target_fopt.payload) - 6
    stream_bytes[fopt_payload_end:fopt_payload_end] = b"\0" * 6
    fopt_options = struct.unpack_from("<H", stream_bytes, target_fopt.offset)[0]
    fopt_version = fopt_options & 0x000F
    fopt_instance = fopt_options >> 4
    if fopt_instance < 1:
        raise RuntimeError("OfficeArtFOPT property count is invalid")
    struct.pack_into(
        "<H",
        stream_bytes,
        target_fopt.offset,
        ((fopt_instance - 1) << 4) | fopt_version,
    )
    anchor_index = (property_offset - target_fopt.offset - 8) // 6
    if complex_index > anchor_index:
        complex_index -= 1
    struct.pack_into(
        "<I",
        stream_bytes,
        target_fopt.offset + 8 + complex_index * 6 + 2,
        complex_length + 6,
    )
    if len(stream_bytes) != len(stream):
        raise RuntimeError("fixture record rewrite changed stream length")

    kind, first_sector, stream_size = compound.entries["PowerPoint Document"]
    if kind != 2 or stream_size < compound.mini_cutoff:
        raise RuntimeError("PowerPoint Document stream is not a regular FAT stream")
    chain = compound._chain(first_sector, compound.fat)
    for stream_offset in range(0, len(stream_bytes), compound.sector_size):
        sector_number = stream_offset // compound.sector_size
        physical_offset = (chain[sector_number] + 1) * compound.sector_size
        chunk = stream_bytes[
            stream_offset:stream_offset + compound.sector_size
        ]
        raw[physical_offset:physical_offset + len(chunk)] = chunk
    path.write_bytes(raw)
    return {
        "stream_offset": property_offset,
        "removed_property_id": 135,
        "padding_property_id": complex_property_id,
    }


def make_placeholder_anchor_fixture(destination: Path) -> dict[str, object]:
    destination = destination.resolve()
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        destination.unlink()

    pythoncom.CoInitialize()
    app = None
    presentation = None
    owned_pid = None
    try:
        app, owned_pid = _launch_powerpoint()
        presentation = app.Presentations.Add(True)
        presentation.PageSetup.SlideWidth = 720
        presentation.PageSetup.SlideHeight = 540
        slide = presentation.Slides.Add(1, PP_LAYOUT_TITLE_ONLY)
        title = slide.Shapes.Title
        title.TextFrame.TextRange.Text = "Inherited bottom anchor"

        master_title = presentation.SlideMaster.Shapes.Title
        master_title.Left = 60
        master_title.Top = 60
        master_title.Width = 600
        master_title.Height = 240
        master_title.TextFrame.VerticalAnchor = MSO_ANCHOR_BOTTOM
        master_title.TextFrame.TextRange.Font.Name = "Arial"
        master_title.TextFrame.TextRange.Font.Size = 40
        master_title.Fill.Visible = MSO_FALSE
        master_title.Line.Visible = MSO_FALSE

        presentation.SaveAs(str(destination), PP_SAVE_AS_PRESENTATION)
        result = {
            "path": str(destination),
            "powerpoint_version": str(app.Version),
            "owned_pid": owned_pid,
            "slide_count": int(presentation.Slides.Count),
            "slide_width_pt": float(presentation.PageSetup.SlideWidth),
            "slide_height_pt": float(presentation.PageSetup.SlideHeight),
            "title_left_pt": float(title.Left),
            "title_top_pt": float(title.Top),
            "title_width_pt": float(title.Width),
            "title_height_pt": float(title.Height),
            "title_vertical_anchor": int(title.TextFrame.VerticalAnchor),
        }
        presentation.Close()
        presentation = None
        result["binary_patch"] = _clear_slide_anchor_override(destination)
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
        default=Path("tests/fixtures/visual_placeholder_anchor.ppt"),
    )
    args = parser.parse_args()
    print(make_placeholder_anchor_fixture(args.output))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
