"""Create a one-slide PPT containing one legacy master arc."""
from __future__ import annotations

import argparse
from pathlib import Path
import struct
import tempfile

import pythoncom

from ppt2pptx.cfb import CompoundFile

from make_zero_extent_line_fixture import (
    PP_LAYOUT_BLANK,
    PP_SAVE_AS_PRESENTATION,
    _download_seed,
    _launch_powerpoint,
)


MSO_FREEFORM = 5
MSO_GROUP = 6
MSO_FALSE = 0


def _direct_records(data: bytes, offset: int):
    _version, _record_type, length = struct.unpack_from("<HHI", data, offset)
    cursor, end = offset + 8, offset + 8 + length
    while cursor + 8 <= end:
        version, record_type, child_length = struct.unpack_from(
            "<HHI", data, cursor
        )
        if cursor + 8 + child_length > end:
            break
        yield (
            cursor,
            version & 0xF,
            version >> 4,
            record_type,
            child_length,
        )
        cursor += 8 + child_length


def _all_records(data: bytes, start: int = 0, end: int | None = None):
    cursor = start
    limit = len(data) if end is None else end
    while cursor + 8 <= limit:
        version, record_type, length = struct.unpack_from(
            "<HHI", data, cursor
        )
        payload_end = cursor + 8 + length
        if payload_end > limit:
            break
        yield cursor, version & 0xF, version >> 4, record_type, length
        if version & 0xF == 0xF:
            yield from _all_records(data, cursor + 8, payload_end)
        cursor = payload_end


def _shape_containers(data: bytes):
    for offset, _version, _instance, record_type, length in _all_records(data):
        if record_type != 0xF004:
            continue
        children = list(_direct_records(data, offset))
        shape = next(
            (child for child in children if child[3] == 0xF00A),
            None,
        )
        anchor = next(
            (child for child in children if child[3] in (0xF00F, 0xF010)),
            None,
        )
        properties = next(
            (child for child in children if child[3] == 0xF00B),
            None,
        )
        if shape is not None:
            yield offset, length + 8, shape, anchor, properties


def _inject_legacy_arc(destination: Path, source_blob: bytes) -> int:
    source = CompoundFile(source_blob)
    source_document = source.open_stream("PowerPoint Document")
    target = CompoundFile.from_path(destination)
    target_document = bytearray(target.open_stream("PowerPoint Document"))

    target_candidates = [
        value
        for value in _shape_containers(target_document)
        if (
            value[2][2] == 0
            and value[3] is not None
            and value[3][3] == 0xF00F
            and value[4] is not None
            and value[4][4] > 500
        )
    ]
    if not target_candidates:
        raise RuntimeError("generated fixture contains no grouped freeform arc")
    first_anchor = target_candidates[0][3]
    assert first_anchor is not None
    anchor_bytes = target_document[
        first_anchor[0] + 8:first_anchor[0] + 8 + first_anchor[4]
    ]

    source_arc = next(
        (
            value
            for value in _shape_containers(source_document)
            if (
                value[2][2] == 19
                and value[3] is not None
                and source_document[
                    value[3][0] + 8:value[3][0] + 8 + value[3][4]
                ] == anchor_bytes
            )
        ),
        None,
    )
    if source_arc is None:
        raise RuntimeError("Apache POI seed contains no matching legacy arc")
    source_offset, source_size, source_shape, _anchor, _properties = source_arc
    source_record = bytearray(
        source_document[source_offset:source_offset + source_size]
    )
    source_shape_relative = source_shape[0] - source_offset

    rewritten = 0
    for offset, size, shape, anchor, properties in target_candidates:
        assert anchor is not None
        if (
            target_document[
                anchor[0] + 8:anchor[0] + 8 + anchor[4]
            ] != anchor_bytes
        ):
            continue
        replacement = bytearray(source_record)
        target_shape_id = struct.unpack_from(
            "<I", target_document, shape[0] + 8
        )[0]
        struct.pack_into(
            "<I", replacement, source_shape_relative + 8, target_shape_id
        )
        struct.pack_into("<I", replacement, 4, size - 8)
        padding = size - len(replacement)
        if padding < 8:
            raise RuntimeError("generated arc record is too small for injection")
        replacement.extend(struct.pack("<HHI", 0, 0, padding - 8))
        replacement.extend(bytes(padding - 8))
        target_document[offset:offset + size] = replacement
        rewritten += 1
    if not rewritten:
        raise RuntimeError("no generated arc records were rewritten")

    # PowerPoint retains the copied group's non-rendering custom geometry in
    # incremental revisions. Keep those hidden records structurally harmless
    # without deleting or resizing the stream: mark them as rectangles and
    # rename their pVertices/pSegmentInfo property IDs so diagnostics see
    # exactly the visible arc.
    for _offset, _size, shape, _anchor, properties in _shape_containers(
        target_document
    ):
        if shape[2] != 0 or properties is None:
            continue
        shape_header = struct.unpack_from("<H", target_document, shape[0])[0]
        struct.pack_into(
            "<H",
            target_document,
            shape[0],
            (shape_header & 0x000F) | (1 << 4),
        )
        count = min(properties[2], properties[4] // 6)
        for index in range(count):
            entry = properties[0] + 8 + index * 6
            opid = struct.unpack_from("<H", target_document, entry)[0]
            pid = opid & 0x3FFF
            if pid == 325:
                struct.pack_into(
                    "<H", target_document, entry, (opid & 0xC000) | 0x3FFE
                )
            elif pid == 326:
                struct.pack_into(
                    "<H", target_document, entry, (opid & 0xC000) | 0x3FFD
                )

    kind, first_sector, stream_size = target.by_name["powerpoint document"]
    if kind != 2 or stream_size != len(target_document):
        raise RuntimeError("PowerPoint Document stream is not rewriteable")
    file_bytes = bytearray(destination.read_bytes())
    cursor = 0
    for sector in target._chain(first_sector, target.fat):
        count = min(target.sector_size, len(target_document) - cursor)
        if count <= 0:
            break
        physical = (sector + 1) * target.sector_size
        file_bytes[physical:physical + count] = target_document[
            cursor:cursor + count
        ]
        cursor += count
    if cursor != len(target_document):
        raise RuntimeError("PowerPoint Document sector chain is truncated")
    destination.write_bytes(file_bytes)
    return rewritten


def make_master_arc_fixture(destination: Path) -> dict[str, object]:
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
        with tempfile.TemporaryDirectory(
            prefix="ppt2pptx-master-arc-fixture-"
        ) as directory:
            seed = Path(directory) / "apache-poi-37625.ppt"
            seed.write_bytes(_download_seed())
            source_presentation = app.Presentations.Open(
                str(seed),
                ReadOnly=True,
                Untitled=True,
                WithWindow=False,
            )
            source_master = source_presentation.SlideMaster
            source_group = None
            source_arc = None
            for shape_index in range(1, source_master.Shapes.Count + 1):
                shape = source_master.Shapes(shape_index)
                if int(shape.Type) != MSO_GROUP:
                    continue
                for item_index in range(1, shape.GroupItems.Count + 1):
                    item = shape.GroupItems(item_index)
                    if (
                        int(item.Type) == MSO_FREEFORM
                        and str(item.Name).casefold().startswith("arc")
                    ):
                        source_group = shape
                        source_arc = item
                        break
                if source_arc is not None:
                    break
            if source_arc is None:
                raise RuntimeError("Apache POI seed contains no master arc")

            presentation = app.Presentations.Add(True)
            presentation.PageSetup.SlideWidth = 720
            presentation.PageSetup.SlideHeight = 540
            presentation.Slides.Add(1, PP_LAYOUT_BLANK)
            master = presentation.SlideMaster
            for shape_index in range(master.Shapes.Count, 0, -1):
                master.Shapes(shape_index).Delete()

            assert source_group is not None
            source_group.Copy()
            pasted = master.Shapes.Paste()
            if pasted.Count != 1:
                raise RuntimeError("expected one pasted master group")
            group = pasted(1)
            for item_index in range(1, group.GroupItems.Count + 1):
                item = group.GroupItems(item_index)
                if str(item.Name) == str(source_arc.Name):
                    continue
                try:
                    item.Line.Visible = MSO_FALSE
                except Exception:
                    pass
                try:
                    item.Fill.Visible = MSO_FALSE
                except Exception:
                    pass

            presentation.SaveAs(str(destination), PP_SAVE_AS_PRESENTATION)
            group = master.Shapes(1)
            arc = next(
                group.GroupItems(index)
                for index in range(1, group.GroupItems.Count + 1)
                if str(group.GroupItems(index).Name).casefold().startswith("arc")
            )
            result = {
                "path": str(destination),
                "powerpoint_version": str(app.Version),
                "owned_pid": owned_pid,
                "slide_count": int(presentation.Slides.Count),
                "slide_width_pt": float(presentation.PageSetup.SlideWidth),
                "slide_height_pt": float(presentation.PageSetup.SlideHeight),
                "master_shape_count": int(master.Shapes.Count),
                "master_group_item_count": int(group.GroupItems.Count),
                "master_shape_type": int(arc.Type),
                "master_auto_shape_type": int(arc.AutoShapeType),
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
    result["rewritten_arc_records"] = _inject_legacy_arc(
        destination, _download_seed()
    )
    return result


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("tests/fixtures/visual_master_arc.ppt"),
    )
    args = parser.parse_args()
    print(make_master_arc_fixture(args.output))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
