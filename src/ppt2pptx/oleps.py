"""Bounded parser for core metadata in an OLE SummaryInformation stream."""
from __future__ import annotations

from datetime import datetime, timedelta, timezone
import struct

from .errors import InvalidPpt
from .ppt import CoreProperties

FMTID = bytes.fromhex("e0859ff2f94f6810ab9108002b27b3d9")
STRING_FIELDS = {2: "title", 3: "subject", 4: "creator", 5: "keywords", 6: "description", 8: "last_modified_by", 9: "revision"}
TIME_FIELDS = {11: "last_printed", 12: "created", 13: "modified"}

def _range(data: bytes, start: int, size: int, end: int, label: str) -> None:
    if start < 0 or size < 0 or end > len(data) or start > end - size:
        raise InvalidPpt(f"{label} exceeds the SummaryInformation property set")

def _type(data: bytes, start: int, end: int) -> int:
    _range(data, start, 4, end, "property type")
    kind, padding = struct.unpack_from("<HH", data, start)
    if padding:
        raise InvalidPpt("SummaryInformation property has invalid type padding")
    return kind

def _string(data: bytes, start: int, end: int, codepage: int) -> str | None:
    kind = _type(data, start, end)
    if kind not in (0x1E, 0x1F):
        return None
    _range(data, start + 4, 4, end, "string length")
    count = struct.unpack_from("<I", data, start + 4)[0]
    wide = kind == 0x1F or codepage in (1200, 1201)
    size = count * (2 if wide else 1)
    _range(data, start + 8, size, end, "string value")
    raw = data[start + 8:start + 8 + size]
    if wide:
        while raw.endswith(b"\0\0"):
            raw = raw[:-2]
    else:
        # PowerPoint 7 occasionally counts all alignment padding bytes as
        # part of an LPSTR value instead of counting only one terminator.
        raw = raw.rstrip(b"\0")
    codec = "utf-16le" if kind == 0x1F else "utf-16be" if codepage == 1201 else "utf-16le" if codepage == 1200 else "utf-8" if codepage == 65001 else f"cp{codepage}"
    try:
        value = raw.decode(codec, "replace")
    except LookupError:
        value = raw.decode("cp1252", "replace")
    value = "".join(char if ord(char) in (9, 10, 13) or ord(char) >= 32 else "�" for char in value)
    return value or None

def _time(data: bytes, start: int, end: int) -> str | None:
    if _type(data, start, end) != 0x40:
        return None
    _range(data, start + 4, 8, end, "FILETIME value")
    ticks = struct.unpack_from("<Q", data, start + 4)[0]
    if not ticks:
        return None
    try:
        value = datetime(1601, 1, 1, tzinfo=timezone.utc) + timedelta(microseconds=ticks // 10)
    except OverflowError as exc:
        raise InvalidPpt("SummaryInformation FILETIME is invalid") from exc
    return value.isoformat(timespec="seconds").replace("+00:00", "Z")

def read_summary_information(data: bytes) -> CoreProperties:
    if len(data) < 48 or len(data) > 2 * 1024 * 1024:
        raise InvalidPpt("SummaryInformation stream has an invalid size")
    byte_order, version = struct.unpack_from("<HH", data)
    if byte_order != 0xFFFE or version not in (0, 1):
        raise InvalidPpt("SummaryInformation header is invalid")
    count = struct.unpack_from("<I", data, 24)[0]
    if count not in (1, 2):
        raise InvalidPpt("SummaryInformation property-set count is invalid")
    _range(data, 28, count * 20, len(data), "property-set descriptors")
    set_offset = None
    for index in range(count):
        position = 28 + index * 20
        if data[position:position + 16] == FMTID:
            set_offset = struct.unpack_from("<I", data, position + 16)[0]
    if set_offset is None or set_offset % 4:
        raise InvalidPpt("SummaryInformation format identifier is absent")
    _range(data, set_offset, 8, len(data), "property set")
    set_size, property_count = struct.unpack_from("<II", data, set_offset)
    if set_size < 8 or set_offset > len(data) - set_size or property_count > (set_size - 8) // 8:
        raise InvalidPpt("SummaryInformation property table is invalid")
    set_end, table_end = set_offset + set_size, set_offset + 8 + property_count * 8
    offsets: dict[int, int] = {}
    for index in range(property_count):
        prop_id, relative = struct.unpack_from("<II", data, set_offset + 8 + index * 8)
        if prop_id in offsets or relative % 4 or set_offset + relative < table_end or set_offset + relative >= set_end:
            raise InvalidPpt("SummaryInformation property offset is invalid")
        offsets[prop_id] = set_offset + relative
    ordered = sorted((position, prop_id) for prop_id, position in offsets.items())
    bounds = {prop_id: (position, ordered[index + 1][0] if index + 1 < len(ordered) else set_end) for index, (position, prop_id) in enumerate(ordered)}
    codepage = 1252
    if 1 in bounds and _type(data, *bounds[1]) == 2:
        _range(data, bounds[1][0] + 4, 2, bounds[1][1], "codepage")
        codepage = struct.unpack_from("<H", data, bounds[1][0] + 4)[0] or 1252
    values: dict[str, str | None] = {}
    for prop_id, field in STRING_FIELDS.items():
        if prop_id in bounds:
            values[field] = _string(data, *bounds[prop_id], codepage)
    for prop_id, field in TIME_FIELDS.items():
        if prop_id in bounds:
            values[field] = _time(data, *bounds[prop_id])
    return CoreProperties(**values)
