"""Bounded parser for the portions of MS-PPT needed for editable slide text."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import math
import struct
import zlib

from .errors import InvalidPpt

RT_DOCUMENT = 1000
RT_SLIDE = 1006
RT_SLIDE_LIST_WITH_TEXT = 4080
RT_SLIDE_PERSIST_ATOM = 1011
RT_SLIDE_SHOW_SLIDE_INFO_ATOM = 0x03F9
RT_PERSIST_DIRECTORY_ATOM = 6002
RT_TEXT_CHARS_ATOM = 4000
RT_TEXT_BYTES_ATOM = 4008
RT_OUTLINE_TEXT_REF_ATOM = 3998
RT_DOCUMENT_ATOM = 1001
RT_OFFICEART_SPGR_CONTAINER = 0xF003
RT_OFFICEART_SP_CONTAINER = 0xF004
RT_OFFICEART_FSPGR = 0xF009
RT_OFFICEART_CLIENT_ANCHOR = 0xF010
RT_OFFICEART_CHILD_ANCHOR = 0xF00F
RT_OFFICEART_CLIENT_DATA = 0xF011
RT_OFFICEART_CLIENT_TEXTBOX = 0xF00D
RT_OFFICEART_BSE = 0xF007
RT_OFFICEART_FOPT = 0xF00B
RT_OE_PLACEHOLDER_ATOM = 3011
RT_ROUNDTRIP_OPAQUE_MIN = 1053
RT_ROUNDTRIP_OPAQUE_MAX = 1064
RT_SOUND_COLLECTION = 0x07E4
RT_SOUND = 0x07E6
RT_SOUND_DATA_BLOB = 0x07E7
RT_EXTERNAL_OBJECT_REF_ATOM = 0x0BC1
RT_EXTERNAL_OLE_OBJECT_ATOM = 0x0FC3
RT_EXTERNAL_OLE_EMBED = 0x0FCC
RT_EXTERNAL_OLE_LINK = 0x0FCE
RT_EXTERNAL_OLE_CONTROL = 0x0FEE
RT_ANIMATION_INFO_ATOM = 0x0FF1
RT_EXTERNAL_VIDEO = 0x1005
RT_EXTERNAL_AVI_MOVIE = 0x1006
RT_EXTERNAL_MCI_MOVIE = 0x1007
RT_EXTERNAL_OLE_OBJECT_STG = 0x1011
RT_ANIMATION_INFO = 0x1014
RT_PROG_BINARY_TAG = 0x138A
RT_BINARY_TAG_DATA_BLOB = 0x138B
RT_CHART_BUILD = 0x2B04
RT_CHART_BUILD_ATOM = 0x2B05
RT_DIAGRAM_BUILD = 0x2B06
RT_DIAGRAM_BUILD_ATOM = 0x2B07
RT_ROUNDTRIP_ANIMATION_ATOM = 0x2B0B
RT_ROUNDTRIP_ANIMATION_HASH_ATOM = 0x2B0D
RT_TIME_NODE = 0xF127
RT_TIME_PROPERTY_LIST = 0xF13D
RT_TIME_VARIANT = 0xF142
RT_TIME_EXT_TIME_NODE = 0xF144
RT_VISUAL_SHAPE_ATOM = 0x2AFB
RT_CSTRING = 0x0FBA
CONTAINER_VERSION = 0xF
DEFAULT_SLIDE_WIDTH = 5760
DEFAULT_SLIDE_HEIGHT = 4320
SHAPE_PRESETS = {
    1: "rect", 2: "roundRect", 3: "ellipse", 4: "diamond",
    5: "triangle", 6: "rtTriangle", 7: "parallelogram", 8: "trapezoid",
    9: "hexagon", 10: "octagon", 11: "plus", 12: "star5", 13: "rightArrow",
    19: "arc", 20: "line",
    32: "straightConnector1", 33: "bentConnector2",
    34: "bentConnector3", 35: "bentConnector4", 36: "bentConnector5",
    37: "curvedConnector2", 38: "curvedConnector3",
    39: "curvedConnector4", 40: "curvedConnector5",
    56: "pentagon", 57: "hexagon", 58: "heptagon", 59: "octagon",
    66: "star4", 67: "star5", 68: "star6", 69: "star8", 70: "star16",
    84: "bevel", 85: "leftBracket", 86: "rightBracket",
    87: "leftBrace", 88: "rightBrace", 89: "leftUpArrow",
    90: "bentUpArrow", 91: "bentArrow",
    125: "diamond", 183: "ellipse", 184: "moon",
}

@dataclass(frozen=True, slots=True)
class Record:
    offset: int; version: int; instance: int; type: int; payload: bytes

@dataclass(frozen=True, slots=True)
class TextBox:
    text: str
    left: int
    top: int
    width: int
    height: int
    runs: tuple["TextRun", ...] = ()
    paragraph_alignments: tuple[str | None, ...] = ()
    paragraph_bullets: tuple[bool, ...] = ()
    rotation: int = 0
    flip_horizontal: bool = False
    flip_vertical: bool = False
    fill_color: str | None = None
    line_color: str | None = None
    line_dash: str | None = None
    paragraph_levels: tuple[int, ...] = ()
    paragraph_bullet_chars: tuple[str | None, ...] = ()
    auto_fit: bool = False
    fit_shape_to_text: bool = False
    vertical_anchor: str | None = None
    preset: str = "rect"
    wrap_text: bool = True
    paragraph_left_margins: tuple[int | None, ...] = ()
    paragraph_indents: tuple[int | None, ...] = ()
    paragraph_line_spacings: tuple[int | None, ...] = ()
    paragraph_space_before: tuple[int | None, ...] = ()
    paragraph_space_after: tuple[int | None, ...] = ()
    inset_left: int | None = None
    inset_top: int | None = None
    inset_right: int | None = None
    inset_bottom: int | None = None
    is_placeholder: bool = False
    line_width: int | None = None
    default_tab_size: int | None = None
    tab_stops: tuple[tuple[int, str], ...] = ()
    fill_pattern: str | None = None
    fill_back_color: str | None = None
    line_head: tuple[str, str | None, str | None] | None = None
    line_tail: tuple[str, str | None, str | None] | None = None

@dataclass(frozen=True, slots=True)
class TextRun:
    text: str
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    font_size: int | None = None
    color: str | None = None
    typeface: str | None = None
    hyperlink: str | None = None
    baseline: int | None = None

@dataclass(frozen=True, slots=True)
class TextContent:
    text: str
    runs: tuple[TextRun, ...] = ()
    paragraph_alignments: tuple[str | None, ...] = ()
    paragraph_bullets: tuple[bool, ...] = ()
    paragraph_levels: tuple[int, ...] = ()
    paragraph_bullet_chars: tuple[str | None, ...] = ()
    paragraph_left_margins: tuple[int | None, ...] = ()
    paragraph_indents: tuple[int | None, ...] = ()
    paragraph_line_spacings: tuple[int | None, ...] = ()
    paragraph_space_before: tuple[int | None, ...] = ()
    paragraph_space_after: tuple[int | None, ...] = ()
    text_type: int = 4
    default_tab_size: int | None = None
    tab_stops: tuple[tuple[int, str], ...] = ()

@dataclass(frozen=True, slots=True)
class _MasterStyle:
    run: TextRun
    alignment: str | None
    bullet: bool | None
    bullet_char: str | None = None
    left_margin: int | None = None
    indent: int | None = None
    line_spacing: int | None = None
    space_before: int | None = None
    space_after: int | None = None

@dataclass(frozen=True, slots=True)
class _ParagraphProperties:
    alignment: str | None = None
    bullet: bool | None = None
    bullet_char: str | None = None
    left_margin: int | None = None
    indent: int | None = None
    line_spacing: int | None = None
    space_before: int | None = None
    space_after: int | None = None

@dataclass(frozen=True, slots=True)
class Picture:
    data: bytes
    extension: str
    content_type: str
    left: int
    top: int
    width: int
    height: int
    crop_left: int = 0
    crop_top: int = 0
    crop_right: int = 0
    crop_bottom: int = 0
    rotation: int = 0
    flip_horizontal: bool = False
    flip_vertical: bool = False
    transparent_color: str | None = None
    line_color: str | None = None
    line_dash: str | None = None
    line_width: int | None = None

@dataclass(frozen=True, slots=True)
class BasicShape:
    preset: str
    left: int
    top: int
    width: int
    height: int
    fill_color: str | None = None
    line_color: str | None = None
    rotation: int = 0
    flip_horizontal: bool = False
    flip_vertical: bool = False
    line_dash: str | None = None
    path: tuple[tuple[object, ...], ...] | None = None
    path_width: int = 21600
    path_height: int = 21600
    line_width: int | None = None
    fill_pattern: str | None = None
    fill_back_color: str | None = None
    line_head: tuple[str, str | None, str | None] | None = None
    line_tail: tuple[str, str | None, str | None] | None = None
    adjustments: tuple[int, ...] = ()

@dataclass(frozen=True, slots=True)
class Comment:
    author: str
    initials: str
    text: str
    left: int
    top: int
    created: str | None = None

@dataclass(frozen=True, slots=True)
class HeaderFooter:
    date_text: str | None = None
    date_is_auto: bool = False
    header_text: str | None = None
    footer_text: str | None = None
    show_slide_number: bool = False

@dataclass(frozen=True, slots=True)
class TableCell:
    text: str
    runs: tuple[TextRun, ...]
    left: int
    top: int
    width: int
    height: int
    fill_color: str | None = None
    row: int = 0
    col: int = 0

@dataclass(frozen=True, slots=True)
class Table:
    left: int
    top: int
    width: int
    height: int
    rows: int
    cols: int
    cells: tuple[TableCell, ...]  # row-major, len == rows * cols

@dataclass(frozen=True, slots=True)
class Slide:
    text_boxes: tuple[TextBox, ...]
    pictures: tuple[Picture, ...] = ()
    shapes: tuple[BasicShape, ...] = ()
    background_color: str | None = None
    comments: tuple[Comment, ...] = ()
    header_footer: HeaderFooter | None = None
    background_color_end: str | None = None
    notes: tuple[str, ...] = ()
    hidden: bool = False
    background_gradient_angle: int | None = None
    background_gradient_type: int | None = None
    tables: tuple[Table, ...] = ()
    excluded_offsets: frozenset[int] = frozenset()

@dataclass(frozen=True, slots=True)
class CoreProperties:
    title: str | None = None
    subject: str | None = None
    creator: str | None = None
    keywords: str | None = None
    description: str | None = None
    last_modified_by: str | None = None
    revision: str | None = None
    created: str | None = None
    modified: str | None = None
    last_printed: str | None = None

@dataclass(frozen=True, slots=True)
class Presentation:
    width: int
    height: int
    slides: tuple[Slide, ...]
    core_properties: CoreProperties = CoreProperties()
    excluded_offsets: frozenset[int] = frozenset()

@dataclass(frozen=True, slots=True)
class LossyFeatureLocation:
    slide_index: int | None
    record_type: int
    record_offset: int
    object_kind: str

@dataclass(frozen=True, slots=True)
class LossyFeature:
    code: str
    message: str
    count: int
    record_types: tuple[int, ...] = ()
    locations: tuple[LossyFeatureLocation, ...] = ()

def records(data: bytes, start: int = 0, end: int | None = None):
    end = len(data) if end is None else end
    cursor = start
    while cursor < end:
        # Trailing padding / truncated tails are common in real files; stop cleanly
        # once a well-formed prefix has already been consumed.
        if end - cursor < 8:
            if cursor == start:
                raise InvalidPpt("PowerPoint record header is truncated")
            return
        version_instance, record_type, length = struct.unpack_from("<HHI", data, cursor)
        payload_start, payload_end = cursor + 8, cursor + 8 + length
        if payload_end > end:
            if cursor == start:
                raise InvalidPpt("PowerPoint record extends beyond its container")
            return
        yield Record(cursor, version_instance & 0xF, version_instance >> 4, record_type, data[payload_start:payload_end])
        cursor = payload_end

def descendants(record: Record):
    if record.version != CONTAINER_VERSION: return
    # PowerPoint 2007+ round-trip blobs embed OOXML/ZIP, not nested PPT records.
    if RT_ROUNDTRIP_OPAQUE_MIN <= record.type <= RT_ROUNDTRIP_OPAQUE_MAX:
        return
    for child in records(record.payload):
        yield child
        yield from descendants(child)

def persist_directory(data: bytes) -> dict[int, int]:
    """Read the newest observed persist mappings.  A PPT may contain several saves."""
    mapping: dict[int, int] = {}
    for record in records(data):
        if record.type != RT_PERSIST_DIRECTORY_ATOM: continue
        cursor = 0
        while cursor + 4 <= len(record.payload):
            info = struct.unpack_from("<I", record.payload, cursor)[0]; cursor += 4
            count, first = info >> 20, info & 0xFFFFF
            if not count or cursor + count * 4 > len(record.payload): break
            for index in range(count):
                mapping[first + index] = struct.unpack_from("<I", record.payload, cursor)[0]
                cursor += 4
    return mapping

def _slide_entries(document: Record) -> list[tuple[int, int]]:
    entries: list[tuple[int, int]] = []
    for container in descendants(document):
        if container.type != RT_SLIDE_LIST_WITH_TEXT or container.instance != 0:
            continue
        for record in records(container.payload):
            if record.type != RT_SLIDE_PERSIST_ATOM or len(record.payload) < 12:
                continue
            # SlidePersistAtom stores persistIdRef first and the stable slideId
            # after flags and cTexts at byte offset 12.
            reference = struct.unpack_from("<I", record.payload, 0)[0]
            if reference:
                slide_id = (
                    struct.unpack_from("<I", record.payload, 12)[0]
                    if len(record.payload) >= 16
                    else 0
                )
                entries.append((reference, slide_id))
    return entries


def _slide_refs(document: Record) -> list[int]:
    return [reference for reference, _slide_id in _slide_entries(document)]

def _notes_refs(document: Record) -> list[int]:
    refs: list[int] = []
    for container in descendants(document):
        if container.type != RT_SLIDE_LIST_WITH_TEXT or container.instance != 2:
            continue
        refs.extend(struct.unpack_from("<I", record.payload)[0]
                    for record in records(container.payload)
                    if record.type == RT_SLIDE_PERSIST_ATOM and len(record.payload) >= 4)
    return refs

def _text(record: Record) -> str | None:
    if record.type == RT_TEXT_CHARS_ATOM:
        value = record.payload.decode("utf-16le", "replace").rstrip("\x00")
        return value.translate({
            0xF030: "0",
            0xF03D: "=",
            0xF03E: ">",
            0xF044: "Δ",
            0xF072: "ρ",
            0xF073: "σ",
            0xF0AE: "→",
            0xF0B3: "≥",
            0xF0BB: "≈",
            0xF0BC: "…",
            0xF0DE: "⇒",
        })
    if record.type == RT_TEXT_BYTES_ATOM:
        return record.payload.decode("cp1252", "replace").rstrip("\x00")
    return None

def _read_paragraph_properties(
    payload: bytes, position: int, mask: int
) -> tuple[int, _ParagraphProperties]:
    alignment: str | None = None
    bullet: bool | None = None
    bullet_char: str | None = None
    left_margin = indent = line_spacing = space_before = space_after = None
    fixed = ((0xF, 2), (0x80, 2), (0x10, 2), (0x40, 2), (0x20, 4),
             (0x800, 2), (0x1000, 2), (0x2000, 2), (0x4000, 2),
             (0x100, 2), (0x400, 2), (0x8000, 2))
    for property_mask, size in fixed:
        if mask & property_mask:
            if position + size > len(payload):
                return len(payload), _ParagraphProperties(
                    alignment, bullet, bullet_char, left_margin, indent,
                    line_spacing, space_before, space_after
                )
            raw = payload[position:position + size]
            value = int.from_bytes(raw, "little", signed=size == 2)
            if property_mask == 0xF:
                bullet = bool(value & 1)
            elif property_mask == 0x80:
                codepoint = int.from_bytes(raw, "little")
                bullet_char = chr(codepoint) if codepoint else None
            elif property_mask == 0x800:
                alignment = {0: "l", 1: "ctr", 2: "r", 3: "just", 4: "dist"}.get(value)
            elif property_mask == 0x100:
                left_margin = value
            elif property_mask == 0x400:
                indent = value
            elif property_mask == 0x1000:
                line_spacing = value
            elif property_mask == 0x2000:
                space_before = value
            elif property_mask == 0x4000:
                space_after = value
            position += size
    if mask & 0x100000:
        if position + 2 > len(payload):
            return len(payload), _ParagraphProperties(
                alignment, bullet, bullet_char, left_margin, indent,
                line_spacing, space_before, space_after
            )
        count = struct.unpack_from("<H", payload, position)[0]
        position += min(2 + count * 4, len(payload) - position)
    for property_mask in (0x10000, 0xE0000, 0x200000):
        if mask & property_mask:
            position = min(position + 2, len(payload))
    return position, _ParagraphProperties(
        alignment, bullet, bullet_char, left_margin, indent,
        line_spacing, space_before, space_after
    )

def _skip_paragraph_properties(
    payload: bytes, position: int, mask: int
) -> tuple[int, str | None, bool | None, str | None]:
    position, style = _read_paragraph_properties(payload, position, mask)
    return position, style.alignment, style.bullet, style.bullet_char

def _parse_text_ruler(
    payload: bytes,
) -> tuple[tuple[int | None, int | None], ...]:
    return _parse_text_ruler_details(payload)[0]

def _parse_text_ruler_details(
    payload: bytes,
) -> tuple[
    tuple[tuple[int | None, int | None], ...],
    int | None,
    tuple[tuple[int, str], ...],
]:
    """Return margins, default tab size, and explicit tab stops."""
    values: list[list[int | None]] = [[None, None] for _ in range(5)]
    if len(payload) < 4:
        return tuple((left, indent) for left, indent in values), None, ()
    mask = struct.unpack_from("<I", payload)[0]
    position = 4
    default_tab_size = None
    tabs: list[tuple[int, str]] = []

    def skip_short() -> bool:
        nonlocal position
        if position + 2 > len(payload):
            position = len(payload)
            return False
        position += 2
        return True

    # TextRuler serializes cLevels before defaultTabSize, irrespective of bit order.
    if mask & 0x2 and not skip_short():
        return tuple((left, indent) for left, indent in values), None, ()
    if mask & 0x1:
        if position + 2 > len(payload):
            return tuple((left, indent) for left, indent in values), None, ()
        default_tab_size = struct.unpack_from("<h", payload, position)[0]
        position += 2
    if mask & 0x4:
        if position + 2 > len(payload):
            return (
                tuple((left, indent) for left, indent in values),
                default_tab_size,
                (),
            )
        count = struct.unpack_from("<H", payload, position)[0]
        position += 2
        for _ in range(count):
            if position + 4 > len(payload):
                position = len(payload)
                break
            tab_position, tab_alignment = struct.unpack_from(
                "<hH", payload, position
            )
            tabs.append((
                tab_position,
                {0: "l", 1: "ctr", 2: "r", 3: "dec"}.get(tab_alignment, "l"),
            ))
            position += 4
    for level in range(5):
        if mask & (1 << (3 + level)):
            if position + 2 > len(payload):
                break
            values[level][0] = struct.unpack_from("<h", payload, position)[0]
            position += 2
        if mask & (1 << (8 + level)):
            if position + 2 > len(payload):
                break
            values[level][1] = struct.unpack_from("<h", payload, position)[0]
            position += 2
    return (
        tuple((left, indent) for left, indent in values),
        default_tab_size,
        tuple(tabs),
    )

def _text_color(value: int, scheme: tuple[str, ...]) -> str | None:
    high = value >> 24
    if high == 0xFE:
        return f"{value & 0xFF:02X}{(value >> 8) & 0xFF:02X}{(value >> 16) & 0xFF:02X}"
    if value & 0xFFFFFF == 0 and high < len(scheme):
        return scheme[high]
    return _office_color(value, scheme)

def _character_style(payload: bytes, position: int, mask: int, fonts: tuple[str, ...], scheme: tuple[str, ...]) -> tuple[int, TextRun]:
    bold = italic = underline = None
    font_size = None
    color = None
    font_index = None
    baseline = None
    if mask & 0xFFFF:
        if position + 2 > len(payload):
            return len(payload), TextRun("")
        flags = struct.unpack_from("<H", payload, position)[0]; position += 2
        bold = bool(flags & 1) if mask & 1 else None
        italic = bool(flags & 2) if mask & 2 else None
        underline = bool(flags & 4) if mask & 4 else None
    for property_mask in (0x10000, 0x200000, 0x400000, 0x800000):
        if mask & property_mask:
            if position + 2 > len(payload):
                return len(payload), TextRun("")
            index_value = struct.unpack_from("<H", payload, position)[0]
            if property_mask == 0x10000:
                font_index = index_value
            position += 2
    if mask & 0x20000 and position + 2 <= len(payload):
        font_size = struct.unpack_from("<H", payload, position)[0]; position += 2
    if mask & 0x40000 and position + 4 <= len(payload):
        color = _text_color(struct.unpack_from("<I", payload, position)[0], scheme); position += 4
    if mask & 0x80000 and position + 2 <= len(payload):
        baseline = struct.unpack_from("<h", payload, position)[0]
        position += 2
    return position, TextRun("", bold, italic, underline, font_size, color,
                             fonts[font_index] if font_index is not None and font_index < len(fonts) else None,
                             baseline=baseline)

def _merge_style(value: str, explicit: TextRun, master: TextRun | None) -> TextRun:
    if master is None:
        return TextRun(value, explicit.bold, explicit.italic, explicit.underline,
                       explicit.font_size, explicit.color, explicit.typeface,
                       explicit.hyperlink, explicit.baseline)
    return TextRun(value,
                   explicit.bold if explicit.bold is not None else master.bold,
                   explicit.italic if explicit.italic is not None else master.italic,
                   explicit.underline if explicit.underline is not None else master.underline,
                   explicit.font_size or master.font_size,
                   explicit.color or master.color,
                   explicit.typeface or master.typeface,
                   explicit.hyperlink,
                   explicit.baseline if explicit.baseline is not None else master.baseline)

def _master_at_level(masters: tuple[_MasterStyle, ...], level: int) -> _MasterStyle | None:
    if not masters:
        return None
    effective = masters[0]
    for current in masters[1:min(level, len(masters) - 1) + 1]:
        effective = _MasterStyle(
            _merge_style("", current.run, effective.run),
            current.alignment if current.alignment is not None else effective.alignment,
            current.bullet if current.bullet is not None else effective.bullet,
            current.bullet_char or effective.bullet_char,
            current.left_margin if current.left_margin is not None else effective.left_margin,
            current.indent if current.indent is not None else effective.indent,
            current.line_spacing if current.line_spacing is not None else effective.line_spacing,
            current.space_before if current.space_before is not None else effective.space_before,
            current.space_after if current.space_after is not None else effective.space_after,
        )
    return effective

def _style_text(text: str, payload: bytes, fonts: tuple[str, ...],
                masters: tuple[_MasterStyle, ...], scheme: tuple[str, ...],
                text_type: int,
                ruler: tuple[tuple[int | None, int | None], ...] = (),
                default_tab_size: int | None = None,
                tab_stops: tuple[tuple[int, str], ...] = ()) -> TextContent:
    position = handled = 0
    paragraph_runs: list[tuple[int, int, int, _ParagraphProperties]] = []
    while position + 10 <= len(payload) and handled < len(text) + 1:
        count = struct.unpack_from("<I", payload, position)[0]
        level = max(0, min(4, struct.unpack_from("<h", payload, position + 4)[0]))
        mask = struct.unpack_from("<I", payload, position + 6)[0]
        position, paragraph_style = _read_paragraph_properties(
            payload, position + 10, mask
        )
        if not count:
            break
        paragraph_runs.append((handled, handled + count, level, paragraph_style))
        handled += count
    character_runs: list[tuple[int, int, TextRun]] = []
    handled = 0
    while position + 8 <= len(payload) and handled < len(text) + 1:
        count, mask = struct.unpack_from("<II", payload, position)
        position += 8
        if not count:
            break
        position, explicit = _character_style(payload, position, mask, fonts, scheme)
        end = min(handled + count, len(text))
        if end > handled:
            character_runs.append((handled, end, explicit))
        handled += count

    boundaries = {0, len(text)}
    for start, end, *_rest in paragraph_runs:
        boundaries.update((max(0, min(len(text), start)), max(0, min(len(text), end))))
    for start, end, _run in character_runs:
        boundaries.update((start, end))
    ordered = sorted(boundaries)
    styled_runs: list[TextRun] = []
    for start, end in zip(ordered, ordered[1:]):
        if end <= start:
            continue
        paragraph_style = next((run for run in paragraph_runs if run[0] <= start < run[1]), None)
        level = paragraph_style[2] if paragraph_style else 0
        master = _master_at_level(masters, level)
        explicit = next((run for left, right, run in character_runs if left <= start < right), TextRun(""))
        styled_runs.append(_merge_style(text[start:end], explicit, master.run if master else None))

    alignments: list[str | None] = []
    bullets: list[bool] = []
    levels: list[int] = []
    bullet_chars: list[str | None] = []
    left_margins: list[int | None] = []
    indents: list[int | None] = []
    line_spacings: list[int | None] = []
    space_before: list[int | None] = []
    space_after: list[int | None] = []
    paragraph_start = 0
    for paragraph in text.split("\r"):
        style = next((run for run in paragraph_runs if run[0] <= paragraph_start < run[1]), None)
        level = style[2] if style else 0
        paragraph_style = style[3] if style else _ParagraphProperties()
        master = _master_at_level(masters, level)
        ruler_left, ruler_indent = ruler[level] if level < len(ruler) else (None, None)
        alignments.append(
            paragraph_style.alignment
            if paragraph_style.alignment is not None
            else master.alignment if master else None
        )
        bullets.append(
            paragraph_style.bullet
            if paragraph_style.bullet is not None
            else bool(master.bullet) if master else False
        )
        levels.append(level)
        bullet_chars.append(
            paragraph_style.bullet_char
            if paragraph_style.bullet_char is not None
            else master.bullet_char if master else None
        )
        left_margins.append(
            paragraph_style.left_margin
            if paragraph_style.left_margin is not None
            else ruler_left if ruler_left is not None
            else master.left_margin if master else None
        )
        indents.append(
            paragraph_style.indent
            if paragraph_style.indent is not None
            else ruler_indent if ruler_indent is not None
            else master.indent if master else None
        )
        line_spacings.append(
            paragraph_style.line_spacing
            if paragraph_style.line_spacing is not None
            else master.line_spacing if master else None
        )
        space_before.append(
            paragraph_style.space_before
            if paragraph_style.space_before is not None
            else master.space_before if master else None
        )
        space_after.append(
            paragraph_style.space_after
            if paragraph_style.space_after is not None
            else master.space_after if master else None
        )
        paragraph_start += len(paragraph) + 1
    return TextContent(
        text=text,
        runs=tuple(styled_runs),
        paragraph_alignments=tuple(alignments),
        paragraph_bullets=tuple(bullets),
        paragraph_levels=tuple(levels),
        paragraph_bullet_chars=tuple(bullet_chars),
        paragraph_left_margins=tuple(left_margins),
        paragraph_indents=tuple(indents),
        paragraph_line_spacings=tuple(line_spacings),
        paragraph_space_before=tuple(space_before),
        paragraph_space_after=tuple(space_after),
        text_type=text_type,
        default_tab_size=default_tab_size,
        tab_stops=tab_stops,
    )

def _apply_hyperlinks(content: TextContent, spans: list[tuple[int, int, str]]) -> TextContent:
    if not spans:
        return content
    output: list[TextRun] = []
    offset = 0
    for run in content.runs:
        boundaries = {offset, offset + len(run.text)}
        for start, end, _url in spans:
            if offset < start < offset + len(run.text): boundaries.add(start)
            if offset < end < offset + len(run.text): boundaries.add(end)
        ordered = sorted(boundaries)
        for left, right in zip(ordered, ordered[1:]):
            value = run.text[left - offset:right - offset]
            url = next((target for start, end, target in spans if start <= left < end), None)
            output.append(TextRun(value, run.bold, run.italic, True if url else run.underline,
                                  run.font_size, run.color, run.typeface, url, run.baseline))
        offset += len(run.text)
    return TextContent(
        text=content.text,
        runs=tuple(output),
        paragraph_alignments=content.paragraph_alignments,
        paragraph_bullets=content.paragraph_bullets,
        paragraph_levels=content.paragraph_levels,
        paragraph_bullet_chars=content.paragraph_bullet_chars,
        paragraph_left_margins=content.paragraph_left_margins,
        paragraph_indents=content.paragraph_indents,
        paragraph_line_spacings=content.paragraph_line_spacings,
        paragraph_space_before=content.paragraph_space_before,
        paragraph_space_after=content.paragraph_space_after,
        text_type=content.text_type,
        default_tab_size=content.default_tab_size,
        tab_stops=content.tab_stops,
    )

def _text_contents(source_records: list[Record], fonts: tuple[str, ...], masters: dict[int, tuple[_MasterStyle, ...]], hyperlinks: dict[int, str], scheme: tuple[str, ...]) -> list[TextContent]:
    result: list[TextContent] = []
    text_type = 4
    for index, record in enumerate(source_records):
        if record.type == 3999 and len(record.payload) >= 4:
            text_type = struct.unpack_from("<I", record.payload)[0]
            continue
        value = _text(record)
        if value is None:
            continue
        tail = source_records[index + 1:]
        next_text = next((position for position, candidate in enumerate(tail) if candidate.type in (RT_TEXT_CHARS_ATOM, RT_TEXT_BYTES_ATOM, RT_SLIDE_PERSIST_ATOM)), len(tail))
        related = tail[:next_text]
        style = next((candidate for candidate in related if candidate.type == 4001), None)
        ruler_atom = next((candidate for candidate in related if candidate.type == 4006), None)
        ruler, default_tab_size, tab_stops = (
            _parse_text_ruler_details(ruler_atom.payload)
            if ruler_atom else ((), None, ())
        )
        master = masters.get(text_type, ())
        base = master[0] if master else None
        if style is not None and style.type == 4001:
            content = _style_text(
                value, style.payload, fonts, master, scheme, text_type, ruler,
                default_tab_size, tab_stops,
            )
        else:
            paragraphs = value.split("\r")
            ruler_left, ruler_indent = ruler[0] if ruler else (None, None)
            content = TextContent(
                text=value,
                runs=(_merge_style(value, TextRun(""), base.run if base else None),),
                paragraph_alignments=tuple(
                    base.alignment if base else None for _ in paragraphs
                ),
                paragraph_bullets=tuple(
                    bool(base.bullet) if base else False for _ in paragraphs
                ),
                paragraph_levels=tuple(0 for _ in paragraphs),
                paragraph_bullet_chars=tuple(
                    base.bullet_char if base else None for _ in paragraphs
                ),
                paragraph_left_margins=tuple(
                    ruler_left if ruler_left is not None
                    else base.left_margin if base else None
                    for _ in paragraphs
                ),
                paragraph_indents=tuple(
                    ruler_indent if ruler_indent is not None
                    else base.indent if base else None
                    for _ in paragraphs
                ),
                paragraph_line_spacings=tuple(
                    base.line_spacing if base else None for _ in paragraphs
                ),
                paragraph_space_before=tuple(
                    base.space_before if base else None for _ in paragraphs
                ),
                paragraph_space_after=tuple(
                    base.space_after if base else None for _ in paragraphs
                ),
                text_type=text_type,
                default_tab_size=default_tab_size,
                tab_stops=tab_stops,
            )
        spans: list[tuple[int, int, str]] = []
        pending_id: int | None = None
        for candidate in related:
            if candidate.type == 4082 and candidate.version == CONTAINER_VERSION:
                atom = next((child for child in records(candidate.payload) if child.type == 4083 and len(child.payload) >= 8), None)
                pending_id = struct.unpack_from("<I", atom.payload, 4)[0] if atom else None
            elif candidate.type == 4063 and len(candidate.payload) >= 8 and pending_id in hyperlinks:
                start, end = struct.unpack_from("<II", candidate.payload)
                spans.append((start, end, hyperlinks[pending_id]))
                pending_id = None
        result.append(_apply_hyperlinks(content, spans))
    return result

def _fonts(document: Record) -> tuple[str, ...]:
    result: list[str] = []
    for atom in descendants(document):
        if atom.type != 4023 or len(atom.payload) < 2:
            continue
        name = atom.payload[:64].decode("utf-16le", "replace").split("\0", 1)[0]
        result.append(name or "Arial")
    return tuple(result)

def _master_text_styles(powerpoint_document: bytes, document: Record, fonts: tuple[str, ...], scheme: tuple[str, ...]) -> dict[int, tuple[_MasterStyle, ...]]:
    result: dict[int, tuple[_MasterStyle, ...]] = {}
    atoms = list(descendants(document))
    for root in records(powerpoint_document):
        if root.type == 1016:
            atoms.extend(descendants(root))
    for atom in atoms:
        if atom.type != 4003 or len(atom.payload) < 2 or atom.instance in result:
            continue
        levels = struct.unpack_from("<H", atom.payload)[0]
        position = 2
        styles: list[_MasterStyle] = []
        for _level in range(min(levels, 5)):
            if atom.instance >= 5:
                position += 2
            if position + 4 > len(atom.payload):
                break
            paragraph_mask = struct.unpack_from("<I", atom.payload, position)[0]
            position, paragraph_style = _read_paragraph_properties(
                atom.payload, position + 4, paragraph_mask
            )
            if position + 4 > len(atom.payload):
                break
            character_mask = struct.unpack_from("<I", atom.payload, position)[0]
            position, run = _character_style(atom.payload, position + 4, character_mask, fonts, scheme)
            styles.append(_MasterStyle(
                run,
                paragraph_style.alignment,
                paragraph_style.bullet,
                paragraph_style.bullet_char,
                paragraph_style.left_margin,
                paragraph_style.indent,
                paragraph_style.line_spacing,
                paragraph_style.space_before,
                paragraph_style.space_after,
            ))
        if styles:
            result[atom.instance] = tuple(styles)
    for child_type, parent_type in ((5, 1), (6, 0), (7, 1), (8, 1)):
        child, parent = result.get(child_type), result.get(parent_type)
        if not child or not parent:
            continue
        merged: list[_MasterStyle] = []
        for index, child_style in enumerate(child):
            parent_style = parent[min(index, len(parent) - 1)]
            merged.append(_MasterStyle(
                _merge_style("", child_style.run, parent_style.run),
                child_style.alignment if child_style.alignment is not None else parent_style.alignment,
                child_style.bullet if child_style.bullet is not None else parent_style.bullet,
                child_style.bullet_char or parent_style.bullet_char,
                child_style.left_margin if child_style.left_margin is not None else parent_style.left_margin,
                child_style.indent if child_style.indent is not None else parent_style.indent,
                child_style.line_spacing if child_style.line_spacing is not None else parent_style.line_spacing,
                child_style.space_before if child_style.space_before is not None else parent_style.space_before,
                child_style.space_after if child_style.space_after is not None else parent_style.space_after,
            ))
        result[child_type] = tuple(merged)
    return result

def _hyperlinks(document: Record) -> dict[int, str]:
    result: dict[int, str] = {}
    for container in descendants(document):
        if container.type != 4055 or container.version != CONTAINER_VERSION:
            continue
        link_id = None
        strings: list[str] = []
        for child in records(container.payload):
            if child.type == 4051 and len(child.payload) >= 4:
                link_id = struct.unpack_from("<I", child.payload)[0]
            elif child.type == 4026:
                strings.append(child.payload.decode("utf-16le", "replace").rstrip("\0"))
        if link_id is not None and strings:
            target = strings[-1]
            if target.startswith(("http://", "https://", "mailto:")):
                result[link_id] = target
    return result

def _direct_children(record: Record) -> list[Record]:
    return list(records(record.payload)) if record.version == CONTAINER_VERSION else []

@dataclass(frozen=True, slots=True)
class _GroupSpace:
    coord_left: int
    coord_top: int
    coord_right: int
    coord_bottom: int
    abs_left: int
    abs_top: int
    abs_right: int
    abs_bottom: int
    a: float | None = None
    b: float = 0.0
    c: float = 0.0
    d: float | None = None
    tx: float = 0.0
    ty: float = 0.0
    flip_horizontal: bool = False
    flip_vertical: bool = False

def _client_rect(payload: bytes) -> tuple[int, int, int, int] | None:
    if len(payload) < 8:
        return None
    top, left, right, bottom = struct.unpack_from("<4h", payload)
    return left, top, right, bottom

def _child_rect(payload: bytes) -> tuple[int, int, int, int] | None:
    if len(payload) < 16:
        return None
    return struct.unpack_from("<4i", payload)

def _space_matrix(space: _GroupSpace) -> tuple[float, float, float, float, float, float]:
    if space.a is not None and space.d is not None:
        return space.a, space.b, space.c, space.d, space.tx, space.ty
    coord_width = max(1, space.coord_right - space.coord_left)
    coord_height = max(1, space.coord_bottom - space.coord_top)
    scale_x = (space.abs_right - space.abs_left) / coord_width
    scale_y = (space.abs_bottom - space.abs_top) / coord_height
    return (scale_x, 0.0, 0.0, scale_y,
            space.abs_left - scale_x * space.coord_left,
            space.abs_top - scale_y * space.coord_top)

def _compose_matrix(parent: tuple[float, float, float, float, float, float],
                    child: tuple[float, float, float, float, float, float]
                    ) -> tuple[float, float, float, float, float, float]:
    pa, pb, pc, pd, ptx, pty = parent
    ca, cb, cc, cd, ctx, cty = child
    return (pa * ca + pc * cb, pb * ca + pd * cb,
            pa * cc + pc * cd, pb * cc + pd * cd,
            pa * ctx + pc * cty + ptx, pb * ctx + pd * cty + pty)

def _transform_box(rect: tuple[int, int, int, int],
                   matrix: tuple[float, float, float, float, float, float]
                   ) -> tuple[int, int, int, int]:
    left, top, right, bottom = rect
    a, b, c, d, tx, ty = matrix
    center_x, center_y = (left + right) / 2, (top + bottom) / 2
    mapped_x = a * center_x + c * center_y + tx
    mapped_y = b * center_x + d * center_y + ty
    width = max(1, round(math.hypot(a, b) * abs(right - left)))
    height = max(1, round(math.hypot(c, d) * abs(bottom - top)))
    return (round(mapped_x - width / 2), round(mapped_y - height / 2), width, height)

def _transform_box_in_space(rect: tuple[int, int, int, int],
                            space: _GroupSpace) -> tuple[int, int, int, int]:
    matrix = _space_matrix(space)
    a, b, c, d, tx, ty = matrix
    scale_x, scale_y = math.hypot(a, b), math.hypot(c, d)
    smaller = min(scale_x, scale_y)
    # A rotated nested group can encode a very anisotropic intermediate anchor
    # even though its children retain their aspect ratio.  PowerPoint resolves
    # that case with the area-preserving (determinant) scale around the group
    # center; applying the two raw axis scales produces tall, misplaced slivers.
    if smaller > 0 and max(scale_x, scale_y) / smaller > 4 and (abs(b) > 1e-12 or abs(c) > 1e-12):
        left, top, right, bottom = rect
        coord_center_x = (space.coord_left + space.coord_right) / 2
        coord_center_y = (space.coord_top + space.coord_bottom) / 2
        abs_center_x = a * coord_center_x + c * coord_center_y + tx
        abs_center_y = b * coord_center_x + d * coord_center_y + ty
        rect_center_x, rect_center_y = (left + right) / 2, (top + bottom) / 2
        scale = math.sqrt(abs(a * d - b * c))
        angle = math.atan2(b, a)
        delta_x = (rect_center_x - coord_center_x) * scale
        delta_y = (rect_center_y - coord_center_y) * scale
        mapped_x = abs_center_x + math.cos(angle) * delta_x - math.sin(angle) * delta_y
        mapped_y = abs_center_y + math.sin(angle) * delta_x + math.cos(angle) * delta_y
        width = max(1, round(abs(right - left) * scale))
        height = max(1, round(abs(bottom - top) * scale))
        return (round(mapped_x - width / 2), round(mapped_y - height / 2), width, height)
    return _transform_box(rect, matrix)

def _rect_to_box(left: int, top: int, right: int, bottom: int) -> tuple[int, int, int, int]:
    return left, top, max(1, right - left), max(1, bottom - top)

def _anchor(children: list[Record], fallback_index: int, space: _GroupSpace | None = None) -> tuple[int, int, int, int]:
    client = next((child for child in children if child.type == RT_OFFICEART_CLIENT_ANCHOR), None)
    if client is not None:
        rect = _client_rect(client.payload)
        if rect is not None:
            return _rect_to_box(*rect)
    child_anchor = next((child for child in children if child.type == RT_OFFICEART_CHILD_ANCHOR), None)
    if child_anchor is not None:
        rect = _child_rect(child_anchor.payload)
        if rect is not None:
            left, top, right, bottom = rect
            if space is not None:
                return _transform_box_in_space(rect, space)
            if max(abs(left), abs(top), abs(right), abs(bottom)) > 100_000:
                left, top, right, bottom = (round(value * 576 / 914400)
                                            for value in (left, top, right, bottom))
            return _rect_to_box(left, top, right, bottom)
    return 288, 288 + fallback_index * 576, 5184, 432

def _group_space(group_shape: Record, parent: _GroupSpace | None) -> _GroupSpace | None:
    children = _direct_children(group_shape)
    fspgr = next((child for child in children if child.type == RT_OFFICEART_FSPGR and len(child.payload) >= 16), None)
    if fspgr is None:
        return parent
    coord_left, coord_top, coord_right, coord_bottom = struct.unpack_from("<4i", fspgr.payload)
    if coord_right <= coord_left or coord_bottom <= coord_top:
        return parent
    if max(abs(coord_left), abs(coord_top), abs(coord_right), abs(coord_bottom)) > 10_000_000:
        return parent
    client = next((child for child in children if child.type == RT_OFFICEART_CLIENT_ANCHOR), None)
    child_anchor = next((child for child in children if child.type == RT_OFFICEART_CHILD_ANCHOR), None)
    rect = _client_rect(client.payload) if client is not None else (
        _child_rect(child_anchor.payload) if child_anchor is not None else None
    )
    if rect is None:
        return parent
    left, top, right, bottom = rect
    if parent is None and child_anchor is not None and max(map(abs, rect)) > 100_000:
        left, top, right, bottom = (round(value * 576 / 914400) for value in rect)
    coord_width = coord_right - coord_left
    coord_height = coord_bottom - coord_top
    scale_x = (right - left) / coord_width
    scale_y = (bottom - top) / coord_height
    fopt = next((child for child in children if child.type == RT_OFFICEART_FOPT), None)
    properties = _fopt_properties(fopt) if fopt else {}
    rotation, flip_horizontal, flip_vertical = _transform(children, properties)
    if flip_horizontal:
        base_a, base_tx = -scale_x, right + scale_x * coord_left
    else:
        base_a, base_tx = scale_x, left - scale_x * coord_left
    if flip_vertical:
        base_d, base_ty = -scale_y, bottom + scale_y * coord_top
    else:
        base_d, base_ty = scale_y, top - scale_y * coord_top
    angle = math.radians(rotation / 60000)
    cosine, sine = math.cos(angle), math.sin(angle)
    center_x, center_y = (left + right) / 2, (top + bottom) / 2
    own = (
        cosine * base_a, sine * base_a,
        -sine * base_d, cosine * base_d,
        cosine * (base_tx - center_x) - sine * (base_ty - center_y) + center_x,
        sine * (base_tx - center_x) + cosine * (base_ty - center_y) + center_y,
    )
    matrix = _compose_matrix(_space_matrix(parent), own) if parent is not None else own
    abs_left, abs_top, width, height = _transform_box(
        (coord_left, coord_top, coord_right, coord_bottom), matrix
    )
    return _GroupSpace(coord_left, coord_top, coord_right, coord_bottom,
                       abs_left, abs_top, abs_left + width, abs_top + height,
                       *matrix,
                       (parent.flip_horizontal if parent else False) ^ flip_horizontal,
                       (parent.flip_vertical if parent else False) ^ flip_vertical)

def _combine_transform(transform: tuple[int, bool, bool],
                       space: _GroupSpace | None) -> tuple[int, bool, bool]:
    if space is None:
        return transform
    rotation, flip_horizontal, flip_vertical = transform
    a, b, _c, _d, _tx, _ty = _space_matrix(space)
    group_rotation = round(math.degrees(math.atan2(b, a)) * 60000)
    return (rotation + group_rotation,
            flip_horizontal ^ space.flip_horizontal,
            flip_vertical ^ space.flip_vertical)

def _iter_sp_containers(record: Record, space: _GroupSpace | None = None):
    """Yield (OfficeArtSpContainer, group space) with nested group transforms applied."""
    if record.type == RT_OFFICEART_SPGR_CONTAINER and record.version == CONTAINER_VERSION:
        children = list(records(record.payload))
        if not children:
            return
        nested = _group_space(children[0], space)
        for child in children[1:]:
            yield from _iter_sp_containers(child, nested)
        return
    if record.type == RT_OFFICEART_SP_CONTAINER:
        yield record, space
        return
    if record.version == CONTAINER_VERSION and not (RT_ROUNDTRIP_OPAQUE_MIN <= record.type <= RT_ROUNDTRIP_OPAQUE_MAX):
        for child in records(record.payload):
            yield from _iter_sp_containers(child, space)

def _is_placeholder(shape: Record) -> bool:
    for child in _direct_children(shape):
        if child.type != RT_OFFICEART_CLIENT_DATA:
            continue
        try:
            for atom in records(child.payload) if child.version == CONTAINER_VERSION else ():
                if atom.type == RT_OE_PLACEHOLDER_ATOM:
                    return True
            for atom in descendants(child):
                if atom.type == RT_OE_PLACEHOLDER_ATOM:
                    return True
        except InvalidPpt:
            continue
    return False

def _is_background_shape(children: list[Record]) -> bool:
    sp = next((child for child in children if child.type == 0xF00A and len(child.payload) >= 8), None)
    if sp is None:
        return False
    return bool(struct.unpack_from("<I", sp.payload, 4)[0] & 0x400)

def _character_width(character: str, font_size: int) -> float:
    if character == " ":
        factor = 0.28
    elif character in "ilIjtfr.,:;'|!":
        factor = 0.28
    elif character in "MW@%&":
        factor = 0.85
    elif ord(character) > 0x2FF:
        factor = 1.0
    elif character.isupper():
        factor = 0.65
    else:
        factor = 0.5
    return font_size * 8 * factor

def _minimum_unwrapped_width(content: TextContent) -> int:
    """Estimate the width PowerPoint gives an auto-sized, non-wrapping text box."""
    line_width = maximum = 0.0
    for run in content.runs or (TextRun(content.text),):
        font_size = run.font_size or 18
        for character in run.text:
            if character in "\r\n":
                maximum = max(maximum, line_width)
                line_width = 0.0
                continue
            line_width += _character_width(character, font_size)
    maximum = max(maximum, line_width)
    # Default DrawingML left/right text insets are 0.1in each.
    # A small safety allowance prevents LibreOffice from wrapping at a
    # borderline glyph advance that PowerPoint kept on one line.
    return round((maximum + 2 * 57.6) * 1.15)

def _minimum_wrapped_height(
    content: TextContent,
    width: int,
    inset_left: int | None,
    inset_top: int | None,
    inset_right: int | None,
    inset_bottom: int | None,
) -> int:
    # Text insets in OfficeArt FOPT are already EMUs, while anchors and the
    # estimates in this module use the legacy 576-units-per-inch coordinate
    # system.  Mixing those units can turn a normal text box into a shape
    # hundreds of slides tall.
    def inset_units(value: int | None, default: float) -> float:
        return default if value is None else value * 576 / 914400

    available = max(
        1,
        width
        - inset_units(inset_left, 57.6)
        - inset_units(inset_right, 57.6),
    )
    paragraphs: list[tuple[float, int]] = [(0.0, 18)]
    for run in content.runs or (TextRun(content.text),):
        font_size = run.font_size or 18
        for character in run.text:
            if character in "\r\n":
                paragraphs.append((0.0, font_size))
                continue
            line_width, maximum_size = paragraphs[-1]
            paragraphs[-1] = (
                line_width + _character_width(character, font_size),
                max(maximum_size, font_size),
            )
    text_height = sum(
        max(1, math.ceil(line_width / available)) * font_size * 8 * 1.15
        for line_width, font_size in paragraphs
    )
    vertical_insets = (
        inset_units(inset_top, 29.5)
        + inset_units(inset_bottom, 29.5)
    )
    return round(text_height + vertical_insets)

def _shape_text_boxes(slide: Record, external_text: list[TextContent], fonts: tuple[str, ...], masters: dict[int, tuple[_MasterStyle, ...]], hyperlinks: dict[int, str], scheme: tuple[str, ...], *, skip_placeholders: bool = False, exclude: set[int] | None = None) -> list[TextBox]:
    result: list[TextBox] = []
    for shape, space in _iter_sp_containers(slide):
        if exclude is not None and shape.offset in exclude:
            continue
        is_placeholder = _is_placeholder(shape)
        if skip_placeholders and is_placeholder:
            continue
        children = _direct_children(shape)
        textbox = next((child for child in children if child.type == RT_OFFICEART_CLIENT_TEXTBOX), None)
        if textbox is None:
            continue
        contents = _text_contents(list(descendants(textbox)), fonts, masters, hyperlinks, scheme)
        content = contents[0] if contents else None
        if content is None:
            reference = next((child for child in descendants(textbox) if child.type == RT_OUTLINE_TEXT_REF_ATOM and len(child.payload) >= 4), None)
            if reference is not None:
                index = struct.unpack_from("<I", reference.payload)[0]
                if index < len(external_text):
                    content = external_text[index]
        if content is None:
            continue
        left, top, width, height = _anchor(children, len(result), space)
        fopt = next((child for child in children if child.type == RT_OFFICEART_FOPT), None)
        properties = _fopt_properties(fopt) if fopt else {}
        fill, line, dash, fill_pattern, fill_back = _shape_style(properties, scheme)
        transform = _combine_transform(_transform(children, properties), space)
        sp = next((child for child in children if child.type == 0xF00A), None)
        preset = SHAPE_PRESETS.get(sp.instance, "rect") if sp is not None else "rect"
        text_anchor = properties.get(135, 0)
        vertical_anchor = (
            "ctr" if text_anchor in (1, 4)
            else "b" if text_anchor in (2, 5)
            else None
        )
        if 135 not in properties and is_placeholder and content.text_type in (0, 6):
            vertical_anchor = "ctr"
        text_flags = properties.get(191, 0)
        auto_fit = (
            (is_placeholder and content.text_type in (1, 5, 7, 8))
            or bool(text_flags & 0x40000 and text_flags & 0x4)
        )
        fit_shape_to_text = bool(text_flags & 0x20000 and text_flags & 0x2)
        wrap_text = properties.get(133, 0) != 2
        if fit_shape_to_text and not wrap_text and fill is None and line is None:
            fitted_width = _minimum_unwrapped_width(content)
            if fitted_width > width:
                growth = fitted_width - width
                alignments = tuple(
                    alignment for alignment in content.paragraph_alignments
                    if alignment is not None
                )
                if alignments and all(alignment == "ctr" for alignment in alignments):
                    left -= round(growth / 2)
                elif alignments and all(alignment == "r" for alignment in alignments):
                    left -= growth
                width = fitted_width
        elif fit_shape_to_text and wrap_text:
            fitted_height = _minimum_wrapped_height(
                content,
                width,
                properties.get(129) if 129 in properties else None,
                properties.get(130) if 130 in properties else None,
                properties.get(131) if 131 in properties else None,
                properties.get(132) if 132 in properties else None,
            )
            if fitted_height > height:
                growth = fitted_height - height
                if vertical_anchor == "ctr":
                    top -= round(growth / 2)
                elif vertical_anchor == "b":
                    top -= growth
                height = fitted_height
        result.append(TextBox(
            text=content.text,
            left=left,
            top=top,
            width=width,
            height=height,
            runs=content.runs,
            paragraph_alignments=content.paragraph_alignments,
            paragraph_bullets=content.paragraph_bullets,
            rotation=transform[0],
            flip_horizontal=transform[1],
            flip_vertical=transform[2],
            fill_color=fill,
            line_color=line,
            line_dash=dash,
            paragraph_levels=content.paragraph_levels,
            paragraph_bullet_chars=content.paragraph_bullet_chars,
            auto_fit=auto_fit,
            fit_shape_to_text=fit_shape_to_text,
            vertical_anchor=vertical_anchor,
            preset=preset,
            wrap_text=wrap_text,
            paragraph_left_margins=content.paragraph_left_margins,
            paragraph_indents=content.paragraph_indents,
            paragraph_line_spacings=content.paragraph_line_spacings,
            paragraph_space_before=content.paragraph_space_before,
            paragraph_space_after=content.paragraph_space_after,
            inset_left=properties.get(129) if 129 in properties else None,
            inset_top=properties.get(130) if 130 in properties else None,
            inset_right=properties.get(131) if 131 in properties else None,
            inset_bottom=properties.get(132) if 132 in properties else None,
            is_placeholder=is_placeholder,
            line_width=_line_width(properties, space),
            # Legacy text boxes commonly use a leading tab as a compact
            # alignment aid.  DrawingML's larger implicit tab interval makes
            # those lines wrap, so preserve the legacy half-inch fallback.
            default_tab_size=(
                content.default_tab_size
                if content.default_tab_size is not None
                else 288 if "\t" in content.text and not content.tab_stops
                else None
            ),
            tab_stops=content.tab_stops,
            fill_pattern=fill_pattern,
            fill_back_color=fill_back,
            line_head=_line_end(properties, 464, 466, 467),
            line_tail=_line_end(properties, 465, 468, 469),
        ))
    return result

def _fopt_properties(record: Record) -> dict[int, int]:
    count = min(record.instance, len(record.payload) // 6)
    properties: dict[int, int] = {}
    for index in range(count):
        opid, value = struct.unpack_from("<HI", record.payload, index * 6)
        properties[opid & 0x3FFF] = value
    return properties

def _fopt_complex_properties(record: Record) -> dict[int, bytes]:
    count = min(record.instance, len(record.payload) // 6)
    ordered: list[tuple[int, int]] = []
    for index in range(count):
        opid, value = struct.unpack_from("<HI", record.payload, index * 6)
        if opid & 0x8000:
            ordered.append((opid & 0x3FFF, value))
    cursor = count * 6
    result: dict[int, bytes] = {}
    for pid, size in ordered:
        actual_size = size
        # OfficeArt pVertices uses a packed-array header whose six bytes are
        # not included in the FOPTE size.  Without accounting for it, the
        # following pSegmentInfo starts six bytes early, drops the final
        # vertices, and can accidentally close an otherwise open path.
        if pid == 325 and cursor + 6 <= len(record.payload):
            item_count, _allocated, item_size = struct.unpack_from(
                "<3H", record.payload, cursor
            )
            element_size = 4 if item_size in (0, 0xFFF0) else item_size
            packed_size = item_count * element_size
            if size == packed_size and cursor + size + 6 <= len(record.payload):
                actual_size += 6
        result[pid] = record.payload[cursor:cursor + actual_size]
        cursor += actual_size
    return result

def _has_fill(properties: dict[int, int]) -> bool:
    flags = properties.get(447)
    if flags is None:
        return 385 in properties
    return bool(flags & 0x10)

def _has_line(properties: dict[int, int]) -> bool:
    flags = properties.get(511)
    if flags is None:
        return 448 in properties
    return bool(flags & 0x8)

def _line_dash(properties: dict[int, int]) -> str | None:
    # Exact MSOLINEDASHING -> DrawingML preset mapping.
    return {
        1: "sysDash",
        2: "sysDot",
        3: "sysDashDot",
        4: "sysDashDotDot",
        5: "dot",
        6: "dash",
        7: "lgDash",
        8: "dashDot",
        9: "lgDashDot",
        10: "lgDashDotDot",
    }.get(properties.get(462))

def _line_end(
    properties: dict[int, int],
    type_property: int,
    width_property: int,
    length_property: int,
) -> tuple[str, str | None, str | None] | None:
    kind = {
        1: "triangle",
        2: "stealth",
        3: "diamond",
        4: "oval",
        5: "arrow",
    }.get(properties.get(type_property))
    if kind is None:
        return None
    sizes = {0: "sm", 1: "med", 2: "lg"}
    return (
        kind,
        sizes.get(properties.get(width_property)),
        sizes.get(properties.get(length_property)),
    )

def _line_width(
    properties: dict[int, int], space: _GroupSpace | None
) -> int | None:
    width = properties.get(459)
    if width is None:
        return None
    if space is not None:
        a, b, c, d, _tx, _ty = _space_matrix(space)
        width = round(width * math.sqrt(abs(a * d - b * c)))
    return max(0, width)

def _parse_imso_points(data: bytes) -> list[tuple[int, int]]:
    if len(data) < 6:
        return []
    count, _alloc, cb_elem = struct.unpack_from("<3H", data)
    if cb_elem == 0xFFF0:
        size = 4
    elif cb_elem == 0xFFF8:
        size = 8
    else:
        size = cb_elem or 4
    points: list[tuple[int, int]] = []
    cursor = 6
    for _ in range(count):
        if cursor + size > len(data):
            break
        if size >= 8:
            x, y = struct.unpack_from("<2i", data, cursor)
        else:
            x, y = struct.unpack_from("<2h", data, cursor)
        points.append((x, y))
        cursor += size
    return points

def _parse_imso_segments(data: bytes) -> list[int]:
    if len(data) < 6:
        return []
    count, _alloc, cb_elem = struct.unpack_from("<3H", data)
    size = 2 if cb_elem in (0, 0xFFF0) else (cb_elem or 2)
    segments: list[int] = []
    cursor = 6
    limit = count if count else (len(data) - 6) // size
    for _ in range(limit):
        if cursor + 2 > len(data):
            break
        segments.append(struct.unpack_from("<H", data, cursor)[0])
        cursor += size
    return segments

def _freeform_path(properties: dict[int, int], complex_props: dict[int, bytes]) -> tuple[tuple[tuple[object, ...], ...], int, int] | None:
    points = _parse_imso_points(complex_props.get(325, b""))
    if len(points) < 2:
        return None
    segments = _parse_imso_segments(complex_props.get(326, b""))
    commands: list[tuple[object, ...]] = []
    index = 0
    for segment in segments:
        command = segment >> 8
        if command == 0x40:  # moveTo
            if index < len(points):
                commands.append(("M", points[index])); index += 1
        elif command == 0x00:  # lineTo
            if index < len(points):
                commands.append(("L", points[index])); index += 1
        elif command == 0x20:  # curveTo
            if index + 2 < len(points):
                commands.append(("C", points[index], points[index + 1], points[index + 2]))
                index += 3
        elif command == 0x60:  # close
            commands.append(("Z",))
        elif command == 0x80:  # end
            break
        # escapes and unknowns are skipped
    if len(commands) < 2:
        commands = [("M", points[0])]
        for point in points[1:]:
            commands.append(("L", point))
        if len(points) >= 3:
            commands.append(("Z",))
    width = max(properties.get(322, 0), max((x for x, _y in points), default=0), 1)
    height = max(properties.get(323, 0), max((y for _x, y in points), default=0), 1)
    return tuple(commands), width, height

def _uses_custom_geometry(
    shape_type: int, complex_props: dict[int, bytes]
) -> bool:
    # Preset shapes such as legacy arcs may carry pVertices for adjustment
    # bookkeeping without pSegmentInfo.  Treating those vertices as an
    # independent polygon turns smooth curves into jagged diamonds.
    return shape_type == 0 or (325 in complex_props and 326 in complex_props)

def _legacy_arc_path(
    complex_props: dict[int, bytes],
) -> tuple[tuple[tuple[object, ...], ...], int, int] | None:
    """Convert the quarter-ellipse geometry carried by a legacy arc."""
    points = _parse_imso_points(complex_props.get(325, b""))
    if len(points) < 4:
        return None
    # The third and fourth guide points are the visible endpoints.  Their
    # vertical order distinguishes the two quarter-arc orientations used to
    # compose smooth sine waves.  11931 is the cubic Bézier circle constant
    # (0.55228475 * 21600) expressed as the control-point distance.
    if points[2][1] > points[3][1]:
        path = (
            ("M", (0, 21600)),
            ("C", (0, 9669), (9669, 0), (21600, 0)),
        )
    else:
        path = (
            ("M", (0, 0)),
            ("C", (11931, 0), (21600, 9669), (21600, 21600)),
        )
    return path, 21600, 21600

def _shape_style(
    properties: dict[int, int], scheme: tuple[str, ...]
) -> tuple[str | None, str | None, str | None, str | None, str | None]:
    fill = _office_color(properties.get(385), scheme) if _has_fill(properties) else None
    back = _office_color(properties.get(387), scheme)
    fill_type = properties.get(384, 0)
    pattern = "dkUpDiag" if fill is not None and back is not None and fill_type == 1 else None
    if fill is not None and fill_type >= 4 and back is not None:
        fill = back
        back = None
    line = _office_color(properties.get(448), scheme) if _has_line(properties) else None
    return fill, line, _line_dash(properties) if line else None, pattern, back if pattern else None

def _transform(children: list[Record], properties: dict[int, int]) -> tuple[int, bool, bool]:
    rotation_raw = properties.get(4, 0)
    rotation_signed = struct.unpack("<i", struct.pack("<I", rotation_raw))[0]
    rotation = round(rotation_signed / 65536 * 60000)
    sp = next((child for child in children if child.type == 0xF00A and len(child.payload) >= 8), None)
    flags = struct.unpack_from("<I", sp.payload, 4)[0] if sp else 0
    return rotation, bool(flags & 0x40), bool(flags & 0x80)

def _gradient_angle(properties: dict[int, int]) -> int:
    """Translate OfficeArt's counter-clockwise, bottom-up angle to DrawingML."""
    raw = properties.get(395, 0)
    signed = struct.unpack("<i", struct.pack("<I", raw))[0]
    degrees = signed / 65536
    return round(((90 - degrees) % 360) * 60000)

def _shape_adjustments(shape_type: int, properties: dict[int, int]) -> tuple[int, ...]:
    """Normalize legacy connector bend positions to DrawingML guide values."""
    if shape_type != 34 or 327 not in properties:
        return ()
    signed = struct.unpack("<i", struct.pack("<I", properties[327]))[0]
    return (round(signed * 100000 / 21600),)

def _connector_transform(
    shape_type: int, transform: tuple[int, bool, bool]
) -> tuple[int, bool, bool]:
    """Preserve endpoint direction when normalizing elbow connectors."""
    rotation, flip_horizontal, flip_vertical = transform
    if shape_type != 34 or flip_vertical:
        return transform
    rotation = (rotation + 10800000) % 21600000
    if not flip_horizontal:
        flip_horizontal = True
        flip_vertical = True
    return rotation, flip_horizontal, flip_vertical

def _office_color(value: int | None, scheme: tuple[str, ...] = ()) -> str | None:
    if value is None:
        return None
    if value & 0x08000000:
        index = value & 0xFF
        return scheme[index] if index < len(scheme) else None
    if value & 0x17000000:
        return None
    red, green, blue = value & 0xFF, (value >> 8) & 0xFF, (value >> 16) & 0xFF
    return f"{red:02X}{green:02X}{blue:02X}"

def _basic_shapes(slide: Record, image_map: dict[int, tuple[bytes, str, str]], scheme: tuple[str, ...], exclude: set[int] | None = None) -> list[BasicShape]:
    result: list[BasicShape] = []
    for shape, space in _iter_sp_containers(slide):
        if exclude is not None and shape.offset in exclude:
            continue
        children = _direct_children(shape)
        if _is_background_shape(children) or _is_placeholder(shape):
            continue
        # Text-bearing shapes are emitted as TextBox with their own stroke/fill.
        textbox = next((child for child in children if child.type == RT_OFFICEART_CLIENT_TEXTBOX), None)
        if textbox is not None:
            has_visible_text = any(
                child.type in (RT_TEXT_CHARS_ATOM, RT_TEXT_BYTES_ATOM, RT_OUTLINE_TEXT_REF_ATOM)
                for child in descendants(textbox)
            )
            if has_visible_text:
                continue
        sp = next((child for child in children if child.type == 0xF00A and len(child.payload) >= 8), None)
        has_anchor = any(child.type in (RT_OFFICEART_CLIENT_ANCHOR, RT_OFFICEART_CHILD_ANCHOR) for child in children)
        if sp is None or not has_anchor:
            continue
        flags = struct.unpack_from("<I", sp.payload, 4)[0]
        if flags & 0x1:  # group coordinator shape
            continue
        fopt = next((child for child in children if child.type == RT_OFFICEART_FOPT), None)
        if fopt is None:
            continue
        properties = _fopt_properties(fopt)
        if properties.get(260, 0) in image_map:
            continue
        fill, line, dash, fill_pattern, fill_back = _shape_style(properties, scheme)
        if fill is None and 385 not in properties and _has_fill(properties):
            # Non-text preset shapes inherit OfficeArt's white fill default.
            # Text boxes use different inheritance and must remain transparent
            # unless a fill color is explicit.
            fill = "FFFFFF"
        complex_props = _fopt_complex_properties(fopt)
        path = path_width = path_height = None
        if sp.instance == 19 and 325 in complex_props and 326 not in complex_props:
            path_info = _legacy_arc_path(complex_props)
            if path_info is not None:
                path, path_width, path_height = path_info
        elif _uses_custom_geometry(sp.instance, complex_props):
            path_info = _freeform_path(properties, complex_props)
            if path_info is not None:
                path, path_width, path_height = path_info
        if sp.instance == 0:
            if path is None:
                continue
            preset = "custom"
        else:
            preset = SHAPE_PRESETS.get(sp.instance)
            if preset is None and path is None:
                continue
            if path is not None:
                preset = "custom"
            elif preset is None:
                continue
        if fill is None and line is None and path is None:
            continue
        left, top, width, height = _anchor(children, len(result), space)
        transform = _combine_transform(
            _connector_transform(sp.instance, _transform(children, properties)),
            space,
        )
        result.append(BasicShape(
            preset=preset,
            left=left,
            top=top,
            width=width,
            height=height,
            fill_color=fill,
            line_color=line,
            rotation=transform[0],
            flip_horizontal=transform[1],
            flip_vertical=transform[2],
            line_dash=dash,
            path=path,
            path_width=path_width or 21600,
            path_height=path_height or 21600,
            line_width=_line_width(properties, space),
            fill_pattern=fill_pattern,
            fill_back_color=fill_back,
            line_head=_line_end(properties, 464, 466, 467),
            line_tail=_line_end(properties, 465, 468, 469),
            adjustments=_shape_adjustments(sp.instance, properties),
        ))
    return result

def _cluster_axis(positions: list[int], sizes: list[int]) -> list[dict[str, float]] | None:
    """Group sorted positions into clusters; each cluster keeps a running
    center and average size. Returns None for empty input."""
    if not positions:
        return None
    order = sorted(range(len(positions)), key=lambda i: positions[i])
    clusters: list[dict[str, float]] = []
    for i in order:
        p, s = positions[i], sizes[i]
        if clusters:
            tolerance = max(8.0, 0.25 * clusters[-1]["size"])
            if abs(p - clusters[-1]["center"]) <= tolerance:
                n = len(clusters[-1]["items"])
                clusters[-1]["center"] = (clusters[-1]["center"] * n + p) / (n + 1)
                clusters[-1]["size"] = (clusters[-1]["size"] * n + s) / (n + 1)
                clusters[-1]["items"].append(i)
                continue
        clusters.append({"center": float(p), "size": float(s), "items": [i]})
    return clusters


def _detect_tables(
    slide: Record,
    external_text: list[TextContent],
    fonts: tuple[str, ...],
    masters: dict[int, tuple[_MasterStyle, ...]],
    hyperlinks: dict[int, str],
    scheme: tuple[str, ...],
) -> tuple[list[Table], set[int]]:
    """Detect legacy PowerPoint tables (regular grids of text-bearing rectangle
    cells) and return ``(tables, excluded_offsets)``.

    Detection is intentionally strict: a table must form a complete ``R x C`` grid
    (``R >= 2`` and ``C >= 2``) of rectangle autoshape cells with consistent
    column widths and row heights. Looser layouts fall back to the existing
    flattened rendering so unrelated text arrangements are never mis-detected.
    The returned offsets let the caller suppress the flattened cell shapes and
    the misleading freeform warning that legacy gridline geometry would raise.
    """
    cells: list[dict[str, object]] = []
    thin_shapes: list[tuple[int, int, int, int, int]] = []  # offset, left, top, right, bottom
    for shape, space in _iter_sp_containers(slide):
        children = _direct_children(shape)
        if _is_background_shape(children) or _is_placeholder(shape):
            continue
        sp = next((child for child in children if child.type == 0xF00A and len(child.payload) >= 8), None)
        if sp is None:
            continue
        textbox = next((child for child in children if child.type == RT_OFFICEART_CLIENT_TEXTBOX), None)
        if textbox is not None:
            has_visible_text = any(
                child.type in (RT_TEXT_CHARS_ATOM, RT_TEXT_BYTES_ATOM, RT_OUTLINE_TEXT_REF_ATOM)
                for child in descendants(textbox)
            )
            if has_visible_text and sp.instance == 1:
                fopt = next((child for child in children if child.type == RT_OFFICEART_FOPT), None)
                properties = _fopt_properties(fopt) if fopt else {}
                left, top, width, height = _anchor(children, 0, space)
                contents = _text_contents(list(descendants(textbox)), fonts, masters, hyperlinks, scheme)
                content = contents[0] if contents else None
                if content is not None:
                    fill, _line, _dash, _fp, _fb = _shape_style(properties, scheme)
                    cells.append({
                        "offset": shape.offset,
                        "left": left, "top": top, "width": width, "height": height,
                        "content": content, "fill": fill,
                    })
                continue
        left, top, width, height = _anchor(children, 0, space)
        if width <= 2 or height <= 2:
            thin_shapes.append((shape.offset, left, top, left + width, top + height))
    if len(cells) < 4:
        return [], set()

    row_clusters = _cluster_axis([c["top"] for c in cells], [c["height"] for c in cells])
    col_clusters = _cluster_axis([c["left"] for c in cells], [c["width"] for c in cells])
    if row_clusters is None or col_clusters is None or len(row_clusters) < 2 or len(col_clusters) < 2:
        return [], set()

    grid: dict[tuple[int, int], dict[str, object]] = {}
    for c in cells:
        row = min(range(len(row_clusters)), key=lambda i: abs(c["top"] - row_clusters[i]["center"]))
        col = min(range(len(col_clusters)), key=lambda i: abs(c["left"] - col_clusters[i]["center"]))
        if (row, col) in grid:
            return [], set()
        grid[(row, col)] = c
    rows, cols = len(row_clusters), len(col_clusters)
    if len(grid) != rows * cols:
        return [], set()

    col_widths = [max(grid[(r, c)]["width"] for r in range(rows)) for c in range(cols)]
    row_heights = [max(grid[(r, c)]["height"] for c in range(cols)) for r in range(rows)]
    for c in cells:
        row = min(range(rows), key=lambda i: abs(c["top"] - row_clusters[i]["center"]))
        col = min(range(cols), key=lambda i: abs(c["left"] - col_clusters[i]["center"]))
        if abs(c["width"] - col_widths[col]) > 0.4 * col_widths[col] + 8:
            return [], set()
        if abs(c["height"] - row_heights[row]) > 0.4 * row_heights[row] + 8:
            return [], set()

    table_left = min(c["left"] for c in cells)
    table_top = min(c["top"] for c in cells)
    table_cells: list[TableCell] = []
    for row in range(rows):
        for col in range(cols):
            c = grid[(row, col)]
            content = c["content"]
            table_cells.append(TableCell(
                text=content.text,
                runs=content.runs,
                left=c["left"], top=c["top"], width=c["width"], height=c["height"],
                fill_color=c["fill"],
                row=row, col=col,
            ))
    table = Table(
        left=table_left, top=table_top,
        width=sum(col_widths), height=sum(row_heights),
        rows=rows, cols=cols, cells=tuple(table_cells),
    )
    excluded: set[int] = {c["offset"] for c in cells}
    for offset, l, t, r_, b in thin_shapes:
        if l >= table_left and r_ <= table_left + table.width and t >= table_top and b <= table_top + table.height:
            excluded.add(offset)
    return [table], excluded


def _background(
    slide: Record, scheme: tuple[str, ...]
) -> tuple[str | None, str | None, int | None, int | None]:
    for shape in descendants(slide):
        if shape.type != RT_OFFICEART_SP_CONTAINER:
            continue
        children = _direct_children(shape)
        sp = next((child for child in children if child.type == 0xF00A and len(child.payload) >= 8), None)
        if sp is None:
            continue
        _shape_id, flags = struct.unpack_from("<II", sp.payload)
        if not flags & 0x400:
            continue
        fopt = next((child for child in children if child.type == RT_OFFICEART_FOPT), None)
        if fopt:
            properties = _fopt_properties(fopt)
            color = _office_color(properties.get(385), scheme)
            if color:
                back = _office_color(properties.get(387), scheme)
                # fillType 4+ are gradients / shades in MS-ODRAW.
                fill_type = properties.get(384, 0)
                if fill_type >= 4 and back:
                    return back, color, _gradient_angle(properties), fill_type
                return color, None, None, None
    return None, None, None, None

def _comments(slide: Record) -> list[Comment]:
    result: list[Comment] = []
    for blob in descendants(slide):
        if blob.type != 5003:
            continue
        try:
            containers = records(blob.payload)
            for container in containers:
                if container.type != 12000 or container.version != CONTAINER_VERSION:
                    continue
                author = initials = text = ""
                left = top = 0
                created = None
                for child in records(container.payload):
                    if child.type == 4026:
                        value = child.payload.decode("utf-16le", "replace").rstrip("\0")
                        if child.instance == 0: author = value
                        elif child.instance == 1: text = value
                        elif child.instance == 2: initials = value
                    elif child.type == 12001 and len(child.payload) >= 28:
                        fields = struct.unpack_from("<I8Hii", child.payload)
                        _number, year, month, _weekday, day, hour, minute, second, milliseconds, left, top = fields
                        try:
                            created = datetime(year, month, day, hour, minute, second, milliseconds * 1000).isoformat(timespec="milliseconds")
                        except ValueError:
                            created = None
                if text:
                    result.append(Comment(author or "Unknown", initials, text, left, top, created))
        except InvalidPpt:
            continue
    return result

def _pictures(document: Record, stream: bytes | None) -> dict[int, tuple[bytes, str, str]]:
    if not stream:
        return {}
    result: dict[int, tuple[bytes, str, str]] = {}
    for index, bse in enumerate((item for item in descendants(document) if item.type == RT_OFFICEART_BSE), 1):
        if len(bse.payload) < 32:
            continue
        offset = struct.unpack_from("<I", bse.payload, 28)[0]
        if offset >= len(stream):
            continue
        try:
            blip = next(records(stream, offset))
        except (InvalidPpt, StopIteration):
            continue
        signatures = ((b"\x89PNG\r\n\x1a\n", "png", "image/png"),
                      (b"\xff\xd8\xff", "jpg", "image/jpeg"),
                      (b"GIF87a", "gif", "image/gif"), (b"GIF89a", "gif", "image/gif"),
                      (b"II*\x00", "tif", "image/tiff"), (b"MM\x00*", "tif", "image/tiff"))
        for signature, extension, content_type in signatures:
            position = blip.payload.find(signature, 0, 40)
            if position >= 0:
                result[index] = (blip.payload[position:], extension, content_type)
                break
        if index not in result and blip.type in (0xF01A, 0xF01B, 0xF01C):
            uid_counts = {0xF01A: 1 if blip.instance == 0x3D4 else 2,
                          0xF01B: 1 if blip.instance == 0x216 else 2,
                          0xF01C: 1 if blip.instance == 0x542 else 2}
            header_start = uid_counts[blip.type] * 16
            if header_start + 34 > len(blip.payload):
                continue
            original_size = struct.unpack_from("<I", blip.payload, header_start)[0]
            left, top, right, bottom = struct.unpack_from("<4i", blip.payload, header_start + 4)
            compressed_size = struct.unpack_from("<I", blip.payload, header_start + 28)[0]
            compression = blip.payload[header_start + 32]
            compressed = blip.payload[header_start + 34:header_start + 34 + compressed_size]
            try:
                if compression == 0:
                    inflater = zlib.decompressobj()
                    raw = inflater.decompress(compressed, min(original_size + 1, 100_000_001))
                    if not inflater.eof or len(raw) > 100_000_000:
                        continue
                else:
                    raw = compressed
            except zlib.error:
                continue
            if blip.type == 0xF01B:
                # OOXML embeds the standard WMF stream.  Adding an Aldus
                # placeable header makes LibreOffice render its transparent
                # background as an opaque white rectangle.
                if raw.startswith(b"\xd7\xcd\xc6\x9a") and len(raw) >= 22:
                    raw = raw[22:]
                result[index] = (raw, "wmf", "image/x-wmf")
            elif blip.type == 0xF01A:
                result[index] = (raw, "emf", "image/x-emf")
            else:
                result[index] = (bytes(512) + raw, "pct", "image/x-pict")
    return result

def _shape_pictures(
    slide: Record,
    image_map: dict[int, tuple[bytes, str, str]],
    scheme: tuple[str, ...] = (),
) -> list[Picture]:
    result: list[Picture] = []
    for shape, space in _iter_sp_containers(slide):
        children = _direct_children(shape)
        fopt = next((child for child in children if child.type == RT_OFFICEART_FOPT), None)
        if fopt is None:
            continue
        properties = _fopt_properties(fopt)
        reference = properties.get(260)
        image = image_map.get(reference or 0)
        if image is None:
            continue
        left, top, width, height = _anchor(children, len(result), space)
        data, extension, content_type = image
        def crop(property_id: int) -> int:
            raw = properties.get(property_id, 0)
            signed = struct.unpack("<i", struct.pack("<I", raw))[0]
            return round(signed / 65536 * 100000)
        transform = _combine_transform(_transform(children, properties), space)
        transparent_color = _office_color(properties.get(263), scheme)
        result.append(Picture(
            data=data,
            extension=extension,
            content_type=content_type,
            left=left,
            top=top,
            width=width,
            height=height,
            crop_left=crop(258),
            crop_top=crop(256),
            crop_right=crop(259),
            crop_bottom=crop(257),
            rotation=transform[0],
            flip_horizontal=transform[1],
            flip_vertical=transform[2],
            transparent_color=transparent_color,
            line_color=(
                (_office_color(properties.get(448), scheme) or "000000")
                if _has_line(properties) else None
            ),
            line_dash=_line_dash(properties) if _has_line(properties) else None,
            line_width=_line_width(properties, space),
        ))
    return result

def _presentation_size(document: Record) -> tuple[int, int]:
    atom = next((child for child in descendants(document) if child.type == RT_DOCUMENT_ATOM), None)
    if atom is None or len(atom.payload) < 8:
        return DEFAULT_SLIDE_WIDTH, DEFAULT_SLIDE_HEIGHT
    width, height = struct.unpack_from("<2I", atom.payload)
    if not width or not height or width > 100_000 or height > 100_000:
        return DEFAULT_SLIDE_WIDTH, DEFAULT_SLIDE_HEIGHT
    return width, height

def _external_slide_text(document: Record, fonts: tuple[str, ...], masters: dict[int, tuple[_MasterStyle, ...]], hyperlinks: dict[int, str], scheme: tuple[str, ...]) -> dict[int, list[TextContent]]:
    """Map slide persist references to text held in SlideListWithText.

    PowerPoint 97-era files commonly keep the actual text in the document
    container and put only an OutlineTextRefAtom in each drawing shape.
    """
    grouped: dict[int, list[Record]] = {}
    for container in descendants(document):
        if container.type != RT_SLIDE_LIST_WITH_TEXT or container.instance != 0:
            continue
        current_reference: int | None = None
        for child in records(container.payload):
            if child.type == RT_SLIDE_PERSIST_ATOM and len(child.payload) >= 4:
                current_reference = struct.unpack_from("<I", child.payload)[0]
                grouped.setdefault(current_reference, [])
                continue
            if current_reference is not None:
                grouped[current_reference].append(child)
    return {reference: _text_contents(items, fonts, masters, hyperlinks, scheme) for reference, items in grouped.items()}

def _color_scheme(document: Record, masters: tuple[Record, ...]) -> tuple[str, ...]:
    candidates: list[Record] = []
    for master in masters:
        candidates.extend(child for child in descendants(master) if child.type == 2032 and child.instance == 1)
    candidates.extend(child for child in descendants(document) if child.type == 2032 and child.instance == 1)
    if not candidates:
        return ("FFFFFF", "000000", "808080", "000000", "FFFFFF", "0000FF", "FF0000", "FFFF00")
    payload = candidates[0].payload
    if len(payload) < 32:
        return ()
    return tuple(f"{payload[index]:02X}{payload[index + 1]:02X}{payload[index + 2]:02X}"
                 for index in range(0, 32, 4))

def _header_footer(record: Record, base: HeaderFooter | None = None, *, instance: int) -> HeaderFooter | None:
    container = next((child for child in descendants(record)
                      if child.type == 4057 and child.version == CONTAINER_VERSION and child.instance == instance), None)
    if container is None:
        return base
    children = list(records(container.payload))
    atom = next((child for child in children if child.type == 4058 and len(child.payload) >= 4), None)
    if atom is None:
        return base
    mask = struct.unpack_from("<H", atom.payload, 2)[0]
    strings = {child.instance: child.payload.decode("utf-16le", "replace").rstrip("\x00")
               for child in children if child.type == 4026}
    inherited = base or HeaderFooter()
    date_text = strings.get(0) if mask & 0x04 else None
    header_text = strings.get(1, inherited.header_text) if mask & 0x10 else None
    footer_text = strings.get(2, inherited.footer_text) if mask & 0x20 else None
    return HeaderFooter(
        date_text=date_text,
        date_is_auto=bool(mask & 0x01 and mask & 0x02),
        header_text=header_text,
        footer_text=footer_text,
        show_slide_number=bool(mask & 0x08),
    )

def _parse_slide(slide_record: Record, image_map: dict[int, tuple[bytes, str, str]], external_text: list[TextContent], fonts: tuple[str, ...], masters: dict[int, tuple[_MasterStyle, ...]], hyperlinks: dict[int, str], header_footer: HeaderFooter | None, scheme: tuple[str, ...], master_record: Record | None, notes: tuple[str, ...], slide_width: int, slide_height: int) -> Slide:
    tables, table_excluded = _detect_tables(
        slide_record, external_text, fonts, masters, hyperlinks, scheme
    )
    slide_boxes = _shape_text_boxes(
        slide_record, external_text, fonts, masters, hyperlinks, scheme,
        exclude=table_excluded,
    )
    if not slide_boxes:
        texts = [value for child in descendants(slide_record) if (value := _text(child))]
        slide_boxes = [
            TextBox(text, 288, 288 + i * 576, 5184, 432, (TextRun(text),))
            for i, text in enumerate(texts)
        ]
    background, background_end, background_angle, background_type = _background(
        slide_record, scheme
    )
    if master_record is not None:
        (
            master_background,
            master_background_end,
            master_background_angle,
            master_background_type,
        ) = _background(master_record, scheme)
        # Prefer master's explicit gradient when the slide only has a flat scheme fill.
        if master_background_end and not background_end:
            background, background_end = master_background, master_background_end
            background_angle, background_type = (
                master_background_angle,
                master_background_type,
            )
        elif background is None:
            background, background_end = master_background, master_background_end
            background_angle, background_type = (
                master_background_angle,
                master_background_type,
            )
    # Flatten non-placeholder master decorations onto the slide so common
    # template chrome survives conversion to a blank OOXML layout.
    master_boxes: list[TextBox] = []
    master_pictures: list[Picture] = []
    shapes: list[BasicShape] = []
    if master_record is not None:
        for box in _shape_text_boxes(master_record, [], fonts, masters, hyperlinks, scheme,
                                     skip_placeholders=True):
            if box.text.strip() in ("", "*"):
                continue
            master_boxes.append(box)
        for picture in _shape_pictures(master_record, image_map, scheme):
            master_pictures.append(picture)
        for shape in _basic_shapes(master_record, image_map, scheme):
            left, top = max(0, shape.left), max(0, shape.top)
            right = min(slide_width, shape.left + shape.width)
            bottom = min(slide_height, shape.top + shape.height)
            if right > left and bottom > top:
                shapes.append(BasicShape(shape.preset, left, top, right - left, bottom - top,
                                         shape.fill_color, shape.line_color, shape.rotation,
                                         shape.flip_horizontal, shape.flip_vertical,
                                         shape.line_dash, shape.path, shape.path_width,
                                         shape.path_height, shape.line_width,
                                         shape.fill_pattern, shape.fill_back_color,
                                         shape.line_head, shape.line_tail,
                                         shape.adjustments))
    for shape in _basic_shapes(slide_record, image_map, scheme, exclude=table_excluded):
        left, top = max(0, shape.left), max(0, shape.top)
        right = min(slide_width, shape.left + shape.width)
        bottom = min(slide_height, shape.top + shape.height)
        if right > left and bottom > top:
            shapes.append(BasicShape(shape.preset, left, top, right - left, bottom - top,
                                     shape.fill_color, shape.line_color, shape.rotation,
                                     shape.flip_horizontal, shape.flip_vertical,
                                     shape.line_dash, shape.path, shape.path_width,
                                     shape.path_height, shape.line_width,
                                     shape.fill_pattern, shape.fill_back_color,
                                     shape.line_head, shape.line_tail,
                                     shape.adjustments))
    boxes = master_boxes + slide_boxes
    pictures = master_pictures + list(
        _shape_pictures(slide_record, image_map, scheme)
    )
    slide_show_info = next((child for child in descendants(slide_record)
                            if child.type == RT_SLIDE_SHOW_SLIDE_INFO_ATOM
                            and len(child.payload) >= 12), None)
    hidden = bool(struct.unpack_from("<H", slide_show_info.payload, 10)[0] & 0x0004) if slide_show_info else False
    return Slide(
        text_boxes=tuple(boxes),
        pictures=tuple(pictures),
        shapes=tuple(shapes),
        background_color=background,
        comments=tuple(_comments(slide_record)),
        header_footer=_header_footer(slide_record, header_footer, instance=0),
        background_color_end=background_end,
        notes=notes,
        hidden=hidden,
        background_gradient_angle=background_angle,
        background_gradient_type=background_type,
        tables=tuple(tables),
        excluded_offsets=frozenset(table_excluded),
    )

def extract_presentation(powerpoint_document: bytes, pictures_stream: bytes | None = None) -> Presentation:
    roots = list(records(powerpoint_document))
    document = next((r for r in roots if r.type == RT_DOCUMENT and r.version == CONTAINER_VERSION), None)
    if document is None:
        raise InvalidPpt("PowerPoint Document record is missing")
    mapping = persist_directory(powerpoint_document)
    image_map = _pictures(document, pictures_stream)
    fonts = _fonts(document)
    master_records = tuple(root for root in roots if root.type == 1016)
    scheme = _color_scheme(document, master_records)
    masters = _master_text_styles(powerpoint_document, document, fonts, scheme)
    hyperlinks = _hyperlinks(document)
    external_text = _external_slide_text(document, fonts, masters, hyperlinks, scheme)
    header_footer = _header_footer(document, instance=3)
    notes_by_slide_id: dict[int, tuple[str, ...]] = {}
    unbound_note_values: list[tuple[str, ...]] = []
    for reference in _notes_refs(document):
        offset = mapping.get(reference)
        if offset is None or offset >= len(powerpoint_document):
            unbound_note_values.append(())
            continue
        try:
            note_record = next(records(powerpoint_document, offset))
        except (InvalidPpt, StopIteration):
            unbound_note_values.append(())
            continue
        if note_record.type != 1008:
            unbound_note_values.append(())
            continue
        values = tuple(box.text for box in _shape_text_boxes(note_record, [], fonts, masters, hyperlinks, scheme)
                       if box.text.strip() not in ("", "*"))
        notes_atom = next(
            (
                child
                for child in descendants(note_record)
                if child.type == 1009 and len(child.payload) >= 4
            ),
            None,
        )
        slide_id = (
            struct.unpack_from("<I", notes_atom.payload)[0]
            if notes_atom is not None
            else 0
        )
        if slide_id:
            notes_by_slide_id[slide_id] = values
        else:
            unbound_note_values.append(values)
    width, height = _presentation_size(document)
    slides: list[Slide] = []
    seen_offsets: set[int] = set()
    unbound_note_index = 0
    for reference, slide_id in _slide_entries(document):
        offset = mapping.get(reference)
        if offset is None or offset >= len(powerpoint_document) or offset in seen_offsets:
            continue
        try:
            slide_record = next(records(powerpoint_document, offset))
        except (InvalidPpt, StopIteration):
            continue
        if slide_record.type != RT_SLIDE:
            continue
        seen_offsets.add(offset)
        notes = notes_by_slide_id.get(slide_id)
        if notes is None:
            notes = (
                unbound_note_values[unbound_note_index]
                if unbound_note_index < len(unbound_note_values)
                else ()
            )
            unbound_note_index += 1
        slides.append(_parse_slide(slide_record, image_map, external_text.get(reference, []), fonts, masters, hyperlinks, header_footer, scheme, master_records[0] if master_records else None, notes, width, height))
    if not slides:
        for slide_record in descendants(document):
            if slide_record.type == RT_SLIDE:
                slides.append(_parse_slide(slide_record, image_map, [], fonts, masters, hyperlinks, header_footer, scheme, master_records[0] if master_records else None, (), width, height))
    excluded_offsets: set[int] = set()
    for slide in slides:
        excluded_offsets |= slide.excluded_offsets
    return Presentation(width, height, tuple(slides), excluded_offsets=frozenset(excluded_offsets))

def extract_slides(powerpoint_document: bytes) -> list[list[str]]:
    return [[box.text for box in slide.text_boxes] for slide in extract_presentation(powerpoint_document).slides]

_LOSSY_RECORD_CODES: dict[int, tuple[str, str]] = {
    RT_ROUNDTRIP_ANIMATION_ATOM: ("ANIMATION_OMITTED", "animation timeline was omitted"),
    RT_ROUNDTRIP_ANIMATION_HASH_ATOM: ("ANIMATION_OMITTED", "animation timeline was omitted"),
    RT_CHART_BUILD: ("ANIMATION_OMITTED", "chart build animation was omitted"),
    RT_CHART_BUILD_ATOM: ("ANIMATION_OMITTED", "chart build animation was omitted"),
    RT_SOUND_COLLECTION: ("AUDIO_OMITTED", "embedded audio was omitted"),
    RT_SOUND: ("AUDIO_OMITTED", "embedded audio was omitted"),
    RT_SOUND_DATA_BLOB: ("AUDIO_OMITTED", "embedded audio was omitted"),
    RT_EXTERNAL_VIDEO: ("VIDEO_OMITTED", "embedded video was omitted"),
    RT_EXTERNAL_AVI_MOVIE: ("VIDEO_OMITTED", "embedded video was omitted"),
    RT_EXTERNAL_MCI_MOVIE: ("VIDEO_OMITTED", "embedded video was omitted"),
    RT_DIAGRAM_BUILD: ("DIAGRAM_OR_SMARTART_OMITTED", "diagram/SmartArt build was omitted"),
    RT_DIAGRAM_BUILD_ATOM: ("DIAGRAM_OR_SMARTART_OMITTED", "diagram/SmartArt build was omitted"),
}

_OLE_TYPE_DIAGNOSTICS = {
    0: (
        "EMBEDDED_OLE_OMITTED",
        "embedded OLE payload and editability were omitted; preview media may be preserved",
        "ole",
    ),
    1: (
        "LINKED_OLE_OMITTED",
        "linked OLE behavior was omitted; preview media may be preserved",
        "linked_ole",
    ),
    2: (
        "ACTIVEX_CONTROL_OMITTED",
        "ActiveX/OLE control behavior was omitted; preview media may be preserved",
        "activex_control",
    ),
}
_OLE_CONTAINER_DIAGNOSTICS = {
    RT_EXTERNAL_OLE_EMBED: _OLE_TYPE_DIAGNOSTICS[0],
    RT_EXTERNAL_OLE_LINK: _OLE_TYPE_DIAGNOSTICS[1],
    RT_EXTERNAL_OLE_CONTROL: _OLE_TYPE_DIAGNOSTICS[2],
}

_CHART_MARKERS = (
    "MSGraph.Chart",
    "Excel.Chart",
    "MSGraph",
    "orgchart",
)
_DIAGRAM_MARKERS = (
    "SmartArt",
    "Office.SmartArt",
    "schemas.openxmlformats.org/drawingml/2006/diagram",
)

def _utf16_payload_text(payload: bytes) -> str:
    try:
        return payload.decode("utf-16le", errors="ignore")
    except Exception:
        return ""

def _payload_has_marker(payload: bytes, markers: tuple[str, ...]) -> bool:
    text = _utf16_payload_text(payload)
    folded = text.casefold()
    for marker in markers:
        if marker.casefold() in folded:
            return True
        if marker.encode("utf-16le") in payload:
            return True
        if marker.encode("ascii") in payload:
            return True
    return False

def _iter_all_records(data: bytes):
    """Yield every record with offsets absolute to ``data``."""

    def walk(start: int, end: int):
        for record in records(data, start, end):
            yield record
            if record.version != CONTAINER_VERSION:
                continue
            if RT_ROUNDTRIP_OPAQUE_MIN <= record.type <= RT_ROUNDTRIP_OPAQUE_MAX:
                continue
            payload_start = record.offset + 8
            yield from walk(payload_start, payload_start + len(record.payload))

    yield from walk(0, len(data))

def _slide_byte_ranges(data: bytes) -> tuple[tuple[int, int, int], ...]:
    """Return (1-based slide index, start offset, end offset) for each normal slide."""
    roots = list(records(data))
    document = next((root for root in roots if root.type == RT_DOCUMENT and root.version == CONTAINER_VERSION), None)
    if document is None:
        return ()
    mapping = persist_directory(data)
    ranges: list[tuple[int, int, int]] = []
    seen: set[int] = set()
    for index, reference in enumerate(_slide_refs(document), 1):
        offset = mapping.get(reference)
        if offset is None or offset in seen or offset >= len(data):
            continue
        try:
            slide = next(records(data, offset))
        except (InvalidPpt, StopIteration):
            continue
        if slide.type != RT_SLIDE:
            continue
        seen.add(offset)
        ranges.append((index, slide.offset, slide.offset + 8 + len(slide.payload)))
    if ranges:
        return tuple(ranges)
    # Synthetic streams may embed slides directly under Document without persist ids.
    embedded: list[tuple[int, int, int]] = []
    payload_start = document.offset + 8
    payload_end = payload_start + len(document.payload)
    for index, slide in enumerate(
        (record for record in records(data, payload_start, payload_end) if record.type == RT_SLIDE),
        1,
    ):
        embedded.append((index, slide.offset, slide.offset + 8 + len(slide.payload)))
    return tuple(embedded)

def _slide_index_for_offset(offset: int, ranges: tuple[tuple[int, int, int], ...]) -> int | None:
    for slide_index, start, end in ranges:
        if start <= offset < end:
            return slide_index
    return None

def _object_kind_for_code(code: str) -> str:
    return {
        "ANIMATION_OMITTED": "animation",
        "AUDIO_OMITTED": "audio",
        "VIDEO_OMITTED": "video",
        "EMBEDDED_OLE_OMITTED": "ole",
        "LINKED_OLE_OMITTED": "linked_ole",
        "ACTIVEX_CONTROL_OMITTED": "activex_control",
        "CHART_OMITTED": "chart",
        "DIAGRAM_OR_SMARTART_OMITTED": "diagram",
        "COMPLEX_FREEFORM_OMITTED": "freeform",
    }.get(code, "object")


def _animation_object_diagnostics(
    data: bytes, ranges: tuple[tuple[int, int, int], ...]
) -> tuple[
    tuple[str, str, tuple[int, ...], int, tuple[LossyFeatureLocation, ...]], ...
]:
    """Report one object per legacy or PP10 shape-animation effect."""
    all_records = list(_iter_all_records(data))

    def direct_children(record: Record) -> list[Record]:
        start = record.offset + 8
        return list(records(data, start, start + len(record.payload)))

    def tree(record: Record):
        yield record
        if record.version != CONTAINER_VERSION:
            return
        for child in direct_children(record):
            yield from tree(child)

    shape_containers = [
        record for record in all_records if record.type == RT_OFFICEART_SP_CONTAINER
    ]
    legacy: list[tuple[Record, tuple[Record, ...], int | None, int | None]] = []
    handled_atoms: set[int] = set()
    for container in (
        record for record in all_records if record.type == RT_ANIMATION_INFO
    ):
        start = container.offset + 8
        end = start + len(container.payload)
        atoms = tuple(
            record
            for record in all_records
            if record.type == RT_ANIMATION_INFO_ATOM
            and start <= record.offset < end
        )
        handled_atoms.update(record.offset for record in atoms)
        owner = min(
            (
                shape
                for shape in shape_containers
                if shape.offset + 8 <= container.offset
                < shape.offset + 8 + len(shape.payload)
            ),
            key=lambda shape: len(shape.payload),
            default=None,
        )
        shape_id = None
        if owner is not None:
            fsp = next(
                (
                    child
                    for child in direct_children(owner)
                    if child.type == 0xF00A and len(child.payload) >= 4
                ),
                None,
            )
            if fsp is not None:
                shape_id = struct.unpack_from("<I", fsp.payload)[0]
        legacy.append(
            (
                container,
                atoms,
                _slide_index_for_offset(container.offset, ranges),
                shape_id,
            )
        )

    diagnostics: list[
        tuple[str, str, tuple[int, ...], int, tuple[LossyFeatureLocation, ...]]
    ] = []
    handled_legacy: set[int] = set()
    seen_timing_nodes: set[int] = set()
    for blob in (
        record for record in all_records if record.type == RT_BINARY_TAG_DATA_BLOB
    ):
        blob_start = blob.offset + 8
        try:
            top_level = list(records(data, blob_start, blob_start + len(blob.payload)))
        except InvalidPpt:
            continue
        roots = [
            record
            for record in top_level
            if record.type == RT_TIME_EXT_TIME_NODE
            and record.version == CONTAINER_VERSION
            and record.instance == 1
        ]
        for root in roots:
            for node in tree(root):
                if (
                    node.type != RT_TIME_EXT_TIME_NODE
                    or node.offset in seen_timing_nodes
                ):
                    continue
                seen_timing_nodes.add(node.offset)
                children = direct_children(node)
                time_node = next(
                    (child for child in children if child.type == RT_TIME_NODE),
                    None,
                )
                property_list = next(
                    (
                        child
                        for child in children
                        if child.type == RT_TIME_PROPERTY_LIST
                    ),
                    None,
                )
                if property_list is None:
                    continue
                variants = [
                    child
                    for child in direct_children(property_list)
                    if child.type == RT_TIME_VARIANT
                ]
                # TL_TPID_EffectType (instance 11) marks an actual effect node;
                # root, sequence, trigger, and behavior nodes are scaffolding.
                effect_types = [
                    struct.unpack_from("<i", variant.payload, 1)[0]
                    for variant in variants
                    if variant.instance == 11
                    and len(variant.payload) >= 5
                    and variant.payload[0] == 1
                ]
                if not any(1 <= value <= 6 for value in effect_types):
                    continue
                shape_ids = {
                    struct.unpack_from("<I", item.payload, 8)[0]
                    for item in tree(node)
                    if item.type == RT_VISUAL_SHAPE_ATOM
                    and len(item.payload) >= 12
                }
                slide_index = _slide_index_for_offset(node.offset, ranges)
                shape_id = next(iter(shape_ids)) if len(shape_ids) == 1 else None
                matched = next(
                    (
                        index
                        for index, (_, _, legacy_slide, legacy_shape) in enumerate(legacy)
                        if index not in handled_legacy
                        and legacy_slide == slide_index
                        and shape_id is not None
                        and legacy_shape == shape_id
                    ),
                    None,
                )
                record_types = {
                    node.type,
                    property_list.type,
                    *(variant.type for variant in variants),
                }
                if time_node is not None:
                    record_types.add(time_node.type)
                if matched is not None:
                    handled_legacy.add(matched)
                    legacy_container, atoms, _, _ = legacy[matched]
                    record_types.add(legacy_container.type)
                    record_types.update(atom.type for atom in atoms)
                diagnostics.append(
                    (
                        "ANIMATION_OMITTED",
                        "shape animation timeline and effect were omitted",
                        tuple(sorted(record_types)),
                        node.offset,
                        (
                            LossyFeatureLocation(
                                slide_index=slide_index,
                                record_type=node.type,
                                record_offset=node.offset,
                                object_kind="animation",
                            ),
                        ),
                    )
                )

    for index, (container, atoms, slide_index, _shape_id) in enumerate(legacy):
        if index in handled_legacy:
            continue
        diagnostics.append(
            (
                "ANIMATION_OMITTED",
                "shape animation timing and effect were omitted",
                tuple(sorted({container.type, *(record.type for record in atoms)})),
                container.offset,
                (
                    LossyFeatureLocation(
                        slide_index=slide_index,
                        record_type=container.type,
                        record_offset=container.offset,
                        object_kind="animation",
                    ),
                ),
            )
        )
    for atom in all_records:
        if (
            atom.type != RT_ANIMATION_INFO_ATOM
            or atom.offset in handled_atoms
        ):
            continue
        diagnostics.append(
            (
                "ANIMATION_OMITTED",
                "orphaned shape animation information was omitted",
                (atom.type,),
                atom.offset,
                (
                    LossyFeatureLocation(
                        slide_index=_slide_index_for_offset(atom.offset, ranges),
                        record_type=atom.type,
                        record_offset=atom.offset,
                        object_kind="animation",
                    ),
                ),
            )
        )
    return tuple(diagnostics)


def _ole_object_diagnostics(
    data: bytes, ranges: tuple[tuple[int, int, int], ...]
) -> tuple[
    tuple[
        tuple[str, str, tuple[int, ...], int, tuple[LossyFeatureLocation, ...]],
        ...,
    ],
    frozenset[int],
]:
    """Collapse OLE storage records into external objects and bind them to slides."""
    all_records = list(_iter_all_records(data))
    ole_containers = [
        record for record in all_records if record.type in _OLE_CONTAINER_DIAGNOSTICS
    ]
    references: dict[int, list[Record]] = {}
    atoms: list[tuple[Record, int, int]] = []
    for record in all_records:
        if record.type == RT_EXTERNAL_OBJECT_REF_ATOM and len(record.payload) >= 4:
            ex_obj_id = struct.unpack_from("<I", record.payload)[0]
            references.setdefault(ex_obj_id, []).append(record)
        elif record.type == RT_EXTERNAL_OLE_OBJECT_ATOM and len(record.payload) >= 20:
            ole_type = struct.unpack_from("<I", record.payload, 4)[0]
            ex_obj_id = struct.unpack_from("<I", record.payload, 8)[0]
            atoms.append((record, ole_type, ex_obj_id))

    diagnostics: list[
        tuple[str, str, tuple[int, ...], int, tuple[LossyFeatureLocation, ...]]
    ] = []
    handled_chart_markers: set[int] = set()
    for atom, ole_type, ex_obj_id in atoms:
        code, message, object_kind = _OLE_TYPE_DIAGNOSTICS.get(
            ole_type,
            ("OLE_OBJECT_OMITTED", "OLE object of unknown type was omitted", "ole"),
        )
        container = next(
            (
                item
                for item in ole_containers
                if item.offset + 8 <= atom.offset < item.offset + 8 + len(item.payload)
            ),
            None,
        )
        chart_markers = [
            item
            for item in all_records
            if container is not None
            and container.offset + 8 <= item.offset < container.offset + 8 + len(container.payload)
            and item.type in (RT_CSTRING, RT_PROG_BINARY_TAG, RT_BINARY_TAG_DATA_BLOB)
            and _payload_has_marker(item.payload, _CHART_MARKERS)
        ]
        if chart_markers:
            code = "CHART_OMITTED"
            message = (
                "legacy chart data and editability were omitted; "
                "preview media may be preserved"
            )
            object_kind = "chart"
            handled_chart_markers.update(item.offset for item in chart_markers)
        object_references = references.get(ex_obj_id, [])
        locations = tuple(
            LossyFeatureLocation(
                slide_index=_slide_index_for_offset(reference.offset, ranges),
                record_type=reference.type,
                record_offset=reference.offset,
                object_kind=object_kind,
            )
            for reference in object_references
        )
        if not locations:
            locations = (
                LossyFeatureLocation(
                    slide_index=_slide_index_for_offset(atom.offset, ranges),
                    record_type=atom.type,
                    record_offset=atom.offset,
                    object_kind=object_kind,
                ),
            )
        record_types = tuple(
            sorted(
                {
                    atom.type,
                    *(item.type for item in object_references),
                    *(item.type for item in chart_markers),
                }
            )
        )
        diagnostics.append((code, message, record_types, atom.offset, locations))

    atom_offsets = {atom.offset for atom, _ole_type, _ex_obj_id in atoms}
    fallback_containers = 0
    for record in all_records:
        mapped = _OLE_CONTAINER_DIAGNOSTICS.get(record.type)
        if mapped is None:
            continue
        start = record.offset + 8
        end = start + len(record.payload)
        if any(start <= atom_offset < end for atom_offset in atom_offsets):
            continue
        code, message, object_kind = mapped
        diagnostics.append(
            (
                code,
                message,
                (record.type,),
                record.offset,
                (
                    LossyFeatureLocation(
                        slide_index=_slide_index_for_offset(record.offset, ranges),
                        record_type=record.type,
                        record_offset=record.offset,
                        object_kind=object_kind,
                    ),
                ),
            )
        )
        fallback_containers += 1

    if not atoms and fallback_containers == 0:
        for record in all_records:
            if record.type != RT_EXTERNAL_OLE_OBJECT_STG:
                continue
            diagnostics.append(
                (
                    "EMBEDDED_OLE_OMITTED",
                    "orphaned embedded OLE storage was omitted",
                    (record.type,),
                    record.offset,
                    (
                        LossyFeatureLocation(
                            slide_index=_slide_index_for_offset(record.offset, ranges),
                            record_type=record.type,
                            record_offset=record.offset,
                            object_kind="ole",
                        ),
                    ),
                )
            )
    return tuple(diagnostics), frozenset(handled_chart_markers)


def _unparsed_freeform_locations(
    data: bytes, ranges: tuple[tuple[int, int, int], ...], exclude: set[int] | None = None
) -> tuple[LossyFeatureLocation, ...]:
    locations: list[LossyFeatureLocation] = []
    for record in _iter_all_records(data):
        if record.type != RT_OFFICEART_SP_CONTAINER:
            continue
        if exclude is not None and record.offset in exclude:
            continue
        children = list(records(record.payload))
        sp = next((child for child in children if child.type == 0xF00A and len(child.payload) >= 8), None)
        fopt = next((child for child in children if child.type == RT_OFFICEART_FOPT), None)
        if sp is None or fopt is None:
            continue
        complex_props = _fopt_complex_properties(fopt)
        if not _uses_custom_geometry(sp.instance, complex_props):
            continue
        properties = _fopt_properties(fopt)
        if _freeform_path(properties, complex_props) is None:
            locations.append(
                LossyFeatureLocation(
                    slide_index=_slide_index_for_offset(record.offset, ranges),
                    record_type=record.type,
                    record_offset=record.offset,
                    object_kind="freeform",
                )
            )
    return tuple(locations)

def detect_lossy_features(powerpoint_document: bytes, exclude_offsets: set[int] | None = None) -> tuple[LossyFeature, ...]:
    """Return object-backed loss diagnostics for unsupported/approximated content."""
    ranges = _slide_byte_ranges(powerpoint_document)
    buckets: dict[str, dict[str, object]] = {}
    animation_diagnostics = _animation_object_diagnostics(
        powerpoint_document, ranges
    )
    ole_diagnostics, handled_chart_markers = _ole_object_diagnostics(
        powerpoint_document, ranges
    )

    def add(
        code: str,
        message: str,
        *,
        record_type: int | None = None,
        record_offset: int | None = None,
        amount: int = 1,
        locations: tuple[LossyFeatureLocation, ...] = (),
        record_types: tuple[int, ...] = (),
    ) -> None:
        entry = buckets.setdefault(
            code,
            {"message": message, "count": 0, "types": set(), "locations": []},
        )
        entry["count"] = int(entry["count"]) + amount
        collected = entry["locations"]
        assert isinstance(collected, list)
        if locations:
            collected.extend(locations)
        elif record_type is not None and record_offset is not None:
            collected.append(
                LossyFeatureLocation(
                    slide_index=_slide_index_for_offset(record_offset, ranges),
                    record_type=record_type,
                    record_offset=record_offset,
                    object_kind=_object_kind_for_code(code),
                )
            )
        if record_type is not None:
            types = entry["types"]
            assert isinstance(types, set)
            types.add(record_type)
        if record_types:
            types = entry["types"]
            assert isinstance(types, set)
            types.update(record_types)

    for record in _iter_all_records(powerpoint_document):
        mapped = _LOSSY_RECORD_CODES.get(record.type)
        if mapped is not None:
            code, message = mapped
            add(code, message, record_type=record.type, record_offset=record.offset)
        if record.type in (RT_CSTRING, RT_PROG_BINARY_TAG, RT_BINARY_TAG_DATA_BLOB):
            if (
                record.offset not in handled_chart_markers
                and _payload_has_marker(record.payload, _CHART_MARKERS)
            ):
                add(
                    "CHART_OMITTED",
                    "chart content was omitted or left as non-editable media",
                    record_type=record.type,
                    record_offset=record.offset,
                )
            if _payload_has_marker(record.payload, _DIAGRAM_MARKERS):
                add(
                    "DIAGRAM_OR_SMARTART_OMITTED",
                    "diagram/SmartArt content was omitted or left as non-editable media",
                    record_type=record.type,
                    record_offset=record.offset,
                )

    for code, message, record_types, record_offset, locations in animation_diagnostics:
        add(
            code,
            message,
            record_offset=record_offset,
            amount=1,
            locations=locations,
            record_types=record_types,
        )

    for code, message, record_types, record_offset, locations in ole_diagnostics:
        add(
            code,
            message,
            record_offset=record_offset,
            amount=1,
            locations=locations,
            record_types=record_types,
        )

    freeform_locations = _unparsed_freeform_locations(powerpoint_document, ranges, exclude_offsets)
    if freeform_locations:
        add(
            "COMPLEX_FREEFORM_OMITTED",
            "complex freeform geometry could not be reconstructed as an editable path",
            amount=len(freeform_locations),
            locations=freeform_locations,
        )

    features: list[LossyFeature] = []
    for code, entry in sorted(buckets.items()):
        types = entry["types"]
        assert isinstance(types, set)
        locations = entry["locations"]
        assert isinstance(locations, list)
        # Keep a stable, de-duplicated location list for report consumers.
        unique: list[LossyFeatureLocation] = []
        seen: set[tuple[int | None, int, int, str]] = set()
        for location in locations:
            key = (
                location.slide_index,
                location.record_type,
                location.record_offset,
                location.object_kind,
            )
            if key in seen:
                continue
            seen.add(key)
            unique.append(location)
        features.append(
            LossyFeature(
                code=code,
                message=str(entry["message"]),
                count=int(entry["count"]),
                record_types=tuple(sorted(types)),
                locations=tuple(unique),
            )
        )
    return tuple(features)
