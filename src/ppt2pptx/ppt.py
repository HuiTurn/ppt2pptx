"""Bounded parser for the portions of MS-PPT needed for editable slide text."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
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
RT_OFFICEART_SP_CONTAINER = 0xF004
RT_OFFICEART_CLIENT_ANCHOR = 0xF010
RT_OFFICEART_CLIENT_TEXTBOX = 0xF00D
RT_OFFICEART_BSE = 0xF007
RT_OFFICEART_FOPT = 0xF00B
CONTAINER_VERSION = 0xF
DEFAULT_SLIDE_WIDTH = 5760
DEFAULT_SLIDE_HEIGHT = 4320

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

@dataclass(frozen=True, slots=True)
class TextContent:
    text: str
    runs: tuple[TextRun, ...] = ()
    paragraph_alignments: tuple[str | None, ...] = ()
    paragraph_bullets: tuple[bool, ...] = ()

@dataclass(frozen=True, slots=True)
class _MasterStyle:
    run: TextRun
    alignment: str | None
    bullet: bool | None

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

def records(data: bytes, start: int = 0, end: int | None = None):
    end = len(data) if end is None else end
    cursor = start
    while cursor < end:
        if end - cursor < 8: raise InvalidPpt("PowerPoint record header is truncated")
        version_instance, record_type, length = struct.unpack_from("<HHI", data, cursor)
        payload_start, payload_end = cursor + 8, cursor + 8 + length
        if payload_end > end: raise InvalidPpt("PowerPoint record extends beyond its container")
        yield Record(cursor, version_instance & 0xF, version_instance >> 4, record_type, data[payload_start:payload_end])
        cursor = payload_end

def descendants(record: Record):
    if record.version != CONTAINER_VERSION: return
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

def _slide_refs(document: Record) -> list[int]:
    refs: list[int] = []
    for container in descendants(document):
        if container.type != RT_SLIDE_LIST_WITH_TEXT or container.instance != 0:
            continue
        for record in records(container.payload):
            if record.type != RT_SLIDE_PERSIST_ATOM or len(record.payload) < 12:
                continue
            # SlidePersistAtom begins with persistIdRef.  The following fields
            # are slide identifier and flags.
            reference = struct.unpack_from("<I", record.payload, 0)[0]
            if reference:
                refs.append(reference)
    return refs

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
        return record.payload.decode("utf-16le", "replace").rstrip("\x00")
    if record.type == RT_TEXT_BYTES_ATOM:
        return record.payload.decode("cp1252", "replace").rstrip("\x00")
    return None

def _skip_paragraph_properties(payload: bytes, position: int, mask: int) -> tuple[int, str | None, bool | None]:
    alignment: str | None = None
    bullet: bool | None = None
    fixed = ((0xF, 2), (0x80, 2), (0x10, 2), (0x40, 2), (0x20, 4),
             (0x800, 2), (0x1000, 2), (0x2000, 2), (0x4000, 2),
             (0x100, 2), (0x400, 2), (0x8000, 2))
    for property_mask, size in fixed:
        if mask & property_mask:
            if position + size > len(payload):
                return len(payload), alignment, bullet
            value = int.from_bytes(payload[position:position + size], "little", signed=size == 2)
            if property_mask == 0xF:
                bullet = bool(value & 1)
            elif property_mask == 0x800:
                alignment = {0: "l", 1: "ctr", 2: "r", 3: "just", 4: "dist"}.get(value)
            position += size
    if mask & 0x100000:
        if position + 2 > len(payload):
            return len(payload), alignment, bullet
        count = struct.unpack_from("<H", payload, position)[0]
        position += min(2 + count * 4, len(payload) - position)
    for property_mask in (0x10000, 0xE0000, 0x200000):
        if mask & property_mask:
            position = min(position + 2, len(payload))
    return position, alignment, bullet

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
    if mask & 0x80000:
        position = min(position + 2, len(payload))
    return position, TextRun("", bold, italic, underline, font_size, color,
                             fonts[font_index] if font_index is not None and font_index < len(fonts) else None)

def _merge_style(value: str, explicit: TextRun, master: TextRun | None) -> TextRun:
    if master is None:
        return TextRun(value, explicit.bold, explicit.italic, explicit.underline,
                       explicit.font_size, explicit.color, explicit.typeface, explicit.hyperlink)
    return TextRun(value,
                   explicit.bold if explicit.bold is not None else master.bold,
                   explicit.italic if explicit.italic is not None else master.italic,
                   explicit.underline if explicit.underline is not None else master.underline,
                   explicit.font_size or master.font_size,
                   explicit.color or master.color,
                   explicit.typeface or master.typeface,
                   explicit.hyperlink)

def _style_text(text: str, payload: bytes, fonts: tuple[str, ...], master: _MasterStyle | None, scheme: tuple[str, ...]) -> TextContent:
    position = handled = 0
    paragraph_runs: list[tuple[int, int, str | None, bool | None]] = []
    while position + 10 <= len(payload) and handled < len(text) + 1:
        count = struct.unpack_from("<I", payload, position)[0]
        _indent = struct.unpack_from("<h", payload, position + 4)[0]
        mask = struct.unpack_from("<I", payload, position + 6)[0]
        position, alignment, bullet = _skip_paragraph_properties(payload, position + 10, mask)
        if not count:
            break
        paragraph_runs.append((handled, handled + count, alignment, bullet))
        handled += count
    character_runs: list[TextRun] = []
    handled = 0
    while position + 8 <= len(payload) and handled < len(text) + 1:
        count, mask = struct.unpack_from("<II", payload, position)
        position += 8
        if not count:
            break
        position, explicit = _character_style(payload, position, mask, fonts, scheme)
        end = min(handled + count, len(text))
        if end > handled:
            character_runs.append(_merge_style(text[handled:end], explicit, master.run if master else None))
        handled += count
    if handled < len(text):
        character_runs.append(_merge_style(text[handled:], TextRun(""), master.run if master else None))
    alignments: list[str | None] = []
    bullets: list[bool] = []
    paragraph_start = 0
    for paragraph in text.split("\r"):
        style = next((run for run in paragraph_runs if run[0] <= paragraph_start < run[1]), None)
        alignments.append(style[2] if style and style[2] is not None else master.alignment if master else None)
        bullets.append(style[3] if style and style[3] is not None else bool(master.bullet) if master else False)
        paragraph_start += len(paragraph) + 1
    return TextContent(text, tuple(character_runs), tuple(alignments), tuple(bullets))

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
            output.append(TextRun(value, run.bold, run.italic, run.underline,
                                  run.font_size, run.color, run.typeface, url))
        offset += len(run.text)
    return TextContent(content.text, tuple(output), content.paragraph_alignments, content.paragraph_bullets)

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
        master = masters.get(text_type, ())
        base = master[0] if master else None
        if style is not None and style.type == 4001:
            content = _style_text(value, style.payload, fonts, base, scheme)
        else:
            paragraphs = value.split("\r")
            content = TextContent(value, (_merge_style(value, TextRun(""), base.run if base else None),),
                                  tuple(base.alignment if base else None for _ in paragraphs),
                                  tuple(bool(base.bullet) if base else False for _ in paragraphs))
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
            position, alignment, bullet = _skip_paragraph_properties(atom.payload, position + 4, paragraph_mask)
            if position + 4 > len(atom.payload):
                break
            character_mask = struct.unpack_from("<I", atom.payload, position)[0]
            position, run = _character_style(atom.payload, position + 4, character_mask, fonts, scheme)
            styles.append(_MasterStyle(run, alignment, bullet))
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

def _shape_text_boxes(slide: Record, external_text: list[TextContent], fonts: tuple[str, ...], masters: dict[int, tuple[_MasterStyle, ...]], hyperlinks: dict[int, str], scheme: tuple[str, ...]) -> list[TextBox]:
    result: list[TextBox] = []
    for shape in descendants(slide):
        if shape.type != RT_OFFICEART_SP_CONTAINER:
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
        anchor = next((child for child in children if child.type == RT_OFFICEART_CLIENT_ANCHOR), None)
        if anchor is not None and len(anchor.payload) >= 8:
            top, left, right, bottom = struct.unpack_from("<4h", anchor.payload)
            width, height = max(1, right - left), max(1, bottom - top)
        else:
            # A deterministic cascade is preferable to overlapping all text
            # when a producer omitted the optional client anchor.
            left, top, width, height = 288, 288 + len(result) * 576, 5184, 432
        fopt = next((child for child in children if child.type == RT_OFFICEART_FOPT), None)
        transform = _transform(children, _fopt_properties(fopt) if fopt else {})
        result.append(TextBox(content.text, left, top, width, height, content.runs,
                              content.paragraph_alignments, content.paragraph_bullets,
                              *transform))
    return result

def _anchor(children: list[Record], fallback_index: int) -> tuple[int, int, int, int]:
    anchor = next((child for child in children if child.type == RT_OFFICEART_CLIENT_ANCHOR), None)
    if anchor is not None and len(anchor.payload) >= 8:
        top, left, right, bottom = struct.unpack_from("<4h", anchor.payload)
        return left, top, max(1, right - left), max(1, bottom - top)
    child_anchor = next((child for child in children if child.type == 0xF00F), None)
    if child_anchor is not None and len(child_anchor.payload) >= 16:
        left, top, right, bottom = struct.unpack_from("<4i", child_anchor.payload)
        # Some producers store child anchors in EMUs even though the enclosing
        # PowerPoint drawing uses master units (576 units/inch).
        if max(abs(left), abs(top), abs(right), abs(bottom)) > 100_000:
            left, top, right, bottom = (round(value * 576 / 914400)
                                        for value in (left, top, right, bottom))
        return left, top, max(1, right - left), max(1, bottom - top)
    return 288, 288 + fallback_index * 576, 5184, 432

def _fopt_properties(record: Record) -> dict[int, int]:
    count = min(record.instance, len(record.payload) // 6)
    properties: dict[int, int] = {}
    for index in range(count):
        opid, value = struct.unpack_from("<HI", record.payload, index * 6)
        properties[opid & 0x3FFF] = value
    return properties

def _transform(children: list[Record], properties: dict[int, int]) -> tuple[int, bool, bool]:
    rotation_raw = properties.get(4, 0)
    rotation_signed = struct.unpack("<i", struct.pack("<I", rotation_raw))[0]
    rotation = round(rotation_signed / 65536 * 60000)
    sp = next((child for child in children if child.type == 0xF00A and len(child.payload) >= 8), None)
    flags = struct.unpack_from("<I", sp.payload, 4)[0] if sp else 0
    return rotation, bool(flags & 0x40), bool(flags & 0x80)

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

def _basic_shapes(slide: Record, image_map: dict[int, tuple[bytes, str, str]], scheme: tuple[str, ...]) -> list[BasicShape]:
    presets = {0: "rect", 1: "rect", 2: "roundRect", 3: "ellipse", 4: "diamond", 19: "arc", 20: "line",
               32: "rightArrow", 33: "leftArrow", 34: "upArrow", 35: "downArrow",
               56: "pentagon", 57: "hexagon", 58: "heptagon", 59: "octagon"}
    result: list[BasicShape] = []
    for shape in descendants(slide):
        if shape.type != RT_OFFICEART_SP_CONTAINER:
            continue
        children = _direct_children(shape)
        sp = next((child for child in children if child.type == 0xF00A and len(child.payload) >= 8), None)
        anchor = next((child for child in children if child.type in (RT_OFFICEART_CLIENT_ANCHOR, 0xF00F)), None)
        if sp is None or anchor is None or len(anchor.payload) < 8:
            continue
        fopt = next((child for child in children if child.type == RT_OFFICEART_FOPT), None)
        properties = _fopt_properties(fopt) if fopt else {}
        if 385 not in properties and 448 not in properties:
            continue
        if properties.get(260, 0) in image_map:
            continue
        preset = presets.get(sp.instance)
        if preset is None:
            continue
        left, top, width, height = _anchor(children, len(result))
        result.append(BasicShape(preset, left, top, width, height,
                                 _office_color(properties.get(385), scheme),
                                 _office_color(properties.get(448), scheme),
                                 *_transform(children, properties)))
    return result

def _background(slide: Record, scheme: tuple[str, ...]) -> tuple[str | None, str | None]:
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
                return (back, color) if properties.get(384, 0) and back else (color, None)
    return None, None

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
                bounds = [max(-32768, min(32767, value)) for value in (left, top, right, bottom)]
                checksum = 0xCDD7 ^ 0x9AC6 ^ bounds[0] ^ bounds[1] ^ bounds[2] ^ bounds[3] ^ 72
                placeable = struct.pack("<IH4hHIH", 0x9AC6CDD7, 0, *bounds, 72, 0, checksum & 0xFFFF)
                result[index] = (placeable + raw, "wmf", "image/x-wmf")
            elif blip.type == 0xF01A:
                result[index] = (raw, "emf", "image/x-emf")
            else:
                result[index] = (bytes(512) + raw, "pct", "image/x-pict")
    return result

def _shape_pictures(slide: Record, image_map: dict[int, tuple[bytes, str, str]]) -> list[Picture]:
    result: list[Picture] = []
    for shape in descendants(slide):
        if shape.type != RT_OFFICEART_SP_CONTAINER:
            continue
        children = _direct_children(shape)
        fopt = next((child for child in children if child.type == RT_OFFICEART_FOPT), None)
        if fopt is None:
            continue
        properties = _fopt_properties(fopt)
        reference = properties.get(260)
        image = image_map.get(reference or 0)
        if image is None:
            continue
        left, top, width, height = _anchor(children, len(result))
        data, extension, content_type = image
        def crop(property_id: int) -> int:
            raw = properties.get(property_id, 0)
            signed = struct.unpack("<i", struct.pack("<I", raw))[0]
            return round(signed / 65536 * 100000)
        result.append(Picture(data, extension, content_type, left, top, width, height,
                              crop(258), crop(256), crop(259), crop(257),
                              *_transform(children, properties)))
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
    boxes = _shape_text_boxes(slide_record, external_text, fonts, masters, hyperlinks, scheme)
    if not boxes:
        texts = [value for child in descendants(slide_record) if (value := _text(child))]
        boxes = [TextBox(text, 288, 288 + i * 576, 5184, 432, (TextRun(text),)) for i, text in enumerate(texts)]
    background, background_end = _background(slide_record, scheme)
    if master_record is not None:
        master_background, master_background_end = _background(master_record, scheme)
        if background is None or (background_end is None and master_background_end == background):
            background, background_end = master_background, master_background_end
    # Master freeform geometry uses group-relative coordinates.  It is safer
    # to omit those decorative shapes than to flatten them at the wrong scale;
    # slide-owned editable shapes are retained below.
    shapes: list[BasicShape] = []
    for shape in _basic_shapes(slide_record, image_map, scheme):
        left, top = max(0, shape.left), max(0, shape.top)
        right = min(slide_width, shape.left + shape.width)
        bottom = min(slide_height, shape.top + shape.height)
        if right > left and bottom > top:
            shapes.append(BasicShape(shape.preset, left, top, right - left, bottom - top,
                                     shape.fill_color, shape.line_color, shape.rotation,
                                     shape.flip_horizontal, shape.flip_vertical))
    slide_show_info = next((child for child in descendants(slide_record)
                            if child.type == RT_SLIDE_SHOW_SLIDE_INFO_ATOM
                            and len(child.payload) >= 12), None)
    hidden = bool(struct.unpack_from("<H", slide_show_info.payload, 10)[0] & 0x0004) if slide_show_info else False
    return Slide(
        text_boxes=tuple(boxes),
        pictures=tuple(_shape_pictures(slide_record, image_map)),
        shapes=tuple(shapes),
        background_color=background,
        comments=tuple(_comments(slide_record)),
        header_footer=_header_footer(slide_record, header_footer, instance=0),
        background_color_end=background_end,
        notes=notes,
        hidden=hidden,
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
    note_values: list[tuple[str, ...]] = []
    for reference in _notes_refs(document):
        offset = mapping.get(reference)
        if offset is None or offset >= len(powerpoint_document):
            note_values.append(())
            continue
        try:
            note_record = next(records(powerpoint_document, offset))
        except (InvalidPpt, StopIteration):
            note_values.append(())
            continue
        if note_record.type != 1008:
            note_values.append(())
            continue
        values = tuple(box.text for box in _shape_text_boxes(note_record, [], fonts, masters, hyperlinks, scheme)
                       if box.text.strip() not in ("", "*"))
        note_values.append(values)
    width, height = _presentation_size(document)
    slides: list[Slide] = []
    seen_offsets: set[int] = set()
    for slide_index, reference in enumerate(_slide_refs(document)):
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
        notes = note_values[slide_index] if slide_index < len(note_values) else ()
        slides.append(_parse_slide(slide_record, image_map, external_text.get(reference, []), fonts, masters, hyperlinks, header_footer, scheme, master_records[0] if master_records else None, notes, width, height))
    if not slides:
        for slide_record in descendants(document):
            if slide_record.type == RT_SLIDE:
                slides.append(_parse_slide(slide_record, image_map, [], fonts, masters, hyperlinks, header_footer, scheme, master_records[0] if master_records else None, (), width, height))
    return Presentation(width, height, tuple(slides))

def extract_slides(powerpoint_document: bytes) -> list[list[str]]:
    return [[box.text for box in slide.text_boxes] for slide in extract_presentation(powerpoint_document).slides]
