from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
import tempfile
import re
from xml.sax.saxutils import escape
from xml.etree import ElementTree
import zipfile

from .errors import InvalidPpt
from .ppt import BasicShape, HeaderFooter, Picture, Presentation, Table, TextBox, TextRun

CT = "http://schemas.openxmlformats.org/package/2006/content-types"

def _xml(value: str) -> str: return escape(value, {'"': '&quot;'})
def _element(prefix: str, name: str, value: str | None, attributes: str = "") -> str:
    return "" if value is None else f"<{prefix}:{name}{attributes}>{_xml(value)}</{prefix}:{name}>"

def _paragraphs(box: TextBox, hyperlink_ids: dict[str, str]) -> str:
    paragraphs: list[list[tuple[str, TextRun]]] = [[]]
    for run in box.runs or (TextRun(box.text),):
        pieces = re.split(r"([\r\n])", run.text.replace("\v", "\n"))
        for piece in pieces:
            if piece in ("\r", "\n"):
                paragraphs.append([])
            elif piece:
                paragraphs[-1].append((piece, run))
    result: list[str] = []
    for index, fragments in enumerate(paragraphs):
        fragments = list(fragments)
        alignment = box.paragraph_alignments[index] if index < len(box.paragraph_alignments) else None
        bullet = box.paragraph_bullets[index] if index < len(box.paragraph_bullets) else False
        level = box.paragraph_levels[index] if index < len(box.paragraph_levels) else 0
        leading_tabs = 0
        for fragment_index, (value, style) in enumerate(fragments):
            count = len(value) - len(value.lstrip("\t"))
            leading_tabs += count
            if count:
                fragments[fragment_index] = (value[count:], style)
            if fragments[fragment_index][0]:
                break
        attributes = f' algn="{alignment}"' if alignment else ""
        if level:
            attributes += f' lvl="{level}"'
        if box.default_tab_size is not None and box.default_tab_size > 0:
            attributes += f' defTabSz="{_emu(box.default_tab_size)}"'
        left_margin = (
            box.paragraph_left_margins[index]
            if index < len(box.paragraph_left_margins) else None
        )
        indent = (
            box.paragraph_indents[index]
            if index < len(box.paragraph_indents) else None
        )
        # LibreOffice applies leading tabs after paragraph centering/autofit,
        # which shifts or wraps lines that legacy PowerPoint kept aligned.
        # Express a leading tab as a paragraph margin for left-aligned text;
        # centered text already has the intended placement.
        if (
            leading_tabs
            and alignment != "ctr"
            and box.tab_stops
            and left_margin in (None, 0)
        ):
            left_margin = box.tab_stops[
                min(leading_tabs - 1, len(box.tab_stops) - 1)
            ][0]
            indent = left_margin
        if left_margin is not None:
            attributes += f' marL="{_emu(left_margin)}"'
            attributes += f' indent="{_emu((indent or 0) - left_margin)}"'
        elif indent is not None:
            attributes += f' indent="{_emu(indent)}"'
        if bullet:
            if left_margin is None and indent is None and box.is_placeholder:
                margin = _emu(216 + level * 252)
                generic_indent = _emu(216 if level == 0 else 180)
                attributes += f' marL="{margin}" indent="-{generic_indent}"'
            bullet_char = box.paragraph_bullet_chars[index] if index < len(box.paragraph_bullet_chars) else None
            bullet_xml = f'<a:buChar char="{_xml(bullet_char or "•")}"/>'
        else:
            bullet_xml = '<a:buNone/>'
        tabs_xml = ""
        if box.tab_stops:
            tabs_xml = "<a:tabLst>" + "".join(
                f'<a:tab pos="{_emu(position)}" algn="{alignment}"/>'
                for position, alignment in box.tab_stops
                if position >= 0
            ) + "</a:tabLst>"
        spacing_xml = ""
        for tag, values in (
            ("lnSpc", box.paragraph_line_spacings),
            ("spcBef", box.paragraph_space_before),
            ("spcAft", box.paragraph_space_after),
        ):
            spacing = values[index] if index < len(values) else None
            if spacing is None:
                continue
            if spacing >= 0:
                value_xml = f'<a:spcPct val="{spacing * 1000}"/>'
            else:
                value_xml = f'<a:spcPts val="{round(-spacing * 12.5)}"/>'
            spacing_xml += f"<a:{tag}>{value_xml}</a:{tag}>"
        ppr = f'<a:pPr{attributes}>{spacing_xml}{bullet_xml}{tabs_xml}</a:pPr>'
        runs: list[str] = []
        for value, style in fragments:
            attrs = ['lang="en-US"']
            if style.bold is not None: attrs.append(f'b="{1 if style.bold else 0}"')
            if style.italic is not None: attrs.append(f'i="{1 if style.italic else 0}"')
            if style.underline is not None: attrs.append(f'u="{"sng" if style.underline else "none"}"')
            if style.font_size: attrs.append(f'sz="{style.font_size * 100}"')
            if style.baseline is not None: attrs.append(f'baseline="{style.baseline * 1000}"')
            color = f'<a:solidFill><a:srgbClr val="{style.color}"/></a:solidFill>' if style.color else ''
            typeface = f'<a:latin typeface="{_xml(style.typeface)}"/><a:ea typeface="{_xml(style.typeface)}"/>' if style.typeface else ''
            hyperlink = f'<a:hlinkClick r:id="{hyperlink_ids[style.hyperlink]}"/>' if style.hyperlink in hyperlink_ids else ''
            runs.append(f'<a:r><a:rPr {" ".join(attrs)}>{color}{typeface}{hyperlink}</a:rPr><a:t xml:space="preserve">{_xml(value)}</a:t></a:r>')
        result.append(f'<a:p>{ppr}{"".join(runs)}</a:p>')
    return "".join(result)
def _emu(value: int) -> int:
    return round(value * 914400 / 576)

def _xfrm_attributes(rotation: int, flip_horizontal: bool, flip_vertical: bool) -> str:
    values = []
    if rotation: values.append(f'rot="{rotation}"')
    if flip_horizontal: values.append('flipH="1"')
    if flip_vertical: values.append('flipV="1"')
    return (" " + " ".join(values)) if values else ""

def _xfrm_box(
    left: int, top: int, width: int, height: int, rotation: int
) -> tuple[int, int, int, int]:
    normalized = rotation % 21600000
    if normalized in (5400000, 16200000):
        return (
            round(left + (width - height) / 2),
            round(top + (height - width) / 2),
            height,
            width,
        )
    return left, top, width, height

def _fill_xml(
    color: str | None,
    pattern: str | None = None,
    back_color: str | None = None,
) -> str:
    if color and pattern and back_color:
        return (
            f'<a:pattFill prst="{pattern}">'
            f'<a:fgClr><a:srgbClr val="{color}"/></a:fgClr>'
            f'<a:bgClr><a:srgbClr val="{back_color}"/></a:bgClr>'
            f'</a:pattFill>'
        )
    return f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>' if color else '<a:noFill/>'

def _line_xml(
    color: str | None,
    dash: str | None = None,
    width: int | None = None,
    head: tuple[str, str | None, str | None] | None = None,
    tail: tuple[str, str | None, str | None] | None = None,
) -> str:
    width_attr = f' w="{width}"' if width is not None else ""
    if not color:
        return f'<a:ln{width_attr}><a:noFill/></a:ln>'
    dash_xml = f'<a:prstDash val="{dash}"/>' if dash else ''
    ends = ""
    for tag, value in (("headEnd", head), ("tailEnd", tail)):
        if value is None:
            continue
        kind, arrow_width, arrow_length = value
        attributes = f' type="{kind}"'
        if arrow_width:
            attributes += f' w="{arrow_width}"'
        if arrow_length:
            attributes += f' len="{arrow_length}"'
        ends += f"<a:{tag}{attributes}/>"
    return f'<a:ln{width_attr}><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>{dash_xml}{ends}</a:ln>'

def _path_xml(shape: BasicShape) -> str:
    if not shape.path:
        adjustments = "".join(
            f'<a:gd name="adj{index}" fmla="val {value}"/>'
            for index, value in enumerate(shape.adjustments, 1)
        )
        return (
            f'<a:prstGeom prst="{shape.preset}">'
            f'<a:avLst>{adjustments}</a:avLst></a:prstGeom>'
        )
    commands: list[str] = []
    for item in shape.path:
        kind = item[0]
        if kind == "M":
            x, y = item[1]
            commands.append(f'<a:moveTo><a:pt x="{x}" y="{y}"/></a:moveTo>')
        elif kind == "L":
            x, y = item[1]
            commands.append(f'<a:lnTo><a:pt x="{x}" y="{y}"/></a:lnTo>')
        elif kind == "C":
            (x1, y1), (x2, y2), (x3, y3) = item[1], item[2], item[3]
            commands.append(
                f'<a:cubicBezTo><a:pt x="{x1}" y="{y1}"/><a:pt x="{x2}" y="{y2}"/><a:pt x="{x3}" y="{y3}"/></a:cubicBezTo>'
            )
        elif kind == "Z":
            commands.append('<a:close/>')
    fill_attribute = ' fill="none"' if shape.fill_color is None else ""
    return (
        f'<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>'
        f'<a:rect l="l" t="t" r="r" b="b"/>'
        f'<a:pathLst><a:path w="{shape.path_width}" h="{shape.path_height}"{fill_attribute}>'
        f'{"".join(commands)}</a:path></a:pathLst></a:custGeom>'
    )

def _field_shape(shape_id: int, name: str, left: int, top: int, width: int, height: int,
                 value: str, alignment: str, field_type: str | None = None) -> str:
    if field_type:
        field_id = f"{{00000000-0000-0000-0000-{shape_id:012d}}}"
        run = f'<a:fld id="{field_id}" type="{field_type}"><a:rPr lang="en-US" sz="1000"/><a:t>{_xml(value)}</a:t></a:fld>'
    else:
        run = f'<a:r><a:rPr lang="en-US" sz="1000"/><a:t xml:space="preserve">{_xml(value)}</a:t></a:r>'
    return f'<p:sp><p:nvSpPr><p:cNvPr id="{shape_id}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{_emu(left)}" y="{_emu(top)}"/><a:ext cx="{_emu(width)}" cy="{_emu(height)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:pPr algn="{alignment}"><a:buNone/></a:pPr>{run}</a:p></p:txBody></p:sp>'

def _header_footer_shapes(value: HeaderFooter | None, slide_width: int, slide_height: int,
                          slide_number: int, first_id: int) -> list[str]:
    if value is None:
        return []
    result: list[str] = []
    if value.header_text:
        result.append(_field_shape(first_id + len(result), "Header", 288, 96, slide_width - 576, 240,
                                   value.header_text, "ctr"))
    if value.date_is_auto or value.date_text:
        result.append(_field_shape(first_id + len(result), "Date", 288, slide_height - 336, 1440, 240,
                                   value.date_text or datetime.now().strftime("%m/%d/%Y"), "l",
                                   "datetime1" if value.date_is_auto else None))
    if value.footer_text:
        result.append(_field_shape(first_id + len(result), "Footer", 1728, slide_height - 336,
                                   max(576, slide_width - 3456), 240, value.footer_text, "ctr"))
    if value.show_slide_number:
        result.append(_field_shape(first_id + len(result), "Slide Number", slide_width - 864,
                                   slide_height - 336, 576, 240, str(slide_number), "r", "slidenum"))
    return result

# Built-in Office table style (Medium Style 2 - Accent 1). Referencing a
# built-in GUID lets PowerPoint resolve the style without a custom tableStyles
# part while still drawing clean borders around the converted cells.
_TABLE_STYLE_ID = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"

def _table_frames(tables: tuple[Table, ...], start_id: int) -> tuple[list[str], int]:
    frames: list[str] = []
    next_id = start_id
    for table in tables:
        rows, cols = table.rows, table.cols
        col_widths = [
            max(cell.width for cell in table.cells if cell.col == c) for c in range(cols)
        ]
        row_heights = [
            max(cell.height for cell in table.cells if cell.row == r) for r in range(rows)
        ]
        grid = "".join(f'<a:gridCol w="{_emu(w)}"/>' for w in col_widths)
        rows_xml = ""
        for r in range(rows):
            cells_xml = ""
            for c in range(cols):
                cell = next(x for x in table.cells if x.row == r and x.col == c)
                textbox = TextBox(text=cell.text, runs=cell.runs, left=cell.left, top=cell.top, width=cell.width, height=cell.height, fill_color=cell.fill_color, wrap_text=True)
                body = _paragraphs(textbox, {})
                tcpr = f'<a:tcPr marL="45720" marR="45720" marT="45720" marB="45720">{_fill_xml(cell.fill_color)}</a:tcPr>'
                cells_xml += f'<a:tc><a:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/>{body}</a:txBody>{tcpr}</a:tc>'
            rows_xml += f'<a:tr h="{_emu(row_heights[r])}">{cells_xml}</a:tr>'
        tbl_borders = (
            '<a:tblBorders>'
            '<a:top w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:top>'
            '<a:left w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:left>'
            '<a:bottom w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:bottom>'
            '<a:right w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:right>'
            '<a:insideH w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:insideH>'
            '<a:insideV w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:insideV>'
            '</a:tblBorders>'
        )
        tbl = (
            f'<a:tbl><a:tblPr firstRow="0" bandRow="0">'
            f'<a:tableStyleId>{_TABLE_STYLE_ID}</a:tableStyleId>{tbl_borders}</a:tblPr>'
            f'<a:tblGrid>{grid}</a:tblGrid>{rows_xml}</a:tbl>'
        )
        frame = (
            f'<p:graphicFrame><p:nvGraphicFramePr>'
            f'<p:cNvPr id="{next_id}" name="Table {next_id}"/>'
            f'<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>'
            f'<p:nvPr/></p:nvGraphicFramePr>'
            f'<p:xfrm><a:off x="{_emu(table.left)}" y="{_emu(table.top)}"/>'
            f'<a:ext cx="{_emu(table.width)}" cy="{_emu(table.height)}"/></p:xfrm>'
            f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
            f'{tbl}</a:graphicData></a:graphic></p:graphicFrame>'
        )
        frames.append(frame)
        next_id += 1
    return frames, next_id

def _slide(parts: tuple[TextBox, ...], pictures: list[tuple[Picture, str]], basic_shapes: tuple[BasicShape, ...], background_color: str | None, background_color_end: str | None, background_gradient_angle: int | None, background_gradient_type: int | None, hyperlink_ids: dict[str, str], header_footer: HeaderFooter | None, slide_width: int, slide_height: int, slide_number: int, tables: tuple[Table, ...], hidden: bool) -> str:
    background_drawing_shapes = []
    foreground_drawing_shapes = []
    for index, shape in enumerate(basic_shapes, 2):
        fill = _fill_xml(shape.fill_color, shape.fill_pattern, shape.fill_back_color)
        line = _line_xml(
            shape.line_color, shape.line_dash, shape.line_width,
            shape.line_head, shape.line_tail,
        )
        geom = _path_xml(shape)
        left, top, width, height = _xfrm_box(
            shape.left, shape.top, shape.width, shape.height, shape.rotation
        )
        shape_xml = f'<p:sp><p:nvSpPr><p:cNvPr id="{index}" name="{shape.preset} {index}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm{_xfrm_attributes(shape.rotation, shape.flip_horizontal, shape.flip_vertical)}><a:off x="{_emu(left)}" y="{_emu(top)}"/><a:ext cx="{_emu(width)}" cy="{_emu(height)}"/></a:xfrm>{geom}{fill}{line}</p:spPr></p:sp>'
        target = (
            background_drawing_shapes
            if shape.fill_color and shape.width * shape.height >= 1_000_000
            else foreground_drawing_shapes
        )
        target.append(shape_xml)
    text_shapes = []
    for index, box in enumerate(parts, len(basic_shapes) + 2):
        paragraphs = _paragraphs(box, hyperlink_ids)
        fill = _fill_xml(box.fill_color, box.fill_pattern, box.fill_back_color)
        line = _line_xml(
            box.line_color, box.line_dash, box.line_width,
            box.line_head, box.line_tail,
        )
        anchor = f' anchor="{box.vertical_anchor}"' if box.vertical_anchor else ''
        autofit = (
            '<a:spAutoFit/>' if box.fit_shape_to_text
            else '<a:normAutofit/>' if box.auto_fit
            else ''
        )
        wrap = "square" if box.wrap_text else "none"
        insets = "".join(
            f' {name}="{value}"'
            for name, value in (
                ("lIns", box.inset_left),
                ("tIns", box.inset_top),
                ("rIns", box.inset_right),
                ("bIns", box.inset_bottom),
            )
            if value is not None
        )
        left, top, width, height = _xfrm_box(
            box.left, box.top, box.width, box.height, box.rotation
        )
        text_shapes.append(f'<p:sp><p:nvSpPr><p:cNvPr id="{index}" name="Text Box {index-1}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm{_xfrm_attributes(box.rotation, box.flip_horizontal, box.flip_vertical)}><a:off x="{_emu(left)}" y="{_emu(top)}"/><a:ext cx="{_emu(width)}" cy="{_emu(height)}"/></a:xfrm><a:prstGeom prst="{box.preset}"><a:avLst/></a:prstGeom>{fill}{line}</p:spPr><p:txBody><a:bodyPr wrap="{wrap}"{anchor}{insets}>{autofit}</a:bodyPr><a:lstStyle/>{paragraphs}</p:txBody></p:sp>')
    footer_shapes = _header_footer_shapes(header_footer, slide_width, slide_height, slide_number,
                                          len(basic_shapes) + len(parts) + 2)
    base_picture_shapes = []
    overlay_picture_shapes = []
    for index, (picture, relation_id) in enumerate(pictures, len(basic_shapes) + len(parts) + len(footer_shapes) + 2):
        crop = f'<a:srcRect l="{picture.crop_left}" t="{picture.crop_top}" r="{picture.crop_right}" b="{picture.crop_bottom}"/>' if any((picture.crop_left, picture.crop_top, picture.crop_right, picture.crop_bottom)) else ''
        stretch = (
            '<a:stretch/>'
            if picture.extension in ("wmf", "emf")
            else '<a:stretch><a:fillRect/></a:stretch>'
        )
        transparent = (
            f'<a:clrChange useA="1"><a:clrFrom><a:srgbClr val="{picture.transparent_color}"/>'
            f'</a:clrFrom><a:clrTo><a:srgbClr val="{picture.transparent_color}">'
            f'<a:alpha val="0"/></a:srgbClr></a:clrTo></a:clrChange>'
            if picture.transparent_color else ""
        )
        left, top, width, height = _xfrm_box(
            picture.left, picture.top, picture.width, picture.height,
            picture.rotation
        )
        line = _line_xml(
            picture.line_color, picture.line_dash, picture.line_width
        )
        picture_xml = f'<p:pic><p:nvPicPr><p:cNvPr id="{index}" name="Picture {index}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="{relation_id}">{transparent}</a:blip>{crop}{stretch}</p:blipFill><p:spPr><a:xfrm{_xfrm_attributes(picture.rotation, picture.flip_horizontal, picture.flip_vertical)}><a:off x="{_emu(left)}" y="{_emu(top)}"/><a:ext cx="{_emu(width)}" cy="{_emu(height)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>{line}</p:spPr></p:pic>'
        # Large pictures are normally screenshots, plots, or photo backdrops
        # that annotations must overlay.  Small pictures are commonly WMF
        # equations or clip art and need to remain above filled diagram boxes.
        target = (
            base_picture_shapes
            if picture.width * picture.height >= 1_000_000
            else overlay_picture_shapes
        )
        target.append(picture_xml)
    if background_color and background_color_end:
        stops = (
            f'<a:gs pos="0"><a:srgbClr val="{background_color}"/></a:gs>'
            f'<a:gs pos="50000"><a:srgbClr val="{background_color_end}"/></a:gs>'
            f'<a:gs pos="100000"><a:srgbClr val="{background_color}"/></a:gs>'
            if background_gradient_type == 7
            else
            f'<a:gs pos="0"><a:srgbClr val="{background_color}"/></a:gs>'
            f'<a:gs pos="100000"><a:srgbClr val="{background_color_end}"/></a:gs>'
        )
        angle = (
            background_gradient_angle
            if background_gradient_angle is not None else 2700000
        )
        background = f'<p:bg><p:bgPr><a:gradFill rotWithShape="0"><a:gsLst>{stops}</a:gsLst><a:lin ang="{angle}" scaled="1"/></a:gradFill><a:effectLst/></p:bgPr></p:bg>'
    else:
        background = f'<p:bg><p:bgPr><a:solidFill><a:srgbClr val="{background_color}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>' if background_color else ''
    show = ' show="0"' if hidden else ''
    table_next_id = len(basic_shapes) + len(parts) + len(footer_shapes) + len(pictures) + 2
    table_frames, _ = _table_frames(tables, table_next_id)
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"' + show + '><p:cSld>' + background + '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' + ''.join(background_drawing_shapes) + ''.join(base_picture_shapes) + ''.join(foreground_drawing_shapes) + ''.join(overlay_picture_shapes) + ''.join(table_frames) + ''.join(text_shapes) + ''.join(footer_shapes) + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'

def _notes_slide(values: tuple[str, ...]) -> str:
    text = "\r".join(values)
    paragraphs = _paragraphs(TextBox(text, 0, 0, 0, 0, (TextRun(text, font_size=12),)), {})
    group_transform = '<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>'
    tree = '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>' + group_transform + '</p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image"/><p:cNvSpPr/><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="1143000" y="685800"/><a:ext cx="4572000" cy="3429000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Body"/><p:cNvSpPr txBox="1"/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="4343400"/><a:ext cx="5486400" cy="4114800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>' + paragraphs + '</p:txBody></p:sp></p:spTree>'
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld>' + tree + '</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>'

def _validate_package(path: Path) -> None:
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        required = {'[Content_Types].xml', '_rels/.rels', 'ppt/presentation.xml',
                    'ppt/slideMasters/slideMaster1.xml', 'ppt/slideLayouts/slideLayout1.xml',
                    'ppt/theme/theme1.xml'}
        missing = required - names
        if missing:
            raise InvalidPpt(f"generated package is missing {sorted(missing)[0]}")
        for name in names:
            if name.endswith(('.xml', '.rels')):
                try:
                    ElementTree.fromstring(archive.read(name))
                except ElementTree.ParseError as exc:
                    raise InvalidPpt(f"generated XML is malformed: {name}") from exc

def write_pptx(destination: str | Path, presentation: Presentation) -> None:
    slides = presentation.slides
    has_notes = any(slide.notes for slide in slides)
    author_ids: dict[tuple[str, str], int] = {}
    for slide in slides:
        for comment in slide.comments:
            author_ids.setdefault((comment.author, comment.initials), len(author_ids))
    target = Path(destination); target.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(dir=target.parent, prefix=f".{target.name}.", suffix=".tmp", delete=False) as temp:
        temporary = Path(temp.name)
    try:
        now = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace('+00:00', 'Z')
        with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as archive:
            overrides = ''.join(f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' for i in range(1, len(slides)+1))
            comment_overrides = ''.join(f'<Override PartName="/ppt/comments/comment{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.comments+xml"/>' for i, slide in enumerate(slides, 1) if slide.comments)
            note_overrides = ''.join(f'<Override PartName="/ppt/notesSlides/notesSlide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>' for i, slide in enumerate(slides, 1) if slide.notes)
            if has_notes:
                note_overrides += (
                    '<Override PartName="/ppt/notesMasters/notesMaster1.xml" '
                    'ContentType="application/vnd.openxmlformats-officedocument.'
                    'presentationml.notesMaster+xml"/>'
                    '<Override PartName="/ppt/theme/theme2.xml" '
                    'ContentType="application/vnd.openxmlformats-officedocument.'
                    'theme+xml"/>'
                )
            if author_ids:
                comment_overrides += '<Override PartName="/ppt/commentAuthors.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"/>'
            archive.writestr('[Content_Types].xml', f'<?xml version="1.0" encoding="UTF-8"?><Types xmlns="{CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/><Default Extension="tif" ContentType="image/tiff"/><Default Extension="emf" ContentType="image/x-emf"/><Default Extension="wmf" ContentType="image/x-wmf"/><Default Extension="pct" ContentType="image/x-pict"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>{overrides}{comment_overrides}{note_overrides}</Types>')
            archive.writestr('_rels/.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>')
            core = presentation.core_properties
            created = core.created or now
            core_values = ''.join((
                _element('dc', 'title', core.title), _element('dc', 'subject', core.subject),
                _element('dc', 'creator', core.creator), _element('cp', 'keywords', core.keywords),
                _element('dc', 'description', core.description),
                _element('cp', 'lastModifiedBy', core.last_modified_by),
                _element('cp', 'revision', core.revision),
                _element('dcterms', 'created', created, ' xsi:type="dcterms:W3CDTF"'),
                _element('dcterms', 'modified', core.modified, ' xsi:type="dcterms:W3CDTF"'),
                _element('cp', 'lastPrinted', core.last_printed),
            ))
            archive.writestr('docProps/core.xml', f'<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">{core_values}</cp:coreProperties>')
            archive.writestr('docProps/app.xml', f'<?xml version="1.0" encoding="UTF-8"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>ppt2pptx</Application><Slides>{len(slides)}</Slides></Properties>')
            rels, ids = [], []
            image_index = 1
            for i in range(1, len(slides)+1):
                rels.append(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>')
                ids.append(f'<p:sldId id="{255+i}" r:id="rId{i}"/>')
                picture_refs: list[tuple[Picture, str]] = []
                picture_rels: list[str] = []
                next_relation = 2
                for picture in slides[i-1].pictures:
                    relation_id = f"rId{next_relation}"
                    next_relation += 1
                    filename = f"image{image_index}.{picture.extension}"
                    picture_refs.append((picture, relation_id))
                    picture_rels.append(f'<Relationship Id="{relation_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/{filename}"/>')
                    archive.writestr(f'ppt/media/{filename}', picture.data)
                    image_index += 1
                if slides[i-1].comments:
                    comment_rid = f"rId{next_relation}"
                    next_relation += 1
                    picture_rels.append(f'<Relationship Id="{comment_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment{i}.xml"/>')
                    comments_xml = []
                    for comment_index, comment in enumerate(slides[i-1].comments, 1):
                        author_id = author_ids[(comment.author, comment.initials)]
                        date = f' dt="{_xml(comment.created)}"' if comment.created else ""
                        comments_xml.append(f'<p:cm authorId="{author_id}"{date} idx="{comment_index}"><p:pos x="{_emu(comment.left)}" y="{_emu(comment.top)}"/><p:text>{_xml(comment.text)}</p:text></p:cm>')
                    archive.writestr(f'ppt/comments/comment{i}.xml', '<?xml version="1.0" encoding="UTF-8"?><p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' + ''.join(comments_xml) + '</p:cmLst>')
                hyperlink_ids: dict[str, str] = {}
                for box in slides[i-1].text_boxes:
                    for run in box.runs:
                        if run.hyperlink and run.hyperlink not in hyperlink_ids:
                            relation_id = f"rId{next_relation}"
                            next_relation += 1
                            hyperlink_ids[run.hyperlink] = relation_id
                            picture_rels.append(f'<Relationship Id="{relation_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{_xml(run.hyperlink)}" TargetMode="External"/>')
                if slides[i-1].notes:
                    notes_relation = f"rId{next_relation}"
                    picture_rels.append(f'<Relationship Id="{notes_relation}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide{i}.xml"/>')
                    archive.writestr(f'ppt/notesSlides/notesSlide{i}.xml', _notes_slide(slides[i-1].notes))
                    archive.writestr(f'ppt/notesSlides/_rels/notesSlide{i}.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide' + str(i) + '.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMasters/notesMaster1.xml"/></Relationships>')
                archive.writestr(f'ppt/slides/slide{i}.xml', _slide(slides[i-1].text_boxes, picture_refs, slides[i-1].shapes, slides[i-1].background_color, slides[i-1].background_color_end, slides[i-1].background_gradient_angle, slides[i-1].background_gradient_type, hyperlink_ids, slides[i-1].header_footer, presentation.width, presentation.height, i, slides[i-1].tables, slides[i-1].hidden))
                archive.writestr(f'ppt/slides/_rels/slide{i}.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' + ''.join(picture_rels) + '</Relationships>')
            master_rid = len(slides) + 1
            rels.append(f'<Relationship Id="rId{master_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>')
            notes_master_rid = master_rid + 1 if has_notes else None
            if notes_master_rid is not None:
                rels.append(f'<Relationship Id="rId{notes_master_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="notesMasters/notesMaster1.xml"/>')
            if author_ids:
                author_rid = master_rid + (2 if has_notes else 1)
                rels.append(f'<Relationship Id="rId{author_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors" Target="commentAuthors.xml"/>')
                authors_xml = ''.join(f'<p:cmAuthor id="{author_id}" name="{_xml(author)}" initials="{_xml(initials)}" lastIdx="{sum(len(slide.comments) for slide in slides)}" clrIdx="{author_id % 8}"/>' for (author, initials), author_id in author_ids.items())
                archive.writestr('ppt/commentAuthors.xml', '<?xml version="1.0" encoding="UTF-8"?><p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' + authors_xml + '</p:cmAuthorLst>')
            archive.writestr('ppt/_rels/presentation.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + ''.join(rels) + '</Relationships>')
            notes_master_ids = f'<p:notesMasterIdLst><p:notesMasterId r:id="rId{notes_master_rid}"/></p:notesMasterIdLst>' if notes_master_rid is not None else ''
            archive.writestr('ppt/presentation.xml', f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId{master_rid}"/></p:sldMasterIdLst>{notes_master_ids}<p:sldIdLst>{"".join(ids)}</p:sldIdLst><p:sldSz cx="{_emu(presentation.width)}" cy="{_emu(presentation.height)}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>')
            # PowerPoint rejects empty grpSpPr / empty txStyles / incomplete fmtScheme lists.
            grp_xfrm = '<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>'
            sp_tree = f'<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>{grp_xfrm}</p:grpSpPr></p:spTree>'
            master_level = (
                '<a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">'
                '<a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
                '<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>'
            )
            tx_styles = (
                f'<p:txStyles><p:titleStyle>{master_level}</p:titleStyle>'
                f'<p:bodyStyle>{master_level}</p:bodyStyle>'
                f'<p:otherStyle>{master_level}</p:otherStyle></p:txStyles>'
            )
            if has_notes:
                notes_style = f'<p:notesStyle>{master_level}</p:notesStyle>'
                archive.writestr('ppt/notesMasters/notesMaster1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:notesMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld>' + sp_tree + '</p:cSld><p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/><p:hf hdr="1" ftr="1" dt="1" sldNum="1"/>' + notes_style + '</p:notesMaster>')
                archive.writestr('ppt/notesMasters/_rels/notesMaster1.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme2.xml"/></Relationships>')
            archive.writestr('ppt/slideMasters/slideMaster1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld>' + sp_tree + '</p:cSld><p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>' + tx_styles + '</p:sldMaster>')
            archive.writestr('ppt/slideMasters/_rels/slideMaster1.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>')
            archive.writestr('ppt/slideLayouts/slideLayout1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1"><p:cSld name="Blank">' + sp_tree + '</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>')
            archive.writestr('ppt/slideLayouts/_rels/slideLayout1.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>')
            theme_fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            theme_line = '<a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>'
            theme_effect = '<a:effectStyle><a:effectLst/></a:effectStyle>'
            theme_xml = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="ppt2pptx">'
                '<a:themeElements><a:clrScheme name="Office">'
                '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
                '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
                '<a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2>'
                '<a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2>'
                '<a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4>'
                '<a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6>'
                '<a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink>'
                '</a:clrScheme><a:fontScheme name="Office">'
                '<a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
                '<a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>'
                '</a:fontScheme><a:fmtScheme name="Office">'
                f'<a:fillStyleLst>{theme_fill * 3}</a:fillStyleLst>'
                f'<a:lnStyleLst>{theme_line * 3}</a:lnStyleLst>'
                f'<a:effectStyleLst>{theme_effect * 3}</a:effectStyleLst>'
                f'<a:bgFillStyleLst>{theme_fill * 3}</a:bgFillStyleLst>'
                '</a:fmtScheme></a:themeElements></a:theme>'
            )
            archive.writestr('ppt/theme/theme1.xml', theme_xml)
            if has_notes:
                archive.writestr('ppt/theme/theme2.xml', theme_xml)
        _validate_package(temporary)
        temporary.replace(target)
    except Exception:
        temporary.unlink(missing_ok=True); raise
