from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
import tempfile
import re
from xml.sax.saxutils import escape
from xml.etree import ElementTree
import zipfile

from .errors import InvalidPpt
from .ppt import BasicShape, HeaderFooter, Picture, Presentation, TextBox, TextRun

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
        alignment = box.paragraph_alignments[index] if index < len(box.paragraph_alignments) else None
        bullet = box.paragraph_bullets[index] if index < len(box.paragraph_bullets) else False
        attributes = f' algn="{alignment}"' if alignment else ""
        bullet_xml = '<a:buChar char="•"/>' if bullet else '<a:buNone/>'
        ppr = f'<a:pPr{attributes}>{bullet_xml}</a:pPr>'
        runs: list[str] = []
        for value, style in fragments:
            attrs = ['lang="en-US"']
            if style.bold is not None: attrs.append(f'b="{1 if style.bold else 0}"')
            if style.italic is not None: attrs.append(f'i="{1 if style.italic else 0}"')
            if style.underline is not None: attrs.append(f'u="{"sng" if style.underline else "none"}"')
            if style.font_size: attrs.append(f'sz="{style.font_size * 100}"')
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

def _slide(parts: tuple[TextBox, ...], pictures: list[tuple[Picture, str]], basic_shapes: tuple[BasicShape, ...], background_color: str | None, background_color_end: str | None, hyperlink_ids: dict[str, str], header_footer: HeaderFooter | None, slide_width: int, slide_height: int, slide_number: int, hidden: bool) -> str:
    drawing_shapes = []
    for index, shape in enumerate(basic_shapes, 2):
        fill = f'<a:solidFill><a:srgbClr val="{shape.fill_color}"/></a:solidFill>' if shape.fill_color else '<a:noFill/>'
        line = (f'<a:ln><a:solidFill><a:srgbClr val="{shape.line_color}"/></a:solidFill></a:ln>'
                if shape.line_color else '<a:ln><a:noFill/></a:ln>')
        drawing_shapes.append(f'<p:sp><p:nvSpPr><p:cNvPr id="{index}" name="{shape.preset} {index}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm{_xfrm_attributes(shape.rotation, shape.flip_horizontal, shape.flip_vertical)}><a:off x="{_emu(shape.left)}" y="{_emu(shape.top)}"/><a:ext cx="{_emu(shape.width)}" cy="{_emu(shape.height)}"/></a:xfrm><a:prstGeom prst="{shape.preset}"><a:avLst/></a:prstGeom>{fill}{line}</p:spPr></p:sp>')
    text_shapes = []
    for index, box in enumerate(parts, len(basic_shapes) + 2):
        paragraphs = _paragraphs(box, hyperlink_ids)
        text_shapes.append(f'<p:sp><p:nvSpPr><p:cNvPr id="{index}" name="Text Box {index-1}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm{_xfrm_attributes(box.rotation, box.flip_horizontal, box.flip_vertical)}><a:off x="{_emu(box.left)}" y="{_emu(box.top)}"/><a:ext cx="{_emu(box.width)}" cy="{_emu(box.height)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square"/><a:lstStyle/>{paragraphs}</p:txBody></p:sp>')
    footer_shapes = _header_footer_shapes(header_footer, slide_width, slide_height, slide_number,
                                          len(basic_shapes) + len(parts) + 2)
    picture_shapes = []
    for index, (picture, relation_id) in enumerate(pictures, len(basic_shapes) + len(parts) + len(footer_shapes) + 2):
        crop = f'<a:srcRect l="{picture.crop_left}" t="{picture.crop_top}" r="{picture.crop_right}" b="{picture.crop_bottom}"/>' if any((picture.crop_left, picture.crop_top, picture.crop_right, picture.crop_bottom)) else ''
        picture_shapes.append(f'<p:pic><p:nvPicPr><p:cNvPr id="{index}" name="Picture {index}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="{relation_id}"/>{crop}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm{_xfrm_attributes(picture.rotation, picture.flip_horizontal, picture.flip_vertical)}><a:off x="{_emu(picture.left)}" y="{_emu(picture.top)}"/><a:ext cx="{_emu(picture.width)}" cy="{_emu(picture.height)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>')
    if background_color and background_color_end:
        background = f'<p:bg><p:bgPr><a:gradFill rotWithShape="0"><a:gsLst><a:gs pos="0"><a:srgbClr val="{background_color}"/></a:gs><a:gs pos="100000"><a:srgbClr val="{background_color_end}"/></a:gs></a:gsLst><a:lin ang="0" scaled="1"/></a:gradFill><a:effectLst/></p:bgPr></p:bg>'
    else:
        background = f'<p:bg><p:bgPr><a:solidFill><a:srgbClr val="{background_color}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>' if background_color else ''
    show = ' show="0"' if hidden else ''
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"' + show + '><p:cSld>' + background + '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' + ''.join(drawing_shapes) + ''.join(text_shapes) + ''.join(footer_shapes) + ''.join(picture_shapes) + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'

def _notes_slide(values: tuple[str, ...]) -> str:
    text = "\r".join(values)
    paragraphs = _paragraphs(TextBox(text, 0, 0, 0, 0, (TextRun(text, font_size=12),)), {})
    tree = '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image"/><p:cNvSpPr/><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="1143000" y="685800"/><a:ext cx="4572000" cy="3429000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Body"/><p:cNvSpPr txBox="1"/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="4343400"/><a:ext cx="5486400" cy="4114800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>' + paragraphs + '</p:txBody></p:sp></p:spTree>'
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
                note_overrides += '<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>'
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
                archive.writestr(f'ppt/slides/slide{i}.xml', _slide(slides[i-1].text_boxes, picture_refs, slides[i-1].shapes, slides[i-1].background_color, slides[i-1].background_color_end, hyperlink_ids, slides[i-1].header_footer, presentation.width, presentation.height, i, slides[i-1].hidden))
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
            archive.writestr('ppt/presentation.xml', f'<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId{master_rid}"/></p:sldMasterIdLst>{notes_master_ids}<p:sldIdLst>{"".join(ids)}</p:sldIdLst><p:sldSz cx="{_emu(presentation.width)}" cy="{_emu(presentation.height)}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>')
            sp_tree = '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree>'
            if has_notes:
                archive.writestr('ppt/notesMasters/notesMaster1.xml', '<?xml version="1.0" encoding="UTF-8"?><p:notesMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld>' + sp_tree + '</p:cSld><p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/><p:hf hdr="1" ftr="1" dt="1" sldNum="1"/><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:notesMaster>')
                archive.writestr('ppt/notesMasters/_rels/notesMaster1.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>')
            archive.writestr('ppt/slideMasters/slideMaster1.xml', f'<?xml version="1.0" encoding="UTF-8"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld>{sp_tree}</p:cSld><p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/><p:sldLayoutIdLst><p:sldLayoutId id="1" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>')
            archive.writestr('ppt/slideMasters/_rels/slideMaster1.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>')
            archive.writestr('ppt/slideLayouts/slideLayout1.xml', f'<?xml version="1.0" encoding="UTF-8"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank"><p:cSld name="Blank">{sp_tree}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>')
            archive.writestr('ppt/slideLayouts/_rels/slideLayout1.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>')
            archive.writestr('ppt/theme/theme1.xml', '<?xml version="1.0" encoding="UTF-8"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="ppt2pptx"><a:themeElements><a:clrScheme name="Default"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Default"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Default"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>')
        _validate_package(temporary)
        temporary.replace(target)
    except Exception:
        temporary.unlink(missing_ok=True); raise
