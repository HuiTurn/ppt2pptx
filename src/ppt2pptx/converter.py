from __future__ import annotations

from dataclasses import dataclass, replace
from pathlib import Path

from .cfb import CompoundFile, Limits
from .diagnostics import ConversionReport
from .encryption import decrypt_powerpoint_document
from .errors import EncryptedPresentationError, InvalidPpt, UnsupportedPptVersionError, UnsafeOutputPathError
from .ooxml import write_pptx
from .oleps import read_summary_information
from .ppt import Presentation, detect_lossy_features, extract_presentation

@dataclass(frozen=True, slots=True)
class ConversionResult:
    output_path: Path
    report: ConversionReport
    slide_count: int
    presentation: Presentation

def _load(source: str | Path, limits: Limits | None = None, report: ConversionReport | None = None, password: str | None = None) -> tuple[Presentation, bytes]:
    compound = CompoundFile.from_path(source, limits)
    current_user = compound.open_stream("Current User")
    document_stream = compound.open_stream("PowerPoint Document")
    if len(current_user) >= 16 and current_user[12:16] == b"\xdf\xc4\xd1\xf3":
        document_stream = decrypt_powerpoint_document(document_stream, current_user, password)
        if report is not None and "encryptedsummary" in compound.by_name:
            report.warning("ENCRYPTED_METADATA_OMITTED", "encrypted document properties were not copied")
    if len(current_user) < 8 or current_user[2:4] != b"\xf6\x0f":
        raise UnsupportedPptVersionError("PowerPoint 95 and earlier presentations are outside the PowerPoint 97–2003 format")
    pictures = compound.open_stream("Pictures") if "pictures" in compound.by_name else None
    presentation = extract_presentation(document_stream, pictures)
    if "\x05summaryinformation" in compound.by_name:
        try:
            properties = read_summary_information(compound.open_stream("\x05SummaryInformation"))
            presentation = replace(presentation, core_properties=properties)
        except InvalidPpt as exc:
            if report is not None:
                report.warning("SUMMARY_INFORMATION_MALFORMED", "malformed optional presentation metadata was omitted", reason=str(exc))
    return presentation, document_stream

def inspect_ppt(source: str | Path, *, limits: Limits | None = None, password: str | None = None) -> dict[str, object]:
    presentation, _document = _load(source, limits, password=password)
    return {
        "source": str(source),
        "slide_count": len(presentation.slides),
        "slide_size": {"width": presentation.width, "height": presentation.height},
        "text_box_count": sum(len(slide.text_boxes) for slide in presentation.slides),
        "picture_count": sum(len(slide.pictures) for slide in presentation.slides),
        "shape_count": sum(len(slide.shapes) for slide in presentation.slides),
        "comment_count": sum(len(slide.comments) for slide in presentation.slides),
        "note_count": sum(len(slide.notes) for slide in presentation.slides),
        "core_properties": {
            name: value for name in presentation.core_properties.__dataclass_fields__
            if (value := getattr(presentation.core_properties, name)) is not None
        },
        "slides": [
            {
                "index": i + 1,
                "text_boxes": [
                    {"text": box.text, "left": box.left, "top": box.top,
                     "width": box.width, "height": box.height,
                     "paragraph_alignments": box.paragraph_alignments,
                     "paragraph_bullets": box.paragraph_bullets,
                     "paragraph_levels": box.paragraph_levels,
                     "paragraph_bullet_chars": box.paragraph_bullet_chars,
                     "paragraph_left_margins": box.paragraph_left_margins,
                     "paragraph_indents": box.paragraph_indents,
                     "paragraph_line_spacings": box.paragraph_line_spacings,
                     "paragraph_space_before": box.paragraph_space_before,
                     "paragraph_space_after": box.paragraph_space_after,
                     "auto_fit": box.auto_fit,
                     "fit_shape_to_text": box.fit_shape_to_text,
                     "vertical_anchor": box.vertical_anchor,
                     "inset_left": box.inset_left, "inset_top": box.inset_top,
                     "inset_right": box.inset_right, "inset_bottom": box.inset_bottom,
                     "is_placeholder": box.is_placeholder,
                     "preset": box.preset,
                     "wrap_text": box.wrap_text,
                     "fill_color": box.fill_color,
                     "fill_pattern": box.fill_pattern,
                     "fill_back_color": box.fill_back_color,
                     "line_color": box.line_color,
                     "line_dash": box.line_dash,
                     "line_width": box.line_width,
                     "line_head": box.line_head,
                     "line_tail": box.line_tail,
                     "default_tab_size": box.default_tab_size,
                     "tab_stops": box.tab_stops,
                     "rotation": box.rotation, "flip_horizontal": box.flip_horizontal,
                     "flip_vertical": box.flip_vertical,
                     "runs": [
                         {"text": run.text, "bold": run.bold, "italic": run.italic,
                          "underline": run.underline, "font_size": run.font_size,
                          "color": run.color, "typeface": run.typeface,
                          "hyperlink": run.hyperlink, "baseline": run.baseline}
                         for run in box.runs
                     ]}
                    for box in slide.text_boxes
                ],
                "pictures": [
                    {"extension": picture.extension, "content_type": picture.content_type,
                     "byte_count": len(picture.data), "left": picture.left,
                     "top": picture.top, "width": picture.width, "height": picture.height,
                     "crop_left": picture.crop_left, "crop_top": picture.crop_top,
                     "crop_right": picture.crop_right, "crop_bottom": picture.crop_bottom,
                     "transparent_color": picture.transparent_color,
                     "line_color": picture.line_color,
                     "line_dash": picture.line_dash,
                     "line_width": picture.line_width,
                     "rotation": picture.rotation, "flip_horizontal": picture.flip_horizontal,
                     "flip_vertical": picture.flip_vertical}
                    for picture in slide.pictures
                ],
                "shapes": [
                    {"preset": shape.preset, "left": shape.left, "top": shape.top,
                     "width": shape.width, "height": shape.height,
                     "fill_color": shape.fill_color,
                     "fill_pattern": shape.fill_pattern,
                     "fill_back_color": shape.fill_back_color,
                     "line_color": shape.line_color,
                     "line_dash": shape.line_dash, "line_width": shape.line_width,
                     "line_head": shape.line_head, "line_tail": shape.line_tail,
                     "adjustments": shape.adjustments,
                     "rotation": shape.rotation, "flip_horizontal": shape.flip_horizontal,
                     "flip_vertical": shape.flip_vertical}
                    for shape in slide.shapes
                ],
                "background_color": slide.background_color,
                "background_color_end": slide.background_color_end,
                "background_gradient_angle": slide.background_gradient_angle,
                "background_gradient_type": slide.background_gradient_type,
                "background_image": ({
                    "extension": slide.background_image[1],
                    "content_type": slide.background_image[2],
                    "byte_count": len(slide.background_image[0]),
                } if slide.background_image else None),
                "hidden": slide.hidden,
                "notes": list(slide.notes),
                "header_footer": ({
                    "date_text": slide.header_footer.date_text,
                    "date_is_auto": slide.header_footer.date_is_auto,
                    "header_text": slide.header_footer.header_text,
                    "footer_text": slide.header_footer.footer_text,
                    "show_slide_number": slide.header_footer.show_slide_number,
                } if slide.header_footer else None),
                "comments": [
                    {"author": comment.author, "initials": comment.initials,
                     "text": comment.text, "left": comment.left,
                     "top": comment.top, "created": comment.created}
                    for comment in slide.comments
                ],
            }
            for i, slide in enumerate(presentation.slides)
        ],
    }

def convert(source: str | Path, destination: str | Path | None = None, *, limits: Limits | None = None, password: str | None = None) -> ConversionResult:
    source_path = Path(source)
    output = Path(destination) if destination is not None else source_path.with_suffix(".pptx")
    if output.suffix.casefold() != ".pptx": raise UnsafeOutputPathError("destination must use the .pptx extension")
    if source_path.resolve() == output.resolve(): raise UnsafeOutputPathError("destination must not overwrite the source presentation")
    report = ConversionReport(str(source_path), str(output))
    presentation, document_stream = _load(source_path, limits, report, password)
    for feature in detect_lossy_features(document_stream, presentation.excluded_offsets):
        report.warning(
            feature.code,
            feature.message,
            count=feature.count,
            record_types=list(feature.record_types),
            locations=[
                {
                    "slide_index": location.slide_index,
                    "record_type": location.record_type,
                    "record_offset": location.record_offset,
                    "object_kind": location.object_kind,
                }
                for location in feature.locations
            ],
        )
    if not presentation.slides: report.warning("NO_SLIDES_FOUND", "no normal slide records could be recovered from the presentation")
    elif any(not (slide.text_boxes or slide.pictures or slide.shapes or slide.background_image
                  or slide.header_footer or slide.notes)
             for slide in presentation.slides):
        report.warning("EMPTY_SLIDE_CONTENT", "one or more slides had no recoverable editable content")
    write_pptx(output, presentation)
    return ConversionResult(output, report, len(presentation.slides), presentation)
