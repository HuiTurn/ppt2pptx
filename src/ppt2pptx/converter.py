from __future__ import annotations

from dataclasses import dataclass, replace
from pathlib import Path

from .cfb import CompoundFile, Limits
from .diagnostics import ConversionReport
from .encryption import decrypt_powerpoint_document
from .errors import EncryptedPresentationError, InvalidPpt, UnsupportedPptVersionError, UnsafeOutputPathError
from .ooxml import write_pptx
from .oleps import read_summary_information
from .ppt import Presentation, extract_presentation

@dataclass(frozen=True, slots=True)
class ConversionResult:
    output_path: Path
    report: ConversionReport
    slide_count: int
    presentation: Presentation

def _load(source: str | Path, limits: Limits | None = None, report: ConversionReport | None = None, password: str | None = None) -> Presentation:
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
    return presentation

def inspect_ppt(source: str | Path, *, limits: Limits | None = None, password: str | None = None) -> dict[str, object]:
    presentation = _load(source, limits, password=password)
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
                     "rotation": box.rotation, "flip_horizontal": box.flip_horizontal,
                     "flip_vertical": box.flip_vertical,
                     "runs": [
                         {"text": run.text, "bold": run.bold, "italic": run.italic,
                          "underline": run.underline, "font_size": run.font_size,
                          "color": run.color, "typeface": run.typeface,
                          "hyperlink": run.hyperlink}
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
                     "rotation": picture.rotation, "flip_horizontal": picture.flip_horizontal,
                     "flip_vertical": picture.flip_vertical}
                    for picture in slide.pictures
                ],
                "shapes": [
                    {"preset": shape.preset, "left": shape.left, "top": shape.top,
                     "width": shape.width, "height": shape.height,
                     "fill_color": shape.fill_color, "line_color": shape.line_color,
                     "rotation": shape.rotation, "flip_horizontal": shape.flip_horizontal,
                     "flip_vertical": shape.flip_vertical}
                    for shape in slide.shapes
                ],
                "background_color": slide.background_color,
                "background_color_end": slide.background_color_end,
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
    presentation = _load(source_path, limits, report, password)
    report.warning("ADVANCED_FEATURES_APPROXIMATED", "charts, animations, audio/video, and complex freeform master geometry may be omitted or approximated")
    if not presentation.slides: report.warning("NO_SLIDES_FOUND", "no normal slide records could be recovered from the presentation")
    elif any(not (slide.text_boxes or slide.pictures or slide.shapes or slide.header_footer or slide.notes)
             for slide in presentation.slides):
        report.warning("EMPTY_SLIDE_CONTENT", "one or more slides had no recoverable editable content")
    write_pptx(output, presentation)
    return ConversionResult(output, report, len(presentation.slides), presentation)
