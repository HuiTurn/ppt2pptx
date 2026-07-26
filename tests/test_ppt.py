import struct
import unittest
from ppt2pptx.ppt import (
    SHAPE_PRESETS, TextContent, TextRun, _GroupSpace, _MasterStyle, _anchor, _fopt_complex_properties,
    _connector_transform, _fopt_properties, _freeform_path, _gradient_angle, _line_dash, _line_end,
    _parse_text_ruler, _shape_style,
    _parse_text_ruler_details,
    _minimum_unwrapped_width, _minimum_wrapped_height, _parse_slide, _pictures,
    _shape_adjustments,
    _skip_paragraph_properties, _style_text,
    _legacy_arc_path, _text, _uses_custom_geometry, extract_presentation,
    extract_slides, records,
)

def rec(kind, payload=b"", version=0xF, instance=0): return struct.pack("<HHI", (instance << 4) | version, kind, len(payload)) + payload

class PptParserTests(unittest.TestCase):
    def test_uses_valid_ooxml_name_for_legacy_arrow(self):
        self.assertEqual(SHAPE_PRESETS[13], "rightArrow")
        self.assertEqual(SHAPE_PRESETS[34], "bentConnector3")
        self.assertEqual(SHAPE_PRESETS[88], "rightBrace")
        self.assertEqual(_line_dash({462: 2}), "sysDot")
        self.assertEqual(_line_dash({462: 7}), "lgDash")
        self.assertEqual(
            _line_end({465: 1, 468: 2, 469: 0}, 465, 468, 469),
            ("triangle", "lg", "sm"),
        )
        self.assertEqual(
            _shape_style({384: 1, 385: 0, 387: 0x0000FFFF}, ()),
            ("000000", None, None, "dkUpDiag", "FFFF00"),
        )
        self.assertEqual(_gradient_angle({395: 0xFF2E0000}), 18000000)
        self.assertEqual(_shape_adjustments(34, {327: 4669}), (21616,))
        self.assertEqual(_shape_adjustments(34, {327: 0xFFFFFFDD}), (-162,))
        self.assertEqual(
            _connector_transform(34, (5400000, True, False)),
            (16200000, True, False),
        )
        self.assertEqual(
            _connector_transform(34, (0, False, False)),
            (10800000, True, True),
        )
        self.assertGreater(
            _minimum_unwrapped_width(
                TextContent("A much wider line", (TextRun("A much wider line", font_size=20),))
            ),
            500,
        )
        self.assertGreater(
            _minimum_wrapped_height(
                TextContent("A long line that wraps", (
                    TextRun("A long line that wraps", font_size=20),
                )),
                500, None, None, None, None,
            ),
            400,
        )
        self.assertLess(
            _minimum_wrapped_height(
                TextContent("Five\rshort\rlines\rfit\rnormally", (
                    TextRun("Five\rshort\rlines\rfit\rnormally", font_size=20),
                )),
                1566, 90488, 90488, 45720, 45720,
            ),
            1200,
        )
        self.assertFalse(_uses_custom_geometry(19, {325: b"vertices"}))
        self.assertTrue(
            _uses_custom_geometry(19, {325: b"vertices", 326: b"segments"})
        )
        arc_vertices = struct.pack(
            "<3H8h", 4, 4, 4,
            0, 0, 21600, 21600, 0, 21600, 21600, 0,
        )
        self.assertEqual(
            _legacy_arc_path({325: arc_vertices})[0][-1],
            ("C", (0, 9669), (9669, 0), (21600, 0)),
        )
        symbol = next(records(rec(4000, "\uf044 \uf0de \uf0bb".encode("utf-16le"), 0)))
        self.assertEqual(_text(symbol), "Δ ⇒ ≈")

    def test_stops_safely_on_truncated_paragraph_properties(self):
        self.assertEqual(
            _skip_paragraph_properties(b"", 0, 0x80),
            (0, None, None, None),
        )

    def test_applies_master_style_for_each_paragraph_level(self):
        text = "Top\rNested"
        paragraph_runs = (
            struct.pack("<IhI", 4, 0, 0)
            + struct.pack("<IhI", 7, 1, 0)
        )
        character_runs = struct.pack("<II", len(text) + 1, 0)
        masters = (
            _MasterStyle(TextRun("", font_size=28, color="FFFFFF", typeface="Arial"),
                         None, True, "•"),
            _MasterStyle(TextRun("", font_size=24), None, None, "–"),
        )

        content = _style_text(text, paragraph_runs + character_runs, (), masters, (), 1)

        self.assertEqual([(run.text, run.font_size) for run in content.runs],
                         [("Top\r", 28), ("Nested", 24)])
        self.assertEqual(content.paragraph_levels, (0, 1))
        self.assertEqual(content.paragraph_bullet_chars, ("•", "–"))
        self.assertEqual((content.runs[1].color, content.runs[1].typeface),
                         ("FFFFFF", "Arial"))

    def test_preserves_paragraph_layout_baseline_and_text_ruler(self):
        text = "Tc"
        paragraph_mask = 0x100 | 0x400 | 0x1000 | 0x2000 | 0x4000
        paragraph = (
            struct.pack("<IhI", len(text) + 1, 0, paragraph_mask)
            + struct.pack("<5h", 90, -20, 10, 100, 20)
        )
        characters = (
            struct.pack("<II", 1, 0)
            + struct.pack("<IIh", 1, 0x80000, -25)
        )
        content = _style_text(text, paragraph + characters, (), (), (), 4)

        self.assertEqual(content.runs[1].baseline, -25)
        self.assertEqual(content.paragraph_left_margins, (100,))
        self.assertEqual(content.paragraph_indents, (20,))
        self.assertEqual(content.paragraph_line_spacings, (90,))
        self.assertEqual(content.paragraph_space_before, (-20,))
        self.assertEqual(content.paragraph_space_after, (10,))
        self.assertEqual(_parse_text_ruler(struct.pack("<Ih", 0x8, 39))[0],
                         (39, None))
        levels, default_tab, tabs = _parse_text_ruler_details(
            bytes.fromhex("050000006c0001001c020000")
        )
        self.assertEqual(levels[0], (None, None))
        self.assertEqual(default_tab, 108)
        self.assertEqual(tabs, ((540, "l"),))

    def test_extracts_text_from_direct_slide_record(self):
        text = rec(4000, "Hello\rWorld".encode("utf-16le"), 0)
        slide = rec(1006, text)
        document = rec(1000, slide)
        self.assertEqual(extract_slides(document), [["Hello\rWorld"]])

    def test_resolves_slide_through_persist_directory(self):
        text = rec(4000, "Persisted".encode("utf-16le"), 0)
        slide = rec(1006, text)
        # Persist id 2 points to the record after the document and directory.
        persist_atom = rec(1011, struct.pack("<5I", 2, 0, 0, 0, 0), 0)
        document = rec(1000, rec(4080, persist_atom))
        offset = len(document) + 8 + 12  # persist header + entry descriptor
        directory = rec(6002, struct.pack("<I2I", (2 << 20) | 1, 0, offset), 0)
        self.assertEqual(extract_slides(document + directory + slide), [["Persisted"]])

    def test_reads_document_dimensions(self):
        document_atom = rec(1001, struct.pack("<2I", 7200, 4050) + bytes(32), 1)
        presentation = extract_presentation(rec(1000, document_atom))
        self.assertEqual((presentation.width, presentation.height), (7200, 4050))

    def test_extracts_png_picture_and_anchor(self):
        bse_payload = bytearray(36)
        struct.pack_into("<I", bse_payload, 28, 0)
        bse = rec(0xF007, bytes(bse_payload), 2, 6)
        persist_atom = rec(1011, struct.pack("<5I", 2, 0, 0, 0, 0), 0)
        document = rec(1000, bse + rec(4080, persist_atom))
        fopt = rec(
            0xF00B,
            b"".join((
                struct.pack("<HI", 0x4104, 1),
                struct.pack("<HI", 263, 0x000000FF),
                struct.pack("<HI", 448, 0x08000000),
                struct.pack("<HI", 459, 19050),
                struct.pack("<HI", 511, 0x00080008),
            )),
            3,
            5,
        )
        anchor = rec(0xF010, struct.pack("<4h", 576, 288, 1728, 1152), 0)
        slide = rec(1006, rec(0xF004, fopt + anchor))
        offset = len(document) + 20
        directory = rec(6002, struct.pack("<I2I", (2 << 20) | 1, 0, offset), 0)
        png = b"\x89PNG\r\n\x1a\n" + bytes(16)
        pictures = rec(0xF01E, bytes(16) + b"\xff" + png, 0, 0x6E0)
        presentation = extract_presentation(document + directory + slide, pictures)
        picture = presentation.slides[0].pictures[0]
        self.assertEqual(picture.data, png)
        self.assertEqual((picture.left, picture.top, picture.width, picture.height), (288, 576, 1440, 576))
        self.assertEqual(picture.transparent_color, "FF0000")
        self.assertEqual(picture.line_color, "FFFFFF")
        self.assertEqual(picture.line_width, 19050)

    def test_embeds_standard_wmf_without_an_aldus_placeable_header(self):
        bse_payload = bytearray(36)
        struct.pack_into("<I", bse_payload, 28, 0)
        document = next(records(rec(1000, rec(0xF007, bytes(bse_payload), 2, 3))))
        raw = b"\x01\x00\x09\x00\x00\x03" + bytes(12)
        blip_payload = bytearray(16 + 34)
        struct.pack_into("<I4i", blip_payload, 16, len(raw), 0, 0, 100, 100)
        struct.pack_into("<I", blip_payload, 44, len(raw))
        blip_payload[48] = 0xFE
        stream = rec(0xF01B, bytes(blip_payload) + raw, 0, 0x216)

        data, extension, content_type = _pictures(document, stream)[1]

        self.assertEqual(data, raw)
        self.assertEqual((extension, content_type), ("wmf", "image/x-wmf"))
        self.assertFalse(data.startswith(b"\xd7\xcd\xc6\x9a"))

    def test_places_master_picture_behind_slide_picture(self):
        def picture_shape(reference):
            fopt = rec(0xF00B, struct.pack("<HI", 0x4104, reference), 3, 1)
            anchor = rec(0xF010, struct.pack("<4h", 100, 100, 500, 500), 0)
            return rec(0xF004, fopt + anchor)

        master = next(records(rec(1016, picture_shape(1))))
        slide = next(records(rec(1006, picture_shape(2))))
        image_map = {
            1: (b"master", "png", "image/png"),
            2: (b"slide", "png", "image/png"),
        }

        parsed = _parse_slide(
            slide, image_map, [], (), {}, {}, None, (), master, (), 5760, 4320
        )

        self.assertEqual([picture.data for picture in parsed.pictures],
                         [b"master", b"slide"])

    def test_resolves_legacy_external_slide_text(self):
        persist_atom = rec(1011, struct.pack("<5I", 2, 0, 0, 0, 0), 0)
        external = persist_atom + rec(3999, struct.pack("<I", 1), 0) + rec(4008, b"Legacy text", 0)
        document = rec(1000, rec(4080, external, 0xF, 0))
        textbox = rec(0xF00D, rec(3998, struct.pack("<I", 0), 0))
        anchor = rec(0xF010, struct.pack("<4h", 100, 200, 1200, 500), 0)
        slide = rec(1006, rec(0xF004, anchor + textbox))
        offset = len(document) + 20
        directory = rec(6002, struct.pack("<I2I", (2 << 20) | 1, 0, offset), 0)
        presentation = extract_presentation(document + directory + slide)
        box = presentation.slides[0].text_boxes[0]
        self.assertEqual(box.text, "Legacy text")
        self.assertEqual((box.left, box.top, box.width, box.height), (200, 100, 1000, 400))

    def test_uses_only_normal_slide_list_and_reads_footer(self):
        master_ref = rec(1011, struct.pack("<5I", 2, 0, 0, 0, 0), 0)
        slide_ref = rec(1011, struct.pack("<5I", 3, 0, 0, 0, 0), 0)
        options = rec(4058, struct.pack("<2H", 0, 0x28), 0)
        footer = rec(4026, "Global footer".encode("utf-16le"), 0, 2)
        document = rec(1000, rec(4080, master_ref, instance=1) + rec(4080, slide_ref, instance=0)
                       + rec(4057, options + footer, instance=3))
        directory_size = len(rec(6002, struct.pack("<I2I", (2 << 20) | 2, 0, 0), 0))
        master_offset = len(document) + directory_size
        master = rec(1016)
        slide_offset = master_offset + len(master)
        directory = rec(6002, struct.pack("<I2I", (2 << 20) | 2, master_offset, slide_offset), 0)
        presentation = extract_presentation(document + directory + master + rec(1006))
        self.assertEqual(len(presentation.slides), 1)
        self.assertEqual(presentation.slides[0].header_footer.footer_text, "Global footer")
        self.assertTrue(presentation.slides[0].header_footer.show_slide_number)

    def test_recovers_speaker_notes_without_counting_note_as_slide(self):
        slide_ref = rec(1011, struct.pack("<5I", 2, 0, 0, 0, 0), 0)
        note_ref = rec(1011, struct.pack("<5I", 3, 0, 0, 0, 0), 0)
        document = rec(1000, rec(4080, slide_ref, instance=0) + rec(4080, note_ref, instance=2))
        directory_size = len(rec(6002, struct.pack("<I2I", (2 << 20) | 2, 0, 0), 0))
        slide_offset = len(document) + directory_size
        slide = rec(1006, rec(4008, b"Slide", 0))
        note_offset = slide_offset + len(slide)
        directory = rec(6002, struct.pack("<I2I", (2 << 20) | 2, slide_offset, note_offset), 0)
        textbox = rec(0xF00D, rec(4008, b"Speaker note", 0))
        anchor = rec(0xF010, struct.pack("<4h", 100, 100, 1000, 500), 0)
        note = rec(1008, rec(0xF004, anchor + textbox))
        presentation = extract_presentation(document + directory + slide + note)
        self.assertEqual(len(presentation.slides), 1)
        self.assertEqual(presentation.slides[0].notes, ("Speaker note",))

    def test_binds_sparse_speaker_notes_by_slide_id(self):
        slide_ref1 = rec(1011, struct.pack("<5I", 2, 0, 0, 11, 0), 0)
        slide_ref2 = rec(1011, struct.pack("<5I", 3, 0, 0, 22, 0), 0)
        note_ref = rec(1011, struct.pack("<5I", 4, 0, 0, 0, 0), 0)
        document = rec(
            1000,
            rec(4080, slide_ref1 + slide_ref2, instance=0)
            + rec(4080, note_ref, instance=2),
        )
        directory_size = len(
            rec(6002, struct.pack("<I3I", (3 << 20) | 2, 0, 0, 0), 0)
        )
        slide1_offset = len(document) + directory_size
        slide1 = rec(1006, rec(4008, b"Slide 1", 0))
        slide2_offset = slide1_offset + len(slide1)
        slide2 = rec(1006, rec(4008, b"Slide 2", 0))
        note_offset = slide2_offset + len(slide2)
        directory = rec(
            6002,
            struct.pack(
                "<I3I",
                (3 << 20) | 2,
                slide1_offset,
                slide2_offset,
                note_offset,
            ),
            0,
        )
        textbox = rec(0xF00D, rec(4008, b"Second slide note", 0))
        anchor = rec(0xF010, struct.pack("<4h", 100, 100, 1000, 500), 0)
        note = rec(
            1008,
            rec(1009, struct.pack("<I", 22), 0) + rec(0xF004, anchor + textbox),
        )

        presentation = extract_presentation(
            document + directory + slide1 + slide2 + note
        )

        self.assertEqual(len(presentation.slides), 2)
        self.assertEqual(presentation.slides[0].notes, ())
        self.assertEqual(presentation.slides[1].notes, ("Second slide note",))

    def test_preserves_hidden_slide_flag(self):
        slide_ref = rec(1011, struct.pack("<5I", 2, 0, 0, 0, 0), 0)
        document = rec(1000, rec(4080, slide_ref, instance=0))
        directory_size = len(rec(6002, struct.pack("<2I", (1 << 20) | 2, 0), 0))
        slide_offset = len(document) + directory_size
        directory = rec(6002, struct.pack("<2I", (1 << 20) | 2, slide_offset), 0)
        slide_show_info = bytearray(16)
        struct.pack_into("<H", slide_show_info, 10, 0x0004)
        slide = rec(1006, rec(0x03F9, bytes(slide_show_info), 0))
        presentation = extract_presentation(document + directory + slide)
        self.assertTrue(presentation.slides[0].hidden)

    def test_recovers_filled_table_cells_as_editable_shapes(self):
        slide_ref = rec(1011, struct.pack("<5I", 2, 0, 0, 0, 0), 0)
        document = rec(1000, rec(4080, slide_ref, instance=0))
        directory_size = len(rec(6002, struct.pack("<2I", (1 << 20) | 2, 0), 0))
        slide_offset = len(document) + directory_size
        directory = rec(6002, struct.pack("<2I", (1 << 20) | 2, slide_offset), 0)
        fsp = rec(0xF00A, struct.pack("<2I", 1, 0), 2, 1)
        fill = rec(0xF00B, struct.pack("<HI", 385, 0x00D59B5B), 3, 1)
        anchor = rec(0xF00F, struct.pack("<4i", 100, 200, 900, 600), 0)
        cell = rec(0xF004, fsp + fill + anchor + rec(0xF00D), 0xF)
        presentation = extract_presentation(document + directory + rec(1006, cell))
        shape = presentation.slides[0].shapes[0]
        self.assertEqual(shape.preset, "rect")
        self.assertEqual(shape.fill_color, "5B9BD5")
        self.assertEqual((shape.left, shape.top, shape.width, shape.height), (100, 200, 800, 400))

    def test_maps_child_anchor_text_through_group_transform(self):
        slide_ref = rec(1011, struct.pack("<5I", 2, 0, 0, 0, 0), 0)
        document = rec(1000, rec(4080, slide_ref, instance=0))
        directory_size = len(rec(6002, struct.pack("<2I", (1 << 20) | 2, 0), 0))
        slide_offset = len(document) + directory_size
        directory = rec(6002, struct.pack("<2I", (1 << 20) | 2, slide_offset), 0)
        fspgr = rec(0xF009, struct.pack("<4i", 0, 0, 1000, 1000), 1, 0)
        group_anchor = rec(0xF010, struct.pack("<4h", 100, 200, 600, 500), 0)  # top,left,right,bottom
        group_sp = rec(0xF004, fspgr + rec(0xF00A, struct.pack("<2I", 1, 0x1), 2, 0) + group_anchor)
        textbox = rec(0xF00D, rec(4008, b"Grouped", 0))
        child_anchor = rec(0xF00F, struct.pack("<4i", 100, 200, 500, 400), 0)
        member = rec(0xF004, rec(0xF00A, struct.pack("<2I", 2, 0), 2, 1) + child_anchor + textbox)
        group = rec(0xF003, group_sp + member)
        presentation = extract_presentation(document + directory + rec(1006, group))
        box = presentation.slides[0].text_boxes[0]
        self.assertEqual(box.text, "Grouped")
        # child (100,200)-(500,400) in 1000x1000 space maps into abs (200,100)-(600,500)
        self.assertEqual((box.left, box.top, box.width, box.height), (240, 180, 160, 80))

    def test_keeps_large_child_anchor_in_parent_group_coordinates(self):
        space = _GroupSpace(5_715_376, 3_000_987, 8_115_743, 5_660_414,
                            3600, 1890, 5112, 3566)
        child = next(records(rec(0xF00F, struct.pack(
            "<4i", 5_715_376, 3_000_987, 8_115_743, 5_660_414
        ), 0)))
        self.assertEqual(_anchor([child], 0, space), (3600, 1890, 1512, 1676))

    def test_respects_explicit_no_line_flag_with_stored_line_color(self):
        slide_ref = rec(1011, struct.pack("<5I", 2, 0, 0, 0, 0), 0)
        document = rec(1000, rec(4080, slide_ref, instance=0))
        directory_size = len(rec(6002, struct.pack("<2I", (1 << 20) | 2, 0), 0))
        slide_offset = len(document) + directory_size
        directory = rec(6002, struct.pack("<2I", (1 << 20) | 2, slide_offset), 0)
        properties = struct.pack("<HIHI", 448, 0, 511, 0x90000)
        fopt = rec(0xF00B, properties, 3, 2)
        textbox = rec(0xF00D, rec(4008, b"No border", 0))
        anchor = rec(0xF010, struct.pack("<4h", 100, 100, 1000, 500), 0)
        shape = rec(0xF004, fopt + anchor + textbox)
        box = extract_presentation(document + directory + rec(1006, shape)).slides[0].text_boxes[0]
        self.assertIsNone(box.line_color)

    def test_preserves_geometry_for_text_bearing_shape(self):
        ellipse = rec(0xF00A, struct.pack("<2I", 1, 0), 2, 3)
        options = rec(0xF00B, struct.pack("<HIHI", 133, 2, 135, 1), 3, 2)
        textbox = rec(0xF00D, rec(4008, b"Ellipse", 0))
        anchor = rec(0xF010, struct.pack("<4h", 100, 100, 600, 600), 0)
        presentation = extract_presentation(rec(1000, rec(1006, rec(
            0xF004, ellipse + options + anchor + textbox
        ))))
        box = presentation.slides[0].text_boxes[0]
        self.assertEqual(box.preset, "ellipse")
        self.assertEqual(box.vertical_anchor, "ctr")
        self.assertFalse(box.auto_fit)
        self.assertFalse(box.wrap_text)

    def test_reads_both_legacy_text_autofit_modes(self):
        def shape(text: bytes, top: int, flags: int):
            options = rec(0xF00B, struct.pack("<HI", 191, flags), 3, 1)
            textbox = rec(0xF00D, rec(4008, text, 0))
            anchor = rec(0xF010, struct.pack("<4h", top, 100, top + 300, 600), 0)
            return rec(0xF004, options + anchor + textbox)

        presentation = extract_presentation(rec(1000, rec(
            1006,
            shape(b"Shrink text", 100, 0x40004)
            + shape(b"Grow shape", 500, 0xF0002),
        )))
        shrink, grow = presentation.slides[0].text_boxes
        self.assertTrue(shrink.auto_fit)
        self.assertFalse(shrink.fit_shape_to_text)
        self.assertFalse(grow.auto_fit)
        self.assertTrue(grow.fit_shape_to_text)

    def test_keeps_packed_vertex_header_out_of_following_segments(self):
        vertices = (
            struct.pack("<3H", 3, 3, 0xFFF0)
            + struct.pack("<6h", 0, 100, 50, 0, 100, 100)
        )
        segments = (
            struct.pack("<3H", 4, 4, 2)
            + struct.pack("<4H", 0x4000, 0x0001, 0x0001, 0x8000)
        )
        entries = struct.pack("<HIHI", 0xC145, 12, 0xC146, len(segments))
        fopt = next(records(rec(0xF00B, entries + vertices + segments, 3, 2)))
        complex_properties = _fopt_complex_properties(fopt)

        self.assertEqual(len(complex_properties[325]), 18)
        self.assertEqual(complex_properties[326][:6], struct.pack("<3H", 4, 4, 2))
        path, width, height = _freeform_path(
            _fopt_properties(fopt), complex_properties
        )
        self.assertEqual(path, (("M", (0, 100)), ("L", (50, 0)), ("L", (100, 100))))
        self.assertEqual((width, height), (100, 100))

    def test_ignores_trailing_garbage_after_valid_records(self):
        text = rec(4000, "OK".encode("utf-16le"), 0)
        slide = rec(1006, text)
        document = rec(1000, slide)
        presentation = extract_presentation(document + b"\x00" * 32)
        self.assertEqual(presentation.slides[0].text_boxes[0].text, "OK")
