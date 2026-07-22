import struct
import unittest
from ppt2pptx.ppt import extract_presentation, extract_slides

def rec(kind, payload=b"", version=0xF, instance=0): return struct.pack("<HHI", (instance << 4) | version, kind, len(payload)) + payload

class PptParserTests(unittest.TestCase):
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
        fopt = rec(0xF00B, struct.pack("<HI", 0x4104, 1), 3, 1)
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
