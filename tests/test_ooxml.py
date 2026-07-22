from pathlib import Path
import tempfile
import unittest
from xml.etree import ElementTree
import zipfile

from ppt2pptx.ooxml import write_pptx
from ppt2pptx.ppt import Comment, CoreProperties, HeaderFooter, Presentation, Slide, TextBox, TextRun

class OoxmlTests(unittest.TestCase):
    def test_writes_positioned_text_and_complete_layout_relationships(self):
        presentation = Presentation(5760, 4320, (Slide((TextBox("Hello", 288, 576, 1152, 288),)),))
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "output.pptx"
            write_pptx(path, presentation)
            with zipfile.ZipFile(path) as archive:
                xml = archive.read("ppt/slides/slide1.xml")
                root = ElementTree.fromstring(xml)
                ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
                offset = root.find(".//a:off", ns)
                extent = root.find(".//a:ext", ns)
                self.assertEqual(offset.attrib, {"x": "457200", "y": "914400"})
                self.assertEqual(extent.attrib, {"cx": "1828800", "cy": "457200"})
                self.assertIn("ppt/slideLayouts/slideLayout1.xml", archive.namelist())
                self.assertIn("ppt/theme/theme1.xml", archive.namelist())

    def test_writes_core_properties(self):
        presentation = Presentation(5760, 4320, (), CoreProperties(title="Résumé", creator="Ada", created="2020-01-02T03:04:05Z"))
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "metadata.pptx"
            write_pptx(path, presentation)
            with zipfile.ZipFile(path) as archive:
                core = ElementTree.fromstring(archive.read("docProps/core.xml"))
        self.assertEqual(core.find("{http://purl.org/dc/elements/1.1/}title").text, "Résumé")
        self.assertEqual(core.find("{http://purl.org/dc/elements/1.1/}creator").text, "Ada")

    def test_writes_styles_hyperlinks_comments_and_footer_fields(self):
        box = TextBox("Styled link", 100, 100, 2000, 400,
                      (TextRun("Styled link", bold=True, italic=True, underline=True,
                               font_size=24, color="FF0000", typeface="Times New Roman",
                               hyperlink="https://example.com"),))
        slide = Slide((box,), comments=(Comment("Ada", "A", "Review this", 10, 20),),
                      header_footer=HeaderFooter(footer_text="Footer", show_slide_number=True),
                      notes=("Speaker note",))
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "features.pptx"
            write_pptx(path, Presentation(5760, 4320, (slide,)))
            with zipfile.ZipFile(path) as archive:
                xml = archive.read("ppt/slides/slide1.xml").decode()
                rels = archive.read("ppt/slides/_rels/slide1.xml.rels").decode()
                names = set(archive.namelist())
        self.assertIn('b="1"', xml)
        self.assertIn('i="1"', xml)
        self.assertIn('u="sng"', xml)
        self.assertIn('type="slidenum"', xml)
        self.assertIn("Footer", xml)
        self.assertIn('Target="https://example.com"', rels)
        self.assertIn("ppt/comments/comment1.xml", names)
        self.assertIn("ppt/notesSlides/notesSlide1.xml", names)
        self.assertIn("ppt/notesMasters/notesMaster1.xml", names)
