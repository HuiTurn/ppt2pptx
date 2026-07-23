from pathlib import Path
import tempfile
import unittest
from xml.etree import ElementTree
import zipfile

from ppt2pptx.ooxml import _xfrm_box, write_pptx
from ppt2pptx.ppt import (
    BasicShape, Comment, CoreProperties, HeaderFooter, Picture, Presentation,
    Slide, TextBox, TextRun,
)

class OoxmlTests(unittest.TestCase):
    def test_swaps_quarter_turn_shape_extents_around_the_same_center(self):
        self.assertEqual(
            _xfrm_box(96, 1584, 5472, 240, 5400000),
            (2712, -1032, 240, 5472),
        )

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

    def test_marks_hidden_slides_in_presentation(self):
        presentation = Presentation(5760, 4320, (Slide(()), Slide((), hidden=True)))
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "hidden.pptx"
            write_pptx(path, presentation)
            with zipfile.ZipFile(path) as archive:
                visible = ElementTree.fromstring(archive.read("ppt/slides/slide1.xml"))
                hidden = ElementTree.fromstring(archive.read("ppt/slides/slide2.xml"))
        self.assertNotIn("show", visible.attrib)
        self.assertEqual(hidden.attrib["show"], "0")

    def test_writes_autofit_vertical_anchor_and_nested_bullet(self):
        box = TextBox(
            "Top\rNested", 100, 100, 2000, 400,
            (TextRun("Top\r", font_size=28), TextRun("Nested", font_size=24)),
            paragraph_bullets=(True, True),
            paragraph_levels=(0, 1),
            paragraph_bullet_chars=("•", "–"),
            auto_fit=True,
            vertical_anchor="ctr",
        )
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "text-layout.pptx"
            write_pptx(path, Presentation(5760, 4320, (Slide((box,)),)))
            with zipfile.ZipFile(path) as archive:
                xml = archive.read("ppt/slides/slide1.xml").decode()
        self.assertIn('<a:bodyPr wrap="square" anchor="ctr"><a:normAutofit/></a:bodyPr>', xml)
        self.assertIn('lvl="1"', xml)
        self.assertIn('<a:buChar char="–"/>', xml)

    def test_writes_text_shape_geometry(self):
        box = TextBox("Inside", 100, 100, 500, 500, preset="ellipse",
                      wrap_text=False, fit_shape_to_text=True)
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "ellipse.pptx"
            write_pptx(path, Presentation(5760, 4320, (Slide((box,)),)))
            with zipfile.ZipFile(path) as archive:
                xml = archive.read("ppt/slides/slide1.xml").decode()
        self.assertIn('<a:prstGeom prst="ellipse">', xml)
        self.assertIn('<a:bodyPr wrap="none"><a:spAutoFit/></a:bodyPr>', xml)

    def test_normalizes_leading_tabs_for_centered_and_left_text(self):
        boxes = (
            TextBox(
                "\tCentered", 0, 0, 1000, 300,
                paragraph_alignments=("ctr",),
                tab_stops=((120, "l"),),
            ),
            TextBox(
                "\tIndented", 0, 400, 1000, 300,
                paragraph_alignments=("l",),
                paragraph_left_margins=(0,),
                paragraph_indents=(0,),
                tab_stops=((120, "l"),),
            ),
        )
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "tabs.pptx"
            write_pptx(path, Presentation(5760, 4320, (Slide(boxes),)))
            with zipfile.ZipFile(path) as archive:
                xml = archive.read("ppt/slides/slide1.xml").decode()
        self.assertNotIn(">\\tCentered<", xml)
        self.assertNotIn(">\\tIndented<", xml)
        self.assertIn('algn="l" marL="190500" indent="0"', xml)

    def test_writes_detailed_text_line_and_picture_properties(self):
        box = TextBox(
            "Tc",
            100,
            100,
            1000,
            400,
            runs=(TextRun("T"), TextRun("c", baseline=-25)),
            paragraph_bullets=(True,),
            paragraph_left_margins=(39,),
            paragraph_indents=(0,),
            paragraph_line_spacings=(90,),
            paragraph_space_before=(-20,),
            paragraph_space_after=(10,),
            inset_left=0,
            inset_top=0,
            inset_right=0,
            inset_bottom=0,
            line_color="112233",
            line_width=38100,
            default_tab_size=108,
            tab_stops=((540, "l"),),
        )
        picture = Picture(
            b"image",
            "png",
            "image/png",
            100,
            500,
            1000,
            1000,
            transparent_color="FF0000",
        )
        shape = BasicShape(
            "line", 50, 50, 500, 1, fill_color="000000",
            line_color="000000", line_width=12700,
            fill_pattern="dkUpDiag", fill_back_color="FFFF00",
            line_head=("stealth", None, None),
            line_tail=("triangle", "lg", "sm"),
        )
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "details.pptx"
            write_pptx(
                path,
                Presentation(5760, 4320, (Slide((box,), (picture,), (shape,)),)),
            )
            with zipfile.ZipFile(path) as archive:
                xml = archive.read("ppt/slides/slide1.xml").decode()
        self.assertIn('baseline="-25000"', xml)
        self.assertIn('marL="61912" indent="-61912"', xml)
        self.assertIn('<a:lnSpc><a:spcPct val="90000"/></a:lnSpc>', xml)
        self.assertIn('<a:spcBef><a:spcPts val="250"/></a:spcBef>', xml)
        self.assertIn('<a:spcAft><a:spcPct val="10000"/></a:spcAft>', xml)
        self.assertIn('lIns="0" tIns="0" rIns="0" bIns="0"', xml)
        self.assertIn('<a:ln w="38100">', xml)
        self.assertIn('<a:ln w="12700">', xml)
        self.assertIn('<a:headEnd type="stealth"/>', xml)
        self.assertIn('<a:tailEnd type="triangle" w="lg" len="sm"/>', xml)
        self.assertIn('<a:pattFill prst="dkUpDiag">', xml)
        self.assertIn('<a:bgClr><a:srgbClr val="FFFF00"/></a:bgClr>', xml)
        self.assertIn('defTabSz="171450"', xml)
        self.assertIn('<a:tab pos="857250" algn="l"/>', xml)
        self.assertIn('<a:clrFrom><a:srgbClr val="FF0000"/></a:clrFrom>', xml)
        self.assertIn('<a:alpha val="0"/>', xml)
        self.assertLess(xml.index("<p:pic>"), xml.index("<p:sp>"))
        self.assertLess(xml.index("<p:pic>"), xml.index("<p:txBody>"))
