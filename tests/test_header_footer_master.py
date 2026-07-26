from __future__ import annotations

import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from xml.etree import ElementTree
import zipfile

from ppt2pptx.converter import convert, inspect_ppt


ROOT = Path(__file__).resolve().parents[1]
FIXTURE = ROOT / "tests" / "fixtures" / "visual_header_footer_master.ppt"
CORPUS = ROOT / "tests" / "real_samples" / "37625.ppt"
COMPARE = ROOT / "scripts" / "compare_powerpoint_visual.py"


def powerpoint_available() -> bool:
    if sys.platform != "win32":
        return False
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return False
    pythoncom.CoInitialize()
    app = None
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        return bool(app.Version)
    except Exception:
        return False
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


class HeaderFooterMasterTests(unittest.TestCase):
    def test_inherits_date_geometry_and_style_from_master_placeholder(self):
        self.assertTrue(FIXTURE.is_file(), f"missing fixture: {FIXTURE}")
        inspected = inspect_ppt(FIXTURE)
        header_footer = inspected["slides"][0]["header_footer"]
        self.assertEqual(header_footer["date_text"], "DATE42")
        self.assertEqual(
            header_footer["date_placeholder"],
            {
                "left": 48,
                "top": 4080,
                "width": 1200,
                "height": 288,
                "alignment": "l",
                "vertical_anchor": "b",
                "font_size": 14,
                "color": "40458C",
                "typeface": "Tahoma",
            },
        )

        with tempfile.TemporaryDirectory() as directory:
            result = convert(FIXTURE, Path(directory) / "fields.pptx")
            with zipfile.ZipFile(result.output_path) as archive:
                root = ElementTree.fromstring(
                    archive.read("ppt/slides/slide1.xml")
                )

        self.assertEqual(result.report.warnings, [])
        ns = {
            "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
        }
        date_shape = next(
            shape
            for shape in root.findall(".//p:sp", ns)
            if (
                shape.find("./p:nvSpPr/p:cNvPr", ns) is not None
                and shape.find("./p:nvSpPr/p:cNvPr", ns).attrib.get("name")
                == "Date"
            )
        )
        offset = date_shape.find("./p:spPr/a:xfrm/a:off", ns)
        extent = date_shape.find("./p:spPr/a:xfrm/a:ext", ns)
        body = date_shape.find("./p:txBody/a:bodyPr", ns)
        paragraph = date_shape.find("./p:txBody/a:p/a:pPr", ns)
        run = date_shape.find("./p:txBody/a:p/a:r/a:rPr", ns)
        color = run.find("./a:solidFill/a:srgbClr", ns)
        latin = run.find("./a:latin", ns)
        self.assertEqual(offset.attrib, {"x": "76200", "y": "6477000"})
        self.assertEqual(extent.attrib, {"cx": "1905000", "cy": "457200"})
        self.assertEqual(body.attrib.get("anchor"), "b")
        self.assertEqual(paragraph.attrib.get("algn"), "l")
        self.assertEqual(run.attrib.get("sz"), "1400")
        self.assertEqual(color.attrib.get("val"), "40458C")
        self.assertEqual(latin.attrib.get("typeface"), "Tahoma")
        self.assertEqual(
            date_shape.find("./p:txBody/a:p/a:r/a:t", ns).text,
            "DATE42",
        )

    def test_suppresses_master_fields_when_master_objects_or_title_fields_are_off(self):
        inspected = inspect_ppt(CORPUS)
        slides = inspected["slides"]
        self.assertIsNone(slides[0]["header_footer"])
        self.assertIsNotNone(slides[1]["header_footer"])
        self.assertIsNone(slides[2]["header_footer"])
        self.assertIsNone(slides[28]["header_footer"])


@unittest.skipUnless(powerpoint_available(), "Microsoft PowerPoint COM is required")
class HeaderFooterMasterPowerPointVisualTests(unittest.TestCase):
    def test_master_date_placeholder_has_exact_bilateral_visual(self):
        with tempfile.TemporaryDirectory() as directory:
            evidence = Path(directory) / "evidence"
            env = os.environ.copy()
            env["PYTHONPATH"] = str(ROOT / "src")
            completed = subprocess.run(
                [
                    sys.executable,
                    str(COMPARE),
                    str(FIXTURE),
                    "-o",
                    str(evidence),
                    "--width",
                    "960",
                    "--height",
                    "720",
                    "--timeout",
                    "240",
                ],
                capture_output=True,
                text=True,
                env=env,
                timeout=360,
                check=False,
            )
            self.assertEqual(
                completed.returncode,
                0,
                msg=f"stdout:\n{completed.stdout}\nstderr:\n{completed.stderr}",
            )
            report = json.loads(
                (evidence / "report.json").read_text(encoding="utf-8")
            )

        self.assertEqual(report["provider"], "office")
        self.assertEqual(report["powerpoint_version"], "16.0")
        self.assertEqual(report["hard_differences"], [])
        self.assertEqual(report["slide_count_source"], 1)
        self.assertEqual(report["slide_count_output"], 1)
        self.assertEqual(report["slide_width_pt"], 720.0)
        self.assertEqual(report["slide_height_pt"], 540.0)
        self.assertEqual(report["hidden_source"], [False])
        self.assertEqual(report["hidden_output"], [False])
        self.assertEqual(report["summary"]["mean_mae"], 0.0)
        self.assertEqual(report["summary"]["mean_rmse"], 0.0)
        self.assertEqual(
            report["summary"]["mean_changed_pixel_ratio"],
            0.0,
        )
        self.assertEqual(report["summary"]["mean_ssim"], 1.0)
        self.assertEqual(report["conversion_warnings"]["warnings"], [])
        self.assertEqual(
            report["office_structure"]["source"]["shape_count"],
            1,
        )
        self.assertEqual(
            report["office_structure"]["output"]["shape_count"],
            1,
        )
        self.assertEqual(report["office_structure"]["differences"], [])
