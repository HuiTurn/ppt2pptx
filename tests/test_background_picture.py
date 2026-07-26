from __future__ import annotations

import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import zipfile

from ppt2pptx.converter import convert, inspect_ppt


ROOT = Path(__file__).resolve().parents[1]
FIXTURE = ROOT / "tests" / "fixtures" / "visual_background_picture.ppt"
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


class BackgroundPictureTests(unittest.TestCase):
    def test_parses_picture_fill_and_writes_editable_slide_background(self):
        self.assertTrue(FIXTURE.is_file(), f"missing fixture: {FIXTURE}")
        inspected = inspect_ppt(FIXTURE)
        background = inspected["slides"][0]["background_image"]
        self.assertEqual(background["extension"], "png")
        self.assertEqual(background["content_type"], "image/png")
        self.assertGreater(background["byte_count"], 100)
        self.assertEqual(inspected["picture_count"], 0)
        self.assertEqual(inspected["text_box_count"], 1)

        with tempfile.TemporaryDirectory() as directory:
            result = convert(FIXTURE, Path(directory) / "background.pptx")
            with zipfile.ZipFile(result.output_path) as archive:
                slide_xml = archive.read("ppt/slides/slide1.xml").decode()
                rels_xml = archive.read(
                    "ppt/slides/_rels/slide1.xml.rels"
                ).decode()
                media = [
                    name for name in archive.namelist()
                    if name.startswith("ppt/media/")
                ]

        self.assertEqual(result.report.warnings, [])
        self.assertIn("<p:bg><p:bgPr><a:blipFill", slide_xml)
        self.assertIn('<a:blip r:embed="rId2"/>', slide_xml)
        self.assertIn("<a:stretch><a:fillRect/></a:stretch>", slide_xml)
        self.assertNotIn("<p:pic>", slide_xml)
        self.assertIn(
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships/image"',
            rels_xml,
        )
        self.assertEqual(media, ["ppt/media/image1.png"])


@unittest.skipUnless(powerpoint_available(), "Microsoft PowerPoint COM is required")
class BackgroundPicturePowerPointVisualTests(unittest.TestCase):
    def test_picture_background_has_powerpoint_bilateral_visual_evidence(self):
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
        self.assertEqual(report["hard_differences"], [])
        self.assertEqual(report["slide_count_source"], 1)
        self.assertEqual(report["slide_count_output"], 1)
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
