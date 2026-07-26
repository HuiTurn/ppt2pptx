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
FIXTURE = ROOT / "tests" / "fixtures" / "visual_bullet_font.ppt"
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


class BulletFontTests(unittest.TestCase):
    def test_preserves_textpfexception_bullet_style(self):
        self.assertTrue(FIXTURE.is_file(), f"missing fixture: {FIXTURE}")
        inspected = inspect_ppt(FIXTURE)
        self.assertEqual(inspected["slide_count"], 1)
        boxes = inspected["slides"][0]["text_boxes"]
        self.assertEqual(len(boxes), 1)
        self.assertEqual(
            boxes[0]["paragraph_bullet_chars"],
            ("w", "n", "w"),
        )
        self.assertEqual(
            boxes[0]["paragraph_bullet_typefaces"],
            ("Wingdings", "Wingdings", "Wingdings"),
        )
        self.assertEqual(
            boxes[0]["paragraph_bullet_colors"],
            ("6F89F7", "40458C", "6F89F7"),
        )
        self.assertEqual(
            boxes[0]["paragraph_bullet_sizes"],
            (110, 60, 110),
        )

        with tempfile.TemporaryDirectory() as directory:
            result = convert(
                FIXTURE, Path(directory) / "bullet-font.pptx"
            )
            with zipfile.ZipFile(result.output_path) as archive:
                xml = archive.read("ppt/slides/slide1.xml").decode()

        self.assertEqual(
            xml.count('<a:buFont typeface="Wingdings"/>'), 3
        )
        self.assertEqual(xml.count('<a:buChar char="w"/>'), 2)
        self.assertEqual(xml.count('<a:buChar char="n"/>'), 1)
        self.assertEqual(
            xml.count('<a:buClr><a:srgbClr val="6F89F7"/></a:buClr>'),
            2,
        )
        self.assertEqual(
            xml.count('<a:buClr><a:srgbClr val="40458C"/></a:buClr>'),
            1,
        )
        self.assertEqual(xml.count('<a:buSzPct val="110000"/>'), 2)
        self.assertEqual(xml.count('<a:buSzPct val="60000"/>'), 1)
        self.assertEqual(result.report.warnings, [])


@unittest.skipUnless(
    powerpoint_available(), "Microsoft PowerPoint COM is required"
)
class BulletFontPowerPointVisualTests(unittest.TestCase):
    def test_wingdings_bullets_have_exact_bilateral_visual(self):
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
                msg=(
                    f"stdout:\n{completed.stdout}\n"
                    f"stderr:\n{completed.stderr}"
                ),
            )
            report = json.loads(
                (evidence / "report.json").read_text(encoding="utf-8")
            )

        self.assertEqual(report["provider"], "office")
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
            report["summary"]["mean_changed_pixel_ratio"], 0.0
        )
        self.assertEqual(report["summary"]["mean_ssim"], 1.0)
        self.assertEqual(report["conversion_warnings"]["warnings"], [])
        self.assertEqual(
            report["office_structure"]["source"]["text_shape_count"], 1
        )
        self.assertEqual(
            report["office_structure"]["output"]["text_shape_count"], 1
        )
        self.assertEqual(report["office_structure"]["differences"], [])


if __name__ == "__main__":
    unittest.main()
