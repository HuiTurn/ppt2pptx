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

from ppt2pptx.converter import convert


ROOT = Path(__file__).resolve().parents[1]
FIXTURE = ROOT / "tests" / "fixtures" / "visual_zero_extent_lines.ppt"
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


class ZeroExtentLineTests(unittest.TestCase):
    def test_preserves_zero_axis_in_model_and_drawingml(self):
        self.assertTrue(FIXTURE.is_file(), f"missing fixture: {FIXTURE}")
        with tempfile.TemporaryDirectory() as directory:
            result = convert(FIXTURE, Path(directory) / "zero-lines.pptx")
            with zipfile.ZipFile(result.output_path) as archive:
                root = ElementTree.fromstring(
                    archive.read("ppt/slides/slide1.xml")
                )

        lines = result.presentation.slides[0].shapes
        self.assertEqual(len(lines), 54)
        self.assertTrue(all(shape.preset == "line" for shape in lines))
        self.assertTrue(
            all(
                (shape.width == 0 and shape.height > 0)
                or (shape.height == 0 and shape.width > 0)
                for shape in lines
            )
        )
        self.assertEqual(result.report.warnings, [])

        ns = {
            "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
        }
        extents = root.findall(".//p:sp/p:spPr/a:xfrm/a:ext", ns)
        self.assertEqual(len(extents), 54)
        self.assertTrue(
            all(
                extent.attrib["cx"] == "0" or extent.attrib["cy"] == "0"
                for extent in extents
            )
        )


@unittest.skipUnless(powerpoint_available(), "Microsoft PowerPoint COM is required")
class ZeroExtentLinePowerPointVisualTests(unittest.TestCase):
    def test_master_grid_has_bounded_bilateral_visual_difference(self):
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
        self.assertLess(report["summary"]["mean_mae"], 0.001)
        self.assertLess(report["summary"]["mean_rmse"], 0.1)
        self.assertLess(
            report["summary"]["mean_changed_pixel_ratio"],
            0.00001,
        )
        self.assertGreater(report["summary"]["mean_ssim"], 0.9999)
        self.assertEqual(report["conversion_warnings"]["warnings"], [])
        self.assertEqual(
            report["office_structure"]["source"]["shape_count"],
            0,
        )
        self.assertEqual(
            report["office_structure"]["output"]["shape_count"],
            54,
        )
