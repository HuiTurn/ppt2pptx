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
FIXTURE = ROOT / "tests" / "fixtures" / "visual_master_shape_z_order.ppt"
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


class MasterShapeZOrderTests(unittest.TestCase):
    def test_master_shape_precedes_slide_shape_in_drawingml(self):
        self.assertTrue(FIXTURE.is_file(), f"missing fixture: {FIXTURE}")
        with tempfile.TemporaryDirectory() as directory:
            result = convert(FIXTURE, Path(directory) / "master-z-order.pptx")
            with zipfile.ZipFile(result.output_path) as archive:
                root = ElementTree.fromstring(
                    archive.read("ppt/slides/slide1.xml")
                )

        shapes = result.presentation.slides[0].shapes
        self.assertEqual(len(shapes), 2)
        self.assertEqual(
            [(shape.preset, shape.from_master) for shape in shapes],
            [("line", True), ("rect", False)],
        )
        ns = {
            "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
        }
        shape_ids = [
            int(value.attrib["id"])
            for value in root.findall(".//p:sp/p:nvSpPr/p:cNvPr", ns)
        ]
        self.assertEqual(shape_ids, [2, 3])
        self.assertEqual(result.report.warnings, [])


@unittest.skipUnless(
    powerpoint_available(), "Microsoft PowerPoint COM is required"
)
class MasterShapeZOrderPowerPointVisualTests(unittest.TestCase):
    def test_master_line_remains_behind_slide_fill(self):
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
            report["office_structure"]["source"]["shape_count"], 1
        )
        self.assertEqual(
            report["office_structure"]["output"]["shape_count"], 2
        )


if __name__ == "__main__":
    unittest.main()
