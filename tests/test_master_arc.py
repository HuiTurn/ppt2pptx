from __future__ import annotations

import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import zipfile

from ppt2pptx.converter import convert


ROOT = Path(__file__).resolve().parents[1]
FIXTURE = ROOT / "tests" / "fixtures" / "visual_master_arc.ppt"
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


class MasterArcTests(unittest.TestCase):
    def test_legacy_arc_uses_editable_preset_geometry(self):
        self.assertTrue(FIXTURE.is_file(), f"missing fixture: {FIXTURE}")
        with tempfile.TemporaryDirectory() as directory:
            result = convert(FIXTURE, Path(directory) / "master-arc.pptx")
            with zipfile.ZipFile(result.output_path) as archive:
                xml = archive.read("ppt/slides/slide1.xml").decode()

        shapes = result.presentation.slides[0].shapes
        self.assertEqual(len(shapes), 1)
        self.assertEqual(shapes[0].preset, "arc")
        self.assertIsNone(shapes[0].path)
        self.assertTrue(shapes[0].from_master)
        self.assertIn('<a:prstGeom prst="arc">', xml)
        self.assertNotIn("<a:custGeom>", xml)
        self.assertEqual(result.report.warnings, [])


@unittest.skipUnless(
    powerpoint_available(), "Microsoft PowerPoint COM is required"
)
class MasterArcPowerPointVisualTests(unittest.TestCase):
    def test_native_arc_improves_bilateral_visual(self):
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
        self.assertLess(report["summary"]["mean_mae"], 0.007)
        self.assertLess(report["summary"]["mean_rmse"], 0.92)
        self.assertLess(
            report["summary"]["mean_changed_pixel_ratio"],
            0.0001,
        )
        self.assertGreater(report["summary"]["mean_ssim"], 0.986)
        self.assertEqual(report["conversion_warnings"]["warnings"], [])
        self.assertEqual(
            report["office_structure"]["source"]["shape_count"], 0
        )
        self.assertEqual(
            report["office_structure"]["output"]["shape_count"], 1
        )


if __name__ == "__main__":
    unittest.main()
