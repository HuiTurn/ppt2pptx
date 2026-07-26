from __future__ import annotations

import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import zipfile

from ppt2pptx import convert


ROOT = Path(__file__).resolve().parents[1]
FIXTURE = ROOT / "tests" / "fixtures" / "visual_equation_ole.ppt"
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


class EquationOleTests(unittest.TestCase):
    def test_preserves_editable_mathtype_storage_and_preview(self):
        with tempfile.TemporaryDirectory() as directory:
            result = convert(FIXTURE, Path(directory) / "equation.pptx")
            with zipfile.ZipFile(result.output_path) as archive:
                xml = archive.read("ppt/slides/slide1.xml").decode()
                embedded = archive.read("ppt/embeddings/oleObject1.bin")

        self.assertEqual(result.report.warnings, [])
        self.assertEqual(result.presentation.preserved_external_object_ids, {2})
        picture = result.presentation.slides[0].pictures[0]
        self.assertEqual(picture.embedded_object_prog_id, "Equation.DSMT4")
        self.assertEqual(picture.embedded_object_name, "Equation")
        self.assertEqual(embedded, picture.embedded_object_data)
        self.assertTrue(embedded.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"))
        self.assertIn(b"Equation.DSMT4", embedded)
        self.assertIn("<p:oleObj", xml)
        self.assertIn('progId="Equation.DSMT4"', xml)
        self.assertIn("<p:pic>", xml)


@unittest.skipUnless(
    powerpoint_available(), "Microsoft PowerPoint COM is required"
)
class EquationOlePowerPointVisualTests(unittest.TestCase):
    def test_equation_remains_one_pixel_exact_ole_object(self):
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
        self.assertEqual(report["office_structure"]["source"]["ole_count"], 1)
        self.assertEqual(report["office_structure"]["output"]["ole_count"], 1)
        self.assertEqual(report["conversion_warnings"]["warnings"], [])


if __name__ == "__main__":
    unittest.main()
