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
from ppt2pptx.cfb import CompoundFile


ROOT = Path(__file__).resolve().parents[1]
FIXTURE = ROOT / "tests" / "fixtures" / "visual_excel_ole.ppt"
CORPUS = ROOT / "tests" / "real_samples" / "testPPT_oleWorkbook.ppt"
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


class ExcelOleTests(unittest.TestCase):
    def test_preserves_editable_excel_storage_and_preview(self):
        with tempfile.TemporaryDirectory() as directory:
            result = convert(FIXTURE, Path(directory) / "worksheet.pptx")
            with zipfile.ZipFile(result.output_path) as archive:
                xml = archive.read("ppt/slides/slide1.xml").decode()
                embedded = archive.read("ppt/embeddings/oleObject1.bin")

        self.assertEqual(result.report.warnings, [])
        self.assertEqual(len(result.presentation.preserved_external_object_ids), 1)
        picture = result.presentation.slides[0].pictures[0]
        self.assertEqual(picture.embedded_object_prog_id, "Excel.Sheet.12")
        self.assertEqual(picture.embedded_object_name, "Worksheet")
        self.assertEqual(embedded, picture.embedded_object_data)
        compound = CompoundFile(embedded)
        self.assertIn("Package", compound.entries)
        self.assertTrue(compound.open_stream("Package").startswith(b"PK\x03\x04"))
        self.assertIn(b"Excel.Sheet.12", compound.open_stream("\x01CompObj"))
        self.assertIn("<p:oleObj", xml)
        self.assertIn('progId="Excel.Sheet.12"', xml)
        self.assertIn("<p:pic>", xml)

    @unittest.skipUnless(
        CORPUS.is_file(), "Apache POI OLE workbook sample is not downloaded"
    )
    def test_real_sample_removes_only_its_embedded_ole_warning(self):
        with tempfile.TemporaryDirectory() as directory:
            result = convert(CORPUS, Path(directory) / "workbook.pptx")

        self.assertEqual(result.report.warnings, [])
        self.assertEqual(len(result.presentation.preserved_external_object_ids), 1)
        embedded = [
            picture
            for slide in result.presentation.slides
            for picture in slide.pictures
            if picture.embedded_object_data is not None
        ]
        self.assertEqual(len(embedded), 1)
        self.assertEqual(embedded[0].embedded_object_prog_id, "Excel.Sheet.12")


@unittest.skipUnless(
    powerpoint_available(), "Microsoft PowerPoint COM is required"
)
class ExcelOlePowerPointVisualTests(unittest.TestCase):
    def test_worksheet_remains_one_pixel_exact_ole_object(self):
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
        self.assertEqual(report["office_structure"]["source"]["picture_count"], 0)
        self.assertEqual(report["office_structure"]["output"]["picture_count"], 0)
        self.assertEqual(report["conversion_warnings"]["warnings"], [])


if __name__ == "__main__":
    unittest.main()
