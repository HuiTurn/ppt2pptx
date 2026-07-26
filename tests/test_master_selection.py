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

from ppt2pptx.cfb import CompoundFile
from ppt2pptx.converter import convert
from ppt2pptx.ppt import (
    CONTAINER_VERSION,
    RT_DOCUMENT,
    RT_MAIN_MASTER,
    _master_entries,
    _slide_entries,
    _slide_master_id,
    persist_directory,
    records,
)


ROOT = Path(__file__).resolve().parents[1]
FIXTURE = ROOT / "tests" / "fixtures" / "visual_master_selection.ppt"
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


class MasterSelectionTests(unittest.TestCase):
    def test_uses_slide_master_id_instead_of_first_master(self):
        self.assertTrue(FIXTURE.is_file(), f"missing fixture: {FIXTURE}")
        stream = CompoundFile.from_path(FIXTURE).open_stream(
            "PowerPoint Document"
        )
        roots = list(records(stream))
        document = next(
            record
            for record in reversed(roots)
            if record.type == RT_DOCUMENT
            and record.version == CONTAINER_VERSION
        )
        mapping = persist_directory(stream)
        master_entries = _master_entries(document)
        self.assertGreater(len(master_entries), 1)
        slide_reference, _slide_id = _slide_entries(document)[0]
        slide = next(records(stream, mapping[slide_reference]))
        selected_master_id = _slide_master_id(slide)
        self.assertIsNotNone(selected_master_id)
        self.assertNotEqual(selected_master_id, master_entries[0][1])
        selected_reference = next(
            reference
            for reference, master_id in master_entries
            if master_id == selected_master_id
        )
        selected_master = next(records(stream, mapping[selected_reference]))
        self.assertEqual(selected_master.type, RT_MAIN_MASTER)

        with tempfile.TemporaryDirectory() as directory:
            result = convert(FIXTURE, Path(directory) / "master.pptx")
            with zipfile.ZipFile(result.output_path) as archive:
                root = ElementTree.fromstring(
                    archive.read("ppt/slides/slide1.xml")
                )

        slide_model = result.presentation.slides[0]
        self.assertEqual(slide_model.background_color, "0000FF")
        self.assertEqual(len(slide_model.shapes), 0)
        self.assertEqual(len(slide_model.text_boxes), 0)
        self.assertEqual(len(slide_model.pictures), 0)
        self.assertEqual(result.report.warnings, [])
        ns = {
            "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
        }
        color = root.find(".//p:bg//a:srgbClr", ns)
        self.assertIsNotNone(color)
        assert color is not None
        self.assertEqual(color.attrib.get("val"), "0000FF")


@unittest.skipUnless(powerpoint_available(), "Microsoft PowerPoint COM is required")
class MasterSelectionPowerPointVisualTests(unittest.TestCase):
    def test_selected_master_has_exact_bilateral_visual(self):
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
            0,
        )
        self.assertEqual(
            report["office_structure"]["output"]["shape_count"],
            0,
        )
        self.assertEqual(report["office_structure"]["differences"], [])
