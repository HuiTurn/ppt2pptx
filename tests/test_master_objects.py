from __future__ import annotations

import json
import os
from pathlib import Path
import struct
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
    RT_SLIDE_ATOM,
    _slide_entries,
    descendants,
    persist_directory,
    records,
)


ROOT = Path(__file__).resolve().parents[1]
FIXTURE = ROOT / "tests" / "fixtures" / "visual_master_objects_disabled.ppt"
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


class MasterObjectsTests(unittest.TestCase):
    def test_omits_master_shapes_when_slide_flag_is_clear(self):
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
        reference, _slide_id = _slide_entries(document)[0]
        slide = next(records(stream, persist_directory(stream)[reference]))
        slide_atom = next(
            child
            for child in descendants(slide)
            if child.type == RT_SLIDE_ATOM
        )
        slide_flags = struct.unpack_from("<H", slide_atom.payload, 20)[0]
        self.assertEqual(slide_flags & 0x0001, 0)

        with tempfile.TemporaryDirectory() as directory:
            result = convert(FIXTURE, Path(directory) / "master-hidden.pptx")
            with zipfile.ZipFile(result.output_path) as archive:
                root = ElementTree.fromstring(
                    archive.read("ppt/slides/slide1.xml")
                )

        slide_model = result.presentation.slides[0]
        self.assertEqual(len(slide_model.shapes), 1)
        self.assertEqual(len(slide_model.text_boxes), 0)
        self.assertEqual(len(slide_model.pictures), 0)
        self.assertEqual(result.report.warnings, [])
        ns = {
            "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
        }
        self.assertEqual(len(root.findall(".//p:sp", ns)), 1)


@unittest.skipUnless(powerpoint_available(), "Microsoft PowerPoint COM is required")
class MasterObjectsPowerPointVisualTests(unittest.TestCase):
    def test_disabled_master_objects_have_exact_bilateral_visual(self):
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
        self.assertEqual(
            report["office_structure"]["differences"],
            [],
        )
