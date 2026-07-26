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
    RT_OFFICEART_FOPT,
    _direct_children,
    _fopt_properties,
    _iter_sp_containers,
    _master_placeholder_text_anchors,
    _placeholder_placement_id,
    _slide_entries,
    persist_directory,
    records,
)


ROOT = Path(__file__).resolve().parents[1]
FIXTURE = ROOT / "tests" / "fixtures" / "visual_placeholder_anchor.ppt"
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


class PlaceholderAnchorTests(unittest.TestCase):
    def test_inherits_title_anchor_from_master_placeholder(self):
        self.assertTrue(FIXTURE.is_file(), f"missing fixture: {FIXTURE}")
        compound = CompoundFile.from_path(FIXTURE)
        stream = compound.open_stream("PowerPoint Document")
        roots = list(records(stream))
        document = next(
            record
            for record in reversed(roots)
            if record.type == RT_DOCUMENT
            and record.version == CONTAINER_VERSION
        )
        reference, _slide_id = _slide_entries(document)[0]
        slide = next(records(stream, persist_directory(stream)[reference]))
        title_shape = next(
            shape
            for shape, _space in _iter_sp_containers(slide)
            if _placeholder_placement_id(shape) == 13
        )
        fopt = next(
            (
                child
                for child in _direct_children(title_shape)
                if child.type == RT_OFFICEART_FOPT
            ),
            None,
        )
        self.assertNotIn(135, _fopt_properties(fopt) if fopt else {})
        masters = [record for record in roots if record.type == 1016]
        self.assertTrue(
            any(
                _master_placeholder_text_anchors(master).get(13) == 2
                for master in masters
            )
        )

        with tempfile.TemporaryDirectory() as directory:
            result = convert(FIXTURE, Path(directory) / "anchor.pptx")
            with zipfile.ZipFile(result.output_path) as archive:
                root = ElementTree.fromstring(
                    archive.read("ppt/slides/slide1.xml")
                )

        boxes = result.presentation.slides[0].text_boxes
        self.assertEqual(len(boxes), 1)
        self.assertEqual(boxes[0].vertical_anchor, "b")
        self.assertEqual(result.report.warnings, [])
        ns = {
            "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        }
        body_properties = root.find(".//a:bodyPr", ns)
        self.assertIsNotNone(body_properties)
        assert body_properties is not None
        self.assertEqual(body_properties.attrib.get("anchor"), "b")


@unittest.skipUnless(powerpoint_available(), "Microsoft PowerPoint COM is required")
class PlaceholderAnchorPowerPointVisualTests(unittest.TestCase):
    def test_inherited_title_anchor_has_exact_bilateral_visual(self):
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
