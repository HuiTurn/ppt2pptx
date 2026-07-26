from __future__ import annotations

import io
import json
import os
from pathlib import Path
import struct
import subprocess
import sys
import tempfile
import unittest
import zipfile

from ppt2pptx import convert
from ppt2pptx.ppt import (
    RT_OFFICEART_SP_CONTAINER,
    RT_OFFICEART_TERTIARY_FOPT,
    detect_lossy_features,
)


ROOT = Path(__file__).resolve().parents[1]
FIXTURE = ROOT / "tests" / "fixtures" / "visual_smartart.ppt"
COMPARE = ROOT / "scripts" / "compare_powerpoint_visual.py"


def rec(kind: int, payload: bytes = b"", version: int = 0xF, instance: int = 0) -> bytes:
    return struct.pack("<HHI", (instance << 4) | version, kind, len(payload)) + payload


def metro_blob_document(*members: str) -> bytes:
    package = io.BytesIO()
    with zipfile.ZipFile(package, "w", zipfile.ZIP_STORED) as archive:
        for member in members:
            archive.writestr(member, b"<root/>")
    blob = package.getvalue()
    metro_blob = struct.pack("<HI", 0xC3A9, len(blob)) + blob
    tertiary = rec(
        RT_OFFICEART_TERTIARY_FOPT,
        metro_blob,
        version=3,
        instance=1,
    )
    shape = rec(RT_OFFICEART_SP_CONTAINER, tertiary)
    return rec(1000, rec(1006, shape))


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
        app = None
        pythoncom.CoUninitialize()


class SmartArtDiagnosticTests(unittest.TestCase):
    def test_detects_one_slide_bound_smartart_metro_blob(self):
        features = {
            item.code: item
            for item in detect_lossy_features(
                metro_blob_document(
                    "[Content_Types].xml",
                    "drs/diagrams/data1.xml",
                )
            )
        }

        self.assertEqual(set(features), {"DIAGRAM_OR_SMARTART_OMITTED"})
        feature = features["DIAGRAM_OR_SMARTART_OMITTED"]
        self.assertEqual(feature.count, 1)
        self.assertEqual(
            feature.record_types,
            (RT_OFFICEART_SP_CONTAINER, RT_OFFICEART_TERTIARY_FOPT),
        )
        self.assertEqual(len(feature.locations), 1)
        self.assertEqual(feature.locations[0].slide_index, 1)
        self.assertEqual(
            feature.locations[0].record_type,
            RT_OFFICEART_TERTIARY_FOPT,
        )
        self.assertEqual(feature.locations[0].object_kind, "smartart")

    def test_ignores_non_diagram_metro_blob(self):
        document = metro_blob_document(
            "[Content_Types].xml",
            "drs/shapes/shape1.xml",
        )
        self.assertEqual(detect_lossy_features(document), ())

    def test_controlled_fixture_reports_one_smartart_object(self):
        self.assertTrue(FIXTURE.is_file(), f"missing fixture: {FIXTURE}")
        with tempfile.TemporaryDirectory() as directory:
            result = convert(FIXTURE, Path(directory) / "smartart.pptx")

        warnings = {item["code"]: item for item in result.report.warnings}
        self.assertEqual(set(warnings), {"DIAGRAM_OR_SMARTART_OMITTED"})
        warning = warnings["DIAGRAM_OR_SMARTART_OMITTED"]
        self.assertEqual(warning["count"], 1)
        self.assertEqual(
            warning["record_types"],
            [RT_OFFICEART_SP_CONTAINER, RT_OFFICEART_TERTIARY_FOPT],
        )
        self.assertEqual(len(warning["locations"]), 1)
        self.assertEqual(warning["locations"][0]["slide_index"], 1)
        self.assertEqual(warning["locations"][0]["object_kind"], "smartart")


@unittest.skipUnless(powerpoint_available(), "Microsoft PowerPoint COM is required")
class SmartArtPowerPointVisualTests(unittest.TestCase):
    def test_smartart_preview_has_bounded_visual_loss_and_exact_diagnostic(self):
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
                ],
                capture_output=True,
                text=True,
                env=env,
                timeout=180,
                check=False,
            )
            self.assertEqual(
                completed.returncode,
                0,
                msg=f"stdout:\n{completed.stdout}\nstderr:\n{completed.stderr}",
            )
            report = json.loads((evidence / "report.json").read_text(encoding="utf-8"))

        self.assertEqual(report["provider"], "office")
        self.assertEqual(report["slide_count_source"], 1)
        self.assertEqual(report["slide_count_output"], 1)
        self.assertEqual(report["slide_width_pt"], 960.0)
        self.assertEqual(report["slide_height_pt"], 540.0)
        self.assertEqual(report["hidden_source"], [False])
        self.assertEqual(report["hidden_output"], [False])
        self.assertEqual(report["hard_differences"], [])
        self.assertLessEqual(report["summary"]["mean_mae"], 0.33)
        self.assertLessEqual(report["summary"]["mean_rmse"], 4.7)
        self.assertLessEqual(
            report["summary"]["mean_changed_pixel_ratio"],
            0.012,
        )
        self.assertGreaterEqual(report["summary"]["mean_ssim"], 0.997)

        structure = report["office_structure"]
        self.assertEqual(structure["source"]["smartart_count"], 1)
        self.assertEqual(structure["output"]["smartart_count"], 0)
        self.assertEqual(structure["source"]["picture_count"], 0)
        self.assertEqual(structure["output"]["picture_count"], 1)
        self.assertEqual(
            [(item["slide_index"], item["field"]) for item in structure["differences"]],
            [(1, "picture_count"), (1, "smartart_count")],
        )
        warnings = {
            item["code"]: item
            for item in report["conversion_warnings"]["warnings"]
        }
        self.assertEqual(set(warnings), {"DIAGRAM_OR_SMARTART_OMITTED"})
        warning = warnings["DIAGRAM_OR_SMARTART_OMITTED"]
        self.assertEqual(warning["count"], 1)
        self.assertEqual(warning["locations"][0]["slide_index"], 1)
        self.assertEqual(warning["locations"][0]["object_kind"], "smartart")


if __name__ == "__main__":
    unittest.main()
