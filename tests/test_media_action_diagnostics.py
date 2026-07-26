from __future__ import annotations

import json
import os
from pathlib import Path
import struct
import subprocess
import sys
import tempfile
import unittest

from ppt2pptx import convert
from ppt2pptx.ppt import (
    RT_INTERACTIVE_INFO,
    RT_INTERACTIVE_INFO_ATOM,
    RT_OFFICEART_CLIENT_DATA,
    RT_OFFICEART_SP_CONTAINER,
    detect_lossy_features,
)


ROOT = Path(__file__).resolve().parents[1]
FIXTURE = ROOT / "tests" / "fixtures" / "visual_video.ppt"
COMPARE = ROOT / "scripts" / "compare_powerpoint_visual.py"


def rec(kind: int, payload: bytes = b"", version: int = 0xF) -> bytes:
    return struct.pack("<HHI", version, kind, len(payload)) + payload


def interactive_shape_document(action: int) -> bytes:
    atom_payload = b"\0" * 8 + bytes((action,)) + b"\0" * 7
    atom = rec(RT_INTERACTIVE_INFO_ATOM, atom_payload, version=0)
    interactive = rec(RT_INTERACTIVE_INFO, atom)
    client_data = rec(RT_OFFICEART_CLIENT_DATA, interactive)
    shape = rec(RT_OFFICEART_SP_CONTAINER, client_data)
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


class MediaActionDiagnosticTests(unittest.TestCase):
    def test_detects_one_slide_bound_media_action(self):
        features = {
            item.code: item
            for item in detect_lossy_features(interactive_shape_document(0x06))
        }

        self.assertEqual(set(features), {"MEDIA_ACTION_OMITTED"})
        feature = features["MEDIA_ACTION_OMITTED"]
        self.assertEqual(feature.count, 1)
        self.assertEqual(
            feature.record_types,
            (
                RT_INTERACTIVE_INFO,
                RT_INTERACTIVE_INFO_ATOM,
                RT_OFFICEART_SP_CONTAINER,
                RT_OFFICEART_CLIENT_DATA,
            ),
        )
        self.assertEqual(len(feature.locations), 1)
        self.assertEqual(feature.locations[0].slide_index, 1)
        self.assertEqual(feature.locations[0].record_type, RT_INTERACTIVE_INFO_ATOM)
        self.assertEqual(feature.locations[0].object_kind, "media")

    def test_ignores_non_media_interactive_action(self):
        self.assertEqual(detect_lossy_features(interactive_shape_document(0x04)), ())

    def test_controlled_fixture_reports_one_media_object(self):
        self.assertTrue(FIXTURE.is_file(), f"missing fixture: {FIXTURE}")
        with tempfile.TemporaryDirectory() as directory:
            result = convert(FIXTURE, Path(directory) / "video.pptx")

        warnings = {item["code"]: item for item in result.report.warnings}
        self.assertEqual(set(warnings), {"MEDIA_ACTION_OMITTED"})
        warning = warnings["MEDIA_ACTION_OMITTED"]
        self.assertEqual(warning["count"], 1)
        self.assertEqual(
            warning["record_types"],
            [
                RT_INTERACTIVE_INFO,
                RT_INTERACTIVE_INFO_ATOM,
                RT_OFFICEART_SP_CONTAINER,
                RT_OFFICEART_CLIENT_DATA,
            ],
        )
        self.assertEqual(len(warning["locations"]), 1)
        self.assertEqual(warning["locations"][0]["slide_index"], 1)
        self.assertEqual(warning["locations"][0]["record_offset"], 222620)
        self.assertEqual(warning["locations"][0]["object_kind"], "media")


@unittest.skipUnless(powerpoint_available(), "Microsoft PowerPoint COM is required")
class MediaActionPowerPointVisualTests(unittest.TestCase):
    def test_media_preview_is_identical_and_playback_loss_is_reported(self):
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
            report = json.loads(
                (evidence / "report.json").read_text(encoding="utf-8")
            )

        self.assertEqual(report["provider"], "office")
        self.assertTrue(report["powerpoint_version"])
        self.assertEqual(report["slide_count_source"], 1)
        self.assertEqual(report["slide_count_output"], 1)
        self.assertEqual(report["slide_width_pt"], 720.0)
        self.assertEqual(report["slide_height_pt"], 540.0)
        self.assertEqual(report["hidden_source"], [False])
        self.assertEqual(report["hidden_output"], [False])
        self.assertEqual(report["hard_differences"], [])
        self.assertEqual(report["summary"]["mean_mae"], 0.0)
        self.assertEqual(report["summary"]["mean_rmse"], 0.0)
        self.assertEqual(report["summary"]["mean_changed_pixel_ratio"], 0.0)
        self.assertEqual(report["summary"]["mean_ssim"], 1.0)

        structure = report["office_structure"]
        self.assertEqual(structure["source"]["media_count"], 1)
        self.assertEqual(structure["output"]["media_count"], 0)
        self.assertEqual(structure["source"]["picture_count"], 1)
        self.assertEqual(structure["output"]["picture_count"], 1)
        self.assertEqual(
            [(item["slide_index"], item["field"]) for item in structure["differences"]],
            [(1, "media_count")],
        )
        warnings = {
            item["code"]: item
            for item in report["conversion_warnings"]["warnings"]
        }
        self.assertEqual(set(warnings), {"MEDIA_ACTION_OMITTED"})
        warning = warnings["MEDIA_ACTION_OMITTED"]
        self.assertEqual(warning["count"], 1)
        self.assertEqual(warning["locations"][0]["slide_index"], 1)
        self.assertEqual(warning["locations"][0]["object_kind"], "media")


if __name__ == "__main__":
    unittest.main()
