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
    RT_EXTERNAL_OBJECT_REF_ATOM,
    RT_EXTERNAL_OLE_CONTROL,
    RT_EXTERNAL_OLE_EMBED,
    RT_EXTERNAL_OLE_LINK,
    RT_EXTERNAL_OLE_OBJECT_ATOM,
    detect_lossy_features,
)


ROOT = Path(__file__).resolve().parents[1]
FIXTURE = ROOT / "tests" / "fixtures" / "visual_ole.ppt"
COMPARE = ROOT / "scripts" / "compare_powerpoint_visual.py"


def rec(kind: int, payload: bytes = b"", version: int = 0xF, instance: int = 0) -> bytes:
    return struct.pack("<HHI", (instance << 4) | version, kind, len(payload)) + payload


def ole_container(container_type: int, ole_type: int, ex_obj_id: int) -> bytes:
    atom = rec(
        RT_EXTERNAL_OLE_OBJECT_ATOM,
        struct.pack("<6I", 1, ole_type, ex_obj_id, 0, ex_obj_id + 20, 0),
        version=1,
    )
    return rec(container_type, atom)


def slide_with_ole_reference(ex_obj_id: int) -> bytes:
    reference = rec(
        RT_EXTERNAL_OBJECT_REF_ATOM,
        struct.pack("<I", ex_obj_id),
        version=0,
    )
    return rec(1006, reference)


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


class OleDiagnosticTests(unittest.TestCase):
    def test_collapses_storage_records_into_one_slide_bound_object(self):
        document = rec(
            1000,
            ole_container(RT_EXTERNAL_OLE_EMBED, 0, 7)
            + slide_with_ole_reference(7),
        )

        features = {item.code: item for item in detect_lossy_features(document)}

        self.assertEqual(set(features), {"EMBEDDED_OLE_OMITTED"})
        feature = features["EMBEDDED_OLE_OMITTED"]
        self.assertEqual(feature.count, 1)
        self.assertEqual(
            feature.record_types,
            (RT_EXTERNAL_OBJECT_REF_ATOM, RT_EXTERNAL_OLE_OBJECT_ATOM),
        )
        self.assertEqual(len(feature.locations), 1)
        self.assertEqual(feature.locations[0].slide_index, 1)
        self.assertEqual(feature.locations[0].record_type, RT_EXTERNAL_OBJECT_REF_ATOM)
        self.assertEqual(feature.locations[0].object_kind, "ole")

    def test_distinguishes_linked_ole_and_activex_control(self):
        document = rec(
            1000,
            ole_container(RT_EXTERNAL_OLE_LINK, 1, 8)
            + ole_container(RT_EXTERNAL_OLE_CONTROL, 2, 9)
            + slide_with_ole_reference(8)
            + slide_with_ole_reference(9),
        )

        features = {item.code: item for item in detect_lossy_features(document)}

        linked = features["LINKED_OLE_OMITTED"]
        control = features["ACTIVEX_CONTROL_OMITTED"]
        self.assertEqual(linked.count, 1)
        self.assertEqual(linked.locations[0].slide_index, 1)
        self.assertEqual(linked.locations[0].object_kind, "linked_ole")
        self.assertEqual(control.count, 1)
        self.assertEqual(control.locations[0].slide_index, 2)
        self.assertEqual(control.locations[0].object_kind, "activex_control")

    def test_controlled_fixture_reports_one_ole_object_on_slide_one(self):
        self.assertTrue(FIXTURE.is_file(), f"missing fixture: {FIXTURE}")
        with tempfile.TemporaryDirectory() as directory:
            result = convert(FIXTURE, Path(directory) / "visual_ole.pptx")

        warnings = {item["code"]: item for item in result.report.warnings}
        warning = warnings["EMBEDDED_OLE_OMITTED"]
        self.assertEqual(warning["count"], 1)
        self.assertEqual(len(warning["locations"]), 1)
        self.assertEqual(warning["locations"][0]["slide_index"], 1)
        self.assertEqual(warning["locations"][0]["record_type"], RT_EXTERNAL_OBJECT_REF_ATOM)
        self.assertEqual(warning["locations"][0]["object_kind"], "ole")


@unittest.skipUnless(powerpoint_available(), "Microsoft PowerPoint COM is required")
class OlePowerPointVisualTests(unittest.TestCase):
    def test_controlled_ole_fixture_has_exact_visual_and_diagnostic_evidence(self):
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
        self.assertEqual(report["hard_differences"], [])
        self.assertEqual(report["summary"]["mean_mae"], 0.0)
        self.assertEqual(report["summary"]["mean_rmse"], 0.0)
        self.assertEqual(report["summary"]["mean_changed_pixel_ratio"], 0.0)
        self.assertEqual(report["summary"]["mean_ssim"], 1.0)
        self.assertEqual(report["office_structure"]["source"]["ole_count"], 1)
        self.assertEqual(report["office_structure"]["output"]["ole_count"], 0)
        self.assertEqual(report["office_structure"]["output"]["picture_count"], 1)
        warnings = {
            item["code"]: item
            for item in report["conversion_warnings"]["warnings"]
        }
        warning = warnings["EMBEDDED_OLE_OMITTED"]
        self.assertEqual(warning["count"], 1)
        self.assertEqual(warning["locations"][0]["slide_index"], 1)
        self.assertEqual(
            warning["locations"][0]["record_type"],
            RT_EXTERNAL_OBJECT_REF_ATOM,
        )


if __name__ == "__main__":
    unittest.main()
