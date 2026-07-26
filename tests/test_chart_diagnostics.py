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
    RT_CSTRING,
    RT_EXTERNAL_OBJECT_REF_ATOM,
    RT_EXTERNAL_OLE_EMBED,
    RT_EXTERNAL_OLE_OBJECT_ATOM,
    detect_lossy_features,
)


ROOT = Path(__file__).resolve().parents[1]
FIXTURE = ROOT / "tests" / "fixtures" / "visual_chart.ppt"
COMPARE = ROOT / "scripts" / "compare_powerpoint_visual.py"


def rec(kind: int, payload: bytes = b"", version: int = 0xF, instance: int = 0) -> bytes:
    return struct.pack("<HHI", (instance << 4) | version, kind, len(payload)) + payload


def chart_document() -> bytes:
    ex_obj_id = 7
    atom = rec(
        RT_EXTERNAL_OLE_OBJECT_ATOM,
        struct.pack("<6I", 1, 0, ex_obj_id, 4, 27, 0),
        version=1,
    )
    progid = rec(
        RT_CSTRING,
        "MSGraph.Chart.8".encode("utf-16le"),
        version=0,
        instance=2,
    )
    container = rec(RT_EXTERNAL_OLE_EMBED, atom + progid)
    reference = rec(
        RT_EXTERNAL_OBJECT_REF_ATOM,
        struct.pack("<I", ex_obj_id),
        version=0,
    )
    return rec(1000, container + rec(1006, reference))


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


class ChartDiagnosticTests(unittest.TestCase):
    def test_classifies_embedded_chart_as_one_slide_bound_object(self):
        features = {item.code: item for item in detect_lossy_features(chart_document())}

        self.assertEqual(set(features), {"CHART_OMITTED"})
        feature = features["CHART_OMITTED"]
        self.assertEqual(feature.count, 1)
        self.assertEqual(
            feature.record_types,
            (RT_EXTERNAL_OBJECT_REF_ATOM, RT_CSTRING, RT_EXTERNAL_OLE_OBJECT_ATOM),
        )
        self.assertEqual(len(feature.locations), 1)
        self.assertEqual(feature.locations[0].slide_index, 1)
        self.assertEqual(feature.locations[0].record_type, RT_EXTERNAL_OBJECT_REF_ATOM)
        self.assertEqual(feature.locations[0].object_kind, "chart")

    def test_controlled_fixture_reports_chart_without_duplicate_ole_warning(self):
        self.assertTrue(FIXTURE.is_file(), f"missing fixture: {FIXTURE}")
        with tempfile.TemporaryDirectory() as directory:
            result = convert(FIXTURE, Path(directory) / "visual_chart.pptx")

        warnings = {item["code"]: item for item in result.report.warnings}
        self.assertEqual(set(warnings), {"CHART_OMITTED"})
        warning = warnings["CHART_OMITTED"]
        self.assertEqual(warning["count"], 1)
        self.assertEqual(len(warning["locations"]), 1)
        self.assertEqual(warning["locations"][0]["slide_index"], 1)
        self.assertEqual(warning["locations"][0]["record_type"], RT_EXTERNAL_OBJECT_REF_ATOM)
        self.assertEqual(warning["locations"][0]["object_kind"], "chart")


@unittest.skipUnless(powerpoint_available(), "Microsoft PowerPoint COM is required")
class ChartPowerPointVisualTests(unittest.TestCase):
    def test_legacy_chart_preview_has_exact_visual_and_object_diagnostic(self):
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
        self.assertEqual(set(warnings), {"CHART_OMITTED"})
        warning = warnings["CHART_OMITTED"]
        self.assertEqual(warning["count"], 1)
        self.assertEqual(warning["locations"][0]["slide_index"], 1)
        self.assertEqual(
            warning["locations"][0]["record_type"],
            RT_EXTERNAL_OBJECT_REF_ATOM,
        )


if __name__ == "__main__":
    unittest.main()
