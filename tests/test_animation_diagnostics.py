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
    RT_ANIMATION_INFO,
    RT_ANIMATION_INFO_ATOM,
    RT_BINARY_TAG_DATA_BLOB,
    RT_BUILD_ATOM,
    RT_BUILD_LIST,
    RT_PARA_BUILD,
    RT_PARA_BUILD_ATOM,
    RT_TIME_EXT_TIME_NODE,
    RT_TIME_NODE,
    RT_TIME_PROPERTY_LIST,
    RT_TIME_VARIANT,
    RT_VISUAL_SHAPE_ATOM,
    detect_lossy_features,
)


ROOT = Path(__file__).resolve().parents[1]
FIXTURE = ROOT / "tests" / "fixtures" / "visual_animation_object.ppt"
COMPARE = ROOT / "scripts" / "compare_powerpoint_visual.py"


def rec(kind: int, payload: bytes = b"", version: int = 0xF, instance: int = 0) -> bytes:
    return struct.pack("<HHI", (instance << 4) | version, kind, len(payload)) + payload


def animation_document() -> bytes:
    atom = rec(RT_ANIMATION_INFO_ATOM, bytes(28), version=0)
    animation = rec(RT_ANIMATION_INFO, atom)
    return rec(1000, rec(1006, animation))


def timeline_document() -> bytes:
    effect_type = rec(
        RT_TIME_VARIANT,
        b"\x01" + struct.pack("<i", 1),
        version=0,
        instance=11,
    )
    properties = rec(RT_TIME_PROPERTY_LIST, effect_type)
    time_node = rec(RT_TIME_NODE, bytes(32), version=0)
    visual_shape = rec(
        RT_VISUAL_SHAPE_ATOM,
        struct.pack("<IIIII", 0, 1, 42, 0xFFFFFFFF, 0xFFFFFFFF),
        version=0,
    )
    effect = rec(
        RT_TIME_EXT_TIME_NODE,
        time_node + properties + visual_shape,
        instance=1,
    )
    build_atom = rec(
        RT_BUILD_ATOM,
        struct.pack("<IIII", 1, 0, 42, 1),
        version=0,
    )
    paragraph_atom = rec(
        RT_PARA_BUILD_ATOM,
        struct.pack("<IIII", 3, 1, 0x00010000, 0),
        version=1,
    )
    paragraph_build = rec(RT_PARA_BUILD, build_atom + paragraph_atom)
    build_list = rec(RT_BUILD_LIST, paragraph_build)
    extension = rec(
        RT_BINARY_TAG_DATA_BLOB,
        effect + build_list,
        version=0,
    )
    return rec(1000, rec(1006, extension))


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


class AnimationDiagnosticTests(unittest.TestCase):
    def test_collapses_animation_container_and_atom_into_one_object(self):
        features = {item.code: item for item in detect_lossy_features(animation_document())}

        self.assertEqual(set(features), {"ANIMATION_OMITTED"})
        feature = features["ANIMATION_OMITTED"]
        self.assertEqual(feature.count, 1)
        self.assertEqual(
            feature.record_types,
            (RT_ANIMATION_INFO_ATOM, RT_ANIMATION_INFO),
        )
        self.assertEqual(len(feature.locations), 1)
        self.assertEqual(feature.locations[0].slide_index, 1)
        self.assertEqual(feature.locations[0].record_type, RT_ANIMATION_INFO)
        self.assertEqual(feature.locations[0].object_kind, "animation")

    def test_detects_effect_node_inside_pp10_binary_tag(self):
        features = {item.code: item for item in detect_lossy_features(timeline_document())}

        self.assertEqual(set(features), {"ANIMATION_OMITTED"})
        feature = features["ANIMATION_OMITTED"]
        self.assertEqual(feature.count, 1)
        self.assertEqual(
            feature.record_types,
            (
                RT_BUILD_LIST,
                RT_BUILD_ATOM,
                RT_PARA_BUILD,
                RT_PARA_BUILD_ATOM,
                RT_TIME_NODE,
                RT_TIME_PROPERTY_LIST,
                RT_TIME_VARIANT,
                RT_TIME_EXT_TIME_NODE,
            ),
        )
        self.assertEqual(len(feature.locations), 1)
        self.assertEqual(feature.locations[0].slide_index, 1)
        self.assertEqual(feature.locations[0].record_type, RT_TIME_EXT_TIME_NODE)

    def test_controlled_fixture_reports_one_shape_animation(self):
        self.assertTrue(FIXTURE.is_file(), f"missing fixture: {FIXTURE}")
        with tempfile.TemporaryDirectory() as directory:
            result = convert(FIXTURE, Path(directory) / "visual_animation_object.pptx")

        warnings = {item["code"]: item for item in result.report.warnings}
        self.assertEqual(set(warnings), {"ANIMATION_OMITTED"})
        warning = warnings["ANIMATION_OMITTED"]
        self.assertEqual(warning["count"], 1)
        self.assertEqual(
            warning["record_types"],
            [
                RT_ANIMATION_INFO_ATOM,
                RT_ANIMATION_INFO,
                RT_BUILD_LIST,
                RT_BUILD_ATOM,
                RT_PARA_BUILD,
                RT_PARA_BUILD_ATOM,
                RT_TIME_NODE,
                RT_TIME_PROPERTY_LIST,
                RT_TIME_VARIANT,
                RT_TIME_EXT_TIME_NODE,
            ],
        )
        self.assertEqual(len(warning["locations"]), 1)
        self.assertEqual(warning["locations"][0]["slide_index"], 1)
        self.assertEqual(
            warning["locations"][0]["record_type"],
            RT_TIME_EXT_TIME_NODE,
        )
        self.assertEqual(warning["locations"][0]["object_kind"], "animation")


@unittest.skipUnless(powerpoint_available(), "Microsoft PowerPoint COM is required")
class AnimationPowerPointVisualTests(unittest.TestCase):
    def test_shape_animation_has_exact_static_visual_and_object_diagnostic(self):
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
        self.assertEqual(report["office_structure"]["differences"], [])
        self.assertEqual(report["summary"]["mean_mae"], 0.0)
        self.assertEqual(report["summary"]["mean_rmse"], 0.0)
        self.assertEqual(report["summary"]["mean_changed_pixel_ratio"], 0.0)
        self.assertEqual(report["summary"]["mean_ssim"], 1.0)
        warnings = {
            item["code"]: item
            for item in report["conversion_warnings"]["warnings"]
        }
        self.assertEqual(set(warnings), {"ANIMATION_OMITTED"})
        warning = warnings["ANIMATION_OMITTED"]
        self.assertEqual(warning["count"], 1)
        self.assertEqual(warning["locations"][0]["slide_index"], 1)
        self.assertEqual(
            warning["locations"][0]["record_type"],
            RT_TIME_EXT_TIME_NODE,
        )


if __name__ == "__main__":
    unittest.main()
