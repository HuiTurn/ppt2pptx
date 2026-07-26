"""Tests for object-backed lossy-feature diagnostics."""
from __future__ import annotations

import os
import struct
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from ppt2pptx import convert
from ppt2pptx.ppt import (
    RT_ANIMATION_INFO,
    RT_ANIMATION_INFO_ATOM,
    RT_EXTERNAL_OLE_EMBED,
    RT_SOUND,
    detect_lossy_features,
)


ROOT = Path(__file__).resolve().parents[1]
FIXTURE_MINIMAL = ROOT / "tests" / "fixtures" / "visual_minimal.ppt"
FIXTURE_ANIMATION = ROOT / "tests" / "fixtures" / "visual_animation.ppt"
MAKE_ANIMATION = ROOT / "scripts" / "make_animation_fixture.py"


def rec(kind, payload=b"", version=0xF, instance=0):
    return struct.pack("<HHI", (instance << 4) | version, kind, len(payload)) + payload


class LossyFeatureTests(unittest.TestCase):
    def test_detects_animation_audio_and_ole_records(self):
        document = (
            rec(1000, rec(RT_ANIMATION_INFO, rec(RT_ANIMATION_INFO_ATOM, bytes(28), 0)))
            + rec(RT_SOUND, b"\x00" * 8)
            + rec(RT_EXTERNAL_OLE_EMBED, b"\x00" * 8)
        )
        features = {item.code: item for item in detect_lossy_features(document)}
        self.assertEqual(features["ANIMATION_OMITTED"].count, 1)
        self.assertIn(RT_ANIMATION_INFO, features["ANIMATION_OMITTED"].record_types)
        self.assertEqual(features["AUDIO_OMITTED"].count, 1)
        self.assertEqual(features["EMBEDDED_OLE_OMITTED"].count, 1)

    def test_attaches_slide_locations_for_embedded_slide_records(self):
        animation = rec(RT_ANIMATION_INFO, rec(RT_ANIMATION_INFO_ATOM, bytes(28), 0))
        slide = rec(1006, animation)
        document = rec(1000, slide)
        features = {item.code: item for item in detect_lossy_features(document)}
        animation_feature = features["ANIMATION_OMITTED"]
        self.assertGreaterEqual(animation_feature.count, 1)
        self.assertTrue(animation_feature.locations)
        self.assertTrue(any(location.slide_index == 1 for location in animation_feature.locations))
        self.assertTrue(any(location.object_kind == "animation" for location in animation_feature.locations))

    def test_detects_chart_progid_cstring(self):
        progid = "Excel.Chart.8".encode("utf-16le")
        document = rec(1000, rec(0x0FBA, progid, 0))
        features = {item.code: item for item in detect_lossy_features(document)}
        self.assertEqual(features["CHART_OMITTED"].count, 1)
        self.assertTrue(features["CHART_OMITTED"].locations)

    def test_ignores_plain_presentations(self):
        text = rec(4000, "Hello".encode("utf-16le"), 0)
        document = rec(1000, rec(1006, text))
        self.assertEqual(detect_lossy_features(document), ())

    def test_minimal_fixture_no_longer_emits_unconditional_advanced_warning(self):
        if not FIXTURE_MINIMAL.is_file():
            self.skipTest("visual_minimal.ppt fixture is missing")
        with tempfile.TemporaryDirectory() as directory:
            result = convert(FIXTURE_MINIMAL, Path(directory) / "out.pptx")
        codes = [item["code"] for item in result.report.warnings]
        self.assertNotIn("ADVANCED_FEATURES_APPROXIMATED", codes)
        advanced = {
            "ANIMATION_OMITTED",
            "AUDIO_OMITTED",
            "VIDEO_OMITTED",
            "EMBEDDED_OLE_OMITTED",
            "CHART_OMITTED",
            "DIAGRAM_OR_SMARTART_OMITTED",
            "COMPLEX_FREEFORM_OMITTED",
        }
        self.assertTrue(advanced.isdisjoint(codes))

    def test_animation_fixture_reports_located_animation_warning(self):
        if sys.platform != "win32":
            self.skipTest("animation fixture generation requires Windows PowerPoint")
        env = os.environ.copy()
        env["PYTHONPATH"] = str(ROOT / "src")
        FIXTURE_ANIMATION.parent.mkdir(parents=True, exist_ok=True)
        if not FIXTURE_ANIMATION.is_file():
            completed = subprocess.run(
                [sys.executable, str(MAKE_ANIMATION), "-o", str(FIXTURE_ANIMATION)],
                capture_output=True,
                text=True,
                env=env,
                timeout=180,
                check=False,
            )
            if completed.returncode != 0 or not FIXTURE_ANIMATION.is_file():
                self.skipTest(
                    "unable to create animation fixture: "
                    f"{completed.stderr or completed.stdout}"
                )
        with tempfile.TemporaryDirectory() as directory:
            result = convert(FIXTURE_ANIMATION, Path(directory) / "animated.pptx")
        warnings = {item["code"]: item for item in result.report.warnings}
        self.assertIn("ANIMATION_OMITTED", warnings)
        self.assertGreaterEqual(warnings["ANIMATION_OMITTED"]["count"], 1)
        locations = warnings["ANIMATION_OMITTED"]["locations"]
        self.assertTrue(locations)
        self.assertTrue(any(item.get("slide_index") == 1 for item in locations))
        self.assertTrue(any(item.get("object_kind") == "animation" for item in locations))
        # Audio is best-effort depending on the host; assert only when present.
        if "AUDIO_OMITTED" in warnings:
            self.assertGreaterEqual(warnings["AUDIO_OMITTED"]["count"], 1)
            self.assertTrue(warnings["AUDIO_OMITTED"]["locations"])


if __name__ == "__main__":
    unittest.main()
