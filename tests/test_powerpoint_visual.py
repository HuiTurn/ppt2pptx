"""PowerPoint bilateral visual regression tests (Windows + Office only)."""
from __future__ import annotations

import json
import os
from pathlib import Path
import platform
import subprocess
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
FIXTURE = ROOT / "tests" / "fixtures" / "visual_minimal.ppt"
MAKE_FIXTURE = ROOT / "scripts" / "make_visual_fixture.py"
COMPARE = ROOT / "scripts" / "compare_powerpoint_visual.py"


def _powerpoint_available() -> bool:
    if platform.system() != "Windows":
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
        _ = app.Version
        return True
    except Exception:
        return False
    finally:
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


@unittest.skipUnless(_powerpoint_available(), "Microsoft PowerPoint COM is required")
class PowerPointVisualRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        env = os.environ.copy()
        env["PYTHONPATH"] = str(ROOT / "src")
        FIXTURE.parent.mkdir(parents=True, exist_ok=True)
        if FIXTURE.is_file() and FIXTURE.stat().st_size > 0:
            return
        completed = subprocess.run(
            [sys.executable, str(MAKE_FIXTURE), "-o", str(FIXTURE)],
            capture_output=True,
            text=True,
            env=env,
            timeout=120,
            check=False,
        )
        if completed.returncode != 0 or not FIXTURE.is_file():
            raise unittest.SkipTest(
                "unable to create visual fixture via PowerPoint: "
                f"{completed.stderr or completed.stdout}"
            )

    def test_bilateral_visual_pipeline_produces_metrics(self) -> None:
        env = os.environ.copy()
        env["PYTHONPATH"] = str(ROOT / "src")
        with tempfile.TemporaryDirectory(prefix="ppt2pptx-visual-test-") as tmp:
            out = Path(tmp) / "evidence"
            completed = subprocess.run(
                [
                    sys.executable,
                    str(COMPARE),
                    str(FIXTURE),
                    "-o",
                    str(out),
                    "--width",
                    "960",
                    "--height",
                    "720",
                ],
                capture_output=True,
                text=True,
                env=env,
                timeout=240,
                check=False,
            )
            report_path = out / "report.json"
            self.assertTrue(
                report_path.is_file(),
                msg=f"stdout={completed.stdout!r} stderr={completed.stderr!r}",
            )
            report = json.loads(report_path.read_text(encoding="utf-8"))
            self.assertEqual(report["provider"], "office")
            self.assertTrue(report["powerpoint_version"])
            self.assertEqual(report["powerpoint_instances"]["dispatch"], "DispatchEx")
            self.assertIsInstance(report["powerpoint_instances"]["reference_pid"], int)
            self.assertIsInstance(report["powerpoint_instances"]["actual_pid"], int)
            self.assertEqual(
                report["powerpoint_versions"]["reference"],
                report["powerpoint_version"],
            )
            self.assertTrue(report["powerpoint_versions"]["actual"])
            self.assertEqual(report["slide_count_source"], 2)
            self.assertEqual(report["slide_count_output"], 2)
            self.assertEqual(report["hard_differences"], [])
            self.assertEqual(report["hidden_source"], [False, False])
            self.assertEqual(report["hidden_output"], [False, False])
            self.assertEqual(len(report["slides"]), 2)
            for slide in report["slides"]:
                self.assertTrue(Path(slide["reference"]).is_file())
                self.assertTrue(Path(slide["actual"]).is_file())
                self.assertTrue(Path(slide["absolute_diff"]).is_file())
                self.assertTrue(Path(slide["overlay"]).is_file())
                self.assertIn("mae", slide)
                self.assertIn("rmse", slide)
                self.assertIn("changed_pixel_ratio", slide)
                self.assertIn("ssim", slide)
            self.assertIsNotNone(report["summary"]["mean_mae"])
            # Structure counts must be present for editable-object assertions.
            self.assertGreaterEqual(report["structure"]["text_box_count"], 1)
            warnings = report["conversion_warnings"]["warnings"]
            codes = [item["code"] for item in warnings]
            self.assertNotIn("ADVANCED_FEATURES_APPROXIMATED", codes)
            self.assertTrue(
                {
                    "ANIMATION_OMITTED",
                    "AUDIO_OMITTED",
                    "VIDEO_OMITTED",
                    "EMBEDDED_OLE_OMITTED",
                    "CHART_OMITTED",
                    "DIAGRAM_OR_SMARTART_OMITTED",
                    "COMPLEX_FREEFORM_OMITTED",
                }.isdisjoint(codes)
            )

    def test_converted_pptx_reopens_read_only_in_powerpoint(self) -> None:
        from ppt2pptx import convert
        import pythoncom
        import win32com.client

        with tempfile.TemporaryDirectory(prefix="ppt2pptx-reopen-") as tmp:
            pptx = Path(tmp) / "out.pptx"
            result = convert(FIXTURE, pptx)
            self.assertEqual(result.slide_count, 2)
            self.assertFalse(any(slide.hidden for slide in result.presentation.slides))
            pythoncom.CoInitialize()
            app = None
            presentation = None
            try:
                app = win32com.client.DispatchEx("PowerPoint.Application")
                try:
                    app.Visible = True
                except Exception:
                    pass
                try:
                    app.DisplayAlerts = 1  # ppAlertsNone
                except Exception:
                    pass
                try:
                    app.AutomationSecurity = 3
                except Exception:
                    pass
                try:
                    app.AskToUpdateLinks = False
                except Exception:
                    pass
                try:
                    presentation = app.Presentations.Open(str(pptx), True, False, False)
                except Exception:
                    presentation = app.Presentations.Open(str(pptx), True, False, True)
                self.assertEqual(int(presentation.Slides.Count), 2)
            finally:
                # Do not save.
                try:
                    if presentation is not None:
                        presentation.Close()
                except Exception:
                    pass
                try:
                    if app is not None:
                        app.Quit()
                except Exception:
                    pass
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass


if __name__ == "__main__":
    unittest.main()
