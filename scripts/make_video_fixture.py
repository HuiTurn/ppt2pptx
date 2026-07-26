"""Create a minimal legacy PPT containing one PowerPoint media shape."""
from __future__ import annotations

import argparse
import csv
import io
from pathlib import Path
import shutil
import subprocess
import tempfile
import time

try:
    import pythoncom
    import win32com.client
except ImportError as exc:  # pragma: no cover
    raise SystemExit(f"pywin32 is required: {exc}") from exc


def _powerpoint_pids() -> set[int]:
    completed = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/FO", "CSV", "/NH"],
        capture_output=True,
        text=True,
        errors="replace",
        check=False,
    )
    pids: set[int] = set()
    for row in csv.reader(io.StringIO(completed.stdout)):
        if len(row) >= 2 and row[0].casefold() == "powerpnt.exe":
            try:
                pids.add(int(row[1]))
            except ValueError:
                pass
    return pids


def _make_video(path: Path) -> None:
    ffmpeg = shutil.which("ffmpeg")
    if ffmpeg is None:
        raise RuntimeError("ffmpeg is required to generate the controlled video")
    completed = subprocess.run(
        [
            ffmpeg,
            "-hide_banner",
            "-loglevel",
            "error",
            "-f",
            "lavfi",
            "-i",
            "color=c=0x1F4E78:s=160x90:r=10:d=1",
            "-an",
            "-c:v",
            "mpeg4",
            "-q:v",
            "2",
            "-map_metadata",
            "-1",
            "-fflags",
            "+bitexact",
            "-flags:v",
            "+bitexact",
            "-y",
            str(path),
        ],
        capture_output=True,
        text=True,
        check=False,
        timeout=60,
    )
    if completed.returncode != 0 or not path.is_file():
        raise RuntimeError(f"ffmpeg failed: {completed.stderr.strip()}")


def _launch_powerpoint():
    last_error: Exception | None = None
    for attempt in range(3):
        before = _powerpoint_pids()
        app = None
        try:
            app = win32com.client.DispatchEx("PowerPoint.Application")
            try:
                app.Visible = True
            except Exception:
                pass
            app.DisplayAlerts = 1
            try:
                app.AutomationSecurity = 3
            except Exception:
                pass
            for _ in range(20):
                created = _powerpoint_pids() - before
                if len(created) == 1:
                    return app, created.pop()
                if len(created) > 1:
                    break
                time.sleep(0.1)
            raise RuntimeError("unable to record isolated PowerPoint instance PID")
        except Exception as exc:
            last_error = exc
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass
            app = None
            time.sleep(0.5 * (attempt + 1))
    raise RuntimeError("unable to launch stable isolated PowerPoint") from last_error


def make_video_fixture(destination: Path) -> dict[str, object]:
    destination = destination.resolve()
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        destination.unlink()

    pythoncom.CoInitialize()
    app = None
    presentation = None
    owned_pid = None
    try:
        app, owned_pid = _launch_powerpoint()

        presentation = app.Presentations.Add(True)
        presentation.PageSetup.SlideWidth = 720
        presentation.PageSetup.SlideHeight = 540
        slide = presentation.Slides.Add(1, 12)
        with tempfile.TemporaryDirectory(prefix="ppt2pptx-video-fixture-") as directory:
            video = Path(directory) / "controlled.mp4"
            _make_video(video)
            shape = slide.Shapes.AddMediaObject2(
                str(video),
                False,
                True,
                180,
                160,
                360,
                202.5,
            )
            shape.Name = "Controlled embedded video"
            presentation.SaveAs(str(destination), 1)
        return {
            "path": str(destination),
            "powerpoint_version": str(app.Version),
            "owned_pid": owned_pid,
            "slide_count": int(presentation.Slides.Count),
            "media_type": int(shape.MediaType),
        }
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        presentation = None
        app = None
        pythoncom.CoUninitialize()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("tests/fixtures/visual_video.ppt"),
    )
    args = parser.parse_args()
    print(make_video_fixture(args.output))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
