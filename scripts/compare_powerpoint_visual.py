"""PowerPoint COM bilateral visual regression for ppt2pptx.

Exports per-slide reference PNGs from the source .ppt and actual PNGs from the
converted .pptx using the same pixel size, then writes metrics, diffs, and a
JSON report. Conversion itself never uses PowerPoint/COM.
"""
from __future__ import annotations

import argparse
import atexit
import csv
import hashlib
import io
import json
from pathlib import Path
import platform
import shutil
import subprocess
import sys
import tempfile
import time
import traceback
from typing import Any

# Optional heavy deps used only by this validation tool.
try:
    import numpy as np
    from PIL import Image, ImageChops, ImageDraw, ImageFont
except ImportError as exc:  # pragma: no cover - exercised when deps missing
    raise SystemExit(
        "compare_powerpoint_visual.py requires Pillow and numpy "
        f"(pip install pillow numpy). Import error: {exc}"
    ) from exc

try:
    import pythoncom
    import win32com.client
    from win32com.client import constants as pp_constants
except ImportError as exc:  # pragma: no cover
    raise SystemExit(
        "compare_powerpoint_visual.py requires pywin32 on Windows "
        f"(pip install pywin32). Import error: {exc}"
    ) from exc

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT / "src") not in sys.path:
    sys.path.insert(0, str(ROOT / "src"))

from ppt2pptx import convert  # noqa: E402
from ppt2pptx.errors import Ppt2PptxError  # noqa: E402


# PowerPoint.Application instances started by this process (pid -> app).
_OWNED_APPS: dict[int, Any] = {}
_OWNED_PIDS: set[int] = set()


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _read_password(password: str | None, password_file: Path | None) -> str | None:
    if password is not None and password_file is not None:
        raise SystemExit("pass --password or --password-file, not both")
    if password_file is not None:
        return password_file.read_text(encoding="utf-8").splitlines()[0]
    return password


def _copy_workdir(source: Path, destination: Path) -> Path:
    """Copy into a throwaway work dir. Never save presentations back to these paths."""
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, destination)
    return destination


def _release_com(obj: Any) -> None:
    try:
        if obj is not None:
            obj = None
    finally:
        pass


def _quit_owned_app(app: Any, pid: int | None) -> None:
    try:
        if app is not None:
            app.Quit()
    except Exception:
        pass
    if pid is not None:
        _OWNED_APPS.pop(pid, None)
        _OWNED_PIDS.discard(pid)
    _release_com(app)


def _cleanup_owned_apps() -> None:
    for pid, app in list(_OWNED_APPS.items()):
        _quit_owned_app(app, pid)


atexit.register(_cleanup_owned_apps)


def _powerpoint_pids() -> set[int]:
    """Return current POWERPNT.EXE PIDs without obtaining termination handles."""
    completed = subprocess.run(
        [
            "tasklist",
            "/FI",
            "IMAGENAME eq POWERPNT.EXE",
            "/FO",
            "CSV",
            "/NH",
        ],
        capture_output=True,
        text=True,
        errors="replace",
        check=False,
    )
    pids: set[int] = set()
    for row in csv.reader(io.StringIO(completed.stdout)):
        if len(row) < 2 or row[0].casefold() != "powerpnt.exe":
            continue
        try:
            pids.add(int(row[1]))
        except ValueError:
            continue
    return pids


def _launch_powerpoint() -> tuple[Any, int | None, str]:
    pythoncom.CoInitialize()
    last_error: Exception | None = None
    for launch_attempt in range(3):
        before = _powerpoint_pids()
        app = None
        pid = None
        try:
            # DispatchEx is required: attaching to a user's existing PowerPoint
            # process would make cleanup and timeout ownership unsafe.
            app = win32com.client.DispatchEx("PowerPoint.Application")
            # Recent PowerPoint builds reject Visible=False for some hosts.
            try:
                app.Visible = True
            except Exception:
                pass
            try:
                app.DisplayAlerts = pp_constants.ppAlertsNone
            except Exception:
                try:
                    app.DisplayAlerts = 1  # ppAlertsNone
                except Exception:
                    pass
            try:
                app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            except Exception:
                pass
            try:
                app.AskToUpdateLinks = False
            except Exception:
                pass
            # PowerPoint does not expose an HWND property. DispatchEx starts an
            # isolated local server, so identify only the single PID created
            # after our snapshot.
            for _poll_attempt in range(20):
                created = _powerpoint_pids() - before
                if len(created) == 1:
                    pid = created.pop()
                    break
                if len(created) > 1:
                    break
                time.sleep(0.1)
            if pid is None:
                raise RuntimeError(
                    "unable to record the isolated PowerPoint instance PID"
                )
            version = str(getattr(app, "Version", "unknown"))
            _OWNED_APPS[pid] = app
            _OWNED_PIDS.add(pid)
            return app, pid, version
        except Exception as exc:
            last_error = exc
            try:
                if app is not None:
                    app.Quit()
            except Exception:
                pass
            _release_com(app)
            time.sleep(0.5 * (launch_attempt + 1))
    raise RuntimeError(
        "unable to launch a stable isolated PowerPoint instance"
    ) from last_error


def _open_presentation(app: Any, path: Path, *, read_only: bool = True, password: str | None = None):
    # Prefer WithWindow=False; fall back if the host rejects a hidden window.
    open_attempts = (
        {"FileName": str(path), "ReadOnly": read_only, "WithWindow": False},
        {"FileName": str(path), "ReadOnly": read_only, "WithWindow": True},
        {"FileName": str(path), "ReadOnly": read_only},
    )
    last_error: Exception | None = None
    for kwargs in open_attempts:
        if password:
            # Password is passed to COM only; never written to logs/report.
            kwargs = {**kwargs, "Password": password}
        try:
            return app.Presentations.Open(**kwargs)
        except Exception as exc:
            last_error = exc
            continue
    assert last_error is not None
    raise last_error


def _slide_pixel_size(presentation: Any, width: int | None, height: int | None) -> tuple[int, int]:
    # SlideWidth/SlideHeight are in points (72 pt = 1 inch).
    slide_w_pt = float(presentation.PageSetup.SlideWidth)
    slide_h_pt = float(presentation.PageSetup.SlideHeight)
    if width and height:
        return width, height
    # Default ~150 DPI equivalent for stable diffs without huge files.
    px_w = width or max(1, int(round(slide_w_pt * 150 / 72.0)))
    if height:
        px_h = height
    else:
        px_h = max(1, int(round(px_w * slide_h_pt / slide_w_pt)))
    return px_w, px_h


_OFFICE_STRUCTURE_FIELDS = (
    "shape_count",
    "text_shape_count",
    "picture_count",
    "table_count",
    "chart_count",
    "group_count",
    "ole_count",
    "media_count",
    "comment_count",
    "note_text_count",
)


def _com_bool(value: Any) -> bool:
    try:
        return bool(int(value))
    except (TypeError, ValueError):
        return bool(value)


def _office_slide_structure(slide: Any, index: int) -> dict[str, int]:
    counts = {field: 0 for field in _OFFICE_STRUCTURE_FIELDS}
    counts["index"] = index
    shapes = slide.Shapes
    counts["shape_count"] = int(shapes.Count)
    for shape_index in range(1, counts["shape_count"] + 1):
        shape = shapes(shape_index)
        try:
            shape_type = int(shape.Type)
        except Exception:
            shape_type = 0
        try:
            if _com_bool(shape.HasTextFrame) and _com_bool(shape.TextFrame.HasText):
                counts["text_shape_count"] += 1
        except Exception:
            pass
        if shape_type in (11, 13):  # msoLinkedPicture, msoPicture
            counts["picture_count"] += 1
        if shape_type == 6:  # msoGroup
            counts["group_count"] += 1
        if shape_type in (7, 10, 12):  # embedded/linked OLE and OLE controls
            counts["ole_count"] += 1
        if shape_type == 16:  # msoMedia
            counts["media_count"] += 1
        try:
            if _com_bool(shape.HasTable):
                counts["table_count"] += 1
        except Exception:
            pass
        try:
            if _com_bool(shape.HasChart):
                counts["chart_count"] += 1
        except Exception:
            pass
    try:
        counts["comment_count"] = int(slide.Comments.Count)
    except Exception:
        pass
    try:
        note_shapes = slide.NotesPage.Shapes
        for shape_index in range(1, int(note_shapes.Count) + 1):
            shape = note_shapes(shape_index)
            try:
                if int(shape.Type) != 14:  # msoPlaceholder
                    continue
                if int(shape.PlaceholderFormat.Type) != 2:  # ppPlaceholderBody
                    continue
                text = str(shape.TextFrame.TextRange.Text).strip()
                if text:
                    counts["note_text_count"] += 1
            except Exception:
                continue
    except Exception:
        pass
    return counts


def _office_presentation_structure(presentation: Any) -> dict[str, Any]:
    slides = [
        _office_slide_structure(presentation.Slides(index), index)
        for index in range(1, int(presentation.Slides.Count) + 1)
    ]
    return {
        field: sum(slide[field] for slide in slides)
        for field in _OFFICE_STRUCTURE_FIELDS
    } | {"slides": slides}


def _office_structure_differences(
    reference: dict[str, Any], actual: dict[str, Any]
) -> list[dict[str, Any]]:
    differences: list[dict[str, Any]] = []
    reference_slides = reference["slides"]
    actual_slides = actual["slides"]
    for index in range(min(len(reference_slides), len(actual_slides))):
        reference_slide = reference_slides[index]
        actual_slide = actual_slides[index]
        for field in _OFFICE_STRUCTURE_FIELDS:
            if reference_slide[field] != actual_slide[field]:
                differences.append(
                    {
                        "slide_index": index + 1,
                        "field": field,
                        "source": reference_slide[field],
                        "output": actual_slide[field],
                    }
                )
    return differences


def _export_slides(
    path: Path,
    out_dir: Path,
    *,
    width: int | None,
    height: int | None,
    password: str | None,
    label: str,
    timeout_s: float,
) -> dict[str, Any]:
    out_dir.mkdir(parents=True, exist_ok=True)
    started = time.monotonic()
    app = None
    pid = None
    presentation = None
    version = "unknown"
    slides_meta: list[dict[str, Any]] = []
    try:
        app, pid, version = _launch_powerpoint()
        presentation = _open_presentation(app, path, read_only=True, password=password)
        px_w, px_h = _slide_pixel_size(presentation, width, height)
        slide_count = int(presentation.Slides.Count)
        slide_w_pt = float(presentation.PageSetup.SlideWidth)
        slide_h_pt = float(presentation.PageSetup.SlideHeight)
        structure = _office_presentation_structure(presentation)
        for index in range(1, slide_count + 1):
            if time.monotonic() - started > timeout_s:
                raise TimeoutError(
                    f"PowerPoint export timed out after {timeout_s:.0f}s while exporting {label}"
                )
            slide = presentation.Slides(index)
            hidden = bool(slide.SlideShowTransition.Hidden)
            png_path = out_dir / f"slide-{index:03d}.png"
            # Export(FileName, FilterName, ScaleWidth, ScaleHeight)
            slide.Export(str(png_path), "PNG", px_w, px_h)
            slides_meta.append(
                {
                    "index": index,
                    "hidden": hidden,
                    "png": str(png_path),
                    "exists": png_path.is_file(),
                }
            )
        return {
            "path": str(path),
            "powerpoint_version": version,
            "owned_pid": pid,
            "slide_count": slide_count,
            "slide_width_pt": slide_w_pt,
            "slide_height_pt": slide_h_pt,
            "export_width_px": px_w,
            "export_height_px": px_h,
            "structure": structure,
            "slides": slides_meta,
        }
    finally:
        try:
            if presentation is not None:
                presentation.Close()
        except Exception:
            pass
        _quit_owned_app(app, pid)
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _ssim(ref: np.ndarray, act: np.ndarray) -> float:
    """Mean SSIM over RGB channels (window-free luminance approximation)."""
    # Lightweight global SSIM-style metric (not full sliding-window SSIM).
    c1 = (0.01 * 255) ** 2
    c2 = (0.03 * 255) ** 2
    scores = []
    for channel in range(ref.shape[2]):
        x = ref[:, :, channel].astype(np.float64)
        y = act[:, :, channel].astype(np.float64)
        mx, my = x.mean(), y.mean()
        vx, vy = x.var(), y.var()
        cov = ((x - mx) * (y - my)).mean()
        scores.append(((2 * mx * my + c1) * (2 * cov + c2)) / ((mx**2 + my**2 + c1) * (vx + vy + c2)))
    return float(sum(scores) / len(scores))


def _compare_pair(reference: Path, actual: Path, diff_dir: Path, index: int) -> dict[str, Any]:
    diff_dir.mkdir(parents=True, exist_ok=True)
    ref_img = Image.open(reference).convert("RGB")
    act_img = Image.open(actual).convert("RGB")
    size_mismatch = ref_img.size != act_img.size
    if size_mismatch:
        act_img = act_img.resize(ref_img.size, Image.Resampling.NEAREST)
    ref = np.asarray(ref_img, dtype=np.float64)
    act = np.asarray(act_img, dtype=np.float64)
    abs_diff = np.abs(ref - act)
    mae = float(abs_diff.mean())
    rmse = float(np.sqrt((abs_diff**2).mean()))
    changed = float((abs_diff.max(axis=2) > 0).mean())
    ssim = _ssim(ref, act)
    abs_img = Image.fromarray(np.clip(abs_diff, 0, 255).astype(np.uint8))
    abs_path = diff_dir / f"slide-{index:03d}-absdiff.png"
    abs_img.save(abs_path)
    # Overlay: emphasize differences in red on the reference.
    overlay = ref_img.copy().convert("RGBA")
    mask = Image.fromarray(((abs_diff.max(axis=2) > 8) * 180).astype(np.uint8), mode="L")
    red = Image.new("RGBA", overlay.size, (220, 32, 32, 0))
    red.putalpha(mask)
    overlay = Image.alpha_composite(overlay, red)
    overlay_path = diff_dir / f"slide-{index:03d}-overlay.png"
    overlay.save(overlay_path)
    side = Image.new("RGB", (ref_img.width * 2, ref_img.height))
    side.paste(ref_img, (0, 0))
    side.paste(act_img, (ref_img.width, 0))
    draw = ImageDraw.Draw(side)
    draw.line([(ref_img.width, 0), (ref_img.width, ref_img.height)], fill=(255, 0, 0), width=2)
    side_path = diff_dir / f"slide-{index:03d}-side-by-side.png"
    side.save(side_path)
    return {
        "index": index,
        "reference": str(reference),
        "actual": str(actual),
        "absolute_diff": str(abs_path),
        "overlay": str(overlay_path),
        "side_by_side": str(side_path),
        "size_mismatch": size_mismatch,
        "mae": mae,
        "rmse": rmse,
        "changed_pixel_ratio": changed,
        "ssim": ssim,
    }


def _structure_counts(presentation) -> dict[str, int]:
    return {
        "text_box_count": sum(len(slide.text_boxes) for slide in presentation.slides),
        "picture_count": sum(len(slide.pictures) for slide in presentation.slides),
        "shape_count": sum(len(slide.shapes) for slide in presentation.slides),
        "comment_count": sum(len(slide.comments) for slide in presentation.slides),
        "note_count": sum(len(slide.notes) for slide in presentation.slides),
    }


def compare_files(
    source: Path,
    output_dir: Path,
    *,
    password: str | None = None,
    width: int | None = None,
    height: int | None = None,
    timeout_s: float = 180.0,
    destination: Path | None = None,
) -> dict[str, Any]:
    output_dir.mkdir(parents=True, exist_ok=True)
    work = Path(tempfile.mkdtemp(prefix="ppt2pptx-visual-"))
    try:
        source_copy = _copy_workdir(source, work / f"source{source.suffix.lower()}")
        pptx_path = destination or (output_dir / source.with_suffix(".pptx").name)
        conversion = convert(source_copy, pptx_path, password=password)
        # Re-copy product into the work dir so PowerPoint cannot save over the kept output.
        pptx_copy = _copy_workdir(pptx_path, work / "actual.pptx")

        reference_dir = output_dir / "reference"
        actual_dir = output_dir / "actual"
        diff_dir = output_dir / "diff"

        reference_meta = _export_slides(
            source_copy,
            reference_dir,
            width=width,
            height=height,
            password=password,
            label="reference",
            timeout_s=timeout_s,
        )
        actual_meta = _export_slides(
            pptx_copy,
            actual_dir,
            width=reference_meta["export_width_px"],
            height=reference_meta["export_height_px"],
            password=None,
            label="actual",
            timeout_s=timeout_s,
        )

        hard_differences: list[str] = []
        if reference_meta["slide_count"] != actual_meta["slide_count"]:
            hard_differences.append(
                f"slide_count {reference_meta['slide_count']} != {actual_meta['slide_count']}"
            )
        if abs(reference_meta["slide_width_pt"] - actual_meta["slide_width_pt"]) > 0.05:
            hard_differences.append("slide_width mismatch")
        if abs(reference_meta["slide_height_pt"] - actual_meta["slide_height_pt"]) > 0.05:
            hard_differences.append("slide_height mismatch")

        ref_hidden = [s["hidden"] for s in reference_meta["slides"]]
        act_hidden = [s["hidden"] for s in actual_meta["slides"]]
        if ref_hidden != act_hidden[: len(ref_hidden)]:
            hard_differences.append("hidden_state mismatch")

        pair_count = min(reference_meta["slide_count"], actual_meta["slide_count"])
        slide_metrics = []
        for index in range(1, pair_count + 1):
            slide_metrics.append(
                _compare_pair(
                    reference_dir / f"slide-{index:03d}.png",
                    actual_dir / f"slide-{index:03d}.png",
                    diff_dir,
                    index,
                )
            )

        counts = _structure_counts(conversion.presentation)
        report = {
            "provider": "office",
            "powerpoint_version": reference_meta["powerpoint_version"],
            "powerpoint_versions": {
                "reference": reference_meta["powerpoint_version"],
                "actual": actual_meta["powerpoint_version"],
            },
            "powerpoint_instances": {
                "dispatch": "DispatchEx",
                "reference_pid": reference_meta["owned_pid"],
                "actual_pid": actual_meta["owned_pid"],
            },
            "windows_version": platform.platform(),
            "python_version": sys.version.split()[0],
            "source": str(source),
            "source_sha256": _sha256(source),
            "output": str(pptx_path),
            "output_sha256": _sha256(pptx_path),
            "slide_count_source": reference_meta["slide_count"],
            "slide_count_output": actual_meta["slide_count"],
            "slide_width_pt": reference_meta["slide_width_pt"],
            "slide_height_pt": reference_meta["slide_height_pt"],
            "export_width_px": reference_meta["export_width_px"],
            "export_height_px": reference_meta["export_height_px"],
            "hidden_source": ref_hidden,
            "hidden_output": act_hidden,
            "hard_differences": hard_differences,
            "structure": {
                **counts,
                "slide_width_emu_or_master": conversion.presentation.width,
                "slide_height_emu_or_master": conversion.presentation.height,
                "slides": [
                    {
                        "index": i + 1,
                        "hidden": slide.hidden,
                        "text_boxes": len(slide.text_boxes),
                        "pictures": len(slide.pictures),
                        "shapes": len(slide.shapes),
                        "comments": len(slide.comments),
                        "notes": len(slide.notes),
                    }
                    for i, slide in enumerate(conversion.presentation.slides)
                ],
            },
            "office_structure": {
                "source": reference_meta["structure"],
                "output": actual_meta["structure"],
                "differences": _office_structure_differences(
                    reference_meta["structure"],
                    actual_meta["structure"],
                ),
            },
            "conversion_warnings": conversion.report.to_dict(),
            "slides": slide_metrics,
            "summary": {
                "mean_mae": float(np.mean([m["mae"] for m in slide_metrics])) if slide_metrics else None,
                "mean_rmse": float(np.mean([m["rmse"] for m in slide_metrics])) if slide_metrics else None,
                "mean_changed_pixel_ratio": (
                    float(np.mean([m["changed_pixel_ratio"] for m in slide_metrics])) if slide_metrics else None
                ),
                "mean_ssim": float(np.mean([m["ssim"] for m in slide_metrics])) if slide_metrics else None,
                "max_mae": float(np.max([m["mae"] for m in slide_metrics])) if slide_metrics else None,
                "hard_difference_count": len(hard_differences),
            },
        }
        report_path = output_dir / "report.json"
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        report["report_path"] = str(report_path)
        return report
    finally:
        shutil.rmtree(work, ignore_errors=True)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path, help="source .ppt path")
    parser.add_argument("-o", "--output", type=Path, required=True, help="evidence output directory")
    parser.add_argument("--destination", type=Path, help="converted .pptx path (default: under -o)")
    parser.add_argument("--password", help="password for encrypted source (never written to report)")
    parser.add_argument("--password-file", type=Path, help="file containing the password")
    parser.add_argument("--width", type=int, help="export width in pixels")
    parser.add_argument("--height", type=int, help="export height in pixels")
    parser.add_argument("--timeout", type=float, default=180.0, help="per-side export timeout seconds")
    args = parser.parse_args(argv)

    if platform.system() != "Windows":
        parser.error("PowerPoint COM bilateral comparison requires Windows")
    password = _read_password(args.password, args.password_file)
    try:
        report = compare_files(
            args.source.resolve(),
            args.output.resolve(),
            password=password,
            width=args.width,
            height=args.height,
            timeout_s=args.timeout,
            destination=args.destination.resolve() if args.destination else None,
        )
    except (Ppt2PptxError, OSError, RuntimeError, TimeoutError) as exc:
        print(f"FAILED: {type(exc).__name__}: {exc}", file=sys.stderr)
        traceback.print_exc()
        return 2
    summary = report["summary"]
    print(
        json.dumps(
            {
                "provider": report["provider"],
                "powerpoint_version": report["powerpoint_version"],
                "slides": report["slide_count_source"],
                "hard_differences": report["hard_differences"],
                "mean_mae": summary["mean_mae"],
                "mean_rmse": summary["mean_rmse"],
                "mean_changed_pixel_ratio": summary["mean_changed_pixel_ratio"],
                "mean_ssim": summary["mean_ssim"],
                "report": report["report_path"],
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 1 if report["hard_differences"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
