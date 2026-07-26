"""Create a one-slide PPT containing one embedded Excel worksheet."""
from __future__ import annotations

import argparse
import base64
import hashlib
import json
from pathlib import Path
import tempfile
from urllib.request import urlopen

import pythoncom

from make_zero_extent_line_fixture import (
    PP_LAYOUT_BLANK,
    PP_SAVE_AS_PRESENTATION,
    _launch_powerpoint,
)


_SEED_BLOB = "b0c9cc8afb48b4cb40a941112de3222206a9651d"
_SEED_SHA256 = "b5c19872ad7b0417c4c601f97bc0f8ee855fd7bbe9c196ab9dbad8316a71f44c"
MSO_EMBEDDED_OLE_OBJECT = 7


def _download_seed() -> bytes:
    url = f"https://api.github.com/repos/apache/poi/git/blobs/{_SEED_BLOB}"
    with urlopen(url, timeout=60) as response:
        payload = json.load(response)
    content = base64.b64decode(payload["content"])
    digest = hashlib.sha256(content).hexdigest()
    if digest != _SEED_SHA256:
        raise RuntimeError(f"unexpected Apache POI seed SHA-256: {digest}")
    return content


def make_excel_ole_fixture(destination: Path) -> dict[str, object]:
    destination = destination.resolve()
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        destination.unlink()

    pythoncom.CoInitialize()
    app = None
    source_presentation = None
    presentation = None
    owned_pid = None
    result: dict[str, object] | None = None
    try:
        app, owned_pid = _launch_powerpoint()
        try:
            app.AskToUpdateLinks = False
        except Exception:
            pass
        with tempfile.TemporaryDirectory(
            prefix="ppt2pptx-excel-ole-fixture-"
        ) as directory:
            seed = Path(directory) / "apache-poi-testPPT_oleWorkbook.ppt"
            seed.write_bytes(_download_seed())
            source_presentation = app.Presentations.Open(
                str(seed),
                ReadOnly=True,
                Untitled=True,
                WithWindow=False,
            )
            worksheet = None
            for slide_index in range(
                1, source_presentation.Slides.Count + 1
            ):
                slide = source_presentation.Slides(slide_index)
                for shape_index in range(1, slide.Shapes.Count + 1):
                    shape = slide.Shapes(shape_index)
                    if int(shape.Type) != MSO_EMBEDDED_OLE_OBJECT:
                        continue
                    if str(shape.OLEFormat.ProgID) == "Excel.Sheet.12":
                        worksheet = shape
                        break
                if worksheet is not None:
                    break
            if worksheet is None:
                raise RuntimeError(
                    "Apache POI seed contains no Excel.Sheet.12 object"
                )

            presentation = app.Presentations.Add(True)
            slide = presentation.Slides.Add(1, PP_LAYOUT_BLANK)
            worksheet.Copy()
            pasted = slide.Shapes.Paste()
            if pasted.Count != 1:
                raise RuntimeError("expected one pasted worksheet object")
            output_worksheet = pasted(1)
            output_worksheet.Left = 120
            output_worksheet.Top = 135
            presentation.SaveAs(
                str(destination),
                PP_SAVE_AS_PRESENTATION,
            )
            result = {
                "path": str(destination),
                "powerpoint_version": str(app.Version),
                "owned_pid": owned_pid,
                "slide_count": int(presentation.Slides.Count),
                "ole_count": 1,
                "progid": str(output_worksheet.OLEFormat.ProgID),
            }
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        if source_presentation is not None:
            try:
                source_presentation.Close()
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        presentation = None
        source_presentation = None
        app = None
        pythoncom.CoUninitialize()
    if result is None:
        raise RuntimeError("fixture generation did not complete")
    return result


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("tests/fixtures/visual_excel_ole.ppt"),
    )
    args = parser.parse_args()
    print(make_excel_ole_fixture(args.output))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
