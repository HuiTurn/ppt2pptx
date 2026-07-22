from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

from . import __version__
from .cfb import Limits
from .converter import convert, inspect_ppt
from .diagnostics import write_json_file
from .errors import Ppt2PptxError

def _positive_int(value: str) -> int:
    parsed = int(value)
    if parsed <= 0:
        raise argparse.ArgumentTypeError("value must be positive")
    return parsed

def _common(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--max-input-bytes", type=_positive_int, default=Limits().max_input_bytes)
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--password", help="password for an encrypted presentation")
    group.add_argument("--password-file", type=Path, help="UTF-8 file containing the password")

def _password(args: argparse.Namespace) -> str | None:
    if args.password is not None:
        return args.password
    if args.password_file is None:
        return None
    if args.password_file.stat().st_size > 4096:
        raise OSError("password file exceeds 4096 bytes")
    return args.password_file.read_text(encoding="utf-8").removesuffix("\n").removesuffix("\r")

def _conversion_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="ppt2pptx", description="Convert PowerPoint 97–2003 .ppt files to .pptx")
    parser.add_argument("input", type=Path)
    parser.add_argument("-o", "--output", type=Path)
    parser.add_argument("--report", type=Path)
    parser.add_argument("--version", action="version", version=__version__)
    _common(parser)
    return parser

def _inspection_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="ppt2pptx inspect", description="Inspect slides and editable text without converting")
    parser.add_argument("input", type=Path)
    parser.add_argument("--json", action="store_true", dest="as_json")
    _common(parser)
    return parser

def _batch_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="ppt2pptx batch", description="Convert a directory of PowerPoint 97–2003 .ppt files")
    parser.add_argument("input", type=Path)
    parser.add_argument("-o", "--output", type=Path, required=True)
    parser.add_argument("--recursive", action="store_true")
    parser.add_argument("--report", type=Path)
    _common(parser)
    return parser

def _same_path(left: Path, right: Path) -> bool:
    return left.resolve().as_posix().casefold() == right.resolve().as_posix().casefold()

def _batch_sources(root: Path, recursive: bool) -> list[Path]:
    iterator = root.rglob("*") if recursive else root.iterdir()
    return sorted(
        (path for path in iterator if path.is_file() and path.suffix.casefold() == ".ppt" and not path.name.startswith("~$")),
        key=lambda path: path.relative_to(root).as_posix().casefold(),
    )

def _run_batch(args: argparse.Namespace) -> int:
    source_root, output_root = args.input, args.output
    if not source_root.is_dir():
        print(f"ppt2pptx: error: batch input is not a directory: {source_root}", file=sys.stderr)
        return 2
    sources = _batch_sources(source_root, args.recursive)
    password = _password(args)
    if not sources:
        print(f"ppt2pptx: error: no .ppt files found in {source_root}", file=sys.stderr)
        return 2
    destinations = [output_root / source.relative_to(source_root).with_suffix(".pptx") for source in sources]
    if args.report and any(_same_path(args.report, path) for path in (*sources, *destinations)):
        print("ppt2pptx: error: report path would overwrite an input or output", file=sys.stderr)
        return 2
    results: list[dict[str, object]] = []
    succeeded = failed = warning_count = 0
    seen_destinations: set[str] = set()
    for source, destination in zip(sources, destinations):
        destination_key = destination.resolve().as_posix().casefold()
        if destination_key in seen_destinations:
            error = f"multiple inputs map to the same output {destination}"
            print(f"Failed {source}: {error}", file=sys.stderr)
            results.append({"source": str(source), "destination": str(destination), "status": "failed", "error": error})
            failed += 1
            continue
        seen_destinations.add(destination_key)
        try:
            result = convert(source, destination, limits=Limits(max_input_bytes=args.max_input_bytes), password=password)
        except (Ppt2PptxError, OSError) as exc:
            print(f"Failed {source}: {exc}", file=sys.stderr)
            results.append({"source": str(source), "destination": str(destination), "status": "failed", "error": str(exc)})
            failed += 1
            continue
        warnings = len(result.report.warnings)
        warning_count += warnings
        succeeded += 1
        print(f"Converted {source} -> {destination}")
        results.append({"source": str(source), "destination": str(destination), "status": "converted", "slide_count": result.slide_count, "warning_count": warnings, "report": result.report.to_dict()})
    summary = {"input": str(source_root), "output": str(output_root), "recursive": bool(args.recursive), "file_count": len(sources), "succeeded": succeeded, "failed": failed, "warning_count": warning_count, "results": results}
    if args.report:
        write_json_file(args.report, summary)
    print(f"Batch complete: {succeeded} converted, {failed} failed")
    return 1 if failed else 0

def main(argv: list[str] | None = None) -> int:
    raw = list(sys.argv[1:] if argv is None else argv)
    mode = raw.pop(0) if raw and raw[0] in ("inspect", "batch") else "convert"
    parser = {"inspect": _inspection_parser, "batch": _batch_parser}.get(mode, _conversion_parser)()
    args = parser.parse_args(raw)
    if mode == "batch":
        try:
            return _run_batch(args)
        except (Ppt2PptxError, OSError) as exc:
            print(f"ppt2pptx: error: {exc}", file=sys.stderr)
            return 2
    try:
        limits = Limits(max_input_bytes=args.max_input_bytes)
        password = _password(args)
        if mode == "inspect":
            result = inspect_ppt(args.input, limits=limits, password=password)
            print(json.dumps(result, ensure_ascii=False, indent=2) if args.as_json else f"{result['slide_count']} slides; {result['text_box_count']} text boxes")
            return 0
        output = args.output or args.input.with_suffix(".pptx")
        if args.report and (_same_path(args.report, args.input) or _same_path(args.report, output)):
            print("ppt2pptx: error: report path would overwrite an input or output", file=sys.stderr)
            return 2
        result = convert(args.input, args.output, limits=limits, password=password)
        if args.report:
            write_json_file(args.report, result.report.to_dict())
        print(f"Converted {args.input} -> {result.output_path} ({result.slide_count} slides)")
        return 0
    except (Ppt2PptxError, OSError) as exc:
        print(f"ppt2pptx: error: {exc}", file=sys.stderr)
        return 2

if __name__ == "__main__":
    raise SystemExit(main())
