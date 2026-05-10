#!/usr/bin/env python3
from __future__ import annotations

import argparse
import subprocess
import sys
import tempfile
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent


def run_child(args: list[str]) -> None:
    result = subprocess.run(args, text=True)
    if result.returncode != 0:
        raise SystemExit(result.returncode)


def output_path(source: Path, suffix: str) -> Path:
    return source.with_suffix(suffix)


def convert_excel(source: Path) -> Path:
    source = source.resolve()
    drawio = output_path(source, ".drawio")
    with tempfile.TemporaryDirectory(prefix="office-to-mdrow-") as tmp:
        intermediate = Path(tmp) / "intermediate.json"
        run_child([
            sys.executable,
            str(SCRIPT_DIR / "xlsx_to_drawio.py"),
            "--input",
            str(source),
            "--json",
            str(intermediate),
            "--drawio",
            str(drawio),
        ])
    return drawio


def convert_word(source: Path) -> Path:
    source = source.resolve()
    markdown = output_path(source, ".md")
    with tempfile.TemporaryDirectory(prefix="office-to-mdrow-") as tmp:
        intermediate_name = Path(tmp) / "intermediate.json"
        run_child([
            sys.executable,
            str(SCRIPT_DIR / "docx_to_md.py"),
            str(source),
            "--out-dir",
            str(source.parent),
            "--json",
            str(intermediate_name),
            "--md",
            markdown.name,
        ])
    # docx_to_md.py treats --json as a filename under --out-dir when relative,
    # and as an absolute path when absolute Path text is passed. Remove either.
    for candidate in (source.parent / str(intermediate_name), intermediate_name):
        if candidate.exists():
            candidate.unlink()
    return markdown


def main() -> None:
    parser = argparse.ArgumentParser(description="office to mdrow converter wrapper")
    subparsers = parser.add_subparsers(dest="command", required=True)

    excel = subparsers.add_parser("excel-drawio")
    excel.add_argument("source", type=Path)

    word = subparsers.add_parser("word-md")
    word.add_argument("source", type=Path)

    args = parser.parse_args()
    if args.command == "excel-drawio":
        out = convert_excel(args.source)
    elif args.command == "word-md":
        out = convert_word(args.source)
    else:
        raise SystemExit(f"Unsupported command: {args.command}")

    print(out.as_posix())


if __name__ == "__main__":
    main()
