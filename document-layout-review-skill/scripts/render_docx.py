#!/usr/bin/env python3
"""Render DOCX -> PDF -> PNG pages for visual verification.

Requires LibreOffice/soffice. PNG rendering additionally requires pdftoppm.
"""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
from pathlib import Path


def resolve_executable(explicit: str | None, names: list[str]) -> str | None:
    if explicit:
        p = Path(explicit)
        return str(p) if p.exists() else shutil.which(explicit)
    for name in names:
        found = shutil.which(name)
        if found:
            return found
    return None


def run(cmd: list[str], timeout: int = 120) -> None:
    completed = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout, check=False)
    if completed.returncode != 0:
        raise SystemExit(
            "command failed:\n" + " ".join(cmd) + "\n" + completed.stdout + "\n" + completed.stderr
        )


def main() -> int:
    parser = argparse.ArgumentParser(description="Render DOCX to PDF and optional PNG pages.")
    parser.add_argument("docx", type=Path)
    parser.add_argument("--out-dir", type=Path, required=True)
    parser.add_argument("--soffice")
    parser.add_argument("--pdftoppm")
    parser.add_argument("--dpi", type=int, default=120)
    args = parser.parse_args()

    if not args.docx.is_file():
        parser.error(f"file not found: {args.docx}")
    args.out_dir.mkdir(parents=True, exist_ok=True)

    soffice = resolve_executable(args.soffice, ["soffice", "libreoffice", "soffice.exe"])
    if not soffice:
        raise SystemExit("LibreOffice/soffice not found; pass --soffice <path> or use the host renderer.")
    run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(args.out_dir), str(args.docx)])
    pdf = args.out_dir / f"{args.docx.stem}.pdf"
    if not pdf.is_file():
        raise SystemExit(f"rendered PDF not found: {pdf}")

    pages: list[str] = []
    pdftoppm = resolve_executable(args.pdftoppm, ["pdftoppm", "pdftoppm.exe"])
    if pdftoppm:
        prefix = args.out_dir / f"{args.docx.stem}-page"
        run([pdftoppm, "-png", "-r", str(args.dpi), str(pdf), str(prefix)])
        pages = [str(p) for p in sorted(args.out_dir.glob(f"{args.docx.stem}-page-*.png"))]

    result = {"docx": str(args.docx), "pdf": str(pdf), "pages": pages, "png_renderer_available": bool(pdftoppm)}
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
