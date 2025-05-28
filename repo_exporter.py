#!/usr/bin/env python3
# /tools/repo_exporter.py
#
# Export every “relevant” text file in a project to a single Excel workbook.
# --------------------------------------------------------------------------
# 1. Walk the tree (skipping venv, .git, __pycache__, etc.)
# 2. Collect code / text files (*.py, *.md, *.json …)
# 3. Build an Excel file:
#       • Summary tab (metadata table)
#       • README tab (usage notes & dependency list)
#       • One tab per source file, with fenced code for legibility
# --------------------------------------------------------------------------
from __future__ import annotations
from pathlib import Path
from datetime import datetime
import mimetypes, os, sys, textwrap

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font

# ─── Configuration ────────────────────────────────────────────────────────
EXCLUDED_DIRS: set[str] = {
    "venv", ".venv", ".git", "__pycache__", ".idea", ".vscode", "node_modules",
}
# any extension you consider “code / text”; add more as needed
INCLUDE_EXTS = {
    ".py", ".md", ".txt", ".json", ".yaml", ".yml", ".toml",
    ".ini", ".cfg", ".csv", ".tsv", ".js", ".ts", ".html", ".css", ".sh",
    ".bat", ".ps1", ".sql",
}
README_CANDIDATES = {"readme", "read_me"}

# ─── Helpers ──────────────────────────────────────────────────────────────
def is_binary(path: Path) -> bool:
    # Rough heuristic: MIME plus a sniff of first bytes
    mime, _ = mimetypes.guess_type(str(path))
    if mime and not mime.startswith("text/"):
        return True
    try:
        with path.open("rb") as f:
            chunk = f.read(512)
        return b"\x00" in chunk
    except Exception:
        return True  # play it safe

def iter_files(root: Path):
    for dirpath, dirnames, filenames in os.walk(root):
        # prune unwanted dirs in-place
        dirnames[:] = [d for d in dirnames if d not in EXCLUDED_DIRS]
        for name in filenames:
            p = Path(dirpath) / name
            if p.suffix.lower() in INCLUDE_EXTS and not is_binary(p):
                yield p.relative_to(root)

def nice_sheet_name(path: Path) -> str:
    """Excel sheet names max 31 chars and no /:*?[]"""
    safe = str(path).replace(os.sep, "_")[:31]
    return safe.encode("ascii", "replace").decode()

# ─── Main ─────────────────────────────────────────────────────────────────
def build_excel(root: Path, out_file: Path):
    files = list(iter_files(root))
    if not files:
        print("❌ No matching files found.")
        return

    wb = Workbook()
    wb.remove(wb.active)  # start with a clean slate

    # Summary sheet --------------------------------------------------------
    summary = wb.create_sheet("Summary")
    summary.append(
        ["Relative Path", "Size (bytes)", "Last Modified", "Sheet Name"]
    )
    for cell in summary[1]:
        cell.font = Font(bold=True)

    for p in files:
        abs_p = root / p
        stats = abs_p.stat()
        sheet_name = nice_sheet_name(p)
        summary.append(
            [str(p), stats.st_size,
             datetime.fromtimestamp(stats.st_mtime).strftime("%Y-%m-%d %H:%M"),
             sheet_name]
        )

    # auto-size summary cols
    for idx, col in enumerate(summary.columns, 1):
        max_len = max(len(str(c.value or "")) for c in col)
        summary.column_dimensions[get_column_letter(idx)].width = max_len + 2

    # README sheet ---------------------------------------------------------
    readme = wb.create_sheet("README")
    readme.column_dimensions["A"].width = 110
    readme["A1"] = "📦 Project Source Export"
    readme["A3"] = "⚙️ How this workbook was generated:"
    readme["A4"] = (
        "1. Place repo_exporter.py in the project root\n"
        "2. Run:  python repo_exporter.py   (optionally -o output.xlsx)\n"
        "3. Examine the Summary sheet for a file index\n"
        "4. Each source file lives in its own tab inside ``` fences."
    )
    # If a requirements/poetry/conda file exists, dump its text here:
    for req_name in ("requirements.txt", "pyproject.toml", "environment.yml"):
        req_path = root / req_name
        if req_path.exists():
            readme_row = readme.max_row + 2
            readme[f"A{readme_row}"] = f"📃 {req_name}:"
            for i, line in enumerate(req_path.read_text().splitlines(),
                                     start=readme_row + 1):
                readme[f"A{i}"] = line.rstrip()

    # Per-file sheets ------------------------------------------------------
    fenced_fill = PatternFill(
        start_color="F8F8F8", end_color="F8F8F8", fill_type="solid"
    )
    mono_font = Font(name="Consolas")

    for p in files:
        sheet = wb.create_sheet(nice_sheet_name(p))
        sheet.column_dimensions["A"].width = 120
        sheet["A1"] = "```" + p.suffix.lstrip(".")
        try:
            src = (root / p).read_text(encoding="utf-8", errors="replace")
        except Exception as e:
            src = f"<<Error reading file: {e}>>"

        for row_i, line in enumerate(src.splitlines(), start=2):
            cell = sheet.cell(row=row_i, column=1, value=line)
            cell.font = mono_font
            cell.fill = fenced_fill
        sheet.cell(row=row_i + 1, column=1, value="```")

    wb.save(out_file)
    print(f"✅ Export complete → {out_file}")

# ─── CLI ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="Export all code / text files in a project to Excel."
    )
    parser.add_argument(
        "root",
        nargs="?",
        default=".",
        help="Project root directory (default: current dir)",
    )
    parser.add_argument(
        "-o", "--out",
        default="project_source_export.xlsx",
        help="Output XLSX filename",
    )
    args = parser.parse_args()

    build_excel(Path(args.root).resolve(), Path(args.out).resolve())
