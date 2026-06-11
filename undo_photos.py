#!/usr/bin/env python3
"""
undo_photos.py
==============

Reverse a download_photos.py (Google Images) run by deleting ONLY the files
that run downloaded -- leaving photos that were already there (e.g. the
accurate face_url downloads) untouched.

It reads the run's report (photos_report.xlsx). Rows that were freshly
downloaded have a real http URL in ``image_source``; rows that were skipped
because the file already existed say ``(cached file)``. We delete only the
freshly-downloaded ones.

Safe by default: prints what it WOULD delete and changes nothing. Add
``--confirm`` to actually delete.

Usage
-----
    python undo_photos.py photos_report.xlsx              # preview (dry run)
    python undo_photos.py photos_report.xlsx --confirm    # actually delete
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd


def _read(path: Path) -> pd.DataFrame:
    if path.suffix.lower() in (".csv", ".txt"):
        return pd.read_csv(path, dtype=str)
    return pd.read_excel(path, dtype=str)


def run_keep_csv(args: argparse.Namespace) -> int:
    """Reliable undo: keep only the photos listed in a known-good CSV, delete
    every other image in the folder. Works even if a run was Ctrl+C'd."""
    keep_csv = Path(args.keep_csv)
    photos_dir = Path(args.photos_dir)
    if not keep_csv.exists():
        print(f"ERROR: keep-csv not found: {keep_csv}", file=sys.stderr); return 2
    if not photos_dir.is_dir():
        print(f"ERROR: photos dir not found: {photos_dir}", file=sys.stderr); return 2

    df = _read(keep_csv)
    df.columns = [str(c).strip() for c in df.columns]
    if "image" not in df.columns:
        print(f"ERROR: 'image' column not in {keep_csv}. Columns: {list(df.columns)}",
              file=sys.stderr); return 2
    keep = {str(v).strip() for v in df["image"].dropna()}

    on_disk = [p for p in photos_dir.iterdir()
               if p.is_file() and p.suffix.lower() in (".png", ".jpg", ".jpeg")]
    to_delete = [p for p in on_disk if p.name not in keep]

    print(f"Keeping {len(keep)} photos listed in {keep_csv.name}.")
    print(f"{len(on_disk)} image files in {photos_dir}/; "
          f"{len(to_delete)} are NOT in the keep-list.\n")
    for p in to_delete:
        print(("DELETE " if args.confirm else "would delete ") + str(p))

    if not args.confirm:
        print(f"\nDRY RUN — nothing deleted. Re-run with --confirm to delete "
              f"these {len(to_delete)} files.")
        return 0

    deleted = 0
    for p in to_delete:
        try:
            p.unlink(); deleted += 1
        except OSError as exc:
            print(f"  could not delete {p}: {exc}", file=sys.stderr)
    print(f"\nDeleted {deleted} files. The {len(keep)} good photos were kept.")
    return 0


def run(args: argparse.Namespace) -> int:
    if args.keep_csv:
        return run_keep_csv(args)
    report = Path(args.report)
    if not report.exists():
        print(f"ERROR: report not found: {report}", file=sys.stderr)
        return 2

    if report.suffix.lower() in (".csv", ".txt"):
        df = pd.read_csv(report, dtype=str)
    else:
        df = pd.read_excel(report, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    for col in ("image_source", "image_file"):
        if col not in df.columns:
            print(f"ERROR: column '{col}' not in report. Columns: {list(df.columns)}",
                  file=sys.stderr)
            return 2

    # Freshly downloaded this run = image_source is a real URL (not '(cached file)').
    src = df["image_source"].fillna("")
    fresh = df[src.str.startswith("http")]

    to_delete = []
    for _, row in fresh.iterrows():
        path = str(row.get("image_file") or "").strip()
        if path and Path(path).exists():
            to_delete.append(Path(path))

    print(f"Found {len(fresh)} files downloaded by that run; "
          f"{len(to_delete)} still exist on disk.\n")
    for p in to_delete:
        print(("DELETE " if args.confirm else "would delete ") + str(p))

    if not args.confirm:
        print(f"\nDRY RUN — nothing deleted. Re-run with --confirm to delete "
              f"these {len(to_delete)} files.")
        return 0

    deleted = 0
    for p in to_delete:
        try:
            p.unlink()
            deleted += 1
        except OSError as exc:
            print(f"  could not delete {p}: {exc}", file=sys.stderr)
    print(f"\nDeleted {deleted} files. Your other photos were left untouched.")
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("report", nargs="?", default="",
                   help="The photos_report.xlsx from the run to undo "
                        "(report-based mode).")
    p.add_argument("--keep-csv", default="",
                   help="Reliable mode: path to your KNOWN-GOOD CSV (e.g. "
                        "backend/attendees.csv). Keeps only photos listed there "
                        "and deletes the rest from --photos-dir.")
    p.add_argument("--photos-dir", default="backend/attendees",
                   help="Folder of photos to clean (used with --keep-csv).")
    p.add_argument("--confirm", action="store_true",
                   help="Actually delete (without this, it only previews).")
    return p


if __name__ == "__main__":
    sys.exit(run(build_parser().parse_args()))
