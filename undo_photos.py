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


def run(args: argparse.Namespace) -> int:
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
    p.add_argument("report", help="The photos_report.xlsx from the run to undo")
    p.add_argument("--confirm", action="store_true",
                   help="Actually delete (without this, it only previews).")
    return p


if __name__ == "__main__":
    sys.exit(run(build_parser().parse_args()))
