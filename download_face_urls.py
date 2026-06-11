#!/usr/bin/env python3
"""
download_face_urls.py
=====================

Download profile photos directly from a ``face_url`` column in a CSV/Excel file
(e.g. the Premagic event export) and save them as
``firstname_secondname_company.png``.

This is the simple, 100%-accurate path: no searching, no API key -- the photo
URL is already in the file. It also writes the backend CSV (image,name,company).

The image URLs are signed and may be hotlink-protected, so we send a browser
User-Agent and a Referer header (``--referer``) which is what makes them load
outside the gallery. Rows with an empty ``face_url`` are flagged.

Usage
-----
    python download_face_urls.py people.csv \
        --url-col face_url --name-col name --company-col "custom_company name" \
        --out-dir backend/attendees --csv backend/attendees.csv \
        --referer https://gallery.hackerrank.com/ -o face_report.xlsx
"""

from __future__ import annotations

import argparse
import csv
import sys
import time
from io import BytesIO
from pathlib import Path

import pandas as pd
import requests

try:
    from PIL import Image
except ImportError:
    print("ERROR: Pillow is required. Run: pip install -r requirements.txt",
          file=sys.stderr)
    sys.exit(2)

# reuse the filename logic so both scripts name files identically
from download_photos import build_filename, BROWSER_UA

COL_STATUS = "face_status"   # downloaded | no_face_url | download_failed | error
COL_FILE = "image_file"


def read_table(path: Path) -> pd.DataFrame:
    if path.suffix.lower() in (".csv", ".txt"):
        return pd.read_csv(path, dtype=str)
    return pd.read_excel(path, dtype=str)


def download_as_png(url: str, dest: Path, session: requests.Session,
                    referer: str, timeout: int = 30) -> None:
    headers = {"User-Agent": BROWSER_UA}
    if referer:
        headers["Referer"] = referer
    resp = session.get(url, headers=headers, timeout=timeout)
    resp.raise_for_status()
    img = Image.open(BytesIO(resp.content)).convert("RGB")
    dest.parent.mkdir(parents=True, exist_ok=True)
    img.save(dest, "PNG")


def run(args: argparse.Namespace) -> int:
    in_path = Path(args.input)
    if not in_path.exists():
        print(f"ERROR: input file not found: {in_path}", file=sys.stderr)
        return 2

    df = read_table(in_path)
    df.columns = [str(c).strip() for c in df.columns]
    for col in (args.name_col, args.url_col):
        if col not in df.columns:
            print(f"ERROR: column '{col}' not found. Columns: {list(df.columns)}",
                  file=sys.stderr)
            return 2
    has_company = args.company_col in df.columns

    if args.limit and args.limit > 0:
        df = df.head(args.limit).copy()
        print(f"(test mode) processing only the first {len(df)} rows\n")

    out_dir = Path(args.out_dir)
    session = requests.Session()

    statuses, files, csv_rows = [], [], []
    seen: dict[str, int] = {}
    downloaded = no_url = failed = 0
    total = len(df)

    for i, row in df.iterrows():
        name = str(row[args.name_col]).strip() if pd.notna(row[args.name_col]) else ""
        company = str(row[args.company_col]).strip() if has_company and pd.notna(row[args.company_col]) else ""
        url = str(row[args.url_col]).strip() if pd.notna(row[args.url_col]) else ""

        if not url or not url.lower().startswith("http"):
            statuses.append("no_face_url"); files.append("")
            no_url += 1
            print(f"[{i + 1}/{total}] {name or '?'} -> no face_url")
            continue

        # unique filename
        filename = build_filename(name, company)
        stem = filename[:-4]
        seen[stem] = seen.get(stem, 0) + 1
        if seen[stem] > 1:
            filename = f"{stem}_{seen[stem]}.png"
        dest = out_dir / filename

        if dest.exists():
            statuses.append("downloaded"); files.append(str(dest))
            csv_rows.append((filename, name, company))
            downloaded += 1
            print(f"[{i + 1}/{total}] {filename} -> already exists, skip")
            continue

        try:
            download_as_png(url, dest, session, args.referer)
            statuses.append("downloaded"); files.append(str(dest))
            csv_rows.append((filename, name, company))
            downloaded += 1
            print(f"[{i + 1}/{total}] {name} -> {dest}")
        except (requests.RequestException, OSError) as exc:
            statuses.append(f"download_failed: {str(exc)[:80]}"); files.append("")
            failed += 1
            print(f"[{i + 1}/{total}] {name} -> FAILED: {str(exc)[:80]}")

        time.sleep(args.sleep)

    df[COL_STATUS] = statuses
    df[COL_FILE] = files
    df.to_excel(Path(args.output), index=False)

    out_csv = Path(args.csv)
    with out_csv.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh, quoting=csv.QUOTE_ALL)
        writer.writerow(["image", "name", "company"])
        for image_name, real_name, real_company in csv_rows:
            writer.writerow([image_name, real_name, real_company])

    print("\n=== Summary ===")
    print(f"  downloaded:     {downloaded}")
    print(f"  no face_url:    {no_url}")
    print(f"  failed:         {failed}")
    print(f"  photos saved:   {out_dir}/")
    print(f"  backend CSV:    {out_csv}  ({len(csv_rows)} rows)")
    print(f"  report:         {args.output}")
    if failed and downloaded == 0:
        print("\n  All downloads failed -- if status is 403, try a different "
              "--referer (the gallery URL), or the URLs may have expired.")
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("input", help="Input CSV or .xlsx with a face_url column")
    p.add_argument("-o", "--output", default="face_report.xlsx")
    p.add_argument("--url-col", default="face_url")
    p.add_argument("--name-col", default="name")
    p.add_argument("--company-col", default="custom_company name")
    p.add_argument("--out-dir", default="photos")
    p.add_argument("--csv", default="attendees.csv",
                   help="Backend CSV (image,name,company).")
    p.add_argument("--referer", default="https://gallery.hackerrank.com/",
                   help="Referer header sent with each image request (beats "
                        "hotlink protection). Set '' to disable.")
    p.add_argument("--limit", type=int, default=0,
                   help="Only process the first N rows (test run). 0 = all.")
    p.add_argument("--sleep", type=float, default=0.1)
    return p


if __name__ == "__main__":
    sys.exit(run(build_parser().parse_args()))
