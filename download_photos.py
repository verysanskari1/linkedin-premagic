#!/usr/bin/env python3
"""
download_photos.py
==================

Download a profile photo for each person in a spreadsheet by searching their
LinkedIn URL in Google Images (via Serper.dev) and saving the first image
result.

Files are saved as ``firstname_secondname_company.png`` in the output folder.

Input
-----
An .xlsx that has a LinkedIn URL column (default ``linkedin_url`` -- the column
``google_lookup.py`` produces) plus name and company columns for the filename.
If a row has no URL, it falls back to searching '"Name" "Company" linkedin'.

Usage
-----
    export SERPER_API_KEY="your_key"
    python download_photos.py output1.xlsx \
        --url-col linkedin_url --name-col Name --company-col "Company Name" \
        --out-dir photos -o photos_report.xlsx

Notes
-----
* LinkedIn-hosted images can expire quickly, so we download immediately.
* The first Google Images result is usually the right person but NOT guaranteed
  (common names). Check rows where ``image_status`` isn't ``downloaded`` and
  spot-check the rest. ``image_source`` records where each photo came from.
* Already-downloaded files are skipped, so re-runs are cheap.
"""

from __future__ import annotations

import argparse
import csv
import os
import re
import sys
import time
import unicodedata
from io import BytesIO
from pathlib import Path

import pandas as pd
import requests

# On corporate networks that intercept HTTPS, use the OS trust store so SSL
# verification still works (the company root cert is already trusted there).
try:
    import truststore
    truststore.inject_into_ssl()
except Exception:
    pass

try:
    from PIL import Image
except ImportError:
    print("ERROR: Pillow is required. Run: pip install -r requirements.txt",
          file=sys.stderr)
    sys.exit(2)

SERPER_IMAGES_URL = "https://google.serper.dev/images"
BROWSER_UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
              "AppleWebKit/537.36 (KHTML, like Gecko) "
              "Chrome/124.0 Safari/537.36")

COL_STATUS = "image_status"   # downloaded | no_result | download_failed | error | skipped_no_query
COL_SOURCE = "image_source"   # the image URL that was downloaded
COL_FILE = "image_file"       # saved file path


def slugify(text: str) -> str:
    """Lowercase, strip accents/punctuation, collapse spaces -> underscores."""
    text = unicodedata.normalize("NFKD", str(text))
    text = "".join(c for c in text if not unicodedata.combining(c))
    text = re.sub(r"[^a-z0-9\s]", " ", text.lower())
    return re.sub(r"\s+", "_", text).strip("_")


def build_filename(name: str, company: str) -> str:
    """firstname_secondname_company.png (skips missing parts cleanly)."""
    name_toks = [t for t in slugify(name).split("_") if t]
    first = name_toks[0] if name_toks else ""
    second = name_toks[1] if len(name_toks) > 1 else ""
    company_slug = slugify(company)
    parts = [p for p in (first, second, company_slug) if p]
    base = "_".join(parts) or "unknown"
    return f"{base}.png"


def search_first_image(api_key: str, query: str, session: requests.Session,
                       timeout: int = 30) -> str:
    """Return the URL of the first Google Images result, or '' if none."""
    headers = {"X-API-KEY": api_key, "Content-Type": "application/json"}
    resp = session.post(SERPER_IMAGES_URL, headers=headers,
                        json={"q": query, "num": 5}, timeout=timeout)
    resp.raise_for_status()
    images = resp.json().get("images", []) or []
    for img in images:
        url = img.get("imageUrl") or ""
        if url.startswith("http"):
            return url
    return ""


def download_as_png(image_url: str, dest: Path, session: requests.Session,
                    timeout: int = 30) -> bool:
    """Download an image and save it as a real PNG. Returns True on success."""
    resp = session.get(image_url, headers={"User-Agent": BROWSER_UA}, timeout=timeout)
    resp.raise_for_status()
    img = Image.open(BytesIO(resp.content)).convert("RGB")
    dest.parent.mkdir(parents=True, exist_ok=True)
    img.save(dest, "PNG")
    return True


def run(args: argparse.Namespace) -> int:
    api_key = os.environ.get("SERPER_API_KEY", "").strip()
    if not api_key:
        print("ERROR: set SERPER_API_KEY (get one free at https://serper.dev).",
              file=sys.stderr)
        return 2

    in_path = Path(args.input)
    if not in_path.exists():
        print(f"ERROR: input file not found: {in_path}", file=sys.stderr)
        return 2

    if in_path.suffix.lower() in (".csv", ".txt"):
        df = pd.read_csv(in_path, dtype=str)
    else:
        df = pd.read_excel(in_path, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    if args.name_col not in df.columns:
        print(f"ERROR: name column '{args.name_col}' not found. "
              f"Columns: {list(df.columns)}", file=sys.stderr)
        return 2
    has_url = args.url_col in df.columns
    has_company = args.company_col in df.columns

    if args.limit and args.limit > 0:
        df = df.head(args.limit).copy()
        print(f"(test mode) processing only the first {len(df)} rows\n")

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    session = requests.Session()
    if args.insecure:
        session.verify = False
        requests.packages.urllib3.disable_warnings()  # type: ignore
        print("(insecure mode) SSL certificate verification disabled\n")

    statuses, sources, files = [], [], []
    csv_rows = []          # (image_filename, real_name, real_company) for the backend CSV
    seen: dict[str, int] = {}   # ensure unique filenames within this run
    downloaded = no_result = failed = skipped = 0
    total = len(df)

    for i, row in df.iterrows():
        name = str(row[args.name_col]).strip()
        company = str(row[args.company_col]).strip() if has_company and pd.notna(row[args.company_col]) else ""
        url = str(row[args.url_col]).strip() if has_url and pd.notna(row[args.url_col]) else ""

        # Build a UNIQUE filename: two people with the same name+company get _2, _3 ...
        filename = build_filename(name, company)
        stem = filename[:-4]
        seen[stem] = seen.get(stem, 0) + 1
        if seen[stem] > 1:
            filename = f"{stem}_{seen[stem]}.png"
        dest = out_dir / filename

        # Build the search query: the LinkedIn URL if we have one, else name+company.
        if url and url.lower().startswith("http"):
            query = url
        elif name:
            query = f'"{name}" "{company}" linkedin'.replace('""', '')
        else:
            statuses.append("skipped_no_query"); sources.append(""); files.append("")
            skipped += 1
            continue

        # Skip if already downloaded (cheap re-runs).
        if dest.exists():
            statuses.append("downloaded"); sources.append("(cached file)"); files.append(str(dest))
            csv_rows.append((filename, name, company))
            downloaded += 1
            print(f"[{i + 1}/{total}] {filename} -> already exists, skip")
            continue

        try:
            image_url = search_first_image(api_key, query, session)
        except requests.RequestException as exc:
            statuses.append(f"error: {exc}"); sources.append(""); files.append("")
            failed += 1
            print(f"[{i + 1}/{total}] {name} -> search error")
            time.sleep(args.sleep)
            continue

        if not image_url:
            statuses.append("no_result"); sources.append(""); files.append("")
            no_result += 1
            print(f"[{i + 1}/{total}] {name} -> no image result")
            time.sleep(args.sleep)
            continue

        try:
            download_as_png(image_url, dest, session)
            statuses.append("downloaded"); sources.append(image_url); files.append(str(dest))
            csv_rows.append((filename, name, company))
            downloaded += 1
            print(f"[{i + 1}/{total}] {name} -> {dest}")
        except (requests.RequestException, OSError) as exc:
            statuses.append(f"download_failed: {exc}"); sources.append(image_url); files.append("")
            failed += 1
            print(f"[{i + 1}/{total}] {name} -> download failed")

        time.sleep(args.sleep)

    df[COL_STATUS] = statuses
    df[COL_SOURCE] = sources
    df[COL_FILE] = files
    out_report = Path(args.output)
    df.to_excel(out_report, index=False)

    # Backend CSV: image,name,company -- only rows that actually have a photo.
    # Every value is quoted so commas/apostrophes/accents never break parsing.
    out_csv = Path(args.csv)
    if out_csv.parent != Path(""):
        out_csv.parent.mkdir(parents=True, exist_ok=True)
    with out_csv.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh, quoting=csv.QUOTE_ALL)
        writer.writerow(["image", "name", "company"])
        for image_name, real_name, real_company in csv_rows:
            writer.writerow([image_name, real_name, real_company])

    print("\n=== Summary ===")
    print(f"  downloaded:      {downloaded}")
    print(f"  no image result: {no_result}")
    print(f"  failed/errors:   {failed}")
    print(f"  skipped:         {skipped}")
    print(f"  photos saved to: {out_dir}/")
    print(f"  backend CSV:     {out_csv}  ({len(csv_rows)} rows)")
    print(f"  report written:  {out_report}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("input", help="Input .xlsx (e.g. the output of google_lookup.py)")
    p.add_argument("-o", "--output", default="photos_report.xlsx",
                   help="Report .xlsx with image status columns added.")
    p.add_argument("--url-col", default="linkedin_url")
    p.add_argument("--name-col", default="Name")
    p.add_argument("--company-col", default="Company")
    p.add_argument("--out-dir", default="photos", help="Folder to save photos in.")
    p.add_argument("--csv", default="attendees.csv",
                   help="Backend CSV (image,name,company) listing every saved photo.")
    p.add_argument("--limit", type=int, default=0,
                   help="Only process the first N rows (for a test run). 0 = all.")
    p.add_argument("--sleep", type=float, default=0.5,
                   help="Seconds between people.")
    p.add_argument("--insecure", action="store_true",
                   help="Disable SSL verification (last resort for corporate "
                        "HTTPS-inspecting proxies).")
    return p


if __name__ == "__main__":
    sys.exit(run(build_parser().parse_args()))
