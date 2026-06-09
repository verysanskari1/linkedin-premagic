#!/usr/bin/env python3
"""
linkedin_lookup.py
==================

Resolve LinkedIn profile URLs / public IDs for a list of people given their
name, designation (title) and company.

Approach
--------
We use the People Data Labs (PDL) *Person Enrichment* API. PDL is a licensed
data provider rather than a LinkedIn scraper, so:
  * there is no risk of getting a LinkedIn account banned, and
  * you only pay for *successful* matches.

Each lookup sends the person's name + company (+ title when present) and PDL
returns the best matching profile together with a ``likelihood`` score (0-10).
We only accept matches at/above a configurable threshold and flag everything
else as "needs review" so the output stays accurate instead of guessing.

Usage
-----
    export PDL_API_KEY="your_key_here"
    python linkedin_lookup.py input.xlsx -o output.xlsx \
        --name-col Name --title-col Designation --company-col Company \
        --min-likelihood 6

Get a free key (100 lookups/month) at https://www.peopledatalabs.com/.

Swapping providers
------------------
All provider-specific logic lives in ``resolve_profile()``. To switch to
Apollo, RocketReach, etc., reimplement just that function.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import time
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Optional

import pandas as pd
import requests

PDL_ENRICH_URL = "https://api.peopledatalabs.com/v5/person/enrich"

# Output column names appended to the spreadsheet.
COL_URL = "linkedin_url"
COL_ID = "linkedin_id"
COL_SCORE = "match_likelihood"
COL_STATUS = "lookup_status"


@dataclass
class LookupResult:
    linkedin_url: str = ""
    linkedin_id: str = ""
    match_likelihood: Optional[int] = None
    lookup_status: str = ""  # matched | low_confidence | not_found | error


def linkedin_id_from_url(url: str) -> str:
    """Extract the public id, e.g. 'john-doe-123' from a /in/<id> URL."""
    if not url:
        return ""
    cleaned = url.split("?")[0].rstrip("/")
    marker = "/in/"
    if marker in cleaned:
        return cleaned.split(marker, 1)[1]
    return cleaned.rsplit("/", 1)[-1]


def resolve_profile(
    api_key: str,
    name: str,
    company: Optional[str],
    title: Optional[str],
    min_likelihood: int,
    session: requests.Session,
    timeout: int = 30,
) -> LookupResult:
    """Resolve a single person to a LinkedIn profile via People Data Labs."""
    params = {
        "name": name,
        "min_likelihood": 2,          # ask PDL for anything plausible; we filter ourselves
        "include_if_matched": "true",
        "pretty": "false",
    }
    if company:
        params["company"] = company
    if title:
        params["title"] = title

    headers = {"X-Api-Key": api_key, "Accept": "application/json"}

    try:
        resp = session.get(
            PDL_ENRICH_URL, headers=headers, params=params, timeout=timeout
        )
    except requests.RequestException as exc:
        return LookupResult(lookup_status=f"error: {exc}")

    # 404 = no confident match found (this is NOT billed).
    if resp.status_code == 404:
        return LookupResult(lookup_status="not_found")

    if resp.status_code == 401:
        return LookupResult(lookup_status="error: invalid API key (401)")

    if resp.status_code == 429:
        return LookupResult(lookup_status="error: rate limited (429)")

    if resp.status_code != 200:
        return LookupResult(
            lookup_status=f"error: HTTP {resp.status_code} {resp.text[:120]}"
        )

    body = resp.json()
    likelihood = body.get("likelihood")
    data = body.get("data") or {}
    url = data.get("linkedin_url") or ""
    public_id = data.get("linkedin_username") or linkedin_id_from_url(url)

    if not url:
        return LookupResult(match_likelihood=likelihood, lookup_status="not_found")

    status = "matched" if (likelihood or 0) >= min_likelihood else "low_confidence"
    return LookupResult(
        linkedin_url=url,
        linkedin_id=public_id,
        match_likelihood=likelihood,
        lookup_status=status,
    )


def load_cache(path: Optional[Path]) -> dict:
    if path and path.exists():
        try:
            return json.loads(path.read_text())
        except json.JSONDecodeError:
            return {}
    return {}


def save_cache(path: Optional[Path], cache: dict) -> None:
    if path:
        path.write_text(json.dumps(cache, indent=2))


def run(args: argparse.Namespace) -> int:
    api_key = os.environ.get("PDL_API_KEY", "").strip()
    if not api_key:
        print("ERROR: set the PDL_API_KEY environment variable.", file=sys.stderr)
        return 2

    in_path = Path(args.input)
    if not in_path.exists():
        print(f"ERROR: input file not found: {in_path}", file=sys.stderr)
        return 2

    df = pd.read_excel(in_path)
    df.columns = [str(c).strip() for c in df.columns]  # tolerate trailing/leading spaces
    for col in (args.name_col, args.company_col):
        if col not in df.columns:
            print(
                f"ERROR: column '{col}' not in spreadsheet. "
                f"Found: {list(df.columns)}",
                file=sys.stderr,
            )
            return 2
    has_title = args.title_col in df.columns

    cache_path = Path(args.cache) if args.cache else None
    cache = load_cache(cache_path)

    session = requests.Session()
    results: list[dict] = []
    matched = low = missing = errored = 0

    total = len(df)
    for i, row in df.iterrows():
        name = str(row[args.name_col]).strip()
        company = str(row[args.company_col]).strip() if pd.notna(row[args.company_col]) else ""
        title = (
            str(row[args.title_col]).strip()
            if has_title and pd.notna(row[args.title_col])
            else ""
        )

        if not name or name.lower() == "nan":
            results.append(asdict(LookupResult(lookup_status="error: missing name")))
            errored += 1
            continue

        cache_key = f"{name}|{company}|{title}".lower()
        if cache_key in cache:
            res_dict = cache[cache_key]
        else:
            res = resolve_profile(
                api_key, name, company, title, args.min_likelihood, session
            )
            res_dict = asdict(res)
            cache[cache_key] = res_dict
            save_cache(cache_path, cache)
            time.sleep(args.sleep)  # be polite to the API / stay under rate limits

        status = res_dict["lookup_status"]
        if status == "matched":
            matched += 1
        elif status == "low_confidence":
            low += 1
        elif status == "not_found":
            missing += 1
        else:
            errored += 1

        results.append(res_dict)
        print(
            f"[{i + 1}/{total}] {name} @ {company or '?'} -> {status}"
            + (f" ({res_dict['linkedin_url']})" if res_dict["linkedin_url"] else "")
        )

    res_df = pd.DataFrame(results)
    out_df = pd.concat([df.reset_index(drop=True), res_df], axis=1)

    out_path = Path(args.output)
    out_df.to_excel(out_path, index=False)

    print("\n=== Summary ===")
    print(f"  matched (>= {args.min_likelihood}): {matched}")
    print(f"  low_confidence (review):       {low}")
    print(f"  not_found:                     {missing}")
    print(f"  errors:                        {errored}")
    print(f"  written to:                    {out_path}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("input", help="Input .xlsx file")
    p.add_argument("-o", "--output", default="output.xlsx", help="Output .xlsx file")
    p.add_argument("--name-col", default="Name")
    p.add_argument("--title-col", default="Designation")
    p.add_argument("--company-col", default="Company")
    p.add_argument(
        "--min-likelihood",
        type=int,
        default=6,
        help="Accept matches at/above this PDL likelihood (0-10). Higher = stricter.",
    )
    p.add_argument(
        "--sleep",
        type=float,
        default=0.2,
        help="Seconds to wait between API calls.",
    )
    p.add_argument(
        "--cache",
        default=".lookup_cache.json",
        help="Cache file to avoid re-paying for repeated lookups. Pass '' to disable.",
    )
    return p


if __name__ == "__main__":
    sys.exit(run(build_parser().parse_args()))
