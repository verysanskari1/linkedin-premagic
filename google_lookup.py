#!/usr/bin/env python3
"""
google_lookup.py
================

Find LinkedIn profile URLs for a list of people (name + designation + company)
by running a Google search for each person and verifying the result.

How it works
------------
For every row we search Google for:

    "<Name>" "<Company>" site:linkedin.com/in

We then look at the organic results, keep only ``linkedin.com/in/`` URLs, and
**score** each candidate instead of blindly trusting result #1:

    + name tokens present in the result title   (strong signal)
    + company present in the title or snippet    (strong signal)
    + designation present in title/snippet       (small boost)

LinkedIn titles look like ``"Jane Doe - VP Sales - Acme | LinkedIn"``, so this
scoring is reliable. The best candidate is accepted only if its score clears
``--min-score``; otherwise it is flagged ``low_confidence`` for manual review.
This keeps the output accurate rather than guessing.

Search backend
--------------
Google's *official* Custom Search JSON API is closed to new signups (retiring
Jan 2027), so by default we use **Serper.dev** (Google results, 2,500 free
queries, then ~$1/1k). If you already have an official CSE key, pass
--backend google with GOOGLE_API_KEY + GOOGLE_CX set.

Usage
-----
    export SERPER_API_KEY="your_key"          # from https://serper.dev
    python google_lookup.py input.xlsx -o output.xlsx \
        --name-col Name --title-col Designation --company-col Company

    # or, official Google CSE:
    export GOOGLE_API_KEY=...; export GOOGLE_CX=...
    python google_lookup.py input.xlsx --backend google
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
import unicodedata
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Optional

import pandas as pd
import requests

# On corporate networks that intercept HTTPS, use the OS trust store so SSL
# verification still works (the company root cert is already trusted there).
try:
    import truststore
    truststore.inject_into_ssl()
except Exception:
    pass

SERPER_URL = "https://google.serper.dev/search"
GOOGLE_CSE_URL = "https://www.googleapis.com/customsearch/v1"

COL_URL = "linkedin_url"
COL_ID = "linkedin_id"
COL_SCORE = "match_score"
COL_STATUS = "lookup_status"
COL_TITLE = "matched_title"

LINKEDIN_IN_RE = re.compile(r"https?://[\w.]*linkedin\.com/in/[^\s/?#]+", re.I)
STOPWORDS = {"the", "inc", "inc.", "llc", "ltd", "ltd.", "co", "corp", "corporation",
             "company", "pvt", "private", "limited", "&", "and", "group", "technologies",
             "technology", "solutions", "services", "systems", "global", "labs"}


@dataclass
class LookupResult:
    linkedin_url: str = ""
    linkedin_id: str = ""
    match_score: int = 0
    matched_title: str = ""
    lookup_status: str = ""  # matched | low_confidence | not_found | error


def normalize(text: str) -> str:
    """Lowercase, strip accents and punctuation -> space-separated tokens string."""
    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if not unicodedata.combining(c))
    text = re.sub(r"[^a-z0-9\s]", " ", text.lower())
    return re.sub(r"\s+", " ", text).strip()


def tokens(text: str) -> list[str]:
    return [t for t in normalize(text).split() if t]


def linkedin_id_from_url(url: str) -> str:
    cleaned = url.split("?")[0].rstrip("/")
    marker = "/in/"
    if marker in cleaned:
        return cleaned.split(marker, 1)[1]
    return ""


def score_candidate(
    title: str, snippet: str, name: str, company: str, designation: str
) -> int:
    """Return a 0-100 confidence score that this result is the right person."""
    title_n = normalize(title)
    blob_n = normalize(f"{title} {snippet}")

    score = 0

    # --- name: require the meaningful name tokens to appear in the TITLE ---
    # Capped at 50 so that a full name match ALONE cannot clear the default
    # match threshold (60) -- a same-name/different-company person must still
    # be corroborated by the company before we trust it.
    name_toks = [t for t in tokens(name) if len(t) > 1]  # drop single-letter initials
    if name_toks:
        present = sum(1 for t in name_toks if t in title_n)
        frac = present / len(name_toks)
        score += int(50 * frac)          # up to 50 pts for a full name match in title

    # --- company match in title or snippet ---
    comp_toks = [t for t in tokens(company) if t not in STOPWORDS and len(t) > 1]
    if comp_toks:
        if any(t in title_n for t in comp_toks):
            score += 45                  # company in title = very strong corroboration
        elif any(t in blob_n for t in comp_toks):
            score += 30                  # company only in snippet
        # company given but NOT found anywhere -> no points (stays <= name-only)
    else:
        score += 10                      # no company to check; don't over-penalize

    # --- designation: small corroborating boost ---
    # Kept low (5) on purpose: a generic title like "Manager"/"Director" must
    # NOT, together with a name match alone, clear the threshold when the
    # company doesn't match -- company stays the real disambiguator.
    desg_toks = [t for t in tokens(designation) if t not in STOPWORDS and len(t) > 2]
    if desg_toks and any(t in blob_n for t in desg_toks):
        score += 5

    return min(score, 100)


def search_serper(api_key: str, query: str, session: requests.Session,
                  timeout: int = 30) -> list[dict]:
    headers = {"X-API-KEY": api_key, "Content-Type": "application/json"}
    resp = session.post(SERPER_URL, headers=headers,
                        json={"q": query, "num": 10}, timeout=timeout)
    resp.raise_for_status()
    return resp.json().get("organic", []) or []


def search_google_cse(api_key: str, cx: str, query: str,
                      session: requests.Session, timeout: int = 30) -> list[dict]:
    params = {"key": api_key, "cx": cx, "q": query, "num": 10}
    resp = session.get(GOOGLE_CSE_URL, params=params, timeout=timeout)
    resp.raise_for_status()
    items = resp.json().get("items", []) or []
    # normalize to serper-like shape
    return [{"link": it.get("link", ""), "title": it.get("title", ""),
             "snippet": it.get("snippet", "")} for it in items]


def _search_candidates(backend_call, query, name, company, title):
    """Run one query; return the best-scored LinkedIn candidate, None, or a
    ('error', message) tuple describing why the request failed."""
    try:
        results = backend_call(query)
    except requests.HTTPError as exc:
        code = exc.response.status_code if exc.response is not None else "?"
        body = (exc.response.text[:120] if exc.response is not None else "")
        return ("error", f"HTTP {code} {body}")
    except requests.RequestException as exc:
        return ("error", f"{type(exc).__name__}: {str(exc)[:120]}")
    best = None
    for item in results:
        link = item.get("link", "")
        if "linkedin.com/in/" not in link.lower():
            continue
        m = LINKEDIN_IN_RE.search(link)
        url = m.group(0) if m else link
        sc = score_candidate(item.get("title", ""), item.get("snippet", ""),
                             name, company, title)
        if best is None or sc > best.match_score:
            best = LookupResult(
                linkedin_url=url.split("?")[0],
                linkedin_id=linkedin_id_from_url(url),
                match_score=sc,
                matched_title=item.get("title", ""),
            )
    return best


def resolve_profile(backend_call, name, company, title, min_score) -> LookupResult:
    """Search + verify a single person, escalating to looser queries on a miss.

    We try increasingly relaxed queries and stop at the first that returns any
    LinkedIn profile. This rescues people whose company on LinkedIn differs from
    the spreadsheet (e.g. 'Hewlett Packard' vs 'HPE', 'HCL Tech' vs 'HCLTech').
    The company is still scored, so a profile found only via the relaxed query
    typically lands as ``low_confidence`` for you to confirm -- never silently
    accepted as a wrong match.
    """
    queries = [f'"{name}" "{company}" site:linkedin.com/in'] if company else []
    if company:
        queries.append(f'"{name}" {company} site:linkedin.com/in')  # company unquoted
    queries.append(f'"{name}" site:linkedin.com/in')                 # name only (recall)

    last_error = ""
    for query in queries:
        best = _search_candidates(backend_call, query, name, company, title)
        if isinstance(best, tuple) and best[0] == "error":
            last_error = best[1]
            continue
        if best is not None:
            best.lookup_status = (
                "matched" if best.match_score >= min_score else "low_confidence"
            )
            return best

    if last_error:
        return LookupResult(lookup_status=f"error: {last_error}")
    return LookupResult(lookup_status="not_found")


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


def make_backend(args) -> tuple:
    """Return (callable taking query -> results, session)."""
    session = requests.Session()
    if getattr(args, "insecure", False):
        session.verify = False
        requests.packages.urllib3.disable_warnings()  # type: ignore
        print("(insecure mode) SSL certificate verification disabled")
    if args.backend == "serper":
        key = os.environ.get("SERPER_API_KEY", "").strip()
        if not key:
            print("ERROR: set SERPER_API_KEY (get one free at https://serper.dev).",
                  file=sys.stderr)
            sys.exit(2)
        return (lambda q: search_serper(key, q, session)), session
    else:  # google official CSE
        key = os.environ.get("GOOGLE_API_KEY", "").strip()
        cx = os.environ.get("GOOGLE_CX", "").strip()
        if not key or not cx:
            print("ERROR: set GOOGLE_API_KEY and GOOGLE_CX for the --backend google.",
                  file=sys.stderr)
            sys.exit(2)
        return (lambda q: search_google_cse(key, cx, q, session)), session


def run(args: argparse.Namespace) -> int:
    in_path = Path(args.input)
    if not in_path.exists():
        print(f"ERROR: input file not found: {in_path}", file=sys.stderr)
        return 2

    if in_path.suffix.lower() in (".csv", ".txt"):
        df = pd.read_csv(in_path, dtype=str)
    else:
        df = pd.read_excel(in_path, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]  # tolerate trailing/leading spaces
    if args.name_col not in df.columns:
        print(f"ERROR: name column '{args.name_col}' not found. "
              f"Columns: {list(df.columns)}", file=sys.stderr)
        return 2
    has_company = args.company_col in df.columns
    has_title = args.title_col in df.columns

    backend_call, _ = make_backend(args)

    cache_path = Path(args.cache) if args.cache else None
    cache = load_cache(cache_path)
    retry_statuses = {s.strip() for s in args.retry.split(",") if s.strip()}

    results: list[dict] = []
    matched = low = missing = errored = 0
    error_samples: dict[str, int] = {}
    total = len(df)

    for i, row in df.iterrows():
        name = str(row[args.name_col]).strip()
        company = str(row[args.company_col]).strip() if has_company and pd.notna(row[args.company_col]) else ""
        title = str(row[args.title_col]).strip() if has_title and pd.notna(row[args.title_col]) else ""

        if not name or name.lower() == "nan":
            results.append(asdict(LookupResult(lookup_status="error: missing name")))
            errored += 1
            continue

        cache_key = f"{name}|{company}|{title}".lower()
        cached = cache.get(cache_key)
        # Re-run if not cached, or if the cached status is one we were asked to retry.
        cached_status = (cached or {}).get("lookup_status", "")
        force = any(cached_status.startswith(s) for s in retry_statuses) if retry_statuses else False
        if cached is not None and not force:
            res_dict = cached
        else:
            res = resolve_profile(backend_call, name, company, title, args.min_score)
            res_dict = asdict(res)
            cache[cache_key] = res_dict
            save_cache(cache_path, cache)
            time.sleep(args.sleep)

        status = res_dict["lookup_status"]
        matched += status == "matched"
        low += status == "low_confidence"
        missing += status == "not_found"
        if status.startswith("error"):
            errored += 1
            error_samples[status] = error_samples.get(status, 0) + 1

        results.append(res_dict)
        print(f"[{i + 1}/{total}] {name} @ {company or '?'} -> {status} "
              f"(score {res_dict['match_score']}) {res_dict['linkedin_url']}")

    out_df = pd.concat([df.reset_index(drop=True), pd.DataFrame(results)], axis=1)
    out_path = Path(args.output)
    out_df.to_excel(out_path, index=False)

    print("\n=== Summary ===")
    print(f"  matched (>= {args.min_score}): {matched}")
    print(f"  low_confidence (review):    {low}")
    print(f"  not_found:                  {missing}")
    print(f"  errors:                     {errored}")
    if error_samples:
        print("  most common error(s):")
        for msg, n in sorted(error_samples.items(), key=lambda kv: -kv[1])[:3]:
            print(f"    {n}x  {msg}")
        if any(s in m for m in error_samples for s in ("429", "credit", "402")):
            print("  -> looks like Serper rate-limit / out of credits. "
                  "Check https://serper.dev dashboard.")
        print("  (these errors are cached; re-run with --retry error to redo them.)")
    print(f"  written to:                 {out_path}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("input", help="Input .xlsx file")
    p.add_argument("-o", "--output", default="output.xlsx")
    p.add_argument("--name-col", default="Name")
    p.add_argument("--title-col", default="Designation")
    p.add_argument("--company-col", default="Company")
    p.add_argument("--backend", choices=["serper", "google"], default="serper",
                   help="Search backend. 'serper' (default) or official 'google' CSE.")
    p.add_argument("--min-score", type=int, default=60,
                   help="Accept matches at/above this confidence score (0-100).")
    p.add_argument("--retry", default="",
                   help="Comma list of cached statuses to RE-RUN, ignoring the "
                        "cache, e.g. --retry not_found,error or --retry low_confidence. "
                        "Use after improving the script or fixing company names.")
    p.add_argument("--sleep", type=float, default=0.5,
                   help="Seconds between searches.")
    p.add_argument("--cache", default=".google_lookup_cache.json",
                   help="Cache file to avoid repeat charges. Pass '' to disable.")
    p.add_argument("--insecure", action="store_true",
                   help="Disable SSL verification (last resort for corporate "
                        "HTTPS-inspecting proxies).")
    return p


if __name__ == "__main__":
    sys.exit(run(build_parser().parse_args()))
