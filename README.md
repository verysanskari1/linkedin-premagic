# LinkedIn profile finder

Given a spreadsheet of people (**name + designation + company**), find each
person's LinkedIn profile URL / public id and write the results back to Excel.

There are two scripts; pick one:

| Script | Method | Cost | Accuracy | When to use |
|--------|--------|------|----------|-------------|
| **`google_lookup.py`** | Google search (Serper.dev) + result verification | **2,500 free**, then ~$1/1k | High *with verification* | Cheap, no account/legal risk. **Recommended starting point.** |
| `linkedin_lookup.py` | People Data Labs enrichment API | Free 100/mo, then ~$0.28/match | Highest | When you need maximum hit rate and pay-per-match is fine |

> Why not scrape LinkedIn directly or use the official Google Custom Search API?
> Direct scraping violates LinkedIn's ToS and gets accounts banned. Google's
> official Custom Search JSON API is **closed to new signups** (retiring Jan
> 2027), so `google_lookup.py` uses Serper.dev, which returns real Google
> results without that restriction.

## Setup

```bash
pip install -r requirements.txt
```

## Input format

Any `.xlsx` with at least a name column (designation & company strongly
recommended for accuracy). Default column names are `Name`, `Designation`,
`Company` — override with flags if yours differ. See `sample_input.xlsx`.

## Usage — Google (recommended)

```bash
export SERPER_API_KEY="your_key"        # free key: https://serper.dev
python google_lookup.py input.xlsx -o output.xlsx \
    --name-col Name --title-col Designation --company-col Company \
    --min-score 60
```

Already have an official Google CSE key/engine?

```bash
export GOOGLE_API_KEY=...; export GOOGLE_CX=...
python google_lookup.py input.xlsx --backend google
```

## Usage — People Data Labs

```bash
export PDL_API_KEY="your_key"           # free 100/mo: https://peopledatalabs.com
python linkedin_lookup.py input.xlsx -o output.xlsx --min-likelihood 6
```

## Output

The input columns plus:

- `linkedin_url` — the matched profile URL
- `linkedin_id` — the public id (the `/in/<this>` slug)
- `match_score` (Google) / `match_likelihood` (PDL) — confidence
- `lookup_status` — one of:
  - `matched` — confident match, use it
  - `low_confidence` — a profile was found but it didn't clear the threshold; **review manually**
  - `not_found` — no LinkedIn profile surfaced
  - `error: ...` — API/network problem

## How accuracy is enforced (Google script)

For each person we search `"Name" "Company" site:linkedin.com/in`, then **score**
every candidate result rather than trusting result #1:

- up to **50 pts** if the name appears in the result title,
- **+40** (title) / **+25** (snippet) if the company matches,
- **+10** if the designation matches.

A full-name match *alone* tops out at 50, **below** the default accept threshold
of 60 — so a same-name-but-different-company person is flagged for review
instead of being wrongly accepted. Raise `--min-score` for stricter matching,
lower it for more (riskier) hits.

## Notes

- Results are cached (`.google_lookup_cache.json` / `.lookup_cache.json`) so
  re-runs don't re-pay for people already resolved. Delete the cache to force
  a refresh.
- Use `--sleep` to tune the delay between requests.
