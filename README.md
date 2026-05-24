# magic_excel2word — learn-mode prototype

Traceable pipeline that pairs a historical Excel workbook with its finished
Word report, locates every visible Word number, and produces reviewable
mapping artifacts. The current scope is **learn mode only** — no production
rendering, no LLM, no GUI. See `CLAUDE.md` for the full design rules.

> ⚠️ **Status: prototype.** Run only against synthetic samples until the
> trust gate (`learn --strict`) passes on a real pair you have reviewed.

## Quickstart

```bash
# 1. Set up a virtualenv and install pinned dependencies.
python -m venv .venv
source .venv/bin/activate          # on Windows: .venv\Scripts\activate
pip install -r requirements.txt

# 2. Generate a paired synthetic Excel + Word sample.
python -m src.main generate-synthetic --out samples/synthetic

# 3. Run learn mode against the pair.
python -m src.main learn \
    --excel samples/synthetic/historical.xlsx \
    --word  samples/synthetic/finished_report.docx \
    --out   output

# 4. Run the test suite.
python -m pytest
```

## Learn-mode artifacts

`learn` writes four files under `--out`:

| File | Purpose |
| --- | --- |
| `mapping_review.xlsx` | One row per Word number with the top Excel candidate, confidence, interpretation, value/context scores, and overlap tokens. The substrate for the (future) human-confirmed mapping workflow. |
| `confidence_report.md` | Markdown overview that **leads with what needs review** — UNRESOLVED, LOW, ambiguous picks, and EXCLUDED-by-policy rows. |
| `auto_mapping.yml` | Machine-readable map keyed by stable `word_id` (`word_0001`, `word_0002`, …). Every visible Word number is present — including UNRESOLVED, LOW, and EXCLUDED — with location, raw text, recommended Excel source (when available), transform metadata, and a `review_status` (`pending_review`, `needs_review`, `needs_source`, `audited_excluded`). |
| `converted_template.docx` | Copy of the original report where only HIGH/MEDIUM tokens have been replaced by `{{ word_NNNN }}` placeholders. LOW, UNRESOLVED, and EXCLUDED values are intentionally **left visible** so a reviewer can still audit them; any safe replacement that could not be applied (offset/raw drift) is recorded in `auto_mapping.yml`'s `placeholder_status` field. |

The console summary mirrors the report header and, when relevant, lists
sample entries from each problem bucket so you can spot issues without
opening the artifacts first.

> The template builder produces a **review substrate**, not a production
> Word renderer. `{{ word_NNNN }}` placeholders are a contract for a
> future deterministic renderer to honor once a human-confirmed
> `confirmed_mapping.yml` exists; no `run` subcommand exists today.

## Confidence statuses

Every visible Word number lands in exactly one bucket. Coverage is reported
over **eligible** numbers (i.e. excluding `EXCLUDED`), but the total visible
count is always shown so nothing can disappear silently.

| Status | Meaning |
| --- | --- |
| `HIGH` | A single Excel cell matches the value (under some unit interpretation) **and** the row/column context shares a strong label overlap. Safe to confirm. |
| `MEDIUM` | A value+context match exists but is either looser (e.g. 1% rounding) or tied with a runner-up. Needs human review before confirmation. |
| `LOW` | A value match exists but context overlap is weak. **Not safe** to confirm without manual checking. |
| `UNRESOLVED` | No Excel cell matches under any unit interpretation. The Word number has no known source. |
| `EXCLUDED` | The number was captured for the audit trail but the matcher deliberately skipped it (e.g. date/period markers such as `5月`, `2026年`, `第20周`). Reviewers should still scan the EXCLUDED section to confirm the policy applied correctly. |

`EXCLUDED` rows are visible in both `mapping_review.xlsx` and the dedicated
**“EXCLUDED by explicit policy”** section of `confidence_report.md` —
they are part of the audit, not silent drops.

## Trust gate: `learn --strict`

`learn` has two modes:

- **Default (exploratory).** Always writes artifacts. If any eligible
  number is `UNRESOLVED` or `LOW`, prints a loud warning to stderr but
  still exits `0`. Use this when you’re iterating on the matcher or
  exploring an unfamiliar report pair.
- **`--strict`.** Same artifacts, but exits non-zero (code `3`) if any
  eligible number is `UNRESOLVED` or `LOW`. Use this as a gate before
  pointing the tool at real company data or wiring it into automation.

```bash
# Will exit 0 with a warning on the synthetic corpus (it has deliberate
# UNRESOLVED cases like "提升15%" and "提升0.70个百分点").
python -m src.main learn --excel ... --word ... --out output

# Will exit 3 on the same corpus — that is the expected behaviour until
# every eligible Word number has been mapped or explicitly excluded.
python -m src.main learn --excel ... --word ... --out output --strict
```

Other exit codes: `2` if `--excel` or `--word` is missing.

## What this prototype is **not**

- Not a renderer. There is no `run` subcommand yet; production Word output
  must wait until human-confirmed mappings + a deterministic template path.
- Not an LLM client. Matching is deterministic; any future reranker must
  layer on top of, not replace, the candidate list.
- Not safe for real company files until `--strict` passes and a reviewer
  has signed off on the mapping. Keep real data out of the repo per
  `CLAUDE.md`’s privacy rules.
