# magic_excel2word — learn-mode prototype + deterministic renderer

Traceable pipeline that pairs a historical Excel workbook with its finished
Word report, locates every visible Word number, and produces reviewable
mapping artifacts. After a human-confirmed mapping is in hand the pipeline
can resolve it against a NEW Excel workbook (`run-preview`) and then
**deterministically render** a new Word report (`render-docx`) by
substituting `{{ word_NNNN }}` placeholders with display text inferred
from the historical raw token shape — still **no LLM, no GUI, no
Microsoft Word automation**. See `CLAUDE.md` for the full design rules.

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

# 4. Cross-check the four artifacts agree with each other.
python -m src.main validate-artifacts --out output

# 5. After a human reviews `output/mapping_review.xlsx`, promote
#    confirmed rows into `confirmed_mapping.yml`. Will fail (exit 5)
#    on a fresh, unreviewed file — that is the expected gate.
python -m src.main confirm-mapping \
    --auto   output/auto_mapping.yml \
    --review output/mapping_review.xlsx \
    --out    output/confirmed_mapping.yml

# 6. Preview a NEW period against the confirmed mapping.
python -m src.main run-preview \
    --excel     samples/synthetic/historical.xlsx \
    --confirmed output/confirmed_mapping.yml \
    --out       output/run_preview

# 7. Deterministically render the new Word report.
python -m src.main render-docx \
    --template       output/converted_template.docx \
    --run-validation output/run_preview/run_validation.xlsx \
    --out            output/new_report.docx

# 8. Run the test suite.
python -m pytest
```

## Learn-mode artifacts

`learn` writes four files under `--out`. Every row in every file carries
the same stable `word_id` (`word_0001`, `word_0002`, …), so the artifacts
can be cross-referenced by Word ID and verified by
`validate-artifacts` (see below).

| File | Purpose |
| --- | --- |
| `mapping_review.xlsx` | One row per Word number with the top Excel candidate, confidence, interpretation, value/context scores, and overlap tokens. The leading columns (`Word ID`, `Review Status`, `Placeholder Status`, `Placeholder`) cross-reference `auto_mapping.yml` and the converted template — a reviewer can read the audit story off the first few columns before drilling into candidate details. The substrate for the (future) human-confirmed mapping workflow. |
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

## Cross-artifact consistency: `validate-artifacts`

`learn` writes four files; `validate-artifacts` re-reads them and proves
they tell the same story about every Word number — that is the contract
the (future) human-confirmation step will depend on.

```bash
python -m src.main validate-artifacts --out output
```

What it proves on success:

- every `word_id` is **unique** within both `mapping_review.xlsx` and
  `auto_mapping.yml`;
- there is **one XLSX review row per YAML mapping** (and vice versa);
- shared rows **agree on `location`, raw token, and confidence/status**
  — no drift between the human-facing XLSX and the machine-readable YAML;
- a `{{ word_NNNN }}` placeholder appears in `converted_template.docx`
  **only when** the YAML row marks `placeholder_status: applied`. LOW,
  UNRESOLVED, and EXCLUDED tokens must **never** carry a placeholder
  anywhere — that is the audit guarantee CLAUDE.md mandates;
- if `auto_mapping.yml` reports any UNRESOLVED or EXCLUDED counts,
  `confidence_report.md` **mentions them by name**, so a reviewer who
  reads only the Markdown summary still sees what needs attention.

Exit codes: `0` on success, `4` on consistency failure (with a per-issue
list on stdout), `2` if `--out` is missing. The command never writes;
it only reads. Treat a failure as a directive to look at the artifacts,
not to re-run the pipeline.

## Reviewer handoff: `confirm-mapping`

`learn` writes a candidate mapping; `confirm-mapping` is the human-in-the-loop
gate that promotes reviewer-approved rows into `confirmed_mapping.yml`. The
two artifacts a reviewer touches are:

- `output/mapping_review.xlsx` — the per-row review sheet. The trailing four
  columns are **blank by default** and only this command reads them back:

  | Column | Meaning |
  | --- | --- |
  | `Reviewer Decision` | One of `confirm`, `reject`, or blank. Case-insensitive, whitespace-trimmed. Anything else is treated as an invalid decision. |
  | `Reviewer Notes` | Free-text; round-tripped into both the confirmed entry and the `review_required` list so the reason for a hold is preserved. |
  | `Confirmed Sheet` / `Confirmed Cell` | Optional override. Leave both blank to confirm the matcher's recommended source. Fill in *both* (sheet AND cell) to pick one of the YAML `alternatives` instead. A partial override (only one of the two) is invalid by design. |

- `output/auto_mapping.yml` — the machine truth from `learn`. This is the only
  source of legal Excel candidates: an override must name a `(sheet, cell)`
  that already exists as either the row's `recommended_source` or one of its
  `alternatives`. Inventing a fresh cell address is rejected because it would
  bypass the matcher's value/context agreement.

```bash
python -m src.main confirm-mapping \
    --auto   output/auto_mapping.yml \
    --review output/mapping_review.xlsx \
    --out    output/confirmed_mapping.yml \
    [--allow-incomplete]
```

`confirmed_mapping.yml` has three buckets. Every row in `auto_mapping.yml`
ends up in **exactly one** of them — that is the no-silent-omission contract:

| Bucket | What goes here |
| --- | --- |
| `confirmed_mappings` | HIGH/MEDIUM rows the reviewer explicitly marked `confirm` whose source is either the recommended pick (blank override) or matches an alternative (full override). Carries `word_id`, `location`, `raw`, `value`, `unit`, `status`/`confidence`, `recommended_source`, `confirmed_source`, `source_origin` (`recommended`, `reviewer_override:recommended`, `reviewer_override:alternative`), `transform` metadata, `reviewer_decision`, and `reviewer_notes`. |
| `review_required` | Every eligible row that did not pass: blank decision (`blank_decision`), reject (`rejected`), invalid/incomplete override (`invalid_override:…`, `incomplete_override:…`), LOW (`low_confidence_cannot_confirm`), UNRESOLVED (`unresolved_no_candidate`), HIGH/MEDIUM whose template placeholder was skipped by `learn` (`non_renderable_template_skip:…` — the converted .docx has no placeholder for a renderer to substitute into, so promoting the row would silently omit the metric at render time), or any unrecognised decision string (`invalid_decision:…`). Includes the recommended source and alternatives so the reviewer can decide what to do next. |
| `audit_only_excluded` | Every matcher-EXCLUDED row (date/period markers). These never become confirmed mappings — even if a reviewer types `confirm`, the decision is recorded for the audit trail but the row stays audit-only. |

### Exit codes

- `0` — every eligible row is in `confirmed_mappings` or `audit_only_excluded`,
  or `--allow-incomplete` was passed.
- `2` — `--auto` or `--review` is missing, or the two artifacts are out of
  sync (e.g. a Word ID in the XLSX has no entry in the YAML; run
  `validate-artifacts` and re-run `learn` if needed).
- `5` — at least one row is in `review_required`. The output YAML is still
  written so a reviewer can see exactly what is blocking, but the gate
  refuses to claim a complete mapping.

### `--allow-incomplete`

Exploratory escape hatch only. Writes `confirmed_mapping.yml` and exits `0`
even when `review_required` is non-empty. The YAML records
`summary.allow_incomplete: true` and `summary.complete: false`, so any
downstream consumer can refuse to render from a partial file. Never use this
as the default in automation pointing at real data.

## Run-mode preview: `run-preview`

`run-preview` is the bridge from a confirmed mapping to a rendered Word
report. It takes a NEW period's Excel workbook plus the existing
`confirmed_mapping.yml` and asks one question per confirmed row: *does
the recorded `(sheet, cell)` still hold a numeric value the transform
knows how to interpret?* It writes a per-row validation table —
`run_validation.xlsx` — that `render-docx` then consumes to substitute
placeholders in the converted template.

```bash
python -m src.main run-preview \
    --excel     path/to/new_period.xlsx \
    --confirmed output/confirmed_mapping.yml \
    --out       output/run_preview
```

The artifact is `run_validation.xlsx` under `--out`. Columns:

| Column | Meaning |
| --- | --- |
| `Word ID` / `Word Location` | Stable join keys from learn mode. |
| `Word Context` / `Word Raw Token` / `Word Unit` | What the historical Word report said at this position — surfaced so a reviewer can sanity-check that the new Excel value reads naturally back into the same sentence. |
| `Source Sheet` / `Source Cell` | The confirmed `(sheet, cell)` resolved against the new workbook. |
| `Raw Excel Value` | The literal numeric value the new workbook holds at that cell. |
| `Generated Value` | `Raw Excel Value` after applying the recorded transform (e.g. `excel / 10000` for `万元→base_unit`). This is the unrounded number a future renderer would display; formatting is deferred. |
| `Transform Interpretation` | The interpretation label carried over from learn time (`as_written`, `万元→base_unit`, `%→decimal`, …). |
| `Confidence` | Original learn-time confidence (`HIGH` / `MEDIUM`). |
| `Status` | Per-row run-preview outcome — see below. |
| `Detail` | Free-text reason populated when `Status` is not `ok`. |

### Fail-loud gates

`run-preview` exits non-zero on any of:

- `confirmed_mapping.yml` cannot prove it is complete. The check is
  fail-CLOSED: the `summary` block must exist, `summary.complete` must
  be the literal boolean `true` (a missing key, `null`, the string
  `"true"`, or `1` all refuse), `summary.allow_incomplete` must not be
  `true`, `review_required` must be empty, and there must be at least
  one `confirmed_mappings` entry. The CLI exits `6` **without writing
  an artifact**; the input is unusable.
- Any confirmed source `(sheet, cell)` is missing in the new workbook,
  is empty, or holds a non-numeric value. The CLI exits `7` and the
  artifact is **still written** so the reviewer can see which rows broke
  and why. Per-row statuses surfaced in this case:
  `missing_sheet`, `missing_cell`, `non_numeric_cell`,
  `missing_confirmed_source`.
- The transform interpretation on a confirmed row is unknown to the v1
  transform table or absent. Per-row statuses: `transform_unknown`,
  `missing_transform`. CLI exits `7` and the artifact is still written.

Other exit codes: `2` if `--excel` or `--confirmed` is missing.

### v1 transform coverage

The inverse transforms used to derive `Generated Value` from
`Raw Excel Value` mirror the candidate-value table in
`src/number_normalizer.py`. Currently supported:

- `as_written` (no conversion)
- `万元→base_unit`, `万人→base_unit`, `万次→base_unit`, `万个→base_unit`,
  `万单→base_unit`, `万→base_unit` (Excel ÷ 10,000)
- `亿元→base_unit`, `亿单→base_unit` (Excel ÷ 100,000,000)
- `千元→元` (÷ 1,000), `百万元→元` (÷ 1,000,000)
- `%→decimal` (× 100), `‰→decimal` (× 1,000)
- `base→万` (× 10,000)

Any interpretation outside this table is a v1 limitation — `run-preview`
fails loudly rather than silently producing an incorrect rendered number.

### Limitations of v1

- The `Generated Value` is the unrounded numeric value. Number-formatting
  to match the original report's display (`23,456.79万元` vs `23456.789`)
  is the renderer's job — see [`render-docx`](#deterministic-render-render-docx)
  below.
- Date/period markers (`audit_only_excluded`) and `review_required` rows
  do not participate in the preview at all — by design. Only confirmed
  rows are extracted.
- When `source_origin` is `reviewer_override:alternative`, the confirmed
  entry carries that alternative's own transform interpretation rather
  than the recommended pick's — so `run-preview` applies the correct
  unit factor for the cell the reviewer actually chose. The invariant
  is pinned by `test_override_to_alternative_carries_alternative_transform`
  and `test_alternative_override_applies_correct_transform_end_to_end`.

## Deterministic render: `render-docx`

`render-docx` is the production Word-output step. It takes
`converted_template.docx` (from `learn`) plus `run_validation.xlsx`
(from `run-preview`) and substitutes every `{{ word_NNNN }}` placeholder
with deterministic display text — preserving the historical raw token's
unit (`万元`, `亿元`, `%`, `‰`, `元`, `单`, `人`, `次`, `个`), comma
grouping, decimal precision, and sign style (explicit `+`/`-` or
accounting parens). No LLM, no GUI, no network, no Microsoft Word
automation.

```bash
python -m src.main render-docx \
    --template       output/converted_template.docx \
    --run-validation output/run_preview/run_validation.xlsx \
    --out            output/new_report.docx
```

On success it writes two files alongside each other:

| File | Purpose |
| --- | --- |
| `new_report.docx` | The rendered Word report. Every `{{ word_NNNN }}` placeholder has been replaced; no placeholder strings remain. |
| `render_log.yml` | One entry per `word_id` with source sheet/cell, raw Excel value, generated value, raw token, unit, display text, replacement count, and per-row status. The full audit a reviewer needs to verify the docx contains every confirmed metric and nothing else. |

### Fail-loud gates

`render-docx` exits non-zero on any of:

- A `run_validation.xlsx` row whose `Status` is not `ok` — render-docx
  refuses to substitute a value whose run-preview gate already failed.
- A `word_id` whose `Generated Value` is missing or that appears in more
  than one validation row. Every confirmed word_id must have exactly
  one generated value.
- A `{{ word_NNNN }}` placeholder in the template whose `word_id` is
  absent from `run_validation.xlsx` (would render `{{ word_NNNN }}`
  literally into the docx). Exit code `8`, no artifact written.
- A `word_id` in `run_validation.xlsx` that has no matching placeholder
  in the template (would silently drop a confirmed metric from the
  rendered report). Exit code `9`, no artifact written.
- A historical raw token that doesn't parse into the v1 numeric shape,
  or whose unit drifts from the validation `Word Unit` field. Exit
  code `9` with the offending `word_id` named in stderr — the renderer
  refuses to guess.

Per-row failures halt the whole render: a partial docx that silently
omits or mis-renders a confirmed metric is exactly the risk CLAUDE.md
forbids. On any gate failure the docx and log are **not** written and
the offending word_ids are listed on stderr.

Other exit codes: `2` if `--template` or `--run-validation` is missing.

### Duplicate placeholder accounting

A single confirmed `word_id` may legitimately appear in multiple
sentences of the historical report (e.g. a headline number referenced
again in a commentary paragraph). The renderer substitutes **every**
occurrence and the `render_log.yml` records the actual replacement
count per `word_id` (`placeholder_occurrences`). No occurrence is
silently skipped.

### Display formatting v1

The shape the renderer preserves is whatever the historical raw token
implies:

| Historical raw token | Generated value | Rendered display |
| --- | --- | --- |
| `23,456.79万元` | `23456.789` | `23,456.79万元` (comma + 2 decimals + 万元) |
| `1.23亿元` | `12.3456789` | `12.35亿元` |
| `37.50%` | `38.45` | `38.45%` |
| `+5.00%` | `7.5` | `+7.50%` (explicit `+` preserved on positives) |
| `-1,234.56` | `-1234.567` | `-1,234.57` (HALF_UP rounding) |
| `(1,234.56)` | `-1234.56` | `(1,234.56)` (accounting-paren style) |
| `100` | `99.5` | `100` (integer token → HALF_UP rounding) |

The renderer uses `ROUND_HALF_UP` rather than Python's banker's
rounding default so business display matches the original report's
convention.

### Limitations of v1

- The renderer touches paragraph and top-level table cell text only —
  the same surfaces `learn`'s template builder writes into. Headers,
  footers, footnotes, text boxes, chart data, and nested tables are
  out of scope.
- Setting `paragraph.text = …` collapses inline runs into a single
  run. The template builder already does this when inserting
  placeholders, so the renderer is not regressing run-level formatting;
  a future styled-renderer would substitute run-aware.
- Only the v1 unit set (`元`, `万元`, `亿元`, `千元`, `百万元`,
  `万单`, `亿单`, `万人`, `万次`, `万个`, `万`, `单`, `人`, `次`,
  `个`, `%`, `‰`) is recognised. A raw token outside this shape is
  surfaced as a `format_inference_failed` row, not silently rendered.

## What this prototype is **not**

- Not an LLM client. Matching and rendering are deterministic; any
  future reranker must layer on top of, not replace, the candidate
  list and the deterministic formatter.
- Not a Microsoft Word automation. The renderer uses `python-docx`
  directly and works on macOS, Linux, or Windows without Word
  installed.
- Not safe for real company files until `--strict` passes and a reviewer
  has signed off on the mapping. Keep real data out of the repo per
  `CLAUDE.md`'s privacy rules.
