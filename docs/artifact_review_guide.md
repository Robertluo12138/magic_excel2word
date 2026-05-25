# Artifact review guide — manual real-file pilot

This guide tells a reviewer **exactly what to inspect** in each pilot
artifact, using **aggregate counts, status enums, and schema/presence
checks only**. It is the per-artifact companion to
[`real_file_pilot.md`](real_file_pilot.md) §3 ("Artifacts a reviewer
must inspect") and to the redacted reviewer form in
[`pilot_result_template.md`](pilot_result_template.md): use those two
documents for the workflow, use this one when you sit down with an
artifact open and need a checklist of what to look at.

> ⚠️ **Documentation-only.** This guide changes no pipeline behavior,
> runs no command, writes no artifact, and never reads a real file.
> Every check below is something a reviewer performs by eye against
> the artifact already on disk.

---

## 0. What you MUST NOT copy out of any artifact

The whole point of inspecting these artifacts by aggregate / status /
schema is so the review record stays free of report content. Before
you write **any** note, ticket comment, PR description, commit
message, chat paste, screenshot, or filled-in copy of
`pilot_result_template.md`, double-check that none of the following
has leaked into the repo or any doc tracked alongside it:

- **raw business numbers** — any literal numeric value carried by the
  historical or rendered report;
- **raw Word snippets** — any sentence, paragraph fragment, header
  string, or table cell text quoted from the historical or rendered
  Word report (this includes the `display_text` strings the renderer
  produced and the `raw` / `raw_token` strings captured at learn
  time);
- **raw Excel cell values** — any literal value held at any source
  cell in the historical or new-period workbook (including the
  `Raw Excel Value` and `Generated Value` columns of
  `run_validation.xlsx` and the matching fields in `render_log.yml`);
- **source sheet names or cell addresses** that identify real company
  data — the `Source Sheet` / `Source Cell` columns of
  `run_validation.xlsx`, the `source_sheet` / `source_cell` fields of
  `render_log.yml`, and the `recommended_source` / `alternatives`
  blocks of `auto_mapping.yml` all carry this content;
- **company identifiers** — company names, product names, customer
  names, brand names, internal team names, project codenames, or any
  other label that would tie an artifact to a specific organization;
- **individual `word_id` values** when they would correlate (via the
  reviewer's memory of the report) to a sensitive metric — aggregate
  `word_id` counts are fine, named individual `word_id`s are not;
- **reviewer notes** that quote the report's content verbatim — the
  `Reviewer Notes` column of `mapping_review.xlsx` and the
  `reviewer_notes` field of `confirmed_mapping.yml` are designed for
  free-form text and can easily carry content that should stay inside
  the artifact;
- **file paths beyond basenames** — the external pilot directory path
  may itself be sensitive; the reviewer records artifact basenames
  only, never the full absolute paths.

If a check below would require writing any of the above into the repo
or a doc to answer it, **stop**. Either record a non-go in
[`pilot_result_template.md`](pilot_result_template.md) §8 or escalate
to the sponsor pair — never paste the forbidden content "just to
explain the issue".

This list is a strict superset of the "What this template MUST NOT
contain" rules already pinned by
[`pilot_result_template.md`](pilot_result_template.md). It applies to
**every** artifact below, not only to the filled-in template, and to
**every** surface the reviewer might paste into — chat, ticket, PR,
commit, screenshot, or any file inside this repository.

---

## 1. `mapping_review.xlsx`

**Source:** `learn` (Stage 1). Per-row XLSX whose column headers come
from the canonical `REVIEW_HEADERS` list in
`src/mapping_reviewer.py`. The first column is the join key
`Word ID`; further along the row the trust-slice columns
`Review Status`, `Placeholder Status`, and `Placeholder` together
tell the audit story across `auto_mapping.yml` and
`converted_template.docx`; and the trailing four columns are the
reviewer-decision inputs `Reviewer Decision`, `Reviewer Notes`,
`Confirmed Sheet`, `Confirmed Cell` (blank by default on every
`learn` run; only `confirm-mapping` reads them back).

**Aggregate / status / schema checks:**

- **Schema.** The header row matches the canonical `REVIEW_HEADERS`
  list in `src/mapping_reviewer.py`: `Word ID`, `Word Location`,
  `Word Snippet`, `Word Raw Token`, `Word Value`, `Word Unit`,
  `Confidence`, `Review Status`, `Placeholder Status`, `Placeholder`,
  `Top Excel Sheet`, `Top Excel Cell`, `Top Excel Value`,
  `Top Row Context`, `Top Column Context`, `Interpretation`,
  `Value Score`, `Context Score`, `Overlap Tokens`,
  `# Alt Candidates`, `Note`, `Reviewer Decision`, `Reviewer Notes`,
  `Confirmed Sheet`, `Confirmed Cell`. `Word ID` is the first
  column; the trust-slice trio (`Review Status`, `Placeholder Status`,
  `Placeholder`) sits between the per-row word descriptor columns
  and the top-candidate columns; the reviewer-decision quartet is
  the trailing block. A missing or renamed column is a regression —
  file a bug, do not paper over it in the review.
- **Row count.** The total data row count equals the
  `Total visible Word numbers` count printed by the `learn` console
  summary and by `pilot-summary --out <pilot output>`.
- **`Word ID` shape and uniqueness.** Every value matches the
  `word_\d{4,}` shape and every value is unique. Two rows sharing a
  `Word ID` is a matcher regression.
- **`Review Status` distribution.** Every value is one of the four
  recognised review-status strings (`pending_review`, `needs_review`,
  `needs_source`, `audited_excluded`). Anything else is a regression.
  Aggregate the counts; record them via
  `pilot_result_template.md` §4 / §5, not row-by-row.
- **`Placeholder Status` distribution.** Every value is one of the
  matcher's recorded placeholder-status enums: `applied`,
  `skipped_low_confidence`, `skipped_unresolved`,
  `skipped_excluded`, `skipped_offset_out_of_range`,
  `skipped_raw_mismatch`, or `skipped_unknown` (the contract pinned
  by `_write_converted_template` in `src/template_builder.py`). The
  column is never empty — even EXCLUDED rows carry `skipped_excluded`
  (their matching `Review Status` is `audited_excluded`). Note the
  aggregate counts only.
- **`Placeholder` column.** Holds the literal `{{ word_NNNN }}`
  string when the row's `Placeholder Status` is `applied`, and is
  empty otherwise. Confirm the column is empty for every row whose
  `Placeholder Status` is a `skipped_*` reason — a non-empty cell
  paired with a skipped status is a writer regression.

**What you do NOT inspect here:**

- Do not transcribe candidate raw tokens, candidate cell values,
  overlap-token lists, or reviewer notes into any note. Eyeball them
  in the XLSX viewer and move on. The XLSX itself stays inside the
  external pilot directory.

---

## 2. `auto_mapping.yml`

**Source:** `learn` (Stage 1). Machine-readable map keyed by
`word_id`.

**Aggregate / status / schema checks:**

- **Top-level structure.** The YAML loads cleanly with
  `yaml.safe_load` and exposes `schema_version`, `summary`, and
  `mappings`. A missing top-level key is a regression.
- **`summary` block.** Contains `total` and `by_confidence` integer
  counts plus `placeholders_applied`. The `total` value equals the
  `Total visible Word numbers` count seen in the `learn` console
  summary, the `mapping_review.xlsx` row count, and `pilot-summary`.
- **`mappings` length.** Equals `summary.total` and equals the row
  count of `mapping_review.xlsx`. `validate-artifacts` already pins
  this; the human check here is a glance at both totals via
  `pilot-summary` to confirm they match.
- **`word_id` uniqueness.** No `word_id` appears twice across
  `mappings`.
- **`review_status` distribution.** Every entry carries a
  `review_status` from the same four-value set as
  `mapping_review.xlsx` (`pending_review`, `needs_review`,
  `needs_source`, `audited_excluded`). Aggregate the counts only.
- **`placeholder_status` distribution.** Every entry carries one of
  the same seven placeholder-status enum values listed in §1
  (`applied`, `skipped_low_confidence`, `skipped_unresolved`,
  `skipped_excluded`, `skipped_offset_out_of_range`,
  `skipped_raw_mismatch`, `skipped_unknown`); the field is never
  null. HIGH/MEDIUM entries may be `applied` or any `skipped_*`
  reason; LOW, UNRESOLVED, and EXCLUDED entries MUST NOT be
  `applied` — that is the audit guarantee `validate-artifacts`
  enforces by code. Every entry with `placeholder_status == applied`
  also carries a `placeholder` field of `"{{ word_NNNN }}"`; every
  other entry has `placeholder: null`.

**What you do NOT inspect here:**

- Do not copy `recommended_source`, `alternatives`, `raw`, `value`,
  `context.snippet`, or `context.label_context` from any entry into a
  note. Those fields exist for the matcher and the renderer;
  aggregate counts and status enums are enough for the human review.

---

## 3. `confirmed_mapping.yml`

**Source:** `confirm-mapping` (Stage 3). Promotes reviewer-approved
rows.

**Aggregate / status / schema checks:**

- **`summary.complete`.** Must be the literal boolean `true`. Any
  other value (`false`, missing key, the string `"true"`, `null`,
  `1`) means the gate did not clear — `run-preview` would refuse the
  file with exit `6`. Confirm with the YAML open, not by trusting the
  CLI exit alone.
- **`summary.allow_incomplete`.** Must be absent or the literal
  boolean `false`. The `--allow-incomplete` escape hatch is forbidden
  in a real-file pilot.
- **Bucket presence.** The file exposes three top-level lists:
  `confirmed_mappings`, `review_required`, `audit_only_excluded`.
- **`review_required` length.** Must be `0`. Any non-zero value
  halts the pilot here; record the count (not the contents) in §5 of
  `pilot_result_template.md`.
- **`confirmed_mappings` length.** Must be ≥ `1` — `run-preview`
  refuses an empty confirmed set with exit `6`. The length equals the
  row count of `run_validation.xlsx` after Stage 4.
- **`audit_only_excluded` length.** Equals the EXCLUDED count from
  `auto_mapping.yml`'s `summary.by_confidence` — the EXCLUDED policy
  is preserved through the reviewer handoff.
- **`source_origin` distribution.** Every confirmed entry carries
  `source_origin` ∈ {`recommended`, `reviewer_override:recommended`,
  `reviewer_override:alternative`}. Aggregate the counts only.

**What you do NOT inspect here:**

- Do not copy `recommended_source`, `confirmed_source`, `raw`,
  `value`, or `reviewer_notes` from any entry into a note. The
  reviewer-notes field is designed for free-form text and is the
  easiest accidental leak — re-read [§0](#0-what-you-must-not-copy-out-of-any-artifact)
  before pasting anything.

---

## 4. `run_validation.xlsx`

**Source:** `run-preview` (Stage 4). Per-row validation against the
new period's workbook.

**Aggregate / status / schema checks:**

- **Schema.** Column headers exactly match the canonical set
  documented in [`README.md` "run-mode preview"](../README.md#run-mode-preview-run-preview):
  `Word ID`, `Word Location`, `Word Context`, `Word Raw Token`,
  `Word Unit`, `Source Sheet`, `Source Cell`, `Raw Excel Value`,
  `Generated Value`, `Transform Interpretation`, `Confidence`,
  `Status`, `Detail`. A renamed column would later surface as
  `run_validation_schema` from `validate-render`; the eyes-on check
  here is so a schema regression is caught before Stage 6 runs.
- **`Status` distribution.** Every row's `Status` must be `ok`. If
  any row is not `ok`, halt the pilot — `render-docx` would refuse
  the artifact anyway with exit `8`/`9`.
- **Row count.** Equals `len(confirmed_mappings)` from
  `confirmed_mapping.yml`. The two artifacts must agree row-for-row
  by `Word ID`.
- **`Word ID` uniqueness.** No `Word ID` appears twice.
- **`Confidence` distribution.** Every value is either `HIGH` or
  `MEDIUM`. LOW, UNRESOLVED, and EXCLUDED rows cannot reach this
  artifact by construction; any other value is a regression.

**What you do NOT inspect here:**

- Do not transcribe `Raw Excel Value`, `Generated Value`,
  `Source Sheet`, `Source Cell`, `Word Context`, `Word Raw Token`,
  or `Detail` (if any) into a note. Their per-row values are the
  report itself; aggregate counts and the `Status` distribution are
  the only things you record.

---

## 5. `render_log.yml`

**Source:** `render-docx` (Stage 5). One entry per rendered
`word_id` under the top-level `replacements` list, plus a `summary`
block.

**Aggregate / status / schema checks:**

- **Top-level structure.** Loads cleanly with `yaml.safe_load` and
  exposes `schema_version`, `inputs`, `summary`, `out_docx`, and
  `replacements`.
- **`summary` block.** Contains integer `total_rows`, `ok`, `failed`,
  `total_replacements`, and `distinct_placeholder_word_ids`. The
  `failed` count must be `0`; `ok` must equal `total_rows`.
- **`replacements` length.** Equals the row count of
  `run_validation.xlsx`. `validate-render` pins this; the eyes-on
  check is a glance at both totals via `pilot-summary`.
- **Per-entry `status`.** Every entry carries `status = ok`. Any
  other value is a regression — `validate-render` would refuse it
  with exit `10`.
- **Per-entry required-field presence.** Every entry carries
  non-empty `generated_value`, `source_sheet`, `source_cell`, and
  `display_text`; absence of any of those four fields is a
  `render_log_missing_field` regression. The renderer also emits
  `raw_token` and `unit` for the audit trail; inspect their presence
  by eye, but do not describe them as part of the current
  `render_log_missing_field` gate unless `validate-render` is widened
  in code first.
- **Per-entry `placeholder_occurrences`.** Every entry's value is
  ≥ `1`. A zero count means the renderer silently dropped a
  confirmed metric or the log was hand-edited.
- **`word_id` uniqueness.** No `word_id` appears twice across
  `replacements`.

**What you do NOT inspect here:**

- Do not copy `generated_value`, `display_text`, `raw_excel_value`,
  `source_sheet`, `source_cell`, or `raw_token` from any entry into
  a note. The fields exist for the audit-trail walk; their per-entry
  values stay inside the artifact.

---

## 6. `new_report.docx`

**Source:** `render-docx` (Stage 5). The rendered Word report — the
actual deliverable.

**Aggregate / status / schema checks:**

- **No leftover placeholder.** Open in a docx viewer with macros
  disabled (LibreOffice with macro execution off, or a docx text
  dump) and search for the literal string `{{ word_`. The count must
  be `0`. `validate-render` pins this as `leftover_placeholder`, but
  human eyes are the last line of defense before the docx leaves the
  pilot directory.
- **Paragraph + table-cell scope.** The docx only carries surfaces
  the renderer writes into (paragraphs and top-level table cells per
  the v1 scope in
  [`milestone_1_readiness.md` §4](milestone_1_readiness.md#4-known-limitations--not-production-ready-yet)).
  If the report relies on numbers in headers, footers, footnotes,
  text boxes, chart data, or nested tables, record a non-go in §8
  of `pilot_result_template.md` — those surfaces are out of v1 scope.
- **Display-text presence sweep.** The `validate-render` gate
  already proves every `display_text` from `render_log.yml` appears
  in the docx the expected number of times. The reviewer's eyes-on
  check is a paragraph-by-paragraph skim to confirm the rendered
  text reads naturally back into the surrounding sentence — that is
  a human judgment no validator can perform.

**What you do NOT inspect here:**

- Do not transcribe any rendered number, sentence fragment, table
  cell, paragraph, or section heading from the docx into a note,
  screenshot, PR description, or chat. The docx itself is the
  deliverable — share it only with the sponsor pair, never paste
  excerpts into a public or tracked location.

---

## 7. `confidence_report.md`

**Source:** `learn` (Stage 1). Markdown overview that leads with
what needs review.

**Aggregate / status / schema checks:**

- **Section order.** The top sections lead with what needs review
  (`UNRESOLVED`, `LOW`, ambiguous, `EXCLUDED by explicit policy`).
  HIGH wins come after. A regression that hides problem buckets
  below HIGH wins is a matcher or reporter bug — file it.
- **Bucket count parity.** The counts named in each section header
  match `summary.by_confidence` from `auto_mapping.yml` and the row
  counts in `mapping_review.xlsx`. `Total = HIGH + MEDIUM + LOW +
  UNRESOLVED + EXCLUDED` must hold against the §4 counts of
  `pilot_result_template.md`.
- **EXCLUDED policy spot-check.** The EXCLUDED section lists only
  the date/period-marker categories the matcher is supposed to
  exclude. A real metric mis-classified as EXCLUDED is a matcher
  bug — file it, do not hand-edit `auto_mapping.yml`.

**What you do NOT inspect here:**

- Do not paste any sample Word token, paragraph fragment, or section
  body into a note. The Markdown overview lists sample tokens by
  design so a reviewer can locate them in the XLSX — those samples
  are report content and stay inside the artifact.

---

## 8. `pilot-summary` output

**Source:** `pilot-summary --out <pilot output>` (optional, any time
after Stage 1). Read-only, contractually redacted.

**Aggregate / status / schema checks:**

- **Paste-safe contract.** By contract this command's stdout never
  prints raw Word tokens, generated values, raw Excel values, source
  sheet/cell content, reviewer notes, individual `word_id` values,
  or file paths beyond basenames; the `--out` directory is reduced
  to its basename in the header. If any output you observe violates
  this contract, **stop** — that is a `pilot-summary` regression and
  the output is no longer safe to paste anywhere. File a bug instead
  of pasting the offending output into the bug report.
- **Per-stage aggregate counts.** Each section reports artifact
  basename, file size, and the per-stage aggregate counts (e.g.
  `Total visible Word numbers`, `Mapped`, `LOW`, `UNRESOLVED`,
  `EXCLUDED`, `confirmed_mappings`, `review_required`,
  `audit_only_excluded`, run-validation `Status` distribution,
  render-log `status` distribution). These are the only numbers you
  record in §4 and §5 of `pilot_result_template.md`.
- **Next-action hint.** Each stage emits exactly one next-action
  hint. A missing hint, or a hint that names a `word_id` or raw
  value, is a regression.

**What you do NOT inspect here:**

- The `pilot-summary` output is the **one** artifact whose stdout is
  designed to be safe to paste into a chat or ticket on real data.
  Even so, re-read the redaction contract above before pasting — a
  regression that leaks a value is exactly the failure mode the
  redaction was designed to prevent. If in doubt, do not paste.

---

## 9. Cross-artifact consistency the reviewer must observe

`validate-artifacts` and `validate-render` already pin
cross-artifact consistency by code. The reviewer's job is to
confirm the **aggregate counts** line up — never to re-derive
agreement by copying per-row content.

- The `Total visible Word numbers` count is identical across the
  `learn` console summary, `pilot-summary` output, the
  `mapping_review.xlsx` row count, the `auto_mapping.yml`
  `summary.total`, and the bucket-count parity sum in
  `confidence_report.md`.
- The `confirmed_mappings` length in `confirmed_mapping.yml` equals
  the row count of `run_validation.xlsx` and the `replacements`
  length of `render_log.yml`.
- Every observed CLI exit code matches the expected value in
  [`real_file_pilot.md` §2c](real_file_pilot.md#2c-expected-exit-codes)
  and the
  [exit-code map](command_reference.md#exit-code-map-at-a-glance).

If any count disagrees, halt the pilot and re-run from the producing
stage. Do not hand-edit any artifact to "make the count match" —
that bypasses every gate the rest of the workflow depends on, and is
the failure mode
[`real_file_pilot.md` §4](real_file_pilot.md#4-handling-unresolved-low-excluded-non-renderable-and-tampered-rows)
calls out as "tampered or drifted".

---

## 10. Recording the review

After walking every check above:

- Record the aggregate counts and per-stage exit codes in
  [`pilot_result_template.md`](pilot_result_template.md) §2, §4, and
  §5.
- Tick the §3 artifact-presence and §7 reviewer-attestation boxes
  only after the corresponding checks above pass.
- Re-read [§0](#0-what-you-must-not-copy-out-of-any-artifact) before
  sharing the filled-in template anywhere. If any field would
  require pasting forbidden content to answer, leave it blank and
  record a non-go in §8 instead.

This guide is **read-only and documentation-only**. It changes no
pipeline behavior, runs no command, writes no artifact, and is never
auto-populated by the pipeline. The reviewer walks it by hand with
the artifacts already on disk in the external pilot directory.
