# Real-file pilot — reviewer result template (redacted)

A **redacted** result form the reviewer fills in after walking the manual
real-file workflow in [`real_file_pilot.md`](real_file_pilot.md) against a
real Excel + Word pair. It captures only the **aggregate** evidence a sponsor
needs to judge go/no-go: per-stage exit codes, artifact presence, aggregate
coverage counts, reviewer decision counts, accepted known limitations, and
the final recommendation.

By design every field below is an integer count, a policy enum, a boolean
checkbox, or a one-line free-text label — there is **no** field that asks for
a raw Word token, a raw Excel value, a source sheet/cell address, or any
other piece of report content. A reviewer who needs to leak content to
answer a field should leave it blank and record a non-go in §8 instead.

> ⚠️ **Prototype.** This template captures sponsor-facing evidence from a
> single pilot, not a marketing readiness claim. The prototype contract in
> [`milestone_1_readiness.md`](milestone_1_readiness.md) still applies and
> `--allow-incomplete` must not have been used on `confirm-mapping`.

> Before filling this in, the reviewer must have already walked the
> rehearsal in [`synthetic_dry_run_checklist.md`](synthetic_dry_run_checklist.md)
> and the eight verification items in
> [`milestone_1_readiness.md` §5](milestone_1_readiness.md#5-what-a-reviewer-must-verify-before-any-real-file-pilot).

---

## What this template MUST NOT contain

The same template file lives in this repo, so any pilot that fills it in
must keep the rendered form free of:

- raw Word number tokens (the literal text of a number from the report);
- raw Excel cell values (the literal value at any source cell);
- source sheet names or cell addresses that identify real company data;
- company names, product names, customer names, internal team names;
- file paths beyond basenames;
- reviewer notes that quote the report's content verbatim;
- individual `word_id` values when they would correlate to a sensitive metric.

If a filled-in copy of this template ever ends up shared outside the
reviewer-plus-sponsor pair, double-check the rendered Markdown against the
list above before sharing.

---

## 1. Pilot identification

Aggregate identifiers only — do not embed company, product, customer, or
team names in the label.

| Field | Value |
| --- | --- |
| Pilot label (free-text, no company identifiers) | `<label>` |
| Date (YYYY-MM-DD) | `<date>` |
| Operator (name or internal role) | `<who ran the commands>` |
| Reviewer (name or internal role) | `<who walked the artifacts>` |
| Pipeline commit SHA (`git rev-parse HEAD`) | `<sha>` |

## 2. Per-stage exit codes

Record the literal CLI exit code observed at each stage. Expected values are
the contract pinned by
[`command_reference.md`](command_reference.md#exit-code-map-at-a-glance).
A non-zero exit at any stage halts the pilot — the discrepancy is recorded
in §8 as a non-go.

| Stage | Command | Expected | Observed |
| --- | --- | --- | --- |
| (0) Privacy preflight | `pilot-preflight` | `0` | `<code>` |
| (1) Strict learn | `learn --strict` | `0` | `<code>` |
| (2) Cross-check learn artifacts | `validate-artifacts` | `0` | `<code>` |
| (3) Confirm mapping (no `--allow-incomplete`) | `confirm-mapping` | `0` | `<code>` |
| (4) Run preview | `run-preview` | `0` | `<code>` |
| (5) Render docx | `render-docx` | `0` | `<code>` |
| (6) Validate render | `validate-render` | `0` | `<code>` |
| (optional) Paste-safe progress lens | `pilot-summary` | `0` | `<code>` |

## 3. Artifact presence

Tick each box only after the file exists at the expected path under the
pilot output directory. Do **not** record file sizes or content snippets
here — file presence is a boolean.

- [ ] `mapping_review.xlsx`
- [ ] `auto_mapping.yml`
- [ ] `converted_template.docx`
- [ ] `confidence_report.md`
- [ ] `confirmed_mapping.yml` (with `summary.complete: true` and
      `summary.allow_incomplete` absent or `false`)
- [ ] `run_validation.xlsx`
- [ ] `new_report.docx`
- [ ] `render_log.yml`

## 4. Aggregate coverage counts

Source the values from the `learn` console summary or
`pilot-summary --out <pilot output>`. Aggregate counts only — never paste
an individual row, raw value, or `word_id` here.

| Bucket | Count |
| --- | --- |
| Total visible Word numbers | `<count>` |
| Mapped (HIGH + MEDIUM) | `<count>` |
| LOW | `<count>` |
| UNRESOLVED | `<count>` |
| EXCLUDED (date/period markers) | `<count>` |

`Total = Mapped + LOW + UNRESOLVED + EXCLUDED` must hold; the eligible
denominator is `Total − EXCLUDED`. A real-file pilot is only ready to ship
if every eligible row was mapped to HIGH/MEDIUM **and** confirmed in §5.

## 5. Reviewer decision counts

Source these from the `summary` block of `confirmed_mapping.yml`. Counts
only — never paste reviewer notes, `word_id` values, or override sheet/cell
content verbatim.

| Bucket | Count |
| --- | --- |
| `confirmed_mappings` (HIGH/MEDIUM confirmed) | `<count>` |
| `review_required` (blank decision / reject / LOW / UNRESOLVED / invalid or incomplete override / non-renderable template skip) | `<count>` |
| `audit_only_excluded` (matcher-EXCLUDED, always audit-only) | `<count>` |

The confirm-mapping gate is only cleared when `review_required` is `0`
**and** `summary.complete` is the literal boolean `true` **and**
`summary.allow_incomplete` is absent or `false`.

## 6. Known limitations accepted for this pilot

Tick a box only if the reviewer has explicitly confirmed the limit is
acceptable **for the specific report being trialed**. The contract is
[`milestone_1_readiness.md` §4](milestone_1_readiness.md#4-known-limitations--not-production-ready-yet).

- [ ] Surface scope: the report's numeric surface fits within paragraph and
      top-level table cells only (no headers, footers, footnotes, text
      boxes, chart data, or nested tables).
- [ ] Run-level styling: any inline-run collapse on substituted
      placeholders is acceptable for this trial.
- [ ] Unit table: every unit in the report is inside the v1 unit set
      (`元`, `万元`, `亿元`, `千元`, `百万元`, `万单`, `亿单`, `万人`,
      `万次`, `万个`, `万`, `单`, `人`, `次`, `个`, `%`, `‰`).
- [ ] No reviewer UI: the reviewer was willing to edit
      `mapping_review.xlsx` columns by hand.
- [ ] No semantic fallback: the matcher's deterministic value+context
      heuristic is sufficient for this trial — no LLM rerank was needed.
- [ ] No real-file CI coverage: the reviewer accepts that the acceptance
      smoke runs against synthetic fixtures only and this pilot is the
      first real-file evidence.

## 7. Reviewer attestations

A GO recommendation in §8 requires every box below to be ticked. These
are the human-eyes-on checks no test can perform.

- [ ] A human has visually skimmed every number in `new_report.docx` and
      confirmed each matches the corresponding `display_text` and
      `Generated Value` recorded in `render_log.yml`.
- [ ] No file under the external pilot directory (e.g. `~/pilot_data/`)
      is staged, tracked, or referenced in any commit on any branch.
- [ ] `git status` in the repo shows only expected source-code changes
      — no real `.xlsx`, `.docx`, `.yml`, or `.md` artifact has leaked
      into the tree.
- [ ] This template, as filled in, contains no raw Word tokens, raw
      Excel values, source sheet/cell content, company identifiers, or
      reviewer-note quotes from the report.

## 8. Final go / no-go recommendation

Pick exactly one. A GO recommendation requires every observed exit code
in §2 to equal its expected value, every box in §3 / §6 / §7 to be
ticked, the §4 / §5 counts to satisfy the eligibility and completeness
rules stated there, and the prototype contract in
[`milestone_1_readiness.md`](milestone_1_readiness.md) to still hold.

- [ ] **GO.** Every prerequisite above holds; the rendered
      `new_report.docx` is safe to share with the sponsor for this
      single trial.
- [ ] **NO-GO.** Justification — aggregate language only, no raw values
      or report content: `<one sentence>`.

---

This template is **documentation-only**. It runs no command, writes no
artifact, and is never auto-populated by the pipeline. The reviewer
fills it in by hand after walking the workflow in
[`real_file_pilot.md`](real_file_pilot.md). Distribute a filled-in copy
only to the sponsor pair and only after re-reading the
"What this template MUST NOT contain" section above.
