# Real-file pilot readiness packet

This packet describes the **manual** workflow for running the existing
six-stage pipeline against a colleague-held **real** Excel + Word pair,
without committing any of that data into this repository.

It changes none of the pipeline's matching, confirmation, rendering, or
validation logic — every command behaves exactly as documented in
`README.md`. The only addition is a one-time **privacy advisory** that
appears on stderr if a path you pass on the CLI resolves inside this
repo's `samples/` folder; see [§5](#5-privacy-preflight-advisory) below.

> ⚠️ This is still a **prototype**. Do not promote a rendered Word
> report to a customer or stakeholder until a reviewer has walked every
> artifact listed in [§3](#3-artifacts-a-reviewer-must-inspect).

> Before pointing the pipeline at real data, walk the operator
> dry-run checklist in
> [`synthetic_dry_run_checklist.md`](synthetic_dry_run_checklist.md)
> against the synthetic fixture. The rehearsal exercises the same
> gates this workflow relies on, end-to-end, with no real data and
> no committed artifacts.

> For a reviewer-facing snapshot of which gates and tests exist
> today, the known v1 limitations, and the eight things a reviewer
> must verify before sponsoring a real-file pilot, see
> [`milestone_1_readiness.md`](milestone_1_readiness.md).

---

## 1. Where real files must live

Real company Excel workbooks and Word reports must live **outside this
repository**, full stop. The `samples/` folder is reserved for
synthetic fixtures (see CLAUDE.md "Repo Conventions" and "Privacy
Rules"), and the `.gitignore` masking of `samples/synthetic/` is a
convenience — not a guarantee.

Recommended layout (anywhere outside the repo):

```
~/pilot_data/
  <pilot_label>/
    inputs/
      historical.xlsx         # historical period workbook
      finished_report.docx    # the matching finished Word report
      new_period.xlsx         # the workbook for the period to render
    output/                   # all generated artifacts land here
```

Do **not**:

- Place real files under `<repo>/samples/`, `<repo>/output/`,
  `<repo>/tests/`, or any other path inside the repo tree.
- Copy/paste sensitive table content into commit messages, PR
  descriptions, or generated synthetic fixtures.
- Commit a rendered `.docx`, `.xlsx`, or `confidence_report.md` that
  contains real metric values, even temporarily.

If you accidentally placed a real file inside the repo, move it out
**before** running any `git add`. Use `git status` to verify the working
tree is clean of those paths afterwards.

---

## 2. Command sequence

The exact sequence is the same as the synthetic quickstart in
`README.md`, with one change: every `--excel`, `--word`, and `--out`
path points at your external pilot directory.

For a compact per-command reference (purpose, required inputs,
outputs written, exit codes, read-only vs. writes, and which
commands' console output is safe to paste into a chat or ticket),
see [`command_reference.md`](command_reference.md). The exit-code
table in [§2c](#2c-expected-exit-codes) below is the canonical
per-stage view; the reference doc collects the same codes into a
single map.

> **Before the first `learn` invocation**, run the read-only
> [`pilot-preflight`](command_reference.md#pilot-preflight) command
> against the same four paths the sequence below uses. It verifies
> the input files exist with the right suffix, `--out` is a
> directory (or can be created), and — most importantly — that no
> pilot path resolves inside this repo tree. The check never opens
> any document and prints only flag labels and basenames, so its
> output is safe to paste into a ticket. Exit codes: `0` ready,
> `2` per-path issue (missing/wrong-suffix/not-a-dir), `12` privacy
> refusal (a pilot path lands inside the repo tree).
>
> ```bash
> python -m src.main pilot-preflight \
>     --historical-excel "$PILOT/inputs/historical.xlsx" \
>     --historical-word  "$PILOT/inputs/finished_report.docx" \
>     --new-excel        "$PILOT/inputs/new_period.xlsx" \
>     --out              "$PILOT/output"
> ```

### 2a. Environment setup (cross-platform)

The pipeline is pure Python plus `openpyxl` / `python-docx` / `PyYAML`
— it does **not** require Microsoft Word, a venv, or any specific
package manager. Any Python 3.10+ environment that has the deps from
`requirements.txt` installed will work. Pick whichever path matches
your platform and toolchain. The exit-code contract in
[§2c](#2c-expected-exit-codes) is identical no matter which path you use.

**macOS / Linux (bash or zsh, virtualenv):**

```bash
# Either create a project venv ...
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
# ... or install into your existing environment of choice (conda, uv,
# pipx, system python). The pipeline does not care which.
```

**Windows (PowerShell, virtualenv):**

```powershell
# Either create a project venv ...
python -m venv .venv
.\.venv\Scripts\Activate.ps1
# (if blocked: `Set-ExecutionPolicy -Scope Process RemoteSigned` then re-run)
pip install -r requirements.txt
# ... or install into conda, uv, or system python.
```

**Windows (cmd.exe, virtualenv):**

```bat
python -m venv .venv
.\.venv\Scripts\activate.bat
pip install -r requirements.txt
```

If you already have a working Python with the deps installed, you can
skip the activation step entirely and invoke `python -m src.main …`
directly — the CLI does not depend on a virtualenv being active.

### 2b. Pilot variables and command sequence

Set a `PILOT` (bash/zsh) or `$Pilot` (PowerShell) variable that points
at your external pilot directory so the command sequence reads the
same on every shell.

**macOS / Linux (bash or zsh):**

```bash
# Pick any path OUTSIDE the repo — ~/pilot_data is just a convention.
PILOT="$HOME/pilot_data/<pilot_label>"
```

**Windows (PowerShell):**

```powershell
# Pick any path OUTSIDE the repo.
$Pilot = "$HOME\pilot_data\<pilot_label>"
```

**Windows (cmd.exe):**

```bat
set PILOT=%USERPROFILE%\pilot_data\<pilot_label>
```

The command sequence itself is identical across shells; only the
variable-reference syntax differs (`"$PILOT/..."` on bash/zsh,
`"$Pilot\..."` on PowerShell, `"%PILOT%\..."` on cmd.exe). The bash
form below is canonical; translate the variable interpolation as
needed for your shell.

```bash
# (1) Strict learn — the trust gate before any real-file pilot.
#     This MUST exit 0 before you confirm or render.
python -m src.main learn \
    --excel "$PILOT/inputs/historical.xlsx" \
    --word  "$PILOT/inputs/finished_report.docx" \
    --out   "$PILOT/output" \
    --strict

# (2) Cross-check the four learn-mode artifacts agree with each other.
python -m src.main validate-artifacts --out "$PILOT/output"

# (3) Reviewer fills in `Reviewer Decision` / `Reviewer Notes` (and
#     optional `Confirmed Sheet` + `Confirmed Cell` overrides) in
#     `$PILOT/output/mapping_review.xlsx`. See §3.
python -m src.main confirm-mapping \
    --auto   "$PILOT/output/auto_mapping.yml" \
    --review "$PILOT/output/mapping_review.xlsx" \
    --out    "$PILOT/output/confirmed_mapping.yml"

# (4) Resolve confirmed mapping against the new-period workbook.
python -m src.main run-preview \
    --excel     "$PILOT/inputs/new_period.xlsx" \
    --confirmed "$PILOT/output/confirmed_mapping.yml" \
    --out       "$PILOT/output/run_preview"

# (5) Deterministically render the new Word report.
python -m src.main render-docx \
    --template       "$PILOT/output/converted_template.docx" \
    --run-validation "$PILOT/output/run_preview/run_validation.xlsx" \
    --out            "$PILOT/output/new_report.docx"

# (6) Final cross-check of the three rendered-output artifacts.
python -m src.main validate-render \
    --docx           "$PILOT/output/new_report.docx" \
    --render-log     "$PILOT/output/render_log.yml" \
    --run-validation "$PILOT/output/run_preview/run_validation.xlsx"
```

On Windows PowerShell, the same call looks like:

```powershell
# (1) Strict learn — same arguments, only the variable + path separator
#     differ. Backticks ` continue the line in PowerShell.
python -m src.main learn `
    --excel  "$Pilot\inputs\historical.xlsx" `
    --word   "$Pilot\inputs\finished_report.docx" `
    --out    "$Pilot\output" `
    --strict
# ... and so on for steps (2)–(6). pathlib normalizes the slashes on
# Python's side, so forward slashes also work if you prefer them.
```

### 2c. Expected exit codes

The codes below are the contract; automation can branch on them. They
are unchanged from `README.md` — repeated here so a pilot operator does
not have to context-switch between docs.

| Stage | Success | Failure modes (non-zero) |
| --- | --- | --- |
| `learn --strict` | `0` | `2` missing `--excel`/`--word`; `3` at least one eligible Word number is `UNRESOLVED` or `LOW` |
| `validate-artifacts` | `0` | `2` missing `--out`; `4` cross-artifact disagreement |
| `confirm-mapping` | `0` | `2` missing/out-of-sync inputs; `5` at least one row landed in `review_required` (unless `--allow-incomplete`) |
| `run-preview` | `0` | `2` missing inputs; `6` `confirmed_mapping.yml` cannot prove completeness; `7` a confirmed source/transform broke on the new workbook |
| `render-docx` | `0` | `2` missing inputs; `8` fatal input error (no docx written); `9` per-row gate failure (no docx written) |
| `validate-render` | `0` | `2` missing inputs; `10` at least one consistency check failed |
| `pilot-summary` (optional) | `0` | `2` missing `--out` (or not a directory); `11` `auto_mapping.yml` not present — nothing to summarize |
| `pilot-preflight` (optional, pre-pilot) | `0` | `2` per-path issue (missing input, wrong suffix, `--out` exists but is not a directory, `--out` cannot be created because its nearest existing ancestor is not a directory); `12` a pilot path resolves inside the repo tree |

A **non-zero exit at any stage halts the pilot**. Do not pass the
generated `new_report.docx` to a stakeholder until every stage above
returns `0`. The `--allow-incomplete` flag on `confirm-mapping` is an
exploratory escape hatch and must not be used in a real-file pilot.

### 2d. Optional: read-only redacted progress check

At any point in the sequence — after `learn`, mid-review,
post-`render-docx`, etc. — you can ask for a one-shot redacted summary
of the pilot output directory:

```bash
python -m src.main pilot-summary --out "$PILOT/output"
```

The summary reports artifact presence, file sizes, per-stage aggregate
counts (from each artifact's own `summary` block / `Status` column),
and a single next-action hint per stage. It is **safe to paste into a
chat, ticket, or commit**: by contract it never prints raw Word
tokens, generated values, raw Excel values, source sheet/cell content,
reviewer notes, individual `word_id` values, or file paths beyond
basenames. The `--out` directory is reduced to its basename in the
header for the same reason.

Exit codes are deliberately narrow:

- `0` — summary printed.
- `2` — `--out` does not exist, or is not a directory.
- `11` — `--out` exists but `auto_mapping.yml` is absent, so there is
  nothing meaningful to summarize. Run `learn` first.

`pilot-summary` is **read-only**: it never re-runs the pipeline, never
mutates an artifact, never calls out to an LLM, GUI, network, or
Microsoft Word automation. It only inspects what is already on disk.

---

## 3. Artifacts a reviewer must inspect

Before trusting the rendered Word report, walk these artifacts in
order. Each one answers a question the next stage assumes is already
true.

1. **`confidence_report.md`** — read the **top** sections first. The
   doc deliberately leads with UNRESOLVED, LOW, ambiguous, and
   EXCLUDED-by-policy buckets so the items needing human eyes are
   not buried under HIGH wins.
2. **`mapping_review.xlsx`** — for every row, sanity-check that the
   `Word ID`, `Word Location`, raw token, candidate cell, and overlap
   tokens line up with what the historical report actually said.
   The first four columns (`Word ID`, `Review Status`,
   `Placeholder Status`, `Placeholder`) tell the audit story without
   drilling into candidate detail.
3. **`auto_mapping.yml`** — confirms machine truth: every `word_id`,
   its `review_status`, its `placeholder_status`, and any
   `alternatives` a reviewer might pick over the recommended source.
4. **`converted_template.docx`** — open in any docx viewer that does
   not auto-execute macros (e.g., LibreOffice with macro execution
   disabled, or a docx text dump). Confirm:
   - HIGH/MEDIUM tokens are replaced with `{{ word_NNNN }}`
     placeholders.
   - LOW, UNRESOLVED, and EXCLUDED values are **still visible** as
     their original text — this is the audit guarantee CLAUDE.md
     mandates.
5. **`confirmed_mapping.yml`** — after `confirm-mapping`, this file
   must show `summary.complete: true` and `summary.allow_incomplete:
   false` (or the absence of `allow_incomplete`). The
   `review_required` bucket must be **empty** before `run-preview` is
   safe.
6. **`run_validation.xlsx`** — every row's `Status` must be `ok`.
   For each row, sanity-check that:
   - `Source Sheet` + `Source Cell` resolve to the cell you expect in
     the new workbook.
   - `Raw Excel Value`, `Generated Value`, and
     `Transform Interpretation` agree with what the historical report
     said at that position.
7. **`new_report.docx`** + **`render_log.yml`** — open the rendered
   docx and skim every number. For each `word_id` in `render_log.yml`
   confirm:
   - `placeholder_occurrences >= 1`.
   - `display_text` actually appears in the rendered docx (the
     `validate-render` gate enforces this, but eyes-on is the
     last line of defense against a tampered audit).

---

## 4. Handling unresolved, LOW, EXCLUDED, non-renderable, and tampered rows

Every visible Word number must land in exactly one of the buckets
below — that is the no-silent-omission contract from CLAUDE.md. Each
bucket has a single appropriate disposition; do **not** invent new
behaviors at pilot time.

| Row type | Where it surfaces | Disposition in a pilot |
| --- | --- | --- |
| `UNRESOLVED` | `auto_mapping.yml` (`review_status: needs_source`), `confidence_report.md`, `mapping_review.xlsx`, and `run_preview`'s pre-confirmation pass | The matcher could not find an Excel source. **Do not** confirm. Either (a) add the missing source value to the historical workbook and re-run `learn`, (b) update the matcher policy, or (c) abandon the row — i.e., mark `Reviewer Decision: reject` so it lands in `review_required` and the pilot stops at `confirm-mapping` exit `5`. |
| `LOW` | Same as UNRESOLVED, plus `review_status: needs_review` | A value match exists but context overlap is weak. **Never** confirm with a blank decision. Either reject (`Reviewer Decision: reject`) or fill in both `Confirmed Sheet` AND `Confirmed Cell` so the source is explicit. |
| `EXCLUDED` (date/period markers, e.g. `5月`, `2026年`, `第20周`) | `auto_mapping.yml` (`audit_only_excluded`), `confidence_report.md` "EXCLUDED by explicit policy" section | These are audit-only by design and **never** become confirmed mappings. Skim the EXCLUDED section to confirm the policy applied correctly; if a real metric was mis-classified as EXCLUDED, that is a matcher bug — file it, do not edit the YAML by hand. |
| Non-renderable HIGH/MEDIUM (template builder skipped the placeholder due to offset/raw drift) | `auto_mapping.yml` `placeholder_status` is **not** `applied`; `confirm-mapping` puts the row into `review_required` with reason `non_renderable_template_skip:…` | Confirming would silently omit the metric at render time because the converted template has no placeholder to substitute into. **Do not** force-confirm. Either fix the template builder gap and re-run `learn`, or accept that this pilot cannot render that row. |
| Run-preview per-row failures (`missing_sheet`, `missing_cell`, `non_numeric_cell`, `missing_confirmed_source`, `transform_unknown`, `missing_transform`) | `run_validation.xlsx` `Status` and `Detail` columns; CLI exit `7` | The confirmed `(sheet, cell)` or transform broke on the new workbook. **Do not** edit `run_validation.xlsx` by hand. Re-check the new workbook's structure against the historical one and re-run `run-preview`; if the schema genuinely changed, re-run `learn` + `confirm-mapping` against a new historical pair. |
| Render-docx per-row failures (non-ok validation row, missing/duplicated `Generated Value`, template references an unknown `word_id`, validation has an unused `word_id`, raw token does not parse) | CLI exit `8` or `9`; **no** docx written | Fix the upstream artifact (`run_validation.xlsx` from `run-preview`, `converted_template.docx` from `learn`) and re-run. Never paste a value into the docx by hand. |
| `validate-render` consistency failure (`leftover_placeholder`, `missing_render_log_row`, `extra_render_log_row`, `audit_value_drift`, `docx_rendered_text_missing`, `render_log_status_not_ok`, `render_log_missing_field`, `zero_placeholder_occurrences`, `run_validation_status_not_ok`, …) | CLI exit `10` | Treat as **tampered or drifted** until proven otherwise. The artifacts disagree, which means at least one of them lies. **Do not** edit any of the three artifacts by hand to "make the validator happy". Re-run `render-docx` from `run_validation.xlsx` and `converted_template.docx`; if the failure persists, walk back to whichever earlier stage produced the offending artifact and re-run from there. |

The general rule: **the only legal way to change a generated artifact
is to re-run the stage that produced it from earlier audited inputs.**
Hand-editing an artifact bypasses every gate that protects against
silent omission or invention.

---

## 5. Privacy preflight advisory

Every pilot-sequence subcommand — `learn`, `validate-artifacts`,
`confirm-mapping`, `run-preview`, `render-docx`, and `validate-render`
— checks **every path argument you passed** on the command line and
emits a one-time advisory to stderr when any of them resolves inside
this repository's `samples/` folder. The set of checked flags per
command:

| Command | Flags checked |
| --- | --- |
| `learn` | `--excel`, `--word`, `--out` |
| `validate-artifacts` | `--out` |
| `confirm-mapping` | `--auto`, `--review`, `--out` |
| `run-preview` | `--excel`, `--confirmed`, `--out` |
| `render-docx` | `--template`, `--run-validation`, `--out` |
| `validate-render` | `--docx`, `--render-log`, `--run-validation` |

`generate-synthetic` is the one exception: it is intentionally
synthetic-only and its default `--out` (`samples/synthetic`) would
trigger the advisory on every invocation, so it is not instrumented.

A triggered advisory looks like:

```
!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
PRIVACY PREFLIGHT ADVISORY
  - --excel /…/samples/synthetic/historical.xlsx is inside the repo's
    samples/ folder, which is reserved for synthetic fixtures.
Real company Excel/Word files must live OUTSIDE this repo
(e.g. ~/pilot_data/) and must NEVER be committed.
See docs/real_file_pilot.md for the supported pilot workflow.
!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
```

The advisory is **informational only**. It does **not** change
matching, confirmation, rendering, or validation logic, and it does
**not** block the command. The exit code is exactly what it would have
been without the advisory.

If you see it, the right response is:

1. **Stop.** Do not commit anything until you have moved the real
   files out of the repo.
2. `mv` the offending file to an external pilot directory (see §1).
3. Re-invoke the command with the external path.
4. `git status` to confirm nothing leaked into the working tree.

If you are deliberately exercising the pipeline against the synthetic
sample (e.g., to reproduce a bug), the advisory is harmless — synthetic
files are safe to commit if they are genuinely synthetic, but
double-check before doing so.

---

## 6. Sign-off checklist

Before sharing a rendered `new_report.docx` from a pilot, the
operator-plus-reviewer pair should be able to answer **yes** to every
one of these:

- [ ] `learn --strict` exited `0`.
- [ ] `validate-artifacts` exited `0`.
- [ ] `confirm-mapping` exited `0` without `--allow-incomplete`, and
      `confirmed_mapping.yml` shows `summary.complete: true`.
- [ ] `run-preview` exited `0` and every `run_validation.xlsx` row
      has `Status = ok`.
- [ ] `render-docx` exited `0` and `render_log.yml` lists every
      confirmed `word_id` with `placeholder_occurrences >= 1`.
- [ ] `validate-render` exited `0`.
- [ ] A human has visually skimmed every number in `new_report.docx`
      and confirmed it matches the corresponding `display_text` and
      `Generated Value` in `render_log.yml`.
- [ ] No file under `~/pilot_data/` (or wherever real files live) is
      staged, tracked, or referenced in any commit on any branch.
- [ ] `git status` in the repo shows only the expected source-code
      changes — no real `.xlsx`, `.docx`, `.yml`, or `.md` artifacts.

If any box cannot be ticked, the pilot is not ready to ship.
