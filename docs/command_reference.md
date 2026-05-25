# CLI command reference

A compact per-command reference for the eight `python -m src.main`
subcommands. Use it to look up purpose, required inputs, outputs,
and exit-code meanings without re-reading the full `README.md` or
`docs/real_file_pilot.md`.

This file is an **index**, not a contract — it mirrors what those
two longer documents already pin down and never overrides them.
Behavioural details (why a row landed in `review_required`, how
transforms work, what `validate-render` actually checks, the full
sign-off checklist for a pilot, ...) live in:

- [`README.md`](../README.md) — synthetic quickstart and per-stage
  contracts.
- [`docs/real_file_pilot.md`](real_file_pilot.md) — manual real-file
  pilot workflow, artifact-inspection checklist, and disposition
  rules for unresolved / LOW / EXCLUDED / tampered rows.

## Conventions used below

- **Read-only** commands inspect files already on disk and write
  nothing. **Writes artifacts** commands create or overwrite files
  under `--out`.
- **Paste-safe?** answers whether the command's stdout/stderr is
  safe to copy into a chat, ticket, or commit message **without
  further redaction** when run against **real** data. On synthetic
  samples every command's output is safe because no real content
  exists; the column refers to real-file pilots. `pilot-summary` is
  the only command contractually redacted for this purpose.
- Exit codes `0` and `2` are shared success / input-error codes.
  Codes `3` through `11` are non-overlapping gate-specific failures
  so automation can branch on a single integer; see the
  [exit-code map](#exit-code-map-at-a-glance) at the bottom of this
  file.

---

## `generate-synthetic`

- **Purpose** — Emit a paired fake Excel + Word sample so the rest
  of the pipeline can be exercised without any real company data.
- **Required inputs** — none. Optional `--out` (default
  `samples/synthetic`) selects the output directory.
- **Outputs written** — `historical.xlsx` and `finished_report.docx`
  under `--out`.
- **Exit codes** — `0` on success.
- **Read-only?** — no; writes synthetic artifacts under `--out`.
- **Paste-safe?** — yes. The output is two filesystem paths and a
  header line; the artifacts are synthetic by construction.

Privacy preflight: **not** instrumented. The default `--out`
intentionally points inside `samples/` (synthetic fixtures only).

---

## `learn`

- **Purpose** — Profile an Excel + Word pair, match every visible
  Word number, and emit the four reviewable learn-mode artifacts
  plus a console coverage summary. This is the trust gate for the
  rest of the pipeline.
- **Required inputs** — `--excel <historical.xlsx>`,
  `--word <finished_report.docx>`, `--out <directory>`. Optional
  `--strict` to fail on any eligible UNRESOLVED/LOW row.
- **Outputs written** — `mapping_review.xlsx`, `auto_mapping.yml`,
  `converted_template.docx`, `confidence_report.md` under `--out`.
- **Exit codes** — `0` success; `2` missing `--excel`/`--word`;
  `3` `--strict` and at least one eligible Word number is
  UNRESOLVED or LOW (artifacts still written so a reviewer can
  inspect what failed).
- **Read-only?** — no; writes the four learn-mode artifacts.
- **Paste-safe?** — no on real data. The console summary lists
  sample Word tokens from each problem bucket (UNRESOLVED, LOW,
  EXCLUDED). Redact before sharing externally. On synthetic samples
  the output is safe.

Privacy preflight: emits a one-time advisory on stderr if any of
`--excel`, `--word`, `--out` resolves inside the repo's `samples/`
folder. The advisory never blocks the run or changes the exit code.

---

## `validate-artifacts`

- **Purpose** — Cross-check that the four learn-mode artifacts
  written by `learn` tell the same story about every Word number.
  Read-only contract before `confirm-mapping`.
- **Required inputs** — `--out <directory written by learn>`.
- **Outputs written** — none.
- **Exit codes** — `0` success; `2` missing `--out`; `4` cross-
  artifact disagreement (per-issue list printed to stdout).
- **Read-only?** — yes.
- **Paste-safe?** — no on real data. Failure output names
  `word_id` values and concrete drift descriptions. On synthetic
  samples the output is safe.

Privacy preflight: emits an advisory if `--out` resolves inside
`samples/`.

---

## `confirm-mapping`

- **Purpose** — Promote reviewer-approved rows from
  `mapping_review.xlsx` + `auto_mapping.yml` into
  `confirmed_mapping.yml`. Rows that are not safely confirmable
  (blank decision, reject, LOW/UNRESOLVED, invalid override,
  non-renderable template skip) stay visible as `review_required`
  and the run fails the gate.
- **Required inputs** — `--auto <auto_mapping.yml>`,
  `--review <mapping_review.xlsx>`, `--out <confirmed_mapping.yml>`.
  Optional `--allow-incomplete` to write the YAML and exit `0` even
  when `review_required` is non-empty (exploratory escape hatch
  only — never use in a real-file pilot).
- **Outputs written** — `confirmed_mapping.yml` at `--out`. Written
  on `0` and on `5`; not written on `2`.
- **Exit codes** — `0` success; `2` missing or out-of-sync inputs
  (no YAML written); `5` at least one row in `review_required`
  (YAML still written so the reviewer can see exactly what is
  blocking).
- **Read-only?** — no; writes `confirmed_mapping.yml`.
- **Paste-safe?** — no on real data. Output names `word_id` values
  and per-row reasons (e.g. `non_renderable_template_skip:…`,
  `unresolved_no_candidate`). On synthetic samples the output is
  safe.

Privacy preflight: emits an advisory if any of `--auto`,
`--review`, `--out` resolves inside `samples/`.

---

## `run-preview`

- **Purpose** — Resolve confirmed mappings against a NEW period's
  Excel workbook and write the per-row validation table that
  `render-docx` consumes. **Does not render a Word document.**
- **Required inputs** — `--excel <new_period.xlsx>`,
  `--confirmed <confirmed_mapping.yml>`, `--out <directory>`.
- **Outputs written** — `run_validation.xlsx` under `--out`.
  Written on `0` and on `7`; not written on `6`.
- **Exit codes** — `0` success; `2` missing inputs; `6`
  `confirmed_mapping.yml` cannot prove completeness (no artifact
  written, input is unusable); `7` a confirmed source or transform
  broke on the new workbook (artifact still written so each broken
  row is visible to the reviewer).
- **Read-only?** — no; writes `run_validation.xlsx`.
- **Paste-safe?** — no on real data. Output may name source sheets,
  cells, raw values, and per-row failure reasons. On synthetic
  samples the output is safe.

Privacy preflight: emits an advisory if any of `--excel`,
`--confirmed`, `--out` resolves inside `samples/`.

---

## `render-docx`

- **Purpose** — Deterministically substitute `{{ word_NNNN }}`
  placeholders in `converted_template.docx` using
  `run_validation.xlsx`, and write the rendered Word report plus an
  audit log. No LLM, no GUI, no network, no Microsoft Word
  automation.
- **Required inputs** — `--template <converted_template.docx>`,
  `--run-validation <run_validation.xlsx>`,
  `--out <new_report.docx>`.
- **Outputs written** — `new_report.docx` at `--out` and
  `render_log.yml` alongside in the same directory. Written **only**
  on success (`0`); on `8`/`9` neither file is written — CLAUDE.md
  forbids a partial docx that silently omits or mis-renders a
  confirmed metric.
- **Exit codes** — `0` success; `2` missing inputs; `8` fatal input
  error (no docx written); `9` per-row gate failure such as non-ok
  validation row, missing or duplicated `Generated Value`, template
  references an unknown `word_id`, validation has an unused
  `word_id`, or raw-token format inference failed (no docx written).
- **Read-only?** — no; writes the rendered docx and `render_log.yml`.
- **Paste-safe?** — no on real data. The console summary names
  `word_id` values and offending rows; the rendered docx itself is
  the actual report output and is never paste-safe. On synthetic
  samples the console output is safe.

Privacy preflight: emits an advisory if any of `--template`,
`--run-validation`, `--out` resolves inside `samples/`.

---

## `validate-render`

- **Purpose** — Read-only cross-check of the three rendered-output
  artifacts (`new_report.docx`, `render_log.yml`,
  `run_validation.xlsx`). Refuses to bless a docx whose audit story
  has drifted, including a hand-edited number that would pass a
  naive file-existence check.
- **Required inputs** — `--docx <new_report.docx>`,
  `--render-log <render_log.yml>`,
  `--run-validation <run_validation.xlsx>`.
- **Outputs written** — none.
- **Exit codes** — `0` success; `2` missing inputs; `10` at least
  one consistency check failed (per-issue list on stderr with short
  codes like `leftover_placeholder`, `audit_value_drift`,
  `docx_rendered_text_missing`, `render_log_status_not_ok`,
  `zero_placeholder_occurrences`).
- **Read-only?** — yes; never re-renders, never mutates an
  artifact.
- **Paste-safe?** — no on real data. Failure output names
  `word_id` values and the offending issue codes. On synthetic
  samples the output is safe.

Privacy preflight: emits an advisory if any of `--docx`,
`--render-log`, `--run-validation` resolves inside `samples/`.

---

## `pilot-summary`

- **Purpose** — Read-only **redacted** summary of a pilot output
  directory. Reports artifact presence, file sizes, per-stage
  aggregate counts (from each artifact's own `summary` block /
  `Status` column), and a single next-action hint per stage.
- **Required inputs** — `--out <pilot output directory>`.
- **Outputs written** — none.
- **Exit codes** — `0` success; `2` `--out` is missing or not a
  directory; `11` `--out` exists but `auto_mapping.yml` is absent —
  nothing meaningful to summarize (run `learn` first).
- **Read-only?** — yes; never re-runs the pipeline, never mutates
  an artifact, never calls out to an LLM, GUI, network, or
  Microsoft Word automation.
- **Paste-safe?** — **yes**, by contract. Never prints raw Word
  tokens, generated values, raw Excel values, source sheet/cell
  content, reviewer notes, individual `word_id` values, or file
  paths beyond basenames. The `--out` directory is reduced to its
  basename in the header for the same reason.

Privacy preflight: emits an advisory if `--out` resolves inside
`samples/`.

---

## Exit-code map at a glance

The gate-specific non-zero exit codes are intentionally
non-overlapping so automation can branch on a single integer. Codes
`0` and `2` are shared across commands by design.

| Code | Meaning | Emitted by |
| --- | --- | --- |
| `0` | success | every command |
| `2` | missing or invalid input path / directory | every command |
| `3` | `learn --strict` saw eligible UNRESOLVED/LOW | `learn` |
| `4` | learn-mode artifact cross-check failed | `validate-artifacts` |
| `5` | confirmed mapping has `review_required` rows | `confirm-mapping` |
| `6` | `confirmed_mapping.yml` cannot prove completeness | `run-preview` |
| `7` | per-row failure on the new workbook | `run-preview` |
| `8` | render-docx fatal input error (no docx written) | `render-docx` |
| `9` | render-docx per-row gate failure (no docx written) | `render-docx` |
| `10` | rendered-output cross-check failed | `validate-render` |
| `11` | nothing to summarize (`auto_mapping.yml` absent) | `pilot-summary` |

Codes `1` and anything beyond `11` are unused today — treat them
as unexpected errors and inspect stderr.
