# Milestone 1 readiness — synthetic-pipeline prototype

This document is a **reviewer-facing snapshot** of the prototype as it
stands today. It exists so a reviewer can decide, without re-reading
`README.md` and `docs/real_file_pilot.md` end-to-end, whether the
current code is ready to be **trialed against a real Excel + Word
pair** under the manual workflow in `docs/real_file_pilot.md`.

> ⚠️ **Status: prototype.** Nothing in this document claims the tool
> is ready for production, GA, customer-facing use, or unattended
> automation. The synthetic pipeline holds together end-to-end; that
> is the entire milestone. Real-file behavior remains unproven until
> a reviewer walks the artifacts on a real pair.

---

## 1. What is implemented and tested end-to-end

The six-stage deterministic pipeline runs on **synthetic** Excel + Word
fixtures from `generate-synthetic` and is exercised by the test suite:

1. `learn` — profile a paired Excel + Word, match every visible Word
   number, write `mapping_review.xlsx`, `auto_mapping.yml`,
   `converted_template.docx`, and `confidence_report.md`.
2. `validate-artifacts` — cross-check that the four learn-mode
   artifacts agree by stable `word_id`.
3. `confirm-mapping` — promote reviewer-approved rows into
   `confirmed_mapping.yml`; the rest land in `review_required`
   (visible, not silently dropped).
4. `run-preview` — resolve `confirmed_mapping.yml` against a new
   period's Excel and write `run_validation.xlsx`.
5. `render-docx` — deterministically substitute `{{ word_NNNN }}`
   placeholders in `converted_template.docx` using
   `run_validation.xlsx`; write `new_report.docx` and `render_log.yml`.
6. `validate-render` — cross-check the three rendered-output artifacts
   (`new_report.docx`, `render_log.yml`, `run_validation.xlsx`) agree.

Two operator-facing helpers also exist:

- `pilot-summary` — read-only, contractually redacted, paste-safe
  status overview of a pilot output directory.
- `pilot-preflight` — read-only path/metadata gate that refuses
  inside-repo pilot paths before the first `learn` run.

`tests/test_acceptance_smoke.py` walks both the Python-API surface and
the CLI surface through stages 1–6 on an inline synthetic pair with
zero eligible UNRESOLVED/LOW, asserts every artifact exists at every
stage, asserts no `{{ word_NNNN }}` placeholder leaks into the
rendered docx, and includes a negative path that tampers
`run_validation.xlsx` so `validate-render` exits `10`. Per-stage
behavior is additionally pinned by `test_learn_smoke.py`,
`test_validate_artifacts.py`, `test_confirm_mapping.py`,
`test_run_preview.py`, `test_renderer.py`, `test_validate_render.py`,
`test_pilot_summary*.py`, and `test_pilot_preflight.py`.

---

## 2. Commands that prove the synthetic pipeline works

Run the following from the repo root. Each command's expected exit
code is listed; any non-zero exit is a failure of this milestone.

```bash
# (0) Install dependencies.
pip install -r requirements.txt

# (1) Generate the paired synthetic Excel + Word fixture.
python -m src.main generate-synthetic --out samples/synthetic     # exit 0

# (2) Learn mode (exploratory — the synthetic corpus has deliberate
#     UNRESOLVED rows, so --strict would correctly exit 3 here).
python -m src.main learn \
    --excel samples/synthetic/historical.xlsx \
    --word  samples/synthetic/finished_report.docx \
    --out   output                                                # exit 0

# (3) Cross-check the four learn-mode artifacts.
python -m src.main validate-artifacts --out output                # exit 0

# (4) Whole-pipeline acceptance smoke (inline clean synthetic pair;
#     stages 1–6 plus the negative validate-render tamper case).
python -m pytest tests/test_acceptance_smoke.py -v                # exit 0

# (5) Full test suite.
python -m pytest                                                  # exit 0
```

Stages 3–6 of the pipeline (`confirm-mapping`, `run-preview`,
`render-docx`, `validate-render`) are exercised end-to-end by the
acceptance smoke in step (4); the synthetic corpus from step (2)
deliberately includes UNRESOLVED rows so `confirm-mapping` cannot
promote it without `--allow-incomplete`, which is by design.

---

## 3. Current safety gates

Every gate below is implemented and pinned by a test. Each fails
**loudly** (non-zero exit, named failure code, stderr explanation) —
none of them silently drop or invent a Word number.

| Gate | Where | What it refuses | Non-zero exit |
| --- | --- | --- | --- |
| `learn --strict` | `learn` CLI | Any eligible Word number is `UNRESOLVED` or `LOW` | `3` |
| `validate-artifacts` | `validate-artifacts` CLI | The four learn-mode artifacts disagree on `word_id`, location, raw token, confidence, or placeholder status | `4` |
| `confirm-mapping` | `confirm-mapping` CLI | At least one row is in `review_required` (blank decision, reject, LOW, UNRESOLVED, invalid/incomplete override, non-renderable template skip) and `--allow-incomplete` was **not** passed | `5` |
| `run-preview` completeness | `run-preview` CLI | `confirmed_mapping.yml` cannot prove it is complete (missing `summary`, `complete` not literal `true`, `allow_incomplete: true`, non-empty `review_required`, zero confirmed rows) — **no artifact written** | `6` |
| `run-preview` per-row | `run-preview` CLI | A confirmed `(sheet, cell)` is missing/empty/non-numeric on the new workbook, or its transform interpretation is absent/unknown to the v1 table — artifact still written so each broken row is visible | `7` |
| `render-docx` input | `render-docx` CLI | Fatal input error such as a `{{ word_NNNN }}` in the template with no validation row — **no docx or log written** | `8` |
| `render-docx` per-row | `render-docx` CLI | A non-ok validation row, missing/duplicated `Generated Value`, unused `word_id`, or raw-token format inference failure — **no docx or log written** | `9` |
| `validate-render` | `validate-render` CLI | The three rendered-output artifacts disagree: leftover placeholder, missing/extra log row, value drift, missing display text in the docx, zero placeholder occurrences, non-ok status | `10` |
| `pilot-summary` precondition | `pilot-summary` CLI | `--out` exists but `auto_mapping.yml` is absent — nothing meaningful to summarize | `11` |
| `pilot-preflight` privacy | `pilot-preflight` CLI | Any pilot path resolves inside the repo tree (strictly stronger than the `samples/` advisory) | `12` |
| Privacy preflight advisory | `learn`, `validate-artifacts`, `confirm-mapping`, `run-preview`, `render-docx`, `validate-render` | A CLI path argument resolves inside `samples/` — informational stderr advisory only, exit code unchanged | (no code; stderr banner) |
| Repo privacy boundary | `tests/test_repo_privacy_boundary.py` | Any `.xlsx`/`.xls`/`.docx`/`.doc` is tracked by git; any generated artifact basename is tracked; `.gitignore` no longer masks `output/` or `samples/synthetic/`; operator docs no longer point real files outside the repo or reference `pilot-preflight` | test failure |

The complete exit-code map (with the gate-specific code for every
command) lives in `docs/command_reference.md` and is parity-tested
against `src/main.py` by `tests/test_command_reference_docs.py`.

---

## 4. Known limitations — not production-ready yet

The prototype is deliberately narrow. Each item below is a real gap a
reviewer must accept before sponsoring a real-file pilot.

- **Synthetic-only test coverage.** The test suite never reads a real
  company file. No real-file pair has been independently audited by
  the team yet — that is exactly what the manual workflow in
  `docs/real_file_pilot.md` is for.
- **Surface scope.** `learn`'s template builder, `render-docx`, and
  `validate-render` all walk paragraphs and top-level table cells
  only. Headers, footers, footnotes, text boxes, chart data, and
  nested tables are out of scope; a placeholder leaked into one of
  those surfaces would not be caught by v1.
- **Run-level styling.** `render-docx` collapses inline runs to a
  single run when it substitutes a placeholder. Bold/italic/color
  spans inside a number's paragraph are not preserved by v1.
- **Fixed unit table.** Only the v1 unit set (`元`, `万元`, `亿元`,
  `千元`, `百万元`, `万单`, `亿单`, `万人`, `万次`, `万个`, `万`,
  `单`, `人`, `次`, `个`, `%`, `‰`) is recognised. Out-of-table
  units surface as `format_inference_failed` or `transform_unknown`
  rather than guess.
- **No reviewer UI.** Reviewers edit `mapping_review.xlsx` columns by
  hand in Excel/LibreOffice. No GUI, no web app, no inline review
  experience.
- **No LLM, no network, no Microsoft Office automation.** That is a
  deliberate design rule from `CLAUDE.md`, but it also means the
  matcher is exactly as strong as the deterministic
  value+context heuristic. There is no semantic fallback.
- **No real-file pilot has been run end-to-end in CI.** The
  acceptance smoke proves the synthetic loop holds together; it
  does **not** prove the matcher, transform table, or rendered
  display formatting generalise to any specific real report.

---

## 5. What a reviewer must verify before any real-file pilot

A reviewer signs off on real-file readiness only after every item
below is true. None of this is automated — that is the point of a
human review gate.

1. The full `python -m pytest` suite passes on the reviewer's machine
   from a clean checkout.
2. The synthetic commands in [§2](#2-commands-that-prove-the-synthetic-pipeline-works)
   above produce the expected exit codes and write every artifact
   `README.md` documents.
3. The reviewer has read `CLAUDE.md` (design rules, privacy rules)
   and `docs/real_file_pilot.md` (manual real-file workflow and
   sign-off checklist) **in full**.
4. The reviewer has run `pilot-preflight` against the four pilot
   paths intended for the real pair and observed exit `0`.
5. The real Excel + Word pair lives **outside the repo tree** (see
   `docs/real_file_pilot.md` §1). `git status` is clean of those
   paths.
6. The reviewer accepts the [known limitations in §4](#4-known-limitations--not-production-ready-yet)
   for the specific report being trialed — i.e., the report's
   numeric surface fits within paragraph + top-level table cells,
   its units fall inside the v1 unit table, and run-level
   styling drift is acceptable for the trial.
7. The reviewer has a written rollback plan for the trial: any
   non-zero exit at any stage halts the pilot, and any artifact
   discrepancy is investigated by re-running the producing stage
   from earlier audited inputs — never by hand-editing an artifact.

Once all seven hold, the manual workflow in
`docs/real_file_pilot.md` is the supported path. The
`--allow-incomplete` flag on `confirm-mapping` must **not** be used
in a real-file pilot.
