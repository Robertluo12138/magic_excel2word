# Verification matrix — safety gates ↔ tests

A narrow, reviewer-facing mapping from each safety gate in
[`milestone_1_readiness.md` §3](milestone_1_readiness.md#3-current-safety-gates)
to the **exact pytest file(s)** that pin it and the CLI smoke
coverage that exercises it end-to-end. Use this matrix to answer the
question *"if this gate regressed, which test would catch it?"* without
re-reading every test file.

This document is an **index**, not a contract — it mirrors the gate
definitions already pinned by `milestone_1_readiness.md` and the
exit-code map in [`command_reference.md`](command_reference.md). The
producing CLI command and the exit code are repeated here only so a
reader can scan the matrix in isolation; behavioural details live in
those two docs.

A small doc consistency test
(`tests/test_verification_matrix_docs.py`) re-reads this file and
fails if any `tests/<file>.py` it references no longer exists, so a
renamed or deleted test file cannot silently rot this matrix.

---

## Conventions

- **Producing CLI** — the subcommand whose `return <N>` emits the
  exit code (or, for the repo privacy boundary row, the guard test
  whose failure is itself the gate).
- **Exit code** — the gate-specific code listed in
  `command_reference.md`'s
  [exit-code map](command_reference.md#exit-code-map-at-a-glance).
  Shared code `0` (success) is covered separately in
  [§2](#2-shared-exit-code-0-success-coverage); shared code `2`
  (missing/invalid input) is covered in
  [§3](#3-shared-exit-code-2-missinginvalid-input-coverage).
- **Pytest file(s) that pin the gate** — the test file(s) under
  `tests/` whose assertions fail if the gate stops firing. A single
  file may pin multiple gates; a single gate may be pinned by more
  than one file when the per-row failure surface is wide
  (`validate-render`, `run-preview`, `render-docx`).
- **Pinned directly vs. indirectly** — "directly" means at least one
  test asserts the CLI exit-code contract literally (e.g.
  `assert cli_main([...]) == N`). "Indirectly" means the file pins
  the underlying *detection* at the Python-API level (e.g.
  `assert not report.ok`) and the CLI translates that to the
  documented exit code via a one-line mapping in `src/main.py`. The
  gate is still enforced in both cases; the difference matters when
  a reviewer asks "would a test fail if I changed the exit-code
  integer in `src/main.py`?" — only directly-pinned gates would.
- **CLI smoke coverage** — whether the end-to-end acceptance smoke in
  `tests/test_acceptance_smoke.py` exercises the gate on the
  success path, the failure path, or both. "Success path" means the
  smoke invokes the producing CLI, the gate fires, and exit `0` is
  asserted — i.e. the gate is exercised but does not refuse.
  "Failure path" means the smoke deliberately violates the gate
  and asserts the gate-specific non-zero exit. The acceptance smoke
  is happy-path by construction except for the single tamper case
  in `test_full_pipeline_halts_on_post_render_tamper`.

---

## 1. Gate × test matrix

| # | Gate | Producing CLI | Exit code | Pytest file(s) that pin the gate | CLI smoke coverage |
| --- | --- | --- | --- | --- | --- |
| 1 | `learn --strict` refuses any eligible `UNRESOLVED` or `LOW` row | `learn` | `3` | `tests/test_strict_and_docs.py` (directly) | success path (`0`) directly in `tests/test_acceptance_smoke.py::test_full_acceptance_smoke_cli` and `::test_full_pipeline_halts_on_post_render_tamper` (both pass `--strict` and assert `0`). The Python-API smoke does **not** exercise `--strict` (no such flag on the API surface). Failure path (`3`) pinned in `tests/test_strict_and_docs.py`, not in the acceptance smoke. |
| 2 | `validate-artifacts` cross-checks the four learn-mode artifacts agree by `word_id`, location, raw token, confidence, placeholder status | `validate-artifacts` | `4` | `tests/test_validate_artifacts.py` (**indirectly**: **5** failure-path tests assert `not report.ok` directly (`test_missing_artifact_is_flagged`, `test_duplicate_word_id_in_xlsx_is_flagged`, `test_duplicate_word_id_in_yaml_is_flagged`, `test_extra_xlsx_row_breaks_one_to_one`, `test_dropped_yaml_mapping_breaks_one_to_one`); **10 more** tests assert only that a specific `issues` code appears (e.g. `"location_mismatch" in _codes(report)`) — those pin drift *detection* but not the gate boolean. The CLI exit code `4` itself is emitted by the one-line `return 0 if report.ok else 4` in `src/main.py` and is **not** directly asserted by any `cli_main([...]) == 4` test) | success path (`0`) directly in `tests/test_acceptance_smoke.py` (both API and CLI surfaces) and in `tests/test_validate_artifacts.py::test_cli_validate_artifacts_passes_on_fresh_learn`. Failure path (`4`) pinned only indirectly — see the previous column. |
| 3 | `confirm-mapping` refuses to claim completeness when any row is in `review_required` (and `--allow-incomplete` was not passed) | `confirm-mapping` | `5` | `tests/test_confirm_mapping.py` (directly) | success path (`0`) directly in `tests/test_acceptance_smoke.py` (both surfaces). Failure path (`5`) pinned in `tests/test_confirm_mapping.py`. |
| 4 | `run-preview` refuses to start when `confirmed_mapping.yml` cannot prove completeness (missing `summary`, `complete` not literal `true`, `allow_incomplete: true`, non-empty `review_required`, zero confirmed rows). No artifact written. | `run-preview` | `6` | `tests/test_run_preview.py` (directly) | success path (`0`) directly in `tests/test_acceptance_smoke.py` and `tests/test_pilot_summary_pipeline_smoke.py` — the smoke's confirmed mapping satisfies the completeness check, so the gate fires and passes. Failure path (`6`) pinned in `tests/test_run_preview.py`, not in the acceptance smoke. |
| 5 | `run-preview` refuses any per-row failure on the new workbook (missing sheet/cell, empty/non-numeric cell, missing/unknown transform). Artifact still written. | `run-preview` | `7` | `tests/test_run_preview.py` (directly) | success path (`0`) directly in `tests/test_acceptance_smoke.py` and `tests/test_pilot_summary_pipeline_smoke.py`. Failure path (`7`) pinned in `tests/test_run_preview.py`, not in the acceptance smoke. |
| 6 | `render-docx` fatal input gate (a `{{ word_NNNN }}` template placeholder has no validation row). No docx or log written. | `render-docx` | `8` | `tests/test_renderer.py` (directly) | success path (`0`) directly in `tests/test_acceptance_smoke.py` and `tests/test_pilot_summary_pipeline_smoke.py` — every successful `render-docx` invocation passes this gate. Failure path (`8`) pinned in `tests/test_renderer.py`, not in the acceptance smoke. |
| 7 | `render-docx` per-row gate (non-ok validation row, missing/duplicated `Generated Value`, unused `word_id` in validation, raw-token format inference failure). No docx or log written. | `render-docx` | `9` | `tests/test_renderer.py` (directly) | success path (`0`) directly in `tests/test_acceptance_smoke.py` and `tests/test_pilot_summary_pipeline_smoke.py`. Failure path (`9`) pinned in `tests/test_renderer.py`, not in the acceptance smoke. |
| 8 | `validate-render` cross-checks the three rendered-output artifacts agree (leftover placeholder, missing/extra log row, value drift, missing display text in docx, zero placeholder occurrences, non-ok status) | `validate-render` | `10` | `tests/test_validate_render.py` (directly) | success path (`0`) directly in `tests/test_acceptance_smoke.py` and `tests/test_pilot_summary_pipeline_smoke.py`. Failure path (`10`) pinned in both `tests/test_validate_render.py` and `tests/test_acceptance_smoke.py::test_full_pipeline_halts_on_post_render_tamper` (the post-render tamper case is the sole negative path in the acceptance smoke). |
| 9 | `pilot-summary` refuses to summarize when `auto_mapping.yml` is absent — nothing meaningful to report | `pilot-summary` | `11` | `tests/test_pilot_summary.py` (directly) | `pilot-summary` is **not** invoked by the acceptance smoke at all. Success path (`0`) on real pipeline artifacts pinned by `tests/test_pilot_summary_pipeline_smoke.py`. Failure path (`11`) pinned in `tests/test_pilot_summary.py`. |
| 10 | `pilot-preflight` refuses any pilot path that resolves inside the repo tree (strictly stronger than the `samples/` advisory) | `pilot-preflight` | `12` | `tests/test_pilot_preflight.py` (directly) | `pilot-preflight` is **not** invoked by the acceptance smoke at all. Both success path (`0`) and failure path (`12`) pinned in `tests/test_pilot_preflight.py`, including the precedence rule that `12` outranks `2` when both apply. |
| 11 | Privacy preflight advisory: stderr banner fires when any CLI path argument resolves inside `samples/`. Informational only — exit code unchanged. | `learn`, `validate-artifacts`, `confirm-mapping`, `run-preview`, `render-docx`, `validate-render` | (no code; stderr banner) | `tests/test_preflight.py` (directly: asserts `"PRIVACY PREFLIGHT ADVISORY" in stderr` after `cli_main([...])`) | not exercised by the acceptance smoke (the smoke uses `tmp_path`, never `samples/`). Banner + per-command flag coverage pinned in `tests/test_preflight.py`. |
| 12 | Repo privacy boundary: refuses any tracked `.xlsx`/`.xls`/`.docx`/`.doc`, any tracked generated-artifact basename, any `.gitignore` regression that stops masking `output/` or `samples/synthetic/`, and any operator-doc regression that no longer points real files outside the repo or no longer references `pilot-preflight` | guard test (not a CLI) | test failure | `tests/test_repo_privacy_boundary.py` (directly: a test failure *is* the gate) | not a CLI gate; the acceptance smoke does not invoke it. Its enforcement depends entirely on whether pytest's collection includes this file — i.e. on a bare `python -m pytest` it runs, but on a narrowed selector (e.g. `python -m pytest tests/test_acceptance_smoke.py`) it does not. There is no automatic coupling to the acceptance smoke. |

---

## 2. Shared exit code `0` (success) coverage

Exit code `0` is the per-stage success signal. The tests that walk
every pipeline stage end-to-end on a clean inline synthetic pair
and assert `cli_main([...]) == 0` (or the equivalent API success)
at each step are:

- `tests/test_acceptance_smoke.py::test_full_acceptance_smoke_python_api`
  — Python-API surface, stages `learn` → `validate-artifacts` →
  `confirm-mapping` → `run-preview` → `render-docx` → `validate-render`.
  Does **not** drive `learn --strict` (no such flag on the API
  surface).
- `tests/test_acceptance_smoke.py::test_full_acceptance_smoke_cli`
  — CLI surface, same stages via `cli_main([...])`, with
  `learn --strict` exercised explicitly.
- `tests/test_pilot_summary_pipeline_smoke.py::test_pilot_summary_on_real_pipeline_artifacts`
  — the same pipeline plus `pilot-summary`, so the pilot-summary
  success path is covered against real pipeline artifacts.

What the acceptance smoke does and does not exercise:

- **Gates 1, 2, 3, 4, 5, 6, 7, 8 — success paths exercised.** Every
  successful CLI invocation passes through the underlying gate by
  construction (e.g. a clean `run-preview` exit `0` means the
  completeness check at gate 4 fired and accepted the input, and
  the per-row check at gate 5 fired and accepted every row).
- **Gates 9, 10 — not exercised at all.** `pilot-summary` and
  `pilot-preflight` are not invoked by the acceptance smoke. Gate 9
  is exercised at success by the pilot-summary pipeline smoke; gate
  10 lives entirely in `tests/test_pilot_preflight.py`.
- **Gate 11 — not exercised.** The acceptance smoke runs entirely
  under `tmp_path`, so the privacy advisory banner is never
  triggered. Pinned in `tests/test_preflight.py`.
- **Failure paths — only gate 8.** The only failure-path assertion
  inside the acceptance smoke is the post-render tamper case
  asserting `cli_main(["validate-render", ...]) == 10`. Every
  other gate's failure path is pinned by its dedicated per-gate
  test file.

---

## 3. Shared exit code `2` (missing/invalid input) coverage

Exit code `2` is the shared input-error code emitted by every
subcommand. Per-command pins:

| Producing CLI | Pytest file(s) that pin exit `2` |
| --- | --- |
| `learn` | `tests/test_learn_smoke.py` |
| `validate-artifacts` | `tests/test_validate_artifacts.py` |
| `confirm-mapping` | `tests/test_confirm_mapping.py` |
| `run-preview` | `tests/test_run_preview.py` |
| `render-docx` | `tests/test_renderer.py` |
| `validate-render` | `tests/test_validate_render.py` |
| `pilot-summary` | `tests/test_pilot_summary.py` |
| `pilot-preflight` | `tests/test_pilot_preflight.py` |

The acceptance smoke does not exercise exit `2` — its inputs are
always valid paths under `tmp_path`.

---

## 4. Parity tests that prevent drift

Four pure-docs tests already guard the contracts this matrix
depends on. They are listed here so a reader can see the full set
of safety nets at a glance:

| Pytest file | What it guards |
| --- | --- |
| `tests/test_command_reference_docs.py` | Every CLI subcommand has its own `## \`<name>\`` section in `docs/command_reference.md` with all six contract fields. Every gate-specific `return <N>` in `src/main.py` (codes ≥ 3) is listed in the exit-code map. |
| `tests/test_milestone_readiness_docs.py` | `docs/milestone_1_readiness.md` exists, is linked from `README.md` and `docs/real_file_pilot.md`, declares prototype status, and contains no positive-readiness marketing claims. |
| `tests/test_verification_matrix_docs.py` | This matrix doc exists, is linked from `docs/milestone_1_readiness.md`, and **every `tests/<file>.py` it references actually exists on disk** — a renamed or deleted test file cannot silently rot the matrix. |
| `tests/test_pilot_result_template_docs.py` | `docs/pilot_result_template.md` exists and is linked from both `docs/real_file_pilot.md` and `docs/milestone_1_readiness.md`, so the redacted reviewer form a sponsor relies on cannot silently disappear or become orphaned. |

---

## 5. What this matrix is **not**

- Not a substitute for `milestone_1_readiness.md` §3 — that doc is
  the gate-definition source of truth. This matrix only adds the
  test-mapping lens.
- Not a substitute for the exit-code map in
  `command_reference.md` — that map is the per-command contract
  consulted by automation. This matrix only adds the
  test-mapping lens.
- Not a coverage report. "Pinned by file X" means at least one
  assertion in `X` would fail if the gate regressed; it does **not**
  claim every behavioural edge case of the gate is tested.
- Not auto-generated. The matrix is hand-maintained. The doc
  consistency test guarantees the *file references* stay valid;
  individual test-name accuracy is the author's responsibility when
  a gate is reshaped.
