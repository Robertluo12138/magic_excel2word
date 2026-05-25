# Operator dry-run checklist — synthetic-only rehearsal

A narrow, one-pass rehearsal an operator runs on the **synthetic**
fixture before pointing the pipeline at any real Excel + Word pair.
Every stage below uses commands and flags that already exist; this
checklist introduces **no** new runtime behaviour and writes only
into directories that `.gitignore` already masks (`output/`,
`samples/synthetic/`), so a clean run leaves the working tree
untouched.

> ⚠️ This is a **rehearsal**, not a pilot. Do not run any of the
> commands below against real company data. The supported real-file
> workflow lives in [`real_file_pilot.md`](real_file_pilot.md) and
> must only be attempted after this rehearsal completes and the
> reviewer items in
> [`milestone_1_readiness.md` §5](milestone_1_readiness.md#5-what-a-reviewer-must-verify-before-any-real-file-pilot)
> are all true.

The exit codes below are the contract pinned by
[`command_reference.md`](command_reference.md) — they are repeated
inline so the operator does not have to context-switch between docs.
A non-zero exit at any stage where `0` is expected (or `0` at a
stage where a fail-loud gate is expected to fire) means the rehearsal
has uncovered a regression; stop and investigate instead of pressing
on.

---

## 0. Prerequisites

- [ ] Python 3.10+ on `PATH` (`python --version`).
- [ ] Repo deps installed: `pip install -r requirements.txt` exits `0`.
- [ ] `git status` shows a clean working tree before you start — so
      any new artifact that *does* leak is obvious afterwards.
- [ ] You are at the repo root. Every command below is run from there.

No virtualenv is required by the pipeline itself; pick whichever
environment matches your platform. No GUI, no LLM, no network, no
Microsoft Office automation is involved at any stage.

---

## 1. Stage rehearsal

Each stage below lists: the exact command, the expected CLI exit
code, the artifacts that stage writes, and the manual observations a
reviewer should tick off **after** the command returns. Tick a stage
only when every observation passes.

### Stage 1 — Generate the synthetic fixture

**Command**

```bash
python -m src.main generate-synthetic --out samples/synthetic
```

**Expected exit code:** `0`.

**Expected artifacts** (both gitignored under `samples/synthetic/`):

- `samples/synthetic/historical.xlsx`
- `samples/synthetic/finished_report.docx`

**Reviewer observations**

- [ ] Both artifacts exist on disk after the command returns.
- [ ] `git status` still shows no new tracked paths — `samples/synthetic/`
      is gitignored, so neither file appears as untracked-to-commit.

### Stage 2 — Learn (exploratory)

**Command**

```bash
python -m src.main learn \
    --excel samples/synthetic/historical.xlsx \
    --word  samples/synthetic/finished_report.docx \
    --out   output
```

**Expected exit code:** `0`. The synthetic corpus contains deliberate
`UNRESOLVED` rows (e.g. `提升15%`, `提升0.70个百分点`), so the CLI
prints a loud stderr warning but still exits `0` in exploratory mode.

**Expected artifacts** (all under `output/`, which is gitignored):

- `output/mapping_review.xlsx`
- `output/auto_mapping.yml`
- `output/converted_template.docx`
- `output/confidence_report.md`

**Reviewer observations**

- [ ] All four artifacts exist on disk.
- [ ] `output/confidence_report.md`'s top sections lead with what
      needs review (`UNRESOLVED`, `LOW`, ambiguous, `EXCLUDED`) —
      not with `HIGH` wins.
- [ ] `output/mapping_review.xlsx` has stable `Word ID` values
      starting at `word_0001` in its first column.
- [ ] `output/converted_template.docx` opens in a docx viewer
      (LibreOffice with macros disabled, or a docx text dump) and
      shows `{{ word_NNNN }}` placeholders only where the matcher
      reported `HIGH`/`MEDIUM`. `LOW`, `UNRESOLVED`, and `EXCLUDED`
      values remain visible as their original text.

### Stage 3 — Trust gate rehearsal (`learn --strict`)

**Command**

```bash
python -m src.main learn \
    --excel samples/synthetic/historical.xlsx \
    --word  samples/synthetic/finished_report.docx \
    --out   output \
    --strict
```

**Expected exit code:** `3` — this is the gate that protects a real
pilot from confirming an incomplete mapping. The artifacts from
Stage 2 are rewritten in place.

**Expected artifacts:** same four files as Stage 2 (overwritten;
exit `3` does **not** suppress artifact writes).

**Reviewer observations**

- [ ] Stderr contains a `STRICT GATE FAILED` banner naming the
      counts (`UNRESOLVED=…, LOW=…`) and pointing the operator at
      `mapping_review.xlsx` and `confidence_report.md` for the
      offending rows. The stderr banner intentionally does **not**
      list individual Word tokens — sample tokens from each problem
      bucket are emitted on **stdout** by the console summary
      printed above the banner.
- [ ] You can articulate, in one sentence, why the same input that
      exited `0` in Stage 2 now exits `3` — the strictness flag is
      the only difference.

### Stage 4 — Cross-check learn-mode artifacts

**Command**

```bash
python -m src.main validate-artifacts --out output
```

**Expected exit code:** `0`. The four artifacts from Stage 2/3 must
agree on every `word_id`, location, raw token, confidence/status,
and placeholder status.

**Expected artifacts:** none. `validate-artifacts` is read-only.

**Reviewer observations**

- [ ] Stdout reports a clean cross-check with no per-issue list.
- [ ] No file under `output/` has been modified by this command
      (compare timestamps if uncertain).

### Stage 5 — Reviewer-handoff gate rehearsal (`confirm-mapping`)

**Command**

```bash
python -m src.main confirm-mapping \
    --auto   output/auto_mapping.yml \
    --review output/mapping_review.xlsx \
    --out    output/confirmed_mapping.yml
```

**Expected exit code:** `5` — `mapping_review.xlsx` is fresh from
Stage 2 with blank `Reviewer Decision` columns, so every eligible
row lands in `review_required`. The gate is doing its job by
refusing to claim completeness.

**Expected artifacts:** `output/confirmed_mapping.yml` is still
written so the operator can see exactly what is blocking. The file
must record `summary.complete: false`.

**Reviewer observations**

- [ ] `output/confirmed_mapping.yml` exists.
- [ ] `summary.complete` is the literal boolean `false` (or the
      key is present and not `true`).
- [ ] The `review_required` list is non-empty and includes the
      `UNRESOLVED` rows surfaced in Stage 3 (search for
      `unresolved_no_candidate` to confirm).
- [ ] You can articulate, in one sentence, what the operator would
      do next on a real pilot — fill in `Reviewer Decision` /
      `Confirmed Sheet` / `Confirmed Cell` columns by hand, then
      re-run this command — and that `--allow-incomplete` is the
      exploratory escape hatch that **must not** be used in a
      real-file pilot.

### Stage 6 — Paste-safe progress lens (`pilot-summary`)

**Command**

```bash
python -m src.main pilot-summary --out output
```

**Expected exit code:** `0`.

**Expected artifacts:** none. `pilot-summary` is read-only and
contractually redacted.

**Reviewer observations**

- [ ] Stdout reports per-stage aggregate counts and a next-action
      hint per stage.
- [ ] Stdout does **not** print any raw Word token, generated value,
      raw Excel value, source sheet/cell content, reviewer note,
      individual `word_id` value, or file path beyond a basename —
      that is the paste-safe contract this command guarantees.
- [ ] The header reduces `--out` to its basename (`output`), not the
      full path.

### Stage 7 — End-to-end success via the acceptance smoke

The synthetic-corpus rehearsal above stops at `confirm-mapping`
(exit `5`) by design — the corpus has `UNRESOLVED` rows that a
real reviewer would have to resolve before `run-preview` and
`render-docx` can succeed. The acceptance smoke uses an inline
clean synthetic pair (3 `HIGH` + 2 `EXCLUDED`, zero eligible
`UNRESOLVED`/`LOW`) so it can walk all six pipeline stages plus the
post-render tamper case end-to-end without `--allow-incomplete`.

**Command**

```bash
python -m pytest tests/test_acceptance_smoke.py -v
```

**Expected exit code:** `0`.

**Expected artifacts:** none committed. The smoke writes only
inside pytest's `tmp_path` fixture and the directory is deleted on
exit.

**Reviewer observations**

- [ ] Both `test_full_acceptance_smoke_python_api` and
      `test_full_acceptance_smoke_cli` pass.
- [ ] `test_full_pipeline_halts_on_post_render_tamper` passes —
      this is the negative path that proves `validate-render`
      refuses a hand-edited `run_validation.xlsx`.

### Stage 8 — Privacy refusal rehearsal (`pilot-preflight`)

Optional but recommended: practice triggering the `pilot-preflight`
privacy refusal on an inside-repo path so the operator recognises
exit code `12` before encountering it on a real pilot. Use the
synthetic fixture paths so no real data is involved.

**Command**

```bash
python -m src.main pilot-preflight \
    --historical-excel samples/synthetic/historical.xlsx \
    --historical-word  samples/synthetic/finished_report.docx \
    --new-excel        samples/synthetic/historical.xlsx \
    --out              output
```

**Expected exit code:** `12` — every path resolves inside the repo
tree, which `pilot-preflight` refuses by design (strictly stronger
than the `samples/` advisory). The refusal protects a real pilot
from a mistyped inside-repo path before any artifact is written.

**Expected artifacts:** none. `pilot-preflight` is read-only.

**Reviewer observations**

- [ ] Stderr (not stdout — `pilot-preflight` routes the failure
      report to stderr while the success report goes to stdout)
      names each offending flag with status `inside-repo`.
- [ ] You can articulate, in one sentence, the rule the gate is
      enforcing — real Excel/Word files must live outside the repo
      tree, per [`real_file_pilot.md` §1](real_file_pilot.md#1-where-real-files-must-live).

---

## 2. Post-rehearsal hygiene

After every stage above ticks, before claiming the rehearsal done:

- [ ] `git status` shows **no** new tracked or staged paths under
      `output/`, `samples/synthetic/`, or anywhere else. The
      `.gitignore` rules for `output/` and `samples/synthetic/`
      should hide every artifact this rehearsal wrote; if they
      surface as untracked-to-commit, **stop** and investigate
      before any `git add`.
- [ ] You did **not** run `git add -A` or `git add .` at any point
      during the rehearsal — only specific source-file paths if
      any source edits happened.
- [ ] You did **not** commit, share, screenshot, or paste raw
      content from `output/confidence_report.md`,
      `output/mapping_review.xlsx`, or `output/confirmed_mapping.yml`
      anywhere outside your local machine. (On the synthetic
      fixture the content is harmless; the discipline matters
      because real-file pilots use the same artifact names.)
- [ ] Optionally, delete the `output/` directory once the rehearsal
      is complete (`rm -rf output` on macOS/Linux,
      `Remove-Item -Recurse -Force output` on PowerShell). The
      next rehearsal or pilot will recreate it.

---

## 3. Pass/fail summary

The rehearsal passes only when **every** observation above ticks.
A failed observation means one of:

- A safety gate that should have fired did not (e.g. Stage 3 exited
  `0` instead of `3`, or Stage 5 exited `0` instead of `5`).
  Investigate the matcher or the gate before any real-file pilot —
  the deterministic loop has regressed.
- A safety gate that should have passed did not (e.g. Stage 4 exited
  `4` on a fresh synthetic fixture). Investigate the artifact writer
  before any real-file pilot — the artifacts are inconsistent.
- The expected artifacts are missing or live in the wrong directory.
  Investigate `--out` handling and `.gitignore` coverage before any
  real-file pilot — a leaked artifact under a tracked path is a
  privacy hazard.
- The paste-safe redaction in Stage 6 leaked a value it should not
  have. Investigate `pilot-summary` before any real-file pilot —
  the same redaction protects a real ticket / chat paste.

If every observation ticks, the operator has rehearsed the gates
the real-file pilot relies on. Proceed to
[`real_file_pilot.md`](real_file_pilot.md) only after the reviewer
items in
[`milestone_1_readiness.md` §5](milestone_1_readiness.md#5-what-a-reviewer-must-verify-before-any-real-file-pilot)
are also all true.
