"""Repository privacy-boundary tests.

These tests reduce the risk of real company files slipping into the
repo by pinning three boundaries the rest of the pipeline already
relies on:

  1. ``git ls-files`` must not list any Office binary
     (``.xlsx`` / ``.xls`` / ``.docx`` / ``.doc``) or any generated
     learn/run/render artifact basename. Both are the highest-leakage
     surfaces for real metric values; the synthetic quickstart
     regenerates them on demand from ``generate-synthetic``.
  2. ``.gitignore`` continues to mask ``output/`` and
     ``samples/synthetic/``, and ``git check-ignore`` confirms
     representative artifact paths inside those directories are
     ignored. The `git check-ignore` check is the runtime equivalent
     of what ``git add`` would do, so a future ``.gitignore``
     reshuffle that quietly breaks masking surfaces here.
  3. The operator-facing docs (``docs/real_file_pilot.md`` and
     ``docs/command_reference.md``) still point operators to keep
     real files OUTSIDE the repo and to use the read-only
     ``pilot-preflight`` gate before the first ``learn`` run.

The suite is **read-only**: it inspects ``git ls-files``, ``.gitignore``,
and docs only. It never generates fixtures, never writes artifacts,
and never mutates the repo. It does not change matching, confirmation,
rendering, validation, pilot-summary, or pilot-preflight behaviour —
those contracts are pinned by their own dedicated test files.
"""
from __future__ import annotations

import subprocess
from pathlib import Path

import pytest


REPO_ROOT = Path(__file__).resolve().parents[1]


def _tracked_files() -> list[str]:
    """Return the repo-relative paths tracked by git, one per line."""
    result = subprocess.run(
        ["git", "ls-files"],
        cwd=str(REPO_ROOT),
        capture_output=True,
        check=True,
        text=True,
    )
    return [line for line in result.stdout.splitlines() if line]


# ---------------------------------------------------------------------------
# git ls-files — no Office binaries, no generated artifacts
# ---------------------------------------------------------------------------

FORBIDDEN_OFFICE_SUFFIXES = (".xlsx", ".xls", ".docx", ".doc")


def test_no_office_binary_files_are_tracked():
    """No ``.xlsx`` / ``.xls`` / ``.docx`` / ``.doc`` should ever be
    tracked. Even a deliberately-synthetic Office file slipping into
    the repo would establish a precedent that lowers the friction for
    a future real-data leak — and the synthetic generator can rebuild
    samples on demand, so there is no reason to commit them."""
    tracked = _tracked_files()
    offenders = sorted(
        path for path in tracked
        if path.lower().endswith(FORBIDDEN_OFFICE_SUFFIXES)
    )
    assert not offenders, (
        "Office binary files must never be committed — they are the "
        "highest-leakage surface for real company data. Move these "
        "out of the repo (see docs/real_file_pilot.md §1) and "
        "regenerate synthetic samples with "
        "`python -m src.main generate-synthetic`:\n  - "
        + "\n  - ".join(offenders)
    )


# Basenames the pipeline writes under `--out`. Each is enumerated in
# README.md and docs/command_reference.md, so this list is the contract,
# not a regex guess. Office-suffix artifacts (mapping_review.xlsx,
# converted_template.docx, run_validation.xlsx, new_report.docx) are
# also caught by the suffix test above; listing them here makes the
# basename-level guard explicit and survives a future suffix change.
GENERATED_ARTIFACT_BASENAMES = (
    "mapping_review.xlsx",
    "auto_mapping.yml",
    "converted_template.docx",
    "confidence_report.md",
    "confirmed_mapping.yml",
    "run_validation.xlsx",
    "new_report.docx",
    "render_log.yml",
)


def test_no_generated_artifact_basenames_are_tracked():
    """Generated learn/run/render artifacts must live under ``output/``
    (or an external pilot directory). The ``.gitignore`` rule masks
    ``output/``, but a basename-level check catches an artifact that
    landed anywhere else — e.g., a stray ``confidence_report.md``
    copied into ``docs/`` for sharing, or a ``confirmed_mapping.yml``
    pasted into ``samples/`` for "convenience"."""
    tracked = _tracked_files()
    offenders = sorted(
        path for path in tracked
        if Path(path).name in GENERATED_ARTIFACT_BASENAMES
    )
    assert not offenders, (
        "Generated pipeline artifacts must never be committed — they "
        "may carry real metric values. Regenerate them under "
        "`output/` and remove these copies:\n  - "
        + "\n  - ".join(offenders)
    )


# ---------------------------------------------------------------------------
# .gitignore — output/ and samples/synthetic/ stay masked
# ---------------------------------------------------------------------------

REQUIRED_GITIGNORE_RULES = ("output/", "samples/synthetic/")


def test_gitignore_masks_output_and_synthetic_samples():
    """The two on-disk directories where the pipeline writes
    (``output/``) or where ``generate-synthetic`` deposits fake
    fixtures (``samples/synthetic/``) must stay gitignored. Removing
    either rule would let a careless ``git add .`` capture every
    artifact at once."""
    text = (REPO_ROOT / ".gitignore").read_text(encoding="utf-8")
    lines = {line.strip() for line in text.splitlines() if line.strip()}
    for rule in REQUIRED_GITIGNORE_RULES:
        assert rule in lines, (
            f".gitignore must keep `{rule}` masked. Without it, real "
            "and synthetic artifacts could be tracked by an accidental "
            "`git add .` — see CLAUDE.md 'Privacy Rules' and "
            "docs/real_file_pilot.md §1."
        )


@pytest.mark.parametrize(
    "candidate",
    [
        "output/mapping_review.xlsx",
        "output/auto_mapping.yml",
        "output/new_report.docx",
        "samples/synthetic/historical.xlsx",
        "samples/synthetic/finished_report.docx",
    ],
)
def test_git_check_ignore_confirms_artifact_paths_are_ignored(candidate: str):
    """``git check-ignore`` is the runtime contract that matches what
    ``git add`` actually does. Pinning representative artifact paths
    catches a ``.gitignore`` reshuffle that breaks masking even when
    the textual rule still looks correct (e.g., a leading-comment
    typo or a stray negation)."""
    # check-ignore exits 0 when the path is ignored, 1 when it is not.
    # We don't pass check=True because exit 1 is the failure signal
    # this test is designed to surface.
    result = subprocess.run(
        ["git", "check-ignore", "--quiet", candidate],
        cwd=str(REPO_ROOT),
    )
    assert result.returncode == 0, (
        f"`git check-ignore` reports `{candidate}` is NOT ignored. "
        "The `.gitignore` rules for `output/` and `samples/synthetic/` "
        "must keep these paths masked."
    )


# ---------------------------------------------------------------------------
# Operator-facing docs — keep real files outside the repo, use
# pilot-preflight as the pre-pilot gate.
# ---------------------------------------------------------------------------

def test_real_file_pilot_doc_says_real_files_live_outside_repo():
    """``docs/real_file_pilot.md`` §1 is the human boundary that
    protects against real-data leaks. Drop the "outside this repo"
    instruction and an operator has no documented anchor for where
    real files belong."""
    text = (REPO_ROOT / "docs" / "real_file_pilot.md").read_text(encoding="utf-8")
    assert "outside this repo" in text.lower(), (
        "docs/real_file_pilot.md must state that real Excel/Word "
        "files live OUTSIDE this repository. That sentence is the "
        "human boundary that protects against real-data leaks."
    )


def test_real_file_pilot_doc_references_pilot_preflight():
    """The pilot doc must point operators at ``pilot-preflight`` as
    the read-only pre-pilot gate. Removing the reference would leave
    operators without a documented way to catch an inside-repo path
    before the first ``learn`` invocation."""
    text = (REPO_ROOT / "docs" / "real_file_pilot.md").read_text(encoding="utf-8")
    assert "pilot-preflight" in text, (
        "docs/real_file_pilot.md must reference the `pilot-preflight` "
        "command — it is the read-only gate that catches inside-repo "
        "paths before any artifact is written."
    )


def test_command_reference_doc_documents_pilot_preflight():
    """``docs/command_reference.md`` is the per-command lookup. The
    ``pilot-preflight`` contract must appear here so operators can
    look it up without reading ``docs/real_file_pilot.md`` end-to-end."""
    text = (REPO_ROOT / "docs" / "command_reference.md").read_text(encoding="utf-8")
    assert "pilot-preflight" in text, (
        "docs/command_reference.md must document `pilot-preflight` "
        "so operators have a single-page reference for the pre-pilot "
        "privacy gate."
    )
