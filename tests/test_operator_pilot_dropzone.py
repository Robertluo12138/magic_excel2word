"""Tests for the ``operator_pilot/`` learn-only drop-zone.

``operator_pilot/`` is the in-repo drop-zone a non-technical operator
uses to run the *first* real-file learn-only matching test: copy one
historical Excel + the matching finished Word report into
``operator_pilot/input/``, paste the fixed prompt from
``operator_pilot/PROMPT_FOR_AGENT.md`` to an Agent, and read the
artifacts the Agent writes to ``operator_pilot/output/``.

The contracts pinned here keep the drop-zone narrow:

  1. The four drop-zone files exist on disk: ``README.zh-CN.md``,
     ``PROMPT_FOR_AGENT.md``, ``input/.gitkeep``, and
     ``output/.gitkeep``. Removing any of them silently breaks the
     documented workflow.
  2. The Chinese operator README frames the test as learn-only and
     points at both drop-zone subdirectories plus the fixed prompt
     file — without those pointers an operator has no documented
     anchor.
  3. The fixed agent prompt names each of the three allowed
     learn-mode commands (``learn --strict``, ``validate-artifacts``,
     ``pilot-summary``) and explicitly forbids the four post-learn
     commands (``confirm-mapping``, ``run-preview``, ``render-docx``,
     ``validate-render``) plus every category of raw content the
     agent must never paste into chat.
  4. ``.gitignore`` masks every real input and every generated
     learn-mode artifact under ``operator_pilot/input/`` /
     ``operator_pilot/output/`` while keeping the four tracked
     files visible. ``git check-ignore`` is the runtime contract
     that matches what ``git add`` actually does.

The suite is **read-only**: it inspects on-disk docs, ``.gitignore``,
and ``git check-ignore`` only. It never imports from ``src/``, never
executes the CLI, never writes files, and never touches real or
synthetic data. It does not change matching, confirmation,
rendering, validation, pilot-summary, or pilot-preflight behaviour —
those contracts are pinned by their own dedicated test files.
"""
from __future__ import annotations

import subprocess
from pathlib import Path

import pytest


REPO_ROOT = Path(__file__).resolve().parents[1]
PILOT_DIR = REPO_ROOT / "operator_pilot"
README_DOC = PILOT_DIR / "README.zh-CN.md"
PROMPT_DOC = PILOT_DIR / "PROMPT_FOR_AGENT.md"
INPUT_GITKEEP = PILOT_DIR / "input" / ".gitkeep"
OUTPUT_GITKEEP = PILOT_DIR / "output" / ".gitkeep"


# ---------------------------------------------------------------------------
# Drop-zone layout
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    "path",
    [README_DOC, PROMPT_DOC, INPUT_GITKEEP, OUTPUT_GITKEEP],
    ids=lambda p: str(p.relative_to(REPO_ROOT)),
)
def test_operator_pilot_layout_file_exists(path: Path):
    """The four tracked files are the contract: an operator-facing
    Chinese README, the fixed agent prompt, and two .gitkeep markers
    so the input/ and output/ subdirectories survive a fresh clone.
    Drop any of them and the documented workflow breaks silently."""
    assert path.exists(), (
        f"operator_pilot drop-zone file missing: "
        f"{path.relative_to(REPO_ROOT)}. See "
        "operator_pilot/README.zh-CN.md for the documented workflow."
    )


# ---------------------------------------------------------------------------
# README frames the workflow as learn-only and points at the drop-zone
# ---------------------------------------------------------------------------

def test_readme_zh_cn_frames_workflow_as_learn_only():
    """If the Chinese README loses its learn-only framing, the
    drop-zone gets quietly repurposed for confirm/render work that
    is not yet ready for real data. Accept either the English
    ``learn-only`` / ``learn-mode`` spelling or the Chinese
    ``learn 模式`` so a future bilingual rewrite still passes."""
    text = README_DOC.read_text(encoding="utf-8")
    lowered = text.lower()
    assert (
        "learn-only" in lowered
        or "learn-mode" in lowered
        or "learn 模式" in text
    ), (
        "operator_pilot/README.zh-CN.md must describe the test as a "
        "learn-only / learn-mode inspection so an operator does not "
        "accidentally promote it into a confirm/render workflow."
    )


@pytest.mark.parametrize(
    "anchor",
    [
        "operator_pilot/input/historical.xlsx",
        "operator_pilot/input/finished_report.docx",
        "operator_pilot/output",
        "PROMPT_FOR_AGENT.md",
    ],
)
def test_readme_zh_cn_points_at_each_drop_zone_anchor(anchor: str):
    """The README is the only documentation an operator sees before
    pasting the prompt. It must name both drop-zone subdirectories,
    the two required input basenames, and the fixed prompt file —
    without those pointers an operator has nowhere to look."""
    text = README_DOC.read_text(encoding="utf-8")
    assert anchor in text, (
        f"operator_pilot/README.zh-CN.md must mention `{anchor}` so "
        "an operator knows exactly where to drop real files, where "
        "to look for artifacts, or which file to copy as the prompt."
    )


# ---------------------------------------------------------------------------
# Fixed agent prompt: allowed commands
# ---------------------------------------------------------------------------

ALLOWED_COMMAND_TOKENS = (
    "learn --strict",
    "validate-artifacts",
    "pilot-summary",
)


@pytest.mark.parametrize("token", ALLOWED_COMMAND_TOKENS)
def test_prompt_mentions_each_allowed_command(token: str):
    """The fixed agent prompt is the source of truth for what the
    learn-only pilot is allowed to run. Drop any of the three
    learn-mode inspection commands and the drop-zone silently
    becomes a partial pilot."""
    text = PROMPT_DOC.read_text(encoding="utf-8")
    assert token in text, (
        f"operator_pilot/PROMPT_FOR_AGENT.md must instruct the agent "
        f"to run `{token}` — it is one of the three learn-mode "
        "inspection commands the drop-zone covers."
    )


# ---------------------------------------------------------------------------
# Fixed agent prompt: forbidden commands + prohibition language
# ---------------------------------------------------------------------------

FORBIDDEN_COMMAND_TOKENS = (
    "confirm-mapping",
    "run-preview",
    "render-docx",
    "validate-render",
)


@pytest.mark.parametrize("token", FORBIDDEN_COMMAND_TOKENS)
def test_prompt_explicitly_names_each_post_learn_command(token: str):
    """A learn-only pilot must not let the agent slip into confirm,
    run-preview, render-docx, or validate-render. The fixed prompt
    must name each one by string so the agent has no plausible
    excuse to escalate the scope."""
    text = PROMPT_DOC.read_text(encoding="utf-8")
    assert token in text, (
        f"operator_pilot/PROMPT_FOR_AGENT.md must name `{token}` so "
        "the agent has an explicit list of forbidden post-learn "
        "commands. Implicitly forbidding leaves room for the agent "
        "to escalate the scope."
    )


def test_prompt_carries_explicit_prohibition_language():
    """Naming the forbidden commands is not enough — they must sit
    next to explicit prohibition language. Accept either the Chinese
    「禁止」 marker or a common English equivalent so a future
    bilingual rewrite still passes."""
    text = PROMPT_DOC.read_text(encoding="utf-8")
    lowered = text.lower()
    assert (
        "禁止" in text
        or "forbid" in lowered
        or "do not" in lowered
        or "must not" in lowered
    ), (
        "operator_pilot/PROMPT_FOR_AGENT.md must carry explicit "
        "prohibition language (e.g. 「禁止」 / `forbid` / `do not` / "
        "`must not`). Without it, the forbidden-command list reads "
        "as advisory rather than mandatory."
    )


# ---------------------------------------------------------------------------
# Fixed agent prompt: forbidden paste categories
# ---------------------------------------------------------------------------

# Each entry is a (label, acceptable-spellings) pair. The acceptable
# spellings list lets the contract survive small bilingual wording
# tweaks; failing the test means the prompt no longer warns about
# that category at all.
PROHIBITED_PASTE_CATEGORIES = (
    ("Word numbers", ("Word 报告", "Word number", "Word 中的数字")),
    ("Excel raw values", ("Excel", "单元格", "raw excel value")),
    ("source sheet / cell address", ("Sheet1!", "sheet 名", "cell 地址")),
    ("absolute paths", ("绝对路径", "absolute path", "/Users/", "C:\\")),
    ("company identifiers", ("公司名", "company", "客户名", "company identifier")),
    ("commit / stage operations", ("git add", "git commit")),
)


@pytest.mark.parametrize(
    "label, candidates",
    PROHIBITED_PASTE_CATEGORIES,
    ids=[label for label, _ in PROHIBITED_PASTE_CATEGORIES],
)
def test_prompt_names_each_forbidden_paste_category(label, candidates):
    """The fixed prompt must warn the agent about each category of
    raw content the goal forbids leaking into chat or git. A missing
    warning is the same as silent permission — the goal lists these
    six categories explicitly."""
    text = PROMPT_DOC.read_text(encoding="utf-8")
    assert any(token in text for token in candidates), (
        f"operator_pilot/PROMPT_FOR_AGENT.md must warn about the "
        f"`{label}` category by name (any of: {candidates}) so the "
        "agent cannot quietly leak it into chat or git."
    )


# ---------------------------------------------------------------------------
# .gitignore — masking contract via literal rules + git check-ignore
# ---------------------------------------------------------------------------

REQUIRED_GITIGNORE_LINES = (
    "operator_pilot/input/*",
    "!operator_pilot/input/.gitkeep",
    "operator_pilot/output/*",
    "!operator_pilot/output/.gitkeep",
)


@pytest.mark.parametrize("rule", REQUIRED_GITIGNORE_LINES)
def test_gitignore_carries_operator_pilot_masking_rule(rule: str):
    """The four rules together implement "mask everything except
    .gitkeep" for both drop-zone subdirectories. Removing any single
    one breaks the contract: either real files start leaking through
    or the .gitkeep marker disappears from a clean clone."""
    text = (REPO_ROOT / ".gitignore").read_text(encoding="utf-8")
    lines = {line.strip() for line in text.splitlines() if line.strip()}
    assert rule in lines, (
        f".gitignore must keep `{rule}` so real files and generated "
        "outputs under operator_pilot/ stay out of commits while the "
        ".gitkeep markers remain tracked."
    )


# `git check-ignore` is the runtime contract that matches what `git
# add` actually does. Pinning representative paths catches a future
# `.gitignore` reshuffle that breaks masking even when the textual
# rules still look correct (e.g. an unanchored `output/` rule that
# pulls in the nested operator_pilot/output/ directory and silently
# defeats the file-level negation).
DROPZONE_PAYLOAD_PATHS = (
    "operator_pilot/input/historical.xlsx",
    "operator_pilot/input/finished_report.docx",
    "operator_pilot/output/mapping_review.xlsx",
    "operator_pilot/output/auto_mapping.yml",
    "operator_pilot/output/converted_template.docx",
    "operator_pilot/output/confidence_report.md",
)


@pytest.mark.parametrize("candidate", DROPZONE_PAYLOAD_PATHS)
def test_git_check_ignore_masks_dropzone_payload(candidate: str):
    """Every real input and every learn-mode artifact under the
    drop-zone must be reported as ignored by ``git check-ignore``."""
    result = subprocess.run(
        ["git", "check-ignore", "--quiet", candidate],
        cwd=str(REPO_ROOT),
    )
    assert result.returncode == 0, (
        f"`git check-ignore` reports `{candidate}` is NOT ignored. "
        "The operator_pilot drop-zone must mask every real input "
        "and every generated learn-mode artifact under input/ and "
        "output/ — otherwise a stray `git add .` could capture real "
        "company data."
    )


DROPZONE_TRACKED_PATHS = (
    "operator_pilot/README.zh-CN.md",
    "operator_pilot/PROMPT_FOR_AGENT.md",
    "operator_pilot/input/.gitkeep",
    "operator_pilot/output/.gitkeep",
)


@pytest.mark.parametrize("candidate", DROPZONE_TRACKED_PATHS)
def test_git_check_ignore_keeps_tracked_files_visible(candidate: str):
    """The README, the fixed prompt, and the two .gitkeep markers
    are the *only* files that should remain tracked under
    operator_pilot/. A future `.gitignore` rule that masks any of
    them by accident would break a clean clone of the drop-zone."""
    # check-ignore exits 0 when the path is ignored, 1 when it is
    # not. We want 1 here: these files must stay visible to git.
    result = subprocess.run(
        ["git", "check-ignore", "--quiet", candidate],
        cwd=str(REPO_ROOT),
    )
    assert result.returncode == 1, (
        f"`git check-ignore` reports `{candidate}` IS ignored. "
        "operator_pilot/README.zh-CN.md, PROMPT_FOR_AGENT.md, and "
        "the two .gitkeep markers must remain tracked so a fresh "
        "clone sees the drop-zone layout."
    )


# ---------------------------------------------------------------------------
# Bypass-narrowing invariants — keep the in-repo exception scoped
# ---------------------------------------------------------------------------

ALLOWED_DROPZONE_TRACKED_FILES = frozenset(DROPZONE_TRACKED_PATHS)


def test_no_unexpected_files_tracked_under_operator_pilot():
    """``operator_pilot/`` is a deliberate, narrow exception to the
    "real files outside this repo" rule in
    ``docs/real_file_pilot.md`` §1 / §1a. The only files that may
    ever be tracked under it are the README, the fixed prompt, and
    the two ``.gitkeep`` markers. A file outside that set — a stray
    real ``.xlsx`` force-added past ``.gitignore``, a new
    documentation page, an extra config — silently widens the
    bypass surface and breaks the drop-zone's stated scope. CI is
    the last line of defense; this test is it."""
    result = subprocess.run(
        ["git", "ls-files", "operator_pilot/"],
        cwd=str(REPO_ROOT),
        capture_output=True,
        check=True,
        text=True,
    )
    tracked = frozenset(line for line in result.stdout.splitlines() if line)
    extras = sorted(tracked - ALLOWED_DROPZONE_TRACKED_FILES)
    assert not extras, (
        "operator_pilot/ tracked-files contract drifted. The "
        "drop-zone is a narrow learn-only exception to the "
        "'real files outside this repo' rule in "
        "docs/real_file_pilot.md §1 / §1a; widening the tracked "
        "set silently dilutes the privacy gate:\n  - "
        + "\n  - ".join(extras)
    )


# ---------------------------------------------------------------------------
# Docs cross-link — the policy doc and the drop-zone agree
# ---------------------------------------------------------------------------

def test_real_file_pilot_doc_documents_operator_pilot_dropzone():
    """``docs/real_file_pilot.md`` is the authoritative privacy
    doc for real-file pilots. Without an explicit cross-link to
    the operator_pilot drop-zone, the in-repo exception sits in
    silent contradiction with §1's "outside this repo" rule —
    which is exactly the bypass shape a stop-time reviewer would
    flag. The cross-link makes the exception visible and pins its
    narrow scope (learn-only, .gitignore-masked, four-file
    invariant)."""
    text = (REPO_ROOT / "docs" / "real_file_pilot.md").read_text(encoding="utf-8")
    assert "operator_pilot/README.zh-CN.md" in text, (
        "docs/real_file_pilot.md must cross-link to "
        "operator_pilot/README.zh-CN.md (e.g. in a §1a sub-section) "
        "so the in-repo drop-zone exception is documented next to "
        "the standard outside-the-repo rule it overrides."
    )


def test_real_file_pilot_doc_pins_learn_only_drop_zone_scope():
    """The §1a sub-section must explicitly mark the drop-zone as
    learn-only. Without that scoping language, a reader could
    reasonably extend the exception to confirm / render /
    validate-render, which the prompt and the test suite both
    deliberately forbid."""
    text = (REPO_ROOT / "docs" / "real_file_pilot.md").read_text(encoding="utf-8")
    lowered = text.lower()
    assert "learn-only" in lowered and "operator_pilot" in text, (
        "docs/real_file_pilot.md must describe the operator_pilot "
        "drop-zone as a learn-only exception so the scope is "
        "pinned in the authoritative privacy doc and not just in "
        "the operator-facing README."
    )
