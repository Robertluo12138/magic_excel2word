"""Docs consistency check for ``docs/pilot_result_template.md``.

The pilot-result template is the redacted form a reviewer fills in
after walking the manual real-file workflow in
``docs/real_file_pilot.md``. To stay useful it must:

  1. exist on disk;
  2. be linked from ``docs/real_file_pilot.md`` so the reviewer
     completing the workflow can find it without grepping the
     ``docs/`` tree;
  3. be linked from ``docs/milestone_1_readiness.md`` so a sponsor
     reading the readiness snapshot knows where the recorded
     evidence from a pilot lives.

This is a pure docs surface check — it imports nothing from
``src/``, never executes the CLI, never writes files, and never
touches synthetic or real data.
"""
from __future__ import annotations

from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
TEMPLATE_DOC = REPO_ROOT / "docs" / "pilot_result_template.md"
REAL_FILE_PILOT = REPO_ROOT / "docs" / "real_file_pilot.md"
READINESS_DOC = REPO_ROOT / "docs" / "milestone_1_readiness.md"


def test_pilot_result_template_doc_exists():
    """real_file_pilot.md and milestone_1_readiness.md both link to
    this file — the link targets must resolve. If this fails, both
    entry points point at a broken target and a reviewer has no
    structured form to record pilot evidence."""
    assert TEMPLATE_DOC.exists(), (
        "docs/pilot_result_template.md must exist — it is the "
        "redacted result form a reviewer fills in after walking "
        "the manual real-file workflow in docs/real_file_pilot.md."
    )


def test_real_file_pilot_doc_links_to_pilot_result_template():
    """A reviewer completing the manual workflow must be able to
    discover the result template without grepping the ``docs/``
    tree. The link must use the file's relative path so it
    resolves from inside ``docs/``."""
    text = REAL_FILE_PILOT.read_text(encoding="utf-8")
    assert "pilot_result_template.md" in text, (
        "docs/real_file_pilot.md must link to "
        "pilot_result_template.md so the reviewer completing the "
        "manual workflow can find the redacted result form."
    )


def test_milestone_readiness_doc_links_to_pilot_result_template():
    """A sponsor reading the readiness snapshot must be able to
    discover where the recorded evidence from a pilot lives,
    without context-switching back to ``real_file_pilot.md``."""
    text = READINESS_DOC.read_text(encoding="utf-8")
    assert "pilot_result_template.md" in text, (
        "docs/milestone_1_readiness.md must link to "
        "pilot_result_template.md so a sponsor reading the "
        "readiness snapshot knows where the recorded evidence "
        "from a pilot lives."
    )
