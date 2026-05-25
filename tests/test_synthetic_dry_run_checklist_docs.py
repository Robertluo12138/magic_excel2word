"""Docs consistency check for ``docs/synthetic_dry_run_checklist.md``.

The synthetic dry-run checklist is the rehearsal an operator walks
before any real-file pilot. To stay useful it must:

  1. exist on disk;
  2. be linked from ``docs/real_file_pilot.md`` so an operator
     arriving at the real-file workflow can pivot back to the
     rehearsal before pointing the pipeline at real data;
  3. be linked from ``docs/milestone_1_readiness.md`` so a reviewer
     can confirm the rehearsal precondition is part of the
     readiness checklist.

This is a pure docs surface check — it imports nothing from
``src/``, never executes the CLI, never writes files, and never
touches synthetic or real data.
"""
from __future__ import annotations

from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
CHECKLIST_DOC = REPO_ROOT / "docs" / "synthetic_dry_run_checklist.md"
REAL_FILE_PILOT = REPO_ROOT / "docs" / "real_file_pilot.md"
READINESS_DOC = REPO_ROOT / "docs" / "milestone_1_readiness.md"


def test_synthetic_dry_run_checklist_doc_exists():
    """real_file_pilot.md and milestone_1_readiness.md both link to
    this file — the link targets must resolve. If this fails, both
    entry points point at a broken target and the rehearsal
    precondition has no doc behind it."""
    assert CHECKLIST_DOC.exists(), (
        "docs/synthetic_dry_run_checklist.md must exist — it is the "
        "operator rehearsal that docs/real_file_pilot.md and "
        "docs/milestone_1_readiness.md both link to as the "
        "precondition for any real-file pilot."
    )


def test_real_file_pilot_doc_links_to_synthetic_dry_run_checklist():
    """An operator arriving at the real-file workflow must be able
    to discover the rehearsal checklist without grepping the docs/
    tree. The link uses the file's relative path so it resolves
    from inside ``docs/``."""
    text = REAL_FILE_PILOT.read_text(encoding="utf-8")
    assert "synthetic_dry_run_checklist.md" in text, (
        "docs/real_file_pilot.md must link to "
        "synthetic_dry_run_checklist.md so an operator arriving at "
        "the real-file workflow can pivot back to the rehearsal "
        "before pointing the pipeline at real data."
    )


def test_milestone_readiness_doc_links_to_synthetic_dry_run_checklist():
    """A reviewer reading the readiness snapshot must see the
    rehearsal listed as part of the pre-pilot precondition set."""
    text = READINESS_DOC.read_text(encoding="utf-8")
    assert "synthetic_dry_run_checklist.md" in text, (
        "docs/milestone_1_readiness.md must link to "
        "synthetic_dry_run_checklist.md so a reviewer can confirm "
        "the rehearsal precondition is part of the readiness "
        "checklist before sponsoring a real-file pilot."
    )
