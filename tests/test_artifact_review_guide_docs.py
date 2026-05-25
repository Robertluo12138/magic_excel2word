"""Docs consistency check for ``docs/artifact_review_guide.md``.

The artifact review guide tells a reviewer exactly what to inspect
in each pilot artifact using aggregate / status / schema checks
only, and forbids copying raw business numbers, raw Word snippets,
Excel cell values, source sheet/cell content, or company identifiers
into the repo or any doc. To stay useful it must:

  1. exist on disk;
  2. be linked from ``docs/real_file_pilot.md`` so a reviewer
     arriving at the manual real-file workflow can pivot to the
     per-artifact checklist without grepping the ``docs/`` tree;
  3. be linked from ``docs/pilot_result_template.md`` so a reviewer
     filling in the redacted result form can re-read the
     forbidden-content rules before recording any aggregate value.

This is a pure docs surface check — it imports nothing from
``src/``, never executes the CLI, never writes files, and never
touches synthetic or real data.
"""
from __future__ import annotations

from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
GUIDE_DOC = REPO_ROOT / "docs" / "artifact_review_guide.md"
REAL_FILE_PILOT = REPO_ROOT / "docs" / "real_file_pilot.md"
PILOT_RESULT_TEMPLATE = REPO_ROOT / "docs" / "pilot_result_template.md"


def test_artifact_review_guide_doc_exists():
    """real_file_pilot.md and pilot_result_template.md both link to
    this file — the link targets must resolve. If this fails, both
    entry points point at a broken target and a reviewer has no
    per-artifact inspection checklist to walk."""
    assert GUIDE_DOC.exists(), (
        "docs/artifact_review_guide.md must exist — it is the "
        "per-artifact inspection checklist that "
        "docs/real_file_pilot.md and docs/pilot_result_template.md "
        "both link to as the aggregate/status/schema review "
        "contract for the manual real-file pilot."
    )


def test_real_file_pilot_doc_links_to_artifact_review_guide():
    """A reviewer arriving at the manual real-file workflow must be
    able to discover the per-artifact checklist without grepping the
    ``docs/`` tree. The link must use the file's relative path so it
    resolves from inside ``docs/``."""
    text = REAL_FILE_PILOT.read_text(encoding="utf-8")
    assert "artifact_review_guide.md" in text, (
        "docs/real_file_pilot.md must link to "
        "artifact_review_guide.md so a reviewer arriving at the "
        "manual real-file workflow can pivot to the per-artifact "
        "inspection checklist."
    )


def test_pilot_result_template_doc_links_to_artifact_review_guide():
    """A reviewer filling in the redacted result form must be able
    to re-read the forbidden-content rules before recording any
    aggregate value. The link must use the file's relative path so
    it resolves from inside ``docs/``."""
    text = PILOT_RESULT_TEMPLATE.read_text(encoding="utf-8")
    assert "artifact_review_guide.md" in text, (
        "docs/pilot_result_template.md must link to "
        "artifact_review_guide.md so a reviewer filling in the "
        "redacted result form can re-read the forbidden-content "
        "rules before recording any aggregate value."
    )
