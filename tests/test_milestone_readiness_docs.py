"""Docs consistency check for ``docs/milestone_1_readiness.md``.

The reviewer-facing milestone readiness doc is the snapshot a reviewer
reads before sponsoring a real-file pilot. To stay useful it must:

  1. exist on disk;
  2. be linked from ``README.md`` and ``docs/real_file_pilot.md`` so a
     reviewer can find it from either entry point;
  3. explicitly declare prototype status — the entire reason the doc
     exists is to bound expectations, so a missing prototype banner
     would silently flip the contract;
  4. **not** contain any unambiguous positive readiness claim such as
     ``production-grade``, ``battle-tested``, ``enterprise-ready``,
     ``GA release``. These phrases never appear in a disclaimer and
     would mis-signal sponsor sign-off if they slipped in.

This is a pure docs surface check — it imports nothing from
``src/``, never executes the CLI, never writes files, and never
touches synthetic or real data.
"""
from __future__ import annotations

from pathlib import Path

import pytest


REPO_ROOT = Path(__file__).resolve().parents[1]
README = REPO_ROOT / "README.md"
REAL_FILE_PILOT = REPO_ROOT / "docs" / "real_file_pilot.md"
READINESS_DOC = REPO_ROOT / "docs" / "milestone_1_readiness.md"


def test_milestone_readiness_doc_exists():
    """README.md and docs/real_file_pilot.md both link to this file —
    the link targets must resolve. If this fails, the reviewer-facing
    snapshot is gone and both entry points point at a broken target."""
    assert READINESS_DOC.exists(), (
        "docs/milestone_1_readiness.md must exist — it is the "
        "reviewer-facing snapshot of the prototype's gates, tests, "
        "and known limits that README.md and docs/real_file_pilot.md "
        "both link to."
    )


def test_readme_links_to_milestone_readiness_doc():
    """A reviewer arriving from the repo root must be able to discover
    the readiness snapshot without grepping the docs/ tree."""
    text = README.read_text(encoding="utf-8")
    assert "docs/milestone_1_readiness.md" in text, (
        "README.md must link to docs/milestone_1_readiness.md so a "
        "reviewer arriving from the repo root can find the snapshot."
    )


def test_real_file_pilot_doc_links_to_milestone_readiness_doc():
    """A reviewer arriving from the pilot doc must be able to jump to
    the snapshot before walking the manual workflow."""
    text = REAL_FILE_PILOT.read_text(encoding="utf-8")
    assert "milestone_1_readiness.md" in text, (
        "docs/real_file_pilot.md must link to milestone_1_readiness.md "
        "so a reviewer arriving from the pilot doc can find the "
        "snapshot before walking the manual workflow."
    )


def test_milestone_readiness_doc_declares_prototype_status():
    """The doc's reason for existing is to bound reviewer expectations.
    A missing ``prototype`` declaration would silently flip the
    contract from "snapshot of a prototype" to "ready-to-ship summary"."""
    text = READINESS_DOC.read_text(encoding="utf-8").lower()
    assert "prototype" in text, (
        "docs/milestone_1_readiness.md must explicitly declare "
        "prototype status — without it the doc reads as a "
        "ready-to-ship summary instead of a reviewer-facing snapshot."
    )


# Phrases that never appear in a disclaimer — listing any of them would
# unambiguously signal production readiness. Deliberately narrow so the
# test does not false-positive on legitimate disclaimer wording (e.g.
# "not production-ready yet" or "not ready for production"), which is
# why broader strings like "production-ready" alone are not blocked.
PROHIBITED_POSITIVE_CLAIMS = (
    "production-grade",
    "production grade",
    "battle-tested",
    "battle tested",
    "enterprise-ready",
    "enterprise ready",
    "GA release",
    "GA-release",
)


@pytest.mark.parametrize("phrase", PROHIBITED_POSITIVE_CLAIMS)
def test_milestone_readiness_doc_makes_no_positive_readiness_claim(phrase: str):
    """No unambiguous positive readiness phrase may appear in the doc.
    These are descriptors that only show up in marketing-style claims,
    never in a "this is still a prototype" disclaimer."""
    text = READINESS_DOC.read_text(encoding="utf-8").lower()
    assert phrase.lower() not in text, (
        f"docs/milestone_1_readiness.md must not contain the positive "
        f"readiness claim {phrase!r}. The doc bounds reviewer "
        f"expectations; any such phrase would mis-signal that the "
        f"prototype is sponsor-ready for unattended real-file use."
    )
