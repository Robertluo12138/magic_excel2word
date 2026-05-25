"""Docs consistency check for ``docs/verification_matrix.md``.

The verification matrix maps each safety gate in
``docs/milestone_1_readiness.md`` §3 to the exact pytest file(s) that
pin it. To stay useful it must:

  1. exist on disk;
  2. be linked from ``docs/milestone_1_readiness.md`` so a reviewer
     arriving at the readiness snapshot can find it;
  3. reference only pytest files that actually exist under ``tests/``
     — a renamed or deleted test file referenced by the matrix would
     silently rot the gate-to-test mapping and the next reviewer
     would have no way to know.

This is a pure docs surface check — it imports nothing from
``src/``, never executes the CLI, never writes files, and never
touches synthetic or real data.
"""
from __future__ import annotations

import re
from pathlib import Path

import pytest


REPO_ROOT = Path(__file__).resolve().parents[1]
MATRIX_DOC = REPO_ROOT / "docs" / "verification_matrix.md"
READINESS_DOC = REPO_ROOT / "docs" / "milestone_1_readiness.md"
TESTS_DIR = REPO_ROOT / "tests"


def _referenced_test_files() -> set[str]:
    """Extract every ``tests/<file>.py`` reference from the matrix.

    Matches both bare ``tests/foo.py`` and ``tests/foo.py::test_bar``
    citations (the ``::test_bar`` suffix is dropped). Backticked,
    bolded, or bare references are all caught — the regex anchors on
    the literal ``tests/`` prefix and a trailing ``.py``.
    """
    text = MATRIX_DOC.read_text(encoding="utf-8")
    return {
        m.group(1)
        for m in re.finditer(r"tests/([A-Za-z0-9_]+\.py)", text)
    }


def test_verification_matrix_doc_exists():
    """milestone_1_readiness.md links to this file — the link target
    must resolve. If this fails, the gate-to-test mapping doc is
    gone and the readiness snapshot points at a broken target."""
    assert MATRIX_DOC.exists(), (
        "docs/verification_matrix.md must exist — it is the "
        "gate-to-test mapping that docs/milestone_1_readiness.md "
        "links to as the reverse lens on §3's safety gates."
    )


def test_milestone_readiness_doc_links_to_verification_matrix():
    """A reviewer reading the readiness snapshot must be able to
    discover the verification matrix without grepping the docs/
    tree. The link must use the file's relative path so it resolves
    from inside ``docs/``."""
    text = READINESS_DOC.read_text(encoding="utf-8")
    assert "verification_matrix.md" in text, (
        "docs/milestone_1_readiness.md must link to "
        "verification_matrix.md so a reviewer arriving at the "
        "readiness snapshot can pivot to the gate-to-test mapping."
    )


def test_verification_matrix_references_at_least_one_test_file():
    """A matrix with no ``tests/<file>.py`` references would silently
    pass the existence check below — pin the lower bound so an
    accidentally-blanked matrix fails this test instead."""
    referenced = _referenced_test_files()
    assert referenced, (
        "docs/verification_matrix.md must reference at least one "
        "pytest file under tests/ — the whole point of the matrix is "
        "to map gates to the tests that pin them."
    )


@pytest.mark.parametrize("test_file", sorted(_referenced_test_files()))
def test_every_referenced_test_file_exists(test_file: str):
    """Every ``tests/<file>.py`` cited by the matrix must exist on
    disk. A renamed or deleted test file referenced by the matrix
    would otherwise silently rot the gate-to-test mapping — the
    reviewer would read "pinned by tests/foo.py" and have no way to
    know the file is gone."""
    candidate = TESTS_DIR / test_file
    assert candidate.exists(), (
        f"docs/verification_matrix.md references tests/{test_file}, "
        f"but that file does not exist under tests/. Either restore "
        f"the test file, rename it back, or update the matrix to "
        f"point at the test that actually pins the gate."
    )
