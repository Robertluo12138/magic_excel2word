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
     forbidden-content rules before recording any aggregate value;
  4. cite every column header in ``REVIEW_HEADERS`` from
     ``src.mapping_reviewer`` — so a renamed or added
     mapping_review.xlsx column cannot silently rot the §1 schema
     bullet;
  5. cite every column header written by
     ``src.run_preview.write_run_validation`` — so a renamed or
     added run_validation.xlsx column cannot silently rot the §4
     schema bullet;
  6. cite every required field enforced by
     ``src.render_validator._check_render_log_required_fields`` —
     so a widened or narrowed ``render_log_missing_field`` gate
     cannot silently rot the §5 required-field bullet.

This is a pure docs surface check — the only ``src/`` reach-ins are
to import ``REVIEW_HEADERS`` and AST-parse two source files for the
canonical schema lists they pin. It never executes the CLI, never
writes files, and never touches synthetic or real data.
"""
from __future__ import annotations

import ast
from pathlib import Path

import pytest

from src.mapping_reviewer import REVIEW_HEADERS


REPO_ROOT = Path(__file__).resolve().parents[1]
GUIDE_DOC = REPO_ROOT / "docs" / "artifact_review_guide.md"
REAL_FILE_PILOT = REPO_ROOT / "docs" / "real_file_pilot.md"
PILOT_RESULT_TEMPLATE = REPO_ROOT / "docs" / "pilot_result_template.md"
RUN_PREVIEW_PY = REPO_ROOT / "src" / "run_preview.py"
RENDER_VALIDATOR_PY = REPO_ROOT / "src" / "render_validator.py"


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


# ---------------------------------------------------------------------------
# Schema-drift pins: keep §1, §4, and §5 of the guide in lockstep with the
# runtime's canonical column / field sets.
# ---------------------------------------------------------------------------

def _find_local_string_collection(
    source_path: Path,
    function_name: str,
    variable_name: str,
    node_type: type,
) -> list[str]:
    """AST-walk ``source_path`` and return the string literals assigned
    to ``variable_name`` inside ``function_name``. ``node_type`` is
    ``ast.List`` or ``ast.Tuple`` — whichever shape the source uses.

    Returns ``[]`` if the assignment cannot be located. The sentinel
    tests below pin a positive lower bound, so a refactor that
    reshapes the source (different function name, different variable
    name, dynamic construction) fails loudly via that sentinel — a
    silently-empty collection cannot let the per-header drift checks
    false-pass.
    """
    tree = ast.parse(source_path.read_text(encoding="utf-8"))
    for func in ast.walk(tree):
        if not isinstance(func, ast.FunctionDef) or func.name != function_name:
            continue
        for stmt in ast.walk(func):
            if not isinstance(stmt, ast.Assign):
                continue
            if len(stmt.targets) != 1 or not isinstance(stmt.targets[0], ast.Name):
                continue
            if stmt.targets[0].id != variable_name:
                continue
            if not isinstance(stmt.value, node_type):
                continue
            return [
                elt.value
                for elt in stmt.value.elts
                if isinstance(elt, ast.Constant) and isinstance(elt.value, str)
            ]
    return []


def _run_validation_headers_in_writer() -> list[str]:
    """Return the canonical run_validation.xlsx column headers, sourced
    from the literal ``headers`` list inside
    ``write_run_validation`` in ``src/run_preview.py``.

    AST-extraction keeps this test independent of file I/O — we never
    invoke ``write_run_validation`` and never write a workbook to disk.
    """
    return _find_local_string_collection(
        RUN_PREVIEW_PY,
        function_name="write_run_validation",
        variable_name="headers",
        node_type=ast.List,
    )


def _render_log_required_fields_in_validator() -> list[str]:
    """Return the render_log required-field names, sourced from the
    literal ``REQUIRED_FIELDS`` tuple inside
    ``_check_render_log_required_fields`` in
    ``src/render_validator.py``.

    AST-extraction keeps this test independent of file I/O — we never
    load a render log and never invoke the validator.
    """
    return _find_local_string_collection(
        RENDER_VALIDATOR_PY,
        function_name="_check_render_log_required_fields",
        variable_name="REQUIRED_FIELDS",
        node_type=ast.Tuple,
    )


def test_review_headers_is_non_empty():
    """``REVIEW_HEADERS`` is imported directly; a structurally empty
    schema would otherwise silently pass the per-header checks below.
    Pin a positive lower bound so an accidentally-blanked list fails
    here instead of going unnoticed."""
    assert REVIEW_HEADERS, (
        "src.mapping_reviewer.REVIEW_HEADERS is empty — the matcher "
        "would write mapping_review.xlsx without column headers. The "
        "guide checks below would false-pass against an empty list."
    )


def test_run_validation_header_extraction_is_non_empty():
    """The AST extractor returning an empty list would let the guide
    silently shrink alongside ``write_run_validation``. Pin a positive
    lower bound so a refactor that moves, renames, or empties the
    list fails here instead of going unnoticed."""
    assert _run_validation_headers_in_writer(), (
        "AST extraction returned no headers from "
        "src.run_preview.write_run_validation — the function may "
        "have been refactored (renamed, moved, or rebuilt with a "
        "different local variable). Update "
        "_find_local_string_collection's arguments to match the new "
        "shape; the guide checks below would false-pass against an "
        "empty list."
    )


def test_render_log_required_field_extraction_is_non_empty():
    """The AST extractor returning an empty tuple would let the guide
    silently shrink alongside the render-log gate. Pin a positive
    lower bound so a refactor that moves, renames, or empties the
    tuple fails here instead of going unnoticed."""
    assert _render_log_required_fields_in_validator(), (
        "AST extraction returned no REQUIRED_FIELDS from "
        "src.render_validator._check_render_log_required_fields — "
        "the function may have been refactored. Update "
        "_find_local_string_collection's arguments to match the new "
        "shape; the guide checks below would false-pass against an "
        "empty tuple."
    )


@pytest.mark.parametrize("header", REVIEW_HEADERS)
def test_artifact_review_guide_cites_every_mapping_review_header(header: str):
    """The guide's §1 schema bullet must name every column in
    ``REVIEW_HEADERS`` in backticks so a reviewer cross-checking the
    mapping_review.xlsx schema by eye sees the same set the writer
    produces. If a header is renamed or added in
    ``src/mapping_reviewer.py`` without updating the guide, the doc
    would silently mislead the reviewer about what columns to expect.
    """
    text = GUIDE_DOC.read_text(encoding="utf-8")
    assert f"`{header}`" in text, (
        f"docs/artifact_review_guide.md does not mention the "
        f"mapping_review.xlsx column `{header}` from "
        f"src.mapping_reviewer.REVIEW_HEADERS — keep the §1 schema "
        f"bullet in lockstep with REVIEW_HEADERS."
    )


@pytest.mark.parametrize("header", _run_validation_headers_in_writer())
def test_artifact_review_guide_cites_every_run_validation_header(header: str):
    """The guide's §4 schema bullet must name every column written by
    ``write_run_validation`` in backticks so a reviewer cross-checking
    the run_validation.xlsx schema by eye sees the same set the writer
    produces. If a header is renamed or added in
    ``src/run_preview.py`` without updating the guide, the doc would
    silently mislead the reviewer about what columns to expect.
    """
    text = GUIDE_DOC.read_text(encoding="utf-8")
    assert f"`{header}`" in text, (
        f"docs/artifact_review_guide.md does not mention the "
        f"run_validation.xlsx column `{header}` from "
        f"src.run_preview.write_run_validation — keep the §4 schema "
        f"bullet in lockstep with the writer's headers list."
    )


@pytest.mark.parametrize(
    "field_name", _render_log_required_fields_in_validator(),
)
def test_artifact_review_guide_cites_every_render_log_required_field(
    field_name: str,
):
    """The guide's §5 required-field bullet must name every field in
    the ``REQUIRED_FIELDS`` tuple inside
    ``_check_render_log_required_fields`` in backticks — those
    fields are the exact contract that produces a
    ``render_log_missing_field`` issue. If the gate is widened or
    narrowed in ``src.render_validator`` without updating the guide,
    the doc would silently mislead the reviewer about what the gate
    catches.
    """
    text = GUIDE_DOC.read_text(encoding="utf-8")
    assert f"`{field_name}`" in text, (
        f"docs/artifact_review_guide.md does not mention the "
        f"render_log required field `{field_name}` from "
        f"src.render_validator._check_render_log_required_fields — "
        f"keep the §5 required-field bullet in lockstep with the "
        f"gate's REQUIRED_FIELDS tuple."
    )
