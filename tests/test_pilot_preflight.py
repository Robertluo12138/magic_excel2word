"""Focused tests for the read-only ``pilot-preflight`` command.

The command is the explicit pre-pilot path/metadata gate: it checks
that the four pilot paths exist with the right suffix, ``--out`` is
a directory (or can be created), and **no** input/output path lands
inside this repo tree. It never opens any document and prints only
flag labels and basenames.

Each test below pins one slice of that contract:

  * outside-repo paths with valid suffixes succeed and the command
    exits ``0`` without writing anything;
  * any path resolving inside the repo tree (``src/``, ``samples/``,
    ``tests/``, ``docs/``, the repo root) is refused with exit ``12``
    and a REFUSED banner pointing at ``docs/real_file_pilot.md`` §1;
  * missing input files exit ``2`` with a ``missing`` per-path tag;
  * wrong-suffix inputs exit ``2`` with a ``bad-suffix`` per-path tag;
  * the redacted-output contract: full paths and parent directories
    never reach stdout/stderr, only flag labels and basenames do;
  * documented exit-code parity: the runtime emits exactly the codes
    the docs claim, and the docs list every code the runtime emits.

The fixture surface is intentionally minimal: ``tmp_path``-backed
empty ``.xlsx`` / ``.docx`` files (or ``.write_bytes(b"")``) — the
command never opens them, so the byte content is irrelevant.
"""
from __future__ import annotations

import ast
import re
from pathlib import Path
from typing import Tuple

import pytest

from src.main import main as cli_main
from src.pilot_preflight import (
    PreflightReport,
    format_report,
    preflight,
)


REPO_ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = REPO_ROOT / "docs"
COMMAND_REFERENCE = DOCS_DIR / "command_reference.md"
REAL_FILE_PILOT = DOCS_DIR / "real_file_pilot.md"
MAIN_PY = REPO_ROOT / "src" / "main.py"


# ---------------------------------------------------------------------------
# Tiny synthetic-file helper. The command never opens any document, so the
# bytes are irrelevant — we just need the path to exist with the right
# suffix (or wrong suffix, in the bad-suffix tests).
# ---------------------------------------------------------------------------

def _touch(path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(b"")
    return path


def _outside_repo_pilot(tmp_path: Path) -> Tuple[Path, Path, Path, Path]:
    """Return four valid outside-the-repo pilot paths under ``tmp_path``.

    pytest's ``tmp_path`` lives under the OS tempdir, which is OUTSIDE
    the repo tree by construction, so the inside-repo gate stays clean.
    """
    inputs = tmp_path / "pilot_inputs"
    out = tmp_path / "pilot_output"
    hist_xlsx = _touch(inputs / "historical.xlsx")
    hist_docx = _touch(inputs / "finished_report.docx")
    new_xlsx = _touch(inputs / "new_period.xlsx")
    return hist_xlsx, hist_docx, new_xlsx, out


# ---------------------------------------------------------------------------
# Module-level: outside-repo success
# ---------------------------------------------------------------------------

def test_preflight_outside_repo_all_ok_no_existing_out(tmp_path: Path):
    """All four checks pass; ``--out`` is the will-be-created flavor."""
    hist_xlsx, hist_docx, new_xlsx, out = _outside_repo_pilot(tmp_path)
    assert not out.exists()
    report = preflight(hist_xlsx, hist_docx, new_xlsx, out)
    assert report.ok is True
    assert report.inside_repo_count == 0
    assert report.other_failure_count == 0
    statuses = [c.status for c in report.checks]
    assert statuses == ["ok", "ok", "ok", "will-be-created"]


def test_preflight_outside_repo_with_existing_out_dir(tmp_path: Path):
    """Pre-existing ``--out`` directory passes too."""
    hist_xlsx, hist_docx, new_xlsx, out = _outside_repo_pilot(tmp_path)
    out.mkdir()
    report = preflight(hist_xlsx, hist_docx, new_xlsx, out)
    assert report.ok is True
    assert report.checks[-1].status == "ok"


# ---------------------------------------------------------------------------
# CLI: outside-repo success
# ---------------------------------------------------------------------------

def test_cli_outside_repo_returns_0(tmp_path: Path, capsys):
    """The CLI surface returns 0 and emits the redacted summary on stdout."""
    hist_xlsx, hist_docx, new_xlsx, out = _outside_repo_pilot(tmp_path)
    rc = cli_main([
        "pilot-preflight",
        "--historical-excel", str(hist_xlsx),
        "--historical-word", str(hist_docx),
        "--new-excel", str(new_xlsx),
        "--out", str(out),
    ])
    captured = capsys.readouterr()
    assert rc == 0
    assert "pilot-preflight checks:" in captured.out
    assert "All checks passed" in captured.out
    # Nothing on stderr on the clean path.
    assert captured.err == ""


# ---------------------------------------------------------------------------
# Inside-repo refusal: every common inside-repo location must trip the gate
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    "inside_subdir",
    ["samples", "src", "tests", "docs"],
)
def test_preflight_refuses_inside_repo_input(
    tmp_path: Path, inside_subdir: str
):
    """An input path resolving anywhere inside the repo tree is refused,
    not just under ``samples/``. The privacy gate is strictly broader
    than the existing ``preflight.py`` advisory."""
    # Use a name that does not actually exist on disk under the repo —
    # the inside-repo gate must fire BEFORE the existence check, the
    # same ordering rationale that ``test_preflight.py`` pins for the
    # informational advisory.
    bogus = REPO_ROOT / inside_subdir / "__pilot_preflight_does_not_exist__.xlsx"
    assert not bogus.exists()

    # Build the other three paths outside the repo so we isolate the
    # one inside-repo failure.
    _, hist_docx, new_xlsx, out = _outside_repo_pilot(tmp_path)
    report = preflight(bogus, hist_docx, new_xlsx, out)
    assert report.ok is False
    assert report.inside_repo_count == 1
    assert report.other_failure_count == 0
    assert report.checks[0].status == "inside-repo"
    # And the basename is preserved — only the parent tree is hidden.
    assert report.checks[0].basename == bogus.name


def test_preflight_refuses_inside_repo_out_directory(tmp_path: Path):
    """``--out`` resolving inside the repo tree (here: ``output/`` at
    the repo root) is refused. This mirrors the inside-repo input
    refusal but on the directory side, and it must fire even when the
    directory does not yet exist — operators typo'ing ``output`` from
    inside the repo would otherwise pass the will-be-created branch."""
    hist_xlsx, hist_docx, new_xlsx, _ = _outside_repo_pilot(tmp_path)
    bogus_out = REPO_ROOT / "__pilot_preflight_out_does_not_exist__"
    assert not bogus_out.exists()
    report = preflight(hist_xlsx, hist_docx, new_xlsx, bogus_out)
    assert report.ok is False
    assert report.inside_repo_count == 1
    assert report.checks[-1].status == "inside-repo"
    assert report.checks[-1].basename == bogus_out.name


def test_cli_inside_repo_returns_12_with_refused_banner(
    tmp_path: Path, capsys
):
    """The CLI translates the privacy refusal to exit ``12`` and prints
    the redacted summary plus a REFUSED banner on stderr."""
    _, hist_docx, new_xlsx, out = _outside_repo_pilot(tmp_path)
    bogus_inside_repo = REPO_ROOT / "samples" / "__pp_inside_repo_test__.xlsx"
    rc = cli_main([
        "pilot-preflight",
        "--historical-excel", str(bogus_inside_repo),
        "--historical-word", str(hist_docx),
        "--new-excel", str(new_xlsx),
        "--out", str(out),
    ])
    captured = capsys.readouterr()
    assert rc == 12
    # Stderr carries the summary (failure branch) and the banner.
    assert "pilot-preflight checks:" in captured.err
    assert "REFUSED" in captured.err
    assert "OUTSIDE this repo" in captured.err
    assert "docs/real_file_pilot.md" in captured.err
    # Stdout stays clean on the failure branch — automation that only
    # captured stdout would not see a misleading "all clear".
    assert captured.out == ""


def test_cli_inside_repo_wins_over_other_failures(tmp_path: Path, capsys):
    """When BOTH inside-repo and other per-path failures are present, the
    privacy refusal must win: exit ``12`` (not ``2``), so automation
    branching on the integer never masks an inside-repo path."""
    _, hist_docx, _, out = _outside_repo_pilot(tmp_path)
    inside_repo_xlsx = REPO_ROOT / "samples" / "__pp_inside_repo_and_other__.xlsx"
    missing_xlsx = tmp_path / "this_file_does_not_exist.xlsx"
    assert not missing_xlsx.exists()
    rc = cli_main([
        "pilot-preflight",
        "--historical-excel", str(inside_repo_xlsx),
        "--historical-word", str(hist_docx),
        "--new-excel", str(missing_xlsx),
        "--out", str(out),
    ])
    captured = capsys.readouterr()
    assert rc == 12, "inside-repo (12) must outrank other path failures (2)"
    assert "REFUSED" in captured.err
    assert "FAILED" in captured.err  # the other failure still surfaces in the summary


# ---------------------------------------------------------------------------
# Missing files
# ---------------------------------------------------------------------------

def test_cli_missing_input_returns_2(tmp_path: Path, capsys):
    _, hist_docx, new_xlsx, out = _outside_repo_pilot(tmp_path)
    missing = tmp_path / "pilot_inputs" / "absent_historical.xlsx"
    assert not missing.exists()
    rc = cli_main([
        "pilot-preflight",
        "--historical-excel", str(missing),
        "--historical-word", str(hist_docx),
        "--new-excel", str(new_xlsx),
        "--out", str(out),
    ])
    captured = capsys.readouterr()
    assert rc == 2
    assert "missing" in captured.err
    assert "file not found" in captured.err
    # And the basename of the offending input is surfaced — operator
    # can identify which flag tripped without reading the full path.
    assert missing.name in captured.err


def test_cli_all_three_inputs_missing_returns_2(tmp_path: Path, capsys):
    """Multiple missing inputs still resolve to a single ``2`` — the
    command doesn't bail on the first failure, it lists all of them so
    a reviewer fixes them in one pass."""
    _, _, _, out = _outside_repo_pilot(tmp_path)
    pilot_inputs = tmp_path / "pilot_inputs"
    pilot_inputs.mkdir(exist_ok=True)
    a, b, c = (
        pilot_inputs / "a.xlsx",
        pilot_inputs / "b.docx",
        pilot_inputs / "c.xlsx",
    )
    rc = cli_main([
        "pilot-preflight",
        "--historical-excel", str(a),
        "--historical-word", str(b),
        "--new-excel", str(c),
        "--out", str(out),
    ])
    captured = capsys.readouterr()
    assert rc == 2
    # All three offending basenames surface so the operator sees them
    # in one pass instead of fix-rerun-fix-rerun.
    assert a.name in captured.err
    assert b.name in captured.err
    assert c.name in captured.err


# ---------------------------------------------------------------------------
# Suffix checks
# ---------------------------------------------------------------------------

def test_cli_wrong_suffix_returns_2(tmp_path: Path, capsys):
    """An existing file with the wrong suffix is rejected — the command
    is meant as a safety net for typos like ``new_period.xls``."""
    inputs = tmp_path / "pilot_inputs"
    hist_xlsx = _touch(inputs / "historical.xls")     # wrong: .xls not .xlsx
    hist_docx = _touch(inputs / "finished_report.docx")
    new_xlsx = _touch(inputs / "new_period.xlsx")
    out = tmp_path / "pilot_output"
    rc = cli_main([
        "pilot-preflight",
        "--historical-excel", str(hist_xlsx),
        "--historical-word", str(hist_docx),
        "--new-excel", str(new_xlsx),
        "--out", str(out),
    ])
    captured = capsys.readouterr()
    assert rc == 2
    assert "bad-suffix" in captured.err
    assert "expected .xlsx" in captured.err


def test_cli_word_with_xlsx_suffix_returns_2(tmp_path: Path, capsys):
    """Mirror: a ``.xlsx`` swapped into ``--historical-word`` is caught
    as a bad-suffix failure, not silently accepted."""
    inputs = tmp_path / "pilot_inputs"
    hist_xlsx = _touch(inputs / "historical.xlsx")
    hist_docx = _touch(inputs / "finished_report.xlsx")  # wrong: .xlsx not .docx
    new_xlsx = _touch(inputs / "new_period.xlsx")
    out = tmp_path / "pilot_output"
    rc = cli_main([
        "pilot-preflight",
        "--historical-excel", str(hist_xlsx),
        "--historical-word", str(hist_docx),
        "--new-excel", str(new_xlsx),
        "--out", str(out),
    ])
    captured = capsys.readouterr()
    assert rc == 2
    assert "bad-suffix" in captured.err
    assert "expected .docx" in captured.err


def test_cli_uppercase_suffix_is_accepted(tmp_path: Path, capsys):
    """Operator-quality-of-life: ``HISTORICAL.XLSX`` on a case-sensitive
    filesystem must still pass — the suffix check is case-insensitive
    like everyone else in this repo (sheet names, status enums, etc.)."""
    inputs = tmp_path / "pilot_inputs"
    hist_xlsx = _touch(inputs / "HISTORICAL.XLSX")
    hist_docx = _touch(inputs / "FINISHED_REPORT.DOCX")
    new_xlsx = _touch(inputs / "NEW_PERIOD.xlsx")
    out = tmp_path / "pilot_output"
    rc = cli_main([
        "pilot-preflight",
        "--historical-excel", str(hist_xlsx),
        "--historical-word", str(hist_docx),
        "--new-excel", str(new_xlsx),
        "--out", str(out),
    ])
    captured = capsys.readouterr()
    assert rc == 0, captured.err
    assert "All checks passed" in captured.out


# ---------------------------------------------------------------------------
# --out is a regular file (not a directory)
# ---------------------------------------------------------------------------

def test_cli_out_is_a_file_returns_2(tmp_path: Path, capsys):
    """``--out`` pointing at an existing regular file is rejected with
    ``not-a-dir`` — distinct from the will-be-created branch (where the
    path doesn't exist yet)."""
    hist_xlsx, hist_docx, new_xlsx, _ = _outside_repo_pilot(tmp_path)
    out_file = tmp_path / "actually_a_file.txt"
    out_file.write_text("not a directory")
    rc = cli_main([
        "pilot-preflight",
        "--historical-excel", str(hist_xlsx),
        "--historical-word", str(hist_docx),
        "--new-excel", str(new_xlsx),
        "--out", str(out_file),
    ])
    captured = capsys.readouterr()
    assert rc == 2
    assert "not-a-dir" in captured.err


def test_cli_out_cannot_be_created_returns_2(tmp_path: Path, capsys):
    """``--out`` under a regular-file ancestor is rejected with
    ``cannot-create`` — a later ``mkdir(parents=True)`` would fail, so
    the preflight refuses up front instead of letting the downstream
    stage crash mid-pilot.

    Pinned shape: ``<tmp_path>/some_file.txt/output``. The "parent" is
    a regular file, so neither ``output`` itself nor any layer below
    can be materialised.
    """
    hist_xlsx, hist_docx, new_xlsx, _ = _outside_repo_pilot(tmp_path)
    file_blocker = tmp_path / "blocker.txt"
    file_blocker.write_text("regular file, not a directory")
    out_under_file = file_blocker / "output"
    rc = cli_main([
        "pilot-preflight",
        "--historical-excel", str(hist_xlsx),
        "--historical-word", str(hist_docx),
        "--new-excel", str(new_xlsx),
        "--out", str(out_under_file),
    ])
    captured = capsys.readouterr()
    assert rc == 2
    assert "cannot-create" in captured.err
    assert "nearest existing ancestor is not a directory" in captured.err


def test_cli_out_deeper_under_file_blocker_returns_2(tmp_path: Path, capsys):
    """Same failure mode, two layers deeper: ``<tmp>/blocker.txt/a/b/output``.

    The walk-up must reach ``blocker.txt`` and refuse — a regression
    that only checked ``path.parent`` (instead of walking up to the
    nearest existing ancestor) would false-pass this case.
    """
    hist_xlsx, hist_docx, new_xlsx, _ = _outside_repo_pilot(tmp_path)
    file_blocker = tmp_path / "blocker.txt"
    file_blocker.write_text("regular file two layers up")
    out_under_file = file_blocker / "a" / "b" / "output"
    rc = cli_main([
        "pilot-preflight",
        "--historical-excel", str(hist_xlsx),
        "--historical-word", str(hist_docx),
        "--new-excel", str(new_xlsx),
        "--out", str(out_under_file),
    ])
    captured = capsys.readouterr()
    assert rc == 2
    assert "cannot-create" in captured.err


def test_preflight_will_be_created_when_intermediate_dirs_missing(tmp_path: Path):
    """Mirror: when the ancestor chain is clean (every existing layer
    is a directory) the path is creatable — even when the parent
    itself does not yet exist. ``mkdir(parents=True)`` would succeed
    so the preflight returns ``will-be-created``."""
    hist_xlsx, hist_docx, new_xlsx, _ = _outside_repo_pilot(tmp_path)
    deep_out = tmp_path / "does" / "not" / "exist" / "yet" / "output"
    assert not deep_out.exists()
    report = preflight(hist_xlsx, hist_docx, new_xlsx, deep_out)
    assert report.ok is True
    assert report.checks[-1].status == "will-be-created"


# ---------------------------------------------------------------------------
# Redaction contract — output prints labels + basenames only
# ---------------------------------------------------------------------------

def test_format_report_redacts_full_paths_and_parent_directories(tmp_path: Path):
    """Pin the load-bearing redaction claim: the formatter only prints
    flag labels and basenames. The parent tree, the tmp_path root, and
    the full file path must never leak.

    We use a deliberately-sensitive parent tree name so a regression
    that ``str(path)``'d the input would be unmistakable in the
    failure message."""
    sensitive = "customer_alpha_2026Q2_top_secret"
    inputs = tmp_path / sensitive / "pilot_inputs"
    hist_xlsx = _touch(inputs / "real_data_historical.xlsx")
    hist_docx = _touch(inputs / "real_data_finished_report.docx")
    new_xlsx = _touch(inputs / "real_data_new_period.xlsx")
    out = tmp_path / sensitive / "pilot_output"

    report = preflight(hist_xlsx, hist_docx, new_xlsx, out)
    text = format_report(report)

    # Basenames must appear so the operator can match each line to a
    # flag without reading their own filesystem.
    for name in (
        hist_xlsx.name, hist_docx.name, new_xlsx.name, out.name,
    ):
        assert name in text

    # Parent tree, sensitive folder, full absolute path — must not.
    for forbidden in (
        sensitive,
        str(tmp_path),
        str(inputs),
        str(hist_xlsx),
        str(hist_docx),
        str(new_xlsx),
        str(out),
        str(out.parent),
    ):
        assert forbidden not in text, (
            f"redaction failed: pilot-preflight printed {forbidden!r}"
        )


def test_cli_output_is_redacted_on_failure_branch(tmp_path: Path, capsys):
    """The redaction also applies on the failure branch where the report
    is sent to stderr (not stdout). A regression that ``print(args.out)``'d
    on error would otherwise slip past the success-path redaction test."""
    sensitive = "customer_beta_2026Q3_top_secret"
    inputs = tmp_path / sensitive / "pilot_inputs"
    hist_xlsx = _touch(inputs / "h.xlsx")
    hist_docx = _touch(inputs / "h.docx")
    new_xlsx = _touch(inputs / "n.xlsx")
    out = REPO_ROOT / "samples" / "__pp_redact_failure_branch__"
    rc = cli_main([
        "pilot-preflight",
        "--historical-excel", str(hist_xlsx),
        "--historical-word", str(hist_docx),
        "--new-excel", str(new_xlsx),
        "--out", str(out),
    ])
    captured = capsys.readouterr()
    assert rc == 12
    full_streams = captured.out + captured.err
    assert sensitive not in full_streams
    assert str(tmp_path) not in full_streams
    assert str(inputs) not in full_streams


def test_cli_output_never_includes_full_input_path(tmp_path: Path, capsys):
    """Tighter pin: even for a fully-passing run, the full absolute
    paths of the four inputs are never echoed back to the operator."""
    hist_xlsx, hist_docx, new_xlsx, out = _outside_repo_pilot(tmp_path)
    rc = cli_main([
        "pilot-preflight",
        "--historical-excel", str(hist_xlsx),
        "--historical-word", str(hist_docx),
        "--new-excel", str(new_xlsx),
        "--out", str(out),
    ])
    captured = capsys.readouterr()
    assert rc == 0
    combined = captured.out + captured.err
    for forbidden in (
        str(hist_xlsx), str(hist_docx), str(new_xlsx), str(out),
        str(tmp_path),
    ):
        assert forbidden not in combined


# ---------------------------------------------------------------------------
# Read-only contract — the command never writes anything
# ---------------------------------------------------------------------------

def test_preflight_writes_no_files_in_out_dir(tmp_path: Path):
    """Even when ``--out`` is the will-be-created flavor, the command
    must not actually create it — that's the next stage's job. Read-only
    by contract."""
    hist_xlsx, hist_docx, new_xlsx, out = _outside_repo_pilot(tmp_path)
    assert not out.exists()
    report = preflight(hist_xlsx, hist_docx, new_xlsx, out)
    assert report.ok is True
    # Crucially: --out still doesn't exist after the preflight.
    assert not out.exists(), "pilot-preflight must not create --out"
    # And the input fixture directory's listing is unchanged.
    listing_before = sorted(p.name for p in (tmp_path / "pilot_inputs").iterdir())
    assert listing_before == ["finished_report.docx", "historical.xlsx", "new_period.xlsx"]


# ---------------------------------------------------------------------------
# Docs/runtime exit-code parity for the new ``12`` code
# ---------------------------------------------------------------------------

def test_runtime_emits_exit_code_12(tmp_path: Path):
    """Direct functional pin: drive ``pilot-preflight`` to the privacy
    refusal branch and confirm the runtime returns the integer ``12``.

    This is the runtime side of the docs-vs-runtime exit-code parity
    contract enforced by ``test_command_reference_docs.py``."""
    _, hist_docx, new_xlsx, out = _outside_repo_pilot(tmp_path)
    inside = REPO_ROOT / "samples" / "__pp_runtime_emits_12__.xlsx"
    rc = cli_main([
        "pilot-preflight",
        "--historical-excel", str(inside),
        "--historical-word", str(hist_docx),
        "--new-excel", str(new_xlsx),
        "--out", str(out),
    ])
    assert rc == 12


def test_command_reference_exit_code_map_lists_12():
    """The exit-code map at the bottom of docs/command_reference.md
    must carry the new ``12`` row for ``pilot-preflight``. The generic
    docs parity test in ``test_command_reference_docs.py`` covers this
    indirectly via ``EXPECTED_EXIT_CODES``; this test adds a direct,
    grep-able pin so a doc edit that drops the row fails with a
    targeted message."""
    text = COMMAND_REFERENCE.read_text(encoding="utf-8")
    # Match a table row shaped like ``| `12` | … | `pilot-preflight` |``.
    row_re = re.compile(
        r"^\|\s*`12`\s*\|[^|]*\|\s*`pilot-preflight`\s*\|\s*$",
        re.MULTILINE,
    )
    assert row_re.search(text), (
        "docs/command_reference.md exit-code map must list a `12` row "
        "emitted by `pilot-preflight`. Add it back before merging."
    )


def test_command_reference_has_pilot_preflight_section():
    """Belt-and-suspenders: pin the per-command section heading. The
    generic docs parity test already enforces this via
    ``EXPECTED_SUBCOMMANDS``, but a direct heading check fails with a
    sharper message if the section is removed or renamed."""
    text = COMMAND_REFERENCE.read_text(encoding="utf-8")
    heading_re = re.compile(r"^##\s+`pilot-preflight`\s*$", re.MULTILINE)
    assert heading_re.search(text), (
        "docs/command_reference.md must carry a `## `pilot-preflight`` "
        "section heading."
    )


def test_real_file_pilot_doc_links_pilot_preflight_before_command_sequence():
    """``docs/real_file_pilot.md`` must point operators at the new
    command before the full pilot sequence — otherwise the doc still
    reads as if no preflight existed, defeating the point of adding
    it."""
    text = REAL_FILE_PILOT.read_text(encoding="utf-8")
    # Anchor on the first 2a header so we know we're upstream of the
    # actual command sequence in §2b.
    pre_section_2a, _, post_section_2a = text.partition(
        "### 2a. Environment setup"
    )
    assert post_section_2a, "expected §2a heading in docs/real_file_pilot.md"
    assert "pilot-preflight" in pre_section_2a, (
        "docs/real_file_pilot.md must mention `pilot-preflight` before "
        "the §2a Environment setup section so operators see the gate "
        "before the command sequence."
    )
    # And the linked section in command_reference.md must exist (the
    # heading-existence test above ensures the target is real).
    assert "command_reference.md#pilot-preflight" in pre_section_2a


def _main_py_return_codes_for_command(command: str) -> set[int]:
    """Walk ``src/main.py``'s AST and return every ``return <int>`` that
    appears inside the ``if args.command == "<command>":`` block.

    This is the minimal-surgery slice of the same AST trick
    ``test_command_reference_docs.py`` already uses; lifting it here
    keeps the parity check scoped to a single command without coupling
    the two test files.
    """
    tree = ast.parse(MAIN_PY.read_text(encoding="utf-8"))
    codes: set[int] = set()
    for node in ast.walk(tree):
        if not isinstance(node, ast.If):
            continue
        test = node.test
        if (
            isinstance(test, ast.Compare)
            and isinstance(test.left, ast.Attribute)
            and test.left.attr == "command"
            and len(test.ops) == 1
            and isinstance(test.ops[0], ast.Eq)
            and len(test.comparators) == 1
            and isinstance(test.comparators[0], ast.Constant)
            and test.comparators[0].value == command
        ):
            for sub in ast.walk(node):
                if isinstance(sub, ast.Return) and sub.value is not None:
                    for child in ast.walk(sub.value):
                        if (
                            isinstance(child, ast.Constant)
                            and isinstance(child.value, int)
                            and not isinstance(child.value, bool)
                        ):
                            codes.add(child.value)
    return codes


def test_pilot_preflight_runtime_codes_match_docs():
    """The codes ``src/main.py`` returns from the ``pilot-preflight``
    branch must be exactly the codes ``docs/command_reference.md``
    documents for it (``0``, ``2``, ``12``). A drift in either
    direction fails this test instead of slipping past review."""
    runtime_codes = _main_py_return_codes_for_command("pilot-preflight")
    assert runtime_codes == {0, 2, 12}, (
        f"src/main.py emits codes {sorted(runtime_codes)} from the "
        f"`pilot-preflight` branch, expected {{0, 2, 12}}. Update "
        f"either the runtime or the docs to bring them back in sync."
    )


# ---------------------------------------------------------------------------
# format_report shape — keep the operator-visible labels stable
# ---------------------------------------------------------------------------

def test_format_report_lists_every_flag_label(tmp_path: Path):
    """Each per-check line must name its CLI flag so the operator can
    grep ``--historical-excel`` and find which line to fix. A regression
    that printed positional indices instead would fail this test."""
    hist_xlsx, hist_docx, new_xlsx, out = _outside_repo_pilot(tmp_path)
    text = format_report(preflight(hist_xlsx, hist_docx, new_xlsx, out))
    for flag in (
        "--historical-excel", "--historical-word", "--new-excel", "--out",
    ):
        assert flag in text, f"format_report missing label {flag!r}"


def test_format_report_trailing_newline(tmp_path: Path):
    """Stable trailing newline keeps copy-paste tidy in chat clients
    and avoids the 'no newline at end of file' warning when an operator
    redirects the output to a text file. Matches the same contract
    pinned by ``test_pilot_summary.test_format_summary_output_ends_in_newline``."""
    hist_xlsx, hist_docx, new_xlsx, out = _outside_repo_pilot(tmp_path)
    text = format_report(preflight(hist_xlsx, hist_docx, new_xlsx, out))
    assert text.endswith("\n")


def test_preflight_report_ok_property_consistency(tmp_path: Path):
    """``PreflightReport.ok`` must agree with the per-check counts so
    callers can rely on either side."""
    hist_xlsx, hist_docx, new_xlsx, out = _outside_repo_pilot(tmp_path)
    clean = preflight(hist_xlsx, hist_docx, new_xlsx, out)
    assert isinstance(clean, PreflightReport)
    assert clean.ok is True
    assert clean.inside_repo_count == 0
    assert clean.other_failure_count == 0

    inside = REPO_ROOT / "samples" / "__pp_ok_property__.xlsx"
    dirty = preflight(inside, hist_docx, new_xlsx, out)
    assert dirty.ok is False
    assert dirty.inside_repo_count == 1
