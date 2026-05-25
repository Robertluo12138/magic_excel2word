"""Docs consistency check for ``docs/command_reference.md``.

The CLI exposes eight subcommands; the per-command reference must mention
every one by name so an operator can look up its contract without having
to re-read ``README.md`` or run ``--help``. Adding a new subcommand to
``src/main.py`` and forgetting to grow a section in the reference should
fail this test, not slip past review.

This is a pure docs/CLI surface check — it imports ``build_parser`` only
to discover the current subcommand set, never executes the CLI, never
writes files, and never touches synthetic or real data.
"""
from __future__ import annotations

import argparse
import ast
import re
from pathlib import Path

import pytest

from src.main import build_parser


REPO_ROOT = Path(__file__).resolve().parents[1]
COMMAND_REFERENCE = REPO_ROOT / "docs" / "command_reference.md"
MAIN_PY = REPO_ROOT / "src" / "main.py"


# Keep this list in lockstep with build_parser() in src/main.py.
# The cross-check test below guards against drift in either direction.
EXPECTED_SUBCOMMANDS = (
    "generate-synthetic",
    "learn",
    "validate-artifacts",
    "confirm-mapping",
    "run-preview",
    "render-docx",
    "validate-render",
    "pilot-summary",
)


def _discover_cli_subcommands() -> set[str]:
    """Return the set of subcommands currently registered on the CLI parser."""
    parser = build_parser()
    sub_actions = [
        action
        for action in parser._actions
        if isinstance(action, argparse._SubParsersAction)
    ]
    assert len(sub_actions) == 1, "expected exactly one subparsers group on the CLI"
    return set(sub_actions[0].choices.keys())


def test_command_reference_file_exists():
    """README and docs/real_file_pilot.md both link to this file — the link
    targets must resolve. If this fails, neither doc points to a real
    artifact and the operator-facing reference is gone."""
    assert COMMAND_REFERENCE.exists(), (
        "docs/command_reference.md must exist — it is the per-command "
        "reference that README.md and docs/real_file_pilot.md link to."
    )


def test_expected_subcommand_list_matches_cli_parser():
    """Guard against drift: the hardcoded EXPECTED_SUBCOMMANDS list must
    track build_parser() exactly. If you added or removed a subcommand,
    update this list AND the docs/command_reference.md sections."""
    actual = _discover_cli_subcommands()
    expected = set(EXPECTED_SUBCOMMANDS)
    assert actual == expected, (
        "EXPECTED_SUBCOMMANDS drifted from src.main.build_parser(): "
        f"missing_from_test={sorted(expected - actual)}, "
        f"extra_in_cli={sorted(actual - expected)}"
    )


@pytest.mark.parametrize("name", EXPECTED_SUBCOMMANDS)
def test_command_reference_has_section_for_every_cli_subcommand(name: str):
    """Every CLI subcommand must have its own ``## `<name>``` section
    heading in docs/command_reference.md. A plain backticked mention is
    not enough — the bottom exit-code table also lists every command name
    in backticks, so a removed section would otherwise false-pass. The
    heading is what an operator scrolls to looking for the contract."""
    text = COMMAND_REFERENCE.read_text(encoding="utf-8")
    heading_pattern = re.compile(
        rf"^##\s+`{re.escape(name)}`\s*$", re.MULTILINE
    )
    assert heading_pattern.search(text), (
        f"docs/command_reference.md is missing a `## \\`{name}\\`` section "
        f"heading — add a dedicated section for it before merging. A mention "
        f"in the exit-code table or prose is not sufficient."
    )


# Six fields every per-command section must document. Trailing ``?`` on
# "Read-only" / "Paste-safe" is accepted because the current doc style
# uses interrogative labels (``**Read-only?**``); the test only cares
# that the field is present, not its punctuation.
REQUIRED_SECTION_LABELS = (
    "Purpose",
    "Required inputs",
    "Outputs written",
    "Exit codes",
    "Read-only",
    "Paste-safe",
)


def _section_bodies() -> dict[str, str]:
    """Split docs/command_reference.md by ``## `<name>``` headings and
    return a dict mapping subcommand name to the section body (the text
    between one heading and the next, or end-of-file).
    """
    text = COMMAND_REFERENCE.read_text(encoding="utf-8")
    heading_re = re.compile(r"^##\s+`([a-z][a-z-]*)`\s*$", re.MULTILINE)
    matches = list(heading_re.finditer(text))
    bodies: dict[str, str] = {}
    for i, m in enumerate(matches):
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        bodies[m.group(1)] = text[m.end():end]
    return bodies


@pytest.mark.parametrize("name", EXPECTED_SUBCOMMANDS)
@pytest.mark.parametrize("label", REQUIRED_SECTION_LABELS)
def test_command_reference_section_documents_required_label(name: str, label: str):
    """Each per-command section must document Purpose, Required inputs,
    Outputs written, Exit codes, Read-only, and Paste-safe — the six
    contract fields an operator scans for when looking up a command.

    Missing any one of these means the section silently drops a piece of
    the contract README.md and docs/real_file_pilot.md both rely on.
    """
    bodies = _section_bodies()
    body = bodies.get(name, "")
    # ``\*\*<label>[?]?\*\*`` — bold marker, optional trailing question mark.
    label_re = re.compile(rf"\*\*{re.escape(label)}\??\*\*")
    assert label_re.search(body), (
        f"docs/command_reference.md section `## \\`{name}\\`` is missing the "
        f"`**{label}**` label. Each section must document all six fields "
        f"(Purpose, Required inputs, Outputs written, Exit codes, "
        f"Read-only, Paste-safe) so the operator-facing contract stays "
        f"complete."
    )


# Shared codes (0, 2) plus gate-specific codes 3-11. Codes 1 and >11 are
# documented as unused; this test pins the current contract.
EXPECTED_EXIT_CODES = (0, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)


def _exit_code_map_codes() -> set[int]:
    """Parse the exit-code table at the bottom of the reference and
    return the set of integer codes it lists. Matches table rows shaped
    like ``| `<N>` | <meaning> | <emitter> |``."""
    text = COMMAND_REFERENCE.read_text(encoding="utf-8")
    return {
        int(m.group(1))
        for m in re.finditer(r"^\|\s*`(\d+)`\s*\|", text, flags=re.MULTILINE)
    }


def test_exit_code_map_lists_shared_and_gate_specific_codes():
    """The exit-code map must enumerate the shared codes (0, 2) and every
    gate-specific code (3 through 11). Automation branches on a single
    integer, so a missing row in the map means a missing contract.
    """
    map_codes = _exit_code_map_codes()
    missing = [c for c in EXPECTED_EXIT_CODES if c not in map_codes]
    assert not missing, (
        f"docs/command_reference.md exit-code map is missing rows for: "
        f"{missing}. The map must list shared codes 0 and 2 plus "
        f"gate-specific codes 3 through 11."
    )


def _main_py_return_codes() -> set[int]:
    """Walk src/main.py's AST and collect every integer constant that
    appears inside a ``return`` statement. Uses AST rather than regex so
    conditional returns like ``return 0 if report.ok else 4`` correctly
    contribute both codes.
    """
    tree = ast.parse(MAIN_PY.read_text(encoding="utf-8"))
    codes: set[int] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Return) and node.value is not None:
            for child in ast.walk(node.value):
                if (
                    isinstance(child, ast.Constant)
                    and isinstance(child.value, int)
                    and not isinstance(child.value, bool)
                ):
                    codes.add(child.value)
    return codes


def test_no_undocumented_gate_specific_code_in_main_py():
    """If src/main.py emits a ``return <N>`` whose code (>= 3) is not in
    the exit-code map, the runtime and the docs have drifted. Codes 0
    and 2 are shared and code 1 is the documented unexpected-error
    fallback, so all three are excluded from the parity check.
    """
    runtime_codes = _main_py_return_codes()
    gate_codes = {c for c in runtime_codes if c >= 3}
    documented = _exit_code_map_codes()
    undocumented = sorted(gate_codes - documented)
    assert not undocumented, (
        f"src/main.py emits gate-specific exit code(s) {undocumented} "
        f"that are not listed in docs/command_reference.md's exit-code "
        f"map. Add a row for each new code (or remove the return) before "
        f"merging."
    )
