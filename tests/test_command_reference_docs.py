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
import re
from pathlib import Path

import pytest

from src.main import build_parser


REPO_ROOT = Path(__file__).resolve().parents[1]
COMMAND_REFERENCE = REPO_ROOT / "docs" / "command_reference.md"


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
