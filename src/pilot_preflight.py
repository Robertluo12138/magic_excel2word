"""Read-only pilot preflight check.

Before running the real-file pilot sequence (``learn`` →
``validate-artifacts`` → ``confirm-mapping`` → ``run-preview`` →
``render-docx`` → ``validate-render``), an operator can run this
command to verify the pilot's input and output paths are well-formed
**without ever reading file contents**. The check is intentionally
narrow:

  * Three input paths (``--historical-excel``, ``--historical-word``,
    ``--new-excel``) must exist and have the expected ``.xlsx`` /
    ``.docx`` suffix.
  * The output directory (``--out``) must already exist as a
    directory, or be creatable later (i.e. it must not currently be
    something else like a regular file, AND its nearest existing
    ancestor must itself be a directory so a downstream
    ``mkdir(parents=True)`` can build the missing layers).
  * No input or output path may resolve inside this repository tree.
    Real company files belong outside the repo (see
    ``docs/real_file_pilot.md`` §1); the repo's own folders are
    reserved for source code and synthetic fixtures.

This module is intentionally distinct from ``src/preflight.py``:
``preflight.py`` is the **informational privacy advisory** wired into
every pilot-sequence subcommand (it only warns when a path lands
under ``samples/`` and never changes an exit code), while this module
is the **explicit operator-facing preflight gate** with its own CLI
subcommand and a hard refusal exit code when a pilot path lands
anywhere inside the repo tree.

Output is redacted: only the flag label and the path **basename**
ever reach stdout/stderr. Full paths, document content, parent-tree
folders, and anything inside a file never leak.

Failure surface (the CLI translates these to exit codes):

  * Any path resolves inside the repo tree → exit ``12``. This is the
    privacy refusal — strictly stronger than the informational
    ``preflight.py`` advisory.
  * Any other per-path issue (missing input, wrong suffix, ``--out``
    is an existing non-directory) → exit ``2``.

Read-only by design: never opens any document, never mutates a file,
never calls out to an LLM, GUI, network, or Microsoft Word
automation. Matching, confirmation, rendering, and validation logic
are untouched.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import List


# Resolve once at import: ``src/pilot_preflight.py`` -> repo root is parents[1].
# ``.resolve()`` so symlinked checkouts (e.g. /tmp -> /private/tmp on
# macOS) compare against the same canonical form as caller paths.
_REPO_ROOT = Path(__file__).resolve().parents[1]


# Per-check status values. ``ok`` and ``will-be-created`` are passing;
# everything else is a failure. ``inside-repo`` is the privacy refusal
# and routes to exit 12; the rest route to exit 2 (shared input-error).
_OK_STATUSES = frozenset({"ok", "will-be-created"})
_INSIDE_REPO = "inside-repo"


@dataclass
class PathCheck:
    """One check result. The basename is the only path slice ever printed."""

    label: str          # e.g. "--historical-excel"
    basename: str       # path.name only — full path never leaks
    status: str         # "ok", "will-be-created", "missing", "bad-suffix",
                        # "not-a-file", "not-a-dir", "cannot-create",
                        # or "inside-repo"
    detail: str = ""    # short label-safe explanation; no path content


@dataclass
class PreflightReport:
    checks: List[PathCheck] = field(default_factory=list)

    @property
    def inside_repo_count(self) -> int:
        return sum(1 for c in self.checks if c.status == _INSIDE_REPO)

    @property
    def other_failure_count(self) -> int:
        return sum(
            1 for c in self.checks
            if c.status != _INSIDE_REPO and c.status not in _OK_STATUSES
        )

    @property
    def ok(self) -> bool:
        return self.inside_repo_count == 0 and self.other_failure_count == 0


def _resolves_under(path: Path, root: Path) -> bool:
    """True iff ``path`` resolves to ``root`` or a descendant.

    ``Path.resolve()`` handles non-existent targets in Python 3.6+, so
    output paths that have not been created yet still resolve correctly.
    """
    try:
        resolved = Path(path).resolve()
    except (OSError, RuntimeError):
        return False
    try:
        resolved.relative_to(root)
    except ValueError:
        return False
    return True


def _check_input_file(label: str, path: Path, expected_suffix: str) -> PathCheck:
    """Check an input file: inside-repo gate first, then existence/suffix.

    Inside-repo is checked first so an operator who typo'd a real-data
    path under the repo tree gets the privacy-refusal signal even when
    the file does not exist there yet — the same ordering rationale as
    ``preflight.py``'s advisory in the existing pilot subcommands.
    """
    basename = path.name
    if _resolves_under(path, _REPO_ROOT):
        return PathCheck(
            label=label,
            basename=basename,
            status=_INSIDE_REPO,
            detail="refused: real pilot inputs must live outside the repo tree",
        )
    if not path.exists():
        return PathCheck(
            label=label, basename=basename, status="missing",
            detail="file not found",
        )
    if path.suffix.lower() != expected_suffix.lower():
        return PathCheck(
            label=label, basename=basename, status="bad-suffix",
            detail=f"expected {expected_suffix}",
        )
    if not path.is_file():
        return PathCheck(
            label=label, basename=basename, status="not-a-file",
            detail="exists but is not a regular file",
        )
    return PathCheck(label=label, basename=basename, status="ok")


def _nearest_existing_ancestor(path: Path) -> Path:
    """Walk up ``path`` until we reach a component that exists on disk.

    Used by the creatability check: ``Path.mkdir(parents=True)`` can
    only build missing intermediate directories underneath an existing
    ancestor that is itself a directory. If the nearest existing
    ancestor is a regular file, the mkdir would fail — so we want to
    flag the path up front instead of letting the downstream stage
    crash mid-pilot.

    Terminates at the filesystem root (``parent == current``) as a
    defensive guard; in practice the root always exists.
    """
    current = Path(path).resolve()
    while not current.exists():
        parent = current.parent
        if parent == current:
            return current
        current = parent
    return current


def _check_out_dir(label: str, path: Path) -> PathCheck:
    """Check ``--out``: inside-repo gate first, then dir-or-creatable.

    Read-only by contract: we never ``mkdir`` here. Either the directory
    already exists (good), it doesn't but its nearest existing ancestor
    is a directory so a later ``mkdir(parents=True)`` would succeed
    (also good), or one of the failure modes fires: inside-repo,
    "exists but is not a directory", or "ancestor is not a directory
    so the path cannot be created later".
    """
    basename = path.name
    if _resolves_under(path, _REPO_ROOT):
        return PathCheck(
            label=label,
            basename=basename,
            status=_INSIDE_REPO,
            detail="refused: pilot output must live outside the repo tree",
        )
    if path.exists():
        if path.is_dir():
            return PathCheck(label=label, basename=basename, status="ok")
        return PathCheck(
            label=label, basename=basename, status="not-a-dir",
            detail="exists but is not a directory",
        )
    # Non-existent: only safe to defer creation if the nearest existing
    # ancestor is a directory. If it's a regular file, mkdir would fail.
    ancestor = _nearest_existing_ancestor(path)
    if not ancestor.is_dir():
        return PathCheck(
            label=label, basename=basename, status="cannot-create",
            detail="nearest existing ancestor is not a directory",
        )
    return PathCheck(
        label=label, basename=basename, status="will-be-created",
        detail="does not exist yet (downstream stage will create it)",
    )


def preflight(
    historical_excel: Path,
    historical_word: Path,
    new_excel: Path,
    out: Path,
) -> PreflightReport:
    """Run the four checks and return an aggregate report.

    Pure path/metadata: no document is opened, no file is mutated.
    """
    return PreflightReport(checks=[
        _check_input_file("--historical-excel", historical_excel, ".xlsx"),
        _check_input_file("--historical-word", historical_word, ".docx"),
        _check_input_file("--new-excel", new_excel, ".xlsx"),
        _check_out_dir("--out", out),
    ])


def format_report(report: PreflightReport) -> str:
    """Render the report as a redacted operator-facing string.

    Only flag labels and path basenames reach the output — never full
    paths, parent directories, or anything from inside a file.
    """
    label_width = max(len(c.label) for c in report.checks)
    name_width = max(len(c.basename) for c in report.checks)
    lines = ["pilot-preflight checks:"]
    for c in report.checks:
        bracket = c.status if not c.detail else f"{c.status}: {c.detail}"
        lines.append(
            f"  {c.label:<{label_width}} : "
            f"{c.basename:<{name_width}}  [{bracket}]"
        )
    if report.ok:
        lines.append("")
        lines.append("All checks passed. Pilot inputs/output look ready.")
    else:
        lines.append("")
        if report.inside_repo_count:
            lines.append(
                f"REFUSED: {report.inside_repo_count} path(s) resolve inside "
                "this repo tree. Move real files OUTSIDE this repo "
                "(e.g. ~/pilot_data/) and re-run. "
                "See docs/real_file_pilot.md §1."
            )
        if report.other_failure_count:
            lines.append(
                f"FAILED: {report.other_failure_count} path issue(s). "
                "Fix the listed inputs and re-run."
            )
    return "\n".join(lines) + "\n"
