"""Privacy preflight advisory for real-file pilots.

The repo's ``samples/`` folder is reserved for synthetic fixtures (see
CLAUDE.md "Repo Conventions" and "Privacy Rules"). Real company Excel
or Word files must never live inside the repo tree — even if
``.gitignore`` currently masks them — because a future ``git add -A``,
a stray editor save, or an inadvertent rule change could leak them into
version control.

This module is intentionally narrow:

  * It looks only at the command-line paths the operator passed in.
  * It never reads file contents, never blocks execution, never mutates
    matching, confirmation, rendering, or validation logic.
  * It emits a single advisory block to a caller-supplied stream when
    any input/output path resolves inside ``<repo>/samples/``.

The advisory is informational; it is the operator's responsibility to
keep real files outside the repo and to refuse to commit them. See
``docs/real_file_pilot.md`` for the supported pilot workflow.
"""
from __future__ import annotations

from pathlib import Path
from typing import IO, List, Optional, Sequence, Tuple

# Resolve once at import: ``src/preflight.py`` -> repo root is parents[1].
# Using ``.resolve()`` so symlinked checkouts (e.g. /tmp -> /private/tmp
# on macOS) compare against the same canonical form as caller paths.
_REPO_ROOT = Path(__file__).resolve().parents[1]
_SAMPLES_DIR = (_REPO_ROOT / "samples").resolve()


def _resolves_under(path: Path, root: Path) -> bool:
    """Return True iff ``path`` resolves to ``root`` or a descendant.

    Resolves both sides so a relative ``samples/synthetic/foo.xlsx``
    invocation from a different cwd still matches. Non-existent paths
    are still resolvable in Python 3.6+, so this works for output paths
    that have not been written yet.
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


def check_paths(
    labeled_paths: Sequence[Tuple[str, Optional[Path]]],
) -> List[str]:
    """Return one warning string per path that lands inside ``samples/``.

    ``labeled_paths`` is a sequence of ``(flag, path)`` pairs, e.g.
    ``[("--excel", Path("samples/synthetic/historical.xlsx"))]``. ``None``
    paths are skipped so callers can pass optional CLI args directly.

    The function is pure: no I/O, no stderr writes. Callers that want to
    surface the advisory should pass the result to ``format_advisory``
    and print it themselves — that keeps this module trivially testable
    and keeps CLI presentation in ``main.py``.
    """
    warnings: List[str] = []
    for label, path in labeled_paths:
        if path is None:
            continue
        if _resolves_under(Path(path), _SAMPLES_DIR):
            warnings.append(
                f"{label} {path} is inside the repo's samples/ folder, "
                "which is reserved for synthetic fixtures."
            )
    return warnings


def format_advisory(warnings: Sequence[str]) -> str:
    """Format a stderr-ready advisory block. Returns ``""`` when empty.

    The wording is deliberately operator-facing: it names the offending
    flag(s), reminds the reader that real files must stay outside the
    repo, and points at the pilot doc so the next-best action is one
    click away.
    """
    if not warnings:
        return ""
    bar = "!" * 72
    lines = [bar, "PRIVACY PREFLIGHT ADVISORY"]
    lines.extend(f"  - {w}" for w in warnings)
    lines.append(
        "Real company Excel/Word files must live OUTSIDE this repo "
        "(e.g. ~/pilot_data/) and must NEVER be committed."
    )
    lines.append("See docs/real_file_pilot.md for the supported pilot workflow.")
    lines.append(bar)
    return "\n".join(lines)


def emit_advisory(
    labeled_paths: Sequence[Tuple[str, Optional[Path]]],
    stream: IO[str],
) -> bool:
    """Convenience wrapper: check + format + write to ``stream``.

    Returns ``True`` when an advisory was emitted, ``False`` otherwise.
    Never raises on a clean input set; never alters downstream behavior.
    """
    advisory = format_advisory(check_paths(labeled_paths))
    if not advisory:
        return False
    print(advisory, file=stream)
    return True
