"""Persist matcher output as human-reviewable artifacts.

Two outputs:
  * ``mapping_review.xlsx`` — one row per Word number with the top candidate
    and enough metadata for a reviewer to either confirm the mapping or
    correct it. This file is the substrate for the future human-confirmed
    mapping workflow.
  * ``confidence_report.md`` — narrative summary highlighting unresolved and
    low-confidence numbers, so the *first* thing a reviewer sees is what
    needs attention, not what already worked.
"""
from __future__ import annotations

from pathlib import Path
from typing import Dict, List

import openpyxl

from .template_builder import derive_review_status
from .validator import CoverageSummary
from .value_matcher import WordMatch

# The first four columns are the trust-slice keys: a stable Word ID (the
# join key with auto_mapping.yml) plus the three columns that record what
# the template builder did with the token. Keeping them up front means a
# reviewer skimming the XLSX sees the audit story before the candidate
# details.
REVIEW_HEADERS = [
    "Word ID",
    "Word Location",
    "Word Snippet",
    "Word Raw Token",
    "Word Value",
    "Word Unit",
    "Confidence",
    "Review Status",
    "Placeholder Status",
    "Placeholder",
    "Top Excel Sheet",
    "Top Excel Cell",
    "Top Excel Value",
    "Top Row Context",
    "Top Column Context",
    "Interpretation",
    "Value Score",
    "Context Score",
    "Overlap Tokens",
    "# Alt Candidates",
    "Note",
]


def write_mapping_review(
    matches: List[WordMatch],
    out_path: Path,
    word_ids: List[str],
    placeholder_status: Dict[int, str],
) -> Path:
    """Write the ``mapping_review.xlsx`` workbook to ``out_path``.

    ``word_ids`` and ``placeholder_status`` must come from the same
    template-builder run that produced ``auto_mapping.yml``; the validator
    relies on the XLSX agreeing with the YAML on every join key.
    """
    if len(word_ids) != len(matches):
        raise ValueError(
            f"word_ids length {len(word_ids)} != matches length {len(matches)}"
        )
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "mapping_review"
    ws.append(REVIEW_HEADERS)
    for i, m in enumerate(matches):
        wn = m.word_number
        top = m.chosen or (m.candidates[0] if m.candidates else None)
        ph_status = placeholder_status.get(i, "skipped_unknown")
        placeholder = "{{ " + word_ids[i] + " }}" if ph_status == "applied" else ""
        review = derive_review_status(m.confidence, ph_status)
        trust_cols = [word_ids[i]]
        trust_tail = [review, ph_status, placeholder]
        if top is not None:
            cell = top.cell
            row = trust_cols + [
                wn.location,
                _truncate(wn.snippet, 200),
                wn.raw,
                wn.value,
                wn.unit or "",
                m.confidence,
            ] + trust_tail + [
                cell.sheet,
                cell.address,
                cell.numeric_value,
                " | ".join(cell.row_context[:4]),
                " | ".join(cell.column_context[:4]),
                top.interpretation,
                round(top.value_score, 3),
                round(top.context_score, 3),
                ", ".join(top.overlap_tokens),
                max(0, len(m.candidates) - 1),
                m.note,
            ]
        else:
            row = trust_cols + [
                wn.location,
                _truncate(wn.snippet, 200),
                wn.raw,
                wn.value,
                wn.unit or "",
                m.confidence,
            ] + trust_tail + [
                "", "", "", "", "", "", "", "", "", 0, m.note,
            ]
        ws.append(row)

    ws.column_dimensions["A"].width = 12  # Word ID
    ws.column_dimensions["B"].width = 18  # Location
    ws.column_dimensions["C"].width = 50  # Snippet
    ws.column_dimensions["D"].width = 16  # Raw token
    ws.freeze_panes = "B2"
    wb.save(str(out_path))
    return out_path


def write_confidence_report(
    matches: List[WordMatch],
    summary: CoverageSummary,
    out_path: Path,
) -> Path:
    """Write a Markdown overview that leads with what needs review."""
    out_path.parent.mkdir(parents=True, exist_ok=True)
    lines: List[str] = []
    lines.append("# Learn-mode confidence report")
    lines.append("")
    lines.append(f"- Total visible Word numbers: **{summary.total}**")
    lines.append(f"  - EXCLUDED by policy: {summary.by_confidence.get('EXCLUDED', 0)}")
    lines.append(f"- Eligible for mapping: **{summary.eligible}**")
    lines.append(f"  - HIGH: {summary.by_confidence.get('HIGH', 0)}")
    lines.append(f"  - MEDIUM: {summary.by_confidence.get('MEDIUM', 0)}")
    lines.append(f"  - LOW: {summary.by_confidence.get('LOW', 0)}")
    lines.append(f"  - UNRESOLVED: {summary.by_confidence.get('UNRESOLVED', 0)}")
    lines.append(f"- Mapped (HIGH+MEDIUM): **{summary.mapped} ({summary.coverage_ratio:.1%} of eligible)**")
    lines.append("")
    lines.append("> This report exists so a human can see at a glance which Word numbers")
    lines.append("> are not yet trustworthy. Unresolved and low-confidence entries must")
    lines.append("> be reviewed before any mapping is confirmed for production rendering.")
    lines.append("> EXCLUDED entries were skipped by explicit policy (e.g. date markers");
    lines.append("> like '5月' or '2026年'); the reviewer should still scan them to")
    lines.append("> confirm the policy applied correctly.")
    lines.append("")

    lines.append("## UNRESOLVED")
    if not summary.unresolved:
        # Scope the claim to "eligible" — saying "every Word number has a
        # source" would be a false success when EXCLUDED rows exist, since
        # those were never matched and have no candidates.
        msg = "_None — every eligible Word number has at least one candidate Excel source._"
        excluded_count = summary.by_confidence.get("EXCLUDED", 0)
        if excluded_count:
            msg += f" ({excluded_count} number(s) were EXCLUDED by policy — see section below.)"
        lines.append(msg)
    else:
        for m in summary.unresolved:
            wn = m.word_number
            lines.append(f"- `{wn.location}` — raw `{wn.raw}` (value={wn.value}, unit={wn.unit or '∅'})")
            lines.append(f"    - snippet: {_truncate(wn.snippet, 160)}")
    lines.append("")

    lines.append("## LOW confidence")
    if not summary.low_confidence:
        lines.append("_None._")
    else:
        for m in summary.low_confidence:
            wn = m.word_number
            top = m.chosen or (m.candidates[0] if m.candidates else None)
            lines.append(f"- `{wn.location}` — raw `{wn.raw}`")
            lines.append(f"    - snippet: {_truncate(wn.snippet, 160)}")
            if top is not None:
                lines.append(f"    - top guess: {top.cell.sheet}!{top.cell.address} = {top.cell.numeric_value} ({top.interpretation})")
    lines.append("")

    lines.append("## Ambiguous picks (MEDIUM with ties)")
    if not summary.ambiguous:
        lines.append("_None._")
    else:
        for m in summary.ambiguous:
            wn = m.word_number
            lines.append(f"- `{wn.location}` — raw `{wn.raw}`: {m.note}")
            for c in m.candidates[:3]:
                lines.append(f"    - {c.cell.sheet}!{c.cell.address} = {c.cell.numeric_value} (row_ctx={c.cell.row_context[:2]}, col_ctx={c.cell.column_context[:2]})")
    lines.append("")

    lines.append("## EXCLUDED by explicit policy")
    if not summary.excluded:
        lines.append("_None._")
    else:
        lines.append(f"_{len(summary.excluded)} number(s) captured for audit but skipped by the matcher._")
        lines.append("")
        for m in summary.excluded:
            wn = m.word_number
            lines.append(f"- `{wn.location}` — raw `{wn.raw}` — {m.note}")
            lines.append(f"    - snippet: {_truncate(wn.snippet, 160)}")

    out_path.write_text("\n".join(lines), encoding="utf-8")
    return out_path


def _truncate(text: str, limit: int) -> str:
    text = (text or "").replace("\n", " ").strip()
    return text if len(text) <= limit else text[: limit - 1] + "…"
