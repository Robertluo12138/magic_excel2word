"""Coverage and traceability checks over the matcher's output.

CLAUDE.md's first product risk is generating a Word report that looks fine
but silently omits or invents a metric. The validator's job is to make that
risk visible: count what was mapped vs. left hanging, and surface the
hanging cases by location and snippet so a human reviewer can decide.

The accounting distinguishes ``total_visible`` (every numeric token the
profiler found, including policy-excluded ones) from ``eligible`` (the
subset the matcher actually tried to map). The coverage ratio is over
eligible so excluded markers don't dilute the signal — but the visible
total is still shown so a reviewer can confirm nothing was dropped.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List

from .value_matcher import WordMatch


@dataclass
class CoverageSummary:
    total: int = 0  # every Word number found, including EXCLUDED
    by_confidence: Dict[str, int] = field(default_factory=dict)
    unresolved: List[WordMatch] = field(default_factory=list)
    low_confidence: List[WordMatch] = field(default_factory=list)
    ambiguous: List[WordMatch] = field(default_factory=list)
    excluded: List[WordMatch] = field(default_factory=list)

    @property
    def mapped(self) -> int:
        return self.by_confidence.get("HIGH", 0) + self.by_confidence.get("MEDIUM", 0)

    @property
    def eligible(self) -> int:
        """Word numbers the matcher tried to map (excludes policy-skipped)."""
        return self.total - self.by_confidence.get("EXCLUDED", 0)

    @property
    def coverage_ratio(self) -> float:
        return self.mapped / self.eligible if self.eligible else 0.0


def summarize(matches: List[WordMatch]) -> CoverageSummary:
    summary = CoverageSummary(total=len(matches))
    counts = {"HIGH": 0, "MEDIUM": 0, "LOW": 0, "UNRESOLVED": 0, "EXCLUDED": 0}
    for m in matches:
        counts[m.confidence] = counts.get(m.confidence, 0) + 1
        if m.confidence == "UNRESOLVED":
            summary.unresolved.append(m)
        elif m.confidence == "LOW":
            summary.low_confidence.append(m)
        elif m.confidence == "EXCLUDED":
            summary.excluded.append(m)
        if m.note.startswith("ambiguous"):
            summary.ambiguous.append(m)
    summary.by_confidence = counts
    return summary


def format_console_summary(summary: CoverageSummary) -> str:
    lines = [
        "Learn-mode coverage summary",
        "===========================",
        f"Total visible Word numbers : {summary.total}",
        f"  EXCLUDED by policy       : {summary.by_confidence.get('EXCLUDED', 0)}",
        f"Eligible for mapping       : {summary.eligible}",
        f"  HIGH       : {summary.by_confidence.get('HIGH', 0)}",
        f"  MEDIUM     : {summary.by_confidence.get('MEDIUM', 0)}",
        f"  LOW        : {summary.by_confidence.get('LOW', 0)}",
        f"  UNRESOLVED : {summary.by_confidence.get('UNRESOLVED', 0)}",
        f"Mapped (HIGH+MEDIUM)       : {summary.mapped} ({summary.coverage_ratio:.1%} of eligible)",
    ]
    if summary.ambiguous:
        lines.append("")
        lines.append(f"Ambiguous picks needing human review: {len(summary.ambiguous)}")
        for m in summary.ambiguous[:5]:
            lines.append(f"  - {m.word_number.location}: '{m.word_number.raw}' — {m.note}")
    if summary.low_confidence:
        lines.append("")
        lines.append(f"Low-confidence numbers: {len(summary.low_confidence)}")
        for m in summary.low_confidence[:10]:
            lines.append(f"  - {m.word_number.location}: '{m.word_number.raw}' in '{_truncate(m.word_number.snippet, 60)}'")
    if summary.unresolved:
        lines.append("")
        lines.append(f"UNRESOLVED numbers (no Excel source found): {len(summary.unresolved)}")
        for m in summary.unresolved[:10]:
            lines.append(f"  - {m.word_number.location}: '{m.word_number.raw}' in '{_truncate(m.word_number.snippet, 60)}'")
    if summary.excluded:
        lines.append("")
        lines.append(f"EXCLUDED by explicit policy: {len(summary.excluded)}")
        # Sample a few so a reviewer can sanity-check the policy without
        # scrolling through every date marker.
        for m in summary.excluded[:5]:
            lines.append(f"  - {m.word_number.location}: '{m.word_number.raw}' — {m.note}")
        if len(summary.excluded) > 5:
            lines.append(f"  … and {len(summary.excluded) - 5} more (see mapping_review.xlsx)")
    return "\n".join(lines)


def _truncate(text: str, limit: int) -> str:
    text = text.replace("\n", " ").strip()
    return text if len(text) <= limit else text[: limit - 1] + "…"
