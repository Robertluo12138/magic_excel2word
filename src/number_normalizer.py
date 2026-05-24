"""Parse and normalize numeric tokens found in Chinese business reports.

The matcher needs two things from this module:
1. Locate every numeric token in a free-form Chinese text snippet, capturing
   its raw substring, parsed value, unit hint, sign, and character offset.
2. For a parsed token, enumerate plausible *equivalent base values* an Excel
   cell could hold (e.g. ``1,234.56 万元`` ↔ ``12,345,600``), so the matcher
   can compare without committing to a single unit interpretation up front.

Both responsibilities are deterministic. AI rewrites or fuzzy guessing live
elsewhere; this module is the trusted source of "what number is on the page".
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import List, Optional, Tuple

# Longest units first so the alternation matches "万元" before "万".
_UNITS = [
    "万元", "亿元", "千元", "百万元",
    "万单", "亿单",
    "万人", "万次", "万个",
    "万",
    "元", "单", "人", "次", "个",
    "%", "‰",
]
_UNIT_ALT = "|".join(sorted(_UNITS, key=len, reverse=True))

# Plain numeric token: optional sign, digits possibly grouped by commas, optional
# decimal, optional unit. The leading sign is captured but optional so we can
# also pick up bare numbers embedded inside Chinese narrative text.
_NUM_RE = re.compile(
    r"(?P<sign>[-+])?"
    r"(?P<int>\d{1,3}(?:,\d{3})+|\d+)"
    r"(?P<frac>\.\d+)?"
    r"\s*"                       # tolerate "1,234 元" with intermediate space
    r"(?P<unit>" + _UNIT_ALT + r")?"
)

# Numbers immediately followed by one of these are date/period markers, not
# business metrics ("2026年", "5月", "第20周"). We still capture them so a
# reviewer can audit the exclusion in the mapping_review artifact; the
# matcher just skips matching them. Reviewers can extend the set if a real
# report needs it.
_DATE_SUFFIX_CHARS = set("年月日季周期号")

# Accounting parentheses, e.g. ``(1,234.56)`` meaning negative. Matched first
# so its span is reserved before _NUM_RE re-scans the same digits.
_PAREN_RE = re.compile(
    r"\((?P<int>\d{1,3}(?:,\d{3})+|\d+)"
    r"(?P<frac>\.\d+)?"
    r"(?P<unit>" + _UNIT_ALT + r")?\)"
)

WAN_UNITS = {"万元", "万单", "万人", "万次", "万个", "万"}
YI_UNITS = {"亿元", "亿单"}


@dataclass(frozen=True)
class ParsedNumber:
    """A numeric token located in a piece of text.

    ``value`` is the signed face value as written (so ``1.5万元`` parses to
    ``1.5``, not ``15000``). ``candidate_values`` does the unit expansion.

    ``exclusion_reason`` is set when the token is captured for the audit
    trail but should not be passed to the matcher — e.g. date/period markers
    like the ``5`` in ``5月``. Callers must still see these tokens (CLAUDE.md
    "never silently ignore a visible Word number"); the matcher skips them.
    """

    raw: str
    value: float
    unit: Optional[str]
    sign: int
    start: int
    end: int
    exclusion_reason: Optional[str] = None


def find_numbers(text: str) -> List[ParsedNumber]:
    """Return every numeric token in ``text`` ordered by appearance."""
    if not text:
        return []
    found: List[ParsedNumber] = []
    claimed: List[Tuple[int, int]] = []

    for m in _PAREN_RE.finditer(text):
        v = _to_float(m.group("int"), m.group("frac"))
        if v is None:
            continue
        found.append(ParsedNumber(
            raw=m.group(0),
            value=-v,
            unit=m.group("unit"),
            sign=-1,
            start=m.start(),
            end=m.end(),
        ))
        claimed.append((m.start(), m.end()))

    for m in _NUM_RE.finditer(text):
        if _overlaps(m.start(), m.end(), claimed):
            continue
        v = _to_float(m.group("int"), m.group("frac"))
        if v is None:
            continue
        unit = m.group("unit")
        # Date/period policy: unit-less numbers followed by 年/月/日/etc. are
        # captured with an exclusion_reason so they appear in the audit trail
        # but bypass the matcher.
        reason: Optional[str] = None
        if unit is None and m.end() < len(text) and text[m.end()] in _DATE_SUFFIX_CHARS:
            reason = f"date/period marker (followed by '{text[m.end()]}')"
        sign = -1 if m.group("sign") == "-" else 1
        found.append(ParsedNumber(
            raw=m.group(0),
            value=sign * v,
            unit=unit,
            sign=sign,
            start=m.start(),
            end=m.end(),
            exclusion_reason=reason,
        ))

    found.sort(key=lambda p: p.start)
    return found


def candidate_values(p: ParsedNumber) -> List[Tuple[float, str]]:
    """Return ``(base_value, interpretation)`` pairs to compare against Excel.

    The interpretation label is preserved through to the mapping review so a
    human reviewer can see *why* the matcher accepted a candidate, not just
    the bare number. Order does not matter; the matcher picks the best one.
    """
    out: List[Tuple[float, str]] = [(p.value, "as_written")]
    u = p.unit
    if u in WAN_UNITS:
        out.append((p.value * 10_000, f"{u}→base_unit"))
    if u in YI_UNITS:
        out.append((p.value * 100_000_000, f"{u}→base_unit"))
    if u == "千元":
        out.append((p.value * 1_000, "千元→元"))
    if u == "百万元":
        out.append((p.value * 1_000_000, "百万元→元"))
    if u == "%":
        out.append((p.value / 100, "%→decimal"))
    if u == "‰":
        out.append((p.value / 1_000, "‰→decimal"))
    # If the Word text writes a large raw figure but the Excel sheet stores it
    # in 万-units (or vice versa), allow a symmetric reading. Guarded by
    # magnitude so we do not flood candidates for small numbers.
    if u in {None, "元", "单", "人", "次", "个"} and abs(p.value) >= 10_000:
        out.append((p.value / 10_000, "base→万"))
    return out


def approx_equal(a: float, b: float, rel_tol: float = 0.01) -> bool:
    """Loose equality used for rounding tolerance.

    Pure ``math.isclose`` rejects rounded reporting figures (Excel
    ``12345.678`` vs. Word ``12346``) when ``abs_tol`` is small, but a flat
    ``abs_tol`` over-matches small percentages, so we only use a relative
    tolerance plus a tiny epsilon for near-zero values.
    """
    if a == b:
        return True
    scale = max(abs(a), abs(b))
    if scale < 1e-9:
        return abs(a - b) < 1e-9
    return abs(a - b) / scale <= rel_tol


def value_match_score(parsed: ParsedNumber, excel_value: float) -> Tuple[float, str]:
    """Score how well ``excel_value`` could be the source of ``parsed``.

    Returns ``(score, interpretation)`` with score in ``[0, 1]``:
      * 1.00 — exact match under some unit interpretation
      * 0.85 — within 0.1% (typical of 2-decimal rounding on large numbers)
      * 0.70 — within 1%   (typical of 1-decimal rounding or % rounding)
      * 0.00 — no plausible interpretation matches
    """
    best_score = 0.0
    best_interp = ""
    for cand_val, interp in candidate_values(parsed):
        if cand_val == excel_value:
            return 1.0, interp
        if approx_equal(cand_val, excel_value, rel_tol=0.001):
            score = 0.85
        elif approx_equal(cand_val, excel_value, rel_tol=0.01):
            score = 0.70
        else:
            score = 0.0
        if score > best_score:
            best_score = score
            best_interp = interp
    return best_score, best_interp


def _to_float(int_part: str, frac_part: Optional[str]) -> Optional[float]:
    s = int_part.replace(",", "")
    if frac_part:
        s += frac_part
    try:
        return float(s)
    except ValueError:
        return None


def _overlaps(start: int, end: int, spans: List[Tuple[int, int]]) -> bool:
    for s, e in spans:
        if start < e and end > s:
            return True
    return False
