"""Deterministic matching of Word numbers to Excel source cells.

The matcher's contract:
  * Generate candidate Excel sources for *every* Word number — never silently
    drop one (CLAUDE.md privacy/coverage rule).
  * Score by both numeric agreement (with unit interpretations and rounding
    tolerance) and label-context overlap, so context disambiguates equal-value
    collisions.
  * Surface ambiguity instead of guessing: a confident pick requires a clear
    winner over the runner-up.

No LLM or fuzzy inference. The :mod:`llm_reranker` module (deferred) is the
only place where heuristic re-ranking should ever live.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import List, Optional, Set

from .excel_profiler import ExcelCell
from .number_normalizer import value_match_score
from .word_profiler import WordNumber

CONFIDENCE_LEVELS = ("HIGH", "MEDIUM", "LOW", "UNRESOLVED", "EXCLUDED")


@dataclass
class Candidate:
    cell: ExcelCell
    value_score: float
    context_score: float
    overlap_tokens: List[str]
    interpretation: str

    @property
    def combined_score(self) -> float:
        # Value match dominates; context lifts a marginal value match and
        # discriminates between equal-value candidates.
        return self.value_score * (0.55 + 0.45 * self.context_score)


@dataclass
class WordMatch:
    word_number: WordNumber
    candidates: List[Candidate]
    confidence: str
    chosen: Optional[Candidate]
    note: str = ""


def match_word_numbers(
    word_numbers: List[WordNumber],
    excel_cells: List[ExcelCell],
) -> List[WordMatch]:
    """Run the full matcher; returns one :class:`WordMatch` per input number."""
    results: List[WordMatch] = []
    for wn in word_numbers:
        if wn.exclusion_reason:
            # Policy-excluded numbers still get a row in the audit trail; the
            # matcher just doesn't spend compute on them and the reviewer
            # sees the reason inline.
            results.append(WordMatch(
                word_number=wn,
                candidates=[],
                confidence="EXCLUDED",
                chosen=None,
                note=wn.exclusion_reason,
            ))
            continue
        word_tokens = _extract_tokens(wn.snippet)
        for ctx in wn.label_context:
            word_tokens |= _extract_tokens(ctx)
        candidates: List[Candidate] = []
        for cell in excel_cells:
            v_score, interp = value_match_score(_to_parsed(wn), cell.numeric_value)
            if v_score == 0:
                continue
            c_score, overlap = _context_score(word_tokens, cell)
            candidates.append(Candidate(
                cell=cell,
                value_score=v_score,
                context_score=c_score,
                overlap_tokens=overlap,
                interpretation=interp,
            ))
        candidates.sort(key=lambda c: c.combined_score, reverse=True)
        confidence, chosen, note = _classify(candidates)
        results.append(WordMatch(
            word_number=wn,
            candidates=candidates[:10],  # cap to keep review artifact readable
            confidence=confidence,
            chosen=chosen,
            note=note,
        ))
    return results


# ---------------------------------------------------------------------------
# Internals
# ---------------------------------------------------------------------------

# Token extraction is intentionally naive: we take all Chinese substrings of
# length 2-6 (so a metric like "营业收入" is found inside "5月营业收入达") and
# uppercase alphanumeric runs (so GMV/ARPU/MAU style acronyms match). This is
# good enough on the synthetic corpus; a future iteration can swap in jieba
# or a domain dictionary.
_CHINESE_RUN_RE = re.compile(r"[一-鿿]+")
_ASCII_TOKEN_RE = re.compile(r"[A-Za-z][A-Za-z0-9_]+")
_MAX_CHINESE_NGRAM = 6


def _extract_tokens(text: str) -> Set[str]:
    tokens: Set[str] = set()
    if not text:
        return tokens
    for run in _CHINESE_RUN_RE.findall(text):
        max_len = min(len(run), _MAX_CHINESE_NGRAM)
        for length in range(2, max_len + 1):
            for start in range(len(run) - length + 1):
                tokens.add(run[start:start + length])
    for tok in _ASCII_TOKEN_RE.findall(text):
        if len(tok) >= 2:
            tokens.add(tok.upper())
    return tokens


def _context_score(word_tokens: Set[str], cell: ExcelCell) -> tuple[float, List[str]]:
    excel_tokens: Set[str] = set()
    for s in cell.row_context:
        excel_tokens |= _extract_tokens(s)
    for s in cell.column_context:
        excel_tokens |= _extract_tokens(s)
    overlap = word_tokens & excel_tokens
    if not overlap:
        return 0.0, []
    ranked = sorted(overlap, key=lambda t: (len(t), t), reverse=True)
    longest = len(ranked[0])
    # A 4-char Chinese metric name overlap is essentially confirmatory;
    # 2-char overlap (e.g. "收入") is suggestive but weak.
    if longest >= 4:
        score = 1.0
    elif longest == 3:
        score = 0.75
    elif longest == 2:
        score = 0.45
    else:
        score = 0.25
    # Multiple distinct overlaps add a small bonus, capped at 1.0.
    if len(ranked) >= 2:
        score = min(1.0, score + 0.1)
    return score, ranked[:5]


def _classify(cands: List[Candidate]) -> tuple[str, Optional[Candidate], str]:
    if not cands:
        return "UNRESOLVED", None, "no excel cell matched any unit interpretation"
    top = cands[0]
    runner = cands[1] if len(cands) > 1 else None
    ambiguous = (
        runner is not None
        and top.value_score == runner.value_score
        and abs(top.combined_score - runner.combined_score) < 0.02
        and top.cell is not runner.cell
    )
    # 0.85 corresponds to "within 0.1% of the Excel figure" — that is the
    # rounding tolerance a 2-decimal 万元 number gives against unrounded yuan.
    # Requiring exact equality (1.00) would push almost every rounded narrative
    # number out of HIGH, defeating the point of the confidence label.
    if top.value_score >= 0.85 and top.context_score >= 0.7 and not ambiguous:
        return "HIGH", top, ""
    if top.value_score >= 0.7 and top.context_score >= 0.3:
        note = "ambiguous: multiple cells tie on value+context" if ambiguous else ""
        return "MEDIUM", top, note
    if top.value_score > 0:
        note = "value match without strong context overlap"
        return "LOW", top, note
    return "UNRESOLVED", None, "no excel cell matched"


# WordNumber and ParsedNumber share enough fields that we can adapt one to the
# other without importing both forms into number_normalizer.
def _to_parsed(wn: WordNumber):
    from .number_normalizer import ParsedNumber
    return ParsedNumber(
        raw=wn.raw, value=wn.value, unit=wn.unit, sign=wn.sign,
        start=wn.offset, end=wn.offset + len(wn.raw),
    )
