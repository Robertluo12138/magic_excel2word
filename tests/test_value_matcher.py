"""Tests for the matcher's confidence classification.

The matcher carries the highest product risk: a wrong HIGH label can quietly
ship a wrong number. These tests pin down the three behaviors that protect
against that — context disambiguates equal-value cells, loose rounding lands
in MEDIUM not HIGH, and no candidates surface as UNRESOLVED rather than
silently disappearing.
"""
from __future__ import annotations

from src.excel_profiler import ExcelCell
from src.value_matcher import match_word_numbers
from src.word_profiler import WordNumber


def _cell(value, row_ctx, col_ctx, sheet="S", addr="A1"):
    return ExcelCell(
        sheet=sheet, address=addr, row=1, column=1,
        raw_value=value, numeric_value=float(value),
        row_context=list(row_ctx), column_context=list(col_ctx),
    )


def _word(raw, value, unit, snippet, label_context=()):
    return WordNumber(
        location="paragraph:0", snippet=snippet, label_context=list(label_context),
        raw=raw, value=value, unit=unit, sign=1, offset=0,
    )


def test_unique_value_with_strong_context_is_high():
    wn = _word("12,345.68万元", 12345.68, "万元", "5月营业收入达12,345.68万元")
    cells = [_cell(123_456_789, ["营业收入(元)"], ["2026年5月", "指标"])]
    [m] = match_word_numbers([wn], cells)
    assert m.confidence == "HIGH"
    assert m.chosen.cell.numeric_value == 123_456_789


def test_context_disambiguates_equal_values():
    # Three Excel cells share value 234,567; only one has matching labels.
    wn = _word("23.46万单", 23.46, "万单",
               "第18周日均订单数为23.46万单",
               label_context=[])
    cells = [
        _cell(234_567, ["新增用户"], ["2026年4月"], sheet="月度", addr="B6"),
        _cell(234_567, ["日均订单数"], ["第18周"], sheet="周度", addr="B2"),
        _cell(234_567, ["其他"], ["用户数"], sheet="渠道", addr="B6"),
    ]
    [m] = match_word_numbers([wn], cells)
    assert m.chosen.cell.sheet == "周度"
    assert m.confidence in {"HIGH", "MEDIUM"}


def test_unresolved_when_no_value_matches():
    wn = _word("15%", 15, "%", "运营效率提升15%")
    cells = [_cell(99_999, ["其他指标"], ["列1"])]
    [m] = match_word_numbers([wn], cells)
    assert m.confidence == "UNRESOLVED"
    assert m.chosen is None
    assert m.candidates == []


def test_loose_rounding_lands_in_medium_not_high():
    # 1.23 亿元 expanded = 123,000,000 vs Excel 123,456,789 — ~0.37% gap.
    wn = _word("1.23亿元", 1.23, "亿元",
               "本月营业收入约为1.23亿元",
               label_context=[])
    cells = [_cell(123_456_789, ["营业收入(元)"], ["2026年5月"])]
    [m] = match_word_numbers([wn], cells)
    assert m.confidence == "MEDIUM"


def test_ambiguous_pick_carries_note():
    # Two cells tie on value AND share equal context strength → ambiguity note.
    wn = _word("25.00%", 25.00, "%", "同比增长25.00%")
    cells = [
        _cell(25.00, ["营业收入(元)"], ["同比(%)"], sheet="月度", addr="D2"),
        _cell(25.00, ["净利润(元)"], ["同比(%)"], sheet="月度", addr="D3"),
    ]
    [m] = match_word_numbers([wn], cells)
    # Only "同比" matches as a token (length 2) in both — same context score.
    assert m.confidence == "MEDIUM"
    assert "ambiguous" in m.note


def test_every_word_number_produces_a_match_record():
    # Coverage rule: never silently drop a Word number.
    wns = [
        _word("100", 100, None, "abc 100 xyz"),
        _word("200", 200, None, "abc 200 xyz"),
        _word("300", 300, None, "abc 300 xyz"),
    ]
    out = match_word_numbers(wns, [])
    assert len(out) == 3
    assert all(m.confidence == "UNRESOLVED" for m in out)


def test_excluded_numbers_are_passed_through_not_matched():
    # A policy-excluded WordNumber must surface as EXCLUDED with its reason —
    # never get matched against the Excel pool. Even if its value happens to
    # match an Excel cell, the matcher must not propose it.
    excluded = _word("5", 5, None, "5月营业收入达12,345.68万元")
    excluded.exclusion_reason = "date/period marker (followed by '月')"
    cells = [_cell(5.0, ["something"], ["another"])]
    [m] = match_word_numbers([excluded], cells)
    assert m.confidence == "EXCLUDED"
    assert m.chosen is None
    assert m.candidates == []
    assert "date" in m.note
