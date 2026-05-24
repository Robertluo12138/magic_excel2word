"""Unit tests for parsing and candidate-value generation."""
from __future__ import annotations

from src.number_normalizer import (
    ParsedNumber,
    approx_equal,
    candidate_values,
    find_numbers,
    value_match_score,
)


def _values_only(text):
    return [(p.value, p.unit, p.raw) for p in find_numbers(text)]


def test_finds_plain_decimal_with_commas():
    [p] = find_numbers("销售额为 12,345.67 元。")
    assert p.value == 12345.67
    assert p.unit == "元"


def test_finds_chinese_unit_suffix():
    [p] = find_numbers("营业收入12,345.68万元")
    assert p.value == 12345.68
    assert p.unit == "万元"


def test_handles_percentage_and_permille():
    out = _values_only("同比25.00%，环比1.5‰。")
    assert (25.00, "%", "25.00%") in out
    assert (1.5, "‰", "1.5‰") in out


def test_accounting_parens_become_negative():
    [p] = find_numbers("亏损(1,234.50)元")
    assert p.value == -1234.50
    assert p.sign == -1


def test_marks_date_period_markers_as_excluded():
    # 2026, 5, 18 are date/period markers — captured for audit with an
    # exclusion_reason; 1,234.56 万元 is the real metric and has none.
    out = find_numbers("2026年5月第18周收入1,234.56万元")
    by_raw = {p.raw: p.exclusion_reason for p in out}
    assert set(by_raw) == {"2026", "5", "18", "1,234.56万元"}
    assert by_raw["1,234.56万元"] is None
    assert "date" in by_raw["2026"]
    assert "'年'" in by_raw["2026"]
    assert "'月'" in by_raw["5"]
    assert "'周'" in by_raw["18"]


def test_yi_unit_expands_to_yuan():
    [p] = find_numbers("收入约1.23亿元")
    cands = dict((interp, val) for val, interp in candidate_values(p))
    assert cands["亿元→base_unit"] == 123_000_000


def test_wan_unit_expands_to_base():
    [p] = find_numbers("用户1,345.68万人")
    cands = dict((interp, val) for val, interp in candidate_values(p))
    assert cands["万人→base_unit"] == 13_456_800


def test_percentage_offers_decimal_alternative():
    [p] = find_numbers("毛利率36.20%")
    cands = dict((interp, val) for val, interp in candidate_values(p))
    assert abs(cands["%→decimal"] - 0.362) < 1e-9


def test_base_to_wan_only_for_large_unitless():
    small = ParsedNumber(raw="3", value=3, unit=None, sign=1, start=0, end=1)
    large = ParsedNumber(raw="123456", value=123456, unit=None, sign=1, start=0, end=6)
    assert all(interp != "base→万" for _, interp in candidate_values(small))
    assert any(interp == "base→万" and val == 12.3456 for val, interp in candidate_values(large))


def test_approx_equal_rounding_tolerance():
    # Word rounds 123,456,789 to 12,345.68 万元 → 123,456,800
    assert approx_equal(123_456_800, 123_456_789, rel_tol=0.001)
    # 1.23 亿元 → 123,000,000 vs 123,456,789 is ~0.37% off, in 1% bucket only
    assert not approx_equal(123_000_000, 123_456_789, rel_tol=0.001)
    assert approx_equal(123_000_000, 123_456_789, rel_tol=0.01)


def test_value_match_score_buckets():
    p = ParsedNumber(raw="12,345.68万元", value=12345.68, unit="万元", sign=1, start=0, end=11)
    score_exact, interp_exact = value_match_score(p, 123_456_800.0)
    score_close, _ = value_match_score(p, 123_456_789.0)
    score_loose, _ = value_match_score(p, 124_000_000.0)
    score_none, _ = value_match_score(p, 999_999.0)
    assert score_exact == 1.0
    assert interp_exact == "万元→base_unit"
    assert score_close == 0.85
    assert score_loose == 0.70
    assert score_none == 0.0


def test_finds_numbers_inside_table_cell_text():
    out = _values_only("345.68万人")
    assert out == [(345.68, "万人", "345.68万人")]
