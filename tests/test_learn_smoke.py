"""End-to-end smoke: generate-synthetic + learn must produce artifacts and
coverage that match expectations from the synthetic corpus."""
from __future__ import annotations

from pathlib import Path

import openpyxl

from src.excel_profiler import profile_workbook
from src.main import main as cli_main
from src.mapping_reviewer import write_confidence_report, write_mapping_review
from src.synthetic_generator import generate
from src.validator import summarize
from src.value_matcher import match_word_numbers
from src.word_profiler import WordNumber, profile_document


def test_generate_synthetic_writes_both_files(tmp_path: Path):
    xlsx, docx_path = generate(tmp_path)
    assert xlsx.exists() and xlsx.suffix == ".xlsx"
    assert docx_path.exists() and docx_path.suffix == ".docx"
    # Excel should have the three planned sheets.
    wb = openpyxl.load_workbook(str(xlsx))
    assert set(wb.sheetnames) == {"月度核心指标", "周度运营指标", "渠道分析"}
    wb.close()


def test_learn_pipeline_meets_coverage_targets(tmp_path: Path):
    sample_dir = tmp_path / "samples"
    out_dir = tmp_path / "output"
    xlsx, docx_path = generate(sample_dir)

    excel_cells = profile_workbook(xlsx)
    word_numbers = profile_document(docx_path)
    matches = match_word_numbers(word_numbers, excel_cells)
    summary = summarize(matches)

    # Synthetic corpus characteristics we depend on.
    assert len(excel_cells) >= 40, "synthetic Excel should have dozens of numeric cells"
    assert summary.total >= 30, "synthetic Word should have dozens of numeric tokens"

    # The corpus has two deliberate UNRESOLVED cases (0.70 个百分点, 15%).
    assert summary.by_confidence["UNRESOLVED"] >= 2
    # And one MEDIUM case (1.23 亿元 with loose rounding).
    assert summary.by_confidence["MEDIUM"] >= 1
    # The synthetic Word doc has many date/period markers ("5月", "第20周",
    # "2026年", etc.) — they must surface as EXCLUDED in the audit, not vanish.
    assert summary.by_confidence["EXCLUDED"] >= 10
    # Visible accounting: total = eligible + excluded, with no silent losses.
    assert summary.total == summary.eligible + summary.by_confidence["EXCLUDED"]
    # Coverage must dominate over the *eligible* numbers (excluded don't count).
    # Tightening this is fine; relaxing it deserves a code-level reason.
    assert summary.coverage_ratio >= 0.75, summary.by_confidence

    review_path = write_mapping_review(matches, out_dir / "mapping_review.xlsx")
    report_path = write_confidence_report(matches, summary, out_dir / "confidence_report.md")
    assert review_path.exists()
    assert report_path.exists()
    report_text = report_path.read_text(encoding="utf-8")
    assert "UNRESOLVED" in report_text
    assert "EXCLUDED by explicit policy" in report_text


def test_cli_generate_then_learn(tmp_path: Path):
    samples = tmp_path / "samples"
    output = tmp_path / "output"
    assert cli_main(["generate-synthetic", "--out", str(samples)]) == 0
    rc = cli_main([
        "learn",
        "--excel", str(samples / "historical.xlsx"),
        "--word", str(samples / "finished_report.docx"),
        "--out", str(output),
    ])
    assert rc == 0
    assert (output / "mapping_review.xlsx").exists()
    assert (output / "confidence_report.md").exists()
    assert (output / "auto_mapping.yml").exists()
    assert (output / "converted_template.docx").exists()


def test_learn_cli_errors_on_missing_files(tmp_path: Path, capsys):
    rc = cli_main([
        "learn",
        "--excel", str(tmp_path / "nope.xlsx"),
        "--word", str(tmp_path / "nope.docx"),
        "--out", str(tmp_path / "out"),
    ])
    assert rc == 2
    err = capsys.readouterr().err
    assert "not found" in err


def test_unresolved_empty_state_does_not_claim_false_success(tmp_path: Path):
    """If UNRESOLVED is empty but EXCLUDED rows exist, the report's
    empty-state must NOT say "every Word number has at least one candidate
    Excel source" — that would be a false success because EXCLUDED numbers
    were skipped by policy and have no candidates."""
    from src.value_matcher import WordMatch

    eligible_high = WordMatch(
        word_number=WordNumber(
            location="paragraph:1", snippet="收入100元", label_context=[],
            raw="100", value=100.0, unit="元", sign=1, offset=2,
        ),
        candidates=[],
        confidence="HIGH",
        chosen=None,
        note="",
    )
    excluded = WordMatch(
        word_number=WordNumber(
            location="paragraph:0", snippet="2026年5月报告", label_context=[],
            raw="5", value=5.0, unit=None, sign=1, offset=5,
            exclusion_reason="date/period marker (followed by '月')",
        ),
        candidates=[],
        confidence="EXCLUDED",
        chosen=None,
        note="date/period marker (followed by '月')",
    )
    matches = [eligible_high, excluded]
    summary = summarize(matches)
    assert summary.unresolved == []
    assert summary.by_confidence["EXCLUDED"] == 1

    report = write_confidence_report(matches, summary, tmp_path / "report.md")
    text = report.read_text(encoding="utf-8")

    # The false-success phrasing must be gone.
    assert "every Word number has at least one candidate" not in text
    # The honest scoped phrasing must be present.
    assert "every eligible Word number" in text
    # And the empty-state must point the reviewer to the EXCLUDED section.
    assert "EXCLUDED by policy" in text and "see section below" in text
