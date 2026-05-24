"""CLI entry point: ``python -m src.main <subcommand>``.

Only two subcommands exist for this first traceability milestone:
  * ``generate-synthetic`` — emit a paired fake Excel + Word sample so the
    learn pipeline can be exercised without real company data.
  * ``learn`` — profile, match, and produce reviewable artifacts.

Future commands (``run``, ``confirm-mapping``, etc.) are intentionally absent.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import List, Optional

from .excel_profiler import profile_workbook
from .mapping_reviewer import write_confidence_report, write_mapping_review
from .synthetic_generator import generate as generate_synthetic
from .validator import format_console_summary, summarize
from .value_matcher import match_word_numbers
from .word_profiler import profile_document


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="magic_excel2word",
        description="Traceable Excel→Word learn-mode prototype.",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    gen = sub.add_parser(
        "generate-synthetic",
        help="Write paired synthetic Excel + Word sample files",
    )
    gen.add_argument(
        "--out",
        type=Path,
        default=Path("samples/synthetic"),
        help="Directory to write historical.xlsx and finished_report.docx",
    )

    learn = sub.add_parser(
        "learn",
        help="Profile an Excel + Word pair and emit mapping_review.xlsx + confidence_report.md",
    )
    learn.add_argument("--excel", type=Path, required=True, help="Source Excel workbook (.xlsx)")
    learn.add_argument("--word", type=Path, required=True, help="Finished Word report (.docx)")
    learn.add_argument("--out", type=Path, required=True, help="Output directory for review artifacts")

    return parser


def main(argv: Optional[List[str]] = None) -> int:
    args = build_parser().parse_args(argv)

    if args.command == "generate-synthetic":
        xlsx, docx_path = generate_synthetic(args.out)
        print("Wrote synthetic sample:")
        print(f"  excel: {xlsx}")
        print(f"  word : {docx_path}")
        return 0

    if args.command == "learn":
        if not args.excel.exists():
            print(f"error: excel file not found: {args.excel}", file=sys.stderr)
            return 2
        if not args.word.exists():
            print(f"error: word file not found: {args.word}", file=sys.stderr)
            return 2

        excel_cells = profile_workbook(args.excel)
        word_numbers = profile_document(args.word)
        matches = match_word_numbers(word_numbers, excel_cells)
        summary = summarize(matches)

        args.out.mkdir(parents=True, exist_ok=True)
        review_path = write_mapping_review(matches, args.out / "mapping_review.xlsx")
        report_path = write_confidence_report(matches, summary, args.out / "confidence_report.md")

        print(format_console_summary(summary))
        print()
        print(f"Excel cells profiled : {len(excel_cells)}")
        print(f"Word numbers profiled: {len(word_numbers)}")
        print(f"Mapping review       : {review_path}")
        print(f"Confidence report    : {report_path}")
        return 0

    return 1


if __name__ == "__main__":
    sys.exit(main())
