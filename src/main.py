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

from .artifact_validator import format_validation_summary, validate_artifacts
from .excel_profiler import profile_workbook
from .mapping_reviewer import write_confidence_report, write_mapping_review
from .synthetic_generator import generate as generate_synthetic
from .template_builder import assign_word_ids, write_template_artifacts
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
    learn.add_argument(
        "--strict",
        action="store_true",
        help=(
            "Exit non-zero if any eligible Word number is UNRESOLVED or LOW. "
            "Use this as a trust gate before any real-file or production step; "
            "default (exploratory) mode still writes artifacts and only warns."
        ),
    )

    validate = sub.add_parser(
        "validate-artifacts",
        help=(
            "Read-only cross-check of the four learn-mode artifacts under "
            "--out: verify mapping_review.xlsx, auto_mapping.yml, "
            "converted_template.docx, and confidence_report.md tell the same "
            "story about every Word number."
        ),
    )
    validate.add_argument(
        "--out",
        type=Path,
        required=True,
        help="Output directory previously written by `learn`",
    )

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
        # Template artifacts run first so we can pass the resulting word_ids
        # and per-match placeholder status into the XLSX writer — the XLSX
        # and YAML must agree on every join key, and the validator depends
        # on that single source of truth.
        template_artifacts = write_template_artifacts(matches, args.word, args.out)
        word_ids = assign_word_ids(matches)
        review_path = write_mapping_review(
            matches,
            args.out / "mapping_review.xlsx",
            word_ids,
            template_artifacts.placeholder_status,
        )
        report_path = write_confidence_report(matches, summary, args.out / "confidence_report.md")

        print(format_console_summary(summary))
        print()
        print(f"Excel cells profiled : {len(excel_cells)}")
        print(f"Word numbers profiled: {len(word_numbers)}")
        print(f"Mapping review       : {review_path}")
        print(f"Confidence report    : {report_path}")
        print(f"Auto mapping (YAML)  : {template_artifacts.yaml_path}")
        print(
            f"Converted template   : {template_artifacts.docx_path} "
            f"({template_artifacts.placeholders_applied} placeholders applied)"
        )

        # Trust gate: artifacts are always written so a reviewer can inspect
        # what failed, but the exit code (or stderr warning) tells automation
        # and humans whether this run is safe to act on.
        failures = summary.strict_failures
        if failures > 0:
            unresolved = summary.by_confidence.get("UNRESOLVED", 0)
            low = summary.by_confidence.get("LOW", 0)
            detail = f"{failures} eligible Word number(s) are UNRESOLVED or LOW (UNRESOLVED={unresolved}, LOW={low})"
            if args.strict:
                print(
                    "\n" + "!" * 72,
                    f"STRICT GATE FAILED: {detail}.",
                    f"Review {review_path} and {report_path} before any real-file pilot.",
                    "!" * 72,
                    sep="\n",
                    file=sys.stderr,
                )
                return 3
            print(
                "\n" + "!" * 72,
                f"WARNING: {detail}.",
                "Artifacts were still written for exploratory review.",
                "Re-run with --strict to fail loudly before any real-file pilot.",
                "!" * 72,
                sep="\n",
                file=sys.stderr,
            )
        return 0

    if args.command == "validate-artifacts":
        if not args.out.exists():
            print(f"error: output directory not found: {args.out}", file=sys.stderr)
            return 2
        report = validate_artifacts(args.out)
        print(format_validation_summary(report, args.out))
        # 4 keeps this distinct from 2 (input missing) and 3 (strict gate),
        # so automation can tell *which* learn-mode contract failed.
        return 0 if report.ok else 4

    return 1


if __name__ == "__main__":
    sys.exit(main())
