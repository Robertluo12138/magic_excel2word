"""CLI entry point: ``python -m src.main <subcommand>``.

Subcommands for the learn-mode trust loop:
  * ``generate-synthetic`` — emit a paired fake Excel + Word sample so the
    learn pipeline can be exercised without real company data.
  * ``learn`` — profile, match, and produce reviewable artifacts.
  * ``validate-artifacts`` — cross-check the four learn-mode artifacts.
  * ``confirm-mapping`` — read reviewer decisions from
    ``mapping_review.xlsx`` and promote confirmed rows into
    ``confirmed_mapping.yml``. Blank/reject/LOW/UNRESOLVED/invalid rows
    stay visible as ``review_required`` and the run fails unless
    ``--allow-incomplete`` is passed.
  * ``run-preview`` — narrow run-mode preview. Resolves each confirmed
    mapping against a NEW Excel workbook and writes a per-row run
    validation artifact. **Does not render a Word document.**

Future commands (full ``run`` for production Word rendering) are
intentionally absent — `run-preview` is the bridge that proves the
confirmed mappings still resolve cleanly against a new period.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import List, Optional

from .artifact_validator import format_validation_summary, validate_artifacts
from .excel_profiler import profile_workbook
from .mapping_confirmer import (
    confirm_mappings,
    format_console_summary as format_confirm_summary,
    write_confirmed_yaml,
)
from .mapping_reviewer import write_confidence_report, write_mapping_review
from .run_preview import (
    format_console_summary as format_preview_summary,
    run_preview,
    write_run_validation,
)
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

    confirm = sub.add_parser(
        "confirm-mapping",
        help=(
            "Promote reviewer-confirmed rows from mapping_review.xlsx + "
            "auto_mapping.yml into confirmed_mapping.yml. Blank decisions, "
            "rejections, LOW/UNRESOLVED rows, and invalid overrides stay "
            "visible as review_required; the command fails unless "
            "--allow-incomplete is set."
        ),
    )
    confirm.add_argument(
        "--auto",
        type=Path,
        required=True,
        help="Path to auto_mapping.yml (written by `learn`)",
    )
    confirm.add_argument(
        "--review",
        type=Path,
        required=True,
        help="Path to mapping_review.xlsx with reviewer decisions filled in",
    )
    confirm.add_argument(
        "--out",
        type=Path,
        required=True,
        help="Path to write confirmed_mapping.yml",
    )
    confirm.add_argument(
        "--allow-incomplete",
        action="store_true",
        help=(
            "Exploratory escape hatch: write confirmed_mapping.yml and exit 0 "
            "even when review_required is non-empty. Default behaviour is to "
            "fail (exit 5) so unreviewed rows can't sneak past the gate."
        ),
    )

    preview = sub.add_parser(
        "run-preview",
        help=(
            "Resolve confirmed_mapping.yml against a NEW Excel workbook and "
            "write run_validation.xlsx — one row per confirmed word_id with "
            "raw Excel value, generated value, transform, confidence, and "
            "per-row status. Fails loudly on incomplete confirmed mapping, "
            "missing source sheet/cell, non-numeric cells, or unknown "
            "transforms. Does NOT render a Word document."
        ),
    )
    preview.add_argument(
        "--excel",
        type=Path,
        required=True,
        help="New-period Excel workbook to extract values from (.xlsx)",
    )
    preview.add_argument(
        "--confirmed",
        type=Path,
        required=True,
        help="confirmed_mapping.yml produced by `confirm-mapping`",
    )
    preview.add_argument(
        "--out",
        type=Path,
        required=True,
        help="Output directory for run_validation.xlsx",
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

    if args.command == "confirm-mapping":
        if not args.auto.exists():
            print(f"error: auto mapping not found: {args.auto}", file=sys.stderr)
            return 2
        if not args.review.exists():
            print(f"error: review xlsx not found: {args.review}", file=sys.stderr)
            return 2

        confirm_report = confirm_mappings(args.auto, args.review)
        if confirm_report.fatal_errors:
            # Fatal errors mean the inputs themselves are unusable. Surface
            # the reasons and refuse to write a confirmed mapping so a
            # human reviewer can fix the artifacts before retrying.
            print(
                format_confirm_summary(confirm_report, None, args.allow_incomplete),
                file=sys.stderr,
            )
            return 2

        # ``total_word_numbers`` is read back from auto_mapping.yml's
        # summary so the confirmed file can be cross-checked against the
        # learn run without re-parsing the original Word doc.
        import yaml as _yaml  # local: keeps top-level imports minimal

        auto_doc = _yaml.safe_load(args.auto.read_text(encoding="utf-8")) or {}
        total = int((auto_doc.get("summary") or {}).get("total") or 0)

        out_path = write_confirmed_yaml(
            confirm_report,
            args.auto,
            args.review,
            args.out,
            allow_incomplete=args.allow_incomplete,
            total_word_numbers=total,
        )
        print(format_confirm_summary(confirm_report, out_path, args.allow_incomplete))

        # 5 keeps this distinct from 2 (missing input), 3 (strict gate),
        # 4 (validate-artifacts) so automation can tell which gate held.
        if confirm_report.review_required and not args.allow_incomplete:
            return 5
        return 0

    if args.command == "run-preview":
        if not args.excel.exists():
            print(f"error: excel file not found: {args.excel}", file=sys.stderr)
            return 2
        if not args.confirmed.exists():
            print(f"error: confirmed mapping not found: {args.confirmed}", file=sys.stderr)
            return 2

        preview_report = run_preview(args.excel, args.confirmed)
        if preview_report.fatal_errors:
            # Don't write an artifact when the confirmed_mapping.yml is
            # unusable — there's nothing meaningful to populate.
            print(format_preview_summary(preview_report, None), file=sys.stderr)
            # 6 keeps this distinct from 2 (missing inputs), 3 (strict gate),
            # 4 (validate-artifacts), 5 (incomplete confirm) so automation
            # can tell which gate held.
            return 6

        out_path = write_run_validation(preview_report, args.out)
        # Per-row failures get printed to stdout (artifact is written) so
        # a reviewer can inspect every row, but the exit code is 7 so
        # automation halts before any downstream rendering step.
        if preview_report.failures:
            print(format_preview_summary(preview_report, out_path), file=sys.stderr)
            return 7
        print(format_preview_summary(preview_report, out_path))
        return 0

    return 1


if __name__ == "__main__":
    sys.exit(main())
