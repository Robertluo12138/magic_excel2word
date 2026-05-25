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
  * ``render-docx`` — deterministic Word renderer. Substitutes
    ``{{ word_NNNN }}`` placeholders in ``converted_template.docx``
    using ``run_validation.xlsx`` and writes ``new_report.docx`` plus
    ``render_log.yml``. Fails loudly when any validation row is non-ok,
    when the template references a word_id absent from validation,
    when validation has an unused word_id, or when display formatting
    cannot be safely inferred from the historical raw token.
  * ``validate-render`` — read-only cross-check of the three rendered
    outputs (``new_report.docx``, ``render_log.yml``,
    ``run_validation.xlsx``). Fails loudly if the rendered docx still
    carries placeholders, if the log and the validation disagree on
    which word_ids were rendered, if any row is non-ok, if any log
    entry is missing source/display fields, or if placeholder occurrences
    were silently zero. Does not re-render, does not mutate artifacts.
  * ``pilot-summary`` — read-only redacted summary of a pilot output
    directory. Reports artifact presence, aggregate counts (from each
    artifact's own ``summary`` block / ``Status`` column), and a
    next-action hint per stage. Never prints raw Word tokens,
    generated values, source sheet/cell content, file paths beyond
    basenames, or individual ``word_id`` values. Safe to copy-paste
    into a chat/ticket without leaking real-data details. Fails when
    the required minimum (``auto_mapping.yml``) is absent.
  * ``pilot-preflight`` — read-only path/metadata preflight for a
    real-file pilot. Verifies the four pilot paths exist with the
    expected ``.xlsx`` / ``.docx`` suffix, ``--out`` is a directory
    (or can be created later), and **no** input/output path resolves
    inside the repo tree. Never opens any document, never mutates a
    file. Prints only flag labels and path basenames so the output
    is safe to paste.
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
from .pilot_preflight import (
    format_report as format_pilot_preflight_report,
    preflight as run_pilot_preflight,
)
from .pilot_summary import format_summary as format_pilot_summary, summarize_pilot
from .preflight import emit_advisory as emit_preflight_advisory
from .render_validator import (
    format_validation_summary as format_render_validation_summary,
    validate_render,
)
from .renderer import (
    format_console_summary as format_render_summary,
    render_docx,
)
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

    render = sub.add_parser(
        "render-docx",
        help=(
            "Substitute {{ word_NNNN }} placeholders in "
            "converted_template.docx using run_validation.xlsx and write "
            "the rendered Word report plus render_log.yml. Fails loudly "
            "when any validation row is non-ok, when the template "
            "references a word_id absent from validation, when validation "
            "has an unused word_id, or when display text cannot be safely "
            "inferred from the historical raw token. Does NOT use an LLM, "
            "GUI, network, or Microsoft Word automation."
        ),
    )
    render.add_argument(
        "--template",
        type=Path,
        required=True,
        help="converted_template.docx written by `learn`",
    )
    render.add_argument(
        "--run-validation",
        type=Path,
        required=True,
        dest="run_validation",
        help="run_validation.xlsx written by `run-preview`",
    )
    render.add_argument(
        "--out",
        type=Path,
        required=True,
        help=(
            "Path to write the rendered Word report (.docx). "
            "render_log.yml is written alongside in the same directory."
        ),
    )

    vr = sub.add_parser(
        "validate-render",
        help=(
            "Read-only cross-check of the three rendered-output artifacts: "
            "verify the rendered docx contains no {{ word_NNNN }} "
            "placeholders, render_log.yml and run_validation.xlsx agree "
            "one-to-one on every word_id, every row is ok, every log entry "
            "carries source/display fields, and placeholder_occurrences "
            ">= 1. Never re-renders, never mutates artifacts."
        ),
    )
    vr.add_argument(
        "--docx",
        type=Path,
        required=True,
        help="Rendered Word report (.docx) written by `render-docx`",
    )
    vr.add_argument(
        "--render-log",
        type=Path,
        required=True,
        dest="render_log",
        help="render_log.yml written alongside the rendered docx",
    )
    vr.add_argument(
        "--run-validation",
        type=Path,
        required=True,
        dest="run_validation",
        help="run_validation.xlsx written by `run-preview`",
    )

    ps = sub.add_parser(
        "pilot-summary",
        help=(
            "Read-only redacted summary of a pilot output directory. "
            "Reports artifact presence, aggregate counts, and per-stage "
            "next-action hints. Never prints raw Word tokens, generated "
            "values, source sheet/cell content, file paths beyond "
            "basenames, or individual word_ids. Safe to copy-paste."
        ),
    )
    ps.add_argument(
        "--out",
        type=Path,
        required=True,
        help="Pilot output directory previously written by `learn` (and later stages)",
    )

    pp = sub.add_parser(
        "pilot-preflight",
        help=(
            "Read-only path/metadata preflight for a real-file pilot. "
            "Verifies the four pilot paths exist with the expected suffix, "
            "--out is a directory (or can be created later), and no path "
            "resolves inside this repo tree. Never opens any document, "
            "never mutates a file. Prints only flag labels and basenames."
        ),
    )
    pp.add_argument(
        "--historical-excel",
        type=Path,
        required=True,
        dest="historical_excel",
        help="Historical-period Excel workbook (.xlsx)",
    )
    pp.add_argument(
        "--historical-word",
        type=Path,
        required=True,
        dest="historical_word",
        help="Historical-period finished Word report (.docx)",
    )
    pp.add_argument(
        "--new-excel",
        type=Path,
        required=True,
        dest="new_excel",
        help="New-period Excel workbook to render against (.xlsx)",
    )
    pp.add_argument(
        "--out",
        type=Path,
        required=True,
        help="Output directory for pilot artifacts (must be outside the repo)",
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
        # Advisory runs first so a typo'd path inside samples/ is flagged
        # even when the existence check below short-circuits with rc=2.
        # The advisory is purely lexical (no file I/O) and tolerates
        # paths that do not exist yet.
        emit_preflight_advisory(
            [("--excel", args.excel), ("--word", args.word), ("--out", args.out)],
            sys.stderr,
        )
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
        # Advisory before existence check; see learn for rationale.
        emit_preflight_advisory([("--out", args.out)], sys.stderr)
        if not args.out.exists():
            print(f"error: output directory not found: {args.out}", file=sys.stderr)
            return 2

        report = validate_artifacts(args.out)
        print(format_validation_summary(report, args.out))
        # 4 keeps this distinct from 2 (input missing) and 3 (strict gate),
        # so automation can tell *which* learn-mode contract failed.
        return 0 if report.ok else 4

    if args.command == "confirm-mapping":
        # Advisory before existence checks; see learn for rationale.
        emit_preflight_advisory(
            [
                ("--auto", args.auto),
                ("--review", args.review),
                ("--out", args.out),
            ],
            sys.stderr,
        )
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
        # Advisory before existence checks; see learn for rationale.
        emit_preflight_advisory(
            [
                ("--excel", args.excel),
                ("--confirmed", args.confirmed),
                ("--out", args.out),
            ],
            sys.stderr,
        )
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

    if args.command == "render-docx":
        # Advisory before existence checks; see learn for rationale.
        emit_preflight_advisory(
            [
                ("--template", args.template),
                ("--run-validation", args.run_validation),
                ("--out", args.out),
            ],
            sys.stderr,
        )
        if not args.template.exists():
            print(f"error: template file not found: {args.template}", file=sys.stderr)
            return 2
        if not args.run_validation.exists():
            print(
                f"error: run-validation file not found: {args.run_validation}",
                file=sys.stderr,
            )
            return 2

        render_report = render_docx(
            template_path=args.template,
            run_validation_path=args.run_validation,
            out_docx_path=args.out,
        )
        if render_report.fatal_errors:
            # 8 keeps this distinct from 2 (missing inputs), 3 (strict
            # gate), 4 (validate-artifacts), 5 (incomplete confirm),
            # 6/7 (run-preview) so automation can tell which gate held.
            # A fatal error means the inputs themselves are unusable —
            # no docx or log is written.
            print(format_render_summary(render_report), file=sys.stderr)
            return 8
        if render_report.failures:
            # 9 = per-row gate failure (non-ok validation, missing or
            # duplicated generated value, orphan validation row,
            # un-inferable format). No partial docx is written: CLAUDE.md
            # forbids producing a Word report that silently omits or
            # mis-renders a confirmed metric.
            print(format_render_summary(render_report), file=sys.stderr)
            return 9
        print(format_render_summary(render_report))
        return 0

    if args.command == "pilot-summary":
        # Advisory before existence check; see learn for rationale.
        emit_preflight_advisory([("--out", args.out)], sys.stderr)
        if not args.out.exists():
            print(
                f"error: output directory not found: {args.out}",
                file=sys.stderr,
            )
            return 2
        if not args.out.is_dir():
            print(
                f"error: --out must be a directory: {args.out}",
                file=sys.stderr,
            )
            return 2

        summary = summarize_pilot(args.out)
        text = format_pilot_summary(summary)
        # 11 keeps this distinct from 2/3/4/5/6/7/8/9/10 so automation can
        # tell which gate held. A fatal here means the minimum learn-mode
        # artifact is absent — nothing meaningful to summarize.
        if summary.fatal_errors:
            print(text, file=sys.stderr)
            return 11
        print(text)
        return 0

    if args.command == "pilot-preflight":
        # Read-only by contract: never opens any document, never mutates
        # a file. No privacy advisory is wired here because the command's
        # own inside-repo gate is strictly stronger — it refuses any path
        # under the repo tree (exit 12), not just paths under samples/.
        report = run_pilot_preflight(
            historical_excel=args.historical_excel,
            historical_word=args.historical_word,
            new_excel=args.new_excel,
            out=args.out,
        )
        text = format_pilot_preflight_report(report)
        if report.ok:
            print(text, end="")
            return 0
        print(text, end="", file=sys.stderr)
        # 12 keeps the privacy refusal distinct from 2 (generic input
        # error) so automation can branch on it: a 12 means a pilot
        # path landed inside the repo tree (refuse, do NOT retry with
        # the same path), while 2 means a missing or malformed input
        # (operator typo, fixable in place).
        if report.inside_repo_count:
            return 12
        return 2

    if args.command == "validate-render":
        # Advisory before existence checks; see learn for rationale.
        emit_preflight_advisory(
            [
                ("--docx", args.docx),
                ("--render-log", args.render_log),
                ("--run-validation", args.run_validation),
            ],
            sys.stderr,
        )
        for label, path in (
            ("--docx", args.docx),
            ("--render-log", args.render_log),
            ("--run-validation", args.run_validation),
        ):
            if not path.exists():
                print(f"error: {label} file not found: {path}", file=sys.stderr)
                return 2

        vr_report = validate_render(
            docx_path=args.docx,
            render_log_path=args.render_log,
            run_validation_path=args.run_validation,
        )
        # 10 keeps this distinct from 2 (missing inputs), 3 (strict gate),
        # 4 (validate-artifacts), 5 (incomplete confirm), 6/7 (run-preview),
        # 8/9 (render-docx) so automation can tell which gate held.
        if vr_report.ok:
            print(format_render_validation_summary(vr_report))
            return 0
        print(format_render_validation_summary(vr_report), file=sys.stderr)
        return 10

    return 1


if __name__ == "__main__":
    sys.exit(main())
