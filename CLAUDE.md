# CLAUDE.md

## Project Overview

This repository is for an internal Python tool that helps operations users convert BI/FBI report Excel data into weekly or monthly Word reports.

The core product is not a fixed-field mail merge. It is a traceable paired-document learning workflow: given a historical Excel file and its corresponding finished Word report, the system should scan every visible numeric value in the Word document, find candidate Excel sources, and produce reviewable mapping artifacts. Future report generation must be driven by confirmed mappings and deterministic transformations.

The highest product risk is generating a Word report that looks successful while silently omitting, misplacing, or inventing metrics. Optimize early work for coverage, source traceability, and reviewability before formatting polish.

## Boundary With Codex

`AGENTS.md` is intended for Codex only. Claude Code should not read, import, depend on, summarize, or edit `AGENTS.md` unless the user explicitly asks. Use this `CLAUDE.md` file as the implementation source of truth.

## Architecture Notes

- Target language: Python.
- Excel processing: use `openpyxl` so the implementation can preserve sheet names, cell addresses, merged-cell clues, row context, column context, and nearby labels.
- Word processing: use `python-docx` for parsing and later `docxtpl` or equivalent for templating. Do not require Microsoft Word.
- Planned modules: `excel_profiler`, `word_profiler`, `number_normalizer`, `value_matcher`, `llm_reranker`, `mapping_reviewer`, `template_builder`, `renderer`, `validator`, and `synthetic_generator`.
- Learn mode: historical Excel + historical finished Word -> candidate mappings, `mapping_review.xlsx`, `auto_mapping.yml`, `converted_template.docx`, and `confidence_report.md`.
- Run mode: new Excel + `confirmed_mapping.yml` + `converted_template.docx` -> new Word report, validation table, and run log.
- AI/LLM behavior: optional reranking or explanation only. Deterministic matching must generate candidates first. AI must not invent numbers, sources, cells, or mappings.

## Repo Conventions

- The repository is at kickoff stage. Implementation files, package metadata, tests, and docs may not exist yet.
- Prefer a simple `src/` plus `tests/` layout unless the user requests a different structure.
- Use `pathlib` for all paths and keep behavior cross-platform for macOS and Windows.
- Support Chinese filenames, spaces in paths, Chinese sheet names, Chinese metric names, and UTF-8 logs.
- Keep generated outputs under an explicit output directory. Do not mix generated artifacts into source folders unless a test fixture explicitly requires it.
- Keep only synthetic samples in the repo. Real company files belong outside the repo.

## Engineering Rules

- If the `karpathy-guidelines` skill is available, use it for coding, review, and refactor work: keep changes simple, surgical, assumption-aware, and verifiable.
- Prefer deterministic logic over AI guesses, especially for numeric matching and production generation.
- Never silently ignore a visible Word number. It must be mapped, low-confidence, excluded by an explicit policy, or unresolved.
- Do not build around a short manually maintained metric dictionary. Tests should include dozens of synthetic metrics and edge cases.
- Normalize business number formats before matching: yuan vs ten-thousand yuan, counts vs ten-thousand counts, decimals vs percentages, comma-separated values, Chinese units, signs, and rounding tolerance.
- Make ambiguous matches visible to humans instead of overconfidently choosing one.
- Add concise comments around non-obvious parsing, normalization, or matching logic.
- Mark unresolved product choices as TODOs or blockers instead of inventing hidden behavior.

## Privacy Rules

- Do not commit real company Excel files, Word reports, screenshots, copied business content, or logs containing sensitive values.
- Repository examples must be fake but structurally realistic.
- Reports should prefer aggregate validation metrics and representative synthetic snippets over raw business content.
- If a command touches real data outside the repo, keep outputs bounded and redact content unless the user explicitly asks otherwise.

## Verification Commands

Exact commands may evolve after the scaffold exists. Expected verification should include:

- `python -m pytest`
- `python -m src.main generate-synthetic`
- `python -m src.main learn --excel samples/synthetic/historical.xlsx --word samples/synthetic/finished_report.docx --out output`

The learn-mode smoke should report total Word numbers, mapped numbers, unresolved numbers, low-confidence numbers, and output artifact paths.

## Reporting Expectations

When finishing work, report:

- files changed;
- commands run and whether they passed;
- synthetic learn-mode coverage summary when relevant;
- unresolved blockers or TODOs;
- the next narrow recommended step.

Do not claim full product readiness when only scaffold, synthetic tests, or learn-mode prototypes exist.

## Roadmap Reference

This roadmap is background only. The current task must come from the user's prompt, not this section.

- Synthetic data generator for paired Excel/Word samples with many fake metrics and edge cases.
- Learn-mode extraction, normalization, matching, review artifact generation, and fail-loud validation.
- Human-confirmed mapping workflow and template conversion.
- Deterministic run-mode rendering from confirmed mappings.
- Later packaging, GUI, or business-specific workflow polish after the traceability loop is proven.
