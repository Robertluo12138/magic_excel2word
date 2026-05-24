# AGENTS.md

## Project North Star

本项目是一个面向运营同学的内部报表自动化工具：把从 BI/FBI 报表复制出来的 Excel 数据，可靠地转换成周报/月报 Word 文档。

核心不是固定字段的邮件合并，而是“历史成品学习”：给定一份历史 Excel 和对应的成熟 Word 成品，系统应扫描 Word 中所有可见数字，尝试追溯每个数字在 Excel 中的来源，并产出可人工复核的映射表。早期最重要的是可信、透明、可审计，而不是输出格式漂亮。

## Stable Product Boundaries

- 不依赖小型人工维护指标字典作为主要真相来源；真实报表可能有几十个指标。
- Word 中每个可见数字都必须被扫描、匹配、记录或进入 unresolved list；不允许静默跳过。
- AI/LLM 只能作为候选来源 reranker 或解释助手，不能发明数字、来源、单元格或最终口径。
- 首期不做 GUI、exe 打包、生产级 Word 渲染、真实业务数据样例、网络依赖或平台专属流程。
- 不依赖本机安装 Microsoft Excel 或 Word；目标是 macOS 和 Windows 都能运行。
- 仓库只能放合成样例；真实公司数据、真实报表内容、敏感指标明细不得提交。

## End-to-End Workflow

1. Learn mode: 历史 Excel + 历史 Word 成品 -> Excel 数字单元格画像、Word 可见数字画像、候选来源匹配、`mapping_review.xlsx`、`auto_mapping.yml`、`converted_template.docx`、`confidence_report.md`。
2. Human review: 人工确认、拒绝或改选候选来源；低置信和未解决数字必须清楚呈现。
3. Run mode: 新 Excel + `confirmed_mapping.yml` + `converted_template.docx` -> 新 Word 报告、字段来源验证表、运行日志。
4. Production discipline: 生产生成只使用确认过的映射和确定性转换；AI 不能自由猜测业务数字。

## Source of Truth Map

- Stable Codex steering: `AGENTS.md`.
- Claude Code implementation instructions: `CLAUDE.md`.
- Project README / user guide: TODO / not created yet.
- Architecture reference: TODO / not created yet.
- Mapping schema and review contract: TODO / not created yet.
- Synthetic sample contract: TODO / not created yet.
- Tests and validation commands: TODO / not created yet.

## Roadmap Reference

This roadmap is background only and does not define the current task.

- Synthetic paired Excel/Word sample generation with dozens of fake metrics and edge cases.
- Learn-mode numeric extraction, normalization, matching, review output, and fail-loud coverage validation.
- Confirmed mapping and template conversion workflow.
- Deterministic run-mode report generation with source traceability.
- Later usability layers such as packaging, GUI, or business-specific polish only after the traceability loop is proven.

## Architecture Direction

- Python project with modular surfaces: `excel_profiler`, `word_profiler`, `number_normalizer`, `value_matcher`, `llm_reranker`, `mapping_reviewer`, `template_builder`, `renderer`, `validator`, and `synthetic_generator`.
- Use `openpyxl` for Excel profiling where cell address, merged cells, row/column context, and formatting clues matter.
- Use `python-docx` / later `docxtpl` for Word parsing and templating without requiring Microsoft Office.
- Normalize business formats before matching: yuan to ten-thousand yuan, counts to ten-thousand counts, decimals to percentages, commas, Chinese units, signs, and rounding tolerance.
- Prefer deterministic candidate generation first; semantic ranking is optional and must remain bounded by real candidates.
- Use `pathlib`, UTF-8 logs, and tests that cover Chinese filenames, spaces in paths, and cross-platform behavior.

## Codex Role

Codex is the planning supervisor and reviewer. It should steer scope, produce narrow Claude Code prompts, review implementation against this file and live repo evidence, and keep the project conservative.

Codex should not treat assistant summaries as evidence. For review turns, inspect the current files, diffs, generated artifacts, and verification output directly. If the user asks for a stop-gate review, lead with `ALLOW:` or `BLOCK:` and keep the finding grounded in file paths and commands.

## Verification Expectations

- Unit tests for number normalization, Excel profiling, Word profiling, candidate matching, review output, and coverage validation.
- End-to-end synthetic smoke: generate fake paired Excel/Word samples with many metrics, run learn mode, produce review artifacts, and report mapped/unresolved/low-confidence counts.
- Coverage validation must fail loudly unless incomplete coverage is explicitly allowed for an exploratory run.
- Review output must expose source sheet, cell address, row context, column context, raw value, transformed value, confidence, status, and alternatives.
- No real company files should appear in tracked files, fixtures, generated docs, logs, or screenshots.

## Scope Drift Rules

Block or defer changes that:

- make a small fixed metric list the central implementation strategy;
- hide unresolved Word numbers or report success without full coverage accounting;
- let AI invent numbers, sources, mappings, or business definitions;
- introduce real business data into the repo;
- jump to GUI, exe packaging, PDF, production rendering, or cloud integration before learn-mode traceability works;
- depend on local Excel/Word installation or OS-specific paths;
- replace focused tests with broad demos that do not assert numeric coverage.

## Open Questions

- Which Word numbers should be excluded from traceability by policy, such as dates, page numbers, IDs, version numbers, or headings?
- What confidence threshold separates confirmed, needs-review, low-confidence, and unresolved?
- What exact schema should `auto_mapping.yml` / `confirmed_mapping.yml` use?
- How should duplicate Excel values and ambiguous repeated metrics be presented for human review?
- Whether and when to add an LLM reranker, and which provider/runtime constraints apply.
