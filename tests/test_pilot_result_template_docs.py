"""Docs consistency check for ``docs/pilot_result_template.md``.

The pilot-result template is the redacted form a reviewer fills in
after walking the manual real-file workflow in
``docs/real_file_pilot.md``. To stay useful — and to keep its
"redacted" promise — it must:

  1. exist on disk;
  2. be linked from ``docs/real_file_pilot.md`` so the reviewer
     completing the workflow can find it without grepping the
     ``docs/`` tree;
  3. be linked from ``docs/milestone_1_readiness.md`` so a sponsor
     reading the readiness snapshot knows where the recorded
     evidence from a pilot lives;
  4. link forward to ``docs/artifact_review_guide.md`` so a reviewer
     filling in the template can re-read the forbidden-content rules
     before recording any aggregate value;
  5. carry every required aggregate section (per-stage exit codes,
     artifact presence, aggregate coverage counts, reviewer decision
     counts, accepted known limitations, reviewer attestations, and
     the final go/no-go recommendation);
  6. keep an explicit ``## What this template MUST NOT contain``
     warning that enumerates every category of forbidden content;
  7. never invite a reviewer to record raw Word snippets, raw Excel
     values, source sheet/cell content, company identifiers,
     absolute file paths, or individual ``word_id`` values outside
     the explicit warning — i.e., no Markdown table cell or
     ``<placeholder>`` slot in the body may name those categories.

This is a pure docs surface check — it imports nothing from
``src/``, never executes the CLI, never writes files, and never
touches synthetic or real data.
"""
from __future__ import annotations

import re
from pathlib import Path

import pytest


REPO_ROOT = Path(__file__).resolve().parents[1]
TEMPLATE_DOC = REPO_ROOT / "docs" / "pilot_result_template.md"
REAL_FILE_PILOT = REPO_ROOT / "docs" / "real_file_pilot.md"
READINESS_DOC = REPO_ROOT / "docs" / "milestone_1_readiness.md"


MUST_NOT_HEADING = "## What this template MUST NOT contain"


def _read_template() -> str:
    return TEMPLATE_DOC.read_text(encoding="utf-8")


def _warning_region_bounds(text: str) -> tuple[int, int] | None:
    """Return ``(start, end)`` indices of the MUST NOT contain warning.

    The warning starts at the ``## What this template MUST NOT
    contain`` heading and ends at the next horizontal-rule line
    (``\\n---``). The end index points at the start of that
    horizontal rule so the divider itself is kept *outside* the
    warning region.

    Returns ``None`` if the heading is absent — the caller decides
    how to handle the missing warning. The dedicated
    ``test_pilot_result_template_has_must_not_contain_warning`` test
    is the one that fails loudly in that case.
    """
    start = text.find(MUST_NOT_HEADING)
    if start == -1:
        return None
    body_after_heading = text[start + len(MUST_NOT_HEADING):]
    end_offset = body_after_heading.find("\n---")
    if end_offset == -1:
        return (start, len(text))
    return (start, start + len(MUST_NOT_HEADING) + end_offset)


def _warning_body(text: str) -> str:
    bounds = _warning_region_bounds(text)
    if bounds is None:
        return ""
    return text[bounds[0] : bounds[1]]


def _body_outside_warning(text: str) -> str:
    bounds = _warning_region_bounds(text)
    if bounds is None:
        return text
    return text[: bounds[0]] + text[bounds[1] :]


# ---------------------------------------------------------------------------
# Existing checks: discoverability of the template itself.
# ---------------------------------------------------------------------------

def test_pilot_result_template_doc_exists():
    """real_file_pilot.md and milestone_1_readiness.md both link to
    this file — the link targets must resolve. If this fails, both
    entry points point at a broken target and a reviewer has no
    structured form to record pilot evidence."""
    assert TEMPLATE_DOC.exists(), (
        "docs/pilot_result_template.md must exist — it is the "
        "redacted result form a reviewer fills in after walking "
        "the manual real-file workflow in docs/real_file_pilot.md."
    )


def test_real_file_pilot_doc_links_to_pilot_result_template():
    """A reviewer completing the manual workflow must be able to
    discover the result template without grepping the ``docs/``
    tree. The link must use the file's relative path so it
    resolves from inside ``docs/``."""
    text = REAL_FILE_PILOT.read_text(encoding="utf-8")
    assert "pilot_result_template.md" in text, (
        "docs/real_file_pilot.md must link to "
        "pilot_result_template.md so the reviewer completing the "
        "manual workflow can find the redacted result form."
    )


def test_milestone_readiness_doc_links_to_pilot_result_template():
    """A sponsor reading the readiness snapshot must be able to
    discover where the recorded evidence from a pilot lives,
    without context-switching back to ``real_file_pilot.md``."""
    text = READINESS_DOC.read_text(encoding="utf-8")
    assert "pilot_result_template.md" in text, (
        "docs/milestone_1_readiness.md must link to "
        "pilot_result_template.md so a sponsor reading the "
        "readiness snapshot knows where the recorded evidence "
        "from a pilot lives."
    )


# ---------------------------------------------------------------------------
# Anti-drift: the template's redacted aggregate-evidence contract.
#
# Each test below pins one slice of the contract that, if it rotted
# silently, would let a future edit either (a) drop a required
# bucket of sponsor-facing evidence or (b) introduce a fillable
# column that invites the reviewer to transcribe raw business
# content into the repo. The aim is to make either failure mode
# loud at pytest time.
# ---------------------------------------------------------------------------

def test_pilot_result_template_links_to_artifact_review_guide():
    """The artifact review guide carries the per-artifact inspection
    checklist and the forbidden-content list the template's redacted
    form depends on. A reviewer filling in the template must be able
    to pivot to that guide from the template body — if the link
    silently disappears, the reviewer loses the upstream contract
    that protects the template from data-leaking drift."""
    text = _read_template()
    assert "artifact_review_guide.md" in text, (
        "docs/pilot_result_template.md must link to "
        "artifact_review_guide.md so a reviewer filling in the "
        "redacted result form can re-read the forbidden-content "
        "rules before recording any aggregate value."
    )


REQUIRED_AGGREGATE_SECTIONS = [
    ("Per-stage exit codes", "per-stage CLI exit codes"),
    ("Artifact presence", "the artifact-presence checklist"),
    ("Aggregate coverage counts", "the aggregate coverage counts"),
    ("Reviewer decision counts", "the reviewer decision counts"),
    (
        "Known limitations accepted for this pilot",
        "the accepted known limitations",
    ),
    ("Reviewer attestations", "the reviewer attestations"),
    (
        "Final go / no-go recommendation",
        "the final go/no-go recommendation",
    ),
]


@pytest.mark.parametrize(
    "heading,description",
    REQUIRED_AGGREGATE_SECTIONS,
    ids=[h for h, _ in REQUIRED_AGGREGATE_SECTIONS],
)
def test_pilot_result_template_has_required_aggregate_section(
    heading: str, description: str,
) -> None:
    """The redacted form is structured around seven named aggregate
    sections (§2-§8 of the doc). If one disappears or is renamed, a
    reviewer can silently skip that bucket of evidence and a sponsor
    reading the form would not notice the missing piece. Each section
    must remain discoverable by its numbered ``## N. Heading``
    Markdown header so the form's structure cannot rot without this
    test firing."""
    text = _read_template()
    pattern = re.compile(
        rf"^##\s+\d+\.\s+{re.escape(heading)}\s*$", re.MULTILINE,
    )
    assert pattern.search(text), (
        f"docs/pilot_result_template.md is missing the required "
        f"section heading for {description} — expected a numbered "
        f"`## N. {heading}` heading. Restore the section so a "
        "reviewer cannot silently drop this bucket from the "
        "recorded evidence."
    )


def test_pilot_result_template_has_must_not_contain_warning() -> None:
    """The redaction contract relies on a single, explicitly named
    warning that lists the content the form is designed to suppress.
    If the heading is renamed, removed, or demoted, the surrounding
    aggregate sections lose their named protection contract and the
    "outside the warning" framing the checks below depend on is no
    longer self-documented inside the file the reviewer reads."""
    text = _read_template()
    assert MUST_NOT_HEADING in text, (
        "docs/pilot_result_template.md must keep the explicit "
        f"`{MUST_NOT_HEADING}` heading so the redaction warning is "
        "anchored in the file the reviewer reads. Without it, the "
        "aggregate sections lose their named protection contract."
    )


WARNING_FORBIDDEN_CATEGORIES = [
    ("raw Word number tokens", "raw Word snippets / tokens"),
    ("raw Excel cell values", "raw Excel cell values"),
    (
        "source sheet names or cell addresses",
        "source sheet names / cell addresses",
    ),
    ("company names", "company identifiers"),
    ("file paths beyond basenames", "file paths beyond basenames"),
    ("`word_id`", "individual word_id values"),
]


@pytest.mark.parametrize(
    "needle,description",
    WARNING_FORBIDDEN_CATEGORIES,
    ids=[d for _, d in WARNING_FORBIDDEN_CATEGORIES],
)
def test_must_not_contain_warning_enumerates_category(
    needle: str, description: str,
) -> None:
    """The warning must enumerate every category of forbidden content
    by its canonical phrase. If a category silently drops out, the
    template's redaction guarantee weakens to "we forgot to mention
    it" and a reviewer reading only the warning could miss that
    leak vector entirely."""
    body = _warning_body(_read_template())
    assert needle in body, (
        f"docs/pilot_result_template.md's `{MUST_NOT_HEADING}` "
        f"warning no longer enumerates {description} (looking for "
        f"the phrase {needle!r}). Restore the bullet so a reviewer "
        "reading the warning cannot miss this leak vector."
    )


_ALIGNMENT_CELL = re.compile(r"^\s*:?-+:?\s*$")
_LEADING_MD_MARKERS = re.compile(r"^[`*_\[\]]+")
_TRAILING_MD_MARKERS = re.compile(r"[`*_\[\]]+$")


def _is_alignment_row(line: str) -> bool:
    """Return True if ``line`` is a Markdown table separator row.

    Recognises GFM alignment rows in every common shape:
    ``| --- | --- |``, ``--- | ---`` (no outer pipes),
    ``:---|---:|:---:`` (alignment colons), and single-column
    variants like ``| --- |`` or ``---``. The row qualifies only if
    every pipe-separated cell matches ``:?-+:?`` after stripping
    whitespace — that is the property no header or data row can
    share, so it pins the table position unambiguously.
    """
    s = line.strip()
    if "|" not in s:
        return False
    inner = s
    if inner.startswith("|"):
        inner = inner[1:]
    if inner.endswith("|"):
        inner = inner[:-1]
    inner = inner.strip()
    if not inner:
        return False
    return all(_ALIGNMENT_CELL.match(cell) for cell in inner.split("|"))


def _extract_table_cells(text: str) -> list[str]:
    """Return stripped content of every Markdown table cell.

    Handles GFM-style tables with optional leading/trailing pipes.
    Tables are identified by their alignment row (the separator
    line whose cells are all ``:?-+:?``); the header above and data
    rows below are then walked outward until a blank line or a line
    without ``|`` ends the table.

    The previous detector required both leading and trailing pipes
    on every row, which silently skipped valid GFM tables like::

        Field | Value
        --- | ---
        Source Sheet | <sheet>

    The forbidden-label guard depends on this detector seeing every
    table row, so the looser shape is now in scope. Each
    pipe-separated segment is stripped of surrounding whitespace and
    returned verbatim; the per-label check applies its own
    Markdown-marker normalization before matching.
    """
    lines = text.splitlines()
    row_indices: set[int] = set()
    for i, line in enumerate(lines):
        if not _is_alignment_row(line):
            continue
        j = i - 1
        while j >= 0:
            candidate = lines[j].strip()
            if not candidate or "|" not in candidate:
                break
            row_indices.add(j)
            j -= 1
        k = i + 1
        while k < len(lines):
            candidate = lines[k].strip()
            if not candidate or "|" not in candidate:
                break
            row_indices.add(k)
            k += 1
    cells: list[str] = []
    for idx in sorted(row_indices):
        line = lines[idx].strip()
        if line.startswith("|"):
            line = line[1:]
        if line.endswith("|"):
            line = line[:-1]
        for cell in line.split("|"):
            cells.append(cell.strip())
    return cells


def _strip_md_markers(cell: str) -> str:
    """Peel off leading and trailing Markdown formatting markers.

    Repeatedly strips outermost runs of backticks, asterisks,
    underscores, and brackets — plus surrounding whitespace — until
    the result stabilises. That lets the per-label regex below
    compare against the visible label text, ignoring code-fence,
    bold/italic, link, or nested combinations such as
    ``** `Source Sheet` **``.
    """
    s = cell.strip()
    prev = ""
    while s != prev:
        prev = s
        s = _LEADING_MD_MARKERS.sub("", s)
        s = _TRAILING_MD_MARKERS.sub("", s)
        s = s.strip()
    return s


# Canonical column-header / YAML-key strings from artifacts that
# hold raw business content. Drawn from:
#   - ``src.mapping_reviewer.REVIEW_HEADERS`` (mapping_review.xlsx)
#   - ``write_run_validation`` in ``src.run_preview``
#     (run_validation.xlsx)
#   - the render_log.yml audit fields surfaced by ``render-docx``
#     (including the snake_case YAML-key forms ``word_id``,
#     ``source_sheet``, ``source_cell``, ``raw_excel_value``,
#     ``generated_value``, ``display_text``, ``raw_token``)
# Plus the company-identifier / absolute-path labels the warning
# explicitly enumerates. If any of these appears as the *label* of a
# Markdown table cell outside the MUST NOT contain warning, the
# template is asking the reviewer to record forbidden content in a
# fillable position.
FORBIDDEN_FIELD_LABELS = [
    # mapping_review.xlsx + run_validation.xlsx column headers.
    "Word ID",
    "Word Snippet",
    "Word Raw Token",
    "Word Value",
    "Word Context",
    "Raw Excel Value",
    "Generated Value",
    "Source Sheet",
    "Source Cell",
    "Top Excel Sheet",
    "Top Excel Cell",
    "Top Excel Value",
    "Top Row Context",
    "Top Column Context",
    "Reviewer Notes",
    "Confirmed Sheet",
    "Confirmed Cell",
    # render_log.yml YAML-key forms — the same content surface, in
    # lower snake_case. ``display_text`` and ``raw_token`` carry
    # rendered Word text directly.
    "word_id",
    "source_sheet",
    "source_cell",
    "raw_excel_value",
    "generated_value",
    "display_text",
    "raw_token",
    # Company-identifier labels.
    "Company Name",
    "Customer Name",
    "Product Name",
    "Brand Name",
    # Absolute-path labels.
    "Absolute Path",
    "Full Path",
]


@pytest.mark.parametrize("label", FORBIDDEN_FIELD_LABELS)
def test_template_does_not_use_forbidden_label_as_table_cell(
    label: str,
) -> None:
    """The redacted form must never carry a Markdown table cell that
    *asks* for one of the canonical column-header strings (or
    YAML-key forms) the artifacts use for raw business content.

    The check walks every Markdown table cell, peels off leading
    and trailing Markdown formatting markers (backticks, asterisks,
    underscores, brackets) so the comparison is against the visible
    label text, and matches the forbidden label as a word-bounded
    prefix. That catches every realistic shape where a cell asks for
    forbidden content:

      - plain cells: ``| Source Sheet |``
      - cells with parenthetical clarification:
        ``| Source Sheet (override) |``
      - cells with trailing modifier words:
        ``| Source Sheet column |``
      - cells wrapped in code/emphasis Markdown formatting:
        ``| `Source Sheet` |``, ``| **Source Sheet** |``,
        ``| __Source Sheet__ |``
      - cells wrapped as a Markdown link:
        ``| [Source Sheet](#anchor) |``

    Cells that merely mention the label later in a longer phrase
    (e.g. ``| Confirmation status of source sheet | ... |``) are
    not flagged because the cell is then asking for a status, not
    the source sheet itself. The check is bounded to the body
    *outside* the MUST NOT contain warning so the warning itself
    can still enumerate forbidden categories in prose without
    tripping the gate, and prose references in instructional text
    (e.g. ``Generated Value`` in the §7 attestation) are unaffected
    because they are not Markdown table cells.
    """
    outside_warning = _body_outside_warning(_read_template())
    cells = _extract_table_cells(outside_warning)
    label_re = re.compile(
        rf"^{re.escape(label)}\b", flags=re.IGNORECASE,
    )
    offenders = [
        cell for cell in cells
        if label_re.search(_strip_md_markers(cell))
    ]
    assert not offenders, (
        f"docs/pilot_result_template.md uses {label!r} as a "
        f"Markdown table cell outside the `{MUST_NOT_HEADING}` "
        f"warning (offending cells: {offenders!r}). That string is "
        "a canonical column header or YAML-key form from an "
        "artifact that holds raw business content — using it as a "
        "fillable label would invite a reviewer to transcribe "
        "forbidden content into the redacted form. Replace it with "
        "aggregate language (counts, status enums, checkboxes) or "
        "move the mention into the explicit warning region."
    )


# Substrings that, if they appear inside a fillable
# ``<placeholder>`` slot, indicate the template is asking the
# reviewer to record forbidden content. The list is intentionally
# narrow — it names only the words the artifact review guide
# forbids — so a future placeholder that asks for a count, exit
# code, date, label, or sha cannot false-trigger.
FORBIDDEN_PLACEHOLDER_SUBSTRINGS = [
    "raw",
    "snippet",
    "token",
    "sheet",
    "cell",
    "company",
    "customer",
    "product",
    "brand",
    "internal team",
    "absolute",
    "full path",
    "file path",
    "word_id",
    "word id",
]


# Match every angle-bracket placeholder slot in the template,
# whether wrapped in backticks (the template's current convention
# — ``` `<label>` ```) or bare (``<word_id>``). Either form invites
# the reviewer to substitute a value, so both must be subject to
# the forbidden-substring check. The previous pattern required
# backticks on both sides and silently skipped bare ``<...>`` slots
# — a future edit dropping the backticks would have slipped a
# forbidden placeholder past the gate. The inner ``[^>\n]+``
# keeps the match single-line so a stray ``<`` and a far-away
# ``>`` cannot accidentally pair across multiple lines.
PLACEHOLDER_PATTERN = re.compile(r"<([^>\n]+)>")


def test_template_placeholders_only_ask_for_safe_content() -> None:
    """Every fillable ``<...>`` placeholder slot in the template body
    must ask for aggregate, policy, or identifier-free content. A
    slot like ``<raw word token>`` or ``<source sheet>`` would
    invite the reviewer to paste forbidden content into the repo.

    The check matches every ``<...>`` shape in the body, with or
    without backtick wrapping — both ``` `<label>` ``` (the
    template's current convention) and a bare ``<word_id>`` (which
    a future edit could introduce without backticks, e.g. inside
    prose or inside a longer code-quoted CLI example like
    ``` `command --excel <source_sheet>` ```) are in scope, since
    either form invites the reviewer to substitute a value. The
    inner ``[^>\\n]+`` keeps each match single-line so an
    unrelated ``<`` and ``>`` on different lines cannot pair into
    a phantom placeholder.

    The check is bounded to the body *outside* the MUST NOT contain
    warning so the warning itself can still enumerate forbidden
    categories without tripping the gate. If this test fires,
    rewrite the offending placeholder to ask for aggregate/policy
    content (counts, status enums, dates, exit codes, labels
    without company identifiers) or move the mention into the
    explicit warning region.
    """
    outside_warning = _body_outside_warning(_read_template())
    offenders: list[str] = []
    for placeholder in PLACEHOLDER_PATTERN.findall(outside_warning):
        lowered = placeholder.lower()
        for forbidden in FORBIDDEN_PLACEHOLDER_SUBSTRINGS:
            if forbidden in lowered:
                offenders.append(
                    f"`<{placeholder}>` (matched forbidden substring "
                    f"{forbidden!r})"
                )
                break
    assert not offenders, (
        "docs/pilot_result_template.md has fillable `<placeholder>` "
        f"slot(s) that ask for forbidden content outside the "
        f"`{MUST_NOT_HEADING}` warning: "
        + "; ".join(offenders)
        + ". Rewrite each placeholder to ask for aggregate/policy "
        "content (counts, status enums, dates, exit codes, labels "
        "without company identifiers) or move the mention into the "
        "explicit warning region."
    )
