"""Promote reviewer-confirmed rows into ``confirmed_mapping.yml``.

Inputs:
  * ``auto_mapping.yml`` — the machine-readable truth ``learn`` wrote,
    one entry per visible Word number with stable ``word_id``s,
    recommended Excel source, transform metadata, and alternatives.
  * ``mapping_review.xlsx`` — the same set of rows with four
    reviewer-facing columns the human fills in by hand:
    ``Reviewer Decision``, ``Reviewer Notes``, ``Confirmed Sheet``,
    ``Confirmed Cell``. Blank is the default.

Output: ``confirmed_mapping.yml``. Three buckets:

* ``confirmed_mappings`` — HIGH/MEDIUM rows the reviewer explicitly
  marked ``confirm`` whose Excel source (either the matcher's
  recommended pick or a reviewer override that matches one of the
  matcher's alternatives) is provable. Each entry keeps the full
  traceability chain: ``word_id``, ``location``, raw token, value, unit,
  original ``status``/``confidence``, recommended/confirmed source,
  transform metadata, reviewer decision, and notes.
* ``review_required`` — every eligible row that did NOT make it into
  ``confirmed_mappings``: blank decision, reject, invalid override,
  LOW/UNRESOLVED, or any unrecognised decision string. The reviewer
  sees exactly *why* a row was held back.
* ``audit_only_excluded`` — every matcher-EXCLUDED row. These are
  date/period markers the matcher deliberately skipped; they never
  become confirmed mappings even if a reviewer types ``confirm``.

The module's safety contracts (mirrored by tests):

* **Blank is never confirmed.** A blank ``Reviewer Decision`` always
  routes to ``review_required``.
* **No silent omission.** Every ``auto_mapping.yml`` row lands in
  exactly one of the three buckets — asserted in code, not just docs.
* **No bypass of the matcher.** A reviewer override must reference a
  candidate the matcher already produced. Inventing a sheet/cell that
  the matcher never considered is invalid; it routes to
  ``review_required`` with reason ``invalid_override``.
* **EXCLUDED stays audit-only.** A reviewer who marks ``confirm`` on a
  matcher-EXCLUDED row gets that decision recorded for the audit trail
  but the row still goes to ``audit_only_excluded``.

A run with any ``review_required`` rows is considered incomplete. The
CLI returns a non-zero exit code unless ``--allow-incomplete`` is set
explicitly — that flag exists for exploratory drafts and must never be
the default in any automation pointing at real data.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl
import yaml

# Accepted reviewer decisions after lower-casing/stripping. Anything else
# is treated as ``invalid_decision`` and routed to review_required — we
# do not silently coerce typos because a typo is exactly the kind of
# thing that would silently demote a confirmed row.
DECISION_CONFIRM = "confirm"
DECISION_REJECT = "reject"
DECISION_BLANK = ""
KNOWN_DECISIONS = frozenset({DECISION_CONFIRM, DECISION_REJECT, DECISION_BLANK})

# Confidence labels the matcher emits. Only HIGH and MEDIUM can ever be
# promoted to confirmed; LOW and UNRESOLVED always need human follow-up
# beyond a checkbox tick, and EXCLUDED is policy-only.
SAFE_CONFIDENCES = frozenset({"HIGH", "MEDIUM"})

# A HIGH/MEDIUM row can only become a confirmed mapping if the template
# builder successfully placed a ``{{ word_NNNN }}`` placeholder in
# converted_template.docx. Any other ``placeholder_status`` means the
# downstream renderer has no placeholder to substitute the value into,
# so promoting the row would create a silent-omission risk: the YAML
# would claim the metric is rendered while the .docx says nothing.
RENDERABLE_PLACEHOLDER_STATUS = "applied"


@dataclass
class ReviewDecision:
    """One reviewer's decision for a single Word ID, as read from XLSX."""
    word_id: str
    decision: str        # normalized (lower-cased, stripped); "" = blank
    decision_raw: str    # exactly what the cell contained (for error messages)
    notes: str
    confirmed_sheet: str
    confirmed_cell: str


@dataclass
class ConfirmReport:
    """Structured result of one ``confirm-mapping`` run."""
    confirmed: List[Dict] = field(default_factory=list)
    review_required: List[Dict] = field(default_factory=list)
    audit_only_excluded: List[Dict] = field(default_factory=list)
    fatal_errors: List[str] = field(default_factory=list)

    @property
    def total_accounted(self) -> int:
        return (
            len(self.confirmed)
            + len(self.review_required)
            + len(self.audit_only_excluded)
        )

    @property
    def is_complete(self) -> bool:
        return not self.review_required and not self.fatal_errors


def confirm_mappings(
    auto_yaml_path: Path,
    review_xlsx_path: Path,
) -> ConfirmReport:
    """Read both inputs and classify every row into a bucket.

    Does NOT write the output YAML — that is :func:`write_confirmed_yaml`'s
    job — and does NOT decide whether to fail. The caller (CLI) reads
    ``ConfirmReport.is_complete`` and the ``--allow-incomplete`` flag.
    """
    report = ConfirmReport()
    if not auto_yaml_path.exists():
        report.fatal_errors.append(f"auto_mapping not found: {auto_yaml_path}")
        return report
    if not review_xlsx_path.exists():
        report.fatal_errors.append(f"mapping_review not found: {review_xlsx_path}")
        return report

    try:
        auto_doc = yaml.safe_load(auto_yaml_path.read_text(encoding="utf-8"))
    except Exception as exc:
        report.fatal_errors.append(f"cannot parse {auto_yaml_path.name}: {exc}")
        return report
    if not isinstance(auto_doc, dict) or not isinstance(auto_doc.get("mappings"), list):
        report.fatal_errors.append(
            f"{auto_yaml_path.name} missing top-level 'mappings' list"
        )
        return report

    try:
        decisions = _read_review_decisions(review_xlsx_path)
    except Exception as exc:
        report.fatal_errors.append(f"cannot read {review_xlsx_path.name}: {exc}")
        return report

    mappings = auto_doc["mappings"]
    yaml_ids = {m.get("word_id") for m in mappings if m.get("word_id")}
    xlsx_ids = set(decisions.keys())
    extra_in_xlsx = sorted(xlsx_ids - yaml_ids)
    if extra_in_xlsx:
        # An XLSX row whose Word ID has no YAML entry means the two
        # artifacts are out of sync — refuse to confirm anything until a
        # human runs ``validate-artifacts`` and resolves the drift,
        # otherwise we might confirm a row whose source we cannot prove.
        report.fatal_errors.append(
            f"{review_xlsx_path.name} has Word IDs absent from "
            f"{auto_yaml_path.name}: {', '.join(extra_in_xlsx)}. "
            "Run `validate-artifacts` and re-run learn if needed."
        )
        return report

    for entry in mappings:
        word_id = entry.get("word_id")
        if not isinstance(word_id, str) or not word_id:
            report.fatal_errors.append(
                f"{auto_yaml_path.name} contains a mapping with no word_id"
            )
            return report
        decision = decisions.get(
            word_id,
            # An eligible YAML row with no XLSX row is the same shape as
            # a blank decision — treated as "unreviewed" rather than
            # silently dropped, so the reviewer is forced to act.
            ReviewDecision(
                word_id=word_id, decision="", decision_raw="",
                notes="", confirmed_sheet="", confirmed_cell="",
            ),
        )
        _classify_entry(entry, decision, report)

    # No-silent-omission invariant: every YAML mapping must end up in
    # exactly one bucket. If this asserts a future refactor introduced a
    # branch that drops a row — catch it here rather than letting a
    # confirmed_mapping.yml silently omit data.
    if report.total_accounted != len(mappings):
        report.fatal_errors.append(
            f"internal accounting error: {len(mappings)} mappings but "
            f"{report.total_accounted} routed (confirmed="
            f"{len(report.confirmed)}, review_required="
            f"{len(report.review_required)}, audit_only_excluded="
            f"{len(report.audit_only_excluded)})"
        )
    return report


def write_confirmed_yaml(
    report: ConfirmReport,
    auto_yaml_path: Path,
    review_xlsx_path: Path,
    out_path: Path,
    allow_incomplete: bool,
    total_word_numbers: int,
) -> Path:
    """Serialize the report as ``confirmed_mapping.yml``.

    Always written when called — the CLI decides whether to call us in
    the first place. The YAML keeps source-artifact paths in its header
    so a reviewer can trace any confirmed mapping back to the matcher
    output it came from.
    """
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc = {
        "schema_version": 1,
        "source_artifacts": {
            "auto_mapping": str(auto_yaml_path),
            "mapping_review": str(review_xlsx_path),
        },
        "summary": {
            "total_word_numbers": total_word_numbers,
            "confirmed": len(report.confirmed),
            "review_required": len(report.review_required),
            "audit_only_excluded": len(report.audit_only_excluded),
            "allow_incomplete": bool(allow_incomplete),
            "complete": report.is_complete,
        },
        "confirmed_mappings": report.confirmed,
        "review_required": report.review_required,
        "audit_only_excluded": report.audit_only_excluded,
    }
    text = yaml.safe_dump(doc, allow_unicode=True, sort_keys=False, width=1000)
    out_path.write_text(text, encoding="utf-8")
    return out_path


def format_console_summary(
    report: ConfirmReport,
    out_path: Optional[Path],
    allow_incomplete: bool,
) -> str:
    """Human-readable summary for the CLI."""
    lines = [
        "confirm-mapping summary",
        "=======================",
        f"  confirmed          : {len(report.confirmed)}",
        f"  review_required    : {len(report.review_required)}",
        f"  audit_only_excluded: {len(report.audit_only_excluded)}",
    ]
    if out_path is not None:
        lines.append(f"  written to         : {out_path}")
    if report.review_required:
        lines.append("")
        lines.append("Review required (first 10):")
        for entry in report.review_required[:10]:
            lines.append(
                f"  - {entry['word_id']} @ {entry.get('location')!s} "
                f"[{entry.get('reason')}] decision={entry.get('reviewer_decision')!r}"
            )
        if len(report.review_required) > 10:
            lines.append(f"  … and {len(report.review_required) - 10} more")
    if report.fatal_errors:
        lines.append("")
        lines.append("FATAL:")
        for err in report.fatal_errors:
            lines.append(f"  - {err}")
    if report.review_required and not allow_incomplete:
        lines.append("")
        lines.append(
            "Run is incomplete: at least one row needs reviewer action. "
            "Fill in `Reviewer Decision` in mapping_review.xlsx and re-run, "
            "or pass --allow-incomplete for an exploratory partial output."
        )
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# XLSX → ReviewDecision map
# ---------------------------------------------------------------------------

def _read_review_decisions(path: Path) -> Dict[str, ReviewDecision]:
    """Load the reviewer-facing columns keyed by Word ID."""
    wb = openpyxl.load_workbook(str(path), data_only=True, read_only=True)
    try:
        ws = wb.active
        rows_iter = ws.iter_rows(values_only=True)
        try:
            header = list(next(rows_iter))
        except StopIteration:
            raise ValueError(f"{path.name} has no header row")
        col_index: Dict[str, int] = {}
        for i, h in enumerate(header):
            if isinstance(h, str):
                col_index[h] = i
        for required in (
            "Word ID", "Reviewer Decision", "Reviewer Notes",
            "Confirmed Sheet", "Confirmed Cell",
        ):
            if required not in col_index:
                raise ValueError(
                    f"{path.name} missing required reviewer column: {required!r}"
                )

        decisions: Dict[str, ReviewDecision] = {}
        for raw in rows_iter:
            if raw is None:
                continue
            if all(cell is None or cell == "" for cell in raw):
                continue
            word_id = _cell_str(raw, col_index["Word ID"])
            if not word_id:
                # Row without a Word ID can't be joined to YAML — skip it
                # here; the artifact validator's job is to flag it. We do
                # NOT raise, because confirm-mapping must keep running so
                # the reviewer sees every other row's classification.
                continue
            raw_decision = _cell_str(raw, col_index["Reviewer Decision"])
            decisions[word_id] = ReviewDecision(
                word_id=word_id,
                decision=raw_decision.strip().lower(),
                decision_raw=raw_decision,
                notes=_cell_str(raw, col_index["Reviewer Notes"]),
                confirmed_sheet=_cell_str(raw, col_index["Confirmed Sheet"]),
                confirmed_cell=_cell_str(raw, col_index["Confirmed Cell"]),
            )
        return decisions
    finally:
        wb.close()


def _cell_str(row: Tuple, index: int) -> str:
    if index >= len(row):
        return ""
    v = row[index]
    if v is None:
        return ""
    return str(v).strip()


# ---------------------------------------------------------------------------
# Classification
# ---------------------------------------------------------------------------

def _classify_entry(
    entry: Dict,
    decision: ReviewDecision,
    report: ConfirmReport,
) -> None:
    status = entry.get("status")

    if status == "EXCLUDED":
        report.audit_only_excluded.append(_audit_only_payload(entry, decision))
        return

    if status not in SAFE_CONFIDENCES:
        # LOW, UNRESOLVED, or anything unexpected — the matcher had no
        # safe pick, so a reviewer-side ``confirm`` is not enough on its
        # own. Surface the row so the reviewer either fixes the upstream
        # data or extends the matcher.
        reason = (
            "unresolved_no_candidate"
            if status == "UNRESOLVED"
            else "low_confidence_cannot_confirm"
            if status == "LOW"
            else f"unexpected_status:{status}"
        )
        report.review_required.append(
            _review_required_payload(entry, decision, reason)
        )
        return

    # HIGH or MEDIUM from here on.
    if decision.decision == DECISION_BLANK:
        report.review_required.append(
            _review_required_payload(entry, decision, "blank_decision")
        )
        return
    if decision.decision == DECISION_REJECT:
        report.review_required.append(
            _review_required_payload(entry, decision, "rejected")
        )
        return
    if decision.decision not in KNOWN_DECISIONS:
        report.review_required.append(
            _review_required_payload(
                entry, decision,
                f"invalid_decision:{decision.decision_raw!r}",
            )
        )
        return

    # decision == confirm.
    # Renderability guard: if the template builder could not write the
    # placeholder for this row (offset drift, raw mismatch, anything
    # other than ``applied``), the converted .docx has nothing for a
    # future renderer to substitute into. Promoting the row would let a
    # confirmed_mapping.yml claim coverage the .docx silently lacks —
    # exactly the "looks successful while omitting metrics" risk.
    placeholder_status = entry.get("placeholder_status")
    if placeholder_status != RENDERABLE_PLACEHOLDER_STATUS:
        report.review_required.append(
            _review_required_payload(
                entry, decision,
                f"non_renderable_template_skip:{placeholder_status}",
            )
        )
        return

    source, source_origin, source_error = _resolve_source(entry, decision)
    if source is None:
        report.review_required.append(
            _review_required_payload(entry, decision, source_error or "invalid_override")
        )
        return
    report.confirmed.append(
        _confirmed_payload(entry, decision, source, source_origin)
    )


def _resolve_source(
    entry: Dict, decision: ReviewDecision,
) -> Tuple[Optional[Dict], str, Optional[str]]:
    """Pick the Excel source for a confirmed row.

    Returns ``(source_dict_or_none, origin_label, error_reason)``. The
    source must already exist as a candidate the matcher produced —
    either ``recommended_source`` or one of ``alternatives``. Inventing
    a fresh ``(sheet, cell)`` pair is not allowed: doing so would bypass
    the matcher's value-agreement check and let a reviewer paste an
    arbitrary cell address into a confirmed mapping.
    """
    recommended = entry.get("recommended_source")
    alternatives = entry.get("alternatives") or []

    sheet_override = decision.confirmed_sheet
    cell_override = decision.confirmed_cell

    if not sheet_override and not cell_override:
        # No override: use the matcher's recommended pick. If the
        # matcher had no pick (defensive — HIGH/MEDIUM should always
        # have one), refuse to confirm.
        if not isinstance(recommended, dict):
            return None, "", "missing_recommended_source"
        return recommended, "recommended", None

    if bool(sheet_override) != bool(cell_override):
        # Partially-filled override: the reviewer typed only sheet OR
        # only cell. Don't guess which they meant.
        return None, "", "incomplete_override:sheet_and_cell_both_required"

    # Both override fields filled — must match recommended_source or an
    # alternative (case-sensitive cell address; case-insensitive sheet
    # name match would be friendlier but Excel sheet names are
    # case-sensitive on round-trip via openpyxl, so we match exactly).
    candidates: List[Tuple[Dict, str]] = []
    if isinstance(recommended, dict):
        candidates.append((recommended, "recommended"))
    for alt in alternatives:
        if isinstance(alt, dict):
            candidates.append((alt, "alternative"))

    for cand, origin in candidates:
        if (
            str(cand.get("sheet", "")) == sheet_override
            and str(cand.get("address", "")) == cell_override
        ):
            return cand, f"reviewer_override:{origin}", None

    return None, "", (
        f"invalid_override:{sheet_override}!{cell_override} "
        "not in recommended_source or alternatives"
    )


# ---------------------------------------------------------------------------
# Payload builders — every bucket carries the full traceability chain.
# ---------------------------------------------------------------------------

def _confirmed_payload(
    entry: Dict,
    decision: ReviewDecision,
    source: Dict,
    source_origin: str,
) -> Dict:
    return {
        "word_id": entry.get("word_id"),
        "location": entry.get("location"),
        "raw": entry.get("raw"),
        "value": entry.get("value"),
        "unit": entry.get("unit"),
        "sign": entry.get("sign"),
        "status": entry.get("status"),
        "confidence": entry.get("confidence"),
        "review_status": "confirmed",
        "reviewer_decision": decision.decision,
        "reviewer_notes": decision.notes,
        "context": entry.get("context"),
        "recommended_source": entry.get("recommended_source"),
        "confirmed_source": {
            "sheet": source.get("sheet"),
            "address": source.get("address"),
            "value": source.get("value"),
        },
        "source_origin": source_origin,
        "transform": _transform_for_source(entry, source, source_origin),
        "placeholder": entry.get("placeholder"),
        "placeholder_status": entry.get("placeholder_status"),
    }


def _transform_for_source(
    entry: Dict, source: Dict, source_origin: str,
) -> Optional[Dict]:
    """Pick the transform metadata that matches the chosen source cell.

    When the reviewer overrode to an alternative, that alternative's
    own interpretation/scores live on the source dict — auto_mapping.yml
    serialises each alternative with its own ``interpretation``,
    ``value_score``, ``context_score``, and ``overlap_tokens``. Copying
    the recommended pick's transform instead would silently apply the
    wrong unit factor at run-preview time (e.g. ``万元→base_unit`` over
    a cell the reviewer picked because it's stored ``as_written``),
    producing a confidently-wrong rendered number.

    For ``recommended`` or ``reviewer_override:recommended`` origins,
    keep the matcher's recommended transform. If an alternative is
    missing an explicit interpretation (defensive — real auto_mapping
    output always includes one), fall back to the recommended transform
    rather than emit a partial dict.
    """
    if source_origin == "reviewer_override:alternative":
        interpretation = source.get("interpretation")
        if interpretation:
            return {
                "interpretation": interpretation,
                "value_score": source.get("value_score"),
                "context_score": source.get("context_score"),
                "overlap_tokens": list(source.get("overlap_tokens") or []),
            }
    return entry.get("transform")


def _review_required_payload(
    entry: Dict, decision: ReviewDecision, reason: str,
) -> Dict:
    return {
        "word_id": entry.get("word_id"),
        "location": entry.get("location"),
        "raw": entry.get("raw"),
        "value": entry.get("value"),
        "unit": entry.get("unit"),
        "status": entry.get("status"),
        "confidence": entry.get("confidence"),
        "review_status": entry.get("review_status"),
        "reason": reason,
        "reviewer_decision": decision.decision_raw,
        "reviewer_notes": decision.notes,
        "reviewer_confirmed_sheet": decision.confirmed_sheet,
        "reviewer_confirmed_cell": decision.confirmed_cell,
        "context": entry.get("context"),
        "recommended_source": entry.get("recommended_source"),
        "alternatives": entry.get("alternatives") or [],
        "note": entry.get("note"),
    }


def _audit_only_payload(entry: Dict, decision: ReviewDecision) -> Dict:
    return {
        "word_id": entry.get("word_id"),
        "location": entry.get("location"),
        "raw": entry.get("raw"),
        "value": entry.get("value"),
        "status": entry.get("status"),
        "confidence": entry.get("confidence"),
        "review_status": "audit_only_excluded",
        "note": entry.get("note"),
        # Decision is recorded for the audit trail even though EXCLUDED
        # rows never become confirmed — surfacing a reviewer's intent
        # here lets a follow-up policy change pick the row up later.
        "reviewer_decision": decision.decision_raw,
        "reviewer_notes": decision.notes,
    }
