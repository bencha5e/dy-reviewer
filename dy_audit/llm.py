"""Revenue review delegated to an LLM.

Everything mechanical stays in Python. The revenue line items - GPR, the rent
roll it is built from, vacancy, other income, and the T12 reconciliation behind
them - move here, because that work is judgment rather than arithmetic: whether
a Pending-renewal row duplicates a unit already counted, whether pet rent
belongs inside base rent, whether a formula cell in an exported column is a
backfill or a correction. Each of those took a bespoke heuristic to catch.

The workbook is handed over as a real `.xlsx`. Claude's `document` content block
does not accept spreadsheets, so the file goes up through the Files API and is
attached as a `container_upload`, which puts it on the filesystem of the code
execution sandbox. Claude then reads it with openpyxl the same way this package
does - `data_only=False` for formulas, `data_only=True` for cached values - and
can follow a SUMIF into the rent roll instead of reasoning over a flat dump.
The workbook never enters the context window, so a 2.5 MB model costs a few
thousand tokens rather than a few hundred thousand.
"""

from __future__ import annotations

import json
import os
import re
from pathlib import Path
from typing import TYPE_CHECKING, Any

from .model import Finding, Severity, Status
from .osar import Line

if TYPE_CHECKING:  # pragma: no cover - typing only
    from .context import LoanContext


#: Default model. Opus rather than Sonnet: the failure mode that motivated this
#: rewrite is a trace that stops one hop early, which is exactly where the
#: cheaper model gives way first.
DEFAULT_MODEL = "claude-opus-5"

#: Formula tracing across tabs is the coding/agentic shape that warrants xhigh.
DEFAULT_EFFORT = "xhigh"

MAX_TOKENS = 64_000

#: The server-side code-execution loop pauses every 10 iterations. Each pause
#: costs a round trip, so cap the resumes rather than looping forever.
MAX_CONTINUATIONS = 5

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

#: The revenue rules, as prose for the model. Kept out of this file so it can be
#: edited without touching code, and cached as a stable prompt prefix.
RULES_FILENAME = "DY_Revenue_Reviewer_System_Prompt_LLM.md"


#: Check IDs the LLM owns. Findings the deterministic suite raises under these
#: IDs are discarded when the LLM review is on, so a loan is never reviewed
#: twice for the same defect under the same heading.
#:
#: CHK_PERIOD, CHK_MONTH_COUNT and CHK_MGMT_BASE are deliberately absent. They
#: touch revenue but are pure arithmetic - a date comparison, a divisor against
#: months-with-data, a formula's base reference - and Python does that more
#: cheaply and more reliably than any model.
REVENUE_CHECK_IDS = frozenset(
    {
        "CHK_GPR_RECOMPUTE",
        "CHK_RENT_STATUS",
        "CHK_VACANCY_RECOMPUTE",
        "CHK_RENT_REBUILD",
        "CHK_RENT_EDITS",
        "CHK_RR_TOTAL_TIE",
        "CHK_GPR_SUPPLEMENTARY",
        "CHK_VACANCY_SIGN",
        "CHK_VACANCY_FLOOR",
        "CHK_VACANCY_DOUBLE_COUNT",
        "CHK_OTHER_INCOME_BASIS",
        "CHK_GPR_TREND",
        "CHK_REVENUE_DOUBLE_COUNT",
        "CHK_EXCLUSIONS",
    }
)

#: Raised when the review itself could not be completed.
LLM_CHECK_ID = "CHK_REVENUE_LLM"

#: Raised when the model's reading of a clause disagrees with the regex parser's.
#: Nothing in the deterministic suite emits this, so it needs no dedupe entry.
TERM_SHEET_CHECK_ID = "CHK_TERM_SHEET_CONFLICT"

#: R-30, the tenant-status screens of Section 4a - bankruptcy, dark,
#: month-to-month. These are screens only on the loans whose own definition
#: makes them one, so no deterministic check could carry them and there is
#: nothing to dedupe against; without a heading of their own they would land in
#: the escape hatch and read as miscellany rather than as a stated exclusion.
TENANT_STATUS_CHECK_ID = "CHK_TENANT_STATUS_SCREEN"

#: The check IDs the model may file a finding under.
FINDING_CHECK_IDS = sorted(REVENUE_CHECK_IDS) + [
    TERM_SHEET_CHECK_ID,
    TENANT_STATUS_CHECK_ID,
    "CHK_REVENUE_OTHER",
]


# --------------------------------------------------------------------------
# Output schema
# --------------------------------------------------------------------------

#: The prompt's own vocabulary, not the report's. The model reasons in rules
#: (R-01..R-29) and finding types (Quantified, Memo, Latent, ...); the report is
#: organised by check ID and status. Translating in `to_findings` rather than
#: forcing the model into the report's shape keeps each side idiomatic - and the
#: fields with no home in the eight-column Findings sheet (clause, impact,
#: recommendation, the full cell list) are what the Revenue Review sheet renders.
_TERM_SHEET_SCHEMA: dict[str, Any] = {
    "type": "object",
    "description": (
        "Phase 0. Read from the agreement text, not from the parser's values. "
        "Null means the agreement states no such parameter, which is a real "
        "answer and governs which rules can run at all."
    ),
    "properties": {
        "delinquency_threshold_days": {"type": ["integer", "null"]},
        "notice_to_vacate": {
            "type": ["string", "null"],
            "description": (
                "'none' when no NTV exclusion exists, 'no time limit' when the "
                "clause has no window, otherwise the window as written."
            ),
        },
        "vacancy_floor_pct": {"type": ["number", "null"]},
        "vacancy_floor_base": {
            "type": ["string", "null"],
            "description": (
                "Which revenue the factor applies to. Note both where the "
                "agreement sets different factors per revenue limb."
            ),
        },
        "occupancy_cap_pct": {"type": ["number", "null"]},
        "new_lease_window_days": {"type": ["integer", "null"]},
        "concessions_basis": {"type": ["string", "null"]},
        "other_income_basis": {"type": ["string", "null"]},
        "tenant_status_screens": {
            "type": ["string", "null"],
            "description": (
                "Screens the definition states on tenant status - bankruptcy, "
                "dark, month-to-month, free rent, an investment-grade carve-out. "
                "Quote each, give its carve-out, and name the NOI limb it sits "
                "in. Null where the definition states none, which is the common "
                "case: a word used in passing is not a screen (Section 4a)."
            ),
        },
    },
    "required": [
        "delinquency_threshold_days",
        "notice_to_vacate",
        "vacancy_floor_pct",
        "vacancy_floor_base",
        "occupancy_cap_pct",
        "new_lease_window_days",
        "concessions_basis",
        "other_income_basis",
        "tenant_status_screens",
    ],
    "additionalProperties": False,
}

_REBUILD_SCHEMA: dict[str, Any] = {
    "type": "array",
    "description": "Phase 2, one row per in-scope revenue line rebuilt from source.",
    "items": {
        "type": "object",
        "properties": {
            "line": {"type": "string"},
            "osar_cell": {"type": ["string", "null"]},
            "reported": {"type": ["number", "null"]},
            "rebuilt": {"type": ["number", "null"]},
            "variance": {"type": ["number", "null"]},
            "flag": {"type": "string", "enum": ["PASS", "FLAG"]},
            "derivation": {
                "type": "string",
                "description": "How the rebuilt figure was built, so it can be checked.",
            },
        },
        "required": [
            "line",
            "osar_cell",
            "reported",
            "rebuilt",
            "variance",
            "flag",
            "derivation",
        ],
        "additionalProperties": False,
    },
}

_FINDING_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "id": {"type": "string", "description": "F1, F2, ... unique within this review."},
        "severity": {
            "type": "string",
            "enum": ["Blocker", "High", "Medium", "Low", "Info"],
        },
        "type": {
            "type": "string",
            "enum": [
                "Quantified",
                "Memo",
                "Unquantified",
                "Control",
                "Methodology",
                "Latent",
                "Verified",
            ],
        },
        "revenue_line": {"type": "string"},
        "rule": {"type": "string", "description": "The R-number from the catalog."},
        "check_id": {
            "type": "string",
            "enum": FINDING_CHECK_IDS,
            "description": (
                "Which heading this lands under on the reviewer's report. Pick the "
                "closest; CHK_REVENUE_OTHER is the escape hatch."
            ),
        },
        "cells": {
            "type": "array",
            "items": {"type": "string"},
            "description": (
                "Every cell the finding rests on, most important first. Prefix with "
                "a sheet name (Sheet!A1) where the cell is off the OSAR tab."
            ),
        },
        "clause": {
            "type": "string",
            "description": (
                "Verbatim words from the definition being relied on. Empty "
                "string when the finding rests on no clause."
            ),
        },
        "finding": {"type": "string", "description": "What is wrong, in plain sentences."},
        "impact_annualized": {
            "type": ["number", "null"],
            "description": "Dollar effect for a full year; null when not quantifiable.",
        },
        "evidence": {"type": "string", "description": "Empty string when none."},
        "recommendation": {"type": "string", "description": "Empty string when none."},
    },
    "required": [
        "id",
        "severity",
        "type",
        "revenue_line",
        "rule",
        "check_id",
        "cells",
        "clause",
        "finding",
        "impact_annualized",
        "evidence",
        "recommendation",
    ],
    "additionalProperties": False,
}

REVENUE_FINDINGS_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "loan": {"type": "string"},
        "test_date": {"type": "string", "description": "Empty string when not stated."},
        "osar_tab": {
            "type": "string",
            "description": "Empty string when the tab could not be identified.",
        },
        "term_sheet": _TERM_SHEET_SCHEMA,
        "rebuild": _REBUILD_SCHEMA,
        "findings": {"type": "array", "items": _FINDING_SCHEMA},
        "reviewer_note": {
            "type": "string",
            "description": (
                "Six to twelve sentences for a credit officer who will not read the "
                "JSON. This is Block B: structured output constrains the response to "
                "one object, so the prose note travels inside it."
            ),
        },
    },
    "required": [
        "loan",
        "test_date",
        "osar_tab",
        "term_sheet",
        "rebuild",
        "findings",
        "reviewer_note",
    ],
    "additionalProperties": False,
}


# --------------------------------------------------------------------------
# Prompt assembly
# --------------------------------------------------------------------------


def rules_path() -> Path:
    """Location of the revenue rules markdown, repo root by default."""
    override = os.environ.get("DY_REVENUE_RULES")
    if override:
        return Path(override)
    return Path(__file__).resolve().parent.parent / RULES_FILENAME


def load_rules() -> str:
    path = rules_path()
    if not path.is_file():
        raise FileNotFoundError(
            f"Revenue rules not found at {path}. Set DY_REVENUE_RULES to point at "
            f"{RULES_FILENAME}, or run with --no-llm to use the deterministic "
            f"revenue checks instead."
        )
    return path.read_text(encoding="utf-8")


def _params_brief(ctx: LoanContext) -> list[str]:
    """The agreement terms that govern revenue, with the text they came from.

    The quotes matter more than the values. A model told 'the vacancy floor is
    5%' will apply 5%; one shown the clause can notice that the clause says
    something the parser flattened.
    """
    params = ctx.params
    if params is None:
        return ["  (no loan agreement was parsed for this loan)"]

    wanted = [
        ("Vacancy floor", "vacancy_floor"),
        ("Investment-grade carve-out", "ig_vacancy_carveout"),
        ("Management fee %", "mgmt_fee_pct"),
        ("Management fee base, as stated", "mgmt_stated_base"),
        ("Delinquency window", "delinquency"),
        ("Other income window (months)", "other_income_months"),
        ("Concessions window (months)", "concession_months"),
        ("Known-vacate window (days)", "known_vacate_days"),
        ("Rent-step window", "rent_step_window"),
    ]
    lines: list[str] = []
    for label, attribute in wanted:
        parsed = getattr(params, attribute, None)
        if parsed is None:
            continue
        value = getattr(parsed, "value", parsed)
        if value is None:
            continue
        quote = (getattr(parsed, "quote", "") or "").strip()
        lines.append(f"  - {label}: {value}" + (f'  <- "{quote}"' if quote else ""))
    if not lines:
        lines.append("  (nothing governing revenue was parsed out of the agreement)")
    lines.append(
        "  - Management fee base actually enforced: EGI (house rule, overrides the "
        "stated base)"
    )
    return lines


def _rebuild_brief(ctx: LoanContext) -> list[str]:
    """What Python's own rent-roll rebuild found.

    Offered as a second opinion, not an answer. Python locates the tenant table
    from the export's own headers and knows nothing about why a row is there;
    where the two disagree the model should say which is right and why, not
    default to either.
    """
    rebuilds = ctx.facts.get("rebuilds") or []
    if not rebuilds:
        return ["  (Python could not rebuild a rent roll for this loan)"]

    lines: list[str] = []
    for rebuild in rebuilds:
        line = getattr(rebuild, "line", None)
        name = getattr(line, "value", None) or "GPR"
        if not getattr(rebuild, "confident", False):
            notes = "; ".join(getattr(rebuild, "notes", []) or []) or "no detail"
            lines.append(f"  - {name}: table not read confidently ({notes})")
            continue
        included = len(rebuild.included_rows())
        total = len(rebuild.rows)
        rent_column = (rebuild.columns or {}).get("rent")
        basis = "monthly x 12" if rebuild.monthly else "annual"
        lines.append(
            f"  - {name}: rebuilt {rebuild.annualized():,.2f} from column "
            f"{rent_column} ({rebuild.rent_header}, {basis}); {included} of {total} "
            f"rows included; model line reads "
            f"{(rebuild.model_gpr or 0.0):,.2f}"
        )
        for item, amount in getattr(rebuild, "reconciliation", []) or []:
            lines.append(
                f"      * {item}" + (f": {amount:+,.2f}" if amount is not None else "")
            )
    return lines


def _figures_brief(ctx: LoanContext) -> list[str]:
    """The revenue lines as the model currently reports them."""
    wb, tab = ctx.wb, ctx.tab
    wanted = [
        Line.GPR,
        Line.VACANCY,
        Line.BASE_RENT,
        Line.REIMBURSEMENT,
        Line.PERCENTAGE_RENT,
        Line.PARKING,
        Line.OTHER_INCOME,
        Line.EGI,
        Line.OCCUPANCY,
    ]
    lines: list[str] = []
    for line in wanted:
        if not tab.has(line):
            continue
        coord = tab.cell(line)
        if not coord:
            continue
        value = wb.number(tab.sheet, coord)
        formula = wb.formula(tab.sheet, coord)
        rendered = f"{value:,.2f}" if value is not None else "(no cached value)"
        entry = f"  - {line.value}: {tab.sheet}!{coord} = {rendered}"
        if formula:
            entry += f"   formula: {formula}"
        lines.append(entry)
    return lines or ["  (no revenue lines were resolved on the OSAR tab)"]


def _definitions_text(ctx: LoanContext) -> str:
    """The loan agreement's defined terms, verbatim.

    Phase 0 cannot run without this. The parsed `LoanParams` below carry the
    numbers a regex found; only the clause text carries the limbs it flattened -
    which revenue a vacancy factor attaches to, whether a notice-to-vacate
    exclusion has a window at all. These files run 4-10 KB, so sending the whole
    thing is cheaper than deciding what to leave out.
    """
    path = getattr(ctx.files, "defs_path", None)
    if not path or not Path(path).is_file():
        return "  (no loan agreement definitions file was found for this loan)"
    try:
        return Path(path).read_text(encoding="utf-8").strip()
    except OSError as exc:
        return f"  (the definitions file could not be read: {exc})"


def _meta_brief(ctx: LoanContext) -> list[str]:
    """META, per the input contract."""
    period = ctx.tab.period_end
    basis = None
    if ctx.params is not None:
        basis = getattr(getattr(ctx.params, "reserve_basis", None), "value", None)
    # Inferred, not read: the reserve is struck per unit on residential and per
    # square foot on commercial, which is the only property-type signal the
    # workbook carries in a fixed place.
    property_type = {"unit": "multifamily (inferred)", "sf": "commercial (inferred)"}.get(
        str(basis).lower(), "not determined"
    )
    # The property count sits in the overview block near the top of the tab,
    # not in the debt-yield column. Column E first, then D, matching how
    # CHK_UNIT_SF_TIE reads the same row.
    size = None
    count_row = ctx.tab.rows.get(Line.NRSF)
    if count_row is not None:
        for column in ("E", "D", ctx.tab.dy_column):
            size = ctx.wb.number(ctx.tab.sheet, f"{column}{count_row}")
            if size:
                break
    return [
        f"  - Test date: {period.strftime('%Y-%m-%d') if period else 'not read from the workbook'}",
        f"  - Property type: {property_type}",
        f"  - Net rentable units / SF: {size:,.0f}" if size else "  - Net rentable units / SF: not read",
        f"  - OSAR tab selected: {ctx.tab.sheet!r}",
    ]


def build_brief(ctx: LoanContext) -> str:
    """The per-loan half of the prompt. The rules are the stable half."""
    tabs = ctx.facts.get("source_tabs") or {}

    def tab_line(label: str, key: str) -> str:
        name = tabs.get(key)
        return f"  - {label}: {name!r}" if name else f"  - {label}: not found in this workbook"

    sections = [
        f"# Revenue review: {ctx.loan_name}",
        "",
        "The workbook is attached and available in your sandbox. Read it with "
        "openpyxl, twice: `data_only=False` for formula text and `data_only=True` "
        "for the cached values Excel last wrote. Do not recalculate it - several of "
        "these models use functions that blank out on recalculation, and the cached "
        "values are the numbers the analyst actually reported.",
        "",
        "## META",
        *_meta_brief(ctx),
        "",
        "## DEFINITIONS",
        "The defined terms from this loan's own agreement, verbatim. Build the "
        "Phase 0 term sheet from these words, not from the parsed values below. "
        "Where the agreement states no parameter for a term-sheet row, that is a "
        "real answer - return null and let the rules that depend on it stand down.",
        "",
        _definitions_text(ctx),
        "",
        "## Where things are",
        tab_line("Rent roll tab", "rent_roll"),
        tab_line("T12 / operating statement tab", "t12"),
        tab_line("AR aging tab", "ar"),
        f"  - Debt-yield column on the OSAR tab: {ctx.tab.dy_column!r}",
        "",
        "  Those tab names are what this tool resolved by following the model's own "
        "formulas. Treat them as a starting point: if the GPR line actually draws on "
        "a different sheet, or a second rent roll exists for a commercial schedule, "
        "say so and review the one the model really uses.",
        "",
        "## Revenue lines as the model reports them",
        *_figures_brief(ctx),
        "",
        "## The same agreement, as a regex parser read it",
        "  A cross-check, not an authority. Your reading of the clause wins - the "
        "parser flattens, and can miss a limb, a carve-out, or a second percentage. "
        "Where your term sheet and this disagree, that disagreement is a finding "
        "under R-27.",
        *_params_brief(ctx),
        "",
        "## Python's independent rent-roll rebuild",
        "  This is a second opinion, not the answer. It reads the tenant table from "
        "the export's own headers and has no idea why any given row is there. Where "
        "your read disagrees with it, say which is right and why.",
        *_rebuild_brief(ctx),
        "",
        "## What to return",
        "Work the phases in order against this workbook, then return one JSON object "
        "under the schema you have been given. Emit a finding for every rule you "
        "evaluated, including the ones that passed - a silent rule is "
        "indistinguishable from one that was never run. Where you could not complete "
        "a check, say so with type `Unquantified` rather than passing it.",
    ]
    return "\n".join(sections)


# --------------------------------------------------------------------------
# Response parsing
# --------------------------------------------------------------------------

_FENCE = re.compile(r"```(?:json)?\s*(.*?)```", re.S)


def _iter_text_blocks(content: Any) -> list[str]:
    out: list[str] = []
    for block in content or []:
        if getattr(block, "type", None) == "text":
            text = getattr(block, "text", "")
            if text:
                out.append(text)
    return out


def extract_payload(content: Any) -> dict:
    """Pull the findings object out of a response.

    With `output_config.format` the JSON arrives as a plain text block, but a
    code-execution turn puts server-tool blocks around it, so scan from the end
    rather than assuming position 0. The fenced-code fallback covers the case
    where the schema is refused and the model falls back to prose.
    """
    candidates = _iter_text_blocks(content)
    for text in reversed(candidates):
        stripped = text.strip()
        for attempt in (stripped, *(m.group(1).strip() for m in _FENCE.finditer(text))):
            if not attempt.startswith("{"):
                continue
            try:
                payload = json.loads(attempt)
            except json.JSONDecodeError:
                continue
            if isinstance(payload, dict) and "findings" in payload:
                return _restore_nulls(payload)
    raise ValueError(
        "The revenue review returned no parseable findings object. "
        f"Text blocks received: {len(candidates)}."
    )


#: The API caps a schema at 16 union-typed (nullable) parameters. The term
#: sheet and rebuild keep theirs - null is a semantically loaded answer there,
#: and numbers have no other way to say "none" - so these string fields carry
#: "" on the wire instead, translated back to None here so nothing downstream
#: sees a different payload than before.
_EMPTY_MEANS_NULL = ("test_date", "osar_tab")
_EMPTY_MEANS_NULL_FINDING = ("clause", "evidence", "recommendation")


def _restore_nulls(payload: dict) -> dict:
    for key in _EMPTY_MEANS_NULL:
        if payload.get(key) == "":
            payload[key] = None
    for item in payload.get("findings") or []:
        if isinstance(item, dict):
            for key in _EMPTY_MEANS_NULL_FINDING:
                if item.get(key) == "":
                    item[key] = None
    return payload


def _coerce(value: Any, enum: Any, default: Any) -> Any:
    try:
        return enum(str(value).strip().upper())
    except (ValueError, AttributeError):
        return default


#: The prompt classifies a finding by what kind of thing it is; the report asks
#: what the check concluded. `Memo` is the one that needs care: a permitted
#: inclusion the reviewer should understand as a sensitivity is not a defect, so
#: it passes - but at INFO, which the severity totals count regardless of status,
#: so it stays visible rather than reading as nothing to see.
_TYPE_TO_STATUS = {
    "verified": Status.PASS,
    "memo": Status.PASS,
    "unquantified": Status.UNVERIFIABLE,
    "control": Status.FLAG,
    "methodology": Status.FLAG,
    "latent": Status.FLAG,
    "quantified": Status.FLAG,
}


def _split_cells(cells: Any) -> tuple[str | None, str | None, list[str]]:
    """First cell becomes the Sheet/Cell columns; the rest go to evidence.

    The Findings sheet has one Cell column and it is A1-only, so a model that
    writes "Rent Roll (MF)!Q9" has to be taken apart rather than passed through.
    """
    if isinstance(cells, str):
        cells = [cells]
    entries = [c.strip() for c in (cells or []) if isinstance(c, str) and c.strip()]
    if not entries:
        return None, None, []
    primary, rest = entries[0], entries[1:]
    if "!" in primary:
        sheet, cell = primary.rsplit("!", 1)
        return (sheet.strip("'") or None), (cell or None), rest
    return None, primary, rest


def _money(value: Any) -> str | None:
    try:
        amount = float(value)
    except (TypeError, ValueError):
        return None
    return None if abs(amount) < 0.005 else f"{amount:,.2f}"


def to_findings(payload: dict) -> list[Finding]:
    """Translate the review's vocabulary into the report's.

    The model reports rules and finding types; the report renders check IDs and
    statuses. Everything the eight-column Findings sheet has no room for -
    clause, recommendation, the full cell list, the rule number - is folded into
    the message and evidence here, and rendered in full on the Revenue Review
    sheet from the raw payload.
    """
    findings: list[Finding] = []
    for item in payload.get("findings") or []:
        if not isinstance(item, dict):
            continue
        # `finding` is this schema's message field; `message` is the older one.
        message = (item.get("finding") or item.get("message") or "").strip()
        if not message:
            continue

        rule = (item.get("rule") or "").strip()
        if rule:
            message = f"[{rule}] {message}"
        impact = _money(item.get("impact_annualized"))
        if impact:
            message = f"{message} Impact: ${impact} annualised."

        sheet, cell, extra_cells = _split_cells(item.get("cells"))
        if sheet is None:
            sheet = item.get("sheet") or None
        if cell is None:
            raw = item.get("cell")
            cell = raw.rsplit("!", 1)[-1] if isinstance(raw, str) and "!" in raw else raw

        evidence_parts = [
            (item.get("evidence") or "").strip(),
            f"Clause: {item['clause'].strip()}" if (item.get("clause") or "").strip() else "",
            f"Also: {', '.join(extra_cells)}" if extra_cells else "",
            (
                f"Recommend: {item['recommendation'].strip()}"
                if (item.get("recommendation") or "").strip()
                else ""
            ),
        ]
        evidence = " | ".join(part for part in evidence_parts if part) or None

        kind = str(item.get("type") or "").strip().lower()
        status = _TYPE_TO_STATUS.get(kind)
        if status is None:
            # An older payload carrying `status` directly, or a type this schema
            # does not know. Neither may become a silent PASS.
            status = _coerce(item.get("status"), Status, Status.MANUAL_REVIEW)
        severity = _coerce(item.get("severity"), Severity, Severity.HIGH)
        if kind == "memo":
            severity = Severity.INFO

        findings.append(
            Finding(
                (item.get("check_id") or "CHK_REVENUE_OTHER").strip(),
                severity,
                status,
                message,
                sheet=sheet or None,
                cell=cell or None,
                evidence=evidence,
                on_dy_path=item.get("on_dy_path", True),
            )
        )
    return findings


def unverifiable(reason: str) -> Finding:
    """The finding raised when the review could not run.

    HIGH and UNVERIFIABLE rather than a silent skip: the house rule is that a
    check the tool could not complete must never read as a clean pass, and
    UNVERIFIABLE has its own fill on the Findings sheet.
    """
    return Finding(
        LLM_CHECK_ID,
        Severity.HIGH,
        Status.UNVERIFIABLE,
        "The revenue line items were not reviewed: the model review did not "
        f"complete ({reason}). Every revenue check on this loan is outstanding - "
        "review GPR, the rent roll behind it, vacancy, and other income by hand.",
        evidence=reason,
        on_dy_path=True,
    )


# --------------------------------------------------------------------------
# Anthropic
# --------------------------------------------------------------------------


def _anthropic_client():
    try:
        import anthropic
    except ImportError as exc:  # pragma: no cover - depends on install state
        raise RuntimeError(
            "The anthropic package is not installed. `pip install -r "
            "requirements.txt`, or run with --no-llm."
        ) from exc
    if not (os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_AUTH_TOKEN")):
        raise RuntimeError(
            "ANTHROPIC_API_KEY is not set. Export it, or run with --no-llm to use "
            "the deterministic revenue checks."
        )
    return anthropic.Anthropic()


def _call_anthropic(
    xlsx_path: Path,
    rules: str,
    brief: str,
    *,
    model: str,
    effort: str,
    client: Any = None,
) -> tuple[dict, dict]:
    """Run the review. Returns (payload, usage)."""
    client = client or _anthropic_client()

    with xlsx_path.open("rb") as handle:
        uploaded = client.beta.files.upload(
            file=(xlsx_path.name, handle, XLSX_MIME),
            betas=["files-api-2025-04-14"],
        )

    try:
        messages: list[dict] = [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": brief},
                    {"type": "container_upload", "file_id": uploaded.id},
                ],
            }
        ]
        request = {
            "model": model,
            "max_tokens": MAX_TOKENS,
            # The rules are identical for every loan in a run, so loans 2..n read
            # this prefix from cache instead of paying for it again.
            "system": [
                {"type": "text", "text": rules, "cache_control": {"type": "ephemeral"}}
            ],
            "tools": [{"type": "code_execution_20260521", "name": "code_execution"}],
            "output_config": {
                "effort": effort,
                "format": {"type": "json_schema", "schema": REVENUE_FINDINGS_SCHEMA},
            },
            # Code execution is GA; referencing an uploaded file still needs the
            # Files API beta on the message call as well as on the upload.
            "extra_headers": {"anthropic-beta": "files-api-2025-04-14"},
        }

        usage_total = {"input_tokens": 0, "output_tokens": 0, "cache_read_input_tokens": 0}
        content: Any = []
        for _ in range(MAX_CONTINUATIONS + 1):
            with client.messages.stream(messages=messages, **request) as stream:
                response = stream.get_final_message()

            usage = getattr(response, "usage", None)
            for key in usage_total:
                usage_total[key] += getattr(usage, key, 0) or 0

            content = response.content
            if getattr(response, "stop_reason", None) != "pause_turn":
                break
            # The server-side tool loop hit its iteration cap. Hand the paused
            # turn back and it resumes where it stopped; without this the review
            # is silently truncated on the larger workbooks.
            messages = messages[:1] + [{"role": "assistant", "content": content}]
        else:
            raise RuntimeError(
                f"the code-execution loop still had work after "
                f"{MAX_CONTINUATIONS} resumes"
            )

        return extract_payload(content), usage_total
    finally:
        try:
            client.beta.files.delete(uploaded.id)
        except Exception:  # noqa: BLE001 - cleanup must not mask a real failure
            pass


# --------------------------------------------------------------------------
# OpenAI
# --------------------------------------------------------------------------


def _call_openai(
    xlsx_path: Path,
    rules: str,
    brief: str,
    *,
    model: str,
    effort: str,
    client: Any = None,
) -> tuple[dict, dict]:
    """Structural fallback. Same contract, same Findings out.

    OpenAI's equivalent of container_upload is a code-interpreter container, so
    the file is uploaded with `purpose="assistants"` and attached to a tool
    resource rather than to the message.
    """
    if client is None:
        try:
            from openai import OpenAI
        except ImportError as exc:  # pragma: no cover - optional dependency
            raise RuntimeError(
                "The openai package is not installed; install it or use "
                "--llm-provider anthropic."
            ) from exc
        client = OpenAI()

    with xlsx_path.open("rb") as handle:
        uploaded = client.files.create(file=(xlsx_path.name, handle), purpose="assistants")

    try:
        response = client.responses.create(
            model=model,
            reasoning={"effort": "high" if effort in ("xhigh", "max") else effort},
            instructions=rules,
            input=[{"role": "user", "content": [{"type": "input_text", "text": brief}]}],
            tools=[
                {
                    "type": "code_interpreter",
                    "container": {"type": "auto", "file_ids": [uploaded.id]},
                }
            ],
            text={
                "format": {
                    "type": "json_schema",
                    "name": "revenue_findings",
                    "schema": REVENUE_FINDINGS_SCHEMA,
                    "strict": True,
                }
            },
        )
        payload = json.loads(response.output_text)
        usage = getattr(response, "usage", None)
        return payload, {
            "input_tokens": getattr(usage, "input_tokens", 0) or 0,
            "output_tokens": getattr(usage, "output_tokens", 0) or 0,
            "cache_read_input_tokens": 0,
        }
    finally:
        try:
            client.files.delete(uploaded.id)
        except Exception:  # noqa: BLE001
            pass


PROVIDERS = ("anthropic", "openai")


def _provider(name: str):
    """Resolve the provider at call time.

    Looked up on each call rather than bound into a table at import, so a test
    can substitute the transport without reaching inside this module's state.
    """
    if name == "anthropic":
        return _call_anthropic
    if name == "openai":
        return _call_openai
    return None


# --------------------------------------------------------------------------
# Entry point
# --------------------------------------------------------------------------


def review_revenue_with_llm(
    ctx: LoanContext,
    *,
    provider: str = "anthropic",
    model: str | None = None,
    effort: str = DEFAULT_EFFORT,
    client: Any = None,
) -> list[Finding]:
    """Review this loan's revenue lines and return findings.

    Never raises. A review that cannot run produces one UNVERIFIABLE finding, so
    the deterministic findings still ship and the gap is visible on the report
    rather than reading as a clean revenue section.
    """
    call = _provider(provider)
    if call is None:
        return [unverifiable(f"unknown provider {provider!r}")]

    try:
        rules = load_rules()
        brief = build_brief(ctx)
        payload, usage = call(
            ctx.files.xlsx_path,
            rules,
            brief,
            model=model or DEFAULT_MODEL,
            effort=effort,
            client=client,
        )
    except Exception as exc:  # noqa: BLE001 - a failed review must not fail the loan
        return [unverifiable(f"{type(exc).__name__}: {exc}")]

    ctx.facts["llm_usage"] = usage
    # The whole payload, not just the findings: the term sheet, the rebuild
    # table and the reviewer note have no home in the eight-column Findings
    # sheet and are rendered from here onto the Revenue Review sheet.
    ctx.facts["revenue_review"] = payload
    findings = to_findings(payload)
    if not findings:
        return [unverifiable("the review returned no findings")]
    return findings


__all__ = [
    "DEFAULT_EFFORT",
    "DEFAULT_MODEL",
    "FINDING_CHECK_IDS",
    "LLM_CHECK_ID",
    "REVENUE_CHECK_IDS",
    "REVENUE_FINDINGS_SCHEMA",
    "TENANT_STATUS_CHECK_ID",
    "TERM_SHEET_CHECK_ID",
    "build_brief",
    "extract_payload",
    "load_rules",
    "review_revenue_with_llm",
    "rules_path",
    "to_findings",
    "unverifiable",
]
