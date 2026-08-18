#!/usr/bin/env python3
r"""
LP DB Export -> pe-sub-jobs seed extract.

Reads two XLSX inputs from pe-sub-jobs/data/import/:
  * the LP DB Export (EXPORT_FILE) — one row per (facility account, investor)
  * AgentBankSummaryRpt.xlsx — a banded print report; the source of each facility's agent bank,
    loan amount, maturity date and agent-reported status

Writes exactly three CSVs into pe-sub-jobs/data/out/, the directory pe-sub-jobs ingests on startup:
  * lp_master.csv          one golden LP per investor_name; each field chosen independently by
                           majority vote over that investor's non-blank values across its rows
  * lp_facility_seeds.csv  one row per (facility, investor), carrying every per-LP export column
                           (facility-level AccountID/FndName/BBDate excluded)
  * facilities.csv         Agent Bank Summary rows + bank_status (Active when the account appears
                           in the export, else Inactive) + collateral_date := BBDate, plus an
                           "Unknown"-bank Inactive placeholder for every export account the report
                           does not list

Normalization against the editable lists in pe-sub-jobs/data/reference/:
  * Investor Type (investor_types.csv + investor_type_aliases.csv) and Agent LP Category
    (agent_lp_categories.csv) map to canonical values; unmatched values pass through unchanged
  * UBS classification is derived per row from that row's own attributes (agency ratings, pension
    assets, NAV, AUM, HNW/SPV flags, agent category) via the classify_ubs waterfall
  * advance rates (UBSAR and AgentAR) are slotted into discrete rate groups via rate_floor_map.csv
    (>=90 -> 90, 75-89.9 -> 75, 65-74.9 -> 65, 50-64.9 -> 50, <50 -> 0)

The input is dirty by design (name drift, 'A minus' ratings, unit mix-ups, NAV range strings), so
the parsers are tolerant and no row aborts the run. Every export account resolves to a facility,
existing or manufactured, so lp_facility_seeds.csv retains 100% of the export's records; the run
prints those counts and nothing else.

Usage (no command-line arguments): edit EXPORT_FILE below, then run from any working directory:
    python pe-sub-jobs/scripts/lp_db_extract.py
"""
from __future__ import annotations

import csv
import re
import sys
from collections import Counter, OrderedDict
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

try:
    import openpyxl
except ImportError:  # pragma: no cover
    sys.exit("openpyxl is required: pip install openpyxl")

# Paths are anchored to this file's location so the script runs from any working directory.
SCRIPT_DIR = Path(__file__).resolve().parent            # pe-sub-jobs/scripts/
JOBS_ROOT = SCRIPT_DIR.parent                           # pe-sub-jobs/
DATA_DIR = JOBS_ROOT / "data"

# ============================================================================================
#  EDIT THIS for each run — the LP DB Export to process. Absolute, or relative to pe-sub-jobs/.
# ============================================================================================
EXPORT_FILE = DATA_DIR / "import" / "LP DB Export 2026.06.25.xlsx"

AGENT_BANK_SUMMARY_FILE = DATA_DIR / "import" / "AgentBankSummaryRpt.xlsx"
OUT_DIR = DATA_DIR / "out"              # all outputs land here
REFERENCE_DIR = DATA_DIR / "reference"  # normalization lists

# --- source column order (must match the export header exactly) ----------------------------
SRC_COLS = [
    "AccountID", "FndName", "InvestorName", "Parent", "SPV", "InvestorType", "Region", "HQ",
    "InstitutionalHNW", "InvestmentGrade", "Classification", "Notes", "SP", "Moodys", "Fitch",
    "AUM", "NAV", "PensionAssets", "FundingRatio", "UBSAR", "AgentAR", "Commitments",
    "PercentOfCommitments", "Called", "Uncalled", "PercentOfUncalled", "CalledPercent",
    "AgentCL", "UBSCL", "AgentBB", "UBSBB", "BBDate",
]


# The platform's own LP Records export (pe-sub-ui/src/services/lpExportService.ts) writes these same
# 32 columns, in this order, under readable headers - so a workbook exported from the UI can be fed
# straight back in. Keep this map in step with that file: readable header -> LP DB Export column.
PLATFORM_HEADERS = {
    "Account ID": "AccountID",                     "Fund Name": "FndName",
    "Investor Name": "InvestorName",               "Parent": "Parent",
    "SPV": "SPV",                                  "Investor Type": "InvestorType",
    "Region / Location": "Region",                 "High Quality": "HQ",
    "Institutional vs HNW": "InstitutionalHNW",    "Investment Grade": "InvestmentGrade",
    "Agent LP Classification": "Classification",   "Notes": "Notes",
    "S&P": "SP",                                   "Moody's": "Moodys",
    "Fitch": "Fitch",                              "AUM": "AUM",
    "NAV": "NAV",                                  "Pension Assets": "PensionAssets",
    "Funded Ratio (%)": "FundingRatio",            "UBS Advance Rate (%)": "UBSAR",
    "Agent Advance Rate (%)": "AgentAR",           "Capital Commitments": "Commitments",
    "% of Commitments": "PercentOfCommitments",    "Called Capital": "Called",
    "Uncalled Capital": "Uncalled",                "% of Uncalled Capital": "PercentOfUncalled",
    "% of LP Called": "CalledPercent",             "Agent Concentration Limit": "AgentCL",
    "UBS Concentration Limit": "UBSCL",            "Agent Borrowing Base": "AgentBB",
    "UBS Borrowing Base": "UBSBB",                 "Collateral Date": "BBDate",
}

# Mapping the header is not enough: the platform export shapes its *values* for a spreadsheet, not
# for this feed. Percents go out as numbers under a "(%)" header (94, not 0.94) and money as display
# strings ("$428,800,000", not 428800000), so each such column is converted back on the way in -
# without this every rate and ratio would land 100x too big.
PLATFORM_PERCENT_COLS = {"FundingRatio", "UBSAR", "AgentAR",
                         "PercentOfCommitments", "PercentOfUncalled", "CalledPercent"}
PLATFORM_MONEY_COLS = {"Commitments", "Called", "Uncalled", "AgentBB", "UBSBB"}
# A concentration limit is either a percent of uncalled ("7.5%") or an absolute cap ("$25,000,000")
# in the same column - the '%' sign is what tells them apart.
PLATFORM_LIMIT_COLS = {"AgentCL", "UBSCL"}


def _platform_number(v) -> "tuple[Decimal, bool] | None":
    """Strip the display formatting off one cell -> (number, was_a_percent).
    '$428,800,000' -> (428800000, False) - '7.5%' -> (7.5, True) - 94 -> (94, False).
    None when the cell holds nothing numeric, in which case the caller keeps it verbatim."""
    s = str(v).strip().replace(",", "").replace("$", "")
    was_pct = s.endswith("%")
    if was_pct:
        s = s[:-1].strip()
    try:
        return Decimal(s), was_pct
    except InvalidOperation:
        return None


def from_platform(col: str, v):
    """Undo the platform export's display formatting for one cell of a mapped column."""
    if blank(v) or col not in (PLATFORM_PERCENT_COLS | PLATFORM_MONEY_COLS | PLATFORM_LIMIT_COLS):
        return v
    parsed = _platform_number(v)
    if parsed is None:
        return v                                   # unparseable: pass through, same as a dirty feed
    number, was_pct = parsed
    if col in PLATFORM_PERCENT_COLS:               # 94 or '94%' -> 0.94
        return _trim(number / 100)
    if col in PLATFORM_LIMIT_COLS:                 # '7.5%' -> 0.075, '$25,000,000' -> 25000000
        return _trim(number / 100 if was_pct else number)
    return _trim(number)                           # money: '$428,800,000' -> 428800000

# CSV header (column) orders required by the pe-sub-jobs FlatFileItemReaders.
MASTER_COLS = [
    "investor_name", "parent", "spv", "high_quality", "investor_type", "institutional_or_hnw",
    "region_location", "investment_grade", "sp_rating", "moodys_rating", "fitch_rating", "aum", "nav", "pension_assets",
    "funding_ratio", "ubs_lp_category", "ubs_default_advance_rate", "ubs_default_concentration_limit",
    "notes",
]
SEED_COLS = [
    "facility_name", "investor_name", "capital_commitment", "uncalled_capital",
    "agent_lp_category", "agent_advance_rate", "agent_concentration_limit",
    "parent", "spv", "high_quality", "investor_type", "institutional_or_hnw", "region_location",
    "investment_grade", "ubs_lp_category", "sp_rating", "moodys_rating", "fitch_rating", "aum", "nav", "pension_assets",
    "funding_ratio", "pct_of_fund_commitments", "called_capital", "pct_of_fund_uncalled", "pct_lp_called",
    "ubs_concentration_limit", "ubs_advance_rate", "agent_borrowing_base", "ubs_borrowing_base", "notes",
]
FACILITY_COLS = [
    "agent_bank", "name", "account_number", "loan_amount", "maturity_date", "bank_status",
    "bank_status_date", "ubs_participation", "collateral_date",
]

# Agent Bank Summary column layout (must match the report header exactly). Index 4 is an unnamed
# spacer holding the report's subtotal amounts.
ABS_COLS = [
    "Agent", "Borrower", "AccountNumber", "LoanAmount", "", "MaturityDate",
    "FacilityStatus", "FacilityStatusDate",
]
ABS_TOTAL_MARKER = "accesstotalsloanamount"  # _norm() prefix of the subtotal / grand-total rows


# --- formatting helpers --------------------------------------------------------------------
def _trim(dec: Decimal) -> str:
    """Fixed-point string with trailing zeros trimmed; never scientific notation."""
    s = format(dec, "f")
    if "." in s:
        s = s.rstrip("0").rstrip(".")
    return s or "0"


def blank(v) -> bool:
    return v is None or str(v).strip() == ""


def as_is(v) -> str:
    """Passthrough: the raw extracted value verbatim, trimmed."""
    return "" if v is None else str(v).strip()


def yn_bool(v, *, yes=("y", "yes", "true", "1")) -> str:
    return "TRUE" if str(v).strip().lower() in yes else "FALSE"


def pct(v) -> str:
    """Fraction (0.41) -> percent string ('41%'). Exact, no rounding."""
    if blank(v):
        return ""
    try:
        return _trim(Decimal(str(v)) * 100) + "%"
    except InvalidOperation:
        return ""


def dec_str(v) -> str:
    """Decimal rate/limit passthrough as a trimmed string ('0.41')."""
    if blank(v):
        return ""
    try:
        return _trim(Decimal(str(v)))
    except InvalidOperation:
        return str(v).strip()


def money_short(v) -> str:
    """Exact short-currency dollars ($484M, $314.6M, $8B); no rounding."""
    if blank(v):
        return ""
    try:
        d = Decimal(str(v))
    except InvalidOperation:
        return str(v).strip()
    neg = d < 0
    d = abs(d)
    if d >= Decimal("1e9"):
        body, suffix = d / Decimal("1e9"), "B"
    elif d >= Decimal("1e6"):
        body, suffix = d / Decimal("1e6"), "M"
    elif d >= Decimal("1e3"):
        body, suffix = d / Decimal("1e3"), "K"
    else:
        body, suffix = d, ""
    return ("-" if neg else "") + "$" + _trim(body) + suffix


def iso_date(v) -> str:
    """Parse a feed date to ISO (YYYY-MM-DD); '' if unparseable. The export carries text dates;
    the Agent Bank Summary carries real Excel dates, which openpyxl returns as datetime."""
    if blank(v):
        return ""
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    s = str(v).strip()
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%m/%d/%y"):
        try:
            return datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            continue
    return ""


# --- reference lists ------------------------------------------------------------------------
def _norm(s) -> str:
    """Case/punctuation-insensitive key: lowercase, non-alphanumerics collapsed to one space."""
    return re.sub(r"[^a-z0-9]+", " ", str(s or "").lower()).strip()


@dataclass
class Reference:
    itype_lookup: dict[str, str]            # norm(value/alias) -> canonical Investor Type
    agent_lookup: dict[str, str]            # norm(alias) -> canonical Agent LP Category
    rate_floors: list[tuple[float, float]]  # (min_rate_pct, group_pct), sorted highest-min first


def _read_reference_rows(path: Path) -> list[list[str]]:
    """Read a reference CSV, dropping blank lines and '#' comment lines."""
    if not path.is_file():
        raise SystemExit(
            f"Reference file not found: {path}\n"
            "The pe-sub-jobs/data/reference/ lists are required. Restore it (or fix REFERENCE_DIR)."
        )
    rows: list[list[str]] = []
    with path.open(newline="", encoding="utf-8") as fh:
        for cells in csv.reader(fh):
            if not cells or not cells[0].strip() or cells[0].lstrip().startswith("#"):
                continue
            rows.append([c.strip() for c in cells])
    return rows


def load_references(ref_dir: Path) -> Reference:
    itype_canonical = [r[0] for r in _read_reference_rows(ref_dir / "investor_types.csv")[1:]]
    itype_lookup = {_norm(c): c for c in itype_canonical}
    for row in _read_reference_rows(ref_dir / "investor_type_aliases.csv")[1:]:
        if len(row) >= 2 and row[1]:
            itype_lookup[_norm(row[0])] = row[1]

    agent_rows = _read_reference_rows(ref_dir / "agent_lp_categories.csv")[1:]
    agent_lookup = {_norm(r[0]): r[1] for r in agent_rows if len(r) >= 2 and r[1]}

    # Floor Map: a rate takes the group of the highest 'min' it meets or exceeds; the min=0 floor
    # makes it total.
    rate_floors: list[tuple[float, float]] = []
    for row in _read_reference_rows(ref_dir / "rate_floor_map.csv")[1:]:
        if len(row) < 2:
            continue
        try:
            rate_floors.append((float(row[0]), float(row[1])))
        except ValueError:
            continue
    rate_floors.sort(key=lambda t: t[0], reverse=True)

    return Reference(itype_lookup, agent_lookup, rate_floors)


def map_investor_type(raw, ref: Reference) -> tuple[str, bool]:
    """(canonical, True) when the value maps to a supported Investor Type; otherwise the ORIGINAL
    value with (value, False). The record is always kept."""
    s = as_is(raw)
    if not s:
        return "", True
    canon = ref.itype_lookup.get(_norm(s))
    return (canon, True) if canon else (s, False)


def map_agent_cls(raw, ref: Reference) -> tuple[str, bool]:
    """(canonical Agent LP Category, True) when mapped; else the original value with False."""
    s = as_is(raw)
    if not s:
        return "", True
    canon = ref.agent_lookup.get(_norm(s))
    return (canon, True) if canon else (s, False)


# --- UBS classification ---------------------------------------------------------------------
# Canonical class labels (BB Criteria Matrix taxonomy).
CLS_RATED = "Rated Investor"
CLS_CP5 = "Corp Pension > $5Bn Assets"
CLS_CP1 = "Corp Pension > $1Bn Assets"
CLS_NAV1 = "Unrated NAV > $1Bn"
CLS_FOF = "FoF & Other > $10Bn AUM"
CLS_OTHER = "Other Institutional"
CLS_HNW_FEEDER = "HNW Feeder (acceptable)"
CLS_HNW = "HNW (acceptable)"
CLS_EXCLUDED = "Excluded"

# Unified ordinal notch scale (1 = strongest) accepting S&P/Fitch and Moody's notation.
_SP_SCALE = ["AAA", "AA+", "AA", "AA-", "A+", "A", "A-", "BBB+", "BBB", "BBB-",
             "BB+", "BB", "BB-", "B+", "B", "B-", "CCC+", "CCC", "CCC-", "CC", "C", "D"]
_MDY_SCALE = ["Aaa", "Aa1", "Aa2", "Aa3", "A1", "A2", "A3", "Baa1", "Baa2", "Baa3",
              "Ba1", "Ba2", "Ba3", "B1", "B2", "B3", "Caa1", "Caa2", "Caa3", "Ca", "C"]
_RATING_NOTCH = {r.lower(): i + 1 for i, r in enumerate(_SP_SCALE)}
for _i, _r in enumerate(_MDY_SCALE):
    _RATING_NOTCH.setdefault(_r.lower(), _i + 1)
_NOT_RATED = {"nr", "n/r", "notrated", "n/a", "na", "wr", "unrated", "none"}


def rating_notch(v) -> int | None:
    """Normalize one agency rating to its ordinal notch, tolerating manual-entry drift
    ('A minus' -> 'A-', case noise, 'Not Rated'/'N/R'/'' -> None)."""
    s = str(v or "").strip().lower().replace("minus", "-").replace("plus", "+").replace(" ", "")
    if not s or s in _NOT_RATED:
        return None
    return _RATING_NOTCH.get(s)


def eligible_band(sp, mdy, fitch) -> str:
    """Eligible-rating waterfall -> rating band: three ratings -> median, two -> the lower,
    one -> as-is; sub-BBB- clamps to the BBB band. '' when no agency rating is usable."""
    notches = sorted(n for n in (rating_notch(v) for v in (sp, mdy, fitch)) if n is not None)
    if not notches:
        return ""
    n = notches[1] if len(notches) >= 2 else notches[0]
    if n <= 1:
        return "AAA"
    if n <= 4:
        return "AA"
    if n <= 7:
        return "A"
    return "BBB"


_UNIT_MULT = {"": 1.0, "k": 1e3, "m": 1e6, "mn": 1e6, "mm": 1e6,
              "b": 1e9, "bn": 1e9, "t": 1e12, "tn": 1e12, "trn": 1e12}


def parse_money_low(v) -> float | None:
    """Tolerant size parser -> absolute dollars for the export's free-text AUM/NAV/PensionAssets
    ('$14.9B', '240B', '1.33. tn', '$21 bn+', '394.6667', '1-5M', '500M - 2Bn', '>5B', '<100M').
    A range takes the LOW end; a bare number is absolute dollars. None when nothing numeric reads."""
    if blank(v):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().lower().replace(",", "").replace("$", "")
    s = s.lstrip("<>~ ").rstrip("+ ")
    s = re.sub(r"\.(?!\d)", " ", s)          # stray dots ('1.33. tn') — keep decimal points only
    parts = re.findall(r"(\d+(?:\.\d+)?)\s*([a-z]*)", s)
    if not parts:
        return None
    units = [u for _, u in parts if u in _UNIT_MULT and u]
    default_unit = units[-1] if units else ""  # '1-5M': the shared unit applies to both ends
    vals = []
    for num, unit in parts:
        mult = _UNIT_MULT.get(unit) if unit in _UNIT_MULT else _UNIT_MULT.get(default_unit, 1.0)
        vals.append(float(num) * mult)
    return min(vals) if vals else None


def _is_yes(v) -> bool:
    return str(v or "").strip().lower() in ("y", "yes", "true", "1")


def classify_ubs(row: dict, ref: Reference) -> str:
    """Derive the UBS LP Classification from the LP's own attributes. Waterfall:
      1. agent category 'Ineligible Investor' -> Excluded;
      2. any usable agency rating -> Rated Investor;
      3. HNW (InstitutionalHNW flag, or agent 'Designated PWM') -> HNW Feeder (acceptable) when
         the vehicle is an SPV, else HNW (acceptable);
      4. pension assets > $5Bn / > $1Bn -> the two Corp Pension classes;
      5. NAV > $1Bn -> Unrated NAV > $1Bn;
      6. FoF/hedge fund investor type with AUM > $10Bn -> FoF & Other > $10Bn AUM;
      7. catch-all -> Other Institutional, so the mapping always resolves."""
    agent_canon, matched = map_agent_cls(row["Classification"], ref)
    agent = agent_canon if matched else ""
    if agent == "Ineligible Investor":
        return CLS_EXCLUDED
    if eligible_band(row["SP"], row["Moodys"], row["Fitch"]):
        return CLS_RATED
    if _norm(row["InstitutionalHNW"]) == "hnw" or agent == "Designated PWM":
        return CLS_HNW_FEEDER if _is_yes(row["SPV"]) else CLS_HNW
    pension = parse_money_low(row["PensionAssets"])
    if pension is not None and pension > 5e9:
        return CLS_CP5
    if pension is not None and pension > 1e9:
        return CLS_CP1
    nav = parse_money_low(row["NAV"])
    if nav is not None and nav > 1e9:
        return CLS_NAV1
    itype = map_investor_type(row["InvestorType"], ref)[0]
    aum = parse_money_low(row["AUM"])
    if itype in ("Fund of Funds", "Hedge Fund") and aum is not None and aum > 10e9:
        return CLS_FOF
    return CLS_OTHER


def rate_group_pct(v, ref: Reference) -> float | None:
    """Floor Map: slot a raw advance-rate fraction (0.93) into its rate group in percent (90.0) —
    the group of the highest 'min' the rate meets or exceeds. None when missing/unparseable."""
    if blank(v):
        return None
    try:
        rate_pct = float(v) * 100
    except (TypeError, ValueError):
        return None
    for min_pct, group_pct in ref.rate_floors:   # sorted highest-min first
        if rate_pct >= min_pct:
            return group_pct
    return None


def floor_rate_pct(v, ref: Reference) -> str:
    """Floor-mapped rate as a percent string ('90%'); '' when unmapped."""
    group = rate_group_pct(v, ref)
    return "" if group is None else _trim(Decimal(str(group))) + "%"


def floor_rate_frac(v, ref: Reference) -> str:
    """Floor-mapped rate as a fraction string ('0.9'); '' when unmapped."""
    group = rate_group_pct(v, ref)
    return "" if group is None else _trim(Decimal(str(group)) / 100)


# --- extract -------------------------------------------------------------------------------
def read_export(path: Path, sheet: str | None = None) -> list[dict]:
    if not path.is_file():
        raise SystemExit(
            f"Export file not found: {path}\n"
            "Edit the EXPORT_FILE variable near the top of pe-sub-jobs/scripts/lp_db_extract.py "
            "to point at the LP DB Export .xlsx, then re-run."
        )
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if sheet is None:
        ws = wb[wb.sheetnames[0]]
    elif sheet in wb.sheetnames:
        ws = wb[sheet]
    else:
        raise SystemExit(
            f"Sheet '{sheet}' not found in {path.name}. Available: {wb.sheetnames}"
        )
    rows_iter = ws.iter_rows(values_only=True)
    header = ["" if h is None else str(h).strip() for h in next(rows_iter)]
    # Columns are addressed by name, not position: either the LP DB Export's own header or the
    # readable one the platform's LP Records export writes (PLATFORM_HEADERS). Anything else is
    # ignored, so an extra trailing column never breaks the read.
    column_at: dict[str, int] = {}
    platform_cols: set[str] = set()
    for i, name in enumerate(header):
        if name in SRC_COLS:
            column_at.setdefault(name, i)
        elif name in PLATFORM_HEADERS:
            src = PLATFORM_HEADERS[name]
            if src not in column_at:
                column_at[src] = i
                platform_cols.add(src)
    missing = [c for c in SRC_COLS if c not in column_at]
    if missing:
        raise SystemExit(
            f"""Export header in '{ws.title}' is missing {len(missing)} of the {len(SRC_COLS)} expected columns: {missing}
  LP DB Export headers: {SRC_COLS}
  or the platform LP Records export's readable headers: {list(PLATFORM_HEADERS)}
  found: {header}"""
        )
    if platform_cols:
        print(f"  header: platform LP Records export detected - {len(platform_cols)} readable "
              f"column(s) mapped back to the LP DB Export schema, values de-formatted.")
    rows = []
    for r in rows_iter:
        row = {c: (r[column_at[c]] if column_at[c] < len(r) else None) for c in SRC_COLS}
        for c in platform_cols:
            row[c] = from_platform(c, row[c])
        rows.append(row)
    return rows


def read_agent_bank_summary(path: Path) -> tuple[list[list[str]], dict[str, int]]:
    """Read the Agent Bank Summary report into FACILITY_COLS-shaped rows.
    Returns (rows, account_number -> row index).

    The report is a banded print layout, not a flat table:
      * the agent bank appears once on its own group-header row (Agent set, Borrower blank) and is
        carried down onto the facility rows beneath it, which leave the Agent cell blank;
      * each group ends with an 'AccessTotalsLoanAmount:' subtotal row carrying no facility.

    Dirty rows are absorbed rather than fatal: a repeated (AccountNumber, Borrower) pair is a
    reprint and is dropped; a Borrower name already taken by a different account is suffixed with
    its AccountID; one AccountNumber listed against two borrowers keeps the FIRST row as the owner
    of the LP join.

    ubs_participation and collateral_date are not in the report — collateral_date is filled from
    the export's BBDate by upsert_facilities. bank_status comes from FacilityStatus here and is
    then overridden by upsert_facilities with the export-match result."""
    if not path.is_file():
        raise SystemExit(
            f"Agent Bank Summary report not found: {path}\n"
            "It is the source of every facility's agent bank, loan amount, maturity date and "
            "agent-reported status. Drop it in pe-sub-jobs/data/import/, or edit the "
            "AGENT_BANK_SUMMARY_FILE variable near the top of this script, then re-run."
        )
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows_iter = ws.iter_rows(values_only=True)
    header = [as_is(c) for c in next(rows_iter)]
    if header[: len(ABS_COLS)] != ABS_COLS:
        raise SystemExit(
            f"Agent Bank Summary header in '{ws.title}' does not match the expected schema.\n"
            f"  expected: {ABS_COLS}\n  found:    {header}"
        )

    data: list[list[str]] = []
    by_acct: dict[str, int] = {}
    seen_pair: set[tuple[str, str]] = set()
    used_norm: set[str] = set()
    agent = ""

    for raw in rows_iter:
        cells = (list(raw) + [None] * len(ABS_COLS))[: len(ABS_COLS)]
        text = [as_is(c) for c in cells]
        if not any(text):
            continue
        if _norm(text[3]).startswith(ABS_TOTAL_MARKER):   # subtotal / grand-total band
            continue
        if text[0] and not text[1]:                       # agent group header -> carry down
            agent = text[0]
            continue
        name, acct = text[1], text[2]
        if not name or not acct:
            continue
        if (acct, _norm(name)) in seen_pair:              # reprint of a row already taken
            continue
        seen_pair.add((acct, _norm(name)))
        if _norm(name) in used_norm:
            name = f"{name} ({acct})"
        used_norm.add(_norm(name))
        # The row's own Agent cell wins if the report fills it; otherwise the carried-down header.
        # "Unknown" satisfies FacilityRowProcessor's non-blank agent_bank rule.
        data.append([text[0] or agent or "Unknown", name, acct, text[3], iso_date(cells[5]),
                     text[6], iso_date(cells[7]), "", ""])
        by_acct.setdefault(acct, len(data) - 1)

    return data, by_acct


def _best(rows: list[dict], field: str) -> str:
    """Consolidate one field across an investor's rows: the most frequently reported non-blank
    value; ties fall back to first occurrence (Counter preserves insertion order). Blanks do not
    vote. '' only when every row is blank."""
    tally: Counter = Counter()
    for r in rows:
        v = as_is(r[field])
        if v:
            tally[v] += 1
    return tally.most_common(1)[0][0] if tally else ""


def build_master(export: list[dict], ref: Reference) -> list[dict]:
    """One golden LP per investor_name, each attribute chosen independently by majority vote over
    that investor's rows. Investor Type is voted on its canonical mapping; UBS classification is
    derived from the consolidated best attributes."""
    groups: "OrderedDict[str, list[dict]]" = OrderedDict()
    for row in export:
        name = (row["InvestorName"] or "").strip()
        if not name:
            continue
        groups.setdefault(name, []).append(row)

    master_rows: list[dict] = []
    for name, rows in groups.items():
        itype_votes: Counter = Counter()
        for r in rows:
            v = as_is(r["InvestorType"])
            if v:
                itype_votes[map_investor_type(v, ref)[0]] += 1
        investor_type = itype_votes.most_common(1)[0][0] if itype_votes else ""

        best_attrs = {c: _best(rows, c) for c in (
            "Classification", "SP", "Moodys", "Fitch", "InstitutionalHNW", "SPV",
            "PensionAssets", "NAV", "AUM", "InvestorType")}

        master_rows.append({
            "investor_name": name,
            "parent": _best(rows, "Parent"),
            "spv": yn_bool(_best(rows, "SPV")),
            "high_quality": yn_bool(_best(rows, "HQ")),
            "investor_type": investor_type,
            "institutional_or_hnw": _best(rows, "InstitutionalHNW"),
            "region_location": _best(rows, "Region"),
            "investment_grade": yn_bool(_best(rows, "InvestmentGrade")),
            "sp_rating": _best(rows, "SP"),
            "moodys_rating": _best(rows, "Moodys"),
            "fitch_rating": _best(rows, "Fitch"),
            "aum": _best(rows, "AUM"),
            "nav": _best(rows, "NAV"),
            "pension_assets": _best(rows, "PensionAssets"),
            "funding_ratio": pct(_best(rows, "FundingRatio")),
            "ubs_lp_category": classify_ubs(best_attrs, ref),
            "ubs_default_advance_rate": floor_rate_frac(_best(rows, "UBSAR"), ref),
            "ubs_default_concentration_limit": dec_str(_best(rows, "UBSCL")),
            "notes": _best(rows, "Notes"),
        })

    return master_rows


@dataclass
class SeedResult:
    rows: list[dict]
    counts: Counter


def build_seed(export: list[dict], name_by_acct: dict[str, str], ref: Reference) -> SeedResult:
    """One seed row per (facility, investor) carrying every per-LP export column. Every export
    account resolves to a facility, so nothing is rejected for a missing facility. Classification
    and Investor Type are normalised (unmatched passed through); ubs_lp_category is derived per row
    from that row's own attributes."""
    seen: set[tuple[str, str]] = set()
    seed_rows: list[dict] = []
    counts = Counter()

    for row in export:
        acct = (row["AccountID"] or "").strip()
        investor = (row["InvestorName"] or "").strip()
        fac_name = name_by_acct.get(acct)
        if fac_name is None:
            counts["skipped_no_facility"] += 1
            continue
        if not investor:
            counts["skipped_blank_investor"] += 1
            continue
        key = (fac_name, investor)
        if key in seen:
            counts["skipped_duplicate_pair"] += 1
            continue
        seen.add(key)

        agent_cls = map_agent_cls(row["Classification"], ref)[0]
        ubs_cls = classify_ubs(row, ref)

        seed_rows.append({
            "facility_name": fac_name,
            "investor_name": investor,
            "capital_commitment": money_short(row["Commitments"]),
            "uncalled_capital": money_short(row["Uncalled"]),
            "agent_lp_category": agent_cls,
            "agent_advance_rate": floor_rate_pct(row["AgentAR"], ref),
            "agent_concentration_limit": pct(row["AgentCL"]),
            "parent": as_is(row["Parent"]),
            "spv": yn_bool(row["SPV"]),
            "high_quality": yn_bool(row["HQ"]),
            "investor_type": map_investor_type(row["InvestorType"], ref)[0],
            "institutional_or_hnw": as_is(row["InstitutionalHNW"]),
            "region_location": as_is(row["Region"]),
            "investment_grade": yn_bool(row["InvestmentGrade"]),
            "ubs_lp_category": ubs_cls,
            "sp_rating": as_is(row["SP"]),
            "moodys_rating": as_is(row["Moodys"]),
            "fitch_rating": as_is(row["Fitch"]),
            "aum": as_is(row["AUM"]),
            "nav": as_is(row["NAV"]),
            "pension_assets": as_is(row["PensionAssets"]),
            "funding_ratio": pct(row["FundingRatio"]),
            "pct_of_fund_commitments": pct(row["PercentOfCommitments"]),
            "called_capital": money_short(row["Called"]),
            "pct_of_fund_uncalled": pct(row["PercentOfUncalled"]),
            "pct_lp_called": pct(row["CalledPercent"]),
            "ubs_concentration_limit": pct(row["UBSCL"]),
            "ubs_advance_rate": floor_rate_pct(row["UBSAR"], ref),
            "agent_borrowing_base": money_short(row["AgentBB"]),
            "ubs_borrowing_base": money_short(row["UBSBB"]),
            "notes": as_is(row["Notes"]),
        })
        counts["written"] += 1

    return SeedResult(seed_rows, counts)


def upsert_facilities(fac_data: list[list[str]], by_acct: dict[str, int],
                      export: list[dict]) -> tuple[list[list[str]], dict[str, str]]:
    """Facilities from the Agent Bank Summary: bank_status := Active if the account appears in the
    export else Inactive (Active also gets collateral_date := BBDate), overriding the report's own
    FacilityStatus. Export accounts not listed in the report are manufactured as placeholder
    Inactive facilities ("Unknown" bank, name=FndName, account_number=AccountID) so every LP record
    can seed. Returns (rows, account_number -> resolved facility name)."""
    bbdate_by_acct: "OrderedDict[str, str]" = OrderedDict()
    fnd_by_acct: "OrderedDict[str, str]" = OrderedDict()
    for row in export:
        acct = (row["AccountID"] or "").strip()
        if acct and acct not in bbdate_by_acct:
            bbdate_by_acct[acct] = iso_date(row["BBDate"])
            fnd_by_acct[acct] = as_is(row["FndName"])

    out = [list(r) for r in fac_data]
    name_by_acct: dict[str, str] = {}
    used_norm = {_norm(r[1]) for r in out if len(r) > 1 and r[1].strip()}

    # Reported facilities. Status is resolved per row off that row's own account; only the owning
    # row (by_acct) lends its name to the LP seeds.
    for acct, idx in by_acct.items():
        name_by_acct[acct] = out[idx][1].strip()
    for row in out:
        acct = row[2].strip()
        if acct in bbdate_by_acct:
            row[5] = "Active"
            if bbdate_by_acct[acct]:
                row[8] = bbdate_by_acct[acct]
        else:
            row[5] = "Inactive"

    # Orphan export accounts -> placeholder Inactive facilities, name disambiguated by AccountID
    # when already in use, so each account survives as its own facility.
    for acct in bbdate_by_acct:
        if acct in by_acct:
            continue
        name = fnd_by_acct[acct] or f"Unknown Facility {acct}"
        if _norm(name) in used_norm:
            name = f"{name} ({acct})"
        used_norm.add(_norm(name))
        out.append(["Unknown", name, acct, "", "", "Inactive", "", "", bbdate_by_acct[acct]])
        name_by_acct[acct] = name

    return out, name_by_acct


def write_csv(path: Path, header: list[str], rows: list[dict]) -> None:
    with path.open("w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh, quoting=csv.QUOTE_ALL)
        w.writerow(header)
        for r in rows:
            w.writerow([r[c] for c in header])


def write_facilities(path: Path, rows: list[list[str]]) -> None:
    with path.open("w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh, quoting=csv.QUOTE_ALL)
        w.writerow(FACILITY_COLS)
        for r in rows:
            w.writerow(r[: len(FACILITY_COLS)])


def main() -> int:
    export_path = Path(EXPORT_FILE)
    abs_path = Path(AGENT_BANK_SUMMARY_FILE)
    out_dir = Path(OUT_DIR)
    ref_dir = Path(REFERENCE_DIR)
    out_dir.mkdir(parents=True, exist_ok=True)

    ref = load_references(ref_dir)
    # Small file first, so a missing/renamed report fails before the big export is parsed.
    fac_data, by_acct = read_agent_bank_summary(abs_path)
    export = read_export(export_path)

    # Facilities first: manufactures placeholders for orphan accounts and returns the
    # account -> facility name map the seed uses, so every LP record resolves to a facility.
    fac_rows, name_by_acct = upsert_facilities(fac_data, by_acct, export)
    master_rows = build_master(export, ref)
    sr = build_seed(export, name_by_acct, ref)

    # Clear data/out/ so it holds only this run's three CSVs. Done after all inputs are read, so a
    # failed read leaves the last good outputs in place. import/ and reference/ are never touched.
    for old in out_dir.iterdir():
        if old.is_file():
            old.unlink()

    write_csv(out_dir / "lp_master.csv", MASTER_COLS, master_rows)
    write_csv(out_dir / "lp_facility_seeds.csv", SEED_COLS, sr.rows)
    write_facilities(out_dir / "facilities.csv", fac_rows)

    # Retention check: lp_facility_seeds.csv must carry 100% of the export's records.
    dropped = (sr.counts["skipped_duplicate_pair"] + sr.counts["skipped_blank_investor"]
               + sr.counts["skipped_no_facility"])
    print(f"export rows            : {len(export)}")
    print(f"lp_facility_seeds rows : {sr.counts['written']}")
    print(f"dropped                : {dropped} "
          f"(duplicate pair {sr.counts['skipped_duplicate_pair']}, "
          f"blank investor {sr.counts['skipped_blank_investor']}, "
          f"no facility {sr.counts['skipped_no_facility']})")
    retained = f"{sr.counts['written'] / len(export) * 100:.2f}%" if export else "n/a"
    print(f"retained               : {retained}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
