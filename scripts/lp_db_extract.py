#!/usr/bin/env python3
r"""
LP DB Export -> seed extract  (one-off, day-1 bootstrap; see pe-sub-docs/LP_DB_EXTRACT_DESIGN.md)

Standalone utility. Reads the LP DB Export (XLSX) plus the Agent Bank Summary report (XLSX) and
writes ALL outputs into one directory (pe-sub-jobs/data/out/) — the directory pe-sub-jobs reads
directly on startup (see ingest.* in application.yml). Review the output, then restart pe-sub-jobs
(or trigger the /jobs endpoints) to load. pe-sub-jobs/data/mock is untouched dev fixture data.

The Agent Bank Summary report (data/import/AgentBankSummaryRpt.xlsx) is the ORIGINAL SOURCE for
loan-level facility attributes — agent bank, loan amount, maturity date and the agent-reported
status — so the facility list is derived from the agent's own report rather than a hand-maintained
CSV. It is a banded print layout, not a table: see read_agent_bank_summary.

Normalization against editable reference lists in pe-sub-jobs/data/reference/ (seeded from Config):
  * Investor Type  -> supported list (investor_types.csv + investor_type_aliases.csv); unmatched
                      values are PASSED THROUGH unchanged (record kept) and counted on the console.
  * Agent LP Category (export "Classification") -> canonical Agent LP Classification
                      (agent_lp_categories.csv); unmatched passed through and counted.
  * UBS classification: the export has no UBS class column, so it is derived from the LP's own
                      attributes (agency ratings, pension assets, NAV, AUM, HNW/SPV flags, agent
                      category) using the Borrowing Base Criteria Matrix (bb_criteria_matrix.csv,
                      transcribed from pe-sub-docs/BB_CRITERIA_DESIGN.md) — see classify_ubs. The
                      Other Institutional catch-all makes this total, so it always resolves; the
                      row's UBSAR/UBSCL are cross-checked against the matrix's expected advance
                      rate (funded-split aware) / conc limit and deviations counted on the console.
  * Advance rates (the Floor Map): UBSAR and AgentAR are slotted into the platform's discrete rate
                      groups via rate_floor_map.csv — >=90 -> 90, 75-89.9 -> 75, 65-74.9 -> 65,
                      50-64.9 -> 50, <50 -> 0 — before seeding (and before the matrix cross-check).

Outputs — EXACTLY these three ingestion files, always in pe-sub-jobs/data/out/:
  * lp_master.csv               - one distinct golden LP per investor_name (best-record consolidation)
  * lp_facility_seeds.csv       - per (facility, investor) LP-record seed rows
  * facilities.csv              - the Agent Bank Summary report + bank_status Active/Inactive +
                                  collateral_date := BBDate, PLUS a manufactured Inactive placeholder
                                  ("Unknown" bank) for any export account the report does not list —
                                  so 100% of LP records seed (no rejects)
  No review reports and no summary file are written: data/out/ holds only what pe-sub-jobs ingests.
  The run's counts (unmatched Investor Types / Agent LP Categories, UBS classification mix, matrix
  variance) are still computed and printed to the console.

Design decisions this script implements (LP_DB_EXTRACT_DESIGN.md):
  D2  (REVISED) full-column seed: lp_facility_seeds.csv carries EVERY per-LP export column
      (facility-level AccountID/FndName/BBDate excluded), matching the full lp_records insert.
      ubs_cls is derived per row from that row's attributes via reference/bb_criteria_matrix.csv.
      LP Master still holds the consolidated golden profile, used only as fallback for blanks.
  D3  facilities matched on AccountID: a facility from the Agent Bank Summary whose account is in
      the export -> Active (else Inactive) — this match-derived value OVERRIDES the report's own
      FacilityStatus column. An export account with NO facility in the report -> a manufactured
      "Unknown"-bank Inactive placeholder, so its LP records still seed (100% insertion; no rejects).
  D4  AccountID is the join key. Real data is 1:1 (AccountID<->fund); a duplicate FndName across
      orphan accounts (a sim artifact) is disambiguated with the AccountID to keep names unique.
  D6  LP Master is cleared and repopulated with one distinct LP per investor_name, each field
      chosen by majority vote across that investor's rows (see build_master / _best).
  D7  collateral_date := BBDate  (label unchanged; wired as "Last BB Run Date").
  D8  (REVISED) seed ubs_default_adv_rate<-UBSAR (Floor-Mapped to its rate group) and
      ubs_default_conc_limit<-UBSCL; agent_rate/ubs_rate on the seed are likewise Floor-Mapped.
      ubs_classification is derived from the LP's attributes via the BB Criteria Matrix
      (reference/bb_criteria_matrix.csv; classify_ubs waterfall) — NOT from UBSAR: under the
      funded-split matrix a rate no longer identifies a class (90% maps to six classes). UBSAR/UBSCL
      are instead cross-checked against the matrix and the deviations reported on the console.
  D10 no hard-fail: bad rows are skipped and counted; a dirty file never aborts the run.

Best-record consolidation (see build_master / _best): an investor appears on many rows (one per
facility) whose attributes may disagree. Each golden field is chosen INDEPENDENTLY by majority vote
over the non-blank values across those rows (ties -> first occurrence); blanks do not vote, so a
value present in only some rows still fills the gaps. This yields the most-agreed profile per field
rather than trusting one arbitrary row. (The generator keeps an LP's attributes constant across its
rows and the chaos monkey drifts individual cells, so the vote genuinely repairs per-row noise.)

Dirty input is expected, not exceptional: the simulated export is generated by lp_db_generate.py,
whose chaos monkey (pe-sub-docs/"AI Chaos Monkey for Data Quality.md") writes the XLSX at realistic
manual-entry quality — name drift, 'A minus' ratings, unit mix-ups, NAV range strings, categorical
drift on Investor Type / agent Classification. This script therefore reads whatever file it is
given AS-IS (no degradation, no cleaning toggle) and relies on its tolerant normalizers, parsers
and review reports to cope — the same posture it needs for a real LP DB export. The generator's
ground-truth log ('<export name>.chaos_log.csv' next to the XLSX) can be used to verify what the
normalizers absorbed.

Usage (NO command-line arguments):
    1. Edit the EXPORT_FILE variable near the top of this file to point at the export to process.
    2. Run it — from any working directory:
           python pe-sub-jobs/scripts/lp_db_extract.py
       or, on Windows:
           powershell -ExecutionPolicy Bypass -File pe-sub-jobs/scripts/lp_db_extract.ps1
       or, on Linux/macOS:
           bash pe-sub-jobs/scripts/lp_db_extract.sh
    Re-point EXPORT_FILE and re-run for the next file. Outputs always land in pe-sub-jobs/data/out/.
"""
from __future__ import annotations

import csv
import difflib
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

# Everything the tool needs lives inside the pe-sub-jobs project. Paths are anchored to this file's
# location (pe-sub-jobs/scripts/) so the script runs from any working directory and never reaches
# outside the project.
SCRIPT_DIR  = Path(__file__).resolve().parent          # pe-sub-jobs/scripts/
JOBS_ROOT = SCRIPT_DIR.parent                           # pe-sub-jobs/
DATA_DIR = JOBS_ROOT / "data"

# ============================================================================================
#  EDIT THIS for each run — the LP DB Export to process.
#  There are NO command-line arguments: change the line below and re-run the script
#  (python pe-sub-jobs/scripts/lp_db_extract.py, or the .ps1/.sh wrapper). Use an absolute path, or
#  a path relative to the pe-sub-jobs project root. The default is pe-sub-jobs/data/import/ — drop
#  the export there, or repoint this line. Everything else below is a stable default.
# ============================================================================================
EXPORT_FILE = DATA_DIR / "import" / "LP DB Export 2026.06.25.xlsx"

# Stable defaults (rarely changed) — all within the pe-sub-jobs project:
# The Agent Bank Summary report is the ORIGINAL SOURCE for loan-level facility attributes (agent
# bank, loan amount, maturity date, agent-reported status). Read-only; parsed by
# read_agent_bank_summary, which copes with the report's banded layout and its dirty rows.
AGENT_BANK_SUMMARY_FILE = DATA_DIR / "import" / "AgentBankSummaryRpt.xlsx"
OUT_DIR = DATA_DIR / "out"              # ALL outputs land here; never writes into the app tree
REFERENCE_DIR = DATA_DIR / "reference"  # Config-seeded normalization lists

# --- source column order (must match the export header exactly) ----------------------------
SRC_COLS = [
    "AccountID", "FndName", "InvestorName", "Parent", "SPV", "InvestorType", "Region", "HQ",
    "InstitutionalHNW", "InvestmentGrade", "Classification", "Notes", "SP", "Moodys", "Fitch",
    "AUM", "NAV", "PensionAssets", "FundingRatio", "UBSAR", "AgentAR", "Commitments",
    "PercentOfCommitments", "Called", "Uncalled", "PercentOfUncalled", "CalledPercent",
    "AgentCL", "UBSCL", "AgentBB", "UBSBB", "BBDate",
]

# CSV header (column) orders required by the pe-sub-jobs FlatFileItemReaders.
MASTER_COLS = [
    "investor_name", "parent", "spv", "high_quality", "investor_type", "institutional_or_hnw",
    "region_location", "investment_grade", "sp_rating", "moodys_rating", "fitch_rating", "aum", "nav", "pension_assets",
    "funding_ratio", "ubs_lp_category", "ubs_default_advance_rate", "ubs_default_concentration_limit",
    "notes",
]
SEED_COLS = [
    # Legacy 7 first (order preserved so an old 7-column file still parses under the
    # non-strict reader), then every remaining per-LP export column so the seed matches
    # the full lp_records insert. Facility-level columns (AccountID, FndName, BBDate)
    # are excluded — they drive facilities.csv, not the LP record.
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

# Agent Bank Summary report column layout (must match the report header exactly). Index 4 is an
# unnamed spacer the report uses to hold its subtotal amounts, so it carries no header label.
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
    """LP-Size passthrough: keep the raw extracted value verbatim (UI does the formatting)."""
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
    """Decimal rate/limit passthrough as a trimmed string ('0.41'), matching lp_master.csv."""
    if blank(v):
        return ""
    try:
        return _trim(Decimal(str(v)))
    except InvalidOperation:
        return str(v).strip()


def money_short(v) -> str:
    """Exact short-currency dollars ($484M, $314.6M, $8B). No rounding; falls back to a
    full grouped string only if a short form cannot represent the value losslessly."""
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
    """Parse a feed date to ISO (YYYY-MM-DD); '' if unparseable. The LP DB Export carries text
    dates (M/D/YYYY BBDate); the Agent Bank Summary carries real Excel dates, which openpyxl
    hands back as datetime objects — take those directly rather than round-tripping the string."""
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


# --- reference lists (Investor Type, Agent LP Category, UBS classification) -----------------
def _norm(s) -> str:
    """Case/punctuation-insensitive key: lowercase, non-alphanumerics collapsed to one space."""
    return re.sub(r"[^a-z0-9]+", " ", str(s or "").lower()).strip()


@dataclass
class Reference:
    itype_canonical: list[str]              # supported Investor Types (display order)
    itype_lookup: dict[str, str]            # norm(value/alias) -> canonical Investor Type
    agent_canonical: list[str]              # canonical Agent LP Categories (AGENT_CLS_OPTS)
    agent_lookup: dict[str, str]            # norm(alias) -> canonical Agent LP Category
    matrix_rated: dict[str, tuple[float, float, float]]    # band -> (conc_limit, ar_lt40, ar_gte40) %
    matrix_classes: dict[str, tuple[float, float, float]]  # cls  -> (conc_limit, ar_lt40, ar_gte40) %
    rate_floors: list[tuple[float, float]]  # (min_rate_pct, group_pct), sorted highest-min first


def _read_reference_rows(path: Path) -> list[list[str]]:
    """Read a reference CSV/TXT, dropping blank lines and '#'/'##' comment lines. Returns rows
    (each a list of trimmed cells). The caller decides whether the first surviving row is a header."""
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
    # Investor Types: canonical list + case/punct-insensitive lookup (self + aliases).
    itype_canonical = [r[0] for r in _read_reference_rows(ref_dir / "investor_types.csv")[1:]]  # drop header
    itype_lookup = {_norm(c): c for c in itype_canonical}
    alias_rows = _read_reference_rows(ref_dir / "investor_type_aliases.csv")[1:]  # drop header
    for row in alias_rows:
        if len(row) >= 2 and row[1]:
            itype_lookup[_norm(row[0])] = row[1]

    # Agent LP Categories: alias -> canonical. Canonical set = the distinct RHS values.
    agent_rows = _read_reference_rows(ref_dir / "agent_lp_categories.csv")[1:]  # drop header
    agent_lookup = {_norm(r[0]): r[1] for r in agent_rows if len(r) >= 2 and r[1]}
    agent_canonical = list(dict.fromkeys(r[1] for r in agent_rows if len(r) >= 2 and r[1]))

    # BB Criteria Matrix: the same bb_criteria_matrix config key Run Shadow BB resolves against
    # (BB_CRITERIA_DESIGN.md tab 2). "Rated Investor" rows carry a rating band; every other class
    # carries a blank band and is keyed by its classification label.
    matrix_rated: dict[str, tuple[float, float, float]] = {}
    matrix_classes: dict[str, tuple[float, float, float]] = {}
    for row in _read_reference_rows(ref_dir / "bb_criteria_matrix.csv")[1:]:  # drop header
        if len(row) < 5 or not row[0]:
            continue
        try:
            vals = (float(row[2]), float(row[3]), float(row[4]))
        except ValueError:
            continue
        if row[1]:
            matrix_rated[row[1]] = vals
        else:
            matrix_classes[row[0]] = vals

    # The Floor Map: raw advance rate -> discrete rate group (applies to UBSAR AND AgentAR).
    # A rate takes the group of the highest 'min' it meets or exceeds; the min=0 floor makes it
    # total (everything below 50% lands in the 0% group).
    rate_floors: list[tuple[float, float]] = []
    for row in _read_reference_rows(ref_dir / "rate_floor_map.csv")[1:]:  # drop header
        if len(row) < 2:
            continue
        try:
            rate_floors.append((float(row[0]), float(row[1])))
        except ValueError:
            continue
    rate_floors.sort(key=lambda t: t[0], reverse=True)

    return Reference(itype_canonical, itype_lookup, agent_canonical, agent_lookup,
                     matrix_rated, matrix_classes, rate_floors)


def _suggest(raw: str, canonical: list[str]) -> str:
    """Closest canonical value for an unmatched raw string (fuzzy hint only; never auto-applied)."""
    norms = {_norm(c): c for c in canonical}
    hit = difflib.get_close_matches(_norm(raw), list(norms), n=1, cutoff=0.6)
    return norms[hit[0]] if hit else ""


def map_investor_type(raw, ref: Reference) -> tuple[str, bool]:
    """(canonical, True) when the value maps to a supported Investor Type; otherwise the ORIGINAL
    value is passed through with (value, False) — the record is always kept, never dropped."""
    s = as_is(raw)
    if not s:
        return "", True
    canon = ref.itype_lookup.get(_norm(s))
    return (canon, True) if canon else (s, False)


def map_agent_cls(raw, ref: Reference) -> tuple[str, bool]:
    """(canonical Agent LP Category, True) when mapped; else the original value + False."""
    s = as_is(raw)
    if not s:
        return "", True
    canon = ref.agent_lookup.get(_norm(s))
    return (canon, True) if canon else (s, False)


# --- UBS classification via the BB Criteria Matrix -----------------------------------------
# Canonical class labels (bb_criteria_matrix.csv / BB_CRITERIA_DESIGN.md tab 3, Q3 typo fixed).
CLS_RATED = "Rated Investor"
CLS_CP5 = "Corp Pension > $5Bn Assets"
CLS_CP1 = "Corp Pension > $1Bn Assets"
CLS_NAV1 = "Unrated NAV > $1Bn"
CLS_FOF = "FoF & Other > $10Bn AUM"
CLS_OTHER = "Other Institutional"
CLS_HNW_FEEDER = "HNW Feeder (acceptable)"
CLS_HNW = "HNW (acceptable)"
CLS_EXCLUDED = "Excluded"

# Unified ordinal notch scale (1 = strongest), accepting both S&P/Fitch and Moody's notation —
# the Q2 "eligible rating" waterfall needs all three agencies on one scale.
_SP_SCALE = ["AAA", "AA+", "AA", "AA-", "A+", "A", "A-", "BBB+", "BBB", "BBB-",
             "BB+", "BB", "BB-", "B+", "B", "B-", "CCC+", "CCC", "CCC-", "CC", "C", "D"]
_MDY_SCALE = ["Aaa", "Aa1", "Aa2", "Aa3", "A1", "A2", "A3", "Baa1", "Baa2", "Baa3",
              "Ba1", "Ba2", "Ba3", "B1", "B2", "B3", "Caa1", "Caa2", "Caa3", "Ca", "C"]
_RATING_NOTCH = {r.lower(): i + 1 for i, r in enumerate(_SP_SCALE)}
for _i, _r in enumerate(_MDY_SCALE):
    _RATING_NOTCH.setdefault(_r.lower(), _i + 1)
_NOT_RATED = {"nr", "n/r", "notrated", "n/a", "na", "wr", "unrated", "none"}


def rating_notch(v) -> int | None:
    """Normalize one agency rating to its ordinal notch. Tolerates the manual-entry drift real
    exports carry: 'A minus' -> 'A-', case noise, 'Not Rated'/'N/R'/'' -> None."""
    s = str(v or "").strip().lower().replace("minus", "-").replace("plus", "+").replace(" ", "")
    if not s or s in _NOT_RATED:
        return None
    return _RATING_NOTCH.get(s)


def eligible_band(sp, mdy, fitch) -> str:
    """Q2 eligible-rating waterfall -> rating band: three ratings -> median, two -> the lower
    (conservative), one -> as-is; sub-BBB- clamps to the BBB band (a rated LP stays Rated).
    Returns '' when no agency rating is usable (the LP is unrated)."""
    notches = sorted(n for n in (rating_notch(v) for v in (sp, mdy, fitch)) if n is not None)
    if not notches:
        return ""
    n = notches[1] if len(notches) >= 2 else notches[0]  # median of 3 / lower of 2 / single
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
    """Tolerant size parser -> absolute dollars, for the export's free-text AUM/NAV/PensionAssets
    ('$14.9B', '240B', '1.33. tn', '$21 bn+', '394.6667', ranges like '1-5M' / '500M - 2Bn' /
    '$2-6Bn', thresholds '>5B' / '<100M'). A range takes the conservative LOW end; a bare number
    is absolute dollars. Returns None when nothing numeric can be read."""
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


def classify_ubs(row: dict, ref: Reference) -> tuple[str, str, str]:
    """Derive the UBS LP Classification from the LP's OWN attributes per the BB Criteria Matrix
    taxonomy (bb_criteria_matrix.csv; BB_CRITERIA_DESIGN.md §4/§6). Waterfall:
      1. agent category 'Ineligible Investor' -> Excluded (the matrix has no attribute criterion
         for exclusion; agent ineligibility is the only available signal, per AGENT_CLS_UBS_MAP);
      2. any usable agency rating -> Rated Investor, band via the Q2 waterfall (sub-IG clamps to
         BBB and stays Rated — never reclassified out of the Rated bucket);
      3. HNW (InstitutionalHNW flag, or agent 'Designated PWM' per Q5) -> HNW Feeder (acceptable)
         when the vehicle is an SPV (feeder), else HNW (acceptable);
      4. Pension assets > $5Bn / > $1Bn -> the two Corp Pension classes (Q4 boundaries);
      5. NAV > $1Bn -> Unrated NAV > $1Bn;
      6. FoF/asset-manager investor type with AUM > $10Bn -> FoF & Other > $10Bn AUM;
      7. catch-all -> Other Institutional (makes the mapping total — it always resolves).
    Returns (classification, rating_band ('' unless Rated), path slug for the run counters)."""
    agent_canon, matched = map_agent_cls(row["Classification"], ref)
    agent = agent_canon if matched else ""
    if agent == "Ineligible Investor":
        return CLS_EXCLUDED, "", "excluded_agent_ineligible"
    band = eligible_band(row["SP"], row["Moodys"], row["Fitch"])
    if band:
        return CLS_RATED, band, "rated_band"
    if _norm(row["InstitutionalHNW"]) == "hnw" or agent == "Designated PWM":
        if _is_yes(row["SPV"]):
            return CLS_HNW_FEEDER, "", "hnw_feeder_spv"
        return CLS_HNW, "", "hnw"
    pension = parse_money_low(row["PensionAssets"])
    if pension is not None and pension > 5e9:
        return CLS_CP5, "", "pension_gt_5bn"
    if pension is not None and pension > 1e9:
        return CLS_CP1, "", "pension_gt_1bn"
    nav = parse_money_low(row["NAV"])
    if nav is not None and nav > 1e9:
        return CLS_NAV1, "", "nav_gt_1bn"
    itype = map_investor_type(row["InvestorType"], ref)[0]
    aum = parse_money_low(row["AUM"])
    if itype in ("Fund of Funds", "Hedge Fund") and aum is not None and aum > 10e9:
        return CLS_FOF, "", "fof_aum_gt_10bn"
    return CLS_OTHER, "", "other_institutional"


def funded_fraction(row: dict) -> float | None:
    """The LP's funded level (Q1: per-LP called ÷ commitment) driving the matrix's 40% advance-rate
    split. Prefers the export's CalledPercent; falls back to Called ÷ Commitments."""
    v = row["CalledPercent"]
    if not blank(v):
        try:
            return float(v)
        except (TypeError, ValueError):
            pass
    try:
        commit = float(row["Commitments"])
        return float(row["Called"]) / commit if commit else None
    except (TypeError, ValueError, ZeroDivisionError):
        return None


def matrix_expected(cls: str, band: str, funded: float | None,
                    ref: Reference) -> tuple[float, float] | None:
    """The matrix's expected (advance_rate_pct, conc_limit_pct) for a classification: Rated resolves
    by band, everything else by class label; the AR column follows the 40% funded split (>= 40%
    inclusive; unknown funded level takes the conservative < 40% column)."""
    row = ref.matrix_rated.get(band) if cls == CLS_RATED else ref.matrix_classes.get(cls)
    if row is None:
        return None
    conc, ar_lt40, ar_gte40 = row
    return (ar_gte40 if funded is not None and funded >= 0.4 else ar_lt40), conc


def rate_group_pct(v, ref: Reference) -> float | None:
    """The Floor Map (reference/rate_floor_map.csv): slot a raw advance-rate fraction (0.93) into
    its discrete rate group in percent (90.0) — the group of the highest 'min' the rate meets or
    exceeds (>=90 -> 90, 75-89.9 -> 75, 65-74.9 -> 65, 50-64.9 -> 50, <50 -> 0). Applies to BOTH
    UBSAR and AgentAR. None when the value is missing/unparseable."""
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
    """Floor-mapped rate as the seed's percent string ('90%'); '' when unmapped."""
    group = rate_group_pct(v, ref)
    return "" if group is None else _trim(Decimal(str(group))) + "%"


def floor_rate_frac(v, ref: Reference) -> str:
    """Floor-mapped rate as the LP-Master fraction string ('0.9'); '' when unmapped."""
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
    header = list(next(rows_iter))
    if header != SRC_COLS:
        raise SystemExit(
            f"Export header in '{ws.title}' does not match the expected schema.\n"
            f"  expected: {SRC_COLS}\n  found:    {header}"
        )
    return [dict(zip(SRC_COLS, r)) for r in rows_iter]


def read_agent_bank_summary(path: Path) -> tuple[list[list[str]], dict[str, int], Counter]:
    """Read the Agent Bank Summary report — the ORIGINAL SOURCE for loan-level facility attributes
    — into FACILITY_COLS-shaped rows. Returns (rows, account_number -> row index, counts).

    The report is a banded print layout, not a flat table:
      * the agent bank appears once on its own group-header row (Agent set, Borrower blank) and is
        CARRIED DOWN onto every facility row beneath it, which leave the Agent cell blank;
      * each group ends with an 'AccessTotalsLoanAmount:' subtotal row (and the sheet with a grand
        total) that carries no facility — those are skipped.

    Dirty input is expected, not exceptional (D10) — the report is an operational extract, so:
      * a row repeating an (AccountNumber, Borrower) pair already seen is a reprint and is dropped;
      * a Borrower name already taken by a DIFFERENT account is disambiguated with its AccountID,
        because facility identity is the name (D4) — same convention upsert_facilities uses;
      * one AccountNumber listed against two different borrowers keeps the FIRST row as the owner
        of the LP join; the later row survives as its own facility but matches no export rows.
    All four are counted and reported on the console rather than aborting the run.

    Two FACILITY_COLS fields are not in the report: ubs_participation (never reported by the agent
    bank — left blank, which the ingest maps to null) and collateral_date (filled from the LP DB
    Export's BBDate by upsert_facilities). bank_status/bank_status_date are taken from the report's
    FacilityStatus columns, but D3 then overrides bank_status with the export-match result."""
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

    counts: Counter = Counter()
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
            counts["total_rows"] += 1
            continue
        if text[0] and not text[1]:                       # agent group header -> carry down
            agent = text[0]
            counts["agents"] += 1
            continue
        name, acct = text[1], text[2]
        if not name or not acct:
            counts["skipped_incomplete"] += 1
            continue
        if (acct, _norm(name)) in seen_pair:              # reprint of a row already taken
            counts["duplicate_row"] += 1
            continue
        seen_pair.add((acct, _norm(name)))
        if _norm(name) in used_norm:
            name = f"{name} ({acct})"
            counts["renamed_duplicate_name"] += 1
        used_norm.add(_norm(name))
        # agent_bank, name, account_number, loan_amount, maturity_date, bank_status,
        # bank_status_date, ubs_participation, collateral_date
        # The row's own Agent cell wins if the report ever fills it; otherwise the carried-down
        # group header. "Unknown" keeps FacilityRowProcessor's non-blank agent_bank rule satisfied.
        data.append([text[0] or agent or "Unknown", name, acct, text[3], iso_date(cells[5]),
                     text[6], iso_date(cells[7]), "", ""])
        if acct in by_acct:
            counts["duplicate_account"] += 1               # first row keeps the LP join
        else:
            by_acct[acct] = len(data) - 1
        counts["facilities"] += 1

    return data, by_acct, counts


@dataclass
class MasterResult:
    rows: list[dict]
    itype_unmatched: Counter               # raw InvestorType not in reference -> occurrences
    cls_counts: Counter                    # derived UBS classification -> distinct LPs


def _best(rows: list[dict], field: str) -> str:
    """Consolidate one field across all of an investor's rows into its single 'best' value:
    the MOST FREQUENTLY reported non-blank value (majority vote). Ties and all-equal-frequency
    fall back to the first occurrence, because Counter preserves insertion order. Blank/'' cells
    do not vote, so a value present in only some rows still wins over the gaps — which is exactly
    what fills sparse/typo'd real-world data. Returns '' only when every row is blank for it."""
    tally: Counter = Counter()
    for r in rows:
        v = as_is(r[field])
        if v:
            tally[v] += 1
    return tally.most_common(1)[0][0] if tally else ""


def build_master(export: list[dict], ref: Reference) -> MasterResult:
    """Build one distinct golden LP per investor_name by CONSOLIDATING all of that investor's rows
    (e.g. the ~15 occurrences across facilities). Each attribute is chosen independently by majority
    vote over the non-blank values (`_best`) — not a single arbitrary row — so the profile is the
    most-agreed value per field and gaps/typos in a minority of rows are outvoted. Investor Type is
    voted on its *canonical* mapping; UBS classification is derived from the consolidated best
    attributes via the BB Criteria Matrix waterfall (classify_ubs), so the golden class reflects
    the golden ratings/sizes."""
    groups: "OrderedDict[str, list[dict]]" = OrderedDict()
    itype_seen: Counter = Counter()
    for row in export:
        name = (row["InvestorName"] or "").strip()
        if not name:
            continue
        itype_seen[as_is(row["InvestorType"])] += 1
        groups.setdefault(name, []).append(row)

    # Distinct raw Investor Types (across every row) that do not resolve to the supported list.
    itype_unmatched: Counter = Counter()
    for raw, cnt in itype_seen.items():
        if raw and not map_investor_type(raw, ref)[1]:
            itype_unmatched[raw] += cnt

    master_rows: list[dict] = []
    cls_counts: Counter = Counter()
    for name, rows in groups.items():
        # Investor Type: majority vote over the CANONICAL mapping (so "Other Institutional Investors"
        # and "Other Institutional" count together; unmatched values vote as their passed-through text).
        itype_votes: Counter = Counter()
        for r in rows:
            v = as_is(r["InvestorType"])
            if v:
                itype_votes[map_investor_type(v, ref)[0]] += 1
        investor_type = itype_votes.most_common(1)[0][0] if itype_votes else ""

        best_ubsar = _best(rows, "UBSAR")
        # Golden UBS classification: the matrix waterfall over the CONSOLIDATED best attributes,
        # so the class agrees with the golden ratings/sizes rather than one arbitrary row.
        best_attrs = {c: _best(rows, c) for c in (
            "Classification", "SP", "Moodys", "Fitch", "InstitutionalHNW", "SPV",
            "PensionAssets", "NAV", "AUM", "InvestorType")}
        ubs_cls, _band, _path = classify_ubs(best_attrs, ref)
        cls_counts[ubs_cls] += 1

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
            "aum": _best(rows, "AUM"),          # LP Size passthrough (D2)
            "nav": _best(rows, "NAV"),
            "pension_assets": _best(rows, "PensionAssets"),
            "funding_ratio": pct(_best(rows, "FundingRatio")),
            "ubs_lp_category": ubs_cls,      # BB Criteria Matrix waterfall over best attributes
            "ubs_default_advance_rate": floor_rate_frac(best_ubsar, ref),  # Floor-Map group
            "ubs_default_concentration_limit": dec_str(_best(rows, "UBSCL")),
            "notes": _best(rows, "Notes"),
        })

    return MasterResult(master_rows, itype_unmatched, cls_counts)


MATRIX_TOLERANCE_PP = 2.5   # UBSAR/UBSCL deviation (percentage points) beyond which a row counts
                            # as off-matrix in the console run summary


@dataclass
class SeedResult:
    rows: list[dict]
    counts: Counter
    agent_unmatched: Counter               # raw Classification not in reference -> occurrences
    cls_counts: Counter                    # derived per-row UBS classification -> rows
    variance: dict                         # (cls, band, funded_bucket) -> aggregate deviation stats


def _variance_add(variance: dict, cls: str, band: str, funded: float | None,
                  ubsar, ubscl, ref: Reference) -> None:
    """Cross-check one seed row's UBSAR/UBSCL against the matrix's expected AR/CL for its derived
    (classification, band, funded bucket), aggregating deviations per group — the sim's randomized
    rates make a per-row count useless, and a clean feed collapses this to nothing."""
    expected = matrix_expected(cls, band, funded, ref)
    if expected is None:
        return
    exp_ar, exp_cl = expected
    bucket = "gte40" if funded is not None and funded >= 0.4 else "lt40"
    agg = variance.setdefault((cls, band, bucket), {
        "rows": 0, "exp_ar": exp_ar, "exp_cl": exp_cl,
        "ar_n": 0, "ar_sum": 0.0, "ar_beyond": 0, "ar_max": 0.0,
        "cl_n": 0, "cl_sum": 0.0, "cl_beyond": 0, "cl_max": 0.0,
    })
    agg["rows"] += 1
    for key, exp, raw in (("ar", exp_ar, ubsar), ("cl", exp_cl, ubscl)):
        try:
            pct_val = float(raw) * 100
        except (TypeError, ValueError):
            continue
        agg[f"{key}_n"] += 1
        agg[f"{key}_sum"] += pct_val
        dev = abs(pct_val - exp)
        agg[f"{key}_max"] = max(agg[f"{key}_max"], dev)
        if dev > MATRIX_TOLERANCE_PP:
            agg[f"{key}_beyond"] += 1


def build_seed(export: list[dict], name_by_acct: dict[str, str], ref: Reference) -> SeedResult:
    """Per (facility, investor) seed rows carrying EVERY per-LP export column (only the
    facility-level AccountID/FndName/BBDate are excluded), so the seed matches the full
    lp_records insert. Every export account resolves to a facility (existing or a manufactured
    placeholder), so nothing is rejected for a missing facility. Normalises the export's
    Classification (Agent LP Category) to the canonical Agent LP Classification and Investor Type
    to the supported list (unmatched passed through); ubs_cls is derived PER ROW from that row's
    attributes via the BB Criteria Matrix waterfall (classify_ubs), and the row's UBSAR/UBSCL are
    cross-checked against the matrix (aggregated into the variance report)."""
    seen: set[tuple[str, str]] = set()
    seed_rows: list[dict] = []
    counts = Counter()
    agent_unmatched: Counter = Counter()      # raw Classification not in reference -> occurrences
    cls_counts: Counter = Counter()
    variance: dict = {}

    for row in export:
        acct = (row["AccountID"] or "").strip()
        investor = (row["InvestorName"] or "").strip()
        fac_name = name_by_acct.get(acct)
        if fac_name is None:
            # Should not happen: upsert_facilities manufactures a facility for every export account.
            counts["skipped_no_facility"] += 1
            continue
        if not investor:
            counts["skipped_blank_investor"] += 1   # investor_name is the LP key; cannot insert blank
            continue
        key = (fac_name, investor)
        if key in seen:
            counts["skipped_duplicate_pair"] += 1
            continue
        seen.add(key)

        agent_cls, matched = map_agent_cls(row["Classification"], ref)
        if not matched and agent_cls:
            agent_unmatched[agent_cls] += 1   # agent_cls holds the passed-through original here

        ubs_cls, band, _path = classify_ubs(row, ref)
        cls_counts[ubs_cls] += 1
        # Cross-check on the Floor-Mapped rate group, not the raw rate — a raw 0.93 belongs to the
        # 90% group and should compare clean against the matrix's 90, not flag a 3pp jitter.
        ubsar_group = rate_group_pct(row["UBSAR"], ref)
        _variance_add(variance, ubs_cls, band, funded_fraction(row),
                      None if ubsar_group is None else ubsar_group / 100, row["UBSCL"], ref)

        seed_rows.append({
            "facility_name": fac_name,
            "investor_name": investor,
            "capital_commitment": money_short(row["Commitments"]),
            "uncalled_capital": money_short(row["Uncalled"]),
            "agent_lp_category": agent_cls,           # normalised to Agent LP Classification (or passthrough)
            "agent_advance_rate": floor_rate_pct(row["AgentAR"], ref),  # Floor-Map group
            "agent_concentration_limit": pct(row["AgentCL"]),
            # Full per-LP column set (row-level values, no LP-Master consolidation):
            "parent": as_is(row["Parent"]),
            "spv": yn_bool(row["SPV"]),
            "high_quality": yn_bool(row["HQ"]),
            "investor_type": map_investor_type(row["InvestorType"], ref)[0],
            "institutional_or_hnw": as_is(row["InstitutionalHNW"]),
            "region_location": as_is(row["Region"]),
            "investment_grade": yn_bool(row["InvestmentGrade"]),
            "ubs_lp_category": ubs_cls,               # per-row attributes via BB Criteria Matrix waterfall
            "sp_rating": as_is(row["SP"]),
            "moodys_rating": as_is(row["Moodys"]),
            "fitch_rating": as_is(row["Fitch"]),
            "aum": as_is(row["AUM"]),         # LP Size passthrough (UI formats)
            "nav": as_is(row["NAV"]),
            "pension_assets": as_is(row["PensionAssets"]),
            "funding_ratio": pct(row["FundingRatio"]),
            "pct_of_fund_commitments": pct(row["PercentOfCommitments"]),
            "called_capital": money_short(row["Called"]),
            "pct_of_fund_uncalled": pct(row["PercentOfUncalled"]),
            "pct_lp_called": pct(row["CalledPercent"]),
            "ubs_concentration_limit": pct(row["UBSCL"]),
            "ubs_advance_rate": floor_rate_pct(row["UBSAR"], ref),  # Floor-Map group
            "agent_borrowing_base": money_short(row["AgentBB"]),
            "ubs_borrowing_base": money_short(row["UBSBB"]),
            "notes": as_is(row["Notes"]),
        })
        counts["written"] += 1

    return SeedResult(seed_rows, counts, agent_unmatched, cls_counts, variance)


def upsert_facilities(fac_data: list[list[str]], by_acct: dict[str, int],
                      export: list[dict]) -> tuple[list[list[str]], Counter, dict[str, str]]:
    """Facilities from the Agent Bank Summary: bank_status := Active if the account appears in the
    export else Inactive (Active also gets collateral_date := BBDate) — this OVERRIDES the report's
    own FacilityStatus. Export accounts NOT listed in the report are MANUFACTURED as placeholder
    Inactive facilities ("Unknown" bank, name=FndName,
    account_number=AccountID, collateral_date=BBDate) so every LP record can seed — no rejects.
    Returns (rows, counts, account_number -> resolved facility name) for the seed to reuse."""
    bbdate_by_acct: "OrderedDict[str, str]" = OrderedDict()
    fnd_by_acct: "OrderedDict[str, str]" = OrderedDict()
    for row in export:
        acct = (row["AccountID"] or "").strip()
        if acct and acct not in bbdate_by_acct:
            bbdate_by_acct[acct] = iso_date(row["BBDate"])
            fnd_by_acct[acct] = as_is(row["FndName"])

    counts = Counter()
    out = [list(r) for r in fac_data]
    name_by_acct: dict[str, str] = {}
    used_norm = {_norm(r[1]) for r in out if len(r) > 1 and r[1].strip()}

    # 1) Reported facilities -> Active/Inactive. Status is resolved PER ROW off that row's own
    #    account, so where the report lists one account against two borrowers both rows follow
    #    that account; only the owning row (by_acct) lends its name to the LP seeds.
    for acct, idx in by_acct.items():
        name_by_acct[acct] = out[idx][1].strip()
    for row in out:
        acct = row[2].strip()
        if acct in bbdate_by_acct:
            row[5] = "Active"
            if bbdate_by_acct[acct]:
                row[8] = bbdate_by_acct[acct]
            counts["active"] += 1
        else:
            row[5] = "Inactive"
            counts["inactive"] += 1

    # 2) Orphan export accounts -> new placeholder Inactive facilities so their LPs still seed.
    #    A name already in use (existing facility, or another orphan sharing the FndName — a sim
    #    artifact) is disambiguated with the AccountID so each account survives as its own facility.
    for acct in bbdate_by_acct:
        if acct in by_acct:
            continue
        name = fnd_by_acct[acct] or f"Unknown Facility {acct}"
        if _norm(name) in used_norm:
            name = f"{name} ({acct})"
        used_norm.add(_norm(name))
        # agent_bank, name, account_number, loan_amount, maturity_date, bank_status,
        # bank_status_date, ubs_participation, collateral_date
        out.append(["Unknown", name, acct, "", "", "Inactive", "", "", bbdate_by_acct[acct]])
        name_by_acct[acct] = name
        counts["inactive_new"] += 1

    return out, counts, name_by_acct


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
    # No command-line arguments: the run is configured by the EXPORT_FILE variable near the top of
    # this file (edit it, then re-run). All outputs go to OUT_DIR; the app tree is never modified.
    export_path = Path(EXPORT_FILE)
    abs_path = Path(AGENT_BANK_SUMMARY_FILE)
    out_dir = Path(OUT_DIR)
    ref_dir = Path(REFERENCE_DIR)
    out_dir.mkdir(parents=True, exist_ok=True)

    print(f"Source: {export_path}")
    print(f"        {abs_path}")
    ref = load_references(ref_dir)
    # Small file first, so a missing/renamed report fails before the big export is parsed.
    fac_data, by_acct, abs_counts = read_agent_bank_summary(abs_path)
    export = read_export(export_path)

    # Facilities first: this manufactures placeholder facilities for orphan accounts and returns
    # the account -> facility name map the seed uses, so every LP record resolves to a facility.
    fac_rows, fac_counts, name_by_acct = upsert_facilities(fac_data, by_acct, export)
    mr = build_master(export, ref)
    sr = build_seed(export, name_by_acct, ref)

    generated = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Clean slate: remove EVERY file left in data/out/ by a previous run (including the review
    # reports and EXTRACT_SUMMARY.txt earlier versions of this script wrote) so the directory holds
    # only this run's three ingestion CSVs and pe-sub-jobs never sees a stale file.
    # Deliberately done AFTER all inputs are read, so a failed read leaves the last good outputs
    # in place. Only out/ is cleared; import/ (export + chaos log) and reference/ are never touched.
    for old in out_dir.iterdir():
        if old.is_file():
            old.unlink()

    # --- the three ingestion CSVs (read directly from data/out/ by pe-sub-jobs on startup) ---
    write_csv(out_dir / "lp_master.csv", MASTER_COLS, mr.rows)
    write_csv(out_dir / "lp_facility_seeds.csv", SEED_COLS, sr.rows)
    write_facilities(out_dir / "facilities.csv", fac_rows)

    # --- console-only run summary (nothing else is written to data/out/) ---
    total_facilities = fac_counts["active"] + fac_counts["inactive"] + fac_counts["inactive_new"]

    def fmt_cls(counter: Counter) -> str:
        return ", ".join(f"{cls} {n}" for cls, n in counter.most_common())

    ar_beyond = sum(a["ar_beyond"] for a in sr.variance.values())
    cl_beyond = sum(a["cl_beyond"] for a in sr.variance.values())
    summary_lines = [
        "LP DB Export -> seed extract — run summary",
        f"generated           : {generated}",
        f"source export       : {export_path}",
        f"agent bank summary  : {abs_path}",
        f"reference dir       : {ref_dir}",
        f"output directory    : {out_dir}",
        "",
        f"export rows read              : {len(export)}",
        f"agent bank summary read       : {abs_counts['facilities']} facilities "
        f"under {abs_counts['agents']} agent bank(s)",
        f"  dropped: subtotal rows      : {abs_counts['total_rows']} (AccessTotalsLoanAmount bands)",
        f"  dropped: reprinted rows     : {abs_counts['duplicate_row']} (same account + borrower)",
        f"  dropped: incomplete rows    : {abs_counts['skipped_incomplete']} (no borrower or no account)",
        f"  renamed: duplicate names    : {abs_counts['renamed_duplicate_name']} "
        f"(name reused by another account - suffixed with its AccountNumber)",
        f"  accounts listed twice       : {abs_counts['duplicate_account']} "
        f"(2nd borrower on an account; 1st row owns the LP join)",
        f"lp_master.csv (distinct LPs)  : {len(mr.rows)}",
        f"lp_facility_seeds.csv (rows)  : {sr.counts['written']}",
        f"  skipped: duplicate pair     : {sr.counts['skipped_duplicate_pair']} (same LP in same facility)",
        f"  skipped: blank investor     : {sr.counts['skipped_blank_investor']} (no LP key - cannot insert)",
        f"  skipped: no facility        : {sr.counts['skipped_no_facility']} (should be 0)",
        f"facilities.csv                : {total_facilities} total "
        f"({fac_counts['active']} Active, {fac_counts['inactive']} Inactive, "
        f"{fac_counts['inactive_new']} new 'Unknown' placeholder Inactive)",
        "",
        "normalization:",
        f"  investor types unmatched    : {len(mr.itype_unmatched)} distinct value(s)",
        f"  agent categories unmatched  : {len(sr.agent_unmatched)} distinct value(s)",
        f"  ubs classification (BB Criteria Matrix waterfall; rates Floor-Mapped to 90/75/65/50/0):",
        f"    lp_master ({len(mr.rows)} LPs)   : {fmt_cls(mr.cls_counts)}",
        f"    seed rows ({sr.counts['written']} rows) : {fmt_cls(sr.cls_counts)}",
        f"  matrix variance (Floor-Mapped UBSAR / raw UBSCL vs matrix, tol ±{MATRIX_TOLERANCE_PP}pp) : "
        f"{ar_beyond} AR row(s), {cl_beyond} CL row(s) beyond tolerance",
    ]
    # No review report files any more, so the values an analyst would have opened them for are
    # listed here instead (original value kept in the output either way; add a reference alias).
    if mr.itype_unmatched:
        summary_lines.append("")
        summary_lines.append("unmatched investor types (reference/investor_types.csv) — value, rows, suggestion:")
        summary_lines += [f"  {v} ({c}) -> {_suggest(v, ref.itype_canonical) or '?'}"
                          for v, c in mr.itype_unmatched.most_common()]
    if sr.agent_unmatched:
        summary_lines.append("")
        summary_lines.append("unmatched agent LP categories (reference/agent_lp_categories.csv) — value, rows, suggestion:")
        summary_lines += [f"  {v} ({c}) -> {_suggest(v, ref.agent_canonical) or '?'}"
                          for v, c in sr.agent_unmatched.most_common()]
    summary_lines += [
        "",
        "Every export account maps to a facility (existing or manufactured), so no LP record is",
        "rejected. data/out/ holds exactly the three ingestion CSVs above — pe-sub-jobs reads them",
        "directly on startup, so restart it (or trigger its /jobs endpoints) to load. data/mock is",
        "not involved.",
    ]

    # --- console echo (counts onward; the header/paths block is skipped) ---
    print("\n".join(summary_lines[7:]))
    print(f"\nAll outputs written to: {out_dir}")
    print("pe-sub-jobs reads these files directly from data/out/ on startup - restart it "
          "(or trigger its /jobs endpoints) to load.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
