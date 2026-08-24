#!/usr/bin/env python3
r"""
Generate a reproducible sample of the 2026-08-18, 30-column LP DB Export format for
lp_db_extract.py. The workbook contains facility/investor positions with canonical reference
values, reconciled borrowing-base fields, and optional recoverable data-quality variations.

The UBS credit columns are never among those variations. UBS LP Classification, UBS Advance Rate,
UBS Concentration Limit and UBS Borrowing Base are always genuine: the classification is written in
the canonical vocabulary of data/reference/ubs_lp_categories.csv, and the rate, the limit and the
base are the Borrowing Base Criteria Matrix's answer for that exact class, re-derived and asserted
per row before the workbook is written.

A fund financed under both sleeves of one Credit Agreement appears TWICE, as it does in the real
file: "<Fund> (Committed)" and "<Fund> (Uncommitted)" carry the same LPs and the same money under
two AccountIDs and price differently on the agent's side. See the tranche section below.

Output:
    data/import/LP DB Export YYYY.MM.DD.xlsx

Usage (no command-line arguments):
    1. Edit SEED, CHAOS_SEED, TARGET_ROWS, or CHAOS_ENABLED below as needed.
    2. Run from any directory:
             python pe-sub-jobs/scripts/lp_db_generate.py
    3. Set lp_db_extract.py's EXPORT_FILE to the printed workbook path and run:
             python pe-sub-jobs/scripts/lp_db_extract.py

The same SEED, CHAOS_SEED, and settings reproduce the same generated export.
"""
from __future__ import annotations

import csv
import math
import random
import re
from collections import Counter
from datetime import date, timedelta
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import openpyxl

SCRIPT_DIR = Path(__file__).resolve().parent          # pe-sub-jobs/scripts/
DATA_DIR = SCRIPT_DIR.parent / "data"                 # pe-sub-jobs/data/
AGENT_BANK_SUMMARY_FILE = DATA_DIR / "import" / "AgentBankSummaryRpt.xlsx"
REFERENCE_DIR = DATA_DIR / "reference"                # the same lists lp_db_extract normalizes against
EXPORT_OUT = DATA_DIR / "import" / f"LP DB Export {date.today():%Y.%m.%d}.xlsx"
SHEET_NAME = "BBs"

# ── tunables ────────────────────────────────────────────────────────────────
SEED = 20260625
CHAOS_ENABLED = True            # degrade the written XLSX to realistic manual-entry quality
CHAOS_SEED = 20260625           # chaos has its own rng: base data identical with chaos on/off
TARGET_ROWS = 20_000            # lp_records to produce (mirrored tranche rows counted)
REPEAT_MIN, REPEAT_MAX = 4, 12  # facilities each LP participates in

# A subscription line is drawn against uncalled capital, so a facility's LP pool is sized FROM its
# loan amount: the commitment is typically 15-30% of the uncalled capital backing it. Positions
# used to be drawn free-floating at $2-500M each regardless of which facility they landed on, which
# made a facility's total nothing but (LP count x $250M) - so "Uncalled to Facility" read 81,982.7%
# against a $75M loan. Apportioning from loan_amount puts every facility-relative ratio on the
# Shadow BB screen (Facility LTV, Uncalled to Facility, BB to Facility) in credit-officer range.
FACILITY_LTV_MIN, FACILITY_LTV_MAX = 0.15, 0.30
LP_SKEW_SIGMA = 0.9             # log-normal spread of position size within one facility
MONEY_STEP = 100_000            # the $100K grid money_short() carries losslessly ("$132.7M")
MIN_UNCALLED = 200_000          # keep the smallest position clear of that display granularity
# Row COUNT also scales with the loan amount (it used to be a flat uniform(0.4, 3.2)), so position
# size stays comparable across facilities instead of a small facility splitting its pool 400 ways.
FAC_WEIGHT_JITTER = (0.6, 1.6)
DEFAULT_LOAN_AMOUNT = 140_000_000   # orphan accounts: the summary report does not list them
ORPHAN_ACCOUNTS = [             # AccountIDs the report omits -> exercise the "Unknown" agent path
    ("5VZ9001", "TPG AG Asset Based Credit Fund"),
    ("5VZ9002", "TPG AG Asset Based Credit Fund"),
]

# The 30 headers of the 2026-08-18 LP DB Export, in order, EXACTLY as the real file spells them —
# quirks included, because reproducing them is most of the point of this generator: "LP Size" and
# "($ Bil)" are separated by a CRLF inside the one cell, "Insitutional" is misspelt at source, and
# "Moody'S" carries a capital S. lp_db_extract matches headers through _norm(), which absorbs all
# three, so this stays a faithful sample rather than a cleaned-up one.
SRC_HEADERS = [
    "AccountID", "FndName", "Investor Name", "Parent", "SPV", "UBS LP Classification",
    "Insitutional vs HNW", "Investment Grade?", "Agent LP Classification", "S&P", "Moody'S",
    "Fitch", "LP Size\r\n($ Bil)", "LP Size Criteria", "Capital Commitments", "Uncalled Capital",
    "UBS Advance Rate", "Agent Advance Rate", "Agent Concentration Limit",
    "UBS Concentration Limit",
    "% of Capital Commitments", "Called Capital", "% of Uncalled Capital", "% of LP Called",
    "Agent Excess Concentration", "UBS Excess Concentration", "Agent Borrowing Base",
    "UBS Borrowing Base", "Notes", "BBDate",
]

# Agent Bank Summary column layout, which the report header must match exactly. Index 4 is an
# unnamed spacer holding the report's subtotal amounts.
ABS_COLS = [
    "Agent", "Borrower", "AccountNumber", "LoanAmount", "", "MaturityDate",
    "FacilityStatus", "FacilityStatusDate",
]
ABS_TOTAL_MARKER = "accesstotalsloanamount"   # _norm() prefix of the subtotal / grand-total rows

# Internal keys for the same 30 columns, in the same order. Mirrors lp_db_extract.SRC_COLS so the
# chaos monkey can address a column by name and the two scripts stay legible side by side.
SRC_COLS = [
    "AccountID", "FndName", "InvestorName", "Parent", "SPV", "UbsClassification",
    "InstitutionalHNW", "InvestmentGrade", "Classification", "SP", "Moodys", "Fitch",
    "LpSizeBil", "LpSizeCriteria", "Commitments", "Uncalled", "UBSAR", "AgentAR",
    "AgentCL", "UBSCL",
    "PercentOfCommitments", "Called", "PercentOfUncalled", "CalledPercent",
    "AgentExcessConc", "UBSExcessConc", "AgentBB", "UBSBB", "Notes", "BBDate",
]

# ── reference vocabularies (canonical values, so they map cleanly through the extract) ───────
# investor_type -> (name-suffix templates, size measure). Canonical per data/reference/investor_types.csv.
TYPE_SPECS = {
    "Public Pension":        (["Public Employees Retirement System", "State Teachers Retirement", "State Pension Fund"], "pension_assets"),
    "Pension Fund":          (["Pension Fund", "Retirement Trust", "Pension Scheme"], "pension_assets"),
    "Endowment":             (["University Endowment", "Endowment Fund", "College Endowment"], "aum"),
    "Foundation":            (["Foundation", "Charitable Foundation", "Family Foundation"], "aum"),
    "Family Office":         (["Family Office", "Family Capital", "Family Holdings"], "aum"),
    "Insurance Company":     (["Life Insurance Co.", "Mutual Insurance", "Assurance Group"], "aum"),
    "Sovereign Wealth Fund": (["Investment Authority", "Sovereign Fund", "Future Fund"], "aum"),
    "Fund of Funds":         (["Fund of Funds", "Multi-Manager Fund", "Partners Fund"], "nav"),
    "Hedge Fund":            (["Master Fund", "Absolute Return Fund", "Opportunities Fund"], "nav"),
    "Endowment ":            (["Endowment"], "aum"),
    "Corporate":             (["Treasury", "Corporate Holdings", "Group Treasury"], "aum"),
    "Healthcare":            (["Health System", "Hospital Trust", "Healthcare Endowment"], "aum"),
    "Investment Consultant": (["Investment Advisors", "Capital Advisors", "Consulting Group"], "aum"),
    "Institutional Investor":(["Institutional Trust", "Alternative Assets Trust", "Capital Partners"], "aum"),
    "Other Institutional":   (["Strategic Capital Partners", "Alternative Assets", "Global Investors"], "aum"),
}
TYPE_WEIGHTS = {  # rough real-world mix
    "Public Pension": 16, "Pension Fund": 14, "Insurance Company": 12, "Endowment": 9,
    "Foundation": 8, "Sovereign Wealth Fund": 6, "Family Office": 8, "Fund of Funds": 7,
    "Hedge Fund": 4, "Corporate": 4, "Healthcare": 3, "Investment Consultant": 2,
    "Institutional Investor": 3, "Other Institutional": 4,
}

# Agent LP Category (export "Agent LP Classification") -> (advance rate, concentration limit,
# weight). The Agent Advance Rate Schedule is the authority for all three: its CLASSIFICATION
# column states the vocabulary and its Rate / Conc. limit columns state what each is worth.
#
# The rate and limit drawn here are written to the export's Agent Advance Rate / Agent Concentration
# Limit columns AND are what the Agent Borrowing Base below is computed with, so the file reconciles
# on its own terms. Both MUST stay in step with data/reference/agent_rate_map.csv (and therefore
# with the `agent_tiers` config it mirrors): that file is the extract's fallback for a row whose
# cell the chaos monkey - or a real analyst - left blank, and a fallback that resolved to a
# different number than the one the borrowing base was computed with would silently unbalance the
# row it repaired.
#
# Weights hold "Included Investors (Rated)" near the real book's ~40%, which is also the UBS "Rated Investor"
# share: ratings and that bucket are now the same population (see build_investor).
AGENT_CATEGORIES = [
    ("Included Investors (Rated)",       0.90, 0.200, 40),
    ("Included Investors (Non-Rated)",   0.75, 0.150, 28),
    ("Insitutional Designated Investors", 0.60, 0.125, 16),
    ("PWM Designated Investors",          0.50, 0.050,  9),
    ("Excluded Investors",         0.00, 0.000,  7),
]
AGENT_BY_NAME = {c: (c, rate, limit) for c, rate, limit, _ in AGENT_CATEGORIES}

# ── credit-agreement tranches: "<Fund> (Committed)" / "<Fund> (Uncommitted)" ─────────────────
# The real LP DB export carries the same fund twice - "Audax Private Equity Fund VII LP
# (Committed)" and "Audax Private Equity Fund VII LP (Uncommitted)" - with the SAME LP roster, the
# same LP COUNT and the same capital commitments, under two different AccountIDs. That is not
# duplication to be de-duplicated. The two are legally distinct tranches of one Credit Agreement:
# a committed line the lenders are contractually bound to fund up to a capped limit, and an
# uncommitted sleeve (an accordion feature or uncommitted overdraft) drawn only with lender consent
# per draw. Each carries its own facility ID, credit limit and draw availability; each has its own
# fee and pricing mechanics (an unused-commitment fee of 25-50bps accrues on the committed line's
# undrawn portion and on nothing else; an uncommitted draw may price at a different margin); and
# the agent bank requires a SEPARATE Borrowing Base Certificate per facility ID, so the two
# collateral pools are never blended and each tranche's headroom is tracked on its own.
#
# Per the Day 1 agreement with the business, the platform therefore treats each sleeve as its own
# record - which is why the generator emits each as its own facility rather than folding the pair
# into one, and why this section exists to keep it doing so.
#
# What separates the two rows for the same LP is the CREDIT TERMS. The LP that advances at 90% on
# the committed line is haircut a rung down the schedule on the uncommitted one and tested against
# half the concentration cap, so two borrowing bases genuinely diverge off identical collateral -
# which is the whole reason the split has to survive ingestion instead of being reconciled away.
TRANCHE_SUFFIX_RE = re.compile(r"\s*\((committed|uncommitted)\)\s*$", re.I)

# The agent's advance-rate rungs, best to worst: the AGENT_CATEGORIES rates plus a 25% rung below
# "PWM Designated Investors" for the step-down to land on. An uncommitted position prices one rung down.
AGENT_RATE_LADDER = (0.90, 0.75, 0.60, 0.50, 0.25, 0.00)
UNCOMMITTED_CL_FACTOR = 0.5     # and is tested against half the committed sleeve's concentration cap


# ── Borrowing Base Criteria Matrix ───────────────────────────────────────────────────────────
# One row per UBS LP Classification -> (concentration limit %, advance rate % below the funded
# threshold, advance rate % at or above it). Mirrors data/reference/bb_criteria_matrix.csv and the
# `bb_criteria_matrix` config in pe-sub-api (V1_3__config.sql) that BbCriteriaResolver reads.
#
# This is why the UBS Advance Rate and UBS Concentration Limit are NOT random draws. Both columns
# are stated by the export and pe-sub-api prefers the stated value over its own default
# (BbCalculationService.advanceRateFraction / perLpConc), so a random rate would not read as a
# manual override - it would read as the matrix disagreeing with itself, and every borrowing base
# in the file would sit at a rate the platform would never have chosen for that classification.
# "Rated Investor" carries one entry PER RATING BAND; every other class is band-independent.
FUNDED_THRESHOLD_PCT = 0.40
BB_MATRIX = {
    ("Rated Investor", "AAA"):            (0.25, 0.90, 0.90),
    ("Rated Investor", "AA"):             (0.20, 0.90, 0.90),
    ("Rated Investor", "A"):              (0.15, 0.90, 0.90),
    ("Rated Investor", "BBB"):            (0.10, 0.65, 0.90),
    ("Corp Pension > $5Bn Assets", None): (0.25, 0.90, 0.90),
    ("Corp Pension > $1Bn Assets", None): (0.20, 0.90, 0.90),
    ("Unrated NAV > $1Bn", None):         (0.15, 0.90, 0.90),
    ("FoF & Other > $10Bn AUM", None):    (0.10, 0.65, 0.75),
    ("Other Institutional", None):        (0.05, 0.50, 0.65),
    ("HNW Feeder (acceptable)", None):    (0.05, 0.50, 0.65),
    ("HNW (acceptable)", None):           (0.01, 0.00, 0.50),
    ("Excluded", None):                   (0.00, 0.00, 0.00),
}
# The nine classes classify_ubs may return, which are also UBS_CLS_OPTS and the canonical column of
# reference/ubs_lp_categories.csv. Nothing else may reach the "UBS LP Classification" column.
UBS_CLASSES = tuple(dict.fromkeys(cls for cls, _ in BB_MATRIX))

# Best-to-worst, the ordinals the tri-party waterfall sorts on. A present notch that matches no band
# is sub-investment-grade and clamps to BBB (Q2: a rated LP is never reclassified out of Rated).
BAND_ORDER = ("AAA", "AA", "A", "BBB")
SUB_INVESTMENT_GRADE_BAND = "BBB"
AGENCIES = ("sp", "moodys", "fitch")            # the rating_scales.csv agency keys, in column order

# Band mix for a rated LP, weighted toward the middle of the curve. SUB_IG is the small
# below-BBB- tail: still Rated Investor, resolved onto the matrix's BBB row.
RATED_BAND_WEIGHTS = {"AAA": 6, "AA": 24, "A": 40, "BBB": 25, "SUB_IG": 5}
# How many of the three agencies cover one rated LP - "S&P, Moody's and/or Fitch". Most LPs the bank
# rates carry all three; a single-agency LP is what makes the extract's one-rating waterfall branch
# (and pe-sub-api's) reachable at all.
AGENCY_COVERAGE_WEIGHTS = [18, 27, 55]          # P(1 agency), P(2), P(3)

WORD1 = ["Apex", "Meridian", "Granite", "Harborview", "Ironwood", "Cascade", "Dominion", "Everest",
         "Fairview", "Northgate", "Oakmont", "Pinnacle", "Redwood", "Silverpeak", "Trident", "Union",
         "Westbridge", "Yorkshire", "Zenith", "Brookstone", "Cedar", "Kestrel", "Larchmont", "Monarch",
         "Sterling", "Beacon", "Clearwater", "Highgate", "Lakeshore", "Ridgeline", "Summit", "Vantage",
         "Ashford", "Bluewater", "Copperfield", "Draycott", "Eastgate", "Fenwick", "Glenwood", "Halcyon"]
WORD2 = ["", "", "", "Capital", "Global", "Strategic", "Alternative", "Atlantic", "Pacific",
         "Continental", "Heritage", "Legacy", "Pioneer", "Cornerstone", "Evergreen"]


# ── chaos monkey (pe-sub-docs/"AI Chaos Monkey for Data Quality.md") ────────────────────────
# Probability per row that each field family is degraded.
#
# The UBS side of the export - UBS LP Classification and the three figures it determines (UBS
# Advance Rate, UBS Concentration Limit, UBS Borrowing Base, and the excess concentration that
# nets them) - is NOT in this table and cannot be put in it: see CHAOS_SACRED. Those columns are
# the bank's own credit decision on the row, and the sample states them genuinely so that the
# rate, the limit and the base always read as the Borrowing Base Criteria Matrix's answer for the
# classification stated three columns to their left.
CHAOS_RATES = {
    "investor_name": 0.25,     # suffix drops / tranche suffixes / case flips - once per LP, not per row
    "ratings": 0.15,           # 'A-' -> 'A minus', 'Baa 1', case noise (all three agencies)
    "lp_size": 0.15,           # LP Size -> range/threshold/unit strings ('5 - 8', '>10', '2 bn')
    "lp_size_criteria": 0.08,  # size basis left blank (the label itself never drifts)
    "agentar_null": 0.04,      # Agent Advance Rate left blank      -> agent_rate_map.csv rate_pct
    "agentcl_null": 0.04,      # Agent Concentration Limit blank    -> agent_rate_map.csv conc_limit_pct
}

# Never degraded, and enforced in mutate(): the cash and legal-LPA figures, the facility join keys,
# and every column the borrowing bases are reconciled from. A dirty spelling of a name is still that
# name; a dirty commitment is a different commitment.
#
# The whole UBS column family is sacred alongside the money, and for the same reason: these are
# not fields an analyst transcribes, they are the bank's stated credit decision on the position and
# the arithmetic that follows from it. They move together or not at all -
#   UbsClassification  decides the matrix row, so drifting it (even to an alias that resolves back)
#                      is the one degradation that can leave a row saying one class and pricing at
#                      another the moment anything downstream reads the label before the alias list;
#   UBSAR / UBSCL      are that row's advance rate and concentration limit, which pe-sub-api prefers
#                      over its own matrix default (BbCalculationService.advanceRateFraction /
#                      perLpConc) - a blank is not recovered as the seeded value on the LP Master
#                      row (ubs_default_advance_rate), so an emptied cell re-prices the LP;
#   UBSBB / UBSExcessConc are computed FROM the other three against the facility's total uncalled.
# Degrade any one of them and the row no longer reconciles against itself, which is exactly what
# this sample exists to prove it does.
CHAOS_SACRED = ("AccountID", "FndName", "Commitments", "Called", "Uncalled", "BBDate",
                "UbsClassification", "UBSAR", "UBSCL", "UBSBB", "UBSExcessConc",
                "AgentBB", "AgentExcessConc",
                "PercentOfCommitments", "PercentOfUncalled", "CalledPercent")

# Blank-only, and also enforced: these cells may be EMPTIED but never rewritten, because each has a
# documented fallback that resolves to exactly the value that was there - so the row still
# reconciles against a cell an analyst simply did not fill in.
#   AgentAR/AgentCL -> the extract falls back to agent_rate_map.csv keyed by the row's Agent LP
#                      Category, which is where the written values came from (AGENT_CATEGORIES) -
#                      EXCEPT on an uncommitted tranche, whose rate and limit are that sleeve's own
#                      haircut price (agent_terms) and not the schedule's. The map would resolve a
#                      blank there back to the COMMITTED number, so apply_chaos leaves those two
#                      cells alone on an uncommitted row;
#   LpSizeCriteria  -> the figure has no basis without it, so the extract blanks the LP's size on
#                      that row - but build_master consolidates the (size, criteria) pair from the
#                      most recent submission that RESOLVES, so an LP keeps its size as long as one
#                      of its rows still states the basis. apply_chaos guarantees one always does.
CHAOS_BLANKABLE = ("AgentAR", "AgentCL", "LpSizeCriteria")

_NAME_SUFFIX_RE = re.compile(r",?\s+(LLC|L\.L\.C\.|L\.P\.|LP|Ltd\.?|Inc\.?|Limited)$", re.I)

# UBS LP Classification does NOT drift, and neither do the three figures it determines. The export
# states the classification outright and it is the key the Borrowing Base Criteria Matrix is read
# with, so an alias in that cell - even one ubs_lp_categories.csv resolves - is a row that says one
# thing and prices as another to any reader that has not applied the alias list yet. The generator
# writes the canonical spelling ubs_lp_categories.csv states, and only that, which is what lets the
# UBS Advance Rate, UBS Concentration Limit and UBS Borrowing Base in the same row be re-derived
# from the label beside them. validate_chaos_vocabularies() checks that alignment against the
# reference list before a run writes anything.
#
# Agent LP Classification does NOT drift either. Nothing downstream rewrites it, so a drifted label would
# reach lp_records as the classification itself and price at a rate resolved from a class it no
# longer names. The generator emits only what the Agent Advance Rate Schedule states.


def _blank(v) -> bool:
    return v is None or str(v).strip() == ""


def _as_str(v) -> str:
    return "" if v is None else str(v)


def _norm(s: str) -> str:
    """Case/punctuation-insensitive key. Byte-identical to lp_db_extract._norm on purpose: the two
    scripts have to agree on what counts as "the same label", or a spelling this file believes is a
    known alias reaches the extract as an unknown value."""
    return re.sub(r"[^a-z0-9]+", " ", str(s or "").lower()).strip()


def _reference_rows(name: str) -> list[list[str]]:
    """Rows of one data/reference/ CSV (header included), blank and '#' comment lines dropped."""
    path = REFERENCE_DIR / name
    if not path.is_file():
        raise SystemExit(f"Reference file not found: {path}\n"
                         "The pe-sub-jobs/data/reference/ lists are required by both scripts.")
    with path.open(newline="", encoding="utf-8") as fh:
        return [[c.strip() for c in cells] for cells in csv.reader(fh)
                if cells and cells[0].strip() and not cells[0].lstrip().startswith("#")]


def _alias_lookup(name: str) -> dict[str, str]:
    """norm(alias) -> canonical, from one of the two classification reference lists."""
    return {_norm(r[0]): r[1] for r in _reference_rows(name)[1:] if len(r) >= 2 and r[1]}


def load_rating_scales() -> tuple[dict[str, dict[str, list[str]]], dict[str, list[str]]]:
    """reference/rating_scales.csv -> ({agency: {band: [notch, ...]}}, {agency: [sub-IG notch, ...]}).

    Each agency keeps its OWN ladder. S&P and Fitch write AA-/BBB+ where Moody's writes Aa3/Baa1, and
    pe-sub-api's BbCriteriaResolver looks a notch up under the column it arrived in - so an S&P notch
    written into the Moody's column matches no band, clamps that agency to the sub-IG floor, and
    drags the tri-party median down a band. Reading the scales from the file the extract normalizes
    against is what keeps the sample out of that trap."""
    by_band: dict[str, dict[str, list[str]]] = {}
    sub_ig: dict[str, list[str]] = {}
    for row in _reference_rows("rating_scales.csv")[1:]:
        if len(row) < 3:
            continue
        agency, band, notch = row[0], row[1], row[2]
        if band:
            by_band.setdefault(agency, {}).setdefault(band, []).append(notch)
        else:
            sub_ig.setdefault(agency, []).append(notch)
    return by_band, sub_ig


RATING_BANDS, SUB_IG_NOTCHES = load_rating_scales()

# norm(alias) -> canonical UBS LP Classification. The generator only ever writes canonical values,
# so this is used to PROVE that rather than to translate: a class resolves to itself or the run
# stops. It is the same list lp_db_extract normalizes the column against, which is what keeps the
# advance rate, concentration limit and borrowing base in a row aligned to the class beside them.
UBS_CANONICAL = _alias_lookup("ubs_lp_categories.csv")
# The sub-investment-grade notches a fund LP realistically carries. The scales run down to D, but an
# LP the bank has admitted at all is not a defaulted issuer; drawing from the whole tail would put
# CCC- and D ratings on live positions.
SUB_IG_DRAW_DEPTH = 6


def _band_ordinal(agency: str, notch: str) -> int:
    """Best-to-worst ordinal of one agency notch. A present notch in no band is sub-investment-grade
    and clamps to the BBB floor - matching BbCriteriaResolver.addOrdinal, and Q2's rule that a rated
    LP is never reclassified out of Rated Investor by a weak rating."""
    for i, band in enumerate(BAND_ORDER):
        if notch in RATING_BANDS.get(agency, {}).get(band, ()):
            return i
    return BAND_ORDER.index(SUB_INVESTMENT_GRADE_BAND)


def rating_band(sp: str, moodys: str, fitch: str) -> str | None:
    """The bb_criteria_matrix rating band for one LP's agency ratings, or None when it carries none.

    The tri-party "eligible rating" waterfall pe-sub-api runs (BbCriteriaResolver.resolveBand):
    three ratings -> the middle (median), two -> the LOWER of the two (conservative), one -> that
    one. Resolving the band HERE from the notches actually written - rather than writing notches to
    match a band picked in advance - is what lets the three agencies legitimately disagree while the
    concentration limit stays the one the platform will resolve for the same row."""
    ords = [_band_ordinal(agency, str(value).strip())
            for agency, value in zip(AGENCIES, (sp, moodys, fitch)) if not _blank(value)]
    if not ords:
        return None
    ords.sort()                                          # ascending = best -> worst
    chosen = ords[0] if len(ords) == 1 else ords[1] if len(ords) == 2 else ords[(len(ords) - 1) // 2]
    return BAND_ORDER[min(chosen, len(BAND_ORDER) - 1)]


def agency_ratings() -> tuple[str, str, str]:
    """(S&P, Moody's, Fitch) for a rated LP: one to three agencies cover it - "S&P, Moody's and/or
    Fitch" - and an agency that does not is left BLANK. "NR" / "N/R" / "N/A" are never written: the
    real LP DB Export leaves the cell empty, and a token in a rated LP's rating column is a fact
    about that LP, not a spelling of one."""
    target = pick_weighted(list(RATED_BAND_WEIGHTS), list(RATED_BAND_WEIGHTS.values()))
    covered = random.sample(AGENCIES, k=random.choices([1, 2, 3], weights=AGENCY_COVERAGE_WEIGHTS)[0])
    out = dict.fromkeys(AGENCIES, "")
    for agency in covered:
        out[agency] = _draw_notch(agency, target)
    return out["sp"], out["moodys"], out["fitch"]


def _draw_notch(agency: str, target: str) -> str:
    """One notch off `agency`'s own ladder, at the target band with a one-notch jitter, so the
    covering agencies land close together without being forced to agree."""
    ladder = [n for band in BAND_ORDER for n in RATING_BANDS[agency][band]]
    ladder += SUB_IG_NOTCHES[agency][:SUB_IG_DRAW_DEPTH]
    pool = (SUB_IG_NOTCHES[agency][:SUB_IG_DRAW_DEPTH] if target == "SUB_IG"
            else RATING_BANDS[agency][target])
    i = ladder.index(random.choice(pool))
    return ladder[min(len(ladder) - 1, max(0, i + random.randint(-1, 1)))]


def is_investment_grade(sp: str, moodys: str, fitch: str) -> bool:
    """The export's "Investment Grade?" column: every agency that covers the LP has it in a band.
    Derived from the notches rather than drawn, so the flag cannot contradict the ratings beside
    it."""
    covering = [(a, str(v).strip()) for a, v in zip(AGENCIES, (sp, moodys, fitch)) if not _blank(v)]
    return bool(covering) and all(v not in SUB_IG_NOTCHES[a] for a, v in covering)


def bb_criteria(ubs_cls: str, band: str | None, pct_funded: float) -> tuple[float, float]:
    """(UBS advance rate, UBS concentration limit) as fractions, from the Borrowing Base Criteria
    Matrix - the same resolution BbCriteriaResolver performs for the same row.

    The advance rate steps at the funded threshold, so it is a property of the POSITION, not of the
    LP: one LP can advance at 65% on a facility that is 30% called and at 90% on one that is 55%
    called. The concentration limit is funded-independent."""
    key = (ubs_cls, (band or SUB_INVESTMENT_GRADE_BAND) if ubs_cls == "Rated Investor" else None)
    conc_limit, rate_lt40, rate_gte40 = BB_MATRIX[key]
    return (rate_gte40 if pct_funded >= FUNDED_THRESHOLD_PCT else rate_lt40), conc_limit


def validate_chaos_vocabularies() -> None:
    """Assert the sample can only state things the platform understands, BEFORE anything is written.

    Three properties, each of which has a cheap way to go wrong and an expensive way to be found out
    (a reseeded database whose classifications no longer match its rates):
      * every class the generator can emit is one of the nine, and the matrix prices it;
      * the nine agree, name for name, with the canonical column of ubs_lp_categories.csv - so the
        classification a row states is the same string the matrix resolved its UBS advance rate and
        concentration limit from, and the UBS borrowing base built on those two describes the class
        printed beside it;
      * every Agent LP Category the generator emits is one the Agent Advance Rate Schedule states,
        carrying the rate and limit it would fall back to, so blanking either cell cannot change
        the row."""
    problems: list[str] = []

    for cls in UBS_CLASSES:
        if cls != "Rated Investor" and (cls, None) not in BB_MATRIX:
            problems.append(f"UBS class {cls!r} has no BB_MATRIX row")

    # ubs_lp_categories.csv is the vocabulary the UBS side of the export is written in, so the
    # matrix that prices a row and the list that names its classification have to describe the same
    # nine classes. Checked in both directions: a class the generator can emit that the list does
    # not state canonically would reach the export as an alias (or as nothing the platform knows),
    # and a class the list states that the matrix does not price would be a row bb_criteria cannot
    # resolve a rate or a limit for.
    for cls in UBS_CLASSES:
        if UBS_CANONICAL.get(_norm(cls)) != cls:
            problems.append(f"UBS class {cls!r} is not the canonical spelling stated by "
                            f"ubs_lp_categories.csv (resolves to {UBS_CANONICAL.get(_norm(cls))!r}) "
                            "- the export would state a label its own advance rate and "
                            "concentration limit were not resolved from")
    for cls in sorted({c for c in UBS_CANONICAL.values() if c} - set(UBS_CLASSES)):
        problems.append(f"ubs_lp_categories.csv states {cls!r}, which BB_MATRIX does not price")

    agent_map = {r[0]: r[1:3] for r in _reference_rows("agent_rate_map.csv")[1:] if len(r) >= 3}
    for cat, rate, limit, _ in AGENT_CATEGORIES:
        fallback = agent_map.get(cat)
        if fallback is None:
            problems.append(f"Agent LP Category {cat!r} is not stated by agent_rate_map.csv")
        elif (float(fallback[0]) / 100, float(fallback[1]) / 100) != (rate, limit):
            problems.append(f"agent_rate_map.csv {cat!r} = {fallback}, generator uses "
                            f"{rate * 100:g}/{limit * 100:g} - a blanked cell would not be recovered "
                            f"as the value the borrowing base was computed with")

    for agency in AGENCIES:
        missing = [b for b in BAND_ORDER if not RATING_BANDS.get(agency, {}).get(b)]
        if missing:
            problems.append(f"rating_scales.csv has no {missing} notches for agency {agency!r}")

    if problems:
        raise SystemExit("the sample would not follow its own rules:\n  " + "\n  ".join(problems))


def _chaos_name(name: str, rng: random.Random) -> tuple[str, str]:
    """Entity-name drift: dropped legal suffix, ' - Tranche A', case flip."""
    options = [("tranche_suffix", name + " - Tranche A"),
               ("case_flip", name.upper() if rng.random() < 0.5 else name.lower())]
    stripped = _NAME_SUFFIX_RE.sub("", name).strip()
    if stripped != name:
        options.append(("suffix_dropped", stripped))
    return rng.choice(options)


def _chaos_rating(val: str, rng: random.Random) -> tuple[str, str]:
    """Rating FORMAT drift: 'A-' -> 'A minus', 'Baa1' -> 'Baa 1', case noise. The notch itself never
    moves and "NR" is never introduced - a rating is the evidence for the LP's Included Investors (Rated) /
    Rated Investor classification and for its advance rate, so changing or removing one changes what
    the row asserts. lp_db_extract.normalize_rating() maps every form here back to the canonical
    notch, so the seed carries a value pe-sub-api can still band."""
    s = str(val).strip()
    if "-" in s:
        return "sign_spelled", s.replace("-", " minus")
    if "+" in s:
        return "sign_spelled", s.replace("+", " plus")
    if s[-1:].isdigit():                                 # Moody's: 'Baa1' -> 'Baa 1'
        return rng.choice([("space_inserted", s[:-1] + " " + s[-1]), ("case_noise", s.lower())])
    return "case_noise", s.lower()


# The LP Size thresholds classify_ubs decides a classification on: Corp Pension $1Bn / $5Bn, Unrated
# NAV $1Bn, FoF & Other $10Bn AUM. A degraded size string may look nothing like the clean figure, but
# it has to land on the same side of every one of these - otherwise the row's own LP Size contradicts
# the classification stated three columns to its left, and an analyst reading the file cannot tell
# which of the two is the typo.
_SIZE_THRESHOLDS = (1.0, 5.0, 10.0)


def _same_size_side(candidate: float, clean: float) -> bool:
    return all((candidate > t) == (clean > t) for t in _SIZE_THRESHOLDS)


def _chaos_lp_size(val, rng: random.Random) -> tuple[str, str] | None:
    """LP Size manual-entry patterns. The column's unit is $Bn and its clean value is a bare number
    (13.5), so the realistic drift is what argues with that convention: a range, a qualitative
    threshold, or a spelled-out unit. lp_db_extract.parse_size_bil resolves each (a range takes its
    LOW end, an explicit unit wins), and the value it will resolve to is checked against
    _SIZE_THRESHOLDS here before the string is used - a candidate that would cross one is skipped
    for the next pattern.

    NOT in the repertoire, though both are real analyst mistakes: a $Bn figure retyped in millions,
    and one typed in absolute dollars. Nothing downstream can tell either from a genuinely different
    number, so they do not degrade the row's LP Size - they replace it."""
    try:
        bil = float(val)
    except (TypeError, ValueError):
        return None
    if bil <= 0:
        return None
    for pattern in rng.sample(["range", "threshold", "unit_spelled"], k=3):
        if pattern == "range":
            low = round(bil * rng.uniform(0.80, 0.95), 1)
            if low > 0 and _same_size_side(low, bil):
                return pattern, f"{low} - {round(bil * 1.25, 1)}"
        elif pattern == "threshold":
            floor_at = round(bil * 0.9, 1)
            if floor_at > 0 and _same_size_side(floor_at, bil):
                return pattern, f">{floor_at}"
        else:
            return pattern, rng.choice([f"${bil}B", f"{bil} bn", f"${bil}Bn"])
    return None


def apply_chaos(export_rows: list[dict], rng: random.Random) -> list[tuple]:
    """Degrade the export rows in place (CHAOS_RATES per field family). Returns one
    (xlsx_row_no, column, pattern, original, corrupted) record per mutation.

    CHAOS_SACRED and CHAOS_BLANKABLE are enforced here rather than left as documentation: a mutator
    that reached a sacred column - the money, the join keys, or any of the four UBS credit columns -
    or blanked a cell with no fallback behind it, is a bug in the
    chaos monkey and stops the run instead of shipping 20,000 rows nobody can reconcile."""
    muts: list[tuple] = []

    def mutate(row_no: int, row: dict, col: str, result: tuple[str, str | None] | None) -> None:
        if result is None:
            return
        pattern, new = result
        if col in CHAOS_SACRED:
            raise SystemExit(f"chaos monkey tried to rewrite the sacred column {col!r} "
                             f"({pattern}) on row {row_no}")
        if _blank(new) and col not in CHAOS_BLANKABLE:
            raise SystemExit(f"chaos monkey tried to blank {col!r} on row {row_no}, which has no "
                             "fallback to recover it")
        if new == row[col]:
            return
        muts.append((row_no, col, pattern, _as_str(row[col]), _as_str(new)))
        row[col] = new

    # An LP's name is its identity: build_master groups LP Master by investor_name, so a name that
    # drifted row to row would not read as one analyst's spelling of one LP - it would split that LP
    # into two or three golden profiles and leave each seed row linked to whichever one it happened
    # to spell. The spelling is therefore decided ONCE per LP and applied to every row it holds,
    # which keeps the drift (and the extract's tolerance for it) without costing the identity.
    name_variant: dict[str, tuple[str, str]] = {}
    decided: set[str] = set()
    rows_per_lp: Counter = Counter()
    # A sponsor's name is the anchor its feeders' Parent cells point at, and the variant is applied to
    # EVERY row the LP holds - so drifting a sponsor's spelling does not blur one submission, it
    # deletes the canonical name from the export and leaves each child a dangling pointer that
    # pe-sub-api links to nothing. Sponsors keep their spelling; feeders still drift.
    sponsors = {_as_str(row["Parent"]) for row in export_rows if not _blank(row["Parent"])}
    for row in export_rows:
        original = _as_str(row["InvestorName"])
        if not original:
            continue
        rows_per_lp[original] += 1
        if original not in decided:
            decided.add(original)
            if original not in sponsors and rng.random() < CHAOS_RATES["investor_name"]:
                name_variant[original] = _chaos_name(original, rng)

    # An LP whose every row lost its LP Size Criteria would lose its LP Size outright, because
    # build_master can only fall through to another row that still states a basis. Cap the blanking
    # at one short of the LP's row count so at least one submission always resolves.
    criteria_blanked: Counter = Counter()

    for row_no, row in enumerate(export_rows, start=2):   # 2-based: XLSX row incl. header
        investor = _as_str(row["InvestorName"])
        # An uncommitted tranche states a price the Agent Advance Rate Schedule does not: blanking
        # it would not be a recoverable gap, it would silently re-price the sleeve at the committed
        # line's rate. The rng is still drawn for the row, so the degradation of every other row is
        # bit-identical whether the sample happens to contain a tranche or not.
        agent_cells_are_overrides = tranche_of(row["FndName"]) == "Uncommitted"
        if investor in name_variant:
            mutate(row_no, row, "InvestorName", name_variant[investor])
        if rng.random() < CHAOS_RATES["ratings"]:
            for col in ("SP", "Moodys", "Fitch"):
                if not _blank(row[col]):
                    mutate(row_no, row, col, _chaos_rating(row[col], rng))
        if not _blank(row["LpSizeBil"]) and rng.random() < CHAOS_RATES["lp_size"]:
            mutate(row_no, row, "LpSizeBil", _chaos_lp_size(row["LpSizeBil"], rng))
        if (not _blank(row["LpSizeCriteria"])
                and criteria_blanked[investor] < rows_per_lp[investor] - 1
                and rng.random() < CHAOS_RATES["lp_size_criteria"]):
            criteria_blanked[investor] += 1
            mutate(row_no, row, "LpSizeCriteria", ("nulled", None))
        # The two fallback-backed cells. Each blank is repaired downstream with the number that
        # was in it, so the row still reconciles - which is what makes an empty cell a legitimate
        # randomization here and a rewritten one not. Their UBS counterparts are NOT here: the UBS
        # rate and limit have no fallback that restores the seeded value, and the classification
        # they were resolved from is stated as written.
        if (not _blank(row["AgentAR"]) and rng.random() < CHAOS_RATES["agentar_null"]
                and not agent_cells_are_overrides):
            mutate(row_no, row, "AgentAR", ("nulled", None))
        if (not _blank(row["AgentCL"]) and rng.random() < CHAOS_RATES["agentcl_null"]
                and not agent_cells_are_overrides):
            mutate(row_no, row, "AgentCL", ("nulled", None))
    return muts


def facility_key(acct: str, fund: str) -> tuple[str, str]:
    """Identity of one facility: its (account number, fund name) pair, which is what the platform
    keys on and what every per-facility figure in this script is grouped by.

    Neither half identifies a facility on its own. One account number can carry several funds - the
    summary report lists exactly that - and one fund name can be reported under several accounts.
    Grouping by the account alone would pool two facilities' LPs into one borrowing base; grouping
    by the name alone would do the same to two funds that share one."""
    return acct, fund


def tranche_of(fund_name: str) -> str | None:
    """'Committed' / 'Uncommitted' when the fund name names a tranche of a Credit Agreement, else
    None. The real file spells the suffix in both cases ("(Uncommitted)", "(UNCOMMITTED)")."""
    m = TRANCHE_SUFFIX_RE.search(_as_str(fund_name))
    return m.group(1).capitalize() if m else None


def tranche_base_name(fund_name: str) -> str:
    """The fund name with its tranche suffix stripped, casefolded - the key the sleeves of one
    Credit Agreement share and nothing else does."""
    return TRANCHE_SUFFIX_RE.sub("", _as_str(fund_name)).strip().casefold()


def tranche_groups(facilities: list[tuple[str, str, int]]) -> tuple[dict, list]:
    """({lead facility -> its mirror facilities}, [lead facilities]), each a facility_key pair.

    Facilities whose names share a base are the sleeves of one Credit Agreement. The committed
    sleeve leads (falling back to the first listed, so a group the report spells without an
    explicit "(Committed)" member still has exactly one lead) and every other sleeve mirrors it.

    Only a NAMED tranche is ever grouped: two facilities that merely share a fund name - the
    ORPHAN_ACCOUNTS pair, or one fund financed twice under separate accounts - are separate credit
    agreements, not two sleeves of one, and keep their own independent LP rosters. A facility with
    no sibling is its own lead, so the caller can treat every facility the same way."""
    by_base: dict = {}
    for acct, fund, _ in facilities:
        fk = facility_key(acct, fund)
        by_base.setdefault(tranche_base_name(fund) if tranche_of(fund) else fk, []).append(fk)
    siblings: dict = {}
    leads: list = []
    for members in by_base.values():
        lead = next((fk for fk in members if tranche_of(fk[1]) == "Committed"), members[0])
        leads.append(lead)
        siblings[lead] = [fk for fk in members if fk != lead]
    return siblings, leads


def agent_terms(agent_cat: str, fund_name: str) -> tuple[float, float]:
    """(agent advance rate, agent concentration limit) for one position, as the facility it sits on
    prices it: the Agent Advance Rate Schedule's values on a committed line or an untranched
    facility, stepped one rung down AGENT_RATE_LADDER and halved on an uncommitted one.

    Only the AGENT side moves with the sleeve. UBS's Borrowing Base Criteria Matrix is bank policy
    keyed on the LP's classification, and the export's UBSAR/UBSCL are what the extract
    consolidates BY RECENCY into LP Master's bank-wide ubs_default_advance_rate /
    ubs_default_concentration_limit - so a tranche haircut written into those columns would leave a
    facility-specific rate standing as the LP's bank-wide default for every OTHER facility, the
    moment an uncommitted row happened to be that LP's most recent submission. The agent columns
    carry no such consolidation: they are per-BB-run figures, which is where a per-tranche price
    belongs."""
    _, rate, limit = AGENT_BY_NAME[agent_cat]
    if tranche_of(fund_name) != "Uncommitted":
        return rate, limit
    rung = AGENT_RATE_LADDER.index(rate)
    return AGENT_RATE_LADDER[min(rung + 1, len(AGENT_RATE_LADDER) - 1)], limit * UNCOMMITTED_CL_FACTOR


def load_facilities() -> list[tuple[str, str, int]]:
    """(account_number, borrower name, loan amount) per facility the Agent Bank Summary report
    states, plus ORPHAN_ACCOUNTS.

    The report is the same source the extract takes its facilities from, read the same way, and
    that is the point: a sample built off a facility list that had drifted from the report would
    seed positions onto accounts the ingestion cannot place. It is a banded print layout, not a
    flat table:
      * the agent bank appears once on its own group-header row (Agent set, Borrower blank) and is
        carried down onto the facility rows beneath it, which leave the Agent cell blank - the
        sample does not use the agent, so those rows are simply skipped here;
      * each group ends with an "AccessTotalsLoanAmount:" subtotal row carrying no facility.

    A repeated (account, borrower) pair is a reprint of a row already taken and is dropped. Two
    borrowers listed against ONE account are two facilities and both are kept: an account number
    does not identify a facility on its own and neither does a fund name, so the pair of them is
    what every per-facility figure below is grouped by. The name is carried exactly as PRINTED,
    because that is what the export's FndName has to join back to.

    The loan amount is what every position on the facility is apportioned from, so a row that
    states none falls back to DEFAULT_LOAN_AMOUNT rather than to zero.
    """
    if not AGENT_BANK_SUMMARY_FILE.is_file():
        raise SystemExit(
            f"Agent Bank Summary report not found: {AGENT_BANK_SUMMARY_FILE}\n"
            "It states every facility the sample places positions on. Put it in the import "
            "directory, or point AGENT_BANK_SUMMARY_FILE at it, then re-run.")
    wb = openpyxl.load_workbook(AGENT_BANK_SUMMARY_FILE, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows_iter = ws.iter_rows(values_only=True)
    header = [_as_str(c) for c in next(rows_iter)][:len(ABS_COLS)]
    if header != ABS_COLS:
        raise SystemExit(f"Agent Bank Summary header in {ws.title!r} does not match the expected "
                         f"schema.\n  expected: {ABS_COLS}\n  found:    {header}")

    out: list[tuple[str, str, int]] = []
    seen: set[tuple[str, str]] = set()
    for raw in rows_iter:
        cells = [_as_str(c) for c in (list(raw) + [None] * len(ABS_COLS))[:len(ABS_COLS)]]
        if not any(cells):
            continue
        if _norm(cells[3]).startswith(ABS_TOTAL_MARKER):    # subtotal / grand-total band
            continue
        if cells[0] and not cells[1]:                       # agent group header
            continue
        name, acct = cells[1], cells[2]
        if not name or not acct or (acct, _norm(name)) in seen:
            continue
        seen.add((acct, _norm(name)))
        try:
            loan = int(float(cells[3]))
        except ValueError:
            loan = 0
        out.append((acct, name, loan or DEFAULT_LOAN_AMOUNT))
    wb.close()
    out.extend((acct, fund, DEFAULT_LOAN_AMOUNT) for acct, fund in ORPHAN_ACCOUNTS)
    return out


def snap_money(x: float, floor: int) -> int:
    """A dollar figure on the export's $100K grid, never below `floor`. Positions are apportioned as
    real-valued shares of a pool, and the export carries one decimal of millions, so every position
    is put back on that grid before anything is derived from it."""
    return max(floor, int(round(x / MONEY_STEP)) * MONEY_STEP)


def size_bil(measure: str) -> float:
    """One LP-size figure in BILLIONS of dollars — the unit the export's "LP Size ($ Bil)" column
    carries, so no conversion happens on the way out.

    Log-uniform: fund sizes are log-distributed, and it puts mass on both sides of the classification
    boundaries (Corp Pension $1Bn/$5Bn, Unrated NAV $1Bn, FoF & Other $10Bn AUM) so every branch of
    classify_ubs below sees traffic."""
    def log_uniform(lo: float, hi: float) -> float:
        return 10 ** random.uniform(math.log10(lo), math.log10(hi))
    if measure == "pension_assets":
        return round(log_uniform(0.8, 120), 1)
    if measure == "nav":
        return round(log_uniform(0.2, 8), 1)
    return round(log_uniform(0.3, 60), 1)            # aum


# Which LP Size Criteria label goes with each internal size measure. Matches the platform's
# LP_SIZE_CRITERIA_OPTS ("AUM", "NAV", "Assets") so the extract routes the figure straight back into
# the right column.
SIZE_CRITERIA = {"aum": "AUM", "nav": "NAV", "pension_assets": "Assets"}


def classify_ubs(*, agent_cat: str, rated: bool, hnw: bool, spv: bool, itype: str,
                 aum_bil: float | None, nav_bil: float | None,
                 pension_bil: float | None) -> str:
    """The UBS LP Classification for a generated LP.

    The export now STATES this field, so the extract no longer derives it — but the value written has
    to be consistent with the row's other attributes or the sample would be incoherent. This is the
    waterfall lp_db_extract used to run, applied here to the clean values before chaos, which is the
    correct place for it: the generator knows the LP's truth, the reader only knows what it is told.
      1. agent 'Excluded Investors' -> Excluded;
      2. any usable agency rating -> Rated Investor;
    3. HNW (flag, or agent 'PWM Designated Investors') -> HNW Feeder when the vehicle is an SPV, else HNW;
      4. pension assets > $5Bn / > $1Bn -> the two Corp Pension classes;
      5. NAV > $1Bn -> Unrated NAV > $1Bn;
      6. FoF/hedge fund with AUM > $10Bn -> FoF & Other > $10Bn AUM;
      7. catch-all -> Other Institutional."""
    if agent_cat == "Excluded Investors":
        return "Excluded"
    if rated:
        return "Rated Investor"
    if hnw or agent_cat == "PWM Designated Investors":
        return "HNW Feeder (acceptable)" if spv else "HNW (acceptable)"
    if pension_bil is not None and pension_bil > 5:
        return "Corp Pension > $5Bn Assets"
    if pension_bil is not None and pension_bil > 1:
        return "Corp Pension > $1Bn Assets"
    if nav_bil is not None and nav_bil > 1:
        return "Unrated NAV > $1Bn"
    if itype in ("Fund of Funds", "Hedge Fund") and aum_bil is not None and aum_bil > 10:
        return "FoF & Other > $10Bn AUM"
    return "Other Institutional"


def _money(x: float) -> float:
    """A dollar figure at its full stored precision - quantized to the cent, never to the dollar and
    never abbreviated. Goes through Decimal because the alternative is a borrowing base that reads
    54000000.000000007 and a reconciliation check that has to be taught to forgive it."""
    return float(Decimal(repr(float(x))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def _pct8(x: float) -> float:
    """A share (0..1) written as a fraction, kept to 8 decimals - enough that a $100K position in a
    $30Bn facility is still a distinguishable number rather than 0.00."""
    return float(Decimal(repr(float(x))).quantize(Decimal("0.00000001"), rounding=ROUND_HALF_UP))


def pick_weighted(options, weights):
    return random.choices(options, weights=weights, k=1)[0]


def build_investor(idx: int, used_names: set, sponsor_pool: list) -> dict:
    itype = pick_weighted(list(TYPE_WEIGHTS), list(TYPE_WEIGHTS.values()))
    suffixes, measure = TYPE_SPECS[itype]
    while True:
        w2 = random.choice(WORD2)
        base = f"{random.choice(WORD1)}{(' ' + w2) if w2 else ''}"
        name = f"{base} {random.choice(suffixes)}"
        if name not in used_names:
            used_names.add(name)
            break
    agent_cat = pick_weighted([c for c, *_ in AGENT_CATEGORIES],
                              [w for *_, w in AGENT_CATEGORIES])
    spv = random.random() < 0.12
    # PWM Designated Investors and "HNW" are the same fact told by the two banks: the agent's private-wealth
    # tier and UBS's HNW classes. Tying them together here is what keeps the row's "Insitutional vs
    # HNW" flag, its Agent LP Classification and its UBS LP Classification from each describing a
    # different investor.
    hnw = agent_cat == "PWM Designated Investors" or (itype == "Family Office" and random.random() < 0.5)
    if hnw:
        agent_cat = "PWM Designated Investors"
    _, agent_ar, agent_cl = AGENT_BY_NAME[agent_cat]

    # Ratings ARE the "Included Investors (Rated)" bucket, not an attribute sprinkled across it. Every LP in the
    # bucket carries at least one agency rating and no LP outside it carries any, so the agent's
    # label and classify_ubs's answer below are two readings of the same evidence - which is what
    # makes "Included Investors (Non-Rated)" + "Rated Investor", or "Excluded" + a AA- from S&P, unreachable.
    rated = agent_cat == "Included Investors (Rated)"
    sp, mdy, fitch = agency_ratings() if rated else ("", "", "")
    band = rating_band(sp, mdy, fitch)       # None unless rated; drives the matrix row below
    ig = is_investment_grade(sp, mdy, fitch)
    is_pension = measure == "pension_assets"

    # nav-measure LPs (FoF / Hedge Fund) carry manager-level AUM alongside fund NAV, which the
    # "FoF & Other > $10Bn AUM" branch needs. Only the LP's OWN measure reaches the export, as the
    # single LP Size figure — the others exist here purely to classify it.
    aum_bil = size_bil("aum") if measure in ("aum", "nav") else None
    nav_bil = size_bil("nav") if measure == "nav" else None
    pension_bil = size_bil("pension_assets") if is_pension else None
    lp_size_bil = {"aum": aum_bil, "nav": nav_bil, "pension_assets": pension_bil}[measure]

    # No UBS advance rate or concentration limit here: both are resolved per POSITION in main(),
    # because the matrix's rate steps at the facility's funded %, which the LP does not own.
    # Parent names a REAL sponsor LP, drawn from the institutions already built: pe-sub-api resolves
    # the string against lp_master.investor_name, so a parent that names no row links nothing and
    # leaves the sponsor showing no children. Only an SPV has one - a standalone institution IS the
    # ultimate entity, and the empty cell is what pe-sub-api reads as "no parent". Sponsors are drawn
    # from non-SPVs only, which caps the tree at one level and makes a cycle unreachable by
    # construction rather than by check.
    parent = random.choice(sponsor_pool) if spv and sponsor_pool else ""
    if not spv:
        sponsor_pool.append(name)

    return {
        "name": name,
        "parent": parent,
        "spv": "Y" if spv else "N",
        "itype": itype,
        "inst": "HNW" if hnw else "Institutional",
        "ig": "Yes" if ig else "No",
        "cls": agent_cat,
        "agent_ar": agent_ar,
        "agent_cl": agent_cl,
        "sp_rating": sp, "moodys_rating": mdy, "fitch_rating": fitch,
        "band": band,
        "lp_size_bil": lp_size_bil,
        "lp_size_criteria": SIZE_CRITERIA[measure],
        "ubs_cls": classify_ubs(agent_cat=agent_cat, rated=rated, hnw=hnw, spv=spv, itype=itype,
                                aum_bil=aum_bil, nav_bil=nav_bil, pension_bil=pension_bil),
    }


def verify_reconciliation(export_rows: list[dict]) -> None:
    """Re-derive every computed cell from the row's OWN stated inputs and fail the run on any
    disagreement. Called on the clean rows, before the chaos monkey; the columns it checks are all
    sacred or blank-only, so what it proves survives the degradation.

    This is the last line of the user's rule that the calculations have to add up, and it is
    deliberately not a re-run of the code above: it reads the workbook's cells the way pe-sub-api
    will and asks whether they agree with each other. The classification's rate and limit come back
    from BB_MATRIX via the band re-resolved from the notches actually written, the borrowing bases
    from the concentration cap against the facility's own total uncalled.
    """
    # Per FACILITY - the (AccountID, FndName) pair - because one account can carry two funds and
    # pooling their uncalled capital would advance each fund against the other's collateral.
    tot_u_by_fac: Counter = Counter()
    tot_c_by_fac: Counter = Counter()
    for row in export_rows:
        fk = facility_key(row["AccountID"], row["FndName"])
        tot_u_by_fac[fk] += row["Uncalled"]
        tot_c_by_fac[fk] += row["Commitments"]

    problems: list[str] = []

    # Parent is a pointer, so it is checked as one. pe-sub-api resolves it against
    # lp_master.investor_name and links nothing when it matches no row, so a parent naming an entity
    # the export never states is a dangling pointer that shows up downstream as a sponsor with no
    # children. A parent equal to the row's own name is worse: a self-referencing ultimate parent.
    all_names = {row["InvestorName"] for row in export_rows if not _blank(row["InvestorName"])}
    # A sponsor is a non-SPV, so it never carries a parent of its own. Asserting that here is what
    # proves the tree is one level deep and pe-sub-api's cycle guard can never fire on this feed.
    parented = {_as_str(row["InvestorName"]): _as_str(row["Parent"]) for row in export_rows}

    def check(row_no, label, got, want, tol=0.01):
        if abs(float(got) - float(want)) > tol:
            problems.append(f"row {row_no}: {label} = {got}, re-derives to {want}")

    for row_no, row in enumerate(export_rows, start=2):
        fk = facility_key(row["AccountID"], row["FndName"])
        cls, ratings = row["UbsClassification"], (row["SP"], row["Moodys"], row["Fitch"])
        agent_cat = row["Classification"]
        parent = _as_str(row["Parent"])
        if parent:
            if parent == _as_str(row["InvestorName"]):
                problems.append(f"row {row_no}: Parent {parent!r} is the row's own Investor Name")
            elif parent not in all_names:
                problems.append(f"row {row_no}: Parent {parent!r} names no LP in the export")
            elif parented.get(parent):
                problems.append(f"row {row_no}: Parent {parent!r} is itself parented to "
                                f"{parented[parent]!r} - sponsors must be ultimate entities")
        # The label has to be the canonical spelling, not merely a resolvable one: it is the key
        # the row's own UBS Advance Rate, UBS Concentration Limit and UBS Borrowing Base three
        # columns over were resolved with, so anything that is not written exactly as
        # ubs_lp_categories.csv states it is a row whose stated class and stated price were read
        # off two different labels.
        if cls not in UBS_CLASSES or UBS_CANONICAL.get(_norm(cls)) != cls:
            problems.append(f"row {row_no}: UBS LP Classification {cls!r} is not one of the nine "
                            "canonical classes ubs_lp_categories.csv states")
            continue
        stated = [str(r).strip() for r in ratings if not _blank(r)]
        if any(r.upper().replace(".", "").replace("/", "") in ("NR", "NA") for r in stated):
            problems.append(f"row {row_no}: a not-rated token reached an agency column: {stated}")
        if (agent_cat == "Included Investors (Rated)") != bool(stated):
            problems.append(f"row {row_no}: Agent LP Classification {agent_cat!r} with "
                            f"{len(stated)} agency rating(s) - the two disagree")
        if (cls == "Rated Investor") != (agent_cat == "Included Investors (Rated)"):
            problems.append(f"row {row_no}: {cls!r} under agent {agent_cat!r}")
        if (cls == "Excluded") != (agent_cat == "Excluded Investors"):
            problems.append(f"row {row_no}: {cls!r} under agent {agent_cat!r}")

        commit, uncalled, called = row["Commitments"], row["Uncalled"], row["Called"]
        check(row_no, "Called Capital", called, commit - uncalled)
        check(row_no, "% of LP Called", row["CalledPercent"], _pct8(called / commit), 1e-8)
        check(row_no, "% of Capital Commitments", row["PercentOfCommitments"],
              _pct8(commit / tot_c_by_fac[fk]), 1e-8)
        check(row_no, "% of Uncalled Capital", row["PercentOfUncalled"],
              _pct8(uncalled / tot_u_by_fac[fk]), 1e-8)

        want_ar, want_cl = bb_criteria(cls, rating_band(*ratings), called / commit)
        check(row_no, "UBS Advance Rate", row["UBSAR"], want_ar, 1e-9)
        check(row_no, "UBS Concentration Limit", row["UBSCL"], want_cl, 1e-9)
        # The schedule as THIS facility prices it: haircut on an uncommitted sleeve, so the two
        # tranches of one Credit Agreement reconcile to different numbers off the same collateral.
        want_agent_ar, want_agent_cl = agent_terms(agent_cat, row["FndName"])
        check(row_no, "Agent Advance Rate", row["AgentAR"], want_agent_ar, 1e-9)
        check(row_no, "Agent Concentration Limit", row["AgentCL"], want_agent_cl, 1e-9)

        tot_u = tot_u_by_fac[fk]
        excluded = cls == "Excluded"
        ubs_eligible = 0.0 if excluded else min(uncalled, want_cl * tot_u)
        check(row_no, "UBS Excess Concentration", row["UBSExcessConc"],
              _money(uncalled - ubs_eligible))
        check(row_no, "UBS Borrowing Base", row["UBSBB"], _money(ubs_eligible * want_ar))
        agent_ar, agent_cl = want_agent_ar, want_agent_cl
        agent_eligible = uncalled if agent_cl <= 0 else min(uncalled, agent_cl * tot_u)
        agent_excess = 0.0 if excluded else uncalled - agent_eligible
        check(row_no, "Agent Excess Concentration", row["AgentExcessConc"], _money(agent_excess))
        check(row_no, "Agent Borrowing Base", row["AgentBB"],
              _money(0.0 if excluded else (uncalled - agent_excess) * agent_ar))

    if problems:
        head = "\n  ".join(problems[:25])
        raise SystemExit(f"the generated sample does not reconcile ({len(problems)} problem(s)):"
                         f"\n  {head}" + ("\n  ..." if len(problems) > 25 else ""))


def weighted_sample_without_replacement(items, weights, k):
    """Efraimidis-Spirakis: key = U^(1/w); take the k largest keys."""
    keyed = sorted(((random.random() ** (1.0 / w), it) for it, w in zip(items, weights)), reverse=True)
    return [it for _, it in keyed[:k]]


def main() -> int:
    validate_chaos_vocabularies()     # before a single row is built, let alone written
    random.seed(SEED)
    facilities = load_facilities()                    # [(acct, fund, loan_amount), ...]
    base = date(2026, 6, 25)
    # One BBDate per facility, M/D/YYYY (formatted manually: Windows strftime lacks %-m/%-d).
    # Two funds on one account are two facilities running their own BBs, so they date separately.
    fac_bbdate = {}
    for acct, fund, _ in facilities:
        d = base - timedelta(days=random.randint(0, 240))
        fac_bbdate[facility_key(acct, fund)] = f"{d.month}/{d.day}/{d.year}"

    fac_loan = {facility_key(acct, fund): loan for acct, fund, loan in facilities}

    # The sleeves of one Credit Agreement are TWO facilities over ONE LP roster, so they are drawn
    # ONCE: the committed sleeve leads the draw and every other sleeve mirrors the positions it
    # receives (see the mirror pass below). Only leads are ever sampled; an untranched facility is
    # its own lead with no mirrors, so nothing else in this function has to know the difference.
    tranche_siblings, leads = tranche_groups(facilities)

    # Both sleeves are advanced against the SAME uncalled capital - one pool of LP commitments
    # pledged under one Credit Agreement - so the pool a group is apportioned from is sized off the
    # group's COMBINED loan amount, not the lead sleeve's alone. A $120M committed line beside a
    # $180M accordion is $300M of potential draw against one collateral pool, and sizing it from
    # $120M would leave the accordion reading as if it were over-collateralized twice over.
    group_loan = {lead: fac_loan[lead] + sum(fac_loan[s] for s in tranche_siblings[lead])
                  for lead in leads}
    # How many positions a facility draws is proportional to its loan amount, so a $400M facility
    # carries a deeper LP list than an $8M one and the two still hold comparable position sizes.
    fac_weights = {fk: group_loan[fk] * random.uniform(*FAC_WEIGHT_JITTER) for fk in leads}

    used_names: set = set()
    # Sponsor names an SPV may roll up to. Filled with each non-SPV as it is built, so a feeder can
    # only ever name an institution the export actually states - and several feeders can share one.
    sponsor_pool: list[str] = []
    positions: list[dict] = []        # each dict = one lp_record row (pre-percent)
    per_fac: dict = {facility_key(a, f): [] for a, f, _ in facilities}   # sleeves included
    investor_count = 0

    rated_lps = 0
    rows_placed = 0                # positions PLUS the mirrors the tranche pass below will add
    while rows_placed < TARGET_ROWS:
        inv = build_investor(investor_count, used_names, sponsor_pool)
        investor_count += 1
        rated_lps += any(r not in ("NR", "") for r in (inv["sp_rating"], inv["moodys_rating"], inv["fitch_rating"]))
        r = random.randint(REPEAT_MIN, REPEAT_MAX)
        chosen = weighted_sample_without_replacement(leads, [fac_weights[fk] for fk in leads], r)
        for fk in chosen:
            # Trim the last LP on TARGET. A lead that mirrors can overshoot by its sibling count
            # and never by more: a tranche group is placed whole or not at all, because half a
            # Credit Agreement in the sample would be a facility whose sleeve simply went missing.
            if rows_placed >= TARGET_ROWS:
                break
            # Dollars are NOT drawn here: a position's size is a share of its facility's pool, and
            # that pool is not known until every position has been placed. Carry the shape only —
            # how large this position is relative to its peers, and how much of it is still
            # uncalled — and let the apportioning pass below turn the pair into dollars.
            row = {
                **inv,
                "acct": fk[0], "fund": fk[1], "bbdate": fac_bbdate[fk],
                "size_w": random.lognormvariate(0.0, LP_SKEW_SIGMA),
                "uncalled_ratio": random.uniform(0.15, 0.85),
            }
            positions.append(row)
            per_fac[fk].append(row)
            rows_placed += 1 + len(tranche_siblings[fk])

    # Apportioning pass: size each facility's LP pool from its loan amount. A subscription line is
    # advanced against uncalled capital at FACILITY_LTV, so the pool backing a facility is its loan
    # amount grossed up by that fraction, split across its positions on the log-normal weights drawn
    # above. Commitment follows from the position's own uncalled ratio, and called capital is the
    # remainder — the identity commit = called + uncalled that verify_reconciliation checks.
    for fk, rows in per_fac.items():
        if not rows:
            continue
        target_uncalled = group_loan[fk] / random.uniform(FACILITY_LTV_MIN, FACILITY_LTV_MAX)
        tot_w = sum(r["size_w"] for r in rows)
        for r in rows:
            uncalled = snap_money(target_uncalled * r["size_w"] / tot_w, MIN_UNCALLED)
            # Strictly 0 < uncalled < commit: the floor is one grid step above uncalled, so a
            # position can never round to fully-uncalled and leave called capital at zero.
            commit = snap_money(uncalled / r["uncalled_ratio"], uncalled + MONEY_STEP)
            r["uncalled_capital"] = uncalled
            r["commit"] = commit
            r["called"] = commit - uncalled
            r["called_pct"] = r["called"] / commit

    # Mirror pass: give every other sleeve of a Credit Agreement its lead sleeve's roster, position
    # for position. The LPs, their capital commitments, their called and their uncalled capital are
    # IDENTICAL - it is one set of LPs pledging one pool of collateral, and the agent's own file
    # reports the pair with the same LP count for exactly that reason - so only the AccountID, the
    # fund name and the BB date differ. Everything that prices the row is resolved per facility in
    # the pass below, which is where the two sleeves stop agreeing.
    #
    # The mirror is placed directly behind the row it copies, so a tranche pair stays adjacent in
    # the file the way the real export presents it, and dict(row) is a copy rather than a shared
    # reference because that pricing pass writes back into the row it prices.
    if any(tranche_siblings.values()):
        mirrored: list[dict] = []
        for row in positions:
            mirrored.append(row)
            for sib in tranche_siblings.get(facility_key(row["acct"], row["fund"]), ()):
                mirror = dict(row)
                mirror["acct"], mirror["fund"] = sib
                mirror["bbdate"] = fac_bbdate[sib]
                mirrored.append(mirror)
                per_fac[sib].append(mirror)
        positions = mirrored

    # Second pass: everything that needs the facility's totals — the UBS rate and limit, the share
    # percentages, then the two excess-concentration figures and the borrowing bases net of them.
    #
    # This mirrors BbCalculationService.computeOne step for step, because pe-sub-api will re-run
    # exactly this arithmetic over the rows once they are ingested, and a sample that disagreed with
    # the engine would be indistinguishable from the engine being wrong:
    #   * a concentration limit is a fraction of the facility's TOTAL uncalled capital (all LPs,
    #     Excluded ones included - they consume the denominator without contributing base);
    #   * eligible uncalled is the LP's own uncalled capped at that limit, and the excess is the
    #     remainder;
    #   * the borrowing base advances only on the eligible portion.
    # The two sides cut at different points - the agent's tier limit is not UBS's matrix limit - so
    # the same LP legitimately shows different excesses in the two columns.
    #
    # An Excluded LP is the asymmetry worth spelling out: UBS advances nothing and books its ENTIRE
    # uncalled as excess concentration, while the agent side reports zero excess (there is no
    # position to have an excess over). That is computeOne's behaviour, not a rounding artefact.
    for fk, rows in per_fac.items():
        tot_c = sum(r["commit"] for r in rows) or 1
        tot_u = sum(r["uncalled_capital"] for r in rows) or 1
        for r in rows:
            r["pct_commit"] = _pct8(r["commit"] / tot_c)
            r["pct_of_fund_uncalled"] = _pct8(r["uncalled_capital"] / tot_u)
            uncalled = r["uncalled_capital"]
            excluded = r["ubs_cls"] == "Excluded"

            # The advance rate steps at the facility's funded %, so it is resolved per position.
            ubs_ar, ubs_cl = bb_criteria(r["ubs_cls"], r["band"], r["called_pct"])
            r["ubsar"], r["ubs_cl"] = ubs_ar, ubs_cl

            ubs_eligible = 0.0 if excluded else min(uncalled, ubs_cl * tot_u)
            r["ubs_excess_conc"] = _money(uncalled - ubs_eligible)
            r["ubs_borrowing_base"] = _money(ubs_eligible * ubs_ar)

            # The agent prices each sleeve of a Credit Agreement on its own terms and certifies
            # each on its own Borrowing Base Certificate, so the rate and limit are resolved per
            # POSITION from the facility it sits on - not once per LP. UBS's matrix above does not
            # move with the sleeve; agent_terms explains why the asymmetry is deliberate.
            agent_ar, agent_cl = agent_terms(r["cls"], r["fund"])
            r["agent_ar"], r["agent_cl"] = agent_ar, agent_cl
            agent_eligible = uncalled if agent_cl <= 0 else min(uncalled, agent_cl * tot_u)
            agent_excess = 0.0 if excluded else uncalled - agent_eligible
            r["agent_excess_conc"] = _money(agent_excess)
            r["agent_borrowing_base"] = _money(
                0.0 if excluded else (uncalled - agent_excess) * agent_ar)

    # Export rows in SRC_COLS order, as dicts so the chaos monkey can address columns by name.
    export_rows = [dict(zip(SRC_COLS, [
        r["acct"], r["fund"], r["name"], r["parent"], r["spv"], r["ubs_cls"],
        r["inst"], r["ig"], r["cls"], r["sp_rating"], r["moodys_rating"], r["fitch_rating"],
        r["lp_size_bil"], r["lp_size_criteria"], r["commit"], r["uncalled_capital"], r["ubsar"],
        r["agent_ar"], r["agent_cl"], r["ubs_cl"], r["pct_commit"], r["called"], r["pct_of_fund_uncalled"],
        _pct8(r["called_pct"]), r["agent_excess_conc"], r["ubs_excess_conc"],
        r["agent_borrowing_base"], r["ubs_borrowing_base"], "", r["bbdate"],
    ])) for r in positions]

    # Prove the sample adds up while it is still clean; the chaos monkey may only degrade columns
    # this check does not depend on, or blank ones whose fallback restores the checked value.
    verify_reconciliation(export_rows)

    # Degrade what gets written, so the XLSX carries realistic manual-entry quality. Separate rng:
    # the clean base data above is identical whether chaos is on or off.
    chaos_muts: list[tuple] = []
    if CHAOS_ENABLED:
        chaos_muts = apply_chaos(export_rows, random.Random(CHAOS_SEED))

    # Write the workbook.
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    ws.append(SRC_HEADERS)          # the real file's header spellings, not the internal keys
    for row in export_rows:
        ws.append([row[c] for c in SRC_COLS])
    EXPORT_OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(EXPORT_OUT)

    # The export is the only file this script produces; clear any stale chaos log from data/import/.
    EXPORT_OUT.with_name(EXPORT_OUT.stem + ".chaos_log.csv").unlink(missing_ok=True)

    fac_sizes = sorted(len(v) for v in per_fac.values())
    print(f"wrote {EXPORT_OUT}")
    print(f"  rows (lp_records target)  : {len(positions)}")
    print(f"  distinct LPs (lp_master)  : {investor_count}")
    mirrored_sleeves = sum(len(s) for s in tranche_siblings.values())
    tranched = sum(1 for s in tranche_siblings.values() if s)
    print(f"  facilities (incl orphans) : {len(per_fac)}  ({len(ORPHAN_ACCOUNTS)} orphan)")
    print(f"  tranche sleeves mirrored  : {mirrored_sleeves} on {tranched} credit agreement(s)")
    print(f"  LPs/facility  min/med/max : {fac_sizes[0]} / {fac_sizes[len(fac_sizes)//2]} / {fac_sizes[-1]}")
    print(f"  avg repeats per LP        : {len(positions)/investor_count:.1f}")
    print(f"  LPs with agency ratings   : {rated_lps} ({rated_lps/investor_count:.0%})")
    cls_mix = Counter(r["ubs_cls"] for r in positions)
    print("  UBS classification mix    : "
          + ", ".join(f"{c} {n/len(positions):.0%}" for c, n in cls_mix.most_common()))
    band_mix = Counter(r["band"] for r in positions if r["band"])
    print("  rated band mix (positions): "
          + ", ".join(f"{b} {n}" for b, n in band_mix.most_common()))
    print(f"  reconciliation            : all {len(positions)} rows re-derive from their own inputs")
    if CHAOS_ENABLED:
        by_col = Counter(col for _, col, *_ in chaos_muts)
        by_pattern = Counter(pattern for _, _, pattern, *_ in chaos_muts)
        print(f"  chaos monkey (seed {CHAOS_SEED}) : {len(chaos_muts)} value(s) degraded "
              f"({', '.join(f'{c} {n}' for c, n in by_col.most_common())})")
        print(f"    by pattern              : {', '.join(f'{p} {n}' for p, n in by_pattern.most_common())}")
        print(f"    (re-run with CHAOS_SEED={CHAOS_SEED} to reproduce these exactly)")
    else:
        print("  chaos monkey              : disabled (clean export)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
