#!/usr/bin/env python3
r"""
Generate the reproducible sample PAIR lp_db_extract.py reads: the Agent Bank Summary report and
the V2, 32-column LP DB Export. The export contains facility/investor positions with canonical
reference values, reconciled borrowing-base fields, and optional recoverable data-quality
variations; the report states the facilities those positions sit on.

Both are rendered from ONE facility roster (ABS_ROSTER), which is why they agree the way the real
pair does: every account number, borrower name and loan amount in the export is the one the report
prints for that facility, the export's BB date falls inside the facility's own reported life, and
each facility's LP pool is apportioned from the loan amount printed beside it. The two deliberate
disagreements are ORPHAN_ACCOUNTS - positions on accounts the report omits, which is what keeps the
ingestion's "Unknown agent" path exercised - and GROUP_UMBRELLAS, an account the report prints ONCE
under the name of the obligor that signs the credit agreement, with the funds borrowing beneath it
appearing in the export alone. Everything else the export writes - about 99% of its records at the
default settings - joins to a facility the report prints.

The roster itself is MINTED rather than transcribed: build_roster() draws the agents, borrowers,
accounts, amounts and dates from ROSTER_SEED, so the sample states no real institution's book of
business and its size is a tunable (ROSTER_AGENTS, ROSTER_FACILITIES) rather than a rewrite. The
population is random; the SHAPE is not. Umbrellas, split borrowers, reprints, tranche sleeves and
manual-entry damage in the names are generated at the rates the printed report carries them, and
validate_roster() fails the run if a population comes out without them.

The UBS credit columns are never among those variations. UBS LP Classification, UBS Advance Rate,
UBS Concentration Limit and UBS Borrowing Base are always genuine: the classification is written in
the canonical vocabulary of data/reference/ubs_lp_categories.csv, and the rate, the limit and the
base are the Borrowing Base Criteria Matrix's answer for that exact class, re-derived and asserted
per row before the workbook is written.

A fund financed under both sleeves of one Credit Agreement appears TWICE, as it does in the real
file: "<Fund> (Committed)" and "<Fund> (Uncommitted)" carry the same LPs and the same money under
two AccountIDs and price differently on the agent's side. See the tranche section below.

Outputs (both overwritten on every run):
    data/import/AgentBankSummaryRpt.xlsx
    data/import/LP DB Export V2.xlsx

Usage (no command-line arguments):
    1. Edit SEED, CHAOS_SEED, TARGET_ROWS, or CHAOS_ENABLED below as needed.
    2. Run from any directory:
             python pe-sub-jobs/scripts/lp_db_generate.py
    3. Run the extract over the pair:
             python pe-sub-jobs/scripts/lp_db_extract.py

The same SEED, CHAOS_SEED, and settings reproduce the same generated pair.
"""
from __future__ import annotations

import csv
import math
import random
import re
from collections import Counter
from dataclasses import dataclass
from datetime import date, timedelta
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font

SCRIPT_DIR = Path(__file__).resolve().parent          # pe-sub-jobs/scripts/
DATA_DIR = SCRIPT_DIR.parent / "data"                 # pe-sub-jobs/data/
REFERENCE_DIR = DATA_DIR / "reference"                # the same lists lp_db_extract normalizes against
ABS_OUT = DATA_DIR / "import" / "AgentBankSummaryRpt.xlsx"
EXPORT_OUT = DATA_DIR / "import" / "LP DB Export V3.xlsx"
SHEET_NAME = "BBs"
ABS_SHEET_NAME = "Agent Bank Summary"

# ── tunables ────────────────────────────────────────────────────────────────
SEED = 20260908
CHAOS_ENABLED = True            # degrade the written XLSX to realistic manual-entry quality
CHAOS_SEED = 20260908           # chaos has its own rng: base data identical with chaos on/off
TARGET_ROWS = 22_000            # lp_records to produce (mirrored tranche rows counted)
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
# AccountIDs the report omits -> exercise the "Unknown" agent path. Minted with the roster (see
# build_roster), so they are drawn from the same vocabulary and can never collide with a printed
# account. One fund on several accounts, which is the shape the printed report's own orphans take:
# separate credit agreements that must never be grouped as an umbrella.
ORPHAN_COUNT = 2
# The umbrella as the agent actually prints it: ONE row for the account, naming the obligor that
# signs the credit agreement, with the funds that borrow beneath it named only by the export. The
# real file states 5VZ8873 as "Carlyle Buyout Umbrella" and the export puts six Carlyle funds on it.
# This is a different shape from _unit_umbrella, which prints every member, and it is the shape that
# decides whether a report row with no fund to match is read as the GROUP it names or handed to
# whichever member the export happens to list first.
GROUP_UMBRELLA_MEMBERS = 6
AS_OF = date(2026, 6, 25)       # the sample's as-of date: no BB run is dated after it
# How far after a facility's reported FacilityStatusDate its most recent BB run may fall. The
# export's BBDate is not a free-floating date: a borrowing base is certified against a LIVE
# facility, so it sits between the date the agent last stated the facility's status and the date the
# facility matures. Drawing it in that window is what makes the collateral date the extract lifts
# out of the export agree with the facility the report describes, instead of dating a BB run before
# the facility was reported or after it had already matured.
BBDATE_MAX_LAG_DAYS = 240

# The 32 headers of the V2 LP DB Export, in order, EXACTLY as the real file spells them —
# quirks included, because reproducing them is most of the point of this generator: "LP Size" and
# "($ Bil)" are separated by a CRLF inside the one cell, "Insitutional" is misspelt at source, and
# "Moody'S" carries a capital S. lp_db_extract matches headers through _norm(), which absorbs all
# three, so this stays a faithful sample rather than a cleaned-up one.
SRC_HEADERS = [
    "AccountID", "FndName", "Investor Name", "Parent", "Region", "SPV",
    "UBS LP Classification", "Investor Type",
    "Insitutional vs HNW", "Investment Grade?", "Agent LP Classification", "S&P", "Moody'S",
    "Fitch", "LP Size\r\n($ Bil)", "LP Size Criteria", "Capital Commitments", "Uncalled Capital",
    "UBS Advance Rate", "Agent Advance Rate", "Agent Concentration Limit",
    "UBS Concentration Limit",
    "% of Capital Commitments", "Called Capital", "% of Uncalled Capital", "% of LP Called",
    "Agent Excess Concentration", "UBS Excess Concentration", "Agent Borrowing Base",
    "UBS Borrowing Base", "Notes", "BBDate"
]

# ── the multi-tranche declaration ────────────────────────────────────────────────────────────
# The export format is expected to gain two columns so a multi-tranche facility can STATE that it
# is one, instead of leaving the platform to recover the relationship by stripping a
# "(Committed)"/"(Uncommitted)" suffix off a name. A suffix match finds exactly the sleeve names
# someone thought to write the pattern for, and silently misses a currency tranche, a term sleeve
# or an accordion; a stated parent finds all of them.
#
# True writes the columns, which is the shape the extract prefers and the shape to develop
# against. False writes the file exactly as every export to date has arrived, with the sleeve
# suffix as the only evidence - which is the case the extract's legacy reading exists for, and
# which has to keep working for as long as the archive does. Flip it to exercise the other path.
EXPORT_STATES_TRANCHES = True
TRANCHE_SRC_HEADERS = ["Tranche", "Tranche Of"]
TRANCHE_SRC_COLS = ["Tranche", "TrancheOf"]

# Agent Bank Summary column layout, which the report header must match exactly. Index 4 is an
# unnamed spacer holding the report's subtotal amounts.
ABS_COLS = [
    "Agent", "Borrower", "AccountNumber", "LoanAmount", "", "MaturityDate",
    "FacilityStatus", "FacilityStatusDate",
]
ABS_TOTAL_MARKER = "accesstotalsloanamount"   # _norm() prefix of the subtotal / grand-total rows
ABS_TOTAL_LABEL = "AccessTotalsLoanAmount:"       # per-agent subtotal band
ABS_GRAND_TOTAL_LABEL = "AccessTotalsLoanAmount1:"  # the single closing band, and its own label

# The report's print formatting, reproduced cell for cell. It is not decoration: the report is fed
# to lp_db_extract as-is, so a rendering that drifted from the printed one would be a sample of a
# file the bank does not send. Arial 10 throughout with no bold anywhere, money on the report's own
# $ mask, the two date columns as real Excel dates (which is what lets iso_date take them as
# datetimes rather than parsing text), and every column but Agent, Borrower and the unnamed spacer
# centred. Which cells count as money, and when the spacer joins the centred columns, is decided per
# ROW rather than per column - see write_agent_bank_summary's put().
ABS_FONT = Font(name="Arial", size=10)
ABS_ROW_HEIGHT = 12.75
ABS_ZOOM = 110
ABS_COL_WIDTHS = {                                # A..H, as the printed report sets them
    "A": 48.421875, "B": 45.7109375, "C": 27.00390625, "D": 24.28125,
    "E": 21.57421875, "F": 18.421875, "G": 21.8515625, "H": 21.28125,
}
ABS_MONEY_FORMAT = r"[$$-409]#,##0.00;[RED]\-[$$-409]#,##0.00"
ABS_DATE_FORMAT = "mm/dd/yyyy"
ABS_LOAN_COL = 4                                  # LoanAmount, and the totals bands' own label
ABS_SPACER_COL = 5                                # unnamed, and empty except on a totals band
ABS_DATE_COLS = (6, 8)                            # MaturityDate, FacilityStatusDate
ABS_CENTERED_COLS = (3, 4, 6, 7, 8)               # everything but Agent, Borrower and the spacer

# How each agent group's subtotal band writes its SUM. The report does not write them one way, and
# the variants are reproduced rather than tidied up because the file the bank sends carries them:
# a reader that only tolerated one shape would pass here and fail on the real report.
#   range              - =SUM(D3:D18) over the group's own rows;
#   pair               - =SUM(D21,D22), the two-row group spelled as two arguments;
#   single             - =SUM(D25), a one-row group;
#   single_abs         - =SUM($D28), the same with an absolute column;
#   range_one_row_high - =SUM(D76:D77) over a group whose rows are 77 and 78. The range is off by
#                        one, so the band under-reports its own group by the last row and picks up
#                        the agent header above it. That is what the printed report states, and the
#                        subtotal bands are skipped by both scripts, so reproducing the fault costs
#                        nothing and keeps the sample honest about what arrives.
ABS_SUBTOTAL_FORMULA = {
    "range":              lambda a, b: f"=SUM(D{a}:D{b})",
    "pair":               lambda a, b: f"=SUM(D{a},D{b})",
    "single":             lambda a, b: f"=SUM(D{a})",
    "single_abs":         lambda a, b: f"=SUM($D{a})",
    "range_one_row_high": lambda a, b: f"=SUM(D{a - 1}:D{b - 1})",
}

# ── the facility roster: the one model BOTH output files are rendered from ───────────────────
# ABS_ROSTER is [(agent bank, subtotal style, [(borrower, account number, loan amount, maturity,
# facility status, facility status date), ...]), ...], in the printed report's own order: agents
# alphabetical, facilities in the order the agent lists them.
#
# This roster IS the Agent Bank Summary report, and it is also the entire universe of facilities the
# LP DB Export may place a position on. Holding it here rather than reading it back out of the
# rendered workbook is what makes the pair cohesive by construction: the export cannot name an
# account the report does not print, the report cannot print a loan amount the export was not
# apportioned from, and neither file can drift when the other is regenerated. Every LP record the
# export writes therefore joins to a printed facility - bar the deliberate orphans, which are the
# whole point of ORPHAN_COUNT.
#
# The population is MINTED, not transcribed. build_roster() draws the agents, borrowers, accounts,
# amounts and dates from ROSTER_SEED, so the sample states no real institution's book of business
# and its size is a tunable rather than a rewrite. What is not randomized is the SHAPE: the
# structures below are the cases the extract's dirty-row handling exists for, so they are generated
# at the rates the printed report carries them, and validate_roster() fails the run if a population
# comes out without them.
#   * one account carrying several borrowers is an umbrella subscription facility - related funds,
#     feeders and SPVs borrowing under one credit agreement;
#   * ONE row printed for an account whose export carries several funds is the same structure
#     reported at group level: the row names the obligor that signed the credit agreement, not any
#     borrower, and the funds beneath it are named by the export alone (see GROUP_UMBRELLAS);
#   * one borrower printed under two accounts is two separate facilities, not a duplicate, and is
#     never grouped as an umbrella;
#   * the same (account, borrower) printed twice is a reprint, and collapses to ONE facility - it is
#     rendered twice and read once;
#   * "(Committed)" / "(Uncommitted)" pairs are the two sleeves of one credit agreement and are two
#     facilities over one LP roster (see the tranche section below);
#   * names arrive as an operator typed them - truncated mid-word, misspelt, elided across several
#     series ("VI, VII, VIII, … [U]"), upper-cased, trailing-punctuated - and are never cleaned,
#     because the export's FndName has to join back to exactly this string.
ROSTER_SEED = 20260831          # the roster draws from its own rng, so the population holds still
                                # when SEED, CHAOS_SEED or TARGET_ROWS are changed
ROSTER_AGENTS = (25, 35)        # agent groups the report prints
ROSTER_FACILITIES = 150         # facility ROWS the report prints. Reprints collapse on the way in,
                                # so the export sees slightly fewer distinct facilities, plus
                                # ORPHAN_COUNT more that the report never printed. It is a floor: a
                                # population too small to hold one of every structure gets one
                                # anyway, because coverage outranks the row count.

# The quirk mix, as a share of printed rows, at the rates the printed report shows them - its 72
# rows hold 4 tranche pairs, 2 umbrellas over 5 rows, 3 split borrowers and 1 reprint. Each is
# floored at a single unit: a population that rounded a structure away would still render, and would
# silently stop covering the ingestion path that structure exists to exercise.
TRANCHE_ROW_SHARE = 0.11
UMBRELLA_ROW_SHARE = 0.07
SPLIT_ROW_SHARE = 0.08
REPRINT_ROW_SHARE = 0.03

# Manual-entry damage to the borrower names, likewise at the printed report's rates. This is NOT the
# chaos monkey, which degrades the export after the fact and leaves FndName sacred: a name is
# damaged HERE, in the roster, so the report and the export carry the same damaged string and still
# join on it. Damaging it downstream would break the join the sample exists to demonstrate.
NAME_UPPER_RATE = 0.08
NAME_ELIDED_RATE = 0.04
NAME_TRUNCATE_RATE = 0.015
NAME_MISSPELL_RATE = 0.015
NAME_TRAILING_RATE = 0.015
NAME_TRUNCATE_LEN = 43          # the printed report's own truncation point, to the character
ODD_LOAN_RATE = 0.30            # loan amounts that are a pro-rata slice of a line, not a round one

# Name parts. Deliberately invented: the shapes are the printed report's, the words are not anyone's.
AGENT_STEMS = (
    "Ashford", "Baymont", "Calderon", "Dunmore", "Eastvale", "Fairholt", "Granton", "Harlow",
    "Inverness", "Jarrow", "Kelvin", "Langmere", "Marrowby", "Norsted", "Oakbourne", "Pemberly",
    "Quinton", "Ravensden", "Stanmore", "Thornbury", "Ullswater", "Vandermere", "Westcliff",
    "Yarrow", "Zellwood", "Brightlow", "Cravenhurst", "Duxbury", "Elmsworth", "Foxbourne",
)
AGENT_SHAPES = (                # weighted below; "Bank of X" and bare initialisms both occur
    "{a}", "{a} Bank", "{a} National Bank", "{a} Bank USA", "{a} {b}", "Bank of {a}", "{initials}",
)
AGENT_SHAPE_WEIGHTS = (18, 14, 8, 5, 20, 8, 27)
AGENT_INITIAL_LETTERS = "BCDFGHKLMNPRSTUW"

SPONSOR_STEMS = (
    "Ridgeline", "Harborview", "Calder", "Northlake", "Stonebridge", "Kestrel", "Ironwood",
    "Blackford", "Marlowe", "Ashgrove", "Fairhaven", "Brightwater", "Sablecreek", "Westmoor",
    "Cranfield", "Larkspur", "Quarrystone", "Emberly", "Thorncliff", "Baymont", "Glenrock",
    "Silverbrook", "Redstone", "Pinehurst", "Highmark", "Alderwood", "Copperfield", "Drayton",
    "Everly", "Fernbank", "Grayling", "Hollowell", "Ivorydale", "Junction", "Kingsmere", "Loxley",
    "Merriden", "Netherby", "Oysterbay", "Pellworth", "Quillon", "Rothwell", "Standish", "Tarnwick",
    "Upwell", "Vellacott", "Wrenfield", "Yatesbury", "Zephyrus", "Braddock",
)
SPONSOR_SECOND = ("Point", "Ridge", "Creek", "Harbor", "Gate", "Park", "Hill", "Reach", "Cove")
FUND_STRATEGIES = (
    "Buyout", "Credit Opportunities", "Direct Lending", "Strat RE", "Real Estate Opportunistic Debt",
    "Growth", "Secondary Opportunities", "Sen Loan", "Mezz", "Specialty Credit",
    "Asset Based Credit", "GP Stakes", "Continuation", "Tech Adj", "Private Equity Fund",
    "Customized Credit Fund", "Equity Partners", "Capital Partners", "Core Program",
    "Strategic Lending", "Aviation Stable Yield", "Sports Fund", "Keystone Partners Fund",
    "MM LBO", "Priv Inv", "Realty", "Ascendant", "Opportunities Fund", "Global Fund",
)
UMBRELLA_MEMBER_SUFFIXES = ("Fund", "Feeder", "SPV", "Master Fund", "AIF", "Offshore", "Co-Invest")
SERIES_ROMAN = ("II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII", "XIV", "XV")
LEGAL_SUFFIXES = ("", "", "", "", " LP", " L.P.", ", LP", " LLC")
# Round lines on a $25M grid, plus the odd sizes the printed report actually carries.
LOAN_GRID = tuple(range(25_000_000, 401_000_000, 25_000_000)) + (
    45_000_000, 75_000_000, 87_500_000, 112_000_000, 132_500_000, 135_000_000, 165_000_000,
    175_000_000, 215_000_000, 240_000_000, 245_000_000, 273_000_000, 292_000_000,
)


def _mint_sponsor(rng: random.Random) -> str:
    """A sponsor stem, occasionally two words the way the printed report's are."""
    stem = rng.choice(SPONSOR_STEMS)
    return f"{stem} {rng.choice(SPONSOR_SECOND)}" if rng.random() < 0.22 else stem


def _mint_series(rng: random.Random) -> str:
    """The series designation, in the shapes the printed report uses: roman numerals mostly, an
    arabic number now and then, a lettered sleeve ("Xb", "V (A)"), or nothing at all."""
    roll = rng.random()
    if roll < 0.60:
        return rng.choice(SERIES_ROMAN)
    if roll < 0.75:
        return str(rng.randint(4, 9))
    if roll < 0.85:
        roman = rng.choice(SERIES_ROMAN)
        return f"{roman}{rng.choice('ab')}" if rng.random() < 0.5 else f"{roman} ({rng.choice('AB')})"
    return ""


def _misspell(rng: random.Random, name: str) -> str:
    """One inserted letter inside one word - the printed report's "Genercal Atlantic", which is a
    real fund name a real operator really typed wrong, and which the extract must not "correct"."""
    words = name.split(" ")
    candidates = [i for i, w in enumerate(words) if len(w) >= 5 and w.isalpha()]
    if not candidates:
        return name
    i = rng.choice(candidates)
    word = words[i]
    at = rng.randrange(1, len(word) - 1)
    words[i] = word[:at] + rng.choice("acdeilnorst") + word[at:]
    return " ".join(words)


def _elide_series(rng: random.Random, sponsor: str, strategy: str) -> str:
    """Several series of one sponsor collapsed onto one line - "Crestview II, III, VI [U]". The
    facility is real and singular; it is the NAME that covers several funds."""
    picks = rng.sample(SERIES_ROMAN, rng.randint(2, 3))
    tail = ", ".join(picks) + (", …" if rng.random() < 0.3 else "")
    head = f"{sponsor} {strategy}" if rng.random() < 0.5 else sponsor
    return f"{head} {tail} [U]"


def _damage_name(rng: random.Random, name: str, sponsor: str, strategy: str) -> str:
    """At most ONE piece of manual-entry damage per name, as the printed report carries it: a name
    there is truncated or misspelt or upper-cased, not all three."""
    roll = rng.random()
    if roll < NAME_ELIDED_RATE:
        return _elide_series(rng, sponsor, strategy)
    roll -= NAME_ELIDED_RATE
    if roll < NAME_TRUNCATE_RATE:
        return name[:NAME_TRUNCATE_LEN] if len(name) > NAME_TRUNCATE_LEN else name
    roll -= NAME_TRUNCATE_RATE
    if roll < NAME_MISSPELL_RATE:
        return _misspell(rng, name)
    roll -= NAME_MISSPELL_RATE
    if roll < NAME_TRAILING_RATE:
        return name + ";"
    roll -= NAME_TRAILING_RATE
    if roll < NAME_UPPER_RATE:
        return name.upper()
    return name


def _mint_name(rng: random.Random, used: set[str], *, sponsor: str | None = None,
               strategy: str | None = None, damage: bool = True) -> str:
    """One borrower name, unique across the roster under _norm().

    Uniqueness is enforced on the extract's OWN key rather than on the raw string, because two names
    that differ only in punctuation or case are one facility to everything downstream: minting both
    would merge two borrowing bases without a single error being raised."""
    for _ in range(400):
        stem = sponsor or _mint_sponsor(rng)
        strat = rng.choice(FUND_STRATEGIES) if strategy is None else strategy
        parts = [p for p in (stem, strat, _mint_series(rng)) if p]
        name = " ".join(parts) + rng.choice(LEGAL_SUFFIXES)
        if damage:
            name = _damage_name(rng, name, stem, strat)
        key = _norm(name)
        if key and key not in used:
            used.add(key)
            return name
    raise SystemExit("build_roster could not mint a unique borrower name: widen the name vocabulary "
                     "or lower ROSTER_FACILITIES")


def _account_parts(rng: random.Random) -> tuple[str, int, int]:
    """(prefix, number, digit width) for one account, in the two masks the printed report uses:
    5V<letter><4 digits> for most, 5VA<letter><3 digits> for the newer block."""
    if rng.random() < 0.06:
        return "5VA" + rng.choice("ABCD"), rng.randint(0, 998), 3
    return "5V" + rng.choice("TUVWXYZ"), rng.randint(0, 9998), 4


def _mint_account(rng: random.Random, used: set[str]) -> str:
    """One account number no other facility on this roster holds."""
    for _ in range(4000):
        prefix, num, width = _account_parts(rng)
        acct = f"{prefix}{num:0{width}d}"
        if acct not in used:
            used.add(acct)
            return acct
    raise SystemExit("build_roster ran out of account numbers: lower ROSTER_FACILITIES")


def _mint_account_pair(rng: random.Random, used: set[str]) -> tuple[str, str]:
    """Two CONSECUTIVE account numbers, which is how the printed report numbers the two sleeves of
    one credit agreement (5VAB067 beside 5VAB068)."""
    for _ in range(4000):
        prefix, num, width = _account_parts(rng)
        first, second = f"{prefix}{num:0{width}d}", f"{prefix}{num + 1:0{width}d}"
        if first not in used and second not in used:
            used.update((first, second))
            return first, second
    raise SystemExit("build_roster ran out of paired account numbers: lower ROSTER_FACILITIES")


def _mint_agents(rng: random.Random, count: int) -> list[str]:
    """Agent bank names, unique and never "Unknown" - which is the name the extract mints for a
    facility NO agent reported, and would be indistinguishable from a real group if a real group
    could be called that."""
    names: list[str] = []
    seen: set[str] = {"unknown"}
    while len(names) < count:
        shape = rng.choices(AGENT_SHAPES, AGENT_SHAPE_WEIGHTS)[0]
        name = shape.format(
            a=rng.choice(AGENT_STEMS), b=rng.choice(AGENT_STEMS),
            initials="".join(rng.choice(AGENT_INITIAL_LETTERS) for _ in range(rng.randint(3, 4))),
        )
        if _norm(name) not in seen:
            seen.add(_norm(name))
            names.append(name)
    return names


def _facility_dates(rng: random.Random) -> tuple[str, str]:
    """(maturity, facility status date) for one facility, as ISO dates.

    The status date is when the agent last reported the facility, so it is never in the future of
    the sample's own as-of date; the maturity is one to three years past it. Together they are the
    window facility_bbdate draws the export's BB run inside, which is what keeps the collateral date
    the extract lifts out of the export describing the facility the report prints. A facility
    maturing within the next few months is left in rather than pushed out: the printed report has
    those, and they are the ones whose BB window is tightest."""
    status = AS_OF - timedelta(days=rng.randint(14, 380))
    maturity = status + timedelta(days=rng.randint(300, 1150))
    return maturity.isoformat(), status.isoformat()


def _loan_amount(rng: random.Random) -> int:
    """One loan amount: a round line most of the time, a pro-rata slice of one otherwise.

    The printed report carries both - $75,000,000 beside $73,076,924 - because a line syndicated or
    stepped down mid-term prints whatever the arithmetic produced. The odd amounts matter to the
    export: every position on a facility is apportioned FROM this number, so a roster of round
    figures alone would make every reconciliation land suspiciously flat."""
    base = rng.choice(LOAN_GRID)
    if rng.random() < ODD_LOAN_RATE:
        denom = rng.randint(3, 19)
        return max(1_000_000, int(base * rng.randint(denom // 2 + 1, denom) / denom))
    return base


def _facility_row(name: str, acct: str, loan: int, maturity: str, status_date: str) -> tuple:
    """One printed facility row. Every facility the report prints is reported Active: an agent
    listing a facility on its summary is what "active" MEANS here, and the extract reads the status
    straight off this column."""
    return (name, acct, loan, maturity, "Active", status_date)


def _unit_plain(rng: random.Random, accts: set[str], names: set[str]) -> list[tuple]:
    """One ordinary facility: one borrower, one account, on its own."""
    maturity, status = _facility_dates(rng)
    return [_facility_row(_mint_name(rng, names), _mint_account(rng, accts), _loan_amount(rng),
                 maturity, status)]


def _unit_tranche(rng: random.Random, accts: set[str], names: set[str],
                  case_split: bool = False) -> list[tuple]:
    """The two sleeves of one credit agreement: "<Fund> (Committed)" and "<Fund> (Uncommitted)".

    Two facilities over ONE LP roster - the export mirrors the committed sleeve's positions onto the
    uncommitted one - so they share a maturity and a status date, and the report prints them under
    two consecutive accounts. The base name is minted undamaged because both sleeves have to carry
    the SAME base to pair them, and only the CASE is allowed to differ: the sleeves are set up in the
    agent's system one at a time, so "COMVEST CREDIT PARTNERS VII (COMMITTED)" reaching the file
    beside "Comvest Credit Partners VII (Uncommitted)" is an ordinary way for one credit agreement
    to arrive. The extract folds case before pairing, and minting the split spelling here is what
    holds it to that - a pair left ungrouped has its one borrowing base counted once per sleeve.

    `case_split` forces that spelling, for the caller that floors it into every population. One rng
    draw is taken either way, so forcing it moves no other structure off its seed."""
    base = _mint_name(rng, names, damage=False)
    committed, uncommitted = f"{base} (Committed)", f"{base} (Uncommitted)"
    roll = rng.random()
    if case_split or NAME_UPPER_RATE * 2 <= roll < NAME_UPPER_RATE * 3:
        committed = committed.upper()          # one sleeve typed in caps, the other not
    elif roll < NAME_UPPER_RATE * 2:
        committed, uncommitted = committed.upper(), uncommitted.upper()
    names.update((_norm(committed), _norm(uncommitted)))
    first, second = _mint_account_pair(rng, accts)
    maturity, status = _facility_dates(rng)
    loan = _loan_amount(rng)
    # The accordion is sized independently of the committed line about half the time, as the printed
    # report shows it: Arctos prints both sleeves at $100,480,770, Audax at $120M and $180M.
    rows = [_facility_row(committed, first, loan, maturity, status),
            _facility_row(uncommitted, second, loan if rng.random() < 0.5 else _loan_amount(rng),
                 maturity, status)]
    rng.shuffle(rows)     # the report does not list the sleeves in a fixed order
    return rows


def _unit_umbrella(rng: random.Random, accts: set[str], names: set[str], members: int) -> list[tuple]:
    """An umbrella subscription facility: several related borrowers on ONE account.

    A fund, its feeder and its SPV borrow under one credit agreement, which is why the agent reports
    them against a single account number - so they share the account and the credit agreement's
    dates, and each keeps its own loan amount and its own LP roster. This is the structure the
    extract groups on the account number, and it is the exact opposite of the split borrower below,
    which shares a NAME across accounts and must never be grouped."""
    acct = _mint_account(rng, accts)
    sponsor = _mint_sponsor(rng)
    strategy = rng.choice(FUND_STRATEGIES)
    maturity, status = _facility_dates(rng)
    return [_facility_row(_mint_name(rng, names, sponsor=sponsor, strategy=f"{strategy} {suffix}",
                            damage=False),
                 acct, _loan_amount(rng), maturity, status)
            for suffix in rng.sample(UMBRELLA_MEMBER_SUFFIXES, members)]


@dataclass(frozen=True)
class GroupUmbrella:
    """An umbrella the report prints at GROUP level: the obligor's row, and the funds beneath it.

    `obligor` is printed by the report and is NOT a facility - it has no LP roster of its own, and
    the export never names it. `members` are facilities on the same account and are named by the
    export alone. `loan_amount` is the whole agreement's line, as printed; each member carries its
    own share of it."""
    account: str
    obligor: str
    loan_amount: int
    maturity: str
    status_date: str
    members: tuple[tuple[str, int], ...]   # (fund name, its share of the line)


def _unit_group_umbrella(rng: random.Random, accts: set[str], names: set[str],
                         members: int) -> tuple[list[tuple], GroupUmbrella]:
    """(the report's one row, the group it stands for).

    The agent reports the CREDIT AGREEMENT here, not the funds: one row, one account, the whole
    line, named for the obligor that signed it. The funds that draw on it reach the platform only
    through the export, which is why the report row has no fund name to match and why handing it to
    a member is a data loss rather than a near miss - the member takes the agreement's name and its
    entire line, and its siblings, being unprinted, land as Unknown/Inactive placeholders.

    The line is split across the members rather than repeated, so the funds' positions still
    apportion out of the amount the report prints for the account."""
    acct = _mint_account(rng, accts)
    sponsor = _mint_sponsor(rng)
    strategy = rng.choice(FUND_STRATEGIES)
    maturity, status_date = _facility_dates(rng)
    loan = _loan_amount(rng)
    obligor = _mint_name(rng, names, sponsor=sponsor, strategy=f"{strategy} Umbrella", damage=False)

    weights = [rng.uniform(0.5, 2.0) for _ in range(members)]
    total = sum(weights)
    shares = [max(1_000_000, int(loan * w / total)) for w in weights]
    funds = tuple(
        (_mint_name(rng, names, sponsor=sponsor, strategy=f"{strategy} {suffix}", damage=False),
         share)
        for suffix, share in zip(rng.sample(UMBRELLA_MEMBER_SUFFIXES, members), shares))
    return ([_facility_row(obligor, acct, loan, maturity, status_date)],
            GroupUmbrella(acct, obligor, loan, maturity, status_date, funds))


def _unit_split(rng: random.Random, accts: set[str], names: set[str]) -> list[tuple]:
    """One borrower financed under TWO accounts: two credit agreements, two facilities, one name.

    The report prints the same borrower twice with the line split between the accounts, and nothing
    downstream may merge them - not the umbrella grouping (the accounts differ), not the reprint
    dedupe (the pair differs). It is the case that makes (AccountID, FndName) the facility key
    rather than either half of it."""
    name = _mint_name(rng, names)
    maturity, status = _facility_dates(rng)
    total = _loan_amount(rng)
    major = max(1_000_000, int(total * rng.uniform(0.55, 0.90)))
    minor = max(1_000_000, total - major)
    later = (date.fromisoformat(status) - timedelta(days=rng.randint(0, 2))).isoformat()
    return [_facility_row(name, _mint_account(rng, accts), major, maturity, status),
            _facility_row(name, _mint_account(rng, accts), minor, maturity, later)]


def _unit_reprint(rng: random.Random, accts: set[str], names: set[str]) -> list[tuple]:
    """The same (account, borrower) printed twice - a reprint, not a second facility.

    The printed report does this when a facility is restated within the run, and the two rows can
    disagree by a day on the status date. Both scripts dedupe on the (account, name) pair and keep
    the first, so this renders as two rows and reads as one facility; the export never places
    positions on the second."""
    row = _unit_plain(rng, accts, names)[0]
    name, acct, loan, maturity, _status, status_date = row
    restated = (date.fromisoformat(status_date) - timedelta(days=rng.randint(0, 1))).isoformat()
    return [row, _facility_row(name, acct, loan, maturity, restated)]


def _partition(rng: random.Random, units: int, groups: int) -> list[int]:
    """How many units each agent group carries: a long tail, like the printed report's, where two
    agents carry sixteen facilities each and eight carry one. Every group gets at least one, because
    an agent printed with no facilities beneath it is a header row the extract would carry down onto
    the NEXT agent's rows."""
    weights = [rng.paretovariate(1.3) for _ in range(groups)]
    total = sum(weights)
    counts = [max(1, int(units * w / total)) for w in weights]
    while sum(counts) > units:
        i = rng.randrange(groups)
        if counts[i] > 1:
            counts[i] -= 1
    while sum(counts) < units:
        counts[rng.randrange(groups)] += 1
    return counts


def build_roster() -> tuple[list[tuple], list[tuple[str, str]], list[GroupUmbrella]]:
    """(ABS_ROSTER, ORPHAN_ACCOUNTS, GROUP_UMBRELLAS), minted from ROSTER_SEED.

    Facilities are built as UNITS rather than rows, because the structures that matter are
    multi-row: a tranche pair, an umbrella and a split borrower each have to reach the same agent
    group, in adjacent rows, the way an agent reports them. Units are drawn to the quirk shares
    first and the remainder filled with ordinary facilities, then shuffled and dealt out to agents -
    so which agent carries the umbrella moves with the seed, but the report always HAS one.

    The result is sorted by agent name, as the printed report sorts it, and each group's subtotal
    style is chosen to fit its size (a "pair" band cannot spell a three-row group). Exactly one
    eligible group gets the off-by-one band the printed report contains, so that fault stays
    covered."""
    rng = random.Random(ROSTER_SEED)
    accts: set[str] = set()
    names: set[str] = set()

    tranche_units = max(1, round(ROSTER_FACILITIES * TRANCHE_ROW_SHARE / 2))
    split_units = max(1, round(ROSTER_FACILITIES * SPLIT_ROW_SHARE / 2))
    reprint_units = max(1, round(ROSTER_FACILITIES * REPRINT_ROW_SHARE / 2))
    umbrella_rows = max(2, round(ROSTER_FACILITIES * UMBRELLA_ROW_SHARE))

    units: list[list[tuple]] = []
    # One pair is minted with its sleeves cased differently, floored the way every other structure
    # is. Left to its own rate the split spelling rounds away from a population this size about half
    # the time, and a roster without it renders and extracts perfectly while quietly no longer
    # proving the extract folds case before pairing sleeves.
    units += [_unit_tranche(rng, accts, names, case_split=(i == 0)) for i in range(tranche_units)]
    units += [_unit_split(rng, accts, names) for _ in range(split_units)]
    units += [_unit_reprint(rng, accts, names) for _ in range(reprint_units)]
    while umbrella_rows >= 2:
        members = 3 if umbrella_rows >= 3 and rng.random() < 0.4 else 2
        units.append(_unit_umbrella(rng, accts, names, members))
        umbrella_rows -= members
    # Floored at one per roster rather than drawn to a share: it is a single printed row, so a rate
    # would round it away on most populations, and a roster without it renders and extracts
    # perfectly while quietly no longer proving that a report row naming a group is read as one.
    group_rows, group = _unit_group_umbrella(rng, accts, names, GROUP_UMBRELLA_MEMBERS)
    units.append(group_rows)
    umbrella_groups = [group]
    units += [_unit_plain(rng, accts, names)
              for _ in range(max(0, ROSTER_FACILITIES - sum(len(u) for u in units)))]

    rng.shuffle(units)
    groups = max(1, min(rng.randint(*ROSTER_AGENTS), len(units)))
    dealt = _partition(rng, len(units), groups)
    blocks: list[list] = []
    at = 0
    for agent, count in zip(_mint_agents(rng, groups), dealt):
        blocks.append([agent, None, [row for unit in units[at:at + count] for row in unit]])
        at += count

    # Subtotal bands. The style has to FIT the group - a "pair" band cannot spell three rows - but
    # which fitting style a group gets is not left to the draw: every variant ABS_SUBTOTAL_FORMULA
    # spells is placed at least once wherever a group of the right size exists, including the
    # printed report's off-by-one band. A reader that tolerated only one shape would pass on a
    # population that happened to mint one shape and fail on the report the bank actually sends, so
    # the sample carries them all in every run rather than most runs.
    for block in blocks:
        block[1] = "range" if len(block[2]) > 1 else rng.choice(("single", "single_abs"))
    singles = [b for b in blocks if len(b[2]) == 1]
    pairs = [b for b in blocks if len(b[2]) == 2]
    for style, block in zip(("single", "single_abs"), rng.sample(singles, min(2, len(singles)))):
        block[1] = style
    if pairs:
        rng.choice(pairs)[1] = "pair"
    # The off-by-one band goes on a group of three or more where there is one, so it cannot land on
    # the group just given the "pair" band and quietly take that variant back out of the sample.
    faulty = [b for b in blocks if len(b[2]) >= 3] or pairs
    if faulty:
        rng.choice(faulty)[1] = "range_one_row_high"

    roster = [(agent, style, rows)
              for agent, style, rows in sorted(blocks, key=lambda b: _norm(b[0]))]

    # The orphans: one fund financed under several accounts NONE of which the report prints. Minted
    # last, from the same vocabulary and against the same used-account set, so an orphan can never
    # collide with a printed facility and quietly stop being an orphan.
    orphan_name = _mint_name(rng, names, damage=False)
    orphans = [(_mint_account(rng, accts), orphan_name) for _ in range(ORPHAN_COUNT)]
    return roster, orphans, umbrella_groups


# Internal keys for the same columns, in the same order. Mirrors lp_db_extract.SRC_COLS so the
# chaos monkey can address a column by name and the two scripts stay legible side by side.
# Region and Investor Type are V2's two additions, written where V2 puts them.
SRC_COLS = [
    "AccountID", "FndName", "InvestorName", "Parent", "Region", "SPV",
    "UbsClassification", "InvestorType",
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
    "Corporate":             (["Treasury", "Corporate Holdings", "Group Treasury"], "aum"),
    "Healthcare":            (["Health System", "Hospital Trust", "Healthcare Endowment"], "aum"),
    "Investment Consultant": (["Investment Advisors", "Capital Advisors", "Consulting Group"], "aum"),
    "Institutional Investor":(["Institutional Trust", "Alternative Assets Trust", "Capital Partners"], "aum"),
    "Other Institutional":   (["Strategic Capital Partners", "Alternative Assets", "Global Investors"], "aum"),
}
# Region, from the platform's REGION_OPTS. Free text downstream - the extract carries this column
# as fed rather than normalizing it - so the generator states the platform's own spellings and lets
# the chaos monkey supply the drift, rather than inventing a second vocabulary here.
REGION_WEIGHTS = {
    "North America": 46, "Europe": 24, "Asia-Pacific": 16, "Middle East": 9, "Other": 5,
}
# Where a type actually raises from, so the sample does not read as region sprinkled at random over
# investor type. Only the types with a real skew are listed; everything else takes REGION_WEIGHTS.
REGION_BY_TYPE = {
    "Sovereign Wealth Fund": {"Middle East": 45, "Asia-Pacific": 30, "Europe": 12,
                              "North America": 5, "Other": 8},
    "Public Pension":        {"North America": 62, "Europe": 22, "Asia-Pacific": 11,
                              "Middle East": 2, "Other": 3},
    "Endowment":             {"North America": 82, "Europe": 10, "Asia-Pacific": 5,
                              "Middle East": 1, "Other": 2},
    "Healthcare":            {"North America": 80, "Europe": 12, "Asia-Pacific": 5,
                              "Middle East": 1, "Other": 2},
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
# the agent bank requires a SEPARATE Borrowing Base Certificate per facility ID, so each tranche's
# headroom is certified and tracked on its own.
#
# Per the Day 1 agreement with the business, the platform therefore treats each sleeve as its own
# record - which is why the generator emits each as its own facility rather than folding the pair
# into one, and why this section exists to keep it doing so.
#
# What the sleeves DO share is the collateral. One borrower pledges one LP roster once, and both
# certificates are drawn against it - which is why this generator mirrors the committed sleeve's
# positions onto the uncommitted one rather than minting a second roster. Two records over one
# pool, certified twice and pledged once, so a report that adds the sleeves' bases together counts
# the collateral twice over. The extract states the pair as one cross-collateralized group and the
# platform divides that one base between them; see lp_db_extract.assign_umbrellas.
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
    "investor_type": 0.18,     # 'Pension Fund' -> 'Corp Pension' etc. - decided once per LP
    "region": 0.14,            # 'North America' -> 'USA' / 'N. America' - decided once per LP
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

# Region drift. Unlike Investor Type below, region is carried as fed - the extract normalizes it
# against nothing - so these spellings reach LP Master exactly as written. That is the point: region
# is a label the banks write freely, and a sample that only ever spelt it the platform's way would
# not show what the column actually looks like. Drift is decided once per LP, so an LP's own profile
# still reads as one region told one way and not as three.
REGION_DRIFT = {
    "North America": ["USA", "US", "N. America", "United States", "North America (US)"],
    "Europe":        ["EMEA", "Western Europe", "EU", "UK/Europe"],
    "Asia-Pacific":  ["APAC", "Asia Pacific", "Asia", "Asia/Pac"],
    "Middle East":   ["MENA", "Middle East & Africa", "GCC"],
    "Other":         ["Global", "Other / Global", "LatAm"],
}

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
# norm(canonical or alias) -> canonical Investor Type. Used BOTH ways here, unlike the UBS list
# above: to prove the types the generator draws are canonical, and to source the alias spellings the
# chaos monkey drifts them to - so a drifted cell is always one the extract resolves back to the
# type the row was built as.
ITYPE_CANONICAL = [r[0] for r in _reference_rows("investor_types.csv")[1:] if r[0]]
ITYPE_LOOKUP = {_norm(c): c for c in ITYPE_CANONICAL}
ITYPE_ALIASES: dict[str, list[str]] = {}
for _alias, _canon in _alias_lookup("investor_type_aliases.csv").items():
    if _alias not in ITYPE_LOOKUP:
        ITYPE_LOOKUP[_alias] = _canon
for _row in _reference_rows("investor_type_aliases.csv")[1:]:
    if len(_row) >= 2 and _row[1] and _norm(_row[0]) not in {_norm(c) for c in ITYPE_CANONICAL}:
        ITYPE_ALIASES.setdefault(_row[1], []).append(_row[0])
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
        the row;
      * every Investor Type the generator draws is canonical, and every alias the chaos monkey can
        drift one to resolves back to it - so a drifted cell costs the sample its spelling and
        never its type."""
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

    # Investor Type is normalized by the extract, so the drift the chaos monkey applies has to be
    # reversible. Checked in both directions: a type drawn here that the canonical list does not
    # state would reach LP Master as its own dropdown entry, and an alias that resolves to anything
    # other than the type it replaced would silently reclassify the LP.
    for itype in TYPE_WEIGHTS:
        if ITYPE_LOOKUP.get(_norm(itype)) != itype:
            problems.append(f"Investor Type {itype!r} is not a canonical value stated by "
                            f"investor_types.csv (resolves to {ITYPE_LOOKUP.get(_norm(itype))!r})")
        for alias in ITYPE_ALIASES.get(itype, []):
            if ITYPE_LOOKUP.get(_norm(alias)) != itype:
                problems.append(f"Investor Type alias {alias!r} resolves to "
                                f"{ITYPE_LOOKUP.get(_norm(alias))!r}, not {itype!r} - drifting a "
                                "row to it would change the LP's type, not its spelling")
    for region in REGION_WEIGHTS:
        if region not in REGION_DRIFT:
            problems.append(f"Region {region!r} has no REGION_DRIFT spellings")

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
    # Investor Type and Region are attributes of the LP, not of the submission, so they are decided
    # once per LP for the same reason the name is: an LP whose type read "Pension Fund" on one row
    # and "Corp Pension" on the next would not look like one analyst's spelling, it would look like
    # the LP changed - and build_master consolidates from the most recent row, so which spelling
    # survived would be an accident of BBDate ordering.
    itype_variant: dict[str, tuple[str, str]] = {}
    region_variant: dict[str, tuple[str, str]] = {}
    for row in export_rows:
        original = _as_str(row["InvestorName"])
        if not original:
            continue
        rows_per_lp[original] += 1
        if original not in decided:
            decided.add(original)
            if original not in sponsors and rng.random() < CHAOS_RATES["investor_name"]:
                name_variant[original] = _chaos_name(original, rng)
            # Only spellings the alias list resolves BACK to the type the row was built as. The
            # investor type is the one drifted cell with a normalization behind it, so a drift the
            # extract could not resolve would not be dirty data, it would be a different type.
            aliases = ITYPE_ALIASES.get(_as_str(row["InvestorType"]), [])
            if aliases and rng.random() < CHAOS_RATES["investor_type"]:
                itype_variant[original] = ("itype alias", rng.choice(aliases))
            drifts = REGION_DRIFT.get(_as_str(row["Region"]), [])
            if drifts and rng.random() < CHAOS_RATES["region"]:
                region_variant[original] = ("region alias", rng.choice(drifts))

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
        if investor in itype_variant:
            mutate(row_no, row, "InvestorType", itype_variant[investor])
        if investor in region_variant:
            mutate(row_no, row, "Region", region_variant[investor])
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


def tranche_parent_name(fund_name: str) -> str:
    """The facility a sleeve is a sleeve OF, spelt as the file spells it.

    This is what the export's `Tranche Of` column states. Unlike tranche_base_name it is not
    casefolded, because it is read by a person as well as matched by a machine."""
    return TRANCHE_SUFFIX_RE.sub("", _as_str(fund_name)).strip()


def tranche_base_name(fund_name: str) -> str:
    """The fund name with its tranche suffix stripped, casefolded - the key the sleeves of one
    Credit Agreement share and nothing else does."""
    return tranche_parent_name(fund_name).casefold()


def tranche_groups(facilities: list[tuple]) -> tuple[dict, list]:
    """({lead facility -> its mirror facilities}, [lead facilities]), each a facility_key pair.

    Facilities whose names share a base are the sleeves of one Credit Agreement. The committed
    sleeve leads (falling back to the first listed, so a group the report spells without an
    explicit "(Committed)" member still has exactly one lead) and every other sleeve mirrors it.

    Only a NAMED tranche is ever grouped: two facilities that merely share a fund name - the
    ORPHAN_ACCOUNTS pair, or one fund financed twice under separate accounts - are separate credit
    agreements, not two sleeves of one, and keep their own independent LP rosters. A facility with
    no sibling is its own lead, so the caller can treat every facility the same way."""
    by_base: dict = {}
    for acct, fund, *_ in facilities:
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


# Minted here, at module load, rather than beside the tunables that shape it: build_roster() keys
# borrower uniqueness on _norm() - the extract's own matcher, so that two names the platform would
# treat as one facility can never both be printed - and _norm is defined further down this file than
# the roster's constants are. Everything below reads ABS_ROSTER as the fixed model of the run.
ABS_ROSTER, ORPHAN_ACCOUNTS, GROUP_UMBRELLAS = build_roster()
# The obligor rows: printed by the report, and facilities in neither file. Held as a set because
# every pass that reads ABS_ROSTER as the export's facility list has to step over them.
GROUP_OBLIGORS = {(g.account, _norm(g.obligor)) for g in GROUP_UMBRELLAS}


def load_facilities() -> list[tuple[str, str, int, str | None, str | None]]:
    """(account number, borrower name, loan amount, maturity date, facility status date) per
    facility the roster states, plus ORPHAN_ACCOUNTS.

    This is the export's whole facility universe, and it is read out of ABS_ROSTER - the same model
    write_agent_bank_summary renders the report from - so a position can only ever land on an
    account the report prints, at the loan amount printed beside it. The rendered report is not read
    back to obtain it: reading a file this script had just written would make the pair agree by
    coincidence of timing rather than by construction, and would lose exactly the rows the roster
    keeps deliberately (a reprint collapses on the way in, so it would stop being reprinted).

    The dedupe below is the extract's, rule for rule, so the two scripts resolve the roster to the
    same facilities. A repeated (account, borrower) pair is a reprint of a row already taken and is
    dropped. Two borrowers listed against ONE account are two facilities and both are kept: an
    account number does not identify a facility on its own and neither does a fund name, so the pair
    of them is what every per-facility figure below is grouped by. The name is carried exactly as
    PRINTED, because that is what the export's FndName has to join back to.

    The loan amount is what every position on the facility is apportioned from, so a facility that
    states none falls back to DEFAULT_LOAN_AMOUNT rather than to zero. The two dates are the
    facility's reported life, which is the window its BB run has to fall inside; the orphans carry
    neither, because no report row states one for them.
    """
    out: list[tuple[str, str, int, str | None, str | None]] = []
    seen: set[tuple[str, str]] = set()
    for _agent, _style, rows in ABS_ROSTER:
        for name, acct, loan, maturity, _status, status_date in rows:
            if not name or not acct or (acct, _norm(name)) in seen:
                continue
            # An obligor row names the credit agreement, not a borrower: the funds that draw on it
            # are the facilities, and they are added below. Placing positions on the obligor too
            # would double the account's collateral and make the group a seventh fund.
            if (acct, _norm(name)) in GROUP_OBLIGORS:
                continue
            seen.add((acct, _norm(name)))
            out.append((acct, name, int(loan) or DEFAULT_LOAN_AMOUNT, maturity, status_date))
    # The member funds of a group-level umbrella: real facilities on a printed account, each with
    # its own LP roster and its own share of the printed line, and named by the export alone.
    for group in GROUP_UMBRELLAS:
        for fund, share in group.members:
            out.append((group.account, fund, share, group.maturity, group.status_date))
    # The orphans are the sample's one deliberate disagreement between the two files: positions on
    # accounts the report does not print, which is what leaves the ingestion's "Unknown agent"
    # placeholder path reachable. They are facilities in every other respect.
    out.extend((acct, fund, DEFAULT_LOAN_AMOUNT, None, None) for acct, fund in ORPHAN_ACCOUNTS)
    return out


def validate_roster() -> None:
    """Assert the roster can be rendered as the report AND used as the export's facility list,
    BEFORE anything is written.

    Each check is a way the two files could quietly stop describing the same book of business:
      * a subtotal style the renderer does not implement would abort mid-write, after the export had
        already been produced against a roster the report never got;
      * a group whose row count the style cannot spell ("pair" over three rows) would print a
        subtotal that silently omits facilities;
      * an ORPHAN_ACCOUNTS entry the roster also states would not be an orphan at all - the report
        would print it, and the "Unknown agent" path it exists to exercise would go untested;
      * an umbrella (one account, several borrowers) and a tranche pair are the two structures the
        ingestion has dedicated handling for, so the sample asserts it still contains one of each
        rather than leaving a roster edit to remove the coverage unnoticed.
    """
    problems: list[str] = []
    printed: set[tuple[str, str]] = set()
    funds_per_account: dict[str, set[str]] = {}

    # Both scripts skip a totals band by the marker alone, so a relabelled band would not read as a
    # total - it would read as a facility named "AccessTotals..." with no account number.
    for label in (ABS_TOTAL_LABEL, ABS_GRAND_TOTAL_LABEL):
        if not _norm(label).startswith(ABS_TOTAL_MARKER):
            problems.append(f"totals band label {label!r} does not carry the "
                            f"{ABS_TOTAL_MARKER!r} marker both scripts skip it on")

    for agent, style, rows in ABS_ROSTER:
        if style not in ABS_SUBTOTAL_FORMULA:
            problems.append(f"agent group {agent!r} uses subtotal style {style!r}, which "
                            f"ABS_SUBTOTAL_FORMULA does not spell")
        if style in ("single", "single_abs") and len(rows) != 1:
            problems.append(f"agent group {agent!r} spells its subtotal as a single row but lists "
                            f"{len(rows)}")
        if style == "pair" and len(rows) != 2:
            problems.append(f"agent group {agent!r} spells its subtotal as a pair but lists "
                            f"{len(rows)}")
        if not rows:
            problems.append(f"agent group {agent!r} lists no facilities")
        for name, acct, loan, maturity, status, status_date in rows:
            if not name or not acct:
                problems.append(f"agent group {agent!r} has a row with no borrower or account")
                continue
            if int(loan) <= 0:
                problems.append(f"facility {name!r} ({acct}) states loan amount {loan}")
            for label, value in (("maturity", maturity), ("status date", status_date)):
                try:
                    date.fromisoformat(value)
                except (TypeError, ValueError):
                    problems.append(f"facility {name!r} ({acct}) states {label} {value!r}, which "
                                    "is not an ISO date")
            if not status:
                problems.append(f"facility {name!r} ({acct}) states no facility status")
            printed.add((acct, _norm(name)))
            funds_per_account.setdefault(acct, set()).add(_norm(name))

    for acct, fund in ORPHAN_ACCOUNTS:
        if (acct, _norm(fund)) in printed:
            problems.append(f"orphan facility {fund!r} ({acct}) is also printed by the report, so "
                            "it is not an orphan and exercises no 'Unknown agent' path")

    if not any(len(funds) > 1 for funds in funds_per_account.values()):
        problems.append("no account carries more than one borrower - the roster states no umbrella "
                        "subscription facility, so that path is unexercised")

    # The group-level umbrella. Its whole point is that the report row has NO fund to match: the
    # obligor must be printed, must be alone on its account, and must not be a facility in the
    # export - any of those failing turns it back into an ordinary umbrella and stops it proving
    # that an unmatched row on a multi-fund account is read as the group it names.
    if not GROUP_UMBRELLAS:
        problems.append("the roster states no group-level umbrella, so nothing proves an Agent Bank "
                        "Summary row naming an obligor rather than a fund is read as a group")
    for group in GROUP_UMBRELLAS:
        if (group.account, _norm(group.obligor)) not in printed:
            problems.append(f"group umbrella {group.obligor!r} ({group.account}) is not printed by "
                            "the report, so no row states the agreement at all")
        if len(funds_per_account.get(group.account, ())) != 1:
            problems.append(f"account {group.account} prints "
                            f"{len(funds_per_account.get(group.account, ()))} borrowers - a "
                            "group-level umbrella is printed once, as the obligor")
        if len(group.members) < 2:
            problems.append(f"group umbrella {group.obligor!r} states {len(group.members)} member "
                            "fund(s) - one fund on an account is an ordinary rename, not a group")
        for fund, share in group.members:
            if (group.account, _norm(fund)) in printed:
                problems.append(f"member fund {fund!r} ({group.account}) is printed by the report, "
                                "so its group's row would match it and never read as a group")
            if share <= 0:
                problems.append(f"member fund {fund!r} ({group.account}) states share {share}")
        if sum(share for _f, share in group.members) > group.loan_amount:
            problems.append(f"group umbrella {group.obligor!r} allocates more to its members than "
                            "the line the report prints for the account")
    if not any(tranche_of(name) for _a, _s, rows in ABS_ROSTER for name, *_ in rows):
        problems.append("the roster states no '(Committed)' / '(Uncommitted)' pair, so the tranche "
                        "mirror pass has nothing to mirror")
    # The sleeves are set up in the agent's system one at a time, so a pair spelt two ways is an
    # ordinary arrival, and it is the only thing proving the extract folds case before pairing.
    # A pair it fails to group has its one borrowing base counted once per sleeve.
    spellings: dict[str, set[str]] = {}
    for _a, _s, rows in ABS_ROSTER:
        for name, *_ in rows:
            if tranche_of(name):
                spellings.setdefault(tranche_base_name(name), set()).add(
                    TRANCHE_SUFFIX_RE.sub("", name).strip())
    if not any(len(spelt) > 1 for spelt in spellings.values()):
        problems.append("every tranche pair spells its base name identically across both sleeves, "
                        "so nothing states that case alone must not decide whether the sleeves are "
                        "grouped onto one borrowing base")

    # The remaining structures a MINTED population could round away. Each is a distinct thing the
    # ingestion has to get right, and none of them announces its absence: a roster without them
    # renders and extracts perfectly, and simply stops proving anything about the case it dropped.
    all_rows = [row for _a, _s, rows in ABS_ROSTER for row in rows]
    if len(all_rows) == len({(acct, _norm(name)) for name, acct, *_ in all_rows}):
        problems.append("no (account, borrower) pair is printed twice - the roster states no "
                        "reprint, so the dedupe that collapses one is unexercised")
    accounts_per_fund: dict[str, set[str]] = {}
    for name, acct, *_ in all_rows:
        accounts_per_fund.setdefault(_norm(name), set()).add(acct)
    if not any(len(accts) > 1 for accts in accounts_per_fund.values()):
        problems.append("no borrower is printed under more than one account - the roster states no "
                        "split borrower, so nothing proves a fund name alone does not key a "
                        "facility")
    # Subtotal-band coverage, asserted only where a group of the fitting size exists: a style the
    # population never mints is a shape of band the sample stops proving both scripts skip.
    styles = {style for _a, style, _r in ABS_ROSTER}
    sizes = [len(rows) for _a, _s, rows in ABS_ROSTER]
    one_row = sizes.count(1)
    for style, needs in (("single", one_row >= 1), ("single_abs", one_row >= 2),
                         ("pair", 2 in sizes), ("range", sum(s > 1 for s in sizes) >= 2),
                         ("range_one_row_high", any(s > 1 for s in sizes))):
        if needs and style not in styles:
            problems.append(f"no agent group spells its subtotal band {style!r}, though a group of "
                            f"the fitting size exists - that band shape goes unexercised")

    if problems:
        raise SystemExit("the facility roster would not render as the report it stands for:\n  "
                         + "\n  ".join(problems))


def write_agent_bank_summary(bbdates: dict[tuple[str, str], str]) -> tuple[int, int]:
    """Render ABS_ROSTER as the Agent Bank Summary report and return (facility rows, grand total).

    The layout is the printed report's, band for band, because lp_db_extract reads THIS file and a
    tidied-up rendering would be a sample of a report the bank never sends:
      * one header row, with the LoanAmount subtotal column left unnamed;
      * per agent, a group-header row carrying the agent alone (Borrower blank, which is what marks
        it as a header and lets the agent be carried down onto the rows beneath it), then one row
        per facility with the Agent cell EMPTY, then an "AccessTotalsLoanAmount:" band whose amount
        sits in the unnamed column;
      * one closing "AccessTotalsLoanAmount1:" band over the whole report.

    The subtotal bands are written as live formulas over the rows they actually cover, so the file
    recalculates in Excel; the closing band is written as a literal, as the printed report writes
    it, so the total is readable without a formula engine. Both are skipped by both scripts on the
    "AccessTotalsLoanAmount" marker, which is why the report's own subtotal arithmetic never
    reaches the platform.

    `bbdates` is accepted only to be asserted against: every facility the report prints and the
    export placed positions on must have dated its BB run inside the life this report states for
    it.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = ABS_SHEET_NAME
    ws.sheet_view.zoomScale = ABS_ZOOM
    for column, width in ABS_COL_WIDTHS.items():
        ws.column_dimensions[column].width = width

    def put(row_no: int, values: list, *, totals_band: bool = False) -> None:
        """One report row, styled the way the printed report styles that KIND of row.

        The money mask sits on the cells that hold money, which is not the same as sitting on a
        column: the LoanAmount column carries it everywhere except on a totals band, where that
        cell holds the band's label instead, and the unnamed spacer carries it only ON a totals
        band, which is the one place it holds anything. The spacer's centring follows the same
        rule, for the same reason."""
        for col, value in enumerate(values, start=1):
            cell = ws.cell(row_no, col, value)
            cell.font = ABS_FONT
            if col == ABS_LOAN_COL and not totals_band:
                cell.number_format = ABS_MONEY_FORMAT
            elif col == ABS_SPACER_COL and totals_band:
                cell.number_format = ABS_MONEY_FORMAT
            elif col in ABS_DATE_COLS:
                cell.number_format = ABS_DATE_FORMAT
            if col in ABS_CENTERED_COLS or (col == ABS_SPACER_COL and totals_band):
                cell.alignment = Alignment(horizontal="center", vertical="bottom")
        ws.row_dimensions[row_no].height = ABS_ROW_HEIGHT

    # Written from ABS_COLS itself, which is the schema both scripts check the header against
    # positionally - so the file cannot be rendered with a header the reader would reject. The
    # unnamed spacer is written as an EMPTY cell rather than skipped: the column has to be present
    # and unnamed, not absent.
    put(1, [name or None for name in ABS_COLS])

    row_no = 2
    facility_rows = 0
    grand_total = 0
    for agent, style, rows in ABS_ROSTER:
        put(row_no, [agent, None, None, None, None, None, None, None])
        row_no += 1
        first = row_no
        for name, acct, loan, maturity, status, status_date in rows:
            put(row_no, [None, name, acct, int(loan), None,
                         date.fromisoformat(maturity), status, date.fromisoformat(status_date)])
            grand_total += int(loan)
            facility_rows += 1
            row_no += 1
        put(row_no, [None, None, None, ABS_TOTAL_LABEL,
                     ABS_SUBTOTAL_FORMULA[style](first, row_no - 1), None, None, None],
            totals_band=True)
        row_no += 1
    put(row_no, [None, None, None, ABS_GRAND_TOTAL_LABEL, grand_total, None, None, None],
        totals_band=True)

    verify_bbdates_within_facility_life(bbdates)
    ABS_OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(ABS_OUT)
    return facility_rows, grand_total


def verify_bbdates_within_facility_life(bbdates: dict[tuple[str, str], str]) -> None:
    """Fail the run if any facility's BB run is dated outside the life the report states for it.

    The export's BBDate is the collateral date the extract lifts onto the facility row, so it is a
    claim about the facility the report describes on the very next line of the feed: a BB certified
    before the agent last reported the facility's status, or after it matured, is the pair
    disagreeing with itself in the one column that joins them."""
    lives = {(acct, name): (maturity, status_date)
             for _a, _s, rows in ABS_ROSTER
             for name, acct, _loan, maturity, _status, status_date in rows}
    # A member of a group-level umbrella is unprinted but not undated: the agreement's own life is
    # the window its BB run is certified inside, so it is held to that rather than skipped.
    lives.update({(g.account, fund): (g.maturity, g.status_date)
                  for g in GROUP_UMBRELLAS for fund, _share in g.members})
    problems = []
    for fk, bbdate in bbdates.items():
        life = lives.get(fk)
        if life is None:                       # an orphan: no report row states a life for it
            continue
        maturity, status_date = life
        run = _mdy_to_iso(bbdate)
        if not status_date <= run <= maturity:
            problems.append(f"facility {fk[1]!r} ({fk[0]}) dates its BB run {run}, outside the "
                            f"{status_date}..{maturity} life the report states for it")
    if problems:
        raise SystemExit("the export's BB dates do not sit inside the facilities the report "
                         f"describes ({len(problems)} problem(s)):\n  "
                         + "\n  ".join(problems[:25]))


def _mdy_to_iso(mdy: str) -> str:
    """The export's own M/D/YYYY date back to ISO, for comparison against the roster's dates."""
    month, day, year = (int(part) for part in mdy.split("/"))
    return date(year, month, day).isoformat()


def facility_bbdate(maturity: str | None, status_date: str | None, latest: date) -> str:
    """One facility's BB run date, as the export spells it (M/D/YYYY - formatted manually, because
    Windows strftime lacks %-m/%-d).

    A borrowing base is certified against a live facility, so the run is drawn between the date the
    agent last reported the facility's status and the earlier of its maturity and the sample's own
    as-of date. A facility the report does not state - an orphan - has no such window and falls
    back to the sample's own trailing year. Two funds on one account are two facilities running
    their own BBs, so they date separately."""
    if status_date and maturity:
        start = date.fromisoformat(status_date)
        end = min(date.fromisoformat(maturity), latest)
        span = max(0, min((end - start).days, BBDATE_MAX_LAG_DAYS))
        d = start + timedelta(days=random.randint(0, span))
    else:
        d = latest - timedelta(days=random.randint(0, BBDATE_MAX_LAG_DAYS))
    return f"{d.month}/{d.day}/{d.year}"


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
    # Region follows the type where the type implies one (a sovereign fund is not raised in Ohio),
    # and the general mix otherwise. It is an attribute of the LP, so every row this investor holds
    # carries the same value.
    region_weights = REGION_BY_TYPE.get(itype, REGION_WEIGHTS)
    region = pick_weighted(list(region_weights), list(region_weights.values()))
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
        "region": region,
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


def verify_export_against_report(export_rows: list[dict]) -> None:
    """Fail the run if the export names a facility the report does not account for.

    The two files join on the (AccountID, FndName) pair - the report's AccountNumber and Borrower,
    the export's AccountID and FndName - and the extract matches the names through _norm(), so that
    is what is compared here. An export facility neither printed by the report nor declared in
    ORPHAN_ACCOUNTS is not dirty data the ingestion recovers from: it silently becomes a placeholder
    facility with agent bank "Unknown" and an Inactive status, which is a real case the sample
    exercises DELIBERATELY through the orphans and must not acquire by accident."""
    printed = {(acct, _norm(name)) for _a, _s, rows in ABS_ROSTER for name, acct, *_ in rows}
    orphans = {(acct, _norm(fund)) for acct, fund in ORPHAN_ACCOUNTS}
    # A group-level umbrella's members are accounted for by the obligor row printed over them: the
    # report states the account and the agreement, and the funds beneath it are the export's to name.
    members = {(g.account, _norm(fund)) for g in GROUP_UMBRELLAS for fund, _share in g.members}
    unaccounted = sorted({(row["AccountID"], row["FndName"]) for row in export_rows
                          if (row["AccountID"], _norm(row["FndName"]))
                          not in printed | orphans | members})
    if unaccounted:
        listed = "\n  ".join(f"{fund!r} (account {acct})" for acct, fund in unaccounted[:25])
        raise SystemExit(f"the export places positions on {len(unaccounted)} facility(ies) the "
                         f"Agent Bank Summary does not print and ORPHAN_ACCOUNTS does not "
                         f"declare:\n  {listed}")


def weighted_sample_without_replacement(items, weights, k):
    """Efraimidis-Spirakis: key = U^(1/w); take the k largest keys."""
    keyed = sorted(((random.random() ** (1.0 / w), it) for it, w in zip(items, weights)), reverse=True)
    return [it for _, it in keyed[:k]]


def main() -> int:
    validate_chaos_vocabularies()     # before a single row is built, let alone written
    validate_roster()                 # and before the roster is rendered as either file
    random.seed(SEED)
    facilities = load_facilities()    # [(acct, fund, loan, maturity, status_date), ...]
    base = AS_OF                      # the sample's as-of date: no BB run is dated after it
    # One BBDate per facility, drawn inside the life the report states for it, so the collateral
    # date the export carries describes the facility the report prints rather than floating free of
    # it. Two funds on one account are two facilities running their own BBs, so they date
    # separately - and so do the two sleeves of one Credit Agreement, whose mirrored rows take the
    # sleeve's own date in the mirror pass below.
    fac_bbdate = {facility_key(acct, fund): facility_bbdate(maturity, status_date, base)
                  for acct, fund, _loan, maturity, status_date in facilities}

    fac_loan = {facility_key(acct, fund): loan for acct, fund, loan, *_ in facilities}

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
    per_fac: dict = {facility_key(a, f): [] for a, f, *_ in facilities}  # sleeves included
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
        r["acct"], r["fund"], r["name"], r["parent"], r["region"], r["spv"],
        r["ubs_cls"], r["itype"],
        r["inst"], r["ig"], r["cls"], r["sp_rating"], r["moodys_rating"], r["fitch_rating"],
        r["lp_size_bil"], r["lp_size_criteria"], r["commit"], r["uncalled_capital"], r["ubsar"],
        r["agent_ar"], r["agent_cl"], r["ubs_cl"], r["pct_commit"], r["called"], r["pct_of_fund_uncalled"],
        _pct8(r["called_pct"]), r["agent_excess_conc"], r["ubs_excess_conc"],
        r["agent_borrowing_base"], r["ubs_borrowing_base"], "", r["bbdate"],
    ])) for r in positions]

    # The tranche declaration, on every row of a sleeve's roster. The columns are per LP in the
    # file and per FACILITY in meaning - which is why the extract takes a facility's declaration
    # from the first of its rows that carries one rather than from the first row outright.
    #
    # Computed from the CLEAN fund name, before the chaos monkey runs. That is the point of the
    # column: the declaration is a stated fact and stays true when the name it was once inferred
    # from drifts, where a suffix reading would lose the sleeve the moment someone retyped it.
    for r, row in zip(positions, export_rows):
        sleeve = tranche_of(r["fund"])
        row["Tranche"] = sleeve or ""
        row["TrancheOf"] = tranche_parent_name(r["fund"]) if sleeve else ""

    # Prove the sample adds up while it is still clean; the chaos monkey may only degrade columns
    # this check does not depend on, or blank ones whose fallback restores the checked value.
    verify_reconciliation(export_rows)

    # Degrade what gets written, so the XLSX carries realistic manual-entry quality. Separate rng:
    # the clean base data above is identical whether chaos is on or off.
    chaos_muts: list[tuple] = []
    if CHAOS_ENABLED:
        chaos_muts = apply_chaos(export_rows, random.Random(CHAOS_SEED))

    # Prove the two files describe the same book of business before either is written. Every
    # account the export names has to be one the report prints, or a declared orphan; anything else
    # is a position the ingestion would strand on a placeholder facility nobody asked for.
    verify_export_against_report(export_rows)

    # Write the report FIRST: it is the file that states what the export's rows are positions on,
    # and writing the export against a report that then failed to render would leave the pair
    # half-updated on disk.
    abs_rows, abs_total = write_agent_bank_summary(fac_bbdate)

    # Write the export.
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    # The real file's header spellings, not the internal keys. The tranche declaration is appended
    # rather than inserted, because that is how a format gains a column without moving the ones a
    # reader already resolves by position.
    cols = SRC_COLS + (TRANCHE_SRC_COLS if EXPORT_STATES_TRANCHES else [])
    ws.append(SRC_HEADERS + (TRANCHE_SRC_HEADERS if EXPORT_STATES_TRANCHES else []))
    for row in export_rows:
        ws.append([row[c] for c in cols])
    EXPORT_OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(EXPORT_OUT)

    # The two workbooks are all this script produces; clear any stale chaos log from data/import/.
    EXPORT_OUT.with_name(EXPORT_OUT.stem + ".chaos_log.csv").unlink(missing_ok=True)

    fac_sizes = sorted(len(v) for v in per_fac.values())
    unplaced = sorted(fk[1] for fk, rows in per_fac.items() if not rows)
    print(f"wrote {ABS_OUT}")
    print(f"  facility rows printed     : {abs_rows} under {len(ABS_ROSTER)} agent bank(s)")
    print(f"  loan amount (grand total) : ${abs_total:,}")
    if unplaced:
        # Not an error: the report is the bank's record of what it lends against, and a live
        # facility the export cycle simply did not carry LPs for is an empty facility, not a
        # closed one. The extract onboards it Active on the report's word alone.
        print(f"  reported, no LPs this run : {len(unplaced)} ({', '.join(unplaced[:4])}"
              + (", ..." if len(unplaced) > 4 else "") + ")")
    print(f"wrote {EXPORT_OUT}")
    print(f"  rows (lp_records target)  : {len(positions)}")
    print(f"  distinct LPs (lp_master)  : {investor_count}")
    mirrored_sleeves = sum(len(s) for s in tranche_siblings.values())
    tranched = sum(1 for s in tranche_siblings.values() if s)
    group_members = sum(len(g.members) for g in GROUP_UMBRELLAS)
    print(f"  facilities (incl orphans) : {len(per_fac)}  ({len(ORPHAN_ACCOUNTS)} orphan, "
          f"{group_members} under a group-level umbrella, "
          f"{len(per_fac) - len(ORPHAN_ACCOUNTS) - group_members} printed by the report)")
    print(f"  tranche sleeves mirrored  : {mirrored_sleeves} on {tranched} credit agreement(s)")
    # Reported rather than declared: the umbrellas are whatever the finished export turns out to
    # hold, seeded and inherited from the report alike, counted the same way the extract counts them.
    funds_per_account: dict = {}
    for acct, fund in per_fac:
        funds_per_account.setdefault(acct, []).append(fund)
    umbrellas = {a: f for a, f in funds_per_account.items() if len(f) > 1}
    fund_count = sum(len(f) for f in umbrellas.values())
    print(f"  umbrella accounts         : {len(umbrellas)} account(s) carrying {fund_count} funds")
    for acct, funds in umbrellas.items():
        print(f"    {acct:<12}: {', '.join(sorted(funds))}")
    for group in GROUP_UMBRELLAS:
        print(f"  group-level umbrella      : {group.account} is printed ONCE as "
              f"{group.obligor!r} at ${group.loan_amount:,}; its "
              f"{len(group.members)} funds are named by the export alone")
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
