#!/usr/bin/env python3
r"""
Generate simulated Agent Borrowing Base workbooks — one per Active facility — laid out and named
the way real agent certificates are, so the Agent BB directory crawler can walk the result and mint
an import template for every file it finds.

The two simulated source workbooks in the import directory are the ONLY source of truth here; this
script reads them and never rewrites them:

    the Agent Bank summary report  states the book of business — one subdirectory is created per
                                   distinct Agent, and one workbook per Active facility under it,
                                   named from the Borrower column.
    the LP DB export               states the LP positions each of those facilities carries.

Output goes under the import directory, so nothing lands outside the data tree:

    <import>/AgentBBs/
    ├── Ashford Bank/
    │   ├── Silverbrook MM LBO VII - 2026.01.14.xlsx
    │   └── ...
    ├── Bank of Kelvin/
    └── ...

USAGE
    python lp_agent_bb_generate.py [--out DIR] [--seed N] [--agent NAME ...] [--limit N]
                                   [--no-verify] [--quiet]

Point the Agent BB directory crawler at the output root afterwards to collect the templates.

WHY THE LAYOUTS DIFFER
Every agent bank certifies its borrowing base on its own form, and a template that only ever sees
one shape proves nothing about the recognizer. Each agent is therefore assigned ONE archetype
(certificate / grouped / sleeved / feeders / deals / wide — the six shapes the sample AgentBBs tree
carries) and, on top of it, a per-agent DIALECT: which header spelling it prints for each column,
which optional columns it carries at all, where the grid starts, how it writes money and percents,
how it names its sheets and its files. One template per bank is enough because every workbook that
bank writes is the same form; no two banks write the same one.

The randomization is bounded by what the parser needs, not by taste — every generated sheet is
written to clear the recognizer's own thresholds (a header row carrying enough dictionary aliases,
and data rows dense enough to read as an investor grid), and --verify re-runs that recognizer over
each finished workbook before the run is called a success.

WHY THE LP ROWS ARE ALMOST — BUT NOT QUITE — THE EXPORT'S
CARRY_RATE of each facility's rows are the SAME investors the LP DB Export carries for that
facility, at the same capital commitments and the same called/uncalled split: that overlap is what
gives the ingestion's duplicate matching something to match. The rest are investors that exist only
on the agent's file, which is what gives it something to fail to match.

WHY AN UMBRELLA IS ONE FILE AND NOT SIX
Several related funds, feeders and SPVs draw on ONE credit agreement, which is why the export books
them all against a single account number. The agent certifies the AGREEMENT, not each fund, so an
account the export shares between UMBRELLA_MEMBERS borrowers produces ONE certificate carrying every
member's roster — and whatever the report printed over that account, one row for the group or one
row per member, is replaced by that single file. The group is named for what its members' names have
in common, marked as the umbrella it is; it is never named for the account number, which names
nothing a reader of either file could look up.

A carried row is never a copy. The agent states its OWN credit view of the LP, so its LP category,
its agency ratings, its investor type and therefore its advance rate, concentration limit and
borrowing base all differ from the export's — while staying inside the canonical reference vocabularies for
agent rate schedules, investor types and agency rating scales. Differing is asserted per row, not
hoped for: the run fails if a carried row comes out identical, and fails if a stated borrowing base
does not follow from the row's own stated rate and limit.

Usage:
    python lp_agent_bb_generate.py [--out DIR] [--seed N] [--agent NAME ...] [--limit N] [--no-verify]

The same --seed reproduces the same tree, byte-comparable layouts included.
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import logging
import random
import re
import sys
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path
from typing import Callable, Optional

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# The analyzer this generator writes for — the same recognizer the import pipeline runs.
try:
    from parse_excel_templates import ExcelAnalyzer, load_dictionary
except ImportError as e:  # pragma: no cover - environment problem, not a data problem
    sys.exit(f"Failed to import the Excel template analyzer: {e}\n"
             f"Ensure this script sits alongside it.")

SCRIPT_DIR = Path(__file__).resolve().parent
DATA_DIR = SCRIPT_DIR.parent / "data"
REFERENCE_DIR = DATA_DIR / "reference"                 # the canonical vocabularies the extract uses
IMPORT_DIR = DATA_DIR / "import"
ABS_IN = IMPORT_DIR / "AgentBankSummaryRpt.xlsx"
EXPORT_IN = IMPORT_DIR / "LP DB Export V3.xlsx"
DEFAULT_OUT = IMPORT_DIR / "AgentBBs"
ABS_SHEET_NAME = "Agent Bank Summary"
EXPORT_SHEET_NAME = "BBs"

logger = logging.getLogger("lp_agent_bb_generate")

# ── tunables ────────────────────────────────────────────────────────────────────────────────────
SEED = 20260908
# Share of a facility's rows the agent's file and the LP DB Export agree on by INVESTOR. The
# remainder are investors only the agent carries: an LP the bank has not onboarded yet, which is the
# row the match queue has to raise rather than auto-accept. Row COUNT is held at the export's, so
# raising the carry rate trades unmatched rows for matched ones and nothing else.
CARRY_RATE = 0.90
ONLY_ACTIVE = True              # the report prints a FacilityStatus; only Active facilities certify
MONEY_STEP = 100_000            # the export's own $100K grid, so a minted position sits on it too
MIN_UNCALLED = 200_000
MIN_ROWS_PER_SHEET = 4          # a sheet below this cannot clear the analyzer's data-row threshold
MAX_SHEETS = 5
# An agent prices its own schedule, not the bank's: the reference rate/limit for a category is the
# starting point, and the credit agreement moves it. Drawn once per facility, so the whole
# certificate prices consistently.
CL_FACTORS = (1.0, 1.0, 1.0, 0.75, 0.5, 1.25)
RATE_SHIFTS = (0, 0, 0, 0, 0, 0, 1, -1)   # rungs down/up AGENT_RATE_LADDER
# An uncommitted sleeve prices one rung down and is tested against half the committed sleeve's cap.
# The two sleeves of one credit agreement must keep diverging off identical collateral on the
# agent's file exactly as they do in the export.
AGENT_RATE_LADDER = (0.90, 0.75, 0.60, 0.50, 0.25, 0.00)
UNCOMMITTED_CL_FACTOR = 0.5
TRANCHE_SUFFIX_RE = re.compile(r"\s*\((committed|uncommitted)\)\s*$", re.I)

# An umbrella subscription facility is an account number the export shares between this many
# borrowers. Two funds on one account are as often a feeder beside its master as a group, so the
# floor sits where the account stops reading as a pair and starts reading as a book: below it the
# borrowers keep their own certificates.
UMBRELLA_MEMBERS = (4, 6)
# The group has to say it is one, in the two spellings the report itself uses. The name is asserted
# to carry one of them before the workbook is written.
UMBRELLA_MARKERS = ("Umbrella", "[U]")
UMBRELLA_SUFFIX = "Umbrella"
UMBRELLA_MARKER_RE = re.compile(r"\s*(Umbrella|\[U\])\s*$", re.I)

# Windows reserves these in a path segment; a Borrower name carrying one still has to produce a file.
ILLEGAL_PATH_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
EM_DASH = "—"


# ================================================================================================
#  Reference vocabularies — the same ones the extract normalizes against
# ================================================================================================
def _reference_rows(name: str) -> list[list[str]]:
    """Data rows of a reference CSV, '#'-comment and blank lines dropped, header row included."""
    path = REFERENCE_DIR / name
    with path.open(newline="", encoding="utf-8") as fh:
        return [r for r in csv.reader(fh)
                if r and r[0].strip() and not r[0].lstrip().startswith("#")]


def load_agent_categories() -> dict[str, tuple[float, float]]:
    """Agent LP Category -> (default advance rate, default concentration limit), as fractions.

    The rate map states the canonical category vocabulary as well as the two defaults, so a category
    this generator writes is one the extract can resolve and one the platform's agent tier config
    already carries."""
    rows = _reference_rows("agent_rate_map.csv")[1:]
    return {r[0]: (float(r[1]) / 100.0, float(r[2]) / 100.0) for r in rows}


def load_investor_types() -> list[str]:
    """The canonical Investor Type vocabulary the platform offers."""
    return [r[0] for r in _reference_rows("investor_types.csv")[1:] if r[0]]


def load_rating_scales() -> dict[str, dict[str, list[str]]]:
    """{agency: {band: [notch, ...]}} from the agency rating scales.

    The three agencies do NOT share a scale — writing an S&P notch into the Moody's column bands to
    nothing — so every notch this generator emits is drawn from its own column's agency."""
    scales: dict[str, dict[str, list[str]]] = {}
    for agency, band, notch in (r[:3] for r in _reference_rows("rating_scales.csv")[1:]):
        scales.setdefault(agency, {}).setdefault(band or "SUB_IG", []).append(notch)
    return scales


AGENT_TERMS = load_agent_categories()
AGENT_CATEGORIES = list(AGENT_TERMS)
# Category mix on an agent's own file, weighted the way the export's population is weighted.
AGENT_CATEGORY_WEIGHTS = {
    "Included Investors (Rated)": 40,
    "Included Investors (Non-Rated)": 28,
    "Insitutional Designated Investors": 16,
    "PWM Designated Investors": 9,
    "Excluded Investors": 7,
}
RATED_CATEGORY = "Included Investors (Rated)"
PWM_CATEGORY = "PWM Designated Investors"
EXCLUDED_CATEGORY = "Excluded Investors"

INVESTOR_TYPES = load_investor_types()
HNW_TYPES = ("HNW", "Family Office")
INSTITUTIONAL_TYPES = tuple(t for t in INVESTOR_TYPES if t != "HNW")
RATING_SCALES = load_rating_scales()
RATED_BAND_WEIGHTS = {"AAA": 6, "AA": 24, "A": 40, "BBB": 25, "SUB_IG": 5}
AGENCY_COVERAGE_WEIGHTS = (18, 27, 55)   # P(1 agency rates it), P(2), P(3)

# Vocabulary for the investors only the agent's file carries. Same shape as the export's minted
# names so an unmatched row reads as an LP the bank has not onboarded, not as a different species.
NEW_LP_WORD1 = (
    "Ashcombe", "Blackthorn", "Carrowmore", "Dunhaven", "Elmsworth", "Fernbank", "Glenavon",
    "Hartsfield", "Inverleith", "Jarrow", "Kingsmere", "Lyndhurst", "Marchmont", "Netherby",
    "Ormesby", "Penrhyn", "Quarrendon", "Ravensworth", "Stanmore", "Thurlow", "Ullapool",
    "Vinehall", "Wentworth", "Yarrowdale", "Amberley", "Bridewell", "Calderstone", "Duddingston",
)
NEW_LP_WORD2 = ("", "", "Capital", "Global", "Strategic", "Nominees", "Trust", "Partners")
NEW_LP_SUFFIX = {
    "Pension Fund": ("Pension Fund", "Retirement Trust", "Pension Scheme"),
    "Public Pension": ("Public Employees Retirement System", "State Pension Fund"),
    "Endowment": ("University Endowment", "Endowment Fund"),
    "Foundation": ("Foundation", "Charitable Foundation"),
    "Family Office": ("Family Office", "Family Holdings"),
    "Fund of Funds": ("Fund of Funds", "Multi-Manager Fund"),
    "Sovereign Wealth Fund": ("Investment Authority", "Sovereign Fund"),
    "Insurance Company": ("Life Insurance Co.", "Mutual Insurance"),
    "Healthcare": ("Health System", "Hospital Trust"),
    "Corporate": ("Group Treasury", "Corporate Holdings"),
    "Investment Consultant": ("Capital Advisors", "Investment Advisors"),
    "Institutional Investor": ("Institutional Trust", "Capital Partners"),
    "Hedge Fund": ("Master Fund", "Opportunities Fund"),
    "Other Institutional": ("Alternative Assets", "Global Investors"),
    "HNW": ("Private Investment Office", "Private Wealth Nominees"),
}


# ================================================================================================
#  Money / percent helpers — full stored precision, never a rounded or abbreviated dollar
# ================================================================================================
def _money(x: float) -> float:
    """A dollar figure quantized to the cent and no further."""
    return float(Decimal(repr(float(x))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def _pct8(x: float) -> float:
    """A share (0..1) kept to 8 decimals, so a small position in a large facility stays legible."""
    return float(Decimal(repr(float(x))).quantize(Decimal("0.00000001"), rounding=ROUND_HALF_UP))


def snap_money(x: float, floor: int) -> int:
    """A dollar figure on the export's $100K grid, never below `floor`."""
    return max(floor, int(round(x / MONEY_STEP)) * MONEY_STEP)


def money_short(bil: Optional[float]) -> str:
    """An LP-size figure the way an agent prints it in a Net Worth / AUM column: "$46.2B", "$478.0M".

    This is a DISPLAY column on the agent's own form — the figure it abbreviates is the export's
    LP Size ($ Bil), which is where the precise value lives."""
    if bil is None:
        return ""
    return f"${bil:.1f}B" if bil >= 1 else f"${bil * 1000:.1f}M"


def money_range(bil: Optional[float]) -> str:
    """The banded form of the same figure, for the agents whose form asks for a range."""
    if bil is None:
        return ""
    for lo, hi in ((0, 0.25), (0.25, 1), (1, 5), (5, 10), (10, 50), (50, 100)):
        if bil < hi:
            return f"{money_short(lo) if lo else '$0'} - {money_short(hi)}"
    return f"> {money_short(100)}"


def _seed_of(*parts) -> int:
    """A stable seed from the run seed and any key — hash() is salted per process and would make
    the tree unreproducible between runs."""
    raw = "␟".join(str(p) for p in parts).encode("utf-8")
    return int.from_bytes(hashlib.sha256(raw).digest()[:8], "big")


# ================================================================================================
#  Source models
# ================================================================================================
@dataclass
class AgentLp:
    """One LP row as the AGENT states it. Money is the export's; the credit view is the agent's."""
    name: str
    parent: str
    region: str
    spv: str
    itype: str                  # the agent's Investor Type — differs from the export's
    cls: str                    # Agent LP Category
    sp: str
    moodys: str
    fitch: str
    size_bil: Optional[float]
    size_criteria: str
    commitment: float
    called: float
    uncalled: float
    rate: float = 0.0           # agent advance rate, fraction
    conc_limit: float = 0.0     # agent concentration limit, fraction of the facility's uncalled
    eligible: float = 0.0
    excess: float = 0.0
    bb: float = 0.0
    recallable: float = 0.0     # the feeders form reports it instead of called capital
    carried: bool = True        # False = an investor only the agent's file carries
    member: str = ""            # on an umbrella, the borrower under the agreement this row draws on
    sleeve: str = ""
    ga_id: str = ""
    transferred: str = ""
    pct_commit: float = 0.0
    pct_uncalled: float = 0.0
    pct_called: float = 0.0
    pct_bb: float = 0.0
    source: dict = field(default_factory=dict)   # the export row this one was carried from

    @property
    def rated(self) -> bool:
        return any(r not in ("", "NR") for r in (self.sp, self.moodys, self.fitch))


@dataclass
class Facility:
    """One Active facility the summary report prints, with the LP roster its certificate carries."""
    agent: str
    borrower: str
    account: str
    loan: float
    maturity: Optional[date]
    status: str
    status_date: Optional[date]
    bb_date: date
    members: tuple[str, ...] = ()   # the borrowers under one agreement; empty on a single facility
    lps: list[AgentLp] = field(default_factory=list)
    tot_commit: float = 0.0
    tot_called: float = 0.0
    tot_uncalled: float = 0.0
    tot_eligible: float = 0.0
    tot_excess: float = 0.0
    tot_bb: float = 0.0

    @property
    def tranche(self) -> Optional[str]:
        m = TRANCHE_SUFFIX_RE.search(self.borrower)
        return m.group(1).title() if m else None

    @property
    def is_umbrella(self) -> bool:
        return bool(self.members)


def _as_date(v) -> Optional[date]:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str) and v.strip():
        for fmt in ("%m/%d/%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(v.strip(), fmt).date()
            except ValueError:
                continue
    return None


def load_report_facilities(path: Path) -> list[Facility]:
    """Read the Agent Bank Summary report.

    The report is a printed, GROUPED report, not a table: the Agent sits alone on its own band row
    and every Borrower row beneath it belongs to that agent until the next one. The totals bands
    ("AccessTotalsLoanAmount:") carry no borrower and are skipped."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[ABS_SHEET_NAME]
    facilities: list[Facility] = []
    agent = None
    for r in range(2, ws.max_row + 1):
        agent_cell = ws.cell(r, 1).value
        borrower = ws.cell(r, 2).value
        if agent_cell and not borrower:
            agent = str(agent_cell).strip()
            continue
        if not borrower or agent is None:
            continue
        status = str(ws.cell(r, 7).value or "").strip()
        if ONLY_ACTIVE and status.lower() != "active":
            continue
        facilities.append(Facility(
            agent=agent,
            borrower=str(borrower).strip(),
            account=str(ws.cell(r, 3).value or "").strip(),
            loan=float(ws.cell(r, 4).value or 0),
            maturity=_as_date(ws.cell(r, 6).value),
            status=status,
            status_date=_as_date(ws.cell(r, 8).value),
            bb_date=date.today(),   # replaced from the export's BBDate below
        ))
    wb.close()
    return facilities


def load_export_rows(path: Path) -> dict[tuple[str, str], list[dict]]:
    """LP DB Export rows grouped by (AccountID, FndName), in FILE ORDER.

    File order is the ordering key the whole ingestion carries (row_index -> source_seq), so the
    agent's certificate presents its carried rows in the order the export states them rather than
    re-sorting them into something the wizard would then have to undo."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[EXPORT_SHEET_NAME]
    stream = ws.iter_rows(values_only=True)
    headers = [str(h).strip() if h is not None else "" for h in next(stream)]
    grouped: dict[tuple[str, str], list[dict]] = {}
    for values in stream:
        row = dict(zip(headers, values))
        acct = str(row.get("AccountID") or "").strip()
        fund = str(row.get("FndName") or "").strip()
        if not acct or not fund:
            continue
        grouped.setdefault((acct, fund), []).append(row)
    wb.close()
    return grouped


# ================================================================================================
#  Umbrellas — one account, several borrowers, one certificate
# ================================================================================================
def umbrella_members(export: dict[tuple[str, str], list[dict]]) -> dict[str, tuple[str, ...]]:
    """Account number -> the borrowers the export shares it between, for the umbrella accounts only.

    Which accounts those are is read from the EXPORT rather than from the report, because the report
    is free to print an umbrella either way — one row for the whole agreement or one row per member —
    and only the export always states every fund that draws on it. Members keep export order."""
    by_account: dict[str, list[str]] = {}
    for acct, fund in export:
        by_account.setdefault(acct, []).append(fund)
    lo, hi = UMBRELLA_MEMBERS
    return {a: tuple(m) for a, m in by_account.items() if lo <= len(m) <= hi}


def common_name(names: list[str]) -> str:
    """The leading words every name shares, in the first one's own casing.

    Members of one umbrella are the sponsor's own family — a fund, its feeder, its offshore sleeve
    and its SPV — so what their names have in common is the agreement's name, and what follows it is
    what separates one member from the next. A trailing legal token is dropped: "Emberly Fund LP" and
    "Emberly Fund Offshore" share "Emberly Fund", not "Emberly Fund LP"."""
    split = [[t for t in name.replace(",", " ").split() if t] for name in names if name.strip()]
    if not split:
        return ""
    stem: list[str] = []
    for i in range(min(len(tokens) for tokens in split)):
        token = split[0][i]
        if any(tokens[i].casefold() != token.casefold() for tokens in split[1:]):
            break
        stem.append(token)
    while stem and stem[-1].replace(".", "").lower() in LEGAL_TOKENS:
        stem.pop()
    return " ".join(stem)


def umbrella_name(members: tuple[str, ...], printed: str = "") -> str:
    """The borrower name the group's certificate carries.

    It is the name common to the members, marked as the umbrella — the one name that is true of
    every fund on the account and that a reader can find on both files. The account number is never
    it: an account number names nothing, and the platform keys the group on it already, so a name
    that only restates it leaves the group unnamed. Where the members share no name at all the
    printed row's name is used, and failing that the first member's, so there is always one.

    Where the summary report ALREADY printed that name — the members' shared stem, marked — its
    exact spelling is kept, "[U]" as readily as "Umbrella". The report and the certificate then name
    the agreement identically, which is what a reader matching the two files by hand goes on, and
    re-spelling it here would put a second name on one credit agreement for no gain."""
    stem = common_name(list(members))
    if UMBRELLA_MARKER_RE.search(printed or ""):
        printed_stem = UMBRELLA_MARKER_RE.sub("", printed).strip()
        if not stem or printed_stem.casefold() == stem.casefold():
            return printed.strip()
    if not stem:
        stem = UMBRELLA_MARKER_RE.sub("", printed).strip() or " ".join(members[0].split()[:2])
    if any(marker.casefold() in stem.casefold() for marker in UMBRELLA_MARKERS):
        return stem
    return f"{stem} {UMBRELLA_SUFFIX}"


def member_label(member: str, fac: Facility) -> str:
    """What separates one member from its siblings — its name with the group's shared stem removed.

    This is what an agent writes on a per-fund tab of an umbrella's certificate: "Offshore X",
    not the full name repeated on every sheet, which would truncate to the same 31 characters."""
    stem = common_name(list(fac.members))
    if stem and member.casefold().startswith(stem.casefold()):
        return member[len(stem):].strip(" ,-") or member
    return member


def fold_umbrellas(facilities: list[Facility],
                   members_by_account: dict[str, tuple[str, ...]]) -> list[Facility]:
    """Replace every printed row on an umbrella account with ONE facility for the group.

    The agent certifies the credit agreement once, so one file covers it however the report printed
    it. The group takes the first printed row's dates and standing — they are the agreement's, and
    every row on the account restates them — and the whole agreement's line: stated outright where
    the report printed the group, and the sum of the members' shares where it printed them one by
    one."""
    folded: list[Facility] = []
    seen: set[str] = set()
    for fac in facilities:
        members = members_by_account.get(fac.account)
        if members is None:
            folded.append(fac)
            continue
        if fac.account in seen:
            continue                    # a further row on the account: the group already carries it
        seen.add(fac.account)
        printed = [f for f in facilities if f.account == fac.account]
        folded.append(Facility(
            agent=fac.agent,
            borrower=umbrella_name(members, fac.borrower),
            account=fac.account,
            loan=fac.loan if len(printed) == 1 else _money(sum(f.loan for f in printed)),
            maturity=fac.maturity,
            status=fac.status,
            status_date=fac.status_date,
            bb_date=fac.bb_date,
            members=members,
        ))
    return folded


# ================================================================================================
#  The agent's own credit view of a carried LP
# ================================================================================================
def _pick(rng: random.Random, options, weights=None):
    return rng.choices(list(options), weights=weights, k=1)[0]


def draw_ratings(rng: random.Random, avoid: tuple[str, str, str] = ("", "", "")) -> tuple[str, str, str]:
    """Agency notches for a rated LP: 1-3 agencies cover it, each from its OWN scale, all inside one
    band so the three columns describe one credit rather than three. `avoid` is the export's triple —
    the draw is repeated until at least one column disagrees with it, which is what makes the agent's
    rating a genuinely independent opinion rather than a copy."""
    for _ in range(12):
        band = _pick(rng, RATED_BAND_WEIGHTS, list(RATED_BAND_WEIGHTS.values()))
        covered = rng.choices((1, 2, 3), weights=AGENCY_COVERAGE_WEIGHTS, k=1)[0]
        agencies = rng.sample(("sp", "moodys", "fitch"), covered)
        drawn = {a: _pick(rng, RATING_SCALES[a][band]) for a in agencies}
        triple = (drawn.get("sp", "NR"), drawn.get("moodys", "NR"), drawn.get("fitch", "NR"))
        if triple != avoid:
            return triple
    return triple   # pragma: no cover - 12 draws colliding is not reachable in practice


def agent_view(rng: random.Random, src: dict) -> AgentLp:
    """Turn one LP DB Export row into the row the AGENT states for the same investor.

    Same investor, same money — the agent and the bank are looking at one position, and the shared
    name and commitment are what the match queue matches on. Everything that is a CREDIT OPINION is
    re-formed here and asserted different downstream:

      * the Agent LP Category is re-drawn from the canonical five, never the export's;
      * ratings follow the category rather than decorating it — the Rated tier IS the rated
        population, so a re-drawn category that is not Rated withdraws the ratings and one that is
        supplies a fresh set;
      * Investor Type follows the same logic (the private-wealth tier is HNW/Family Office; every
        other tier is institutional) and is likewise never the export's value.
    """
    src_cls = str(src.get("Agent LP Classification") or "").strip()
    src_type = str(src.get("Investor Type") or "").strip()
    src_ratings = (str(src.get("S&P") or "NR").strip() or "NR",
                   str(src.get("Moody'S") or "NR").strip() or "NR",
                   str(src.get("Fitch") or "NR").strip() or "NR")

    choices = [c for c in AGENT_CATEGORIES if c != src_cls] or AGENT_CATEGORIES
    cls = _pick(rng, choices, [AGENT_CATEGORY_WEIGHTS[c] for c in choices])

    if cls == RATED_CATEGORY:
        sp, moodys, fitch = draw_ratings(rng, src_ratings)
    else:
        sp = moodys = fitch = "NR"

    type_pool = HNW_TYPES if cls == PWM_CATEGORY else INSTITUTIONAL_TYPES
    type_choices = [t for t in type_pool if t != src_type] or list(type_pool)
    itype = _pick(rng, type_choices)

    commitment = float(src.get("Capital Commitments") or 0)
    uncalled = float(src.get("Uncalled Capital") or 0)
    called = float(src.get("Called Capital") or max(0.0, commitment - uncalled))
    size_bil = src.get("LP Size\n\n($ Bil)")
    return AgentLp(
        name=str(src.get("Investor Name") or "").strip(),
        parent=str(src.get("Parent") or "").strip(),
        region=str(src.get("Region") or "").strip(),
        spv=str(src.get("SPV") or "N").strip() or "N",
        itype=itype,
        cls=cls,
        sp=sp, moodys=moodys, fitch=fitch,
        size_bil=float(size_bil) if isinstance(size_bil, (int, float)) else None,
        size_criteria=str(src.get("LP Size Criteria") or "").strip(),
        commitment=_money(commitment),
        called=_money(called),
        uncalled=_money(uncalled),
        carried=True,
        source=src,
    )


def mint_lp(rng: random.Random, used: set[str], commit_scale: float) -> AgentLp:
    """An investor only the agent's file carries — no export row to match it to."""
    cls = _pick(rng, AGENT_CATEGORIES, [AGENT_CATEGORY_WEIGHTS[c] for c in AGENT_CATEGORIES])
    itype = _pick(rng, HNW_TYPES if cls == PWM_CATEGORY else INSTITUTIONAL_TYPES)
    for _ in range(64):
        w2 = _pick(rng, NEW_LP_WORD2)
        name = (f"{_pick(rng, NEW_LP_WORD1)}{(' ' + w2) if w2 else ''} "
                f"{_pick(rng, NEW_LP_SUFFIX[itype])}")
        if name not in used:
            break
    used.add(name)
    sp, moodys, fitch = draw_ratings(rng) if cls == RATED_CATEGORY else ("NR", "NR", "NR")

    uncalled = snap_money(commit_scale * rng.lognormvariate(0.0, 0.9) * rng.uniform(0.15, 0.85),
                          MIN_UNCALLED)
    commitment = snap_money(uncalled / rng.uniform(0.15, 0.85), uncalled + MONEY_STEP)
    measure = _pick(rng, ("AUM", "AUM", "NAV", "Assets"))
    return AgentLp(
        name=name,
        parent="",
        region=_pick(rng, ("North America", "Europe", "Asia-Pacific", "Middle East", "Other"),
                     (46, 24, 16, 9, 5)),
        spv="N",
        itype=itype,
        cls=cls,
        sp=sp, moodys=moodys, fitch=fitch,
        size_bil=round(10 ** rng.uniform(-0.5, 1.8), 1),
        size_criteria=measure,
        commitment=_money(commitment),
        called=_money(commitment - uncalled),
        uncalled=_money(uncalled),
        carried=False,
    )


def price_facility(fac: Facility, rng: random.Random) -> None:
    """Resolve every rate, limit, eligible amount, excess and borrowing base on the certificate.

    This is the platform's own borrowing base arithmetic, agent side, and it is re-run here rather
    than copied from the export because the agent's categories are not the export's:

      * a concentration limit is a fraction of the facility's TOTAL uncalled capital — Excluded LPs
        included, they consume the denominator without contributing base;
      * eligible uncalled is the LP's own uncalled capped at that limit, the excess is the rest;
      * the base advances only on the eligible portion;
      * an Excluded LP books no base and no excess — there is no position to have an excess over.

    The facility's own credit agreement moves the schedule (CL_FACTORS/RATE_SHIFTS), and an
    uncommitted sleeve prices one rung down against half the cap, so the two sleeves of one credit
    agreement diverge off identical collateral here exactly as they do in the export.
    """
    cl_factor = _pick(rng, CL_FACTORS)
    rate_shift = _pick(rng, RATE_SHIFTS)
    uncommitted = fac.tranche == "Uncommitted"
    last_rung = len(AGENT_RATE_LADDER) - 1

    # The facility totals are pure money and owe nothing to the credit view, so they resolve first —
    # a concentration limit is a fraction of them, and eligibility has to be known before a rate can
    # be checked against the base it would produce.
    fac.tot_commit = _money(sum(lp.commitment for lp in fac.lps))
    fac.tot_called = _money(sum(lp.called for lp in fac.lps))
    fac.tot_uncalled = _money(sum(lp.uncalled for lp in fac.lps))
    tot_u = fac.tot_uncalled or 1.0
    tot_c = fac.tot_commit or 1.0

    for lp in fac.lps:
        excluded = lp.cls == EXCLUDED_CATEGORY
        base_rate, base_cl = AGENT_TERMS[lp.cls]
        lp.conc_limit = round(base_cl * cl_factor * (UNCOMMITTED_CL_FACTOR if uncommitted else 1.0), 6)

        eligible = lp.uncalled if lp.conc_limit <= 0 else min(lp.uncalled, lp.conc_limit * tot_u)
        lp.excess = _money(0.0 if excluded else lp.uncalled - eligible)
        lp.eligible = _money(0.0 if excluded else lp.uncalled - lp.excess)

        rung = AGENT_RATE_LADDER.index(base_rate) if base_rate in AGENT_RATE_LADDER else 0
        rung = min(max(rung + rate_shift + (1 if uncommitted else 0), 0), last_rung)
        if not excluded:
            # Divergence guard. The agent's category for this LP is never the export's, but two
            # categories can still price to the same rung once the facility's shift is applied, and
            # an identical rate over identical eligible capital restates the export's borrowing base
            # exactly — the one thing a duplicate-matching fixture must not do. The test is the base
            # itself rather than the rate, because the export leaves the rate null on some rows and
            # a null there would let a collision through. Stepping stays off the ladder's 0.00 rung:
            # that rung means Excluded, and this LP is not.
            src_bb = float(lp.source.get("Agent Borrowing Base") or 0) if lp.carried else 0.0
            if src_bb > 0 and abs(_money(lp.eligible * AGENT_RATE_LADDER[rung]) - src_bb) <= 0.01:
                rung = rung + 1 if rung < last_rung - 1 else rung - 1
        lp.rate = 0.0 if excluded else AGENT_RATE_LADDER[rung]
        lp.bb = _money(0.0 if excluded else lp.eligible * lp.rate)
        lp.pct_commit = _pct8(lp.commitment / tot_c)
        lp.pct_uncalled = _pct8(lp.uncalled / tot_u)
        lp.pct_called = _pct8(lp.called / lp.commitment) if lp.commitment else 0.0
        # Recallable distributions are a feeder-form column, not an export one: the portion of
        # called capital the partnership may draw again, which that form adds back to callable.
        lp.recallable = _money(lp.called * rng.uniform(0.04, 0.11))

    fac.tot_eligible = _money(sum(lp.eligible for lp in fac.lps))
    fac.tot_excess = _money(sum(lp.excess for lp in fac.lps))
    fac.tot_bb = _money(sum(lp.bb for lp in fac.lps))
    tot_bb = fac.tot_bb or 1.0
    for lp in fac.lps:
        lp.pct_bb = _pct8(lp.bb / tot_bb)


def _carry_block(rows: list[dict], rng: random.Random) -> tuple[list[AgentLp], int]:
    """(the rows this block carries, how many it drops). A sample WITHOUT replacement, so a dropped
    investor is dropped once and the survivors keep their export file order."""
    total = len(rows)
    keep_n = max(1, min(total, round(total * CARRY_RATE)))
    return [agent_view(rng, rows[i]) for i in sorted(rng.sample(range(total), keep_n))], total - keep_n


def stamp_members(fac: Facility) -> None:
    """Name the borrower every row of an umbrella's certificate draws under.

    A carried row states the fund the export booked it against; a minted one inherits the fund of the
    row above it, so an investor the agent alone carries lands INSIDE a member's block rather than
    splitting it, and the certificate keeps one contiguous block per borrower."""
    if not fac.is_umbrella:
        return
    current = fac.members[0]
    for lp in fac.lps:
        if lp.carried:
            current = str(lp.source.get("FndName") or "").strip() or current
        lp.member = current


def build_roster(fac: Facility, export_rows: list[dict], rng: random.Random,
                 minted_names: set[str]) -> None:
    """Fill the facility's LP roster: CARRY_RATE of the export's rows for it, the rest minted.

    Row COUNT follows the export so the certificate is the same size as the position file. An
    umbrella carries at that rate per MEMBER rather than over the group, so no fund on the agreement
    can be sampled away entirely — a certificate silently missing one of its borrowers is not a
    smaller fixture, it is a different structure."""
    if export_rows:
        blocks = ([[r for r in export_rows if str(r.get("FndName") or "").strip() == m]
                   for m in fac.members] if fac.is_umbrella else [export_rows])
        carried, new_n = [], 0
        for block in blocks:
            if not block:
                continue
            kept, dropped = _carry_block(block, rng)
            carried += kept
            new_n += dropped
        commit_scale = sum(lp.uncalled for lp in carried) / max(1, len(carried))
    else:
        # A facility the report prints but the export cycle carried no positions for. The agent
        # still certifies it; every row on it is new to the bank.
        carried = []
        commit_scale = max(MIN_UNCALLED, fac.loan / rng.uniform(8, 25))
        new_n = rng.randint(12, 40)

    used = {lp.name for lp in carried} | minted_names
    minted = [mint_lp(rng, used, commit_scale) for _ in range(new_n)]
    minted_names.update(lp.name for lp in minted)

    # New investors are interleaved, not appended: an agent's file lists whoever subscribed, and a
    # block of unmatched rows at the bottom would let the match queue be exercised by position.
    roster = list(carried)
    for lp in minted:
        roster.insert(rng.randint(0, len(roster)), lp)
    fac.lps = roster
    stamp_members(fac)


# ================================================================================================
#  Layout engine
# ================================================================================================
# One column of an agent's grid: the header spellings a bank may print for it, and how the value is
# rendered. Every alias listed under a key that a layout counts toward its matched-column budget
# resolves through the recognizer's field dictionary — decorations like " (USD)" survive that
# dictionary's substring matching, which is why they are safe variation rather than noise.
@dataclass(frozen=True)
class Column:
    key: str
    aliases: tuple[str, ...]
    kind: str                       # text | money | pct | rate | flag | size
    value: Callable[[AgentLp, Facility], object]


COLUMNS: dict[str, Column] = {c.key: c for c in (
    Column("investor", ("Investor Name", "Investor", "LP Name", "Limited Partner",
                        "Investor Name (Agent Records)", "Deal Investor Name"),
           "text", lambda lp, f: lp.name),
    Column("parent", ("Parent / Sponsor", "Parent", "Sponsor", "Parent / Sponsor / Manager"),
           "text", lambda lp, f: lp.parent),
    Column("sleeve", ("Fund Sleeve", "Feeder Fund", "Fund Vehicle", "Feeder"),
           "text", lambda lp, f: lp.sleeve),
    Column("region", ("Region", "Domicile", "Region / Location", "Location"),
           "text", lambda lp, f: lp.region),
    Column("spv", ("SPV", "SPV Flag", "Special Purpose Vehicle"),
           "flag", lambda lp, f: lp.spv),
    Column("itype", ("Investor Type", "Investor Classification", "Category of Investor"),
           "text", lambda lp, f: lp.itype),
    Column("lp_category", ("LP Category", "Agent LP Classification", "LP Classification"),
           "text", lambda lp, f: lp.cls),
    Column("incl_excl", ("Included/Excluded Investor", "Included / Excluded", "Eligibility"),
           "flag", lambda lp, f: "Excluded" if lp.cls == EXCLUDED_CATEGORY else "Included"),
    Column("excluded_flag", ("Excluded", "Excluded?"),
           "flag", lambda lp, f: "Y" if lp.cls == EXCLUDED_CATEGORY else "N"),
    Column("defaulting", ("Defaulting?", "Defaulting Investor?"),
           "flag", lambda lp, f: "N"),
    Column("rights", ("Claimed/Exercised Rights?", "Claimed / Exercised Rights?"),
           "flag", lambda lp, f: "N"),
    Column("transferred", ("Transferred From", "Transferor"),
           "text", lambda lp, f: lp.transferred),
    Column("partnership", ("Borrowing Partnership", "Borrower Entity"),
           "text", lambda lp, f: lp.sleeve or f.borrower),
    Column("ga_id", ("GA ID", "Investor ID"),
           "text", lambda lp, f: lp.ga_id),
    Column("sp", ("S&P", "S&P Rating", "Investor S&P"),
           "text", lambda lp, f: lp.sp),
    Column("moodys", ("Moody's", "Moody's Rating", "Investor Moody's"),
           "text", lambda lp, f: lp.moodys),
    Column("fitch", ("Fitch", "Fitch Rating"),
           "text", lambda lp, f: lp.fitch),
    Column("size", ("AUM", "Assets Under Management", "NAV", "Net Asset Value", "Pension Assets",
                    "Net Worth", "Net Assets (range)", "NAV Range (USD)"),
           "size", lambda lp, f: lp.size_bil),
    Column("commitment", ("Capital Commitments", "Committed Capital", "Total Commitment",
                          "Capital Commitments (USD)", "Commitment (USD)"),
           "money", lambda lp, f: lp.commitment),
    Column("commitment_ind", ("Individual Capital Commitment", "Individual Original Commitment"),
           "money", lambda lp, f: lp.commitment),
    Column("called", ("Called Capital", "Drawn Capital", "Funded Capital", "Called Capital (USD)"),
           "money", lambda lp, f: lp.called),
    Column("recallable", ("Recallable Distribution", "Recallable Distributions"),
           "money", lambda lp, f: lp.recallable),
    Column("uncalled", ("Uncalled Capital", "Unfunded Commitment", "Remaining Commitment",
                        "Uncalled Capital (USD)", "Unfunded Commitment (USD)"),
           "money", lambda lp, f: lp.uncalled),
    Column("uncalled_ind", ("Individual Unfunded Capital Commitment", "Remaining Callable Capital"),
           "money", lambda lp, f: lp.uncalled),
    Column("pct_commit", ("% of Capital Commitments", "% Commitment"),
           "pct", lambda lp, f: lp.pct_commit),
    Column("pct_uncalled", ("% of Uncalled Capital", "% Uncalled", "Uncalled %",
                            "% Total Unfunded Commitment"),
           "pct", lambda lp, f: lp.pct_uncalled),
    Column("pct_called", ("% Called", "% Called (Net)"),
           "pct", lambda lp, f: lp.pct_called),
    Column("conc_limit", ("Concentration Limit", "Agent Concentration Limit", "Conc. Limit"),
           "rate", lambda lp, f: lp.conc_limit),
    Column("excess", ("Excess Concentration", "Conc. Overage"),
           "money", lambda lp, f: lp.excess),
    Column("eligible", ("Eligible Unfunded Commitment", "Eligible Commitment",
                        "Post-CL Unfunded Commitment"),
           "money", lambda lp, f: lp.eligible),
    Column("rate", ("Advance Rate", "Agent Advance Rate", "Adv. Rate"),
           "rate", lambda lp, f: lp.rate),
    Column("bb", ("Borrowing Base", "Agent Borrowing Base", "Borrowing Base Contribution",
                  "Agent BB"),
           "money", lambda lp, f: lp.bb),
    Column("pct_bb", ("% of Borrowing Base",),
           "pct", lambda lp, f: lp.pct_bb),
    Column("notes", ("Notes", "Comments", "Remarks"),
           "text", lambda lp, f: ""),
)}

# Six forms, one per shape the sample AgentBBs tree carries. `required` columns always print;
# `optional` ones are drawn per agent, which is what stops two banks on one archetype from
# producing the same template.
ARCHETYPES: dict[str, dict] = {
    "certificate": {
        "sheets": 1,
        "required": ("investor", "itype", "commitment", "uncalled", "size", "sp", "moodys",
                     "rate", "bb"),
        "optional": ("fitch", "conc_limit", "pct_called", "pct_bb", "pct_commit"),
        "grouped": False,
        "sheet_names": ("Borrowing Base", "BB Certificate", "Agent BB", "Borrowing Base Cert"),
    },
    "grouped": {
        "sheets": 1,
        "required": ("investor", "moodys", "sp", "size", "commitment", "called", "uncalled",
                     "pct_uncalled", "conc_limit", "rate"),
        "optional": ("excess", "eligible", "fitch", "lp_category"),
        "grouped": True,
        "sheet_names": ("BB", "Borrowing Base", "BB Detail", "Borrowing Base Detail"),
    },
    "sleeved": {
        "sheets": 1,
        "required": ("investor", "sleeve", "moodys", "sp", "size", "commitment", "called",
                     "uncalled", "pct_uncalled", "conc_limit", "rate", "bb"),
        "optional": ("eligible", "fitch", "region"),
        "grouped": True,
        "sheet_names": ("Borrowing Base", "BB", "Collateral Detail"),
    },
    "feeders": {
        "sheets": (2, MAX_SHEETS),
        "required": ("investor", "excluded_flag", "commitment", "uncalled_ind", "conc_limit"),
        "optional": ("defaulting", "rights", "recallable", "rate"),
        "grouped": False,
        "sheet_names": None,          # feeder/sleeve names, minted per facility
    },
    "deals": {
        "sheets": (2, 4),
        "required": ("transferred", "investor", "partnership", "ga_id", "incl_excl",
                     "commitment", "uncalled", "pct_uncalled", "conc_limit", "excess"),
        "optional": ("eligible", "bb", "rate"),
        "grouped": False,
        "sheet_names": None,          # deal codenames, minted per facility
    },
    "wide": {
        "sheets": 1,
        "required": ("investor", "parent", "sp", "moodys", "size", "commitment_ind", "commitment",
                     "uncalled_ind", "uncalled", "pct_called", "pct_uncalled", "conc_limit",
                     "excess", "eligible", "rate", "bb"),
        "optional": ("pct_bb", "spv", "region"),
        "grouped": True,
        "sheet_names": ("Agent BB", "Borrowing Base", "BB - Proposed"),
    },
}
ARCHETYPE_ORDER = tuple(ARCHETYPES)

# Group banner spellings per Agent LP Category. Each variant still classifies to the same canonical
# category through the recognizer's group classifier — "PWM" is kept in the private-wealth banners
# because that classifier needs it to tell the two Designated tiers apart.
GROUP_BANNERS: dict[str, tuple[str, ...]] = {
    RATED_CATEGORY: ("Included Investors (Rated)", "Rated Included Investors", "A. Rated Investors",
                     "Included Investors - Rated"),
    "Included Investors (Non-Rated)": ("Included Investors (Non-Rated)", "Non-Rated Included Investors",
                                       "B. Non-Rated Investors", "Included Investors - Unrated"),
    "Insitutional Designated Investors": ("Insitutional Designated Investors",
                                          "Institutional Designated Investors",
                                          "C. Institutional Designated Investors"),
    PWM_CATEGORY: ("PWM Designated Investors", "PWM Designated", "D. PWM Designated Investors"),
    EXCLUDED_CATEGORY: ("Excluded Investors", "E. Excluded Investors", "Ineligible Investors"),
}
GROUP_ORDER = (RATED_CATEGORY, "Included Investors (Non-Rated)",
               "Insitutional Designated Investors", PWM_CATEGORY, EXCLUDED_CATEGORY)

SLEEVE_NAMES = ("Onshore Feeder", "Offshore Feeder", "Levered Feeder", "(Cayman) Feeder, L.P.",
                "(Delaware) Feeder, L.P.", "Lux Intermediate", "Lux Non-Treaty Feeder",
                "Master Fund", "Parallel Fund", "AIV I", "BB - Onshore", "BB - Offshore")
DEAL_NAMES = ("Halyard", "Ironclad", "Redstone", "Bluefin", "Northwind", "Cobalt", "Trailhead",
              "Windrow", "Sablefish", "Kingfisher", "Driftwood", "Quarry")

MONEY_FORMATS = ("#,##0", "#,##0.00", '"$"#,##0', '"$"#,##0.00')
PCT_FORMATS = ("0.0%", "0.00%", "0.000%")
RATE_FORMATS = ("0.0%", "0.00%", "0%")
DATE_STYLES = ("%d %b %Y", "%m/%d/%Y", "%Y-%m-%d", "%d-%b-%y")
FILE_DATE_STYLES = ("%Y.%m.%d", "%Y-%m-%d", "%Y%m%d")
FILE_SEPARATORS = (" - ", " ", ". ", "_")
NAME_STYLES = ("full", "dashed", "acronym", "short")

HEADER_FONT = Font(bold=True)
TITLE_FONT = Font(bold=True, size=12)
LABEL_FONT = Font(bold=True)
NEW_LP_FILL = PatternFill("solid", fgColor="CCFFCC")     # the "new to the BB" shading a legend names
XFER_FILL = PatternFill("solid", fgColor="FFF2CC")


@dataclass
class Dialect:
    """One agent bank's house style — fixed for the bank, so all its files share one template."""
    agent: str
    archetype: str
    columns: tuple[str, ...]
    headers: dict[str, str]
    sheet_name: str
    start_col: int
    title_gap: int                  # blank rows between the summary block and the header row
    summary_style: str              # "kv" | "paired" | "block" | "none"
    money_format: str
    pct_format: str
    rate_format: str
    rate_as_text: bool
    date_style: str
    file_date_style: str
    file_separator: str
    name_style: str
    legend: bool
    totals_row: bool
    sheet_count_range: tuple[int, int]


def build_dialect(agent: str, seed: int) -> Dialect:
    """Derive an agent's house style. Seeded on the agent NAME, so a bank's form is stable across
    runs and across the facilities it certifies — which is what makes one template per bank enough."""
    rng = random.Random(_seed_of(seed, "dialect", agent))
    archetype = ARCHETYPE_ORDER[_seed_of(seed, "archetype", agent) % len(ARCHETYPE_ORDER)]
    spec = ARCHETYPES[archetype]

    columns = list(spec["required"])
    for key in spec["optional"]:
        if rng.random() < 0.55:
            columns.insert(rng.randint(1, len(columns)), key)

    headers = {key: _pick(rng, COLUMNS[key].aliases) for key in columns}
    sheets = spec["sheets"]
    sheet_range = sheets if isinstance(sheets, tuple) else (1, 1)
    names = spec["sheet_names"]
    return Dialect(
        agent=agent,
        archetype=archetype,
        columns=tuple(columns),
        headers=headers,
        sheet_name=_pick(rng, names) if names else "",
        start_col=2 if archetype == "wide" else _pick(rng, (1, 1, 1, 2)),
        title_gap=rng.randint(1, 2),
        summary_style=_pick(rng, ("kv", "paired", "block")) if archetype != "deals" else "block",
        money_format=_pick(rng, MONEY_FORMATS),
        pct_format=_pick(rng, PCT_FORMATS),
        rate_format=_pick(rng, RATE_FORMATS),
        rate_as_text=rng.random() < 0.30,
        date_style=_pick(rng, DATE_STYLES),
        file_date_style=_pick(rng, FILE_DATE_STYLES),
        file_separator=_pick(rng, FILE_SEPARATORS),
        name_style=_pick(rng, NAME_STYLES),
        legend=rng.random() < 0.45,
        totals_row=rng.random() < 0.70,
        sheet_count_range=sheet_range,
    )


# ================================================================================================
#  Cell rendering
# ================================================================================================
def render(lp: AgentLp, fac: Facility, key: str, d: Dialect) -> tuple[object, Optional[str]]:
    """(value, number format) for one cell. Money and shares are written at full stored precision;
    the number format is presentation only, which is what POI reads past on the way in."""
    col = COLUMNS[key]
    raw = col.value(lp, fac)
    if col.kind == "money":
        return _money(float(raw or 0)), d.money_format
    if col.kind == "pct":
        return _pct8(float(raw or 0)), d.pct_format
    if col.kind == "rate":
        value = float(raw or 0)
        # Some agents type the rate in as text ("75%", "7.5%") rather than as a number; the
        # extraction has to survive both, so both are generated.
        return (f"{value * 100:g}%", None) if d.rate_as_text else (value, d.rate_format)
    if col.kind == "size":
        header = d.headers.get(key, "")
        if "range" in header.lower():
            return money_range(raw), None
        return money_short(raw), None
    return ("" if raw is None else raw), None


def write_header(ws, row: int, d: Dialect) -> None:
    for i, key in enumerate(d.columns):
        cell = ws.cell(row, d.start_col + i, d.headers[key])
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="bottom")


def write_lp_row(ws, row: int, lp: AgentLp, fac: Facility, d: Dialect) -> None:
    for i, key in enumerate(d.columns):
        value, fmt = render(lp, fac, key, d)
        cell = ws.cell(row, d.start_col + i, value)
        if fmt:
            cell.number_format = fmt
    if d.legend and not lp.carried:
        # The shading a Legend block explains, on the rows it describes: investors new to this BB.
        ws.cell(row, d.start_col).fill = NEW_LP_FILL


def write_total_row(ws, row: int, label: str, lps: list[AgentLp], fac: Facility, d: Dialect) -> None:
    """A subtotal / grand-total band. Its label leads with "Total", which is the keyword
    the recognizer lists as a skip-row so the band never reads as an investor."""
    ws.cell(row, d.start_col, label).font = LABEL_FONT
    for i, key in enumerate(d.columns):
        col = COLUMNS[key]
        if col.kind != "money" or key in ("commitment_ind", "uncalled_ind", "recallable"):
            continue
        total = _money(sum(float(col.value(lp, fac) or 0) for lp in lps))
        cell = ws.cell(row, d.start_col + i, total)
        cell.number_format = d.money_format
        cell.font = LABEL_FONT


def write_summary_block(ws, first_row: int, fac: Facility, d: Dialect,
                        extra: tuple[tuple[str, object], ...] = ()) -> int:
    """The label/value block above the grid. Returns the row after the block.

    Every layout carries an agent label ("Agent Bank" / "Administrative Agent" / "Prepared By"),
    which is the block the recognizer reads an agent name out of when no subdirectory override is
    supplied."""
    pairs: list[tuple[str, object]] = [
        (_pick(random.Random(_seed_of(d.agent, "agentlabel")),
               ("Agent Bank", "Administrative Agent", "Agent")), fac.agent),
        ("Facility", fac.borrower),
        ("As Of Date", fac.bb_date.strftime(d.date_style)),
        ("Currency", "USD"),
        *extra,
    ]
    if fac.is_umbrella:
        # The funds that draw on the agreement. A certificate that covers the group has to name the
        # borrowers it covers, or the only thing stating which funds are on the line is the export.
        pairs += [("Borrowers" if i == 0 else "", m) for i, m in enumerate(fac.members)]
    row = first_row
    if d.summary_style == "paired":
        for i in range(0, len(pairs), 2):
            chunk = pairs[i:i + 2]
            for j, (label, value) in enumerate(chunk):
                ws.cell(row, d.start_col + j * 4, label).font = LABEL_FONT
                ws.cell(row, d.start_col + j * 4 + 1, value)
            row += 1
        return row
    for label, value in pairs:
        ws.cell(row, d.start_col, label).font = LABEL_FONT
        cell = ws.cell(row, d.start_col + 1, value)
        if isinstance(value, float):
            cell.number_format = d.money_format
        row += 1
    return row


def write_legend(ws, row: int, d: Dialect) -> None:
    """The legend footer. Its rows carry the words the recognizer treats as metadata, which
    is exactly why they are here: they must be recognized as a legend and not as investor rows."""
    ws.cell(row, d.start_col, "Legend:").font = LABEL_FONT
    ws.cell(row + 1, d.start_col, "Green Shading: LPs new to the BB not via LP Transfer")
    ws.cell(row + 2, d.start_col, "Yellow Shading: LPs new to the BB via LP Transfer")
    ws.cell(row + 3, d.start_col, "Blue Text: LPs with a change in Commitment Amount")


def group_rows(lps: list[AgentLp]) -> list[tuple[str, list[AgentLp]]]:
    """LPs bucketed into the canonical Agent LP Categories, in schedule order, empties dropped."""
    buckets = {cls: [lp for lp in lps if lp.cls == cls] for cls in GROUP_ORDER}
    return [(cls, rows) for cls, rows in buckets.items() if rows]


def split_sheets(lps: list[AgentLp], count: int) -> list[list[AgentLp]]:
    """Contiguous blocks, so file order survives the split. Falls back to fewer sheets rather than
    write one too thin for the analyzer to read as a grid."""
    count = max(1, min(count, len(lps) // MIN_ROWS_PER_SHEET or 1))
    size = -(-len(lps) // count)
    blocks = [lps[i:i + size] for i in range(0, len(lps), size)]
    return [b for b in blocks if b] or [lps]


def split_by_member(lps: list[AgentLp]) -> list[tuple[str, list[AgentLp]]]:
    """(member, its rows) per block, for the forms that put one borrower on one sheet.

    The roster is already contiguous by member, so the split is a walk. A member too thin to read as
    a grid is folded into the block beside it rather than written to a sheet the analyzer would then
    refuse — its rows stay on the certificate either way, under a neighbour's tab."""
    blocks: list[tuple[str, list[AgentLp]]] = []
    for lp in lps:
        if blocks and blocks[-1][0] == lp.member:
            blocks[-1][1].append(lp)
        else:
            blocks.append((lp.member, [lp]))
    merged: list[tuple[str, list[AgentLp]]] = []
    for name, rows in blocks:
        if merged and len(rows) < MIN_ROWS_PER_SHEET:
            merged[-1][1].extend(rows)
        else:
            merged.append((name, rows))
    while len(merged) > 1 and len(merged[0][1]) < MIN_ROWS_PER_SHEET:
        head = merged.pop(0)
        merged[0][1][:0] = head[1]
    return merged or [("", lps)]


# ================================================================================================
#  Workbook writers — one per archetype
# ================================================================================================
def _new_workbook() -> openpyxl.Workbook:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    return wb


def _autosize(ws, d: Dialect) -> None:
    for i, key in enumerate(d.columns):
        width = 34 if key in ("investor", "partnership", "parent") else 18
        ws.column_dimensions[get_column_letter(d.start_col + i)].width = width


def build_certificate(fac: Facility, d: Dialect, rng: random.Random) -> openpyxl.Workbook:
    """A flat borrowing base certificate: title, summary block, one grid, one total band."""
    wb = _new_workbook()
    ws = wb.create_sheet(d.sheet_name[:31])
    ws.cell(1, d.start_col,
            f"{fac.borrower} {EM_DASH} Agent Borrowing Base Certificate").font = TITLE_FONT
    row = write_summary_block(ws, 3, fac, d, extra=(
        ("Prepared By", f"{fac.agent} Agency Services"),
        ("File Date", fac.bb_date.strftime(d.date_style)),
    ))
    header_row = row + d.title_gap
    write_header(ws, header_row, d)
    row = header_row + 1
    for lp in fac.lps:
        write_lp_row(ws, row, lp, fac, d)
        row += 1
    if d.totals_row:
        write_total_row(ws, row, f"Total {EM_DASH} {len(fac.lps)} Investors", fac.lps, fac, d)
        row += 1
    if d.legend:
        write_legend(ws, row + 1, d)
    _autosize(ws, d)
    return wb


def build_grouped(fac: Facility, d: Dialect, rng: random.Random) -> openpyxl.Workbook:
    """Banner-per-category with a subtotal band under each — the shape the group detector reads."""
    wb = _new_workbook()
    ws = wb.create_sheet(d.sheet_name[:31])
    ws.cell(2, d.start_col, fac.borrower.upper()).font = TITLE_FONT
    row = write_summary_block(ws, 3, fac, d, extra=(
        ("Total Investors", len(fac.lps)),
        ("Total Commitment", fac.tot_commit),
        ("Total Unfunded Commitment", fac.tot_uncalled),
    ))
    header_row = row + d.title_gap
    write_header(ws, header_row, d)
    row = header_row + 1
    banner_rng = random.Random(_seed_of(d.agent, "banners"))
    for cls, rows in group_rows(fac.lps):
        banner = _pick(banner_rng, GROUP_BANNERS[cls])
        ws.cell(row, d.start_col, banner).font = LABEL_FONT
        row += 1
        for lp in rows:
            write_lp_row(ws, row, lp, fac, d)
            row += 1
        write_total_row(ws, row, f"Total {EM_DASH} {banner}", rows, fac, d)
        row += 2
    if d.legend:
        write_legend(ws, row, d)
    _autosize(ws, d)
    return wb


def build_sleeved(fac: Facility, d: Dialect, rng: random.Random) -> openpyxl.Workbook:
    """Grouped, plus a Fund Sleeve column: one sheet covering every feeder of the borrower."""
    wb = _new_workbook()
    ws = wb.create_sheet(d.sheet_name[:31])
    # On an umbrella the sleeve column already has a truth to state — the borrower each row draws
    # under — so the vehicles are not minted over it.
    if fac.is_umbrella:
        for lp in fac.lps:
            lp.sleeve = lp.member
    else:
        sleeves = rng.sample(SLEEVE_NAMES, rng.randint(2, 3))
        for lp in fac.lps:
            lp.sleeve = _pick(rng, sleeves)
    ws.cell(2, d.start_col, f"{fac.borrower} {EM_DASH} Borrowing Base").font = TITLE_FONT
    row = write_summary_block(ws, 3, fac, d, extra=(
        ("Total Investors", len(fac.lps)),
        ("Total Unfunded Commitment", fac.tot_uncalled),
        ("Total Borrowing Base", fac.tot_bb),
        ("Eligibility Basis", "Rated / Non-Rated / Designated"),
    ))
    header_row = row + d.title_gap
    write_header(ws, header_row, d)
    row = header_row + 2
    banner_rng = random.Random(_seed_of(d.agent, "banners"))
    for cls, rows in group_rows(fac.lps):
        banner = _pick(banner_rng, GROUP_BANNERS[cls])
        ws.cell(row, d.start_col, banner).font = LABEL_FONT
        row += 1
        for lp in rows:
            write_lp_row(ws, row, lp, fac, d)
            row += 1
        write_total_row(ws, row, f"Total {EM_DASH} {banner}", rows, fac, d)
        row += 2
    _autosize(ws, d)
    return wb


def build_feeders(fac: Facility, d: Dialect, rng: random.Random) -> openpyxl.Workbook:
    """One sheet per feeder vehicle, each its own small grid with its own total band."""
    wb = _new_workbook()
    if fac.is_umbrella:
        # One sheet per borrower on the agreement: on this form the per-vehicle tab is the shape the
        # agent already uses, and an umbrella's vehicles are its member funds.
        pairs = [(member_label(m, fac), block) for m, block in split_by_member(fac.lps)]
    else:
        blocks = split_sheets(fac.lps, rng.randint(*d.sheet_count_range))
        pairs = list(zip(rng.sample(SLEEVE_NAMES, len(blocks)), blocks))
    for sleeve, block in pairs:
        ws = wb.create_sheet(ILLEGAL_PATH_CHARS.sub("-", sleeve)[:31])
        for lp in block:
            lp.sleeve = lp.member or sleeve
        ws.cell(3, d.start_col, fac.borrower).font = TITLE_FONT
        ws.cell(3, d.start_col + 4, sleeve)
        header_row = 4 + d.title_gap + 2
        write_header(ws, header_row, d)
        row = header_row + 1
        for lp in block:
            write_lp_row(ws, row, lp, fac, d)
            row += 1
        if d.totals_row:
            write_total_row(ws, row, f"Total {EM_DASH} {sleeve}", block, fac, d)
        _autosize(ws, d)
    return wb


def build_deals(fac: Facility, d: Dialect, rng: random.Random) -> openpyxl.Workbook:
    """One sheet per deal, each headed by a Deal Name / Borrowers block instead of a summary."""
    wb = _new_workbook()
    if fac.is_umbrella:
        # This form already prints a Borrowers block and a partnership column, which is what an
        # umbrella needs: the deal is the agreement, and its borrowers are the member funds.
        pairs = [(member_label(m, fac), block) for m, block in split_by_member(fac.lps)]
        partnerships = list(fac.members)
    else:
        blocks = split_sheets(fac.lps, rng.randint(*d.sheet_count_range))
        pairs = list(zip(rng.sample(DEAL_NAMES, len(blocks)), blocks))
        partnerships = [fac.borrower, f"{fac.borrower} (Cayman)"]
    ga = rng.randrange(60, 900)
    for deal, block in pairs:
        ws = wb.create_sheet(ILLEGAL_PATH_CHARS.sub("-", deal)[:31])
        for i, lp in enumerate(block):
            lp.sleeve = lp.member or _pick(rng, partnerships)
            lp.ga_id = f"GA-{ga + i:05d}"
            lp.transferred = "Y" if (not lp.carried and rng.random() < 0.3) else ""
        ga += len(block) + rng.randrange(1, 40)
        ws.cell(4, d.start_col, "Deal Name:").font = LABEL_FONT
        ws.cell(4, d.start_col + 1, deal)
        ws.cell(9, d.start_col, "Borrowers:").font = LABEL_FONT
        for i, p in enumerate(partnerships):
            ws.cell(9 + i, d.start_col + 1, p)
        # The grid clears the Borrowers block rather than sitting at a fixed row: an umbrella lists
        # every fund on the agreement there, and a fixed header row would be written over them.
        header_row = max(13, 9 + len(partnerships) + 1)
        write_header(ws, header_row, d)
        row = header_row + 1
        for lp in block:
            write_lp_row(ws, row, lp, fac, d)
            if lp.transferred:
                ws.cell(row, d.start_col).fill = XFER_FILL
            row += 1
        _autosize(ws, d)
    return wb


def build_wide(fac: Facility, d: Dialect, rng: random.Random) -> openpyxl.Workbook:
    """The analytic form: grid offset off column A, a deep summary block, and category banners."""
    wb = _new_workbook()
    ws = wb.create_sheet(d.sheet_name[:31])
    ws.cell(2, d.start_col,
            f"{fac.borrower} {EM_DASH} Subscription Facility Borrowing Base").font = TITLE_FONT
    block = (
        ("Borrowing Base", fac.tot_bb),
        ("Eligible Remaining Commitments", fac.tot_eligible),
        ("Total Remaining Commitments", fac.tot_uncalled),
        ("Total Original Commitments", fac.tot_commit),
        ("Effective Advance Rate", (fac.tot_bb / fac.tot_eligible) if fac.tot_eligible else 0.0),
        ("Apply Individual Concentration Limits", 1),
    )
    row = 4
    for label, value in block:
        ws.cell(row, d.start_col, label).font = LABEL_FONT
        cell = ws.cell(row, d.start_col + 1, value)
        cell.number_format = d.pct_format if label == "Effective Advance Rate" else d.money_format
        row += 1
    # The legend lives beside the summary block on this form, not under the grid.
    ws.cell(3, d.start_col + 3, "Legend").font = LABEL_FONT
    for i, meaning in enumerate(("Reclassified", "Transferor", "Transferee")):
        ws.cell(4 + i, d.start_col + 3, meaning)
    header_row = row + d.title_gap + 1
    write_header(ws, header_row, d)
    row = header_row + 2
    banner_rng = random.Random(_seed_of(d.agent, "banners"))
    for cls, rows in group_rows(fac.lps):
        ws.cell(row, d.start_col, _pick(banner_rng, GROUP_BANNERS[cls])).font = LABEL_FONT
        row += 1
        for lp in rows:
            write_lp_row(ws, row, lp, fac, d)
            row += 1
        row += 1
    _autosize(ws, d)
    return wb


BUILDERS: dict[str, Callable[[Facility, Dialect, random.Random], openpyxl.Workbook]] = {
    "certificate": build_certificate,
    "grouped": build_grouped,
    "sleeved": build_sleeved,
    "feeders": build_feeders,
    "deals": build_deals,
    "wide": build_wide,
}


# ================================================================================================
#  File naming — the sample tree's conventions, one convention per agent
# ================================================================================================
# Compared against the token with its periods removed, so "L.P.", "L.P" and "LP" are one token.
LEGAL_TOKENS = {"lp", "llc", "ltd", "inc", "llp", "limited", "co", "corp"}
ROMAN_RE = re.compile(r"^[IVXL]+$", re.I)


def _abbreviate(borrower: str, style: str) -> str:
    """The borrower name as this agent writes it on a file name.

    The sample tree carries all four: the name in full ("Petershill IV"), dash-joined
    ("Blue-Owl-GP-Stakes-V"), initialised ("CP VII" for Carlyle Partners VII) and shortened
    ("AEP VII Fund"). The trailing series numeral and any parenthetical are always kept — they are
    what separates one fund in a family, or one sleeve of a credit agreement, from the next."""
    # The sleeve marker is not always last ("... Fund XII (A) LP"), so it is lifted out wherever it
    # sits and re-attached at the end.
    parenthetical = ""
    m = re.search(r"\(([^)]*)\)", borrower)
    core = borrower
    if m:
        parenthetical = f" ({m.group(1)})"
        core = f"{borrower[:m.start()]} {borrower[m.end():]}"

    # The umbrella marker is held out the same way and re-attached whole: initialising it away would
    # leave the one file on the tree that covers a whole credit agreement unmarked as one.
    marker = ""
    m = UMBRELLA_MARKER_RE.search(core)
    if m:
        marker = m.group(1)
        core = core[:m.start()]

    tokens = [t for t in core.replace(",", " ").split() if t]
    while tokens and tokens[-1].replace(".", "").lower() in LEGAL_TOKENS:
        tokens.pop()
    # The series numeral is held out of the acronym — "CP VII", never "CPV".
    tail: list[str] = []
    while tokens and (ROMAN_RE.match(tokens[-1].strip(".")) or tokens[-1].strip(".").isdigit()):
        tail.insert(0, tokens.pop())
    if not tokens:
        tokens, tail = tail, []

    if style == "acronym" and len(tokens) >= 2:
        head = "".join(t[0] for t in tokens if t[0].isalpha()).upper()
    elif style == "short":
        head = " ".join(tokens[:2])
    else:
        head = " ".join(tokens)

    name = " ".join(filter(None, [head, " ".join(tail), marker])) + parenthetical
    if style == "dashed":
        name = re.sub(r"\s+", "-", name.strip())
    return name.strip()


def file_name_for(fac: Facility, d: Dialect) -> str:
    """<name><separator><date>.xlsx, in this agent's convention."""
    stem = f"{_abbreviate(fac.borrower, d.name_style)}{d.file_separator}" \
           f"{fac.bb_date.strftime(d.file_date_style)}"
    stem = ILLEGAL_PATH_CHARS.sub("", stem).strip().rstrip(".")
    return f"{stem or fac.account}.xlsx"


def resolve_path(directory: Path, name: str) -> Path:
    """A free path in `directory`. Two borrowers that abbreviate to one file name under the same
    agent are disambiguated rather than silently overwritten."""
    path = directory / name
    if not path.exists():
        return path
    stem, suffix = path.stem, path.suffix
    for n in range(2, 100):
        alt = directory / f"{stem} ({n}){suffix}"
        if not alt.exists():
            return alt
    raise SystemExit(f"Too many file-name collisions for '{name}' in {directory}")


# ================================================================================================
#  Verification
# ================================================================================================
def verify_reconciliation(fac: Facility) -> None:
    """Re-derive every stated figure from the row's own stated inputs; fail the run on disagreement.

    A certificate whose borrowing base does not follow from the rate and limit printed beside it is
    indistinguishable from the engine being wrong once it has been ingested, so it never reaches
    disk."""
    tot_u = fac.tot_uncalled or 1.0
    for lp in fac.lps:
        excluded = lp.cls == EXCLUDED_CATEGORY
        eligible = lp.uncalled if lp.conc_limit <= 0 else min(lp.uncalled, lp.conc_limit * tot_u)
        want_excess = _money(0.0 if excluded else lp.uncalled - eligible)
        want_eligible = _money(0.0 if excluded else lp.uncalled - want_excess)
        want_bb = _money(0.0 if excluded else want_eligible * lp.rate)
        for label, got, want in (("excess", lp.excess, want_excess),
                                 ("eligible", lp.eligible, want_eligible),
                                 ("borrowing base", lp.bb, want_bb)):
            if abs(got - want) > 0.01:
                raise SystemExit(f"{fac.borrower} / {lp.name}: {label} {got} != {want}")
        if abs(_money(lp.called + lp.uncalled) - lp.commitment) > 0.01:
            raise SystemExit(f"{fac.borrower} / {lp.name}: called + uncalled != commitment")
        if excluded and lp.rate != 0.0:
            raise SystemExit(f"{fac.borrower} / {lp.name}: Excluded row carries a {lp.rate} rate")
        if (lp.cls == RATED_CATEGORY) != lp.rated:
            raise SystemExit(f"{fac.borrower} / {lp.name}: category {lp.cls!r} disagrees with its "
                             f"ratings {(lp.sp, lp.moodys, lp.fitch)}")


def verify_divergence(fac: Facility) -> None:
    """Every carried row must state a credit view the export does not — the point of carrying it is
    to be the SAME investor under a DIFFERENT opinion, and an identical row would test nothing."""
    for lp in fac.lps:
        if not lp.carried:
            continue
        src = lp.source
        if str(src.get("Agent LP Classification") or "").strip() == lp.cls:
            raise SystemExit(f"{fac.borrower} / {lp.name}: LP category identical to the export")
        if str(src.get("Investor Type") or "").strip() == lp.itype:
            raise SystemExit(f"{fac.borrower} / {lp.name}: investor type identical to the export")
        src_ratings = (str(src.get("S&P") or "NR").strip() or "NR",
                       str(src.get("Moody'S") or "NR").strip() or "NR",
                       str(src.get("Fitch") or "NR").strip() or "NR")
        if src_ratings == (lp.sp, lp.moodys, lp.fitch) and src_ratings != ("NR", "NR", "NR"):
            raise SystemExit(f"{fac.borrower} / {lp.name}: ratings identical to the export")
        if abs(float(src.get("Agent Borrowing Base") or 0) - lp.bb) < 0.01 and lp.bb > 0:
            raise SystemExit(f"{fac.borrower} / {lp.name}: borrowing base identical to the export")


def verify_umbrella(fac: Facility) -> None:
    """A group's certificate must say it is one, and must carry every borrower on the agreement.

    Both are what makes the file readable as an umbrella rather than as one more facility with an
    unusually long roster: the name states the structure, and the roster states its extent. The
    account number is checked out of the name for the same reason it is never put there — a name
    that restates the key the group is already filed under tells a reader nothing."""
    if not fac.is_umbrella:
        return
    if not any(marker.casefold() in fac.borrower.casefold() for marker in UMBRELLA_MARKERS):
        raise SystemExit(f"{fac.borrower}: umbrella name carries none of {UMBRELLA_MARKERS}")
    if fac.account and fac.account.casefold() in fac.borrower.casefold():
        raise SystemExit(f"{fac.borrower}: umbrella named for its account number")
    missing = [m for m in fac.members if m not in {lp.member for lp in fac.lps}]
    if missing:
        raise SystemExit(f"{fac.borrower}: certificate carries no rows for {missing}")


def verify_parseable(analyzer: ExcelAnalyzer, path: Path, expected_tabs: int) -> Optional[str]:
    """Run the recognizer this generator writes for. Returns a complaint, or None when the workbook
    is one the directory crawler will turn into a template."""
    analysis = analyzer.analyze_workbook(path, agent_bank_override="verify")
    if not analysis.tabs:
        return "no LP-grid sheet recognized"
    if len(analysis.tabs) != expected_tabs:
        return (f"{len(analysis.tabs)} of {expected_tabs} sheet(s) recognized as grids "
                f"(non-grid: {analysis.non_grid_sheets})")
    thin = [t.sheet_name for t in analysis.tabs if t.matched_canonical < analyzer.min_header_matches]
    if thin:
        return f"header row below the alias threshold on {thin}"
    return None


# ================================================================================================
#  Run
# ================================================================================================
def generate(out_root: Path, seed: int, agents: Optional[set[str]], limit: Optional[int],
             verify: bool) -> tuple[int, int, list[dict]]:
    for source in (ABS_IN, EXPORT_IN):
        if not source.is_file():
            raise SystemExit(f"Source workbook not found: {source}\n"
                             f"Generate the simulated LP database first — it writes both inputs.")

    logger.info(f"Reading {ABS_IN.name}")
    facilities = load_report_facilities(ABS_IN)
    logger.info(f"Reading {EXPORT_IN.name}")
    export = load_export_rows(EXPORT_IN)
    logger.info(f"{len(facilities)} Active facility row(s), {len(export)} facility key(s) in the export")

    umbrellas = umbrella_members(export)
    facilities = fold_umbrellas(facilities, umbrellas)
    if umbrellas:
        logger.info(f"{len(umbrellas)} umbrella account(s): "
                    + ", ".join(f"{a} ({len(m)} borrowers)" for a, m in umbrellas.items()))

    if agents:
        facilities = [f for f in facilities if f.agent in agents]
    if limit:
        facilities = facilities[:limit]

    dialects: dict[str, Dialect] = {}
    minted_names: set[str] = set()
    analyzer = ExcelAnalyzer(load_dictionary()) if verify else None
    records: list[dict] = []
    written = failed = 0

    for fac in facilities:
        if fac.is_umbrella:
            rows = [r for m in fac.members for r in export.get((fac.account, m), [])]
        else:
            rows = export.get((fac.account, fac.borrower), [])
        # The BB date is the export's own collateral date for this facility — the two files then
        # describe the same run rather than two runs that happen to share a roster. A group is
        # certified once, on the LATEST of its members' dates: an agreement is not as of a date one
        # of the funds on it has already moved past.
        if fac.is_umbrella:
            dates = [d for d in (_as_date(r.get("BBDate")) for r in rows) if d]
            bb_date = max(dates) if dates else None
        else:
            bb_date = _as_date(rows[0].get("BBDate")) if rows else None
        fac.bb_date = bb_date or (fac.status_date or date.today())

        d = dialects.get(fac.agent) or build_dialect(fac.agent, seed)
        dialects[fac.agent] = d
        rng = random.Random(_seed_of(seed, fac.agent, fac.account, fac.borrower))

        record = {"Agent": fac.agent, "Borrower": fac.borrower, "Layout": d.archetype,
                  "LPs": "0", "Carried": "0", "File": "", "Status": "UNKNOWN",
                  "Notes": f"umbrella of {len(fac.members)} borrowers" if fac.is_umbrella else ""}
        try:
            build_roster(fac, rows, rng, minted_names)
            price_facility(fac, rng)
            verify_reconciliation(fac)
            verify_divergence(fac)
            verify_umbrella(fac)

            wb = BUILDERS[d.archetype](fac, d, rng)
            directory = out_root / ILLEGAL_PATH_CHARS.sub("-", fac.agent)
            directory.mkdir(parents=True, exist_ok=True)
            path = resolve_path(directory, file_name_for(fac, d))
            wb.save(path)
            wb.close()

            record["LPs"] = str(len(fac.lps))
            record["Carried"] = f"{sum(1 for lp in fac.lps if lp.carried)}"
            record["File"] = path.name
            record["Status"] = "SUCCESS"

            if analyzer is not None:
                complaint = verify_parseable(analyzer, path, len(wb.sheetnames))
                if complaint:
                    record["Status"] = "UNPARSEABLE"
                    record["Notes"] = complaint
                    failed += 1
                    records.append(record)
                    continue
            written += 1
        except SystemExit:
            raise
        except Exception as e:                       # one bad facility must not lose the run
            record["Status"] = "FAILED"
            record["Notes"] = str(e)
            failed += 1
        records.append(record)

    return written, failed, records


def render_run_summary(records: list[dict], written: int, failed: int,
                       dialects: dict[str, Dialect]) -> str:
    fieldnames = ["Agent", "Borrower", "Layout", "LPs", "Carried", "File", "Status", "Notes"]
    widths = [max(len(n), *(len(str(r.get(n, ""))) for r in records)) if records else len(n)
              for n in fieldnames]

    def row(values: list[str]) -> str:
        return "  ".join(v.ljust(w) for v, w in zip(values, widths)).rstrip()

    lines = ["Run summary", row(fieldnames), row(["-" * w for w in widths])]
    lines.extend(row([str(r.get(n, "")) for n in fieldnames]) for r in records)
    lines.extend([
        "",
        f"Workbooks written : {written}",
        f"Failed            : {failed}",
        f"Agent banks       : {len(dialects)}",
    ])
    for agent, d in sorted(dialects.items()):
        lines.append(f"  {agent:<24} {d.archetype:<12} {len(d.columns)} cols, "
                     f"sheet {d.sheet_name or '(per facility)'!r}")
    return "\n".join(lines)


def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="lp_agent_bb_generate.py",
        description="Generate simulated Agent BB workbooks — one per Active facility, in "
                    "subdirectories named for the Agent — under the import directory.",
    )
    p.add_argument("--out", default=str(DEFAULT_OUT),
                   help=f"output root (default: {DEFAULT_OUT})")
    p.add_argument("--seed", type=int, default=SEED, help=f"run seed (default: {SEED})")
    p.add_argument("--agent", action="append", default=None,
                   help="only this Agent (repeatable)")
    p.add_argument("--limit", type=int, default=None,
                   help="stop after this many facilities (smoke runs)")
    p.add_argument("--no-verify", action="store_true",
                   help="skip re-analyzing each finished workbook with the template recognizer")
    p.add_argument("--quiet", action="store_true", help="suppress the per-facility run summary")
    return p


def main(argv: Optional[list[str]] = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass

    handler = logging.StreamHandler(sys.stderr)
    handler.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)

    args = build_arg_parser().parse_args(argv)
    out_root = Path(args.out)
    if not out_root.is_absolute():
        out_root = Path.cwd() / out_root

    dialects_seen: dict[str, Dialect] = {}
    written, failed, records = generate(
        out_root, args.seed, set(args.agent) if args.agent else None, args.limit,
        verify=not args.no_verify,
    )
    for r in records:
        dialects_seen.setdefault(r["Agent"], build_dialect(r["Agent"], args.seed))

    if not args.quiet:
        print("\n" + render_run_summary(records, written, failed, dialects_seen))
    print(f"\nwrote {written} workbook(s) to {out_root}")
    if failed:
        # Never suppressed by --quiet: a workbook the recognizer cannot read is the one result
        # that must not be missed, and the summary table it would otherwise appear in is hidden.
        print(f"{failed} workbook(s) did not verify:")
        for r in records:
            if r["Status"] != "SUCCESS":
                print(f"  {r['Status']:<12} {r['Agent']} / {r['Borrower']} — {r['Notes']}")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
