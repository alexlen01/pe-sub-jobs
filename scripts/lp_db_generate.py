#!/usr/bin/env python3
r"""
LP DB Export generator (dev utility; companion to lp_db_extract.py).

Writes one file: data/import/<EXPORT_OUT>, the XLSX lp_db_extract.py reads. Run the extract
afterwards to produce the three ingestion CSVs.

Shape of the generated data:
  * TARGET_ROWS distinct (facility, investor) positions — the lp_records grain.
  * Each LP holds a position in REPEAT_MIN..REPEAT_MAX different facilities, at most once per
    facility.
  * Facilities are the accounts in data/mock/facilities.csv plus ORPHAN_ACCOUNTS, which are absent
    from that file and exercise the extract's "Unknown" placeholder path.
  * An LP's identity, classification, ratings, scale and UBS credit profile are constant across all
    its rows; only the facility-specific financials (commitment, called/uncalled, BB) vary.
  * Base values use the canonical Agent LP Category / UBS LP Classification vocabularies from
    data/reference/, so with chaos off the extract's unmatched counts come back zero.
  * Agency-rating presence follows the Agent LP Category (RATING_PRESENCE). classify_ubs maps any
    usable rating to "Rated Investor", so tying presence to the agent bucket keeps the UBS mix near
    the real book's ~40% Rated and leaves unrated LPs to exercise the Corp Pension / Unrated NAV /
    FoF / HNW / Other branches.
  * Emits the 2026-08-18 29-column format: the UBS LP Classification is written out as its own
    column (classify_ubs runs HERE now, on the clean values, rather than in the extract), the old
    AUM / NAV / Pension Assets trio is replaced by one "LP Size ($ Bil)" figure plus its criteria
    label, and Agent / UBS Excess Concentration are computed against the facility's total uncalled
    with each borrowing base advanced only on the uncalled that stays within the cap.

Chaos monkey (pe-sub-docs/"AI Chaos Monkey for Data Quality.md"): degrades the values written to
the XLSX, per the analyst "hierarchy of care". Sacred cash/identity columns (CHAOS_SACRED) stay
exact; decision fields (ratings, UBS and agent Classification) get formatting and categorical drift
('A minus', 'Rated', 'PWM'); LP Size Criteria is a closed vocabulary ("AUM"/"NAV"/"Assets") so its
only degradation is a blank cell; afterthought fields (investor names, LP
Size) get suffix drops, ' - Tranche A', case flips, and size strings that argue with the column's
$Bn unit - ranges ('5 - 8'), thresholds ('>12'), spelled units ('13.5 bn') and figures typed in
millions.
The same LP is therefore spelled differently across rows, exercising every parser, normalizer and
report in lp_db_extract.py. Chaos draws from its own CHAOS_SEED rng, so the underlying clean dataset
is identical whether CHAOS_ENABLED is True or False, and re-running with the same seed reproduces
the same degradation. Mutation counts per column and per pattern print to the console.

No command-line arguments — tune the constants below and re-run.
"""
from __future__ import annotations

import csv
import math
import random
import re
from collections import Counter
from datetime import date, timedelta
from pathlib import Path

import openpyxl

SCRIPT_DIR = Path(__file__).resolve().parent          # pe-sub-jobs/scripts/
DATA_DIR = SCRIPT_DIR.parent / "data"                 # pe-sub-jobs/data/
FACILITIES_FILE = DATA_DIR / "mock" / "facilities.csv"
EXPORT_OUT = DATA_DIR / "import" / f"LP DB Export {date.today():%Y.%m.%d}.xlsx"
SHEET_NAME = "BBs"

# ── tunables ────────────────────────────────────────────────────────────────
SEED = 20260625
CHAOS_ENABLED = True            # degrade the written XLSX to realistic manual-entry quality
CHAOS_SEED = 20260625           # chaos has its own rng: base data identical with chaos on/off
TARGET_ROWS = 20_000            # distinct lp_records to produce
REPEAT_MIN, REPEAT_MAX = 4, 12  # facilities each LP participates in
ORPHAN_ACCOUNTS = [             # AccountIDs absent from facilities.csv -> exercise "Unknown" path
    ("5VZ9001", "TPG AG Asset Based Credit Fund"),
    ("5VZ9002", "TPG AG Asset Based Credit Fund"),
]

# The 29 headers of the 2026-08-18 LP DB Export, in order, EXACTLY as the real file spells them —
# quirks included, because reproducing them is most of the point of this generator: "LP Size" and
# "($ Bil)" are separated by a CRLF inside the one cell, "Insitutional" is misspelt at source, and
# "Moody'S" carries a capital S. lp_db_extract matches headers through _norm(), which absorbs all
# three, so this stays a faithful sample rather than a cleaned-up one.
SRC_HEADERS = [
    "AccountID", "FndName", "Investor Name", "Parent", "SPV", "UBS LP Classification",
    "Insitutional vs HNW", "Investment Grade?", "Agent LP Classification", "S&P", "Moody'S",
    "Fitch", "LP Size\r\n($ Bil)", "LP Size Criteria", "Capital Commitments", "Uncalled Capital",
    "UBS Advance Rate", "Agent Concentration Limit", "UBS Concentration Limit",
    "% of Capital Commitments", "Called Capital", "% of Uncalled Capital", "% of LP Called",
    "Agent Excess Concentration", "UBS Excess Concentration", "Agent Borrowing Base",
    "UBS Borrowing Base", "Notes", "BBDate",
]

# Internal keys for the same 29 columns, in the same order. Mirrors lp_db_extract.SRC_COLS so the
# chaos monkey can address a column by name and the two scripts stay legible side by side.
SRC_COLS = [
    "AccountID", "FndName", "InvestorName", "Parent", "SPV", "UbsClassification",
    "InstitutionalHNW", "InvestmentGrade", "Classification", "SP", "Moodys", "Fitch",
    "LpSizeBil", "LpSizeCriteria", "Commitments", "Uncalled", "UBSAR", "AgentCL", "UBSCL",
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

# Agent LP Category (export "Agent LP Classification") -> (agent advance rate, weight). Canonical
# per data/reference/agent_lp_categories.csv, so each value maps to itself.
#
# The rates MUST match data/reference/agent_rate_map.csv. The export no longer carries an Agent
# Advance Rate column, so the extract resolves the rate from the category — and these rates are what
# the Agent Borrowing Base below is computed with. If the two drift apart, the generated file stops
# reconciling: its Agent BB would not equal the rate the extract assigns x the eligible uncalled.
AGENT_CATEGORIES = [
    ("Rated Included",            0.90, 30),
    ("Non-Rated Included",        0.75, 34),
    ("Designated Institutional",  0.60, 18),
    ("Designated PWM",            0.50, 10),
    ("Ineligible Investor",       0.00,  8),
]

# Raw UBSAR bands to draw from (band_lo, band_hi, weight). Rates are continuous: the extract floor-
# maps them to the 90/75/65/50/0 groups (reference/rate_floor_map.csv), derives ubs_classification
# from the LP's attributes via reference/bb_criteria_matrix.csv rather than from the rate, then
# cross-checks rate against matrix and reports the deviation counts.
UBS_BANDS = [
    (0.90, 0.97, 28),
    (0.75, 0.90, 26),
    (0.65, 0.75, 20),
    (0.50, 0.65, 16),
    (0.30, 0.50, 10),
]

IG_RATINGS   = ["AAA", "AA+", "AA", "AA-", "A+", "A", "A-", "BBB+", "BBB", "BBB-"]
NONIG_RATINGS = ["BB+", "BB", "BB-", "B+", "B", "B-"]

# P(LP carries usable agency ratings), by Agent LP Category. The extract's classify_ubs waterfall
# sends any usable rating to "Rated Investor" (sub-IG clamps to the BBB band and stays Rated), so
# presence tracks the agent's own bucket to hold derived Rated near the real book's ~40%. Unrated
# LPs carry NR/blank agency columns. Crossover (a rated LP in a non-rated agent bucket) stays small
# but non-zero, keeping the extract's class-vs-rate variance report exercised.
RATING_PRESENCE = {
    "Rated Included": 1.00,
    "Non-Rated Included": 0.15,
    "Designated Institutional": 0.12,
    "Designated PWM": 0.05,
    "Ineligible Investor": 0.20,
}

WORD1 = ["Apex", "Meridian", "Granite", "Harborview", "Ironwood", "Cascade", "Dominion", "Everest",
         "Fairview", "Northgate", "Oakmont", "Pinnacle", "Redwood", "Silverpeak", "Trident", "Union",
         "Westbridge", "Yorkshire", "Zenith", "Brookstone", "Cedar", "Kestrel", "Larchmont", "Monarch",
         "Sterling", "Beacon", "Clearwater", "Highgate", "Lakeshore", "Ridgeline", "Summit", "Vantage",
         "Ashford", "Bluewater", "Copperfield", "Draycott", "Eastgate", "Fenwick", "Glenwood", "Halcyon"]
WORD2 = ["", "", "", "Capital", "Global", "Strategic", "Alternative", "Atlantic", "Pacific",
         "Continental", "Heritage", "Legacy", "Pioneer", "Cornerstone", "Evergreen"]


# ── chaos monkey (pe-sub-docs/"AI Chaos Monkey for Data Quality.md") ────────────────────────
# Probability per row that each field family is degraded.
CHAOS_RATES = {
    "investor_name": 0.25,     # suffix drops / tranche suffixes / case flips
    "ratings": 0.15,           # 'A-' -> 'A minus', case noise, NR variants (all three agencies)
    "lp_size": 0.15,           # LP Size -> range/threshold/shorthand strings ('5 - 8', '>10', '2bn')
    "lp_size_criteria": 0.08,  # size basis left blank (the label itself never drifts)
    "ubs_classification": 0.12,  # UBS class drift ('Rated', 'Corp Pension >5Bn', 'HNW')
    "agent_category": 0.10,    # agent Classification drift ('Rated', 'PWM', 'Ineligible')
    "ubscl_null": 0.05,        # missing concentration limit
}

# Never degraded: cash/legal LPA figures plus the facility join keys, which analysts keep exact.
CHAOS_SACRED = ("AccountID", "FndName", "Commitments", "Called", "Uncalled", "BBDate",
                "UBSAR", "AgentBB", "UBSBB", "AgentExcessConc", "UBSExcessConc",
                "PercentOfCommitments", "PercentOfUncalled", "CalledPercent")

_NAME_SUFFIX_RE = re.compile(r",?\s+(LLC|L\.L\.C\.|L\.P\.|LP|Ltd\.?|Inc\.?|Limited)$", re.I)

# Canonical value -> manual-entry variants: a mix of alias-resolvable spellings (exercise the
# extract's reference lists) and unknown labels (exercise its unmatched reports).
# UBS LP Classification drift. The export now states this field outright, so it is a decision field
# an analyst types — and therefore drifts. Some variants resolve through ubs_lp_categories.csv,
# others are unknown labels that exercise the extract's unmatched report.
_UBS_CLS_DRIFT = {
    "Rated Investor": ["Rated", "Rated Included", "RATED INVESTOR"],
    "Corp Pension > $5Bn Assets": ["Corp Pension >5Bn", "Corporate Pension > $5Bn Assets",
                                   "Corp Pension 5Bn+"],
    "Corp Pension > $1Bn Assets": ["Corp Pension >1Bn", "Corporate Pension > $1Bn Assets",
                                   "Corp Pension 1Bn+"],
    "Unrated NAV > $1Bn": ["Unrated NAV", "Unrated NAV >1Bn", "NAV > 1Bn"],
    "FoF & Other > $10Bn AUM": ["FoF & Other", "FoF and Other > $10Bn AUM", "FOF >10Bn"],
    "Other Institutional": ["Other Inst", "Other Institutional Investor", "Other"],
    "HNW Feeder (acceptable)": ["HNW Feeder", "HNW Feeder Acceptable"],
    "HNW (acceptable)": ["HNW", "HNW Acceptable"],
    "Excluded": ["Ineligible", "Excluded Investor", "Not Eligible"],
}

_AGENT_DRIFT = {
    "Rated Included": ["Rated", "Rated Included Investors", "rated included"],
    "Non-Rated Included": ["Non Rated Included Investors", "Included Investors", "Non-Rated"],
    "Designated Institutional": ["Designated", "Designated Investors",
                                 "Fund-Of-Fund Designated Investors"],
    "Designated PWM": ["PWM", "Designated - PWM"],
    "Ineligible Investor": ["Ineligible", "Ineligible Investors", "Excluded Investor"],
}


def _blank(v) -> bool:
    return v is None or str(v).strip() == ""


def _as_str(v) -> str:
    return "" if v is None else str(v)


def _chaos_name(name: str, rng: random.Random) -> tuple[str, str]:
    """Entity-name drift: dropped legal suffix, ' - Tranche A', case flip."""
    options = [("tranche_suffix", name + " - Tranche A"),
               ("case_flip", name.upper() if rng.random() < 0.5 else name.lower())]
    stripped = _NAME_SUFFIX_RE.sub("", name).strip()
    if stripped != name:
        options.append(("suffix_dropped", stripped))
    return rng.choice(options)


def _chaos_rating(val: str, rng: random.Random) -> tuple[str, str | None]:
    """Rating formatting drift: 'A-' -> 'A minus', NR variants, case noise."""
    s = str(val).strip()
    if s.upper() in ("NR", "N/A"):
        return "nr_variant", rng.choice(["Not Rated", "N/R", None])
    if "-" in s:
        return "sign_spelled", s.replace("-", " minus")
    if "+" in s:
        return "sign_spelled", s.replace("+", " plus")
    return "case_noise", s.lower()


def _chaos_lp_size(val, rng: random.Random) -> tuple[str, str] | None:
    """LP Size manual-entry patterns. The column's unit is $Bn and its clean value is a bare number
    (13.5), so the realistic corruptions are the ones that argue with that convention: a range, a
    qualitative threshold, a spelled-out unit, or - the costly one - a figure typed in MILLIONS into
    a billions column. Derived from the clean value so the result stays plausible."""
    try:
        bil = float(val)
    except (TypeError, ValueError):
        return None
    if bil <= 0:
        return None
    pattern = rng.choice(["range", "threshold", "unit_spelled", "unit_swap_millions",
                          "absolute_dollars"])
    if pattern == "range":
        low, high = round(bil * 0.7, 1), round(bil * 1.3, 1)
        return pattern, f"{low} - {high}"
    if pattern == "threshold":
        return pattern, f"{rng.choice('><')}{round(bil)}"
    if pattern == "unit_spelled":
        return pattern, rng.choice([f"${bil}B", f"{bil} bn", f"${bil}Bn"])
    if pattern == "unit_swap_millions":
        return pattern, str(round(bil * 1000, 1))       # $Bn figure typed as $mm
    return pattern, str(int(bil * 1e9))                  # absolute dollars in a $Bn column


def apply_chaos(export_rows: list[dict], rng: random.Random) -> list[tuple]:
    """Degrade the export rows in place (CHAOS_RATES per field family; CHAOS_SACRED untouched).
    Returns one (xlsx_row_no, column, pattern, original, corrupted) record per mutation."""
    muts: list[tuple] = []

    def mutate(row_no: int, row: dict, col: str, result: tuple[str, str | None] | None) -> None:
        if result is None:
            return
        pattern, new = result
        if new == row[col]:
            return
        muts.append((row_no, col, pattern, _as_str(row[col]), _as_str(new)))
        row[col] = new

    for row_no, row in enumerate(export_rows, start=2):   # 2-based: XLSX row incl. header
        if not _blank(row["InvestorName"]) and rng.random() < CHAOS_RATES["investor_name"]:
            mutate(row_no, row, "InvestorName", _chaos_name(str(row["InvestorName"]), rng))
        if rng.random() < CHAOS_RATES["ratings"]:
            for col in ("SP", "Moodys", "Fitch"):
                if not _blank(row[col]):
                    mutate(row_no, row, col, _chaos_rating(row[col], rng))
        if not _blank(row["LpSizeBil"]) and rng.random() < CHAOS_RATES["lp_size"]:
            mutate(row_no, row, "LpSizeBil", _chaos_lp_size(row["LpSizeBil"], rng))
        if (not _blank(row["LpSizeCriteria"])
                and rng.random() < CHAOS_RATES["lp_size_criteria"]):
            mutate(row_no, row, "LpSizeCriteria", ("nulled", None))
        if rng.random() < CHAOS_RATES["ubs_classification"]:
            variants = _UBS_CLS_DRIFT.get(_as_str(row["UbsClassification"]))
            if variants:
                mutate(row_no, row, "UbsClassification",
                       ("categorical_drift", rng.choice(variants)))
        if rng.random() < CHAOS_RATES["agent_category"]:
            variants = _AGENT_DRIFT.get(_as_str(row["Classification"]))
            if variants:
                mutate(row_no, row, "Classification", ("categorical_drift", rng.choice(variants)))
        if not _blank(row["UBSCL"]) and rng.random() < CHAOS_RATES["ubscl_null"]:
            mutate(row_no, row, "UBSCL", ("nulled", None))
    return muts


def load_facilities() -> list[tuple[str, str]]:
    """(account_number, fund_name) per facility in facilities.csv, plus ORPHAN_ACCOUNTS."""
    with FACILITIES_FILE.open(newline="", encoding="utf-8") as fh:
        rows = list(csv.reader(fh))
    out: list[tuple[str, str]] = []
    seen: set[str] = set()
    for r in rows[1:]:
        # Skip blanks and "Unknown" placeholder rows the extract may have written back, and dedup by
        # account, so regenerating after copying extract output into mock/ stays idempotent.
        if r and r[0].strip() and r[0].strip() != "Unknown" and r[2].strip() and r[2].strip() not in seen:
            seen.add(r[2].strip())
            out.append((r[2].strip(), r[1].strip()))
    out.extend(ORPHAN_ACCOUNTS)
    return out


def money(lo: int, hi: int, step: int = 100_000) -> int:
    return random.randrange(lo, hi, step)


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
      1. agent 'Ineligible Investor' -> Excluded;
      2. any usable agency rating -> Rated Investor;
      3. HNW (flag, or agent 'Designated PWM') -> HNW Feeder when the vehicle is an SPV, else HNW;
      4. pension assets > $5Bn / > $1Bn -> the two Corp Pension classes;
      5. NAV > $1Bn -> Unrated NAV > $1Bn;
      6. FoF/hedge fund with AUM > $10Bn -> FoF & Other > $10Bn AUM;
      7. catch-all -> Other Institutional."""
    if agent_cat == "Ineligible Investor":
        return "Excluded"
    if rated:
        return "Rated Investor"
    if hnw or agent_cat == "Designated PWM":
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


def rating_triplet(ig: bool) -> tuple[str, str, str]:
    pool = IG_RATINGS if ig else NONIG_RATINGS
    i = random.randrange(len(pool))
    jitter = lambda: pool[min(len(pool) - 1, max(0, i + random.randint(-1, 1)))]
    return jitter(), jitter(), jitter()


def pick_weighted(options, weights):
    return random.choices(options, weights=weights, k=1)[0]


def build_investor(idx: int, used_names: set) -> dict:
    itype = pick_weighted(list(TYPE_WEIGHTS), list(TYPE_WEIGHTS.values()))
    suffixes, measure = TYPE_SPECS[itype]
    while True:
        w2 = random.choice(WORD2)
        base = f"{random.choice(WORD1)}{(' ' + w2) if w2 else ''}"
        name = f"{base} {random.choice(suffixes)}"
        if name not in used_names:
            used_names.add(name)
            break
    agent_cat, agent_ar = pick_weighted(
        [(c, r) for c, r, _ in AGENT_CATEGORIES], [w for *_, w in AGENT_CATEGORIES])
    rated = random.random() < RATING_PRESENCE[agent_cat]
    if rated:
        ig = random.random() < (0.80 if agent_cat == "Rated Included" else 0.35)
        sp, mdy, fitch = rating_triplet(ig)
    else:
        ig = False
        sp, mdy, fitch = (random.choice(["NR", "NR", ""]) for _ in range(3))
    lo, hi, _ = pick_weighted(UBS_BANDS, [w for *_, w in UBS_BANDS])
    ubsar = round(random.uniform(lo, hi), 2)
    is_pension = measure == "pension_assets"
    spv = random.random() < 0.12
    hnw = itype == "Family Office" and random.random() < 0.5

    # nav-measure LPs (FoF / Hedge Fund) carry manager-level AUM alongside fund NAV, which the
    # "FoF & Other > $10Bn AUM" branch needs. Only the LP's OWN measure reaches the export, as the
    # single LP Size figure — the others exist here purely to classify it.
    aum_bil = size_bil("aum") if measure in ("aum", "nav") else None
    nav_bil = size_bil("nav") if measure == "nav" else None
    pension_bil = size_bil("pension_assets") if is_pension else None
    lp_size_bil = {"aum": aum_bil, "nav": nav_bil, "pension_assets": pension_bil}[measure]

    return {
        "name": name,
        "parent": name,
        "spv": "Y" if spv else "N",
        "itype": itype,
        "inst": "HNW" if hnw else "Institutional",
        "ig": "Yes" if ig else "No",
        "cls": agent_cat,
        "agent_ar": agent_ar,
        "sp_rating": sp, "moodys_rating": mdy, "fitch_rating": fitch,
        "lp_size_bil": lp_size_bil,
        "lp_size_criteria": SIZE_CRITERIA[measure],
        "ubs_cls": classify_ubs(agent_cat=agent_cat, rated=rated, hnw=hnw, spv=spv, itype=itype,
                                aum_bil=aum_bil, nav_bil=nav_bil, pension_bil=pension_bil),
        "ubsar": ubsar,
        "agent_cl": round(random.uniform(0.05, 0.15), 2),
        "ubs_cl": round(random.uniform(0.05, 0.15), 2),
    }


def weighted_sample_without_replacement(items, weights, k):
    """Efraimidis-Spirakis: key = U^(1/w); take the k largest keys."""
    keyed = sorted(((random.random() ** (1.0 / w), it) for it, w in zip(items, weights)), reverse=True)
    return [it for _, it in keyed[:k]]


def main() -> int:
    random.seed(SEED)
    facilities = load_facilities()                    # [(acct, fund), ...]
    base = date(2026, 6, 25)
    # One BBDate per facility, M/D/YYYY (formatted manually: Windows strftime lacks %-m/%-d).
    fac_bbdate = {}
    for acct, _ in facilities:
        d = base - timedelta(days=random.randint(0, 240))
        fac_bbdate[acct] = f"{d.month}/{d.day}/{d.year}"

    fac_weights = {acct: random.uniform(0.4, 3.2) for acct, _ in facilities}
    fund_by_acct = dict(facilities)
    accts = [a for a, _ in facilities]

    used_names: set = set()
    positions: list[dict] = []        # each dict = one lp_record row (pre-percent)
    per_fac: dict[str, list[dict]] = {a: [] for a in accts}
    investor_count = 0

    rated_lps = 0
    while len(positions) < TARGET_ROWS:
        inv = build_investor(investor_count, used_names)
        investor_count += 1
        rated_lps += any(r not in ("NR", "") for r in (inv["sp_rating"], inv["moodys_rating"], inv["fitch_rating"]))
        r = random.randint(REPEAT_MIN, REPEAT_MAX)
        r = min(r, TARGET_ROWS - len(positions))      # trim the last LP to land exactly on TARGET
        chosen = weighted_sample_without_replacement(accts, [fac_weights[a] for a in accts], r)
        for acct in chosen:
            commit = money(2_000_000, 500_000_000)
            # Snap uncalled to $100K steps, kept strictly 0 < uncalled < commit.
            uncalled = round(commit * random.uniform(0.15, 0.85) / 100_000) * 100_000
            uncalled = min(max(uncalled, 100_000), commit - 100_000)
            called = commit - uncalled
            row = {
                **inv,
                "acct": acct, "fund": fund_by_acct[acct], "bbdate": fac_bbdate[acct],
                "commit": commit, "called": called, "uncalled_capital": uncalled,
                "called_pct": round(called / commit, 2),
            }
            positions.append(row)
            per_fac[acct].append(row)

    # Second pass: everything that needs the facility's totals — share percentages, then the two
    # excess-concentration figures and the borrowing bases computed net of them.
    #
    # A concentration limit is a fraction of the facility's TOTAL uncalled capital, so an LP's cap is
    # limit x total and its excess is whatever its own uncalled exceeds that cap by. The borrowing
    # base then advances against the uncalled that remains WITHIN the cap — which is why these three
    # columns have to be produced together and after the totals are known. The agent and UBS sides
    # carry their own limits, so they cut at different points and disagree on the same LP.
    for acct, rows in per_fac.items():
        tot_c = sum(r["commit"] for r in rows) or 1
        tot_u = sum(r["uncalled_capital"] for r in rows) or 1
        for r in rows:
            r["pct_commit"] = round(r["commit"] / tot_c, 4)
            r["pct_of_fund_uncalled"] = round(r["uncalled_capital"] / tot_u, 4)
            uncalled = r["uncalled_capital"]
            agent_excess = max(0, round(uncalled - r["agent_cl"] * tot_u))
            ubs_excess = max(0, round(uncalled - r["ubs_cl"] * tot_u))
            r["agent_excess_conc"] = agent_excess
            r["ubs_excess_conc"] = ubs_excess
            r["agent_borrowing_base"] = round(r["agent_ar"] * (uncalled - agent_excess))
            r["ubs_borrowing_base"] = round(r["ubsar"] * (uncalled - ubs_excess))

    # Export rows in SRC_COLS order, as dicts so the chaos monkey can address columns by name.
    export_rows = [dict(zip(SRC_COLS, [
        r["acct"], r["fund"], r["name"], r["parent"], r["spv"], r["ubs_cls"],
        r["inst"], r["ig"], r["cls"], r["sp_rating"], r["moodys_rating"], r["fitch_rating"],
        r["lp_size_bil"], r["lp_size_criteria"], r["commit"], r["uncalled_capital"], r["ubsar"],
        r["agent_cl"], r["ubs_cl"], r["pct_commit"], r["called"], r["pct_of_fund_uncalled"],
        r["called_pct"], r["agent_excess_conc"], r["ubs_excess_conc"],
        r["agent_borrowing_base"], r["ubs_borrowing_base"], "", r["bbdate"],
    ])) for r in positions]

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
    print(f"  facilities (incl orphans) : {len(accts)}  ({len(ORPHAN_ACCOUNTS)} orphan)")
    print(f"  LPs/facility  min/med/max : {fac_sizes[0]} / {fac_sizes[len(fac_sizes)//2]} / {fac_sizes[-1]}")
    print(f"  avg repeats per LP        : {len(positions)/investor_count:.1f}")
    print(f"  LPs with usable ratings   : {rated_lps} ({rated_lps/investor_count:.0%})")
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
