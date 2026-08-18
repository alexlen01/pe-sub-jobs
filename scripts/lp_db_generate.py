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
  * Base values use the canonical Investor Type / Agent LP Category vocabularies from
    data/reference/, so with chaos off the extract's unmatched counts come back zero.
  * Agency-rating presence follows the Agent LP Category (RATING_PRESENCE). classify_ubs derives
    "Rated Investor" from any usable rating, so tying presence to the agent bucket keeps the derived
    UBS mix near the real book's ~40% Rated and leaves unrated LPs to exercise the Corp Pension /
    Unrated NAV / FoF / HNW / Other matrix branches.

Chaos monkey (pe-sub-docs/"AI Chaos Monkey for Data Quality.md"): degrades the values written to
the XLSX, per the analyst "hierarchy of care". Sacred cash/identity columns (CHAOS_SACRED) stay
exact; decision fields (ratings, Investor Type, agent Classification) get formatting and categorical
drift ('A minus', 'SWF', 'PWM'); afterthought fields (investor names, AUM/NAV/PensionAssets) get
suffix drops, ' - Tranche A', case flips, M<->B unit mix-ups and range strings like '500M - 2Bn'.
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
EXPORT_OUT = DATA_DIR / "import" / "LP DB Export 2026.08.17.xlsx"
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

# Must match lp_db_extract.SRC_COLS exactly — the extract validates the header.
SRC_COLS = [
    "AccountID", "FndName", "InvestorName", "Parent", "SPV", "InvestorType", "Region", "HQ",
    "InstitutionalHNW", "InvestmentGrade", "Classification", "Notes", "SP", "Moodys", "Fitch",
    "AUM", "NAV", "PensionAssets", "FundingRatio", "UBSAR", "AgentAR", "Commitments",
    "PercentOfCommitments", "Called", "Uncalled", "PercentOfUncalled", "CalledPercent",
    "AgentCL", "UBSCL", "AgentBB", "UBSBB", "BBDate",
]

# ── reference vocabularies (canonical values, so they map cleanly through the extract) ───────
REGIONS = ["North America", "Europe", "Asia Pacific", "Latin America", "Middle East",
           "United Kingdom", "Africa"]

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

# Agent LP Category (export "Classification") -> (agent advance rate, weight). Canonical per
# data/reference/agent_lp_categories.csv, so each value maps to itself.
AGENT_CATEGORIES = [
    ("Rated Included",            0.90, 30),
    ("Non-Rated Included",        0.75, 34),
    ("Designated Institutional",  0.65, 18),
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
    "aum_scale": 0.10,         # unit mix-ups / style drift on AUM
    "pension_scale": 0.10,     # same on PensionAssets (can flip a Corp Pension classification)
    "nav_text": 0.15,          # NAV -> range/threshold/shorthand strings ('500M - 2Bn', '>5B')
    "investor_type": 0.15,     # categorical drift ('SWF', 'Corporate Pension', 'FoF')
    "agent_category": 0.10,    # agent Classification drift ('Rated', 'PWM', 'Ineligible')
    "ubscl_null": 0.05,        # missing concentration limit
    "funding_ratio_null": 0.08,  # missing pension funding ratio
}

# Never degraded: cash/legal LPA figures plus the facility join keys, which analysts keep exact.
CHAOS_SACRED = ("AccountID", "FndName", "Commitments", "Called", "Uncalled", "BBDate",
                "UBSAR", "AgentAR", "AgentBB", "UBSBB",
                "PercentOfCommitments", "PercentOfUncalled", "CalledPercent")

_NAME_SUFFIX_RE = re.compile(r",?\s+(LLC|L\.L\.C\.|L\.P\.|LP|Ltd\.?|Inc\.?|Limited)$", re.I)

# Canonical value -> manual-entry variants: a mix of alias-resolvable spellings (exercise the
# extract's reference lists) and unknown labels (exercise its unmatched reports).
_ITYPE_DRIFT = {
    "Sovereign Wealth Fund": ["SWF", "Sovereign Wealth"],
    "Fund of Funds": ["FoF", "Fund-of-Funds", "FOF & Other Asset Manager"],
    "Endowment": ["Endowments", "Endowment/Foundation"],
    "Foundation": ["Foundations"],
    "Insurance Company": ["Insurance", "Ins. Co."],
    "Family Office": ["Family Offices", "Single Family Office"],
    "Other Institutional": ["Other Institutional Investors"],
    "Public Pension": ["Public Pension Plan", "Pension - Public"],
    "Pension Fund": ["Corporate Pension", "Pension"],
    "Hedge Fund": ["Hedge Fund Manager"],
    "Corporate": ["Corporate Investor"],
    "Healthcare": ["Healthcare System"],
    "Institutional Investor": ["Institutional"],
    "Investment Consultant": ["Consultant"],
}
_AGENT_DRIFT = {
    "Rated Included": ["Rated", "Rated Included Investors", "rated included"],
    "Non-Rated Included": ["Non Rated Included Investors", "Included Investors", "Non-Rated"],
    "Designated Institutional": ["Designated", "Designated Investors",
                                 "Fund-Of-Fund Designated Investors"],
    "Designated PWM": ["PWM", "Designated - PWM"],
    "Ineligible Investor": ["Ineligible", "Ineligible Investors", "Excluded Investor"],
}
_MONEY_STR_RE = re.compile(r"^\$?(\d+(?:\.\d+)?)\s*([KMBT])$", re.I)
_UNIT_MULT = {"K": 1e3, "M": 1e6, "B": 1e9, "T": 1e12}


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


def _chaos_scale(val: str, rng: random.Random) -> tuple[str, str] | None:
    """Scale drift on '$124.9B'-style strings: M<->B unit swap, style drift ('124.9 bn',
    '$124.9B+'), or raw absolute dollars."""
    m = _MONEY_STR_RE.match(str(val).strip())
    if not m:
        return None
    num, unit = m.group(1), m.group(2).upper()
    options = [("style_naked", f"{num} {unit.lower()}n" if unit in ("B", "T") else f"{num} {unit.lower()}"),
               ("style_plus", f"${num}{unit}+")]
    if unit == "B":
        options.append(("unit_swap", f"${num}M"))
        options.append(("absolute", str(int(float(num) * 1e9))))
    elif unit == "M":
        options.append(("unit_swap", f"${num}B"))
    return rng.choice(options)


def _chaos_nav(val, rng: random.Random) -> tuple[str, str] | None:
    """NAV manual-entry patterns: ranges with shared or mixed units, symbol shorthand, qualitative
    thresholds — derived from the clean value so the result stays plausible."""
    m = _MONEY_STR_RE.match(str(val).strip())
    if not m:
        return None
    dollars = float(m.group(1)) * _UNIT_MULT[m.group(2).upper()]
    if dollars <= 0:
        return None
    mn = dollars / 1e6
    pattern = rng.choice(["range_same_unit", "range_mixed_unit", "symbol_shorthand", "threshold"])
    if pattern == "range_same_unit":
        low, high = round(mn * 0.7), round(mn * 1.3)
        if high >= 1000:
            return pattern, f"{round(low / 1000, 1)}-{round(high / 1000, 1)}B"
        return pattern, f"{low}-{high}M"
    if pattern == "range_mixed_unit":
        if mn > 1000:
            return pattern, f"{round(mn * 0.5)}M - {round(mn * 1.5 / 1000, 1)}Bn"
        return pattern, f"{round(mn)}M"
    if pattern == "symbol_shorthand":
        if mn > 1000:
            return pattern, f"${round(mn * 0.8 / 1000)}-{round(mn * 1.2 / 1000)}Bn"
        return pattern, f"${round(mn)}m"
    unit = "M" if mn < 1000 else "B"
    display = round(mn) if unit == "M" else round(mn / 1000)
    return pattern, f"{rng.choice('><')}{display}{unit}"


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
        if not _blank(row["AUM"]) and rng.random() < CHAOS_RATES["aum_scale"]:
            mutate(row_no, row, "AUM", _chaos_scale(row["AUM"], rng))
        if not _blank(row["PensionAssets"]) and rng.random() < CHAOS_RATES["pension_scale"]:
            mutate(row_no, row, "PensionAssets", _chaos_scale(row["PensionAssets"], rng))
        if not _blank(row["NAV"]) and rng.random() < CHAOS_RATES["nav_text"]:
            mutate(row_no, row, "NAV", _chaos_nav(row["NAV"], rng))
        if rng.random() < CHAOS_RATES["investor_type"]:
            variants = _ITYPE_DRIFT.get(_as_str(row["InvestorType"]))
            if variants:
                mutate(row_no, row, "InvestorType", ("categorical_drift", rng.choice(variants)))
        if rng.random() < CHAOS_RATES["agent_category"]:
            variants = _AGENT_DRIFT.get(_as_str(row["Classification"]))
            if variants:
                mutate(row_no, row, "Classification", ("categorical_drift", rng.choice(variants)))
        if not _blank(row["UBSCL"]) and rng.random() < CHAOS_RATES["ubscl_null"]:
            mutate(row_no, row, "UBSCL", ("nulled", None))
        if not _blank(row["FundingRatio"]) and rng.random() < CHAOS_RATES["funding_ratio_null"]:
            mutate(row_no, row, "FundingRatio", ("nulled", None))
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


def size_label(measure: str) -> str:
    # Log-uniform: fund sizes are log-distributed, and it puts mass on both sides of the matrix
    # boundaries (Corp Pension $1Bn/$5Bn, Unrated NAV $1Bn, FoF & Other $10Bn AUM) so every
    # classify_ubs branch sees traffic.
    def log_uniform(lo: float, hi: float) -> float:
        return 10 ** random.uniform(math.log10(lo), math.log10(hi))
    if measure == "pension_assets":
        return f"${log_uniform(0.8, 120):.1f}B"
    if measure == "nav":
        return f"${log_uniform(0.2, 8):.1f}B"
    return f"${log_uniform(0.3, 60):.1f}B"           # aum


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
    return {
        "name": name,
        "parent": name,
        "spv": "Y" if random.random() < 0.12 else "N",
        "itype": itype,
        "region": random.choice(REGIONS),
        "hq": "Yes" if random.random() < 0.8 else "No",
        "inst": "HNW" if itype == "Family Office" and random.random() < 0.5 else "Institutional",
        "ig": "Yes" if ig else "No",
        "cls": agent_cat,
        "agent_ar": agent_ar,
        "sp_rating": sp, "moodys_rating": mdy, "fitch_rating": fitch,
        # nav-measure LPs (FoF / Hedge Fund) report manager-level AUM alongside fund NAV; the
        # "FoF & Other > $10Bn AUM" branch of classify_ubs needs it.
        "aum": size_label("aum") if measure in ("aum", "nav") else None,
        "nav": size_label("nav") if measure == "nav" else None,
        "pension_assets": size_label("pension_assets") if is_pension else None,
        "funding": round(random.uniform(0.72, 1.08), 2) if is_pension else None,
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
            agent_bb = round(inv["agent_ar"] * uncalled)
            ubs_bb = round(inv["ubsar"] * uncalled)
            row = {
                **inv,
                "acct": acct, "fund": fund_by_acct[acct], "bbdate": fac_bbdate[acct],
                "commit": commit, "called": called, "uncalled_capital": uncalled,
                "called_pct": round(called / commit, 2),
                "agent_borrowing_base": agent_bb, "ubs_borrowing_base": ubs_bb,
            }
            positions.append(row)
            per_fac[acct].append(row)

    # Second pass: per-facility share percentages.
    for acct, rows in per_fac.items():
        tot_c = sum(r["commit"] for r in rows) or 1
        tot_u = sum(r["uncalled_capital"] for r in rows) or 1
        for r in rows:
            r["pct_commit"] = round(r["commit"] / tot_c, 4)
            r["pct_of_fund_uncalled"] = round(r["uncalled_capital"] / tot_u, 4)

    # Export rows in SRC_COLS order, as dicts so the chaos monkey can address columns by name.
    export_rows = [dict(zip(SRC_COLS, [
        r["acct"], r["fund"], r["name"], r["parent"], r["spv"], r["itype"], r["region"], r["hq"],
        r["inst"], r["ig"], r["cls"], "", r["sp_rating"], r["moodys_rating"], r["fitch_rating"],
        r["aum"], r["nav"], r["pension_assets"], r["funding"], r["ubsar"], r["agent_ar"], r["commit"],
        r["pct_commit"], r["called"], r["uncalled_capital"], r["pct_of_fund_uncalled"], r["called_pct"],
        r["agent_cl"], r["ubs_cl"], r["agent_borrowing_base"], r["ubs_borrowing_base"], r["bbdate"],
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
    ws.append(SRC_COLS)
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
