#!/usr/bin/env python3
r"""
Read an LP DB Export workbook and AgentBankSummaryRpt.xlsx, then write the four CSV seed files
used by pe-sub-jobs:
    data/out/lp_master.csv
    data/out/lp_facility_seeds.csv
    data/out/facilities.csv
    data/out/umbrellas.csv

The LP input is the LP DB Export. Columns are matched by header, and common analyst entry
variations are tolerated. A facility is identified by the (AccountID, FndName) pair.

V2 of the format adds Region and Investor Type back to the 30 columns the 2026-08-18 format
carried. Both are read as OPTIONAL, so a V1 workbook - and the platform's own 30-column LP Records
export - still parses; a V1 row simply states neither, which reads downstream as "not resubmitted"
and leaves whatever LP Master already holds untouched.

Usage (no command-line arguments): set EXPORT_FILE below, then run from any directory:
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
AGENT_BANK_SUMMARY_FILE = DATA_DIR / "import" / "AgentBankSummaryRpt.xlsx"
EXPORT_FILE = DATA_DIR / "import" / "LP DB Export V3.xlsx"
REFERENCE_DIR = DATA_DIR / "reference"  # normalization lists
OUT_DIR = DATA_DIR / "out"              # all outputs land here

# --- source columns -------------------------------------------------------------------------
# Internal names for the export's columns, in the order the V2 format lists them. Region and
# Investor Type are the two V2 additions and sit where V2 puts them - after Parent, and after UBS
# LP Classification.
# Columns are located BY HEADER, not by position (see SRC_HEADERS / read_export), so a further
# reshuffle needs no code change - only a new spelling needs one.
SRC_COLS = [
    "AccountID", "FndName", "InvestorName", "Parent", "Region", "SPV", "UbsClassification",
    "InvestorType",
    "InstitutionalHNW", "InvestmentGrade", "Classification", "SP", "Moodys", "Fitch",
    "LpSizeBil", "LpSizeCriteria", "Commitments", "Uncalled", "UBSAR", "AgentAR",
    "AgentCL", "UBSCL",
    "PercentOfCommitments", "Called", "PercentOfUncalled", "CalledPercent",
    "AgentExcessConc", "UBSExcessConc", "AgentBB", "UBSBB", "Notes", "BBDate",
    # The two columns that let a multi-tranche facility STATE that it is one, rather than leaving
    # a reader to recover it from a naming habit. Both optional, so today's V2 workbook - which
    # carries neither - parses unchanged. Neither groups anything (see assign_umbrellas); they are
    # carried so the run report can name the sleeve sets a user has to relate in the platform.
    "Tranche", "TrancheOf",
    # The credit agreement's own reference. Optional for the same reason as the two above: no
    # export written to date carries it, and a column that is not there states nothing about any
    # row. Carried to the platform as the evidence a user relates facilities by - never grouped on.
    "AgreementRef",
]

# Agent Advance Rate is the one column allowed to be ABSENT rather than fatal: workbooks written
# while it was thought to have been dropped do not carry it, and the row is not lost by its absence -
# agent_rate_map.csv resolves the rate from the row's Agent LP Category instead. Every other column
# missing still aborts the run. (AgentCL has the same per-cell fallback but is NOT optional as a
# column: the export has always carried it, so its absence still reads as a malformed workbook.)
# Region and Investor Type join it for a different reason: they are the V2 additions, so every
# workbook written to the 30-column format predates them. Treating their absence as fatal would
# reject the entire back catalogue - and the platform's own LP Records export, which still writes
# the 30-column shape - to gain nothing: a column that is not there states nothing about any row.
#
# Tranche and TrancheOf are optional for a third reason: no export written to date carries either.
# They are the columns the format is expected to gain for multi-tranche support, and reading them
# now means the day one arrives it is reported rather than ignored. Until then every workbook states
# nothing about tranching and the legacy sleeve-suffix reading fills the gap (see
# legacy_sleeve_reading).
OPTIONAL_COLS = {"AgentAR", "Region", "InvestorType", "Tranche", "TrancheOf", "AgreementRef"}

# Accepted header spellings per column. Matching runs through _norm(), which lowercases and
# collapses every run of non-alphanumerics to one space - so it absorbs the format's own quirks
# without needing an entry for each: the embedded CRLF in "LP Size\r\n($ Bil)", the "Insitutional"
# typo, the "Moody'S" capitalisation, and "Investment Grade?"'s trailing question mark.
#
# The lists deliberately cover THREE vocabularies at once, because all three land on this reader:
#   * the current LP DB Export headers (first entry of each list);
#   * the pre-2026-08-18 export's terse headers (FndName, InvestorName, UBSAR, ...), so an archived
#     workbook still parses for the columns it does have;
#   * the readable headers the platform's own LP Records export writes
#     (pe-sub-ui/src/services/lpExportService.ts), so a workbook exported from the UI can be fed
#     straight back in. Keep those in step with that file.
SRC_HEADERS = {
    "AccountID":            ["AccountID", "Account ID"],
    "FndName":              ["FndName", "Fund Name", "Facility Name"],
    "InvestorName":         ["Investor Name", "InvestorName"],
    "Parent":               ["Parent", "Parent / Sponsor"],
    # "Region" is V2's spelling; "Region / Location" is what the pre-2026-08-18 export and the
    # platform's LP Master export both write for the same field.
    "Region":               ["Region", "Region / Location", "Region/Location", "Location"],
    "SPV":                  ["SPV", "SPV Flag"],
    "UbsClassification":    ["UBS LP Classification", "UBS (Internal) LP Classification",
                             "UBS Classification", "UBS LP Category"],
    "InvestorType":         ["Investor Type", "InvestorType", "LP Type"],
    "InstitutionalHNW":     ["Insitutional vs HNW", "Institutional vs HNW", "InstitutionalHNW"],
    "InvestmentGrade":      ["Investment Grade?", "Investment Grade", "InvestmentGrade"],
    "Classification":       ["Agent LP Classification", "Classification", "Agent LP Category"],
    "SP":                   ["S&P", "SP", "S&P Rating"],
    "Moodys":               ["Moody'S", "Moodys", "Moody's Rating"],
    "Fitch":                ["Fitch", "Fitch Rating"],
    "LpSizeBil":            ["LP Size ($ Bil)", "LP Size", "LpSizeBil"],
    "LpSizeCriteria":       ["LP Size Criteria", "Size Measure", "Size Metric Type"],
    "Commitments":          ["Capital Commitments", "Commitments"],
    "Uncalled":             ["Uncalled Capital", "Uncalled"],
    "UBSAR":                ["UBS Advance Rate", "UBSAR", "UBS Advance Rate (%)"],
    "AgentAR":              ["Agent Advance Rate", "AgentAR", "Agent Advance Rate (%)"],
    "AgentCL":              ["Agent Concentration Limit", "AgentCL"],
    "UBSCL":                ["UBS Concentration Limit", "UBSCL"],
    "PercentOfCommitments": ["% of Capital Commitments", "% of Commitments", "PercentOfCommitments"],
    "Called":               ["Called Capital", "Called"],
    "PercentOfUncalled":    ["% of Uncalled Capital", "PercentOfUncalled"],
    "CalledPercent":        ["% of LP Called", "CalledPercent"],
    "AgentExcessConc":      ["Agent Excess Concentration", "Agent Excess Conc Base"],
    "UBSExcessConc":        ["UBS Excess Concentration", "UBS Excess Conc Base"],
    "AgentBB":              ["Agent Borrowing Base", "AgentBB"],
    "UBSBB":                ["UBS Borrowing Base", "UBSBB"],
    "Notes":                ["Notes"],
    "BBDate":               ["BBDate", "Collateral Date", "BB Date"],
    # Which sleeve of a multi-tranche facility this row's facility is. The spellings cover what an
    # agent's system is likely to call the field; the VALUES are normalised separately, by
    # tranche_type(), because the vocabulary drifts far more than the header does.
    "Tranche":              ["Tranche", "Tranche Type", "TrancheType", "Facility Tranche",
                             "Sleeve", "Sleeve Type"],
    # The facility the sleeve belongs to - the credit agreement's own facility, named once and
    # shared by every sleeve of it. Read in preference to matching a name suffix: a currency
    # tranche, a term sleeve or an accordion states its parent here and is reported as the sleeve
    # set it is however it is named, which a suffix match cannot do. It relates nothing on its own -
    # the sleeves hold their own account numbers, and a user relates them in the platform.
    "TrancheOf":            ["Tranche Of", "TrancheOf", "Parent Facility", "ParentFacility",
                             "Facility", "Credit Agreement", "Facility Group"],
    # The reference the credit agreement is administered under - a deal number, a CUSIP-like
    # facility ID, whatever the agent prints over the agreement itself. Distinct from "Credit
    # Agreement" above, which is a NAME and is read as the parent facility of a sleeve.
    "AgreementRef":         ["Agreement Ref", "AgreementRef", "Agreement Reference",
                             "Agreement ID", "AgreementID", "Deal ID", "DealID",
                             "Facility ID", "FacilityID", "Credit Agreement Ref"],
}

# Columns the 2026-08-18 format dropped, kept here only so a stale workbook is diagnosed with a
# useful message instead of a bare "missing column" list. None of them feed the outputs any more:
#   HQ                          - gone from the platform feed entirely.
#   AUM / NAV / PensionAssets   - superseded by LP Size ($ Bil) + LP Size Criteria.
#   FundingRatio                - no longer sourced.
# Investor Type and Region were on this list until V2 brought them back; they are read columns
# again and their old "Region / Location" spelling is an accepted alias above.
RETIRED_HEADERS = {
    "HQ": "High Quality",
    "AUM": "AUM", "NAV": "NAV", "PensionAssets": "Pension Assets",
    "FundingRatio": "Funded Ratio (%)",
}

# --- numeric normalization ------------------------------------------------------------------
# Two vocabularies reach this reader with the SAME headers but differently shaped values. The LP DB
# Export writes rates and shares as fractions (0.154); the platform's own LP Records export writes
# them for a spreadsheet, as percent strings or bare percent numbers ("15.4%", 94) with money as
# display strings ("$428,800,000"). Four headers are identical in both - "% of Uncalled Capital",
# "% of LP Called", "Agent Concentration Limit", "UBS Concentration Limit" - so the header cannot
# tell them apart and a header-keyed rule would divide the export's own fractions by 100.
#
# The shape of the VALUE decides instead, which is unambiguous and idempotent: a share or rate is a
# fraction by definition, so anything carrying a '%' or exceeding 1 is a percent and is scaled down,
# while 0.154 is already correct and left alone.
PERCENT_COLS = {"UBSAR", "AgentAR", "PercentOfCommitments", "PercentOfUncalled",
                "CalledPercent"}
MONEY_COLS = {"Commitments", "Called", "Uncalled", "AgentBB", "UBSBB",
              "AgentExcessConc", "UBSExcessConc"}
# A concentration limit is either a percent of uncalled ("7.5%") or an absolute cap ("$25,000,000")
# in the same column - the '%' sign is what tells them apart; bare numbers pass through for the
# downstream magnitude split to resolve.
LIMIT_COLS = {"AgentCL", "UBSCL"}


def _parsed_number(v) -> "tuple[Decimal, bool] | None":
    """Strip the display formatting off one cell -> (number, carried_a_percent_sign).
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


def normalize_numeric(col: str, v):
    """Bring one numeric cell to the feed's internal shape: fractions for rates and shares, plain
    dollars for money. Unparseable values pass through untouched, same as any other dirty cell."""
    if blank(v) or col not in (PERCENT_COLS | MONEY_COLS | LIMIT_COLS):
        return v
    parsed = _parsed_number(v)
    if parsed is None:
        return v                                   # unparseable: pass through
    number, was_pct = parsed
    if col in PERCENT_COLS:                        # '15.4%' or 94 -> 0.154 / 0.94; 0.154 stays
        return _trim(number / 100 if was_pct or abs(number) > 1 else number)
    if col in LIMIT_COLS:                          # '7.5%' -> 0.075; '$25,000,000' -> 25000000
        return _trim(number / 100 if was_pct else number)
    return _trim(number)                           # money: '$428,800,000' -> 428800000

# CSV header (column) orders required by the pe-sub-jobs FlatFileItemReaders.
# high_quality is gone: the export no longer carries it, and nothing else supplies it. The platform
# keeps its own column on the schema default (TRUE) rather than being fed a fabricated value.
# investor_type and region_location are fed again from V2's two new columns. They still go out
# blank for a V1 workbook that states neither, which is not the same as clearing them: pe-sub-api
# reads a blank as "not resubmitted" and keeps what LP Master already holds.
# funding_ratio stays in the header but goes out BLANK from this feed either way - no column
# implies it and it is governed elsewhere.
MASTER_COLS = [
    "investor_name", "parent", "spv", "investor_type", "institutional_or_hnw",
    "region_location", "investment_grade", "sp_rating", "moodys_rating", "fitch_rating", "aum", "nav", "pension_assets",
    "funding_ratio", "ubs_lp_category", "ubs_default_advance_rate", "ubs_default_concentration_limit",
    "notes",
]
SEED_COLS = [
    "facility_name", "investor_name", "capital_commitment", "uncalled_capital",
    "agent_lp_category", "agent_advance_rate", "agent_concentration_limit",
    "parent", "spv", "investor_type", "institutional_or_hnw", "region_location",
    "investment_grade", "ubs_lp_category", "sp_rating", "moodys_rating", "fitch_rating", "aum", "nav", "pension_assets",
    "funding_ratio", "pct_of_fund_commitments", "called_capital", "pct_of_fund_uncalled", "pct_lp_called",
    "ubs_concentration_limit", "ubs_advance_rate", "agent_excess_concentration",
    "ubs_excess_concentration", "agent_borrowing_base", "ubs_borrowing_base", "notes",
]
# The shape a facility row is BUILT in, which is wider than the shape it is FED in.
#
# tranche_type and tranche_of are working columns: which sleeve of a multi-tranche facility a row
# is, and the facility it is a sleeve OF, as either file states it or as the legacy
# "(Committed)"/"(Uncommitted)" suffix spells it. They no longer GROUP anything - grouping is by
# account number alone, see assign_umbrellas - and they are kept for one purpose: the run report
# lists the sleeve sets it did not group, so whoever maps them in the platform is told they are
# there rather than left to find them.
#
# They are NOT fed. Which sleeve is which is not a fact the platform holds, and a sleeve set is
# related in the platform by a person, so stating it on the feed would put a column in the contract
# that nothing on the far side reads.
FACILITY_WORK_COLS = [
    "agent_bank", "name", "account_number", "loan_amount", "maturity_date", "status",
    # The date the agent last restated that status. Written to hold the column's position: the
    # platform records a facility's standing, not when it was last restated, and dropping the column
    # would shift every column after it in feeds already in production.
    "status_date", "ubs_participation", "collateral_date", "umbrella_name",
    "umbrella_key", "cross_collateralized",
    "tranche_type", "tranche_of",
    # ── the agreement reference ──────────────────────────────────────────────────────────────
    # What the agreement is administered under, where either file prints it - carried through to
    # the platform because it is the best evidence a person has when relating two facilities by
    # hand, and stamped on a group where its members agree on one.
    #
    # It does NOT group facilities that hold their own account numbers. An agent re-papers a
    # reference on a renewal, quotes the wrong one, or prints one reference over two agreements a
    # merger left it administering; acting on it would move a facility's whole borrowing base into
    # a pool on this script's reading of a string. Two accounts under one reference are related in
    # the platform, by a user, and the reference is what they read to do it.
    #
    # ADDITIVE: appended, never inserted, and blank on every file written to date.
    "agreement_ref",
]
TRANCHE_TYPE_COL = FACILITY_WORK_COLS.index("tranche_type")
TRANCHE_OF_COL = FACILITY_WORK_COLS.index("tranche_of")
AGREEMENT_REF_COL = FACILITY_WORK_COLS.index("agreement_ref")
# The member columns the group layer reads back off its own members, for a group the report named
# no row for and so stated nothing about directly.
LOAN_AMOUNT_COL = FACILITY_WORK_COLS.index("loan_amount")
MATURITY_COL = FACILITY_WORK_COLS.index("maturity_date")
COLLATERAL_COL = FACILITY_WORK_COLS.index("collateral_date")

# What actually goes out on the facility feed, and the columns of the working row it is taken from.
FACILITY_COLS = [c for c in FACILITY_WORK_COLS if c not in ("tranche_type", "tranche_of")]
FACILITY_EMIT_IDX = [FACILITY_WORK_COLS.index(c) for c in FACILITY_COLS]
# The group layer, fed separately from its members. A group is not a facility - it has no LP roster
# and no borrowing base of its own - so stating it on the member rows would repeat one credit
# agreement's terms once per fund and leave whichever row loaded last holding them.
#
# The agreement's own loan terms ARE written here. They were once left out on the grounds that the
# member facilities carry them and one fact should be stated once; that reading was wrong about which
# row the fact belongs to. The platform now governs a member from its group - an Active group hands
# down its account number, line, maturity, collateral date and standing - so a group fed without them
# lands as Not Stated and governs nothing until somebody types the terms the file already carried.
UMBRELLA_COLS = [
    "key", "name", "agent_bank", "account_number", "loan_amount",
    "cross_collateralized",
    # agreement_ref is what the platform resolves the group by where it is stated, so a group
    # survives its account being re-papered. Blank means the file stated nothing, which the ingest
    # reads as silence: it never clears a reference an analyst recorded.
    #
    # obligor_name, borrower_entity, sub_limit and liability_type were written here and have been
    # removed. None of the four appears in the Agent Bank Summary or the LP Data Export, none is an
    # input to a Shadow BB, and no user asked for any of them - they were extracted from what a
    # credit agreement contains rather than from what the platform is fed. A column nobody
    # populates does not stay empty; it gets filled in by whoever assumes it must matter.
    "agreement_ref",
    # ADDITIVE: appended, never inserted, so a reader positioned on the older header goes on
    # reading the columns before these unchanged.
    #
    # The agreement's maturity and the date its collateral was last certified, both taken from what
    # the files state about the group - never invented, and blank where they state nothing.
    "maturity_date", "collateral_date",
    # The group's own standing, which is the switch that makes it govern its members. Always stated:
    # a group exists here only because several funds draw on one account with live LP rosters
    # beneath it, and that is a live credit agreement. See group_status for the one exception.
    "facility_status",
]

# One account number carrying MORE THAN ONE facility is an umbrella subscription facility: several
# related funds, feeders or SPVs borrowing under one credit agreement, which is why the agent
# reports them against a single account. Each member keeps its own LP roster and its own borrowing
# base; the umbrella is the layer above them.
#
# The grouping is decided HERE, and stated on the facility row, rather than left to the ingest to
# notice: the ingest reads the feed in chunks and never has the whole file in view, so it cannot
# see that two facilities several hundred rows apart share an account. This script does.
#
# The report prints an umbrella one of two ways, and they need different handling:
#
#   * one row per MEMBER fund, all on the shared account. Each row joins to its own fund by name,
#     and the group itself is never named - so its name is minted from the account number, the only
#     thing the two files then say about it.
#   * one row for the GROUP, naming the obligor that signs the credit agreement - "Carlyle Buyout
#     Umbrella" over an account the export carries six funds on. Here the report DOES state the
#     group's name and terms, and the row is not a facility at all: it has no LP roster, and its
#     loan amount is the whole agreement's, not any one fund's.
#
# The second shape is what a real report carries most often, and it is why an unmatched row on a
# multi-fund account is never handed to a member (see upsert_facilities): doing so puts the group's
# name and the group's loan amount onto one arbitrary fund and leaves its siblings orphaned.
#
# The ingest keys the umbrella on `umbrella_key` - the account number - rather than on the name, so
# an analyst renaming it to what the credit agreement calls it is safe and survives every later run.
UMBRELLA_NAME_PREFIX = "Umbrella "

# ── The other shape of one credit agreement: a multi-tranche facility ─────────────────────────
#
# Reported as one sleeve per tranche, because the sleeves are the SAME borrower on the SAME
# collateral, split by how much of the commitment is contractually firm - so the agent certifies
# each separately and gives each ITS OWN ACCOUNT NUMBER. "<Fund> (Committed)" beside "<Fund>
# (Uncommitted)", Tranche A beside Tranche B.
#
# THIS SCRIPT DOES NOT GROUP THEM, and that is a decision rather than an omission. Every signal
# available here is circumstantial: a name suffix someone typed, or an agreement reference an agent
# re-papers on renewal and mis-quotes often enough to matter. Grouping on either moves a facility's
# entire borrowing base under an allocated share of a collateral pool it may not be part of, and
# does it silently, over a whole book at a time. Two facilities holding two account numbers are two
# facilities here; relating them is a decision, it is made by a user in the platform, and it is
# recorded there as one.
#
# What this section still does is READ the evidence, so the platform's user is handed it rather
# than left to notice. `Tranche` / `TrancheOf` are taken from either file wherever they state them,
# the legacy sleeve suffix fills them where neither does, and the run report lists what came out -
# the sets someone has to map, named, at the end of every run.
#
# A shared ACCOUNT NUMBER is the one relationship that is grouped, in assign_umbrellas: it is not
# an inference at all, but the agent's own statement that these funds draw on one agreement.

# Canonical tranche types, matching the platform's own vocabulary (tranches.tranche_type). The
# spellings an agent's system uses for the same thing are folded onto them; anything unrecognised
# is passed through upper-cased so it reaches the platform as a visible rejection rather than being
# silently re-labelled as something the platform does handle.
TRANCHE_TYPE_ALIASES = {
    "committed": "COMMITTED", "committed tranche": "COMMITTED", "firm": "COMMITTED",
    "uncommitted": "UNCOMMITTED", "uncommitted tranche": "UNCOMMITTED", "discretionary": "UNCOMMITTED",
    "term": "TERM", "term loan": "TERM", "term tranche": "TERM",
    "accordion": "ACCORDION", "incremental": "ACCORDION",
    "delayed draw": "DELAYED_DRAW", "delayeddraw": "DELAYED_DRAW", "ddtl": "DELAYED_DRAW",
}

# The suffix the legacy feeds print a sleeve under. Confined to legacy_sleeve_reading below, which
# is the only caller: it is a fallback reader for a stated fact, not the way tranching is modelled.
LEGACY_SLEEVE_SUFFIX_RE = re.compile(r"\s*\((committed|uncommitted)\)\s*$", re.I)


def tranche_type(value) -> str:
    """One stated tranche value, folded onto the platform's vocabulary. Blank stays blank."""
    raw = as_is(value).strip()
    if not raw:
        return ""
    return TRANCHE_TYPE_ALIASES.get(_norm(raw), raw.upper().replace(" ", "_"))


def legacy_sleeve_reading(name: str) -> "tuple[str, str] | None":
    """(tranche_type, tranche_of) recovered from a legacy sleeve name, or None where it is not one.

    Applied only to a facility whose feed states no tranche of its own, and the result groups
    nothing: it fills two columns the run report reads back out, so a sleeve set the export names
    only by suffix is still reported to whoever maps it. The suffix is what makes a name a sleeve of
    something larger; a name without one is a whole facility, and "Fund X" standing beside "Fund X
    (Committed)" is a separate credit agreement until someone says otherwise."""
    m = LEGACY_SLEEVE_SUFFIX_RE.search(name or "")
    if not m:
        return None
    base = LEGACY_SLEEVE_SUFFIX_RE.sub("", name).strip()
    return (tranche_type(m.group(1)), base) if base else None


# Agent Bank Summary column layout (must match the report header exactly). Index 4 is an unnamed
# spacer holding the report's subtotal amounts.
ABS_COLS = [
    "Agent", "Borrower", "AccountNumber", "LoanAmount", "", "MaturityDate",
    "FacilityStatus", "FacilityStatusDate",
]
# Columns the report may carry AFTER the eight above, located by header rather than by position so
# they may arrive in either order or not at all. The Agent Bank Summary is the facility-level file,
# which makes it the natural place for an agent to state which sleeve of a facility a row is - and
# it is the only file that carries a facility the export has no LPs for. Absent on every report
# written to date; see OPTIONAL_COLS for why that is read as "states nothing" rather than as an
# error.
ABS_OPTIONAL_HEADERS = {
    # The agreement reference an agent prints over a GROUP row rather than over a fund.
    "AgreementRef":   ["Agreement Ref", "AgreementRef", "Agreement Reference", "Agreement ID",
                       "AgreementID", "Deal ID", "DealID", "Facility ID", "FacilityID"],
    "Tranche":   ["Tranche", "Tranche Type", "TrancheType", "Sleeve", "Sleeve Type"],
    "TrancheOf": ["Tranche Of", "TrancheOf", "Parent Facility", "ParentFacility",
                  "Credit Agreement", "Facility Group"],
}
ABS_TOTAL_MARKER = "accesstotalsloanamount"  # _norm() prefix of the subtotal / grand-total rows
# _norm() of the FacilityStatus that onboards a reported facility as Active on the report's word
# alone, with no export match behind it. Every other spelling reads as not-Active.
ABS_ACTIVE_STATUS = "active"


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


def clean_name(value) -> str:
    """Trim a person or organization name for stored LP records.

    Casing is carried exactly as fed: the source spelling of a name is the authoritative one, and
    re-casing it turns acronyms and internal capitals into something the fed name never said."""
    name = as_is(value)
    if not name:
        return ""
    return re.sub(r"\s+", " ", name)


@dataclass
class Reference:
    agent_categories: set[str]              # norm(canonical Agent LP Classification)
    ubs_lookup: dict[str, str]              # norm(alias) -> canonical UBS LP Classification
    itype_lookup: dict[str, str]            # norm(alias or canonical) -> canonical Investor Type
    agent_rates: dict[str, float]           # norm(Agent LP Category) -> advance rate percent
    agent_conc_limits: dict[str, float]     # norm(Agent LP Category) -> conc limit percent
    rate_floors: list[tuple[float, float]]  # (min_rate_pct, group_pct), sorted highest-min first
    rating_notches: dict[str, dict[str, str]]  # agency -> {rating key -> canonical notch}


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
    # Investor Type, back as a column in V2. The canonical list mirrors the platform's
    # INVESTOR_TYPE_OPTS, so a normalized value is one the LP Master screen's dropdown already
    # offers; the alias list carries the spellings the banks write for those same types. Canonicals
    # are seeded into the lookup first so an alias row can never shadow one.
    itype_canonical = [r[0] for r in _read_reference_rows(ref_dir / "investor_types.csv")[1:] if r[0]]
    itype_lookup = {_norm(c): c for c in itype_canonical}
    for row in _read_reference_rows(ref_dir / "investor_type_aliases.csv")[1:]:
        if len(row) >= 2 and row[1] and _norm(row[0]) not in itype_lookup:
            itype_lookup[_norm(row[0])] = row[1]

    ubs_rows = _read_reference_rows(ref_dir / "ubs_lp_categories.csv")[1:]
    ubs_lookup = {_norm(r[0]): r[1] for r in ubs_rows if len(r) >= 2 and r[1]}

    # Agent advance rate AND agent concentration limit by Agent LP Category: the fallbacks for rows
    # whose Agent Advance Rate / Agent Concentration Limit cell is blank, so a value-less row still
    # gets what its category implies. The two columns are loaded independently — a row may legally
    # carry a rate with no limit (short reference row), and the missing one simply has no fallback.
    agent_rates: dict[str, float] = {}
    agent_conc_limits: dict[str, float] = {}
    for row in _read_reference_rows(ref_dir / "agent_rate_map.csv")[1:]:
        if len(row) < 2:
            continue
        try:
            agent_rates[_norm(row[0])] = float(str(row[1]).rstrip("%"))
        except ValueError:
            pass
        if len(row) < 3:
            continue
        try:
            agent_conc_limits[_norm(row[0])] = float(str(row[2]).rstrip("%"))
        except ValueError:
            continue

    # Agency notch scales, per agency. The three do NOT share one: S&P and Fitch write AA-/BBB+
    # where Moody's writes Aa3/Baa1, and pe-sub-api bands a notch under the column it arrived in, so
    # each agency is normalized against its own ladder and never against another's.
    rating_notches: dict[str, dict[str, str]] = {}
    for row in _read_reference_rows(ref_dir / "rating_scales.csv")[1:]:
        if len(row) < 3 or not row[0] or not row[2]:
            continue
        rating_notches.setdefault(row[0].lower(), {})[_rating_key(row[2])] = row[2]

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

    # The Agent Advance Rate Schedule states the canonical Agent LP Classifications.
    return Reference(set(agent_rates), ubs_lookup, itype_lookup, agent_rates, agent_conc_limits,
                     rate_floors, rating_notches)


def check_agent_cls(raw, ref: Reference) -> tuple[str, bool]:
    """(the value exactly as fed, is-canonical). The value is NEVER rewritten - not into a
    canonical spelling, not into anything else. Only the flag changes.

    Canonical means the Agent Advance Rate Schedule states it. Recognition is case- and
    punctuation-insensitive (_norm), so "designated - pwm" still prices at PWM Designated Investors' rate
    and limit while being seeded exactly as typed. A classification the schedule does not state
    gets no rate or limit fallback and is counted for manual correction."""
    s = as_is(raw)
    if not s:
        return "", True
    return s, _norm(s) in ref.agent_categories


# --- agency ratings -------------------------------------------------------------------------
# A rating cell is typed by hand, so it arrives spelled every way a notch can be spelled: 'A minus'
# for A-, 'Baa 1' for Baa1, 'aa+' for AA+. pe-sub-api's BbCriteriaResolver looks the notch up in
# bb_criteria_matrix.ratingBands EXACTLY, and a notch it cannot find is not treated as missing - it
# is treated as present-but-unbanded and clamped to the BBB floor. A dirty spelling therefore does
# not lose the rating, it silently downgrades it, moving the LP's advance rate and concentration
# limit. Normalizing here is what keeps a spelling difference from becoming a pricing difference.
NOT_RATED_TOKENS = {"nr", "na", "nm", "none", "notrated", "norating", "unrated"}


def _rating_key(v) -> str:
    """A notch's identity, independent of how it was typed: case, spacing, punctuation and the
    spelled-out modifiers all collapse. 'A minus' / 'a -' / 'A-' -> 'a-'; 'Baa 1' -> 'baa1'."""
    t = re.sub(r"[\s.]+", " ", str(v or "").strip().lower())
    t = t.replace(" minus", "-").replace(" plus", "+")
    return re.sub(r"\s+", "", t)


def normalize_rating(raw, agency: str, ref: Reference) -> tuple[str, str]:
    """(value to write, outcome) for one agency rating cell.

    Outcomes: 'blank' (nothing fed), 'canonical' (matched the agency's scale, written canonically),
    'not_rated' (an NR / N/R / N/A token - written BLANK, because "this agency does not rate this
    LP" is the absence of a rating, and a token in the column reads to pe-sub-api as a rating it
    cannot band), or 'unmatched' (a notch the agency's scale does not list - written exactly as fed
    and reported, never guessed at)."""
    s = as_is(raw)
    if not s:
        return "", "blank"
    if re.sub(r"[^a-z]", "", s.lower()) in NOT_RATED_TOKENS:
        return "", "not_rated"
    canon = ref.rating_notches.get(agency, {}).get(_rating_key(s))
    return (canon, "canonical") if canon else (s, "unmatched")


def normalized_ratings(sp, moodys, fitch, ref: Reference) -> tuple[dict[str, str], list[str]]:
    """The three cells normalized together -> ({sp/moodys/fitch: value}, [outcome, ...])."""
    pairs = [("sp", sp), ("moodys", moodys), ("fitch", fitch)]
    results = {a: normalize_rating(v, a, ref) for a, v in pairs}
    return {a: r[0] for a, r in results.items()}, [r[1] for r in results.values()]


# --- UBS classification ---------------------------------------------------------------------
# The 2026-08-18 export states the UBS LP Classification outright, so it is no longer derived here.
# The waterfall this replaced inferred it from agency ratings, pension assets, NAV, AUM and the
# HNW/SPV flags - four of which the format no longer carries - so deriving it is neither possible
# nor wanted: the LP DB is the system of record for this field.
def map_ubs_cls(raw, ref: Reference) -> tuple[str, bool]:
    """(canonical UBS LP Classification, True) when the fed value maps to one of the nine classes;
    otherwise the ORIGINAL value with (value, False). The record is always kept - an unrecognised
    class is reported, never dropped or silently rewritten."""
    s = as_is(raw)
    if not s:
        return "", True
    canon = ref.ubs_lookup.get(_norm(s))
    return (canon, True) if canon else (s, False)


# --- investor type --------------------------------------------------------------------------
# Distinct from both LP Category (the bank's borrowing-base risk bucket) and LP Classification (the
# regulatory status): this is the industry/sector profile - pension, endowment, SWF - and the three
# are never interchangeable. Normalizing matters here because the LP Master screen's Investor Type
# filter is built from the DISTINCT values in the table, so every unmapped spelling of one type
# becomes another entry in that dropdown.
def map_investor_type(raw, ref: Reference) -> tuple[str, bool]:
    """(canonical Investor Type, True) when the fed value is one the platform states or a known
    spelling of one; otherwise the ORIGINAL value with (value, False). Same contract as
    map_ubs_cls - the record is always kept and an unrecognised type is reported, never dropped."""
    s = as_is(raw)
    if not s:
        return "", True
    canon = ref.itype_lookup.get(_norm(s))
    return (canon, True) if canon else (s, False)


# --- LP size ---------------------------------------------------------------------------------
# "LP Size ($ Bil)" + "LP Size Criteria" replace the old AUM / NAV / PensionAssets trio. The
# criteria column names which measure the figure is, using the same vocabulary as the platform's
# LP_SIZE_CRITERIA_OPTS ("AUM", "NAV", "Assets"), so the value is routed back into whichever of the
# three schema columns it belongs to. That keeps pe-sub-api's contract and the UI's Size Measure
# derivation (aum ? 'AUM' : nav ? 'NAV' : pension_assets ? 'Assets') working unchanged.
# Keyed by _norm(), so case and punctuation are already absorbed. The spelled-out variants are here
# because the criteria cell is analyst-typed free text in practice ("Total AUM", "Net Asset Value"),
# and an unrecognised label costs the row its whole LP Size - the figure has no meaning without a
# basis to attribute it to.
SIZE_CRITERIA_COL = {
    "aum": "aum", "total aum": "aum", "assets under management": "aum",
    "nav": "nav", "net asset value": "nav", "fund nav": "nav",
    "assets": "pension_assets", "total assets": "pension_assets",
    "pension assets": "pension_assets", "pension": "pension_assets",
}

_SIZE_UNIT_MULT = {"": 1.0, "k": 1e-6, "m": 1e-3, "mn": 1e-3, "mm": 1e-3,
                   "b": 1.0, "bn": 1.0, "t": 1e3, "tn": 1e3, "trn": 1e3}


def parse_size_bil(v) -> float | None:
    """Tolerant parser for the LP Size column -> billions of dollars. The column's unit is $Bn, so a
    bare number is already billions ('13.5' -> 13.5) - unlike the old AUM/NAV free text, where a
    bare number meant absolute dollars. An explicit unit still wins when the analyst typed one
    ('$13.5B', '900M', '1.2 bn', '>5B'), and a range takes its LOW end ('5-8' -> 5).
    None when nothing numeric reads."""
    if blank(v):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().lower().replace(",", "").replace("$", "")
    s = s.lstrip("<>~ ").rstrip("+ ")
    s = re.sub(r"\.(?!\d)", " ", s)          # stray dots ('1.33. bn') - keep decimal points only
    parts = re.findall(r"(\d+(?:\.\d+)?)\s*([a-z]*)", s)
    if not parts:
        return None
    units = [u for _, u in parts if u in _SIZE_UNIT_MULT and u]
    default_unit = units[-1] if units else ""   # '5-8bn': the shared unit applies to both ends
    vals = []
    for num, unit in parts:
        mult = _SIZE_UNIT_MULT.get(unit) if unit in _SIZE_UNIT_MULT else _SIZE_UNIT_MULT.get(default_unit, 1.0)
        vals.append(float(num) * mult)
    return min(vals) if vals else None


def size_display(v) -> str:
    """One LP Size cell as the short-currency display string the aum/nav/pension_assets columns hold
    ('13.5' -> '$13.5B'). Unparseable text passes through verbatim, so a dirty cell stays visible
    for review instead of becoming a wrong number."""
    bil = parse_size_bil(v)
    if bil is None:
        return as_is(v)
    return money_short(Decimal(str(bil)) * Decimal("1e9"))


def size_columns(row: dict) -> dict[str, str]:
    """Route one row's LP Size into the aum / nav / pension_assets column its criteria names.
    An unrecognised or blank criteria leaves all three blank rather than guessing a measure -
    the figure without its basis is not attributable to any of them."""
    out = {"aum": "", "nav": "", "pension_assets": ""}
    col = SIZE_CRITERIA_COL.get(_norm(row["LpSizeCriteria"]))
    if col:
        out[col] = size_display(row["LpSizeBil"])
    return out


# --- agent advance rate ----------------------------------------------------------------------
def agent_rate_frac(raw_rate, raw_category, ref: Reference) -> float | None:
    """The agent advance rate for a row, as a fraction. Read from the export's Agent Advance Rate
    column, which normalize_numeric has already brought to a fraction; the fed value wins outright,
    dirty or not, because the export is the system of record for it. Only a blank or unparseable
    cell falls back to the rate the row's Agent LP Category implies (agent_rate_map.csv), and only
    when that category is one the schedule states.
    None when neither has anything - a made-up rate would be indistinguishable from a fed one."""
    if not blank(raw_rate):
        try:
            return float(raw_rate)
        except (TypeError, ValueError):
            pass                                   # unparseable: fall back to the category
    canon, canonical = check_agent_cls(raw_category, ref)
    if not canonical or not canon:
        return None
    pct_value = ref.agent_rates.get(_norm(canon))
    return None if pct_value is None else pct_value / 100


def agent_conc_limit_frac(raw_limit, raw_category, ref: Reference) -> float | None:
    """The agent concentration limit for a row, as a fraction. Mirrors agent_rate_frac: the fed
    Agent Concentration Limit cell wins outright, and only a blank or unparseable cell falls back
    to the limit the row's Agent LP Category implies (agent_rate_map.csv). None when
    neither has anything - a made-up limit would be indistinguishable from a fed one.

    normalize_numeric has already brought the cell to LIMIT_COLS shape, where a percent-of-uncalled
    is a fraction ('7.5%' -> 0.075) but an ABSOLUTE cap stays in dollars ('$25,000,000' -> 25000000)
    in the very same column. Only the fed value can be an absolute cap, so it is returned untouched
    and the caller renders it exactly as before; the fallback is always a percent, by definition of
    the reference file."""
    if not blank(raw_limit):
        try:
            return float(raw_limit)
        except (TypeError, ValueError):
            pass                                   # unparseable: fall back to the category
    canon, canonical = check_agent_cls(raw_category, ref)
    if not canonical or not canon:
        return None
    pct_value = ref.agent_conc_limits.get(_norm(canon))
    return None if pct_value is None else pct_value / 100


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
    header = ["" if h is None else str(h) for h in next(rows_iter)]
    # Columns are addressed by NAME, not position, so the 2026-08-18 reshuffle needed no change
    # here and a further one will not either. Matching goes through _norm(), which lowercases and
    # collapses runs of non-alphanumerics - that is what absorbs the header quirks the format ships
    # with, notably the embedded CRLF in "LP Size\r\n($ Bil)", the "Insitutional" typo and the
    # trailing "?" on "Investment Grade?". Anything unrecognised is ignored, so an extra trailing
    # column never breaks the read.
    header_to_col = {_norm(alias): col for col, aliases in SRC_HEADERS.items() for alias in aliases}
    column_at: dict[str, int] = {}
    for i, name in enumerate(header):
        col = header_to_col.get(_norm(name))
        if col is not None:
            column_at.setdefault(col, i)
    missing = [c for c in SRC_COLS if c not in column_at and c not in OPTIONAL_COLS]
    if missing:
        # A stale pre-2026-08-18 workbook fails on the columns that format did not carry, so name
        # them and say which retired ones are present - far more useful than a bare missing-column
        # list. Region and Investor Type are NOT among them: they are optional, so their absence
        # never reaches here and a V1 workbook parses instead of being diagnosed.
        found_retired = [h for h in header
                         if _norm(h) in {_norm(v) for v in RETIRED_HEADERS.values()}]
        hint = ""
        if found_retired:
            hint = (f"\n  This looks like a PRE-2026-08-18 export: it still carries "
                    f"{found_retired}, which the current format dropped. Re-export it, or run an "
                    f"older revision of this script against it.")
        raise SystemExit(
            f"""Export header in '{ws.title}' is missing {len(missing)} of the {len(SRC_COLS)} expected columns: {missing}
  Accepted header spellings per column are listed in SRC_HEADERS near the top of this script.
  found: {header}{hint}"""
        )
    rows = []
    for r in rows_iter:
        row = {c: (r[column_at[c]] if c in column_at and column_at[c] < len(r) else None)
               for c in SRC_COLS}
        # Rates/shares arrive as fractions from the LP DB Export and as percents from the platform's
        # own export under several identical headers; normalize_numeric decides on the value, not
        # the header, and is idempotent for values already in the feed's shape.
        for c in PERCENT_COLS | MONEY_COLS | LIMIT_COLS:
            row[c] = normalize_numeric(c, row[c])
        rows.append(row)
    return rows


def read_agent_bank_summary(path: Path) -> tuple[list[list[str]], dict[str, int]]:
    """Read the Agent Bank Summary report into FACILITY_WORK_COLS-shaped rows.
    Returns (rows, account_number -> row index).

    The report is a banded print layout, not a flat table:
      * the agent bank appears once on its own group-header row (Agent set, Borrower blank) and is
        carried down onto the facility rows beneath it, which leave the Agent cell blank;
      * each group ends with an 'AccessTotalsLoanAmount:' subtotal row carrying no facility.

    Dirty rows are absorbed rather than fatal: a repeated (AccountNumber, Borrower) pair is a
    reprint and is dropped; one AccountNumber listed against two borrowers yields two facilities.
    Borrower names are carried exactly as printed - two accounts may share one, and separating
    them is upsert_facilities' final pass, which sees the placeholder facilities too.

    ubs_participation and collateral_date are not in the report — collateral_date is filled from
    the export's BBDate by upsert_facilities. bank_status is the report's own FacilityStatus,
    carried as printed; upsert_facilities normalises it and decides the final value."""
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
    # Anything past the eight fixed columns is optional and located by header. An unrecognised
    # trailing column is ignored, exactly as in the export reader, so the report gaining a column
    # this script does not know about is never an error.
    abs_header_to_col = {_norm(alias): col
                         for col, aliases in ABS_OPTIONAL_HEADERS.items() for alias in aliases}
    abs_optional_at: dict[str, int] = {}
    for i, name in enumerate(header[len(ABS_COLS):], start=len(ABS_COLS)):
        col = abs_header_to_col.get(_norm(name))
        if col is not None:
            abs_optional_at.setdefault(col, i)

    data: list[list[str]] = []
    by_acct: "OrderedDict[str, list[tuple[int, str]]]" = OrderedDict()
    seen_pair: set[tuple[str, str]] = set()
    agent = ""

    width = max([len(ABS_COLS)] + [i + 1 for i in abs_optional_at.values()])
    for raw in rows_iter:
        full = (list(raw) + [None] * width)[:width]
        cells = full[: len(ABS_COLS)]
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
        printed = _norm(name)                             # for the FndName join below
        # The row's own Agent cell wins if the report fills it; otherwise the carried-down header.
        # "Unknown" satisfies FacilityRowProcessor's non-blank agent_bank rule.
        # The trailing blanks are ubs_participation, collateral_date, and the three umbrella
        # columns - none of which the report states. collateral_date is filled in by
        # upsert_facilities and the umbrella columns by assign_umbrellas.
        #
        # The last two are the tranche declaration, taken from the report where it states one. A
        # report that does not is left blank here and resolved later, from the export or from the
        # legacy sleeve reading, by declare_tranches.
        def _opt(col: str) -> str:
            i = abs_optional_at.get(col)
            return as_is(full[i]).strip() if i is not None else ""
        data.append([text[0] or agent or "Unknown", name, acct, text[3], iso_date(cells[5]),
                     text[6], iso_date(cells[7]), "", "", "", "", "",
                     tranche_type(_opt("Tranche")), _opt("TrancheOf"), _opt("AgreementRef")])
        # Every borrower on the account, in report order: an account listed against two borrowers
        # is two facilities and both stay eligible for the LP join.
        by_acct.setdefault(acct, []).append((len(data) - 1, printed))

    return data, by_acct


def facility_key(row: dict) -> tuple[str, str]:
    """Identity of the facility an export row belongs to: the (AccountID, FndName) pair.

    Neither column identifies a facility on its own. One FndName repeats across every investor on
    a facility and can be reported under two accounts; one AccountID can carry several FndNames.
    Both cases are two facilities, and the pair keeps them apart.

    Either half may be blank - the pair is taken as given rather than dropped, and rows blank the
    same way group into one facility. FndName is normalised so case and punctuation drift within
    one account does not split a facility in two."""
    return ((row["AccountID"] or "").strip(), _norm(as_is(row["FndName"])))


def _recent_first(rows: list[dict]) -> list[dict]:
    """An investor's rows ordered by BB run date, most recent first. A later submission supersedes
    an earlier one, so this ordering - not a headcount - decides every consolidated attribute.

    sorted() is stable and stays stable under reverse=True, so rows sharing a BBDate keep their
    export order and the earliest-listed of them wins. Rows with no BBDate sort last: an undated
    submission never outranks a dated one."""
    return sorted(rows, key=lambda r: iso_date(r["BBDate"]) or "", reverse=True)


def _latest(rows: list[dict], field: str) -> str:
    """Consolidate one field across an investor's rows, which MUST already be _recent_first():
    the value from the most recent submission that actually reported it.

    A blank on a newer row means 'not resubmitted this cycle', not 'cleared' - ratings and
    concentration limits are routinely omitted from a submission that still carries a current
    investor type - so the search falls through to the next-most-recent non-blank value rather
    than letting a blank erase a known one. '' only when every row is blank."""
    for r in rows:
        v = as_is(r[field])
        if v:
            return v
    return ""


def build_master(export: list[dict], ref: Reference) -> list[dict]:
    """One golden LP per investor_name, consolidated by recency: each attribute is taken from the
    most recent BB run that reported it (see _recent_first / _latest). A later submission both
    improves on data quality and carries the attributes that are current - investor type, LP
    category, ratings - so it supersedes earlier ones field by field rather than being outvoted by
    a mass of older rows.

    Investor Type is mapped to canonical AFTER the recency pick, so an older row cannot supply the
    label. UBS classification is derived from the recency-consolidated attributes."""
    groups: "OrderedDict[str, list[dict]]" = OrderedDict()
    for row in export:
        name = clean_name(row["InvestorName"])
        if not name:
            continue
        groups.setdefault(name, []).append(row)

    master_rows: list[dict] = []
    for name, rows in groups.items():
        rows = _recent_first(rows)
        # LP Size and its criteria are consolidated as a PAIR from the same submission: a figure
        # taken from one BB run and a measure label from another would mislabel the number. The
        # search falls through to the next-most-recent submission that actually RESOLVES - a row
        # whose criteria drifted to a label no alias covers yields nothing, and stopping there would
        # blank the LP's size even though an older row states it perfectly well.
        size = {"aum": "", "nav": "", "pension_assets": ""}
        for r in rows:
            routed = size_columns(r)
            if any(routed.values()):
                size = routed
                break

        master_rows.append({
            "investor_name": name,
            "parent": clean_name(_latest(rows, "Parent")),
            "spv": yn_bool(_latest(rows, "SPV")),
            # Both fed by V2's new columns, and both blank for a V1 workbook that has neither.
            # Blank means "not resubmitted", which pe-sub-api treats as "keep what LP Master
            # already holds" - not as a clearing edit. funding_ratio below is blank either way: no
            # column implies it in either version.
            "investor_type": map_investor_type(_latest(rows, "InvestorType"), ref)[0],
            "institutional_or_hnw": _latest(rows, "InstitutionalHNW"),
            # Free text, deliberately: region is a label the bank writes, not a governed
            # vocabulary here, so it is carried as fed rather than forced onto a list.
            "region_location": _latest(rows, "Region"),
            "investment_grade": yn_bool(_latest(rows, "InvestmentGrade")),
            "sp_rating": normalize_rating(_latest(rows, "SP"), "sp", ref)[0],
            "moodys_rating": normalize_rating(_latest(rows, "Moodys"), "moodys", ref)[0],
            "fitch_rating": normalize_rating(_latest(rows, "Fitch"), "fitch", ref)[0],
            "aum": size["aum"],
            "nav": size["nav"],
            "pension_assets": size["pension_assets"],
            "funding_ratio": "",
            "ubs_lp_category": map_ubs_cls(_latest(rows, "UbsClassification"), ref)[0],
            "ubs_default_advance_rate": floor_rate_frac(_latest(rows, "UBSAR"), ref),
            # The two columns beside each other are on DIFFERENT scales, and deliberately so: the
            # advance rate is stored as a fraction (0.90), while a concentration limit is stored on
            # the percent-or-dollars encoding a limit shares with the LP record it defaults - a
            # percent of total uncalled (25) or an absolute dollar cap (25000000), told apart by
            # magnitude. The export states both as fractions, so the limit is converted here exactly
            # as it is for the seed rows; passing 0.25 through would be read back as 0.25%.
            "ubs_default_concentration_limit": pct(_latest(rows, "UBSCL")),
            "notes": _latest(rows, "Notes"),
        })

    return master_rows


@dataclass
class SeedResult:
    rows: list[dict]
    counts: Counter
    anomalies: list[str]


def build_seed(export: list[dict], name_by_key: dict[tuple[str, str], str],
               ref: Reference) -> SeedResult:
    """One seed row per export row, carrying every per-LP export column. NOTHING is dropped:
    the export is the system of record and every LP it lists must reach the platform.

    The (AccountID, FndName) pair is the identity of a facility here, and each pair owns a
    distinct facility name - which is what keeps these rows resolvable downstream, where a facility
    is looked up by name alone and a (facility, investor) pair that already exists is skipped.

    Three conditions are recorded as anomalies - a blank investor name, a blank AccountID, and the
    same investor listed twice on one facility - but their rows are written anyway and reported by
    main(). The third is the one that costs a record downstream.

    Classification and Investor Type are normalised (unmatched passed through); ubs_lp_category is
    derived per row from that row's own attributes."""
    seen: set[tuple[str, str]] = set()
    seed_rows: list[dict] = []
    counts = Counter()
    anomalies: list[str] = []

    for row in export:
        acct = (row["AccountID"] or "").strip()
        investor = clean_name(row["InvestorName"])
        # Every key is in the map: upsert_facilities manufactures a placeholder facility for every
        # facility_key the Agent Bank Summary does not report, so this cannot fail.
        fkey = facility_key(row)
        fac_name = name_by_key[fkey]
        if not investor:
            counts["blank_investor"] += 1
            anomalies.append(f"blank investor name on facility {fac_name!r} "
                             f"(account {acct or 'blank'})")
        if not acct:
            counts["blank_account"] += 1
            anomalies.append(f"no AccountID for investor {investor!r} "
                             f"(placed in {fac_name!r} by fund name)")
        key = (fkey, investor)
        if key in seen:
            counts["duplicate_facility_investor"] += 1
            anomalies.append(f"facility {fac_name!r} (account {acct or 'blank'}) lists investor "
                             f"{investor!r} more than once - each occurrence is seeded")
        seen.add(key)

        ratings, rating_outcomes = normalized_ratings(row["SP"], row["Moodys"], row["Fitch"], ref)
        counts["stripped_not_rated"] += rating_outcomes.count("not_rated")
        counts["unmatched_rating"] += rating_outcomes.count("unmatched")

        agent_cls, agent_matched = check_agent_cls(row["Classification"], ref)
        ubs_cls, ubs_matched = map_ubs_cls(row["UbsClassification"], ref)
        itype, itype_matched = map_investor_type(row["InvestorType"], ref)
        if not agent_matched:
            counts["unmatched_agent_category"] += 1
        if not itype_matched:
            counts["unmatched_investor_type"] += 1
        if not ubs_matched:
            counts["unmatched_ubs_classification"] += 1
        if not as_is(row["UbsClassification"]):
            counts["blank_ubs_classification"] += 1
        size = size_columns(row)
        agent_rate = agent_rate_frac(row["AgentAR"], row["Classification"], ref)
        if agent_rate is None:
            counts["unresolved_agent_rate"] += 1
        agent_conc = agent_conc_limit_frac(row["AgentCL"], row["Classification"], ref)
        if agent_conc is None:
            counts["unresolved_agent_conc_limit"] += 1

        seed_rows.append({
            "facility_name": fac_name,
            "investor_name": investor,
            "capital_commitment": money_short(row["Commitments"]),
            "uncalled_capital": money_short(row["Uncalled"]),
            "agent_lp_category": agent_cls,
            # Fed by the export (or, for a blank cell, resolved from the Agent LP Category) and
            # NOT floor-mapped. The floor map exists to slot a UBS rate into the bank's own
            # 90/75/65/50/0 groups; the agent's schedule is not those groups, so flooring this
            # would silently turn Insitutional Designated Investors' 60% into 50%.
            "agent_advance_rate": pct(agent_rate),
            # Same rule as the rate above: fed by the export, or - for a blank cell - resolved from
            # the Agent LP Category. Never floor-mapped; the floor map is an advance-rate device.
            "agent_concentration_limit": pct(agent_conc),
            "parent": clean_name(row["Parent"]),
            "spv": yn_bool(row["SPV"]),
            "investor_type": itype,
            "institutional_or_hnw": as_is(row["InstitutionalHNW"]),
            "region_location": as_is(row["Region"]),
            "investment_grade": yn_bool(row["InvestmentGrade"]),
            "ubs_lp_category": ubs_cls,
            "sp_rating": ratings["sp"],
            "moodys_rating": ratings["moodys"],
            "fitch_rating": ratings["fitch"],
            "aum": size["aum"],
            "nav": size["nav"],
            "pension_assets": size["pension_assets"],
            "funding_ratio": "",
            "pct_of_fund_commitments": pct(row["PercentOfCommitments"]),
            "called_capital": money_short(row["Called"]),
            "pct_of_fund_uncalled": pct(row["PercentOfUncalled"]),
            "pct_lp_called": pct(row["CalledPercent"]),
            "ubs_concentration_limit": pct(row["UBSCL"]),
            "ubs_advance_rate": floor_rate_pct(row["UBSAR"], ref),
            "agent_excess_concentration": money_short(row["AgentExcessConc"]),
            "ubs_excess_concentration": money_short(row["UBSExcessConc"]),
            "agent_borrowing_base": money_short(row["AgentBB"]),
            "ubs_borrowing_base": money_short(row["UBSBB"]),
            "notes": as_is(row["Notes"]),
        })
        counts["written"] += 1

    return SeedResult(seed_rows, counts, anomalies)


@dataclass
class StatedGroup:
    """A report row that names a CREDIT AGREEMENT rather than a facility.

    One row on an account the export carries several funds on, naming none of them: the agent has
    printed the obligor that signs and draws, and the funds beneath it only in the export. The row's
    terms are the agreement's - the loan amount is the whole facility's, the status and maturity are
    the group's - so they are fed at group level and never onto a member."""
    account: str
    name: str                # the printed Borrower: the credit agreement's own name
    agent_bank: str
    loan_amount: str
    maturity_date: str
    bank_status: str         # as printed; normalised by the caller
    bank_status_date: str
    # The agreement's own reference, where the report printed it over the group row. Blank on a
    # report that printed only the eight standard columns, which is every report to date - and
    # blank is what the ingest reads as "not stated".
    agreement_ref: str = ""


def upsert_facilities(fac_data: list[list[str]], by_acct: dict[str, list[tuple[int, str]]],
                      export: list[dict]
                      ) -> tuple[list[list[str]], dict[tuple[str, str], str],
                                 dict[str, StatedGroup]]:
    """Facilities from the Agent Bank Summary, joined to the export by (AccountID, FndName).
    bank_status := Active when EITHER the report states the facility is Active, OR an export
    facility claims the report row (which also gets collateral_date := that facility's most
    recent BBDate). Inactive otherwise.

    The report is the bank's own record of what it lends against, so a facility it reports as
    Active is onboarded Active whether or not the export happens to carry LPs for it: a live
    facility with no LP records yet - newly originated, or an export cycle that simply did not
    include it - is an empty facility, not a closed one. The export match only ever promotes a
    status; it never demotes one the report calls Active.

    The join runs per account, because the two files spell facilities differently:

      0. an account the export carries SEVERAL funds on, whose single report row names none of them,
         is an umbrella printed at group level: that row is the credit agreement, not a facility. It
         is taken out of the join as a StatedGroup, its members are all built from the export, and
         each inherits the group's agent bank, status and maturity - but not its loan amount, which
         is the whole agreement's. Handing that row to a member instead would give one arbitrary
         fund the group's name and the group's loan amount, and orphan the other five;
      1. a report row whose Borrower is the export's FndName owns that facility, which is what
         gives each fund on a shared account its own report row, on a report that prints them;
      2. whatever is left over is handed out in report order, so a facility the two files name
         differently still resolves. On a one-borrower account this pass is the entire join.

    An export facility the report does not list is manufactured as a placeholder Inactive facility
    with agent bank "Unknown", carrying what the export knows - name=FndName,
    account_number=AccountID, collateral_date=last BB date - so its LP records still seed. A member
    of a stated group is manufactured the same way but is neither Unknown nor Inactive: the report
    states its agent and its standing at group level, which is where an umbrella's are stated.

    Facility names are made unique here, because the seed feed resolves a facility by name alone
    and the platform keys facilities by name: two facilities sharing one would send a facility's
    LPs to the other. A facility keeps its printed name unless another facility would share it;
    when several would, EVERY one of them is suffixed with its account number. Suffixing all of
    them rather than only the later ones is what makes the pair visible as a pair - otherwise one
    facility silently keeps the bare name and its sibling reads as a variant of it.

    Returns (rows, facility_key -> resolved facility name, account -> stated group)."""
    # A facility's Last BB date is the most recent BB run across its rows: the export is not
    # guaranteed to arrive in date order. Insertion order stays the export's first-seen order, so
    # the placeholder facilities below keep a stable, reproducible ordering.
    bbdate_by_key: "OrderedDict[tuple[str, str], str]" = OrderedDict()
    fnd_by_key: "OrderedDict[tuple[str, str], str]" = OrderedDict()
    acctno_by_key: "OrderedDict[tuple[str, str], str]" = OrderedDict()
    # The tranche a facility's export rows state it is. Taken from the first row that states one
    # rather than the first row outright: the columns are per-LP in the file but per-FACILITY in
    # meaning, so a blank on one investor's row says nothing about the facility, and letting it win
    # would drop a declaration the rest of the roster carries.
    tranche_by_key: dict[tuple[str, str], tuple[str, str]] = {}
    # The agreement reference a facility's export rows state, taken the same way and for the same
    # reason: the column is per-LP in the file but per-FACILITY in meaning.
    ref_by_key: dict[tuple[str, str], str] = {}
    for row in export:
        key = facility_key(row)
        bbdate = iso_date(row["BBDate"])
        stated_type = tranche_type(row.get("Tranche"))
        stated_of = as_is(row.get("TrancheOf")).strip()
        if (stated_type or stated_of) and key not in tranche_by_key:
            tranche_by_key[key] = (stated_type, stated_of)
        stated_ref = as_is(row.get("AgreementRef")).strip()
        if stated_ref and key not in ref_by_key:
            ref_by_key[key] = stated_ref
        if key not in bbdate_by_key:
            bbdate_by_key[key] = bbdate
            fnd_by_key[key] = as_is(row["FndName"])
            acctno_by_key[key] = key[0]
        else:
            if bbdate > bbdate_by_key[key]:   # ISO dates compare lexicographically; '' loses
                bbdate_by_key[key] = bbdate
            if not fnd_by_key[key]:
                fnd_by_key[key] = as_is(row["FndName"])

    out = [list(r) for r in fac_data]

    # --- join: claim report rows for export facilities, account by account --------------------
    keys_by_acct: "OrderedDict[str, list[tuple[str, str]]]" = OrderedDict()
    for key in bbdate_by_key:
        keys_by_acct.setdefault(key[0], []).append(key)

    claimed: dict[tuple[str, str], int] = {}
    groups: dict[str, StatedGroup] = {}
    group_rows: set[int] = set()
    for acct, keys in keys_by_acct.items():
        free = list(by_acct.get(acct, []))          # [(row index, normalised Borrower), ...]
        # pass 0 - the report row that names the credit agreement rather than a fund on it.
        # Recognised only where it cannot be anything else: the export carries SEVERAL funds on the
        # account, the report prints ONE row for it, and that row names none of them. One fund on
        # the account is the ordinary rename case and stays with pass 2; two report rows and three
        # funds is ambiguous, and an ambiguous row is left as a facility rather than promoted on a
        # guess - demoting a facility to a group would take its LP roster with it.
        if len(keys) > 1 and len(free) == 1 and free[0][1] not in {k[1] for k in keys}:
            idx = free.pop(0)[0]
            cells = out[idx]
            # The reference is read off THIS row and no other, so one stated over a single member
            # fund is never read as the whole agreement's.
            groups[acct] = StatedGroup(acct, cells[1], cells[0], cells[3], cells[4],
                                       cells[5], cells[6],
                                       agreement_ref=cells[AGREEMENT_REF_COL].strip())
            group_rows.add(idx)
        for key in keys:                            # pass 1 - exact FndName == Borrower
            if not key[1]:
                continue
            for pos, (idx, printed) in enumerate(free):
                if printed == key[1]:
                    claimed[key] = idx
                    free.pop(pos)
                    break
        for key in keys:                            # pass 2 - positional, in report order
            if key in claimed or not free:
                continue
            claimed[key] = free.pop(0)[0]

    # Status is per report row: the report's own FacilityStatus, promoted to Active by an export
    # match. A borrower the report does NOT call Active and the export never lists is Inactive -
    # not Active because a sibling on the same account number is.
    key_by_row = {idx: key for key, idx in claimed.items()}
    for i, row in enumerate(out):
        if i in group_rows:
            continue                                # not a facility; fed as an umbrella instead
        key = key_by_row.get(i)
        row[5] = "Active" if (key is not None or _norm(row[5]) == ABS_ACTIVE_STATUS) else "Inactive"
        if key is not None and bbdate_by_key[key]:
            row[8] = bbdate_by_key[key]

    # The group rows come out before anything is named or counted, so a credit agreement never
    # competes with its own member funds for a facility name and never reaches the facility feed.
    if group_rows:
        kept = [i for i in range(len(out)) if i not in group_rows]
        remap = {old: new for new, old in enumerate(kept)}
        out = [out[i] for i in kept]
        claimed = {key: remap[idx] for key, idx in claimed.items()}

    # --- orphan export facilities -> placeholders ---------------------------------------------
    # A key with a blank AccountID lands here too, with a blank account_number - which is the only
    # way such a row reaches the platform.
    row_by_key: dict[tuple[str, str], int] = dict(claimed)
    for key in bbdate_by_key:
        if key in row_by_key:
            continue
        acctno = acctno_by_key[key]
        name = fnd_by_key[key] or (f"Unknown Facility {acctno}" if acctno
                                   else "Unknown Facility (no account)")
        # A member of a stated group is not an orphan: the report does carry it, at the level the
        # agent reports the agreement on. It takes the group's agent bank and maturity, and NOT its
        # loan amount - that figure is what the agreement lends, not what this fund borrows, and
        # stamping it on each member would state the same money once per fund. Active for the same
        # reason a claimed row is: the export carries a live LP roster against it. The agreement's
        # own printed standing is fed at group level, where the agent stated it.
        group = groups.get(acctno)
        if group is not None:
            # The agreement reference is carried down to a member the report never printed a row
            # for: it is a fact about the agreement the fund borrows under, and without it the
            # member would group by account while its printed siblings group by reference.
            out.append([group.agent_bank or "Unknown", name, acctno, "", group.maturity_date,
                        "Active", group.bank_status_date, "", bbdate_by_key[key], "", "", "",
                        "", "", group.agreement_ref])
        else:
            out.append(["Unknown", name, acctno, "", "", "Inactive", "", "", bbdate_by_key[key],
                        "", "", "", "", "", ""])
        row_by_key[key] = len(out) - 1

    # --- the tranche declaration -------------------------------------------------------------
    # The export's stated tranche, over every facility the export reaches - the claimed report rows
    # and the placeholders alike. The report is asked first (read_agent_bank_summary already filled
    # these where it stated them) because it is the facility-level file and states a facility the
    # export may carry no LPs for; the export fills what the report left blank.
    for key, idx in row_by_key.items():
        stated = tranche_by_key.get(key)
        if stated is None:
            continue
        if not out[idx][TRANCHE_TYPE_COL]:
            out[idx][TRANCHE_TYPE_COL] = stated[0]
        if not out[idx][TRANCHE_OF_COL]:
            out[idx][TRANCHE_OF_COL] = stated[1]

    # The same order of precedence for the agreement reference: the report is the facility-level
    # file and is asked first; the export fills what it left blank. Neither file stating one leaves
    # the column blank, and the facility groups by its account exactly as it always has.
    for key, idx in row_by_key.items():
        stated_ref = ref_by_key.get(key)
        if stated_ref and not out[idx][AGREEMENT_REF_COL].strip():
            out[idx][AGREEMENT_REF_COL] = stated_ref

    # Whatever neither file stated, read from the legacy sleeve name. This runs BEFORE the
    # uniquifying pass below, so the reading sees the name as printed rather than one already
    # suffixed with an account number - "Fund X (Committed) (ACC-1)" carries no readable sleeve
    # suffix, and a set spelt that way would silently stop grouping.
    for row in out:
        if row[TRANCHE_TYPE_COL] or row[TRANCHE_OF_COL]:
            continue                              # the feed said so; nothing to infer
        reading = legacy_sleeve_reading(row[1].strip())
        if reading is not None:
            row[TRANCHE_TYPE_COL], row[TRANCHE_OF_COL] = reading

    # --- unique names ------------------------------------------------------------------------
    # One pass over report rows and placeholders together, so a placeholder is weighed against the
    # reported facilities and not just against the placeholders written before it.
    shared = {n for n, c in Counter(_norm(r[1]) for r in out).items() if c > 1}
    used_norm: set[str] = set()
    for row in out:
        name = row[1].strip()
        if _norm(name) in shared:
            acctno = row[2].strip()
            name = f"{name} ({acctno})" if acctno else f"{name} (no account)"
        # The account number separates facilities across accounts; the ordinal covers what it
        # cannot - two funds on ONE account whose names normalise the same, and blank accounts.
        if _norm(name) in used_norm:
            base, n = name, 1
            while _norm(name) in used_norm:
                n += 1
                name = f"{base} #{n}"
        used_norm.add(_norm(name))
        row[1] = name

    name_by_key = {key: out[idx][1] for key, idx in row_by_key.items()}
    return out, name_by_key, groups


@dataclass
class UmbrellaGroup:
    """One account number carrying more than one facility, as this run found it.

    That shared account is the agent's own statement that these funds draw on one credit agreement,
    and it is the only relationship this script groups on - see assign_umbrellas for what it
    deliberately leaves to a user."""
    key: str                    # the account number the members share - what the ingest resolves by
    name: str                   # the report's, where it states one; minted otherwise
    members: list[str]          # member facility names, in run order
    # Whether the members stand on ONE borrowing base. Never set here: it is a term of the credit
    # agreement, neither source file states it, and it is an analyst's to record. The field is kept
    # because the group feed carries the column, and False writes it blank.
    cross_collateralized: bool
    # What the report states about the AGREEMENT, present only where it printed a group row. Blank
    # on a group this script inferred from a shared account number, which states nothing of itself.
    agent_bank: str = ""
    account_number: str = ""
    loan_amount: str = ""
    # The agreement's own reference, where the report printed it over the group row. It is what the
    # platform resolves this group by where it is stated - the key above is then the account, or the
    # sleeves' shared base name, and is only a fallback.
    agreement_ref: str = ""
    # The agreement's own loan terms, which are what the platform governs its members from once the
    # group is Active. Read off the printed group row where the report named one, and off the
    # members the report DID print where it named none - see assign_umbrellas.
    maturity_date: str = ""
    collateral_date: str = ""
    # Not Stated is never written: see UMBRELLA_COLS.
    facility_status: str = "Active"


def assign_umbrellas(rows: list[list[str]],
                     stated: dict[str, StatedGroup] | None = None) -> list[UmbrellaGroup]:
    """Stamp `umbrella_name`, `umbrella_key` and `cross_collateralized` on every facility that
    belongs to a group, and return the groups for the run report.

    ONE relationship groups, and it is not an inference: several facilities reported against the
    SAME ACCOUNT NUMBER are an umbrella subscription facility - separate funds, feeders or SPVs
    borrowing under one credit agreement, which is why the agent administers them on a single
    account. Each member keeps its own LP roster and its own borrowing base; the umbrella is the
    layer above them.

    Facilities holding their OWN account numbers are never grouped here, however plainly they read
    as one agreement - the sleeves of a multi-tranche facility, or two accounts an agent prints one
    agreement reference over. The evidence for those is a name suffix or a reference that gets
    re-papered, and acting on it would move a whole borrowing base into a pool on a string match.
    They are reported at the end of the run and related in the platform, by a user, as the decision
    it is. That is why nothing below reads `tranche_of` or groups on `agreement_ref`.

    Runs over the FINAL facility rows - the reported ones and the manufactured placeholders
    together - because a group is a fact about the credit agreement, and a member the report omits
    is still a member. Running it over the report alone would leave a group half-declared.

    A blank account number never groups. Blank is the absence of an account, not an account that
    several facilities happen to hold in common, and pooling every accountless facility into one
    umbrella would invent a credit agreement out of missing data.

    `stated` carries the group rows the report printed by account (see upsert_facilities). Where an
    account has one, the group takes the name the agent printed and the terms beside it, because
    that name IS the credit agreement's; the account-number form is the fallback for a report that
    named no group, not the preferred answer.

    The agreement's TERMS go out with it - account number, line, maturity, collateral date and
    standing - because the platform governs a member fund from its group, and a group fed without
    them arrives Not Stated and governs nothing. Read off the printed group row where there is one
    and off the members where there is not, always from what a file states and never invented: a
    figure the members disagree on is left blank rather than picked between.

    Cross-collateralization is deliberately NOT inferred. Whether an account's members' commitments
    support one common borrowing base is a term of the credit agreement; neither source file states
    it, and a guess would read downstream as the legal position. The column is left blank, which the
    ingest reads as "not stated" and never as "no".
    """
    stated = stated or {}
    taken_names = {_norm(row[1]) for row in rows}
    groups: list[UmbrellaGroup] = []

    def agreed(idxs: list[int], col: int) -> str:
        """The value a set of member rows AGREES on, or blank where they disagree.

        Disagreement is left blank rather than resolved. Two different figures over one group is a
        contradiction in the file, and picking either would state the group's terms from one member
        and let it govern the rest on that reading. Blank is what the ingest takes as silence, and
        the disagreement stays visible on the facility rows where it was printed."""
        vals = {rows[i][col].strip() for i in idxs}
        vals.discard("")
        return vals.pop() if len(vals) == 1 else ""

    def agreed_ref(idxs: list[int]) -> str:
        """The agreement reference a set of member rows AGREES on, or blank.

        Blank where they disagree, and deliberately so: two references over one group is a
        contradiction in the file, and picking either would resolve the group onto an agreement
        half its members were never said to be under. Blank leaves the group resolving by its key,
        which is the account number, and the disagreement stays visible in the facility rows."""
        return agreed(idxs, AGREEMENT_REF_COL)

    def latest(idxs: list[int], col: int) -> str:
        """The most recent ISO date a set of member rows carries, or blank.

        The collateral date is the one group-level term the members are NOT expected to agree on:
        each is certified against its own borrowing base, on the date its own LP roster was last
        stated. The agreement is collateralized as of the last of them - an earlier date would put
        the group behind evidence it already holds. ISO dates compare lexicographically."""
        return max((rows[i][col].strip() for i in idxs if rows[i][col].strip()), default="")

    def group_status(printed: "StatedGroup | None") -> str:
        """The standing to feed for a group. Active unless the report says otherwise.

        A group is formed here only where several funds are reported against ONE account, and a
        member reaches the feed Active only where the export carries a live LP roster against it -
        so an account with a group on it is an agreement being drawn on. Feeding Not Stated would
        leave every extracted agreement inert until somebody opened it and typed what the file
        already said.

        The one thing that overrides that is the agent's own word: where the report printed a group
        row AND gave it a standing that is not Active, that standing is fed instead. A blank status
        on a printed row is silence, not a denial, and stays Active."""
        stated_status = printed.bank_status.strip() if printed is not None else ""
        if stated_status and _norm(stated_status) != ABS_ACTIVE_STATUS:
            return "Inactive"
        return "Active"

    # Account umbrellas, in the order the facilities were built, so a group is named and
    # reported from the first row that stated it whichever run reads the file.
    by_acct: "OrderedDict[str, list[int]]" = OrderedDict()
    for i, row in enumerate(rows):
        acct = row[2].strip()
        if acct:
            by_acct.setdefault(acct, []).append(i)

    for acct, idxs in by_acct.items():
        if len(idxs) < 2:
            continue
        printed = stated.get(acct)
        # The printed name unless a facility of its own already answers to it: the two must stay
        # apart on screen, and the group is the one that yields, since the facility's name is what
        # its LP records resolve by.
        if printed is not None and printed.name and _norm(printed.name) not in taken_names:
            name = printed.name
        elif printed is not None and printed.name:
            name = f"{UMBRELLA_NAME_PREFIX}{printed.name}"
        else:
            name = f"{UMBRELLA_NAME_PREFIX}{acct}"
        for i in idxs:
            rows[i][9] = name
            rows[i][10] = acct
            rows[i][11] = ""
        groups.append(UmbrellaGroup(
            acct, name, [rows[i][1] for i in idxs], False,
            agent_bank=printed.agent_bank if printed is not None else rows[idxs[0]][0],
            account_number=acct,
            # The whole agreement's line, which is a group-level figure: it is stated once, over
            # every member fund, and no member borrows it alone.
            #
            # Where the report printed no group row it is read off the members, and only where they
            # AGREE: funds on one account repeating one figure are each being shown the agreement's
            # line, which is that line. Funds carrying different figures are being shown their own
            # allocations, and summing those would state a line no file does. Blank then, and the
            # platform derives the group's size from its members' shares instead.
            loan_amount=(printed.loan_amount if printed is not None
                         else agreed(idxs, LOAN_AMOUNT_COL)),
            # The reference the group's members agree on, carried so the platform holds what the
            # agent stated about the agreement. The group still resolves by its account number:
            # the reference is evidence for a person, not a key this script groups on.
            agreement_ref=(printed.agreement_ref if printed is not None and printed.agreement_ref
                           else agreed_ref(idxs)),
            # The agreement's own terms, fed so the group governs its members on arrival rather
            # than landing inert for somebody to retype. The maturity is the agreement's and the
            # members agree on it by construction where one was printed; the collateral date is the
            # last of the members', which is when the agreement was last certified as a whole.
            maturity_date=(printed.maturity_date if printed is not None
                           else agreed(idxs, MATURITY_COL)),
            collateral_date=latest(idxs, COLLATERAL_COL),
            facility_status=group_status(printed)))

    return groups


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
            # Padded as well as trimmed: a short row would write a ragged line the ingest reads
            # one column out of step from the header. Then projected onto the fed columns, which
            # drops the working ones the grouping used and the platform does not hold.
            padded = (r + [""] * len(FACILITY_WORK_COLS))[: len(FACILITY_WORK_COLS)]
            w.writerow([padded[i] for i in FACILITY_EMIT_IDX])


def write_umbrellas(path: Path, groups: list[UmbrellaGroup]) -> None:
    """The group feed. Written even when empty, so a run that found no group states that rather
    than leaving the last run's file to be read as this run's."""
    with path.open("w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh, quoting=csv.QUOTE_ALL)
        w.writerow(UMBRELLA_COLS)
        for g in groups:
            w.writerow([g.key, g.name, g.agent_bank, g.account_number,
                        g.loan_amount, "true" if g.cross_collateralized else "",
                        g.agreement_ref, g.maturity_date, g.collateral_date,
                        g.facility_status])


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
    fac_rows, name_by_key, stated_groups = upsert_facilities(fac_data, by_acct, export)
    # After the names are final, so the umbrella reports the members under the names the platform
    # will know them by rather than the ones the report printed.
    umbrellas = assign_umbrellas(fac_rows, stated_groups)
    master_rows = build_master(export, ref)
    sr = build_seed(export, name_by_key, ref)

    # A seed row resolves to a facility by name, so two facilities sharing one would redirect an
    # entire facility's LP records into the other.
    dup_names = sorted({n for n in name_by_key.values()
                        if list(name_by_key.values()).count(n) > 1})
    if dup_names:
        raise SystemExit(
            "FACILITY NAME COLLISION: these names resolve to more than one "
            f"(AccountID, FndName) facility, so their LP records would be merged: {dup_names}"
        )

    # Clear data/out/ so it holds only this run's four CSVs. Done after all inputs are read, so a
    # failed read leaves the last good outputs in place. import/ and reference/ are never touched.
    for old in out_dir.iterdir():
        if old.is_file():
            old.unlink()

    write_csv(out_dir / "lp_master.csv", MASTER_COLS, master_rows)
    write_csv(out_dir / "lp_facility_seeds.csv", SEED_COLS, sr.rows)
    write_facilities(out_dir / "facilities.csv", fac_rows)
    write_umbrellas(out_dir / "umbrellas.csv", umbrellas)

    # Retention is an invariant, not a metric: every export row must appear in
    # lp_facility_seeds.csv. A mismatch is a bug in this script, so it fails the run.
    # An Active facility the export carries no LPs for is onboarded empty, so it is called out
    # separately: it is an expected state (newly originated, or simply not in this export cycle),
    # not a dropped roster.
    seeded_names = {r["facility_name"] for r in sr.rows}
    active = [r for r in fac_rows if r[5] == "Active"]
    empty_active = [r for r in active if r[1].strip() not in seeded_names]
    print(f"export rows            : {len(export)}")
    print(f"lp_facility_seeds rows : {sr.counts['written']}")
    print(f"lp_master rows         : {len(master_rows)} (one per distinct investor name)")
    print(f"facilities             : {len(fac_rows)} "
          f"({len(active)} active from the report, {len(fac_rows) - len(active)} inactive)")
    if empty_active:
        print(f"active, no LP records  : {len(empty_active)} facility(ies) the report states are "
              "Active that this export carries no LPs for; each is onboarded Active and empty")
        for r in empty_active[:20]:
            print(f"  {r[2] or '(no account)':<12}: {r[1]}")
        if len(empty_active) > 20:
            print(f"  ... and {len(empty_active) - 20} more")
    print(f"export facilities      : {len(name_by_key)} "
          f"(distinct AccountID+FndName pairs over {len({k[0] for k in name_by_key})} accounts)")

    # An account carrying more than one facility is reported rather than failed: it is a normal
    # structure, not a data fault. Members keep their own LP records - a group groups them, it does
    # not merge them.
    if umbrellas:
        print()
        print(f"umbrella facilities    : {len(umbrellas)} account(s) carry more than one facility; "
              "each is fed as one umbrella and its funds as members")
        for g in umbrellas[:20]:
            print(f"  {g.name:<40}: {', '.join(g.members)}")
        if len(umbrellas) > 20:
            print(f"  ... and {len(umbrellas) - 20} more")
        named = [g for g in umbrellas if g.loan_amount]
        if named:
            print(f"  {len(named)} of them {'is' if len(named) == 1 else 'are'} printed at GROUP "
                  "level - the report states the agreement's own line and none of its member "
                  "funds. That line is fed as the umbrella's; the funds come from the export.")
        print("  cross-collateralization is a term of the credit agreement, is stated by neither "
              "source file, and is left for an analyst to set")

    # What this run did NOT group, and hands to a user instead. Facilities holding their own
    # account numbers are never related here (see assign_umbrellas), so the sleeve sets and the
    # shared agreement references are listed by name: they are exactly the mapping work the
    # platform is now expecting someone to do, and an unlisted one is one nobody knows about.
    sleeve_sets: "OrderedDict[str, list[str]]" = OrderedDict()
    for r in fac_rows:
        parent = r[TRANCHE_OF_COL].strip()
        if parent and not r[9].strip():             # declared a sleeve, and grouped by no account
            sleeve_sets.setdefault(_norm(parent), []).append(
                f"{r[1]}{f' [{r[TRANCHE_TYPE_COL]}]' if r[TRANCHE_TYPE_COL] else ''}")
    ref_sets: "OrderedDict[str, set[str]]" = OrderedDict()
    for r in fac_rows:
        ref = r[AGREEMENT_REF_COL].strip()
        if ref:
            ref_sets.setdefault(ref, set()).add(r[2].strip())
    multi_ref = {ref: accts for ref, accts in ref_sets.items() if len(accts) > 1}
    if sleeve_sets or multi_ref:
        print()
        print("left to map manually   : facilities that read as one credit agreement but hold "
              "their OWN account numbers. This script does not group them - the evidence is a name "
              "suffix or a re-papered reference - so they are related in the platform, by a user.")
    for parent, sleeves in list(sleeve_sets.items())[:20]:
        if len(sleeves) > 1:
            print(f"  tranche sleeves        : {', '.join(sleeves)}")
    if multi_ref:
        for ref, accts in list(multi_ref.items())[:20]:
            print(f"  agreement {ref:<12}: accounts {', '.join(sorted(accts))}")

    # Normalization outcomes. These are not failures - the value is written through unchanged - but
    # a non-zero count means a reference list is behind the feed and should be topped up.
    norm_counts = [
        ("unmatched Agent LP Category", sr.counts["unmatched_agent_category"],
         "not stated by the Agent Advance Rate Schedule: seeded as fed, set it by hand"),
        ("unmatched UBS classification", sr.counts["unmatched_ubs_classification"],
         "ubs_lp_categories.csv"),
        ("unmatched investor type", sr.counts["unmatched_investor_type"],
         "not a stated Investor Type: seeded as fed, investor_type_aliases.csv"),
        ("blank UBS classification", sr.counts["blank_ubs_classification"],
         "fed empty by the export"),
        ("unresolved agent advance rate", sr.counts["unresolved_agent_rate"],
         "blank in the export and category unmapped: agent_rate_map.csv"),
        ("unresolved agent conc limit", sr.counts["unresolved_agent_conc_limit"],
         "blank in the export and category unmapped: agent_rate_map.csv"),
        ("not-rated token stripped", sr.counts["stripped_not_rated"],
         "NR/N/A written blank: no rating is what the column means"),
        ("unmatched agency rating", sr.counts["unmatched_rating"],
         "not on that agency's scale: seeded as fed, rating_scales.csv"),
    ]
    if any(n for _, n, _ in norm_counts):
        print()
        print("normalization           : values written through unchanged, listed to be fixed at source")
        for label, n, where in norm_counts:
            if n:
                print(f"  {label:<30}: {n}  ({where})")

    if sr.anomalies:
        # Written, not dropped - listed so they can be dealt with at source.
        print()
        print(f"anomalies              : {len(sr.anomalies)} row(s) written but flagged")
        print(f"  blank investor name          : {sr.counts['blank_investor']}")
        print(f"  missing AccountID            : {sr.counts['blank_account']}")
        print(f"  repeated (facility, investor): {sr.counts['duplicate_facility_investor']}"
              "  <- seeded in full; an investor may hold several positions on one facility")
        for line in sr.anomalies[:20]:
            print(f"    - {line}")
        if len(sr.anomalies) > 20:
            print(f"    ... and {len(sr.anomalies) - 20} more")

    if sr.counts["written"] != len(export):
        raise SystemExit(
            f"RETENTION FAILURE: {len(export)} export rows produced "
            f"{sr.counts['written']} seed rows. Every export row must be written; "
            "the outputs above are incomplete and must not be ingested."
        )
    print()
    print(f"retained               : 100% ({len(export)}/{len(export)} rows)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
