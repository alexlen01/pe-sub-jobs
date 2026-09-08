#!/usr/bin/env python3
r"""Cover the grouping the extractor stamps onto every facility row.

Two shapes of credit agreement cover more than one facility, and getting either wrong is expensive
in opposite directions: a tranche set left ungrouped reports its one borrowing base once per sleeve
and doubles the collateral, while an account umbrella wrongly marked cross-collateralized divides
bases that were never shared and under-reports it.

Run directly, no test framework needed:
    python pe-sub-jobs/scripts/tests/test_umbrella_grouping.py
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from lp_db_extract import (  # noqa: E402
    FACILITY_WORK_COLS, SRC_COLS, assign_umbrellas, legacy_sleeve_reading, upsert_facilities,
)

AGENT, NAME, ACCT, LOAN, MATURITY, STATUS = 0, 1, 2, 3, 4, 5
UMBRELLA, KEY, XCOLL = 9, 10, 11
TRANCHE_TYPE = FACILITY_WORK_COLS.index("tranche_type")
TRANCHE_OF = FACILITY_WORK_COLS.index("tranche_of")
AGREEMENT_REF = FACILITY_WORK_COLS.index("agreement_ref")

failures: list[str] = []


def check(label: str, actual, expected) -> None:
    if actual != expected:
        failures.append(f"{label}\n    expected: {expected!r}\n    actual:   {actual!r}")


def facility(name: str, account: str, agreement_ref: str = "") -> list[str]:
    """A FACILITY_WORK_COLS-shaped row with only the fields the grouping reads populated.

    The tranche declaration is filled here the way upsert_facilities fills it for a file that
    states none - from the legacy sleeve suffix - because the grouping itself reads the
    declaration and nothing else. That separation is the point of the column: the suffix supplies
    a default for a field the feed is expected to carry, and stops deciding anything on its own."""
    row = [""] * len(FACILITY_WORK_COLS)
    row[0], row[NAME], row[ACCT] = "Ashford Bank", name, account
    row[AGREEMENT_REF] = agreement_ref
    reading = legacy_sleeve_reading(name)
    if reading is not None:
        row[TRANCHE_TYPE], row[TRANCHE_OF] = reading
    return row


# ── the suffix is what makes a name a sleeve ──────────────────────────────────

check("a committed sleeve reads as its type and its parent",
      legacy_sleeve_reading("Audax Private Equity Fund VII LP (Committed)"),
      ("COMMITTED", "Audax Private Equity Fund VII LP"))
check("the report's upper-case spelling reads the same",
      legacy_sleeve_reading("AUDAX PRIVATE EQUITY FUND VII LP (UNCOMMITTED)"),
      ("UNCOMMITTED", "AUDAX PRIVATE EQUITY FUND VII LP"))
check("a facility with no suffix is not a sleeve",
      legacy_sleeve_reading("Audax Private Equity Fund VII LP"), None)
check("a parenthetical that is not a tranche is not a suffix",
      legacy_sleeve_reading("Sablecreek Global Fund IX (Series B)"), None)

# ── tranche sets ──────────────────────────────────────────────────────────────

rows = [facility("Merriden Opportunities Fund V (Committed)", "5VZ8873"),
        facility("Merriden Opportunities Fund V (Uncommitted)", "5VZ8874")]
groups = assign_umbrellas(rows)

check("two sleeves on two accounts form one group", len(groups), 1)
check("the group takes the credit agreement's own name",
      groups[0].name, "Merriden Opportunities Fund V")
check("both sleeves are stamped with it",
      [r[UMBRELLA] for r in rows],
      ["Merriden Opportunities Fund V", "Merriden Opportunities Fund V"])
check("the sleeves share one resolution key, not their own account numbers",
      {r[KEY] for r in rows}, {"merriden opportunities fund v"})
check("the shared borrowing base is stated on both",
      [r[XCOLL] for r in rows], ["true", "true"])

# The sleeves are typed into the agent's system separately, so the pair reaches the file spelt the
# same way only by habit. Casing must not decide whether one borrowing base is counted twice.
rows = [facility("MERRIDEN OPPORTUNITIES FUND V (COMMITTED)", "5VZ8873"),
        facility("Merriden Opportunities Fund V (Uncommitted)", "5VZ8874")]
groups = assign_umbrellas(rows)
check("sleeves spelt differently still pair", len(groups), 1)
check("and resolve by one key", {r[KEY] for r in rows}, {"merriden opportunities fund v"})
check("the group is named as the file spells the first sleeve, not folded flat",
      groups[0].name, "MERRIDEN OPPORTUNITIES FUND V")

rows = [facility("Kestrel Real Estate Debt VII (Committed)", "5VZ8873")]
check("a lone sleeve is a facility, not a group", assign_umbrellas(rows), [])
check("and is stamped with nothing", [rows[0][UMBRELLA], rows[0][KEY], rows[0][XCOLL]], ["", "", ""])

# A whole facility standing beside a tranche set of the same name is a separate credit agreement.
# Folding it in would put its full borrowing base under an allocated share of a pool it is not in.
rows = [facility("Zephyrus Mezzanine", "5VZ8870"),
        facility("Zephyrus Mezzanine (Committed)", "5VZ8873"),
        facility("Zephyrus Mezzanine (Uncommitted)", "5VZ8874")]
groups = assign_umbrellas(rows)
check("the untranched facility is left out of the sleeves' group",
      [r[UMBRELLA] for r in rows],
      ["", "Umbrella Zephyrus Mezzanine", "Umbrella Zephyrus Mezzanine"])
check("and the group is renamed so the two are distinguishable on screen",
      groups[0].name, "Umbrella Zephyrus Mezzanine")

# ── account umbrellas ─────────────────────────────────────────────────────────

rows = [facility("Atlas Growth Fund IV", "5VZ8873"),
        facility("Atlas Growth Feeder IV", "5VZ8873"),
        facility("Zenith Partners VII", "5VZ9001")]
groups = assign_umbrellas(rows)

check("one account carrying two borrowers forms one umbrella", len(groups), 1)
check("named and keyed from the account", (groups[0].name, groups[0].key),
      ("Umbrella 5VZ8873", "5VZ8873"))
check("the fund borrowing alone joins nothing", rows[2][UMBRELLA], "")
check("a shared borrowing base is NOT inferred - it is a term of the credit agreement",
      [rows[0][XCOLL], rows[1][XCOLL]], ["", ""])

# Blank is the absence of an account, not an account several facilities hold in common.
rows = [facility("Orphan Fund I", ""), facility("Orphan Fund II", "")]
check("accountless facilities are never pooled into one umbrella", assign_umbrellas(rows), [])

# ── the two passes together ───────────────────────────────────────────────────

# The sleeves are the tighter structure - they share collateral, where account members merely share
# paperwork - so a sleeve is claimed by its tranche set and withheld from the account pass. Here
# that leaves the account with one unclaimed member, which correctly forms nothing.
rows = [facility("Blackford Continuation III (Committed)", "5VZ8873"),
        facility("Blackford Continuation III (Uncommitted)", "5VZ8874"),
        facility("Blackford Co-Invest III", "5VZ8873")]
groups = assign_umbrellas(rows)

check("the tranche set claims its sleeves", len(groups), 1)
check("so the co-invest sharing a sleeve's account is not dragged into a second group",
      rows[2][UMBRELLA], "")
check("and no facility carries two groups",
      [r[UMBRELLA] for r in rows[:2]],
      ["Blackford Continuation III", "Blackford Continuation III"])

# ── a report row that names the credit agreement, not a fund ──────────────────
#
# The shape a real Agent Bank Summary carries for an umbrella: ONE row for the account, naming the
# obligor that signs, with the funds beneath it appearing only in the export. Before this was
# recognised, the positional pass handed that row to whichever fund the export happened to list
# first - giving one arbitrary fund the group's name, the group's whole loan amount and Active
# standing, and leaving its siblings as Unknown/Inactive placeholders.

def report_row(agent: str, borrower: str, acct: str, loan: str = "", maturity: str = "",
               status: str = "Active") -> list[str]:
    """An Agent Bank Summary row as read_agent_bank_summary emits it."""
    row = [""] * len(FACILITY_WORK_COLS)
    row[AGENT], row[NAME], row[ACCT] = agent, borrower, acct
    row[LOAN], row[MATURITY], row[STATUS] = loan, maturity, status
    return row


def report(rows: list[list[str]]) -> tuple[list[list[str]], dict[str, list[tuple[int, str]]]]:
    """(rows, account -> [(row index, normalised Borrower)]) - the pair the reader returns."""
    by_acct: dict[str, list[tuple[int, str]]] = {}
    for i, r in enumerate(rows):
        by_acct.setdefault(r[ACCT], []).append((i, r[NAME].lower()))
    return rows, by_acct


def export_row(acct: str, fund: str, bbdate: str = "2026-06-25") -> dict:
    """An export row with only the columns the facility join reads populated."""
    row = {c: None for c in SRC_COLS}
    row["AccountID"], row["FndName"], row["BBDate"] = acct, fund, bbdate
    return row


CARLYLE = ["Carlyle CGFSP III", "Carlyle CGP", "Carlyle CGP II",
           "Carlyle CP Growth", "Carlyle VII", "Carlyle VIII"]

fac_data, by_acct = report([report_row("Wells Fargo", "Carlyle Buyout Umbrella", "5VZ8873",
                                       loan="1500000000", maturity="2029-03-31")])
rows, name_by_key, stated = upsert_facilities(
    fac_data, by_acct, [export_row("5VZ8873", f) for f in CARLYLE])

check("the group row is not a facility - six funds in, six facilities out", len(rows), 6)
check("every fund keeps its own name, none takes the agreement's",
      sorted(r[NAME] for r in rows), sorted(CARLYLE))
check("every fund's LP records resolve to its own facility",
      sorted(name_by_key.values()), sorted(CARLYLE))
check("no fund is banked to Unknown - the report states the agent at group level",
      {r[AGENT] for r in rows}, {"Wells Fargo"})
check("no fund is Inactive - the export carries a live roster against each",
      {r[STATUS] for r in rows}, {"Active"})
check("the agreement's loan amount is not stamped on any fund",
      {r[LOAN] for r in rows}, {""})
check("the group's maturity is carried down", {r[MATURITY] for r in rows}, {"2029-03-31"})
check("the agreement is reported at group level", list(stated), ["5VZ8873"])
check("with the name the agent printed", stated["5VZ8873"].name, "Carlyle Buyout Umbrella")
check("and the whole agreement's loan amount", stated["5VZ8873"].loan_amount, "1500000000")

groups = assign_umbrellas(rows, stated)
check("the funds group on the shared account", len(groups), 1)
check("under the name the report states, not one minted from the account number",
      groups[0].name, "Carlyle Buyout Umbrella")
check("keyed on the account, so an analyst rename survives the next run",
      groups[0].key, "5VZ8873")
check("the agreement's line rides on the group, not on its members",
      (groups[0].loan_amount, groups[0].agent_bank), ("1500000000", "Wells Fargo"))
check("all six funds are members", sorted(groups[0].members), sorted(CARLYLE))
check("a shared borrowing base is still not inferred", groups[0].cross_collateralized, False)

# One fund on the account is the ordinary rename case: the agent and the export spell the same
# facility differently, and the positional pass is what resolves it. Promoting that row to a group
# would leave a credit agreement with a single member and the fund itself unreported.
fac_data, by_acct = report([report_row("Ashford Bank", "Northwind Cap Ptrs IV", "5VZ9001",
                                       loan="400000000")])
rows, name_by_key, stated = upsert_facilities(
    fac_data, by_acct, [export_row("5VZ9001", "Northwind Capital Partners IV")])
check("a single fund still claims the row positionally", stated, {})
check("and keeps the report's spelling, loan amount and standing",
      [rows[0][NAME], rows[0][LOAN], rows[0][STATUS]],
      ["Northwind Cap Ptrs IV", "400000000", "Active"])

# The generator's shape, and the one the design was built on: the report prints every member fund.
# Each row name-matches its own fund, so nothing is left over and no group row is inferred.
fac_data, by_acct = report([report_row("Ashford Bank", "Atlas Growth Fund IV", "5VZ8873"),
                            report_row("Ashford Bank", "Atlas Growth Feeder IV", "5VZ8873")])
rows, name_by_key, stated = upsert_facilities(
    fac_data, by_acct,
    [export_row("5VZ8873", "Atlas Growth Fund IV"), export_row("5VZ8873", "Atlas Growth Feeder IV")])
check("a report that prints every member states no group row", stated, {})
groups = assign_umbrellas(rows, stated)
check("so the group's name is still minted from the account", groups[0].name, "Umbrella 5VZ8873")
check("and it records no agreement line, because the report printed none over the group",
      groups[0].loan_amount, "")

# Two report rows over three funds: one row may be a group row and one a member, or both members
# spelt differently. An ambiguous row stays a facility - demoting one to a group would take its LP
# roster with it, and the positional pass at least keeps every row addressable.
fac_data, by_acct = report([report_row("Ashford Bank", "Halden Umbrella", "5VZ8873"),
                            report_row("Ashford Bank", "Halden Fund II", "5VZ8873")])
rows, name_by_key, stated = upsert_facilities(
    fac_data, by_acct, [export_row("5VZ8873", f) for f in
                        ["Halden Fund I", "Halden Fund II", "Halden Fund III"]])
check("an ambiguous account is not promoted to a group", stated, {})
check("and every fund still resolves to a facility", len(name_by_key), 3)

# The signing entity may also borrow in its own right, on its own account. Umbrella and facility names
# share one namespace on screen, so the group yields the bare name - a facility's name is what its
# LP records resolve by, and renaming it would strand them.
fac_data, by_acct = report([report_row("Ashford Bank", "Sablecreek Holdings", "5VZ8873"),
                            report_row("Ashford Bank", "Sablecreek Holdings", "5VZ9100")])
rows, name_by_key, stated = upsert_facilities(
    fac_data, by_acct,
    [export_row("5VZ8873", "Sablecreek Fund IX"), export_row("5VZ8873", "Sablecreek Fund X"),
     export_row("5VZ9100", "Sablecreek Holdings")])
check("the group row is read on the multi-fund account only", list(stated), ["5VZ8873"])
groups = assign_umbrellas(rows, stated)
check("a group whose printed name a facility already holds is prefixed apart",
      groups[0].name, "Umbrella Sablecreek Holdings")
check("and the standalone facility keeps the bare name", sorted(name_by_key.values()),
      ["Sablecreek Fund IX", "Sablecreek Fund X", "Sablecreek Holdings"])

# ── the agreement reference ───────────────────────────────────────────────────
#
# The key the grouping now turns on. An account number is how a BANK administers an agreement: it
# is re-papered on a renewal, split across two systems after a merger, reissued after a novation.
# Every one of those events used to cut one credit agreement into two groups, each carrying its own
# borrowing base over the same collateral. A reference identifies the agreement itself.

rows = [facility("Bridgepoint Fund VI", "5VZ8873", "CA-2021-4471"),
        facility("Bridgepoint Feeder VI", "5VZ9001", "CA-2021-4471")]
groups = assign_umbrellas(rows)

check("two accounts under one reference are one credit agreement", len(groups), 1)
check("keyed and named by the reference, because nothing else identifies the agreement",
      (groups[0].key, groups[0].name, groups[0].agreement_ref),
      ("CA-2021-4471", "CA-2021-4471", "CA-2021-4471"))
check("both funds are members", sorted(groups[0].members),
      ["Bridgepoint Feeder VI", "Bridgepoint Fund VI"])
check("and the group records no account, because its members hold two",
      groups[0].account_number, "")
check("a shared borrowing base is not inferred from shared paperwork",
      groups[0].cross_collateralized, False)

# Retention: the reference is the newest column and no file written to date carries it. A book that
# states none must group exactly as it did before the column existed.
rows = [facility("Atlas Growth Fund IV", "5VZ8873"),
        facility("Atlas Growth Feeder IV", "5VZ8873")]
groups = assign_umbrellas(rows)
check("a book that states no reference still groups by account",
      (len(groups), groups[0].key, groups[0].agreement_ref), (1, "5VZ8873", ""))

# One reference over one facility relates it to nothing, so it forms no group - but it still shares
# an account with a sibling, and that account still does.
rows = [facility("Halden Fund I", "5VZ8873", "CA-9000"),
        facility("Halden Fund II", "5VZ8873")]
groups = assign_umbrellas(rows)
check("a lone reference forms no group of its own", len(groups), 1)
check("and the account still groups both funds", (groups[0].key, [r[UMBRELLA] for r in rows]),
      ("5VZ8873", ["Umbrella 5VZ8873", "Umbrella 5VZ8873"]))
check("and the group adopts the one reference stated over it - a silent sibling is silence, "
      "not a second agreement",
      groups[0].agreement_ref, "CA-9000")

# Sleeves share collateral; members of a reference merely share paperwork. The tighter fact is
# settled first, and a sleeve already in a set is never re-cut by a reference the file also states.
rows = [facility("Blackford Continuation III (Committed)", "5VZ8873", "CA-77"),
        facility("Blackford Continuation III (Uncommitted)", "5VZ8874", "CA-77"),
        facility("Blackford Co-Invest III", "5VZ8875", "CA-77")]
groups = assign_umbrellas(rows)
check("the tranche set is settled before the reference is read", len(groups), 1)
check("the sleeves keep their set, keyed on the facility they are sleeves of",
      (groups[0].key, groups[0].cross_collateralized),
      ("blackford continuation iii", True))
check("and the set records the reference its sleeves agree on",
      groups[0].agreement_ref, "CA-77")
check("the co-invest is left ungrouped - one unclaimed member relates to nothing",
      rows[2][UMBRELLA], "")

# Two references over what would otherwise be one group is a contradiction in the file. Picking
# either would resolve the group onto an agreement half its members were never said to be under.
rows = [facility("Sablecreek Fund IX", "5VZ8873", "CA-100"),
        facility("Sablecreek Fund X", "5VZ8873", "CA-200")]
groups = assign_umbrellas(rows)
check("members naming different agreements still group by their shared account",
      (len(groups), groups[0].key), (1, "5VZ8873"))
check("but the group claims neither reference", groups[0].agreement_ref, "")

# The agent's own name for the agreement beats one minted from its reference - and the terms the
# report printed over the group row ride with it.
fac_data, by_acct = report([report_row("Wells Fargo", "Carlyle Buyout Umbrella", "5VZ8873",
                                       loan="1500000000")])
fac_data[0][AGREEMENT_REF] = "CA-2021-4471"
rows, name_by_key, stated = upsert_facilities(
    fac_data, by_acct, [export_row("5VZ8873", f) for f in CARLYLE])
check("the promoted group row carries the reference it was printed with",
      stated["5VZ8873"].agreement_ref, "CA-2021-4471")
check("and every member the report never printed inherits it",
      {r[AGREEMENT_REF] for r in rows}, {"CA-2021-4471"})
groups = assign_umbrellas(rows, stated)
check("so the members group by the reference, not by the account",
      (len(groups), groups[0].key), (1, "CA-2021-4471"))
check("under the name the agent printed", groups[0].name, "Carlyle Buyout Umbrella")
check("with the account still recorded, because all six members are on it",
      groups[0].account_number, "5VZ8873")

# ── report ────────────────────────────────────────────────────────────────────

if failures:
    print(f"FAILED ({len(failures)}):")
    for f in failures:
        print(f"  {f}")
    sys.exit(1)
print("umbrella grouping: all checks passed")
