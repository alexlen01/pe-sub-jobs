#!/usr/bin/env python3
r"""Cover the grouping the extractor stamps onto every facility row.

ONE relationship groups: a shared ACCOUNT NUMBER, which is not an inference but the agent's own
statement that several funds draw on one credit agreement. Getting it wrong is expensive in both
directions - an umbrella missed leaves each fund reporting as a separate agreement, and one invented
pools funds that never shared paperwork.

Everything else is left to a user. Facilities holding their OWN account numbers are never related
here, however plainly they read as one agreement - the sleeves of a multi-tranche facility, or two
accounts an agent printed one reference over - because the evidence is a typed name suffix or a
reference that gets re-papered, and acting on it would move a whole borrowing base into a pool on a
string match. These checks hold the script to that, in both directions: what it must group, and what
it must leave alone.

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
    states none - from the legacy sleeve suffix. It is carried so the run report can name the
    sleeve sets someone has to map; the grouping never reads it."""
    row = [""] * len(FACILITY_WORK_COLS)
    row[0], row[NAME], row[ACCT] = "Ashford Bank", name, account
    row[AGREEMENT_REF] = agreement_ref
    reading = legacy_sleeve_reading(name)
    if reading is not None:
        row[TRANCHE_TYPE], row[TRANCHE_OF] = reading
    return row


# ── the suffix still READS as a sleeve; it just no longer groups ──────────────

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

# ── tranche sets are NOT grouped ──────────────────────────────────────────────
#
# The sleeves of one credit agreement hold one account number each, so nothing here relates them.
# That is the decision, not an oversight: a suffix is what an operator typed, and folding a whole
# facility's borrowing base into an allocated share of a pool on that reading is not a call this
# script gets to make over a book at a time. The sleeves reach the platform as two facilities with
# their declaration intact, and a user relates them there.

rows = [facility("Merriden Opportunities Fund V (Committed)", "5VZ8873"),
        facility("Merriden Opportunities Fund V (Uncommitted)", "5VZ8874")]
groups = assign_umbrellas(rows)

check("two sleeves on two accounts are two facilities, not one group", groups, [])
check("and neither is stamped with a group",
      [(r[UMBRELLA], r[KEY], r[XCOLL]) for r in rows], [("", "", ""), ("", "", "")])
check("the declaration they arrived with is still on the row, for whoever maps them",
      [(r[TRANCHE_TYPE], r[TRANCHE_OF]) for r in rows],
      [("COMMITTED", "Merriden Opportunities Fund V"),
       ("UNCOMMITTED", "Merriden Opportunities Fund V")])

# A borrowing base is never divided on this reading either: cross_collateralized is a term of a
# credit agreement, and no file this script reads states one.
check("nothing is ever marked cross-collateralized",
      [r[XCOLL] for r in rows], ["", ""])

rows = [facility("Zephyrus Mezzanine", "5VZ8870"),
        facility("Zephyrus Mezzanine (Committed)", "5VZ8873"),
        facility("Zephyrus Mezzanine (Uncommitted)", "5VZ8874")]
check("a whole facility standing beside two sleeves of its own name groups nothing",
      (assign_umbrellas(rows), [r[UMBRELLA] for r in rows]), ([], ["", "", ""]))

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

# ── a sleeve that shares an account is grouped BY THAT ACCOUNT ────────────────
#
# Being a sleeve neither creates a group nor exempts a facility from one. The account is read the
# same way whatever the name on the row says, so the committed sleeve and the co-invest sharing its
# account are one umbrella, and the uncommitted sleeve on its own account is in nothing.
rows = [facility("Blackford Continuation III (Committed)", "5VZ8873"),
        facility("Blackford Continuation III (Uncommitted)", "5VZ8874"),
        facility("Blackford Co-Invest III", "5VZ8873")]
groups = assign_umbrellas(rows)

check("the shared account forms the one group", (len(groups), groups[0].key), (1, "5VZ8873"))
check("over the two facilities that hold it, sleeve or not",
      sorted(groups[0].members), ["Blackford Co-Invest III", "Blackford Continuation III (Committed)"])
check("and the sleeve on its own account joins nothing", rows[1][UMBRELLA], "")

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
# Carried, and never grouped on. A reference identifies the agreement rather than the account it is
# administered under, which makes it the best thing a person has to relate two facilities by - but
# an agent re-papers it on renewal, quotes the wrong one, and prints one reference over two
# agreements a merger left it holding. Acting on it here would move a borrowing base into a pool on
# the strength of a string, so it rides on the rows and on the group, and a user does the relating.

rows = [facility("Bridgepoint Fund VI", "5VZ8873", "CA-2021-4471"),
        facility("Bridgepoint Feeder VI", "5VZ9001", "CA-2021-4471")]
groups = assign_umbrellas(rows)

check("two accounts under one reference stay two facilities", groups, [])
check("neither is stamped with a group", [r[UMBRELLA] for r in rows], ["", ""])
check("and each keeps the reference, which is what a user relates them by",
      [r[AGREEMENT_REF] for r in rows], ["CA-2021-4471", "CA-2021-4471"])

rows = [facility("Atlas Growth Fund IV", "5VZ8873"),
        facility("Atlas Growth Feeder IV", "5VZ8873")]
groups = assign_umbrellas(rows)
check("a book that states no reference groups by account, as it always has",
      (len(groups), groups[0].key, groups[0].agreement_ref), (1, "5VZ8873", ""))

# A reference stated over one member of an account group is the agreement's, and the group adopts
# it: the sibling saying nothing is silence, not a second agreement.
rows = [facility("Halden Fund I", "5VZ8873", "CA-9000"),
        facility("Halden Fund II", "5VZ8873")]
groups = assign_umbrellas(rows)
check("the account groups both funds", (len(groups), groups[0].key, [r[UMBRELLA] for r in rows]),
      (1, "5VZ8873", ["Umbrella 5VZ8873", "Umbrella 5VZ8873"]))
check("and the group adopts the one reference stated over it", groups[0].agreement_ref, "CA-9000")

# Two references over one account is a contradiction in the file. Picking either would resolve the
# group onto an agreement half its members were never said to be under.
rows = [facility("Sablecreek Fund IX", "5VZ8873", "CA-100"),
        facility("Sablecreek Fund X", "5VZ8873", "CA-200")]
groups = assign_umbrellas(rows)
check("members naming different agreements still group by their shared account",
      (len(groups), groups[0].key), (1, "5VZ8873"))
check("but the group claims neither reference", groups[0].agreement_ref, "")

# A reference printed over a GROUP row is the agreement's own, so every member inherits it and the
# group carries it - while still resolving by the account, which is what the ingest keys on.
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
check("the members still group by the account they share",
      (len(groups), groups[0].key), (1, "5VZ8873"))
check("under the name the agent printed", groups[0].name, "Carlyle Buyout Umbrella")
check("with the reference recorded on the group", groups[0].agreement_ref, "CA-2021-4471")

# ── report ────────────────────────────────────────────────────────────────────

if failures:
    print(f"FAILED ({len(failures)}):")
    for f in failures:
        print(f"  {f}")
    sys.exit(1)
print("umbrella grouping: all checks passed")
