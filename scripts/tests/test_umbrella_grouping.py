#!/usr/bin/env python3
r"""Cover the grouping the extractor stamps onto every facility row.

TWO statements group, and both are required: the export shares an ACCOUNT NUMBER between several
funds, AND the Agent Bank Summary prints the umbrella's own facility row over that account. Neither
is an inference - the account is how the agent administers the funds, the printed row is the agent
naming the agreement above them. Getting it wrong is expensive in both directions: an umbrella missed
leaves each fund reporting as a separate agreement, and one invented from the account alone pools
funds under a credit agreement no file reported.

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

import csv
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from lp_db_extract import (  # noqa: E402
    FACILITY_WORK_COLS, SRC_COLS, UMBRELLA_COLS, StatedGroup, assign_umbrellas,
    legacy_sleeve_reading, upsert_facilities, write_umbrellas,
)

AGENT, NAME, ACCT, SIZE, MATURITY, STATUS = 0, 1, 2, 3, 4, 5
COLLATERAL = FACILITY_WORK_COLS.index("collateral_date")
UBS_PART = FACILITY_WORK_COLS.index("ubs_participation")
UMBRELLA, KEY, XCOLL = 9, 10, 11
TRANCHE_TYPE = FACILITY_WORK_COLS.index("tranche_type")
TRANCHE_OF = FACILITY_WORK_COLS.index("tranche_of")
AGREEMENT_REF = FACILITY_WORK_COLS.index("agreement_ref")

failures: list[str] = []


def check(label: str, actual, expected) -> None:
    if actual != expected:
        failures.append(f"{label}\n    expected: {expected!r}\n    actual:   {actual!r}")


def printed_over(account: str, name: str = "The Agreement", size: str = "", maturity: str = "",
                 status: str = "Active", agreement_ref: str = "",
                 ubs: str = "") -> dict[str, StatedGroup]:
    """The `stated` map for an account the report printed an umbrella row over."""
    return {account: StatedGroup(account, name, "Ashford Bank", size, maturity, status, "",
                                 ubs_participation=ubs, agreement_ref=agreement_ref)}


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

# ── account umbrellas: the account AND the printed row ────────────────────────

rows = [facility("Atlas Growth Fund IV", "5VZ8873"),
        facility("Atlas Growth Feeder IV", "5VZ8873"),
        facility("Zenith Partners VII", "5VZ9001")]
groups = assign_umbrellas(rows, printed_over("5VZ8873", "Atlas Growth Umbrella"))

check("an account the report prints an umbrella over forms one umbrella", len(groups), 1)
check("named from the printed row and keyed on the account", (groups[0].name, groups[0].key),
      ("Atlas Growth Umbrella", "5VZ8873"))
check("the fund borrowing alone joins nothing", rows[2][UMBRELLA], "")
check("a shared borrowing base is NOT inferred - it is a term of the credit agreement",
      [rows[0][XCOLL], rows[1][XCOLL]], ["", ""])

# ── a shared account with NO umbrella row printed over it groups NOTHING ──────
#
# The account says the agent administers these funds together; it does not say what agreement sits
# above them. Minting one from the account number alone produces a credit agreement no file
# reported - named only for the account, with no line and no standing of its own - and the platform
# governs its members' loan terms from it regardless. The funds are fed as the facilities the report
# states they are, and the umbrella is an analyst's to create.

rows = [facility("Atlas Growth Fund IV", "5VZ8873"),
        facility("Atlas Growth Feeder IV", "5VZ8873")]
groups = assign_umbrellas(rows)

check("a shared account alone forms no umbrella", groups, [])
check("and neither fund is stamped with one",
      [(r[UMBRELLA], r[KEY]) for r in rows], [("", ""), ("", "")])

# The umbrella the report DOES print is unaffected by the one it does not: an account is judged on
# its own row, not on what its neighbours carry.
rows = [facility("Atlas Growth Fund IV", "5VZ8873"),
        facility("Atlas Growth Feeder IV", "5VZ8873"),
        facility("Halden Fund I", "5VZ9001"),
        facility("Halden Fund II", "5VZ9001")]
groups = assign_umbrellas(rows, printed_over("5VZ9001", "Halden Umbrella"))
check("only the account with a printed row is grouped",
      [(g.key, g.name) for g in groups], [("5VZ9001", "Halden Umbrella")])
check("and the other account's funds stay ungrouped",
      [r[UMBRELLA] for r in rows], ["", "", "Halden Umbrella", "Halden Umbrella"])

# Blank is the absence of an account, not an account several facilities hold in common.
rows = [facility("Orphan Fund I", ""), facility("Orphan Fund II", "")]
check("accountless facilities are never pooled into one umbrella",
      assign_umbrellas(rows, printed_over("")), [])

# ── a sleeve that shares a grouped account is grouped BY THAT ACCOUNT ─────────
#
# Being a sleeve neither creates a group nor exempts a facility from one. The account is read the
# same way whatever the name on the row says, so the committed sleeve and the co-invest sharing its
# account are members of the umbrella printed over it, and the uncommitted sleeve on its own account
# is in nothing.
rows = [facility("Blackford Continuation III (Committed)", "5VZ8873"),
        facility("Blackford Continuation III (Uncommitted)", "5VZ8874"),
        facility("Blackford Co-Invest III", "5VZ8873")]
groups = assign_umbrellas(rows, printed_over("5VZ8873", "Blackford Umbrella"))

check("the printed account forms the one group", (len(groups), groups[0].key), (1, "5VZ8873"))
check("over the two facilities that hold it, sleeve or not",
      sorted(groups[0].members), ["Blackford Co-Invest III", "Blackford Continuation III (Committed)"])
check("and the sleeve on its own account joins nothing", rows[1][UMBRELLA], "")

# ── a report row that names the credit agreement, not a fund ──────────────────
#
# The shape a real Agent Bank Summary carries for an umbrella: ONE row for the account, naming the
# obligor that signs, with the funds beneath it appearing only in the export. Before this was
# recognised, the positional pass handed that row to whichever fund the export happened to list
# first - giving one arbitrary fund the group's name, the group's whole line and Active
# standing, and leaving its siblings as Unknown/Inactive placeholders.

def report_row(agent: str, borrower: str, acct: str, size: str = "", maturity: str = "",
               status: str = "Active", ubs: str = "") -> list[str]:
    """An Agent Bank Summary row as read_agent_bank_summary emits it."""
    row = [""] * len(FACILITY_WORK_COLS)
    row[AGENT], row[NAME], row[ACCT] = agent, borrower, acct
    row[SIZE], row[MATURITY], row[STATUS] = size, maturity, status
    row[UBS_PART] = ubs
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
                                       size="1500000000", maturity="2029-03-31",
                                       ubs="420000000")])
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
check("neither of the agreement's amounts is stamped on any fund",
      ({r[SIZE] for r in rows}, {r[UBS_PART] for r in rows}), ({""}, {""}))
check("the group's maturity is carried down", {r[MATURITY] for r in rows}, {"2029-03-31"})
check("the agreement is reported at group level", list(stated), ["5VZ8873"])
check("with the name the agent printed", stated["5VZ8873"].name, "Carlyle Buyout Umbrella")
check("and both of the whole agreement's amounts",
      (stated["5VZ8873"].facility_size, stated["5VZ8873"].ubs_participation),
      ("1500000000", "420000000"))

groups = assign_umbrellas(rows, stated)
check("the funds group on the shared account", len(groups), 1)
check("under the name the report states, not one minted from the account number",
      groups[0].name, "Carlyle Buyout Umbrella")
check("keyed on the account, so an analyst rename survives the next run",
      groups[0].key, "5VZ8873")
check("the agreement's amounts ride on the group, not on its members",
      (groups[0].facility_size, groups[0].ubs_participation, groups[0].agent_bank),
      ("1500000000", "420000000", "Wells Fargo"))
check("all six funds are members", sorted(groups[0].members), sorted(CARLYLE))
check("a shared borrowing base is still not inferred", groups[0].cross_collateralized, False)

# One fund on the account is the ordinary rename case: the agent and the export spell the same
# facility differently, and the positional pass is what resolves it. Promoting that row to a group
# would leave a credit agreement with a single member and the fund itself unreported.
fac_data, by_acct = report([report_row("Ashford Bank", "Northwind Cap Ptrs IV", "5VZ9001",
                                       size="400000000", ubs="140000000")])
rows, name_by_key, stated = upsert_facilities(
    fac_data, by_acct, [export_row("5VZ9001", "Northwind Capital Partners IV")])
check("a single fund still claims the row positionally", stated, {})
check("and keeps the report's spelling, both amounts and standing",
      [rows[0][NAME], rows[0][SIZE], rows[0][UBS_PART], rows[0][STATUS]],
      ["Northwind Cap Ptrs IV", "400000000", "140000000", "Active"])

# The report prints every member fund and no agreement above them. Each row name-matches its own
# fund, so nothing is left over, no group row is stated - and nothing is grouped. Two facilities
# administered on one account is what the report says, and it is what the platform is fed.
fac_data, by_acct = report([report_row("Ashford Bank", "Atlas Growth Fund IV", "5VZ8873"),
                            report_row("Ashford Bank", "Atlas Growth Feeder IV", "5VZ8873")])
rows, name_by_key, stated = upsert_facilities(
    fac_data, by_acct,
    [export_row("5VZ8873", "Atlas Growth Fund IV"), export_row("5VZ8873", "Atlas Growth Feeder IV")])
check("a report that prints every member states no group row", stated, {})
check("and with no group row there is no umbrella", assign_umbrellas(rows, stated), [])
check("both funds reach the platform as the facilities the report printed",
      sorted(name_by_key.values()), ["Atlas Growth Feeder IV", "Atlas Growth Fund IV"])
check("neither carrying a group it was never reported under",
      [r[UMBRELLA] for r in rows], ["", ""])

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
groups = assign_umbrellas(rows, printed_over("5VZ8873"))
check("a book that states no reference groups on the printed row alone",
      (len(groups), groups[0].key, groups[0].agreement_ref), (1, "5VZ8873", ""))

# A reference stated over one member and not on the printed row is the agreement's, and the group
# adopts it: the sibling saying nothing is silence, not a second agreement.
rows = [facility("Halden Fund I", "5VZ8873", "CA-9000"),
        facility("Halden Fund II", "5VZ8873")]
groups = assign_umbrellas(rows, printed_over("5VZ8873", "Halden Umbrella"))
check("the printed account groups both funds",
      (len(groups), groups[0].key, [r[UMBRELLA] for r in rows]),
      (1, "5VZ8873", ["Halden Umbrella", "Halden Umbrella"]))
check("and the group adopts the one reference stated over it", groups[0].agreement_ref, "CA-9000")

# Two references over one account is a contradiction in the file. Picking either would resolve the
# group onto an agreement half its members were never said to be under.
rows = [facility("Sablecreek Fund IX", "5VZ8873", "CA-100"),
        facility("Sablecreek Fund X", "5VZ8873", "CA-200")]
groups = assign_umbrellas(rows, printed_over("5VZ8873"))
check("members naming different agreements still group under the printed row",
      (len(groups), groups[0].key), (1, "5VZ8873"))
check("but the group claims neither reference", groups[0].agreement_ref, "")

# A reference printed over a GROUP row is the agreement's own, so every member inherits it and the
# group carries it - while still resolving by the account, which is what the ingest keys on.
fac_data, by_acct = report([report_row("Wells Fargo", "Carlyle Buyout Umbrella", "5VZ8873",
                                       size="1500000000")])
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

# ── the agreement's own terms ─────────────────────────────────────────────────
#
# The group is what the platform governs a member fund from: while it is Active it hands down its
# account number, both amounts, maturity, collateral date and standing, and the member's own
# terms stand aside. A group fed without those arrives Not Stated and governs nothing, so the extract
# carries them - but only ever from what a file states. Nothing below is inferred, summed or split.

def terms(g) -> tuple[str, str, str, str, str, str]:
    """(account, syndicated line, UBS slice, maturity, collateral date, standing) - the group feed."""
    return (g.account_number, g.facility_size, g.ubs_participation, g.maturity_date,
            g.collateral_date, g.facility_status)


fac_data, by_acct = report([report_row("Wells Fargo", "Carlyle Buyout Umbrella", "5VZ8873",
                                       size="1500000000", maturity="2029-03-31")])
rows, _, stated = upsert_facilities(
    fac_data, by_acct, [export_row("5VZ8873", f) for f in CARLYLE])
groups = assign_umbrellas(rows, stated)
check("a printed group row feeds the agreement's terms, not an empty shell",
      terms(groups[0]),
      ("5VZ8873", "1500000000", "", "2029-03-31", "2026-06-25", "Active"))

# The collateral date is the one term the members are not expected to agree on: each is certified
# against its own roster, on its own date. The agreement stands as of the last of them - an earlier
# date would put the group behind evidence it already holds.
fac_data, by_acct = report([report_row("Wells Fargo", "Carlyle Buyout Umbrella", "5VZ8873",
                                       size="1500000000", maturity="2029-03-31")])
rows, _, stated = upsert_facilities(fac_data, by_acct, [
    export_row("5VZ8873", "Carlyle VII", bbdate="2026-03-31"),
    export_row("5VZ8873", "Carlyle VIII", bbdate="2026-06-25"),
    export_row("5VZ8873", "Carlyle CGP", bbdate="2026-05-29"),
])
groups = assign_umbrellas(rows, stated)
check("the group is collateralized as of its members' latest date",
      groups[0].collateral_date, "2026-06-25")

# An agreement the agent does not call Active is fed as the agent stated it. Governing its members
# off a standing the report contradicts would hand down terms the bank says are not in force.
fac_data, by_acct = report([report_row("Wells Fargo", "Carlyle Buyout Umbrella", "5VZ8873",
                                       size="1500000000", status="Matured")])
rows, _, stated = upsert_facilities(
    fac_data, by_acct, [export_row("5VZ8873", f) for f in CARLYLE])
groups = assign_umbrellas(rows, stated)
check("the agent's own standing wins over the default", groups[0].facility_status, "Inactive")

# ── the members' own terms are never read as the agreement's ──────────────────
#
# Funds on one account repeating one figure are each being shown their own line, and the report is
# the only thing that states the agreement's. Where it printed a row, the group's line and maturity
# come off that row whatever the members carry; the collateral date is the members' latest, because
# it is the one term each is certified on separately.

def member(name: str, account: str, size: str = "", maturity: str = "",
           collateral: str = "") -> list[str]:
    row = facility(name, account)
    row[SIZE], row[MATURITY], row[COLLATERAL] = size, maturity, collateral
    return row


rows = [member("Atlas Growth Fund IV", "5VZ8873", "750000000", "2030-06-30", "2026-06-25"),
        member("Atlas Growth Feeder IV", "5VZ8873", "750000000", "2030-06-30", "2026-04-30")]
check("members agreeing on a figure still do not make an umbrella between them",
      assign_umbrellas(rows), [])

groups = assign_umbrellas(rows, printed_over("5VZ8873", "Atlas Growth Umbrella", size="1500000000",
                                             maturity="2029-03-31", ubs="480000000"))
check("the printed row states the agreement's amounts and maturity, not the members' figures",
      terms(groups[0]),
      ("5VZ8873", "1500000000", "480000000", "2029-03-31", "2026-06-25", "Active"))

# Different figures on one account are the funds' own allocations. Summing them would state a line no
# file does, and picking one would govern every sibling off a number printed for one fund - so the
# members are not read for it at all. A printed row that states no line leaves it blank, which is
# silence, and the platform derives the size from the shares instead.
rows = [member("Sablecreek Fund IX", "5VZ8873", "400000000", "2030-06-30"),
        member("Sablecreek Fund X", "5VZ8873", "250000000", "2031-12-31")]
groups = assign_umbrellas(rows, printed_over("5VZ8873"))
check("members' own lines are never summed into the agreement's, nor picked between",
      (groups[0].facility_size, groups[0].maturity_date), ("", ""))
check("and the group is still Active, which is what makes it govern at all",
      groups[0].facility_status, "Active")

# Not Stated is the one value never fed. It was the old default and it left every extracted
# agreement inert until somebody opened it and retyped what the file already said.
rows = [member("Orphan Group Fund I", "5VZ9100"), member("Orphan Group Fund II", "5VZ9100")]
groups = assign_umbrellas(rows, printed_over("5VZ9100", "Orphan Group Umbrella"))
check("a group stating no terms at all is still Active",
      terms(groups[0]), ("5VZ9100", "", "", "", "", "Active"))

# ── the feed's own shape ──────────────────────────────────────────────────────
#
# The new columns are APPENDED. A reader positioned on the older header goes on reading the columns
# before them unchanged, and the row the writer emits has to stay in step with the header it wrote -
# a ragged row is read one column out of step the whole way across.
check("the terms are appended to the group feed, never inserted",
      UMBRELLA_COLS,
      ["key", "name", "agent_bank", "account_number", "facility_size", "cross_collateralized",
       "agreement_ref", "maturity_date", "collateral_date", "facility_status",
       "ubs_participation"])

with tempfile.TemporaryDirectory() as tmp:
    path = Path(tmp) / "umbrellas.csv"
    write_umbrellas(path, groups)
    written = list(csv.reader(path.open(newline="", encoding="utf-8")))

check("the file's header is the contract", written[0], UMBRELLA_COLS)
check("and every group writes one cell per column, in that order",
      written[1], ["5VZ9100", "Orphan Group Umbrella", "Ashford Bank", "5VZ9100", "", "", "", "",
                   "", "Active", ""])

# ── report ────────────────────────────────────────────────────────────────────

if failures:
    print(f"FAILED ({len(failures)}):")
    for f in failures:
        print(f"  {f}")
    sys.exit(1)
print("umbrella grouping: all checks passed")
