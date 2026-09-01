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
    FACILITY_COLS, assign_umbrellas, tranche_base_name,
)

NAME, ACCT, UMBRELLA, KEY, XCOLL = 1, 2, 9, 10, 11

failures: list[str] = []


def check(label: str, actual, expected) -> None:
    if actual != expected:
        failures.append(f"{label}\n    expected: {expected!r}\n    actual:   {actual!r}")


def facility(name: str, account: str) -> list[str]:
    """A FACILITY_COLS-shaped row with only the fields the grouping reads populated."""
    row = [""] * len(FACILITY_COLS)
    row[0], row[NAME], row[ACCT] = "Ashford Bank", name, account
    return row


# ── the suffix is what makes a name a sleeve ──────────────────────────────────

check("a committed sleeve yields its base",
      tranche_base_name("Audax Private Equity Fund VII LP (Committed)"),
      "Audax Private Equity Fund VII LP")
check("the report's upper-case spelling yields the same shape",
      tranche_base_name("AUDAX PRIVATE EQUITY FUND VII LP (UNCOMMITTED)"),
      "AUDAX PRIVATE EQUITY FUND VII LP")
check("a facility with no suffix is not a sleeve",
      tranche_base_name("Audax Private Equity Fund VII LP"), None)
check("a parenthetical that is not a tranche is not a suffix",
      tranche_base_name("Sablecreek Global Fund IX (Series B)"), None)

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

# ── report ────────────────────────────────────────────────────────────────────

if failures:
    print(f"FAILED ({len(failures)}):")
    for f in failures:
        print(f"  {f}")
    sys.exit(1)
print("umbrella grouping: all checks passed")
