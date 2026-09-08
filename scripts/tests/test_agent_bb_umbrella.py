#!/usr/bin/env python3
r"""Cover the umbrella the Agent BB generator certifies as ONE file.

An account the export shares between several borrowers is one credit agreement, and the agent
certifies the agreement rather than each fund on it. Two things decide whether the resulting file is
readable as that: the NAME, which has to be the one name true of every member and marked as the
group, and the EXTENT, which has to be every member's roster with none of them sampled away.

Run directly, no test framework needed:
    python pe-sub-jobs/scripts/tests/test_agent_bb_umbrella.py
"""
from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from lp_agent_bb_generate import (  # noqa: E402
    AgentLp, Facility, MIN_ROWS_PER_SHEET, UMBRELLA_MEMBERS, common_name, fold_umbrellas,
    member_label, split_by_member, stamp_members, umbrella_members, umbrella_name,
)

failures: list[str] = []


def check(label: str, actual, expected) -> None:
    if actual != expected:
        failures.append(f"{label}\n    expected: {expected!r}\n    actual:   {actual!r}")


def facility(borrower: str, account: str, loan: float = 100_000_000.0,
             members: tuple[str, ...] = ()) -> Facility:
    return Facility(agent="Agent Bank", borrower=borrower, account=account, loan=loan,
                    maturity=None, status="Active", status_date=None, bb_date=date(2026, 6, 25),
                    members=members)


def lp(name: str, carried: bool = True, fund: str = "") -> AgentLp:
    return AgentLp(name=name, parent="", region="", spv="N", itype="Pension Fund",
                   cls="Included Investors (Rated)", sp="A", moodys="A2", fitch="NR",
                   size_bil=1.0, size_criteria="AUM", commitment=1.0, called=0.0, uncalled=1.0,
                   carried=carried, source={"FndName": fund} if fund else {})


MEMBERS = ("Emberly Real Estate Debt Feeder L.P.", "Emberly Real Estate Debt Master Fund XIV",
           "Emberly Real Estate Debt Offshore X, LP", "Emberly Real Estate Debt SPV")

# ── which accounts are umbrellas ──────────────────────────────────────────────

lo, hi = UMBRELLA_MEMBERS
export = {("ACCT1", f"Fund {i}"): [{}] for i in range(hi + 1)}          # one over the ceiling
export |= {("ACCT2", f"Feeder {i}"): [{}] for i in range(lo)}           # exactly at the floor
export |= {("ACCT3", f"Sleeve {i}"): [{}] for i in range(lo - 1)}       # one under it
export[("ACCT4", "Standalone Fund VII")] = [{}]

check("an account is an umbrella only between the floor and the ceiling",
      sorted(umbrella_members(export)), ["ACCT2"])
check("and it carries its members in export order",
      umbrella_members(export)["ACCT2"], tuple(f"Feeder {i}" for i in range(lo)))

# ── what the group is called ──────────────────────────────────────────────────

check("the members' shared name is what they have in common, legal token dropped",
      common_name(list(MEMBERS)), "Emberly Real Estate Debt")
check("and it is marked as the group",
      umbrella_name(MEMBERS), "Emberly Real Estate Debt Umbrella")
check("a name that already says umbrella is not told twice",
      umbrella_name(("Carlyle Buyout Umbrella A", "Carlyle Buyout Umbrella B")),
      "Carlyle Buyout Umbrella")
check("members sharing no name at all fall back to the printed row, never to the account",
      umbrella_name(("Alpha Fund IV", "Beta Fund II"), printed="Kelvin Buyout Umbrella"),
      "Kelvin Buyout Umbrella")
check("and to the first member where the report printed nothing either",
      umbrella_name(("Alpha Fund IV", "Beta Fund II")), "Alpha Fund Umbrella")
group = facility(umbrella_name(MEMBERS), "ACCT2", members=MEMBERS)
check("the member's tab is what separates it from its siblings",
      [member_label(m, group) for m in MEMBERS],
      ["Feeder L.P.", "Master Fund XIV", "Offshore X, LP", "SPV"])

# ── what the report printed is replaced by the group ──────────────────────────

printed = [facility("Ordinary Fund VI", "ACCT9"),
           *[facility(m, "ACCT2", loan=25_000_000.0) for m in ("Feeder 0", "Feeder 1")],
           facility("Later Fund II", "ACCT8")]
folded = fold_umbrellas(printed, {"ACCT2": ("Feeder 0", "Feeder 1")})
check("every printed row on the account collapses into one facility",
      [f.borrower for f in folded], ["Ordinary Fund VI", "Feeder Umbrella", "Later Fund II"])
check("which stands where the first of them stood, so the run order does not move",
      folded[1].account, "ACCT2")
check("and carries the whole agreement's line, not one member's share",
      folded[1].loan, 50_000_000.0)
check("a facility on an unshared account is untouched",
      (folded[0].is_umbrella, folded[0].loan), (False, 100_000_000.0))

# ── every member reaches the certificate ──────────────────────────────────────

group.lps = [lp("Carried A", fund=MEMBERS[0]), lp("Minted A", carried=False),
             lp("Carried B", fund=MEMBERS[1]), lp("Carried C", fund=MEMBERS[2]),
             lp("Minted B", carried=False), lp("Carried D", fund=MEMBERS[3])]
stamp_members(group)
check("a carried row states the borrower the export booked it against",
      [x.member for x in group.lps if x.carried], list(MEMBERS))
check("and a minted row draws under the member it was interleaved into",
      [x.member for x in group.lps if not x.carried], [MEMBERS[0], MEMBERS[2]])

# ── one sheet per borrower, never one too thin to read ────────────────────────

thick = [lp(f"LP {i}", fund=MEMBERS[0]) for i in range(MIN_ROWS_PER_SHEET)]
for x in thick:
    x.member = MEMBERS[0]
thin = lp("Sole LP", fund=MEMBERS[1])
thin.member = MEMBERS[1]
blocks = split_by_member(thick + [thin])
check("a member too thin for the analyzer joins the block beside it, rows kept",
      [(name, len(rows)) for name, rows in blocks],
      [(MEMBERS[0], MIN_ROWS_PER_SHEET + 1)])

# ── report ────────────────────────────────────────────────────────────────────

if failures:
    print(f"FAILED ({len(failures)}):")
    for f in failures:
        print(f"  {f}")
    sys.exit(1)
print("agent BB umbrella: all checks passed")
