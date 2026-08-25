#!/usr/bin/env python3
"""Facility identity in the extract: the (AccountID, FndName) pair.

The LP DB Export is unique on neither column - FndName repeats down every investor on a facility
and can appear under two accounts, and one AccountID can carry several FndNames. Every export row
must reach the database, so each case asserts 100% retention and zero colliding
(facility_name, investor_name) pairs.

Run:  python pe-sub-jobs/scripts/tests/test_facility_identity.py
"""
from __future__ import annotations

import csv
import importlib.util
import shutil
import sys
import tempfile
import unittest
from collections import Counter
from pathlib import Path

import openpyxl

SCRIPTS_DIR = Path(__file__).resolve().parent.parent


def load_extract(tmp: Path, export_name: str):
    """Import lp_db_extract with its module-level paths repointed at a temp tree."""
    spec = importlib.util.spec_from_file_location("lp_db_extract", SCRIPTS_DIR / "lp_db_extract.py")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = mod        # @dataclass resolves annotations through sys.modules
    spec.loader.exec_module(mod)
    mod.EXPORT_FILE = tmp / "import" / export_name
    mod.AGENT_BANK_SUMMARY_FILE = tmp / "import" / "AgentBankSummaryRpt.xlsx"
    mod.OUT_DIR = tmp / "out"
    mod.REFERENCE_DIR = SCRIPTS_DIR.parent / "data" / "reference"
    return mod


def write_report(path: Path, bands: list) -> None:
    """A banded Agent Bank Summary: (agent, [(borrower, account), ...]) per band.

    A facility may be given as (borrower, account, FacilityStatus) to state a status other than
    the Active every row of the real report carries."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Agent", "Borrower", "AccountNumber", "LoanAmount", None,
               "MaturityDate", "FacilityStatus", "FacilityStatusDate"])
    for agent, facilities in bands:
        ws.append([agent, None, None, None, None, None, None, None])
        for facility in facilities:
            borrower, acct = facility[0], facility[1]
            status = facility[2] if len(facility) > 2 else "Active"
            ws.append([None, borrower, acct, "100000000", None,
                       "2030-01-01", status, "2026-01-01"])
        ws.append([None, None, None, "AccessTotalsLoanAmount:", None, None, None, None])
    wb.save(path)


# The source header spellings, quirks included: columns are located by name.
EXPORT_HEADERS = [
    "AccountID", "FndName", "Investor Name", "Parent", "SPV", "UBS LP Classification",
    "Insitutional vs HNW", "Investment Grade?", "Agent LP Classification", "S&P", "Moody'S",
    "Fitch", "LP Size\r\n($ Bil)", "LP Size Criteria", "Capital Commitments", "Uncalled Capital",
    "UBS Advance Rate", "Agent Advance Rate", "Agent Concentration Limit",
    "UBS Concentration Limit", "% of Capital Commitments", "Called Capital",
    "% of Uncalled Capital", "% of LP Called", "Agent Excess Concentration",
    "UBS Excess Concentration", "Agent Borrowing Base", "UBS Borrowing Base", "Notes", "BBDate",
]


def write_export(path: Path, rows: list) -> None:
    """(AccountID, FndName, Investor Name) triples; every other column gets a plausible constant."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BBs"
    ws.append(EXPORT_HEADERS)
    for acct, fnd, investor in rows:
        ws.append([acct, fnd, investor, "", "N", "Rated Investor", "Institutional", "Y",
                   "Rated", "A", "A2", "A", 13.5, "AUM", 100_000_000, 60_000_000,
                   0.9, 0.95, 0.25, 0.2, 0.05, 40_000_000, 0.05, 0.4,
                   0, 0, 54_000_000, 57_000_000, "", "2026-06-30"])
    wb.save(path)


class FacilityIdentityTest(unittest.TestCase):

    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        (self.tmp / "import").mkdir()
        (self.tmp / "out").mkdir()
        self.addCleanup(shutil.rmtree, self.tmp, True)

    def run_extract(self, report_bands, export_rows):
        write_report(self.tmp / "import" / "AgentBankSummaryRpt.xlsx", report_bands)
        write_export(self.tmp / "import" / "Export.xlsx", export_rows)
        mod = load_extract(self.tmp, "Export.xlsx")
        self.assertEqual(0, mod.main())   # main() raises SystemExit on any retention failure
        out = self.tmp / "out"
        seeds = list(csv.DictReader((out / "lp_facility_seeds.csv").open(encoding="utf-8")))
        facs = list(csv.DictReader((out / "facilities.csv").open(encoding="utf-8")))
        return seeds, facs

    def assert_no_record_is_lost(self, seeds, export_rows):
        """One seed row per export row, each landing on its own facility: two funds sharing an
        account must not merge their investors onto one (facility_name, investor_name) pair."""
        self.assertEqual(len(export_rows), len(seeds))
        pairs = Counter((s["facility_name"], s["investor_name"]) for s in seeds)
        self.assertEqual([], [p for p, n in pairs.items() if n > 1],
                         "investors merged onto one facility - the funds were not kept apart")

    def test_one_account_two_fund_names_are_two_facilities(self):
        """An AccountID shared by two FndNames, both listed by the report."""
        export_rows = [
            ("5VZ8873", "Carlyle Buyout Umbrella", "Copperfield Strategic Family Holdings"),
            ("5VZ8873", "Carlyle Buyout Umbrella", "Ridgeline Pension Trust"),
            ("5VZ8873", "Oaktree Opportunities Fund Xb", "Copperfield Strategic Family Holdings"),
            ("5VZ8873", "Oaktree Opportunities Fund Xb", "Meridian Sovereign Fund"),
        ]
        seeds, facs = self.run_extract(
            [("Bank of America", [("Carlyle Buyout Umbrella", "5VZ8873"),
                                  ("Oaktree Opportunities Fund Xb", "5VZ8873")])],
            export_rows,
        )
        self.assert_no_record_is_lost(seeds, export_rows)
        # Each fund keeps its own LPs.
        self.assertEqual({"Carlyle Buyout Umbrella": 2, "Oaktree Opportunities Fund Xb": 2},
                         dict(Counter(s["facility_name"] for s in seeds)))
        self.assertEqual(2, sum(1 for f in facs if f["account_number"] == "5VZ8873"))
        self.assertTrue(all(f["bank_status"] == "Active" for f in facs))

    def test_second_fund_on_a_shared_account_is_absent_from_the_report(self):
        """The report knows one borrower on the account; the export knows two. The unreported fund
        becomes an Unknown-bank placeholder rather than merging into its sibling."""
        export_rows = [
            ("5VZ8873", "Carlyle Buyout Umbrella", "Copperfield Strategic Family Holdings"),
            ("5VZ8873", "Oaktree Opportunities Fund Xb", "Copperfield Strategic Family Holdings"),
        ]
        seeds, facs = self.run_extract(
            [("Bank of America", [("Carlyle Buyout Umbrella", "5VZ8873")])], export_rows)
        self.assert_no_record_is_lost(seeds, export_rows)
        orphan = [f for f in facs if f["name"] == "Oaktree Opportunities Fund Xb"]
        self.assertEqual(1, len(orphan))
        self.assertEqual("Unknown", orphan[0]["agent_bank"])
        self.assertEqual("5VZ8873", orphan[0]["account_number"])
        self.assertEqual("Inactive", orphan[0]["bank_status"])

    def test_report_borrower_naming_drift_still_joins(self):
        """One borrower on the account and a fund name the report spells differently: the
        positional pass lands the LPs on the reported facility, under the report's name."""
        export_rows = [("5VW3197", "Platinum Equity Capital Partners IV", "Meridian Sovereign Fund")]
        seeds, _ = self.run_extract(
            [("Bank of America", [("Platinum Eq IV", "5VW3197")])], export_rows)
        self.assert_no_record_is_lost(seeds, export_rows)
        self.assertEqual("Platinum Eq IV", seeds[0]["facility_name"])

    def test_fund_name_shared_by_two_accounts_stays_two_facilities(self):
        """One FndName under two AccountIDs: the names are disambiguated so neither facility's
        LPs are redirected into the other, and BOTH carry their account - neither one keeps the
        bare name and leaves its sibling reading as a variant of it."""
        export_rows = [
            ("5VZ9001", "TPG AG Asset Based Credit Fund", "Copperfield Strategic Family Holdings"),
            ("5VZ9002", "TPG AG Asset Based Credit Fund", "Copperfield Strategic Family Holdings"),
        ]
        seeds, facs = self.run_extract([], export_rows)
        self.assert_no_record_is_lost(seeds, export_rows)
        self.assertEqual({"TPG AG Asset Based Credit Fund (5VZ9001)",
                          "TPG AG Asset Based Credit Fund (5VZ9002)"},
                         {f["name"] for f in facs})

    def test_two_accounts_share_a_borrower_name_and_both_are_suffixed(self):
        """The report lists one borrower name under two accounts - a real and recurring shape.
        Both facilities are suffixed with their account, both stay Active, and each keeps its own
        LPs. A facility whose name nothing else shares is left exactly as printed."""
        export_rows = [
            ("5VY4509", "Arctos Sports Fund II", "Meridian Sovereign Fund"),
            ("5VY4577", "Arctos Sports Fund II", "Ridgeline Pension Trust"),
            ("5VX1997", "Whitehorse Liq V", "Meridian Sovereign Fund"),
        ]
        seeds, facs = self.run_extract(
            [("Bank of America", [("Arctos Sports Fund II", "5VY4509"),
                                  ("Arctos Sports Fund II", "5VY4577"),
                                  ("Whitehorse Liq V", "5VX1997")])],
            export_rows,
        )
        self.assert_no_record_is_lost(seeds, export_rows)
        by_acct = {f["account_number"]: f for f in facs}
        self.assertEqual("Arctos Sports Fund II (5VY4509)", by_acct["5VY4509"]["name"])
        self.assertEqual("Arctos Sports Fund II (5VY4577)", by_acct["5VY4577"]["name"])
        self.assertEqual("Whitehorse Liq V", by_acct["5VX1997"]["name"])
        self.assertTrue(all(f["bank_status"] == "Active" for f in facs))
        self.assertEqual({"Arctos Sports Fund II (5VY4509)": 1,
                          "Arctos Sports Fund II (5VY4577)": 1,
                          "Whitehorse Liq V": 1},
                         dict(Counter(s["facility_name"] for s in seeds)))

    def test_reported_active_facility_with_no_lps_is_onboarded_active(self):
        """The report states the facility is Active and the export carries no LPs for it. It is
        onboarded Active and empty - a live facility awaiting its first BB, not a closed one."""
        export_rows = [("5VW3197", "Platinum Equity Capital Partners IV", "Meridian Sovereign Fund")]
        seeds, facs = self.run_extract(
            [("Bank of America", [("Platinum Equity Capital Partners IV", "5VW3197"),
                                  ("Vista Equity Partners IX", "5VW4400")])],
            export_rows,
        )
        self.assert_no_record_is_lost(seeds, export_rows)
        unseeded = [f for f in facs if f["name"] == "Vista Equity Partners IX"]
        self.assertEqual(1, len(unseeded))
        self.assertEqual("Active", unseeded[0]["bank_status"])
        self.assertEqual("Bank of America", unseeded[0]["agent_bank"])
        self.assertEqual([], [s for s in seeds if s["facility_name"] == "Vista Equity Partners IX"])

    def test_reported_non_active_facility_with_no_lps_stays_inactive(self):
        """The promotion is the report's word, not a blanket Active: a status the report does not
        call Active, with no export match behind it, stays Inactive."""
        export_rows = [("5VW3197", "Platinum Equity Capital Partners IV", "Meridian Sovereign Fund")]
        _, facs = self.run_extract(
            [("Bank of America", [("Platinum Equity Capital Partners IV", "5VW3197"),
                                  ("Vista Equity Partners IX", "5VW4400", "Closed")])],
            export_rows,
        )
        closed = [f for f in facs if f["name"] == "Vista Equity Partners IX"]
        self.assertEqual("Inactive", closed[0]["bank_status"])

    def test_export_match_promotes_a_non_active_reported_status(self):
        """An export match still promotes to Active: the export carrying LPs for the facility is
        evidence it is live, so the match only ever promotes a status, never demotes one."""
        export_rows = [("5VW3197", "Platinum Equity Capital Partners IV", "Meridian Sovereign Fund")]
        _, facs = self.run_extract(
            [("Bank of America", [("Platinum Equity Capital Partners IV", "5VW3197", "Closed")])],
            export_rows,
        )
        matched = [f for f in facs if f["name"] == "Platinum Equity Capital Partners IV"]
        self.assertEqual("Active", matched[0]["bank_status"])

    def test_blank_fund_name_on_a_shared_account_still_seeds(self):
        """A blank FndName is its own facility on the account, under a generated name."""
        export_rows = [
            ("5VZ8873", "Carlyle Buyout Umbrella", "Meridian Sovereign Fund"),
            ("5VZ8873", "", "Meridian Sovereign Fund"),
        ]
        seeds, _ = self.run_extract(
            [("Bank of America", [("Carlyle Buyout Umbrella", "5VZ8873")])], export_rows)
        self.assert_no_record_is_lost(seeds, export_rows)


# V2 adds Region (after Parent) and Investor Type (after UBS LP Classification) to the 30 columns
# above. Both are read as optional, so the V1 header list every test above writes is the standing
# proof that a pre-V2 workbook still parses; these tests cover what the two new columns do.
EXPORT_HEADERS_V2 = (EXPORT_HEADERS[:4] + ["Region"] + EXPORT_HEADERS[4:6] + ["Investor Type"]
                     + EXPORT_HEADERS[6:])


def write_export_v2(path: Path, rows: list) -> None:
    """(AccountID, FndName, Investor Name, Region, Investor Type) rows in the V2 column order."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BBs"
    ws.append(EXPORT_HEADERS_V2)
    for acct, fnd, investor, region, itype in rows:
        ws.append([acct, fnd, investor, "", region, "N", "Rated Investor", itype,
                   "Institutional", "Y",
                   "Rated", "A", "A2", "A", 13.5, "AUM", 100_000_000, 60_000_000,
                   0.9, 0.95, 0.25, 0.2, 0.05, 40_000_000, 0.05, 0.4,
                   0, 0, 54_000_000, 57_000_000, "", "2026-06-30"])
    wb.save(path)


class V2ColumnsTest(unittest.TestCase):
    """Region and Investor Type: the two columns V2 restored."""

    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        (self.tmp / "import").mkdir()
        (self.tmp / "out").mkdir()
        self.addCleanup(shutil.rmtree, self.tmp, True)

    def run_extract(self, export_rows, writer=write_export_v2):
        write_report(self.tmp / "import" / "AgentBankSummaryRpt.xlsx",
                     [("Bank of America", [("Carlyle Buyout Umbrella", "5VZ8873")])])
        writer(self.tmp / "import" / "Export.xlsx", export_rows)
        mod = load_extract(self.tmp, "Export.xlsx")
        self.assertEqual(0, mod.main())
        out = self.tmp / "out"
        return (list(csv.DictReader((out / "lp_facility_seeds.csv").open(encoding="utf-8"))),
                list(csv.DictReader((out / "lp_master.csv").open(encoding="utf-8"))))

    def test_investor_type_is_normalized_to_the_canonical_value(self):
        """A known bank spelling reaches both outputs as the type the platform states, so the LP
        Master Investor Type filter - built from the DISTINCT values in the table - does not grow
        an entry per spelling."""
        seeds, master = self.run_extract([
            ("5VZ8873", "Carlyle Buyout Umbrella", "Ridgeline Pension Trust",
             "North America", "Corp Pension"),
        ])
        self.assertEqual("Pension Fund", seeds[0]["investor_type"])
        self.assertEqual("Pension Fund", master[0]["investor_type"])

    def test_unrecognised_investor_type_is_written_through_unchanged(self):
        """Same contract as an unmatched UBS classification: reported, never dropped or rewritten.
        Folding it into a neighbouring type would be data loss, not normalization."""
        seeds, master = self.run_extract([
            ("5VZ8873", "Carlyle Buyout Umbrella", "Ridgeline Pension Trust",
             "North America", "Crypto Treasury DAO"),
        ])
        self.assertEqual("Crypto Treasury DAO", seeds[0]["investor_type"])
        self.assertEqual("Crypto Treasury DAO", master[0]["investor_type"])

    def test_region_is_carried_exactly_as_fed(self):
        """Region is governed by no reference list here, so it is passed through rather than forced
        onto a vocabulary the banks do not write in."""
        seeds, master = self.run_extract([
            ("5VZ8873", "Carlyle Buyout Umbrella", "Ridgeline Pension Trust", "EMEA", "Endowment"),
        ])
        self.assertEqual("EMEA", seeds[0]["region_location"])
        self.assertEqual("EMEA", master[0]["region_location"])

    def test_v1_workbook_leaves_both_columns_blank(self):
        """A workbook written before V2 states neither field. Blank is "not resubmitted", which
        pe-sub-api reads as "keep what LP Master already holds" - not as a clearing edit."""
        seeds, master = self.run_extract(
            [("5VZ8873", "Carlyle Buyout Umbrella", "Ridgeline Pension Trust")],
            writer=write_export,
        )
        self.assertEqual(("", ""), (seeds[0]["investor_type"], seeds[0]["region_location"]))
        self.assertEqual(("", ""), (master[0]["investor_type"], master[0]["region_location"]))

    def test_most_recent_submission_wins_for_both(self):
        """Both are LP Master attributes consolidated across an LP's rows, so a stale row cannot
        pull back a profile a newer submission restated."""
        write_report(self.tmp / "import" / "AgentBankSummaryRpt.xlsx",
                     [("Bank of America", [("Carlyle Buyout Umbrella", "5VZ8873")])])
        path = self.tmp / "import" / "Export.xlsx"
        write_export_v2(path, [
            ("5VZ8873", "Carlyle Buyout Umbrella", "Ridgeline Pension Trust",
             "North America", "Endowment"),
        ])
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        stale = [c.value for c in ws[2]]
        stale[4], stale[7], stale[-1] = "Europe", "Family Office", "2020-01-01"
        ws.append(stale)
        wb.save(path)

        mod = load_extract(self.tmp, "Export.xlsx")
        self.assertEqual(0, mod.main())
        master = list(csv.DictReader(
            (self.tmp / "out" / "lp_master.csv").open(encoding="utf-8")))
        self.assertEqual(1, len(master))
        self.assertEqual(("Endowment", "North America"),
                         (master[0]["investor_type"], master[0]["region_location"]))


if __name__ == "__main__":
    unittest.main(verbosity=2)
