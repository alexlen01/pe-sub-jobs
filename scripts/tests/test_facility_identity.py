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
    """A banded Agent Bank Summary: (agent, [(borrower, account), ...]) per band."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Agent", "Borrower", "AccountNumber", "LoanAmount", None,
               "MaturityDate", "FacilityStatus", "FacilityStatusDate"])
    for agent, facilities in bands:
        ws.append([agent, None, None, None, None, None, None, None])
        for borrower, acct in facilities:
            ws.append([None, borrower, acct, "100000000", None,
                       "2030-01-01", "Active", "2026-01-01"])
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
        LPs are redirected into the other."""
        export_rows = [
            ("5VZ9001", "TPG AG Asset Based Credit Fund", "Copperfield Strategic Family Holdings"),
            ("5VZ9002", "TPG AG Asset Based Credit Fund", "Copperfield Strategic Family Holdings"),
        ]
        seeds, facs = self.run_extract([], export_rows)
        self.assert_no_record_is_lost(seeds, export_rows)
        self.assertEqual(2, len({f["name"] for f in facs}))

    def test_blank_fund_name_on_a_shared_account_still_seeds(self):
        """A blank FndName is its own facility on the account, under a generated name."""
        export_rows = [
            ("5VZ8873", "Carlyle Buyout Umbrella", "Meridian Sovereign Fund"),
            ("5VZ8873", "", "Meridian Sovereign Fund"),
        ]
        seeds, _ = self.run_extract(
            [("Bank of America", [("Carlyle Buyout Umbrella", "5VZ8873")])], export_rows)
        self.assert_no_record_is_lost(seeds, export_rows)


if __name__ == "__main__":
    unittest.main(verbosity=2)
