#!/usr/bin/env python3
r"""
Parse Excel workbook → generate BB template

Analyzes a raw Agent BB or investor-list workbook and generates a BB-Template-Import-<slug>.xlsx
meta-workbook. pe-sub-jobs watches data/bb-templates/ and automatically picks up and upserts any
BB-Template-Import-*.xlsx file it finds.

USAGE
    python parse_excel_templates.py <workbook.xlsx>

Example:
    python parse_excel_templates.py data/import/agent-bb.xlsx
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys
import urllib.error
import urllib.request
from dataclasses import dataclass, field, asdict
from datetime import datetime
from pathlib import Path
from typing import Optional

try:
    import openpyxl
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover
    sys.exit("openpyxl is required: pip install openpyxl")

# Paths anchored to this file's location (pe-sub-jobs/scripts/), same convention as
# lp_db_extract.py, so the script runs correctly from any working directory.
SCRIPT_DIR = Path(__file__).resolve().parent            # pe-sub-jobs/scripts/
JOBS_ROOT = SCRIPT_DIR.parent                            # pe-sub-jobs/
DATA_DIR = JOBS_ROOT / "data"
DEFAULT_IMPORT_DIR = DATA_DIR / "import"
WATCHED_TEMPLATES_DIR = DATA_DIR / "bb-templates"         # live-watched by BbTemplateDirectoryImporter
REFERENCE_DIR = DATA_DIR / "reference"
DICTIONARY_CACHE_FILE = REFERENCE_DIR / "field_mapping_cache.json"

DEFAULT_API_URL = "http://localhost:3001"   # pe-sub-api; matches PE_SUB_API_URL used platform-wide

# ── Heuristic constants — ported 1:1 from pe-sub-api's TemplateProfiler.java. Kept in sync
#    manually (no shared source between the Java and Python runtimes); if TemplateProfiler's
#    constants change, update these too so the two tools keep making the same judgment calls. ──
MIN_HEADER_MATCHES_DEFAULT = 3     # TemplateProfiler.MIN_HEADER_MATCHES
HEADER_SCAN_ROWS = 15              # TemplateProfiler.HEADER_SCAN_ROWS
MAX_GROUPS = 12                    # TemplateProfiler.MAX_GROUPS
MAX_ROWS_PER_SHEET_DEFAULT = 10_000  # not in the Java profiler (it caps via pe-sub-extraction's
                                      # /api/inspect row limit instead); this tool's own safety cap

AGENT_LABELS = {
    "agent bank", "agent", "administrative agent", "administered by", "prepared by", "lender",
}
# Lowercase, pre-normalized banner keywords (TemplateProfiler.SKIP_BANNERS).
SKIP_BANNER_NORMS = ("total", "subtotal", "sub total", "grand total", "sum", "net total")
# Verbatim default written into bb_template_tabs.skip_row_keywords (BbTemplateTab.java /
# BbTemplateImportService.DEFAULT_SKIP_KEYWORDS) — used as the seed list for every detected tab.
DEFAULT_SKIP_KEYWORDS = ["Total", "Subtotal", "Sub-Total", "Grand Total", "Sum", "Net Total"]

TAB_ROLES = ("LP_GRID", "CONCENTRATION", "CAPITAL_CALL", "TOP_SHEET")  # bb_template_tabs.tab_role CHECK
TEMPLATE_CLASSES = ("A", "B", "C")                                     # BbTemplateRequest @Pattern

# A tiny, explicitly-labeled offline safety net — NOT a substitute for the live dictionary.
# canonical -> a few of its most common Core-tier aliases (pe-sub-api fm_canonical_fields /
# fm_aliases carries ~38 canonical fields and 300+ aliases as of 2026-07; this bundle covers only
# the fields named in the PE Sub platform's own ingestion-pipeline notes). Run with
# --refresh-dictionary once pe-sub-api is reachable to replace this with the real registry.
BUNDLED_FALLBACK_FIELDS: dict[str, list[str]] = {
    "Investor Name": ["Investor Name", "Investor", "LP Name", "Limited Partner"],
    "Fund Sleeve": ["Fund Sleeve", "Fund Vehicle", "Feeder Fund", "Feeder"],
    "Parent / Sponsor": ["Parent / Sponsor", "Parent", "Sponsor"],
    "SPV Flag": ["SPV Flag", "SPV", "Special Purpose Vehicle"],
    "Region / Location": ["Region / Location", "Region", "Location", "Domicile"],
    "Investor Type": ["Investor Type", "Investor Classification", "Category of Investor"],
    "Institutional vs HNW": ["Institutional vs HNW", "HNW Flag", "Investor Segment"],
    "LP Category": ["LP Category", "Agent LP Classification", "LP Classification"],
    "Investment Grade Flag": ["Investment Grade Flag", "Investment Grade", "IG Flag"],
    "S&P Rating": ["S&P", "S&P Rating"],
    "Moody's Rating": ["Moody's", "Moody's Rating", "Moodys"],
    "Fitch Rating": ["Fitch", "Fitch Rating"],
    "Capital Commitments": ["Capital Commitments", "Committed Capital", "Total Commitment"],
    "% of Capital Commitments": ["% of Capital Commitments", "% Commitment"],
    "Called Capital": ["Called Capital", "Drawn Capital", "Funded Capital"],
    "Uncalled Capital": ["Uncalled Capital", "Unfunded Commitment", "Remaining Commitment"],
    "% of Uncalled Capital": ["% of Uncalled Capital", "% Uncalled", "Uncalled %"],
    "AUM": ["AUM", "Assets Under Management", "Total AUM"],
    "NAV": ["NAV", "Net Asset Value"],
    "Pension Assets": ["Pension Assets", "Pension Fund Assets"],
    "Borrowing Base": ["Borrowing Base", "Agent Borrowing Base", "Agent BB"],
    "Advance Rate": ["Advance Rate", "Agent Advance Rate", "Adv. Rate", "Rate"],
    "Concentration Limit": ["Concentration Limit", "Agent Concentration Limit", "Conc. Limit"],
    "Excess Concentration": ["Excess Concentration", "Conc. Overage"],
    "Notes": ["Notes", "Comments", "Remarks"],
}


# ==============================================================================================
#  Small self-describing value types (mirror TemplateProposal's shapes; see docstring above)
# ==============================================================================================
@dataclass
class ProfiledField:
    """Mirrors pe-sub-api TemplateProposal.ProfiledField<T>(value, confidence, evidence)."""
    value: object
    confidence: str   # "high" | "medium" | "low"
    evidence: str

    @staticmethod
    def of(value, confidence: str, evidence: str) -> "ProfiledField":
        return ProfiledField(value, confidence, evidence)


@dataclass
class ProposedGroup:
    header_text: str
    classification: str
    confidence: str
    row_index: int              # 0-based row this banner was found on (diagnostic only)


@dataclass
class SkipRowHit:
    row_index: int              # 0-based
    matched_keyword: str
    sample_text: str


@dataclass
class ColorFlag:
    style: str                  # e.g. "Fill #FFFF00" / "Font #0000FF"
    sample_cell: str            # e.g. "C42"
    occurrences: int


@dataclass
class TabAnalysis:
    sheet_name: str
    tab_sort: int
    header: ProfiledField                       # value = 0-based row index, or None if no grid found
    header_row_span: int
    columns: list[str] = field(default_factory=list)
    matched_canonical: int = 0
    matched_fields: list[dict] = field(default_factory=list)     # [{header, canonical, lpMasterField}]
    unmatched_headers: list[str] = field(default_factory=list)   # never silently dropped — listed
    groups: list[ProposedGroup] = field(default_factory=list)
    skip_row_keywords: list[str] = field(default_factory=list)
    skip_rows_detected: list[SkipRowHit] = field(default_factory=list)
    color_flags: list[ColorFlag] = field(default_factory=list)
    structure_fingerprint: tuple = field(default_factory=tuple)
    duplicate_of: Optional[str] = None
    notes: list[str] = field(default_factory=list)


@dataclass
class WorkbookAnalysis:
    file_name: str
    slug: ProfiledField = None
    agent_name: ProfiledField = None
    title: ProfiledField = None
    template_class: ProfiledField = None
    workbook_shape: ProfiledField = None         # "single" | "multiple" | "none"
    tabs: list[TabAnalysis] = field(default_factory=list)
    non_grid_sheets: list[str] = field(default_factory=list)
    auto_discover_tabs: ProfiledField = None
    detect_keys: list[str] = field(default_factory=list)
    overall_confidence: str = "low"
    dictionary_source: str = ""
    dictionary_field_count: int = 0
    truncated_sheets: list[str] = field(default_factory=list)
    analyzed_at: str = ""


# ==============================================================================================
#  Field Mapping Dictionary — live fetch (pe-sub-api), local cache, bundled fallback
# ==============================================================================================
@dataclass
class Dictionary:
    canonical_fields: list[str]           # ordered, display order
    alias_lookup: dict[str, str]          # normalize(alias) -> canonical field
    lp_master_field: dict[str, str]       # canonical field -> lp_master_field (may be "")
    source: str                           # "live:<url>" | "cache:<path>" | "bundled-fallback"


_WS_RE = re.compile(r"\s+")


def _normalize(s) -> str:
    """Lowercase, non-alnum -> space, whitespace collapsed. Identical to TemplateProfiler.normalize()."""
    if s is None:
        return ""
    return _WS_RE.sub(" ", re.sub(r"[^a-z0-9\s]", " ", str(s).lower())).strip()


def _flatten_alias_groups(raw: list[dict]) -> Dictionary:
    """raw = the JSON body of GET /api/field-mapping/alias-groups (or the cached copy of it):
    [{group, fields:[{canonical, lpMasterField, aliases:[{text, ...}]}]}]."""
    canonical_fields: list[str] = []
    alias_lookup: dict[str, str] = {}
    lp_master_field: dict[str, str] = {}
    for group in raw or []:
        for f in group.get("fields", []):
            canonical = f.get("canonical")
            if not canonical:
                continue
            canonical_fields.append(canonical)
            lp_master_field[canonical] = f.get("lpMasterField") or ""
            alias_lookup[_normalize(canonical)] = canonical
            for a in f.get("aliases", []):
                text = a.get("text")
                if text:
                    alias_lookup[_normalize(text)] = canonical
    return Dictionary(canonical_fields, alias_lookup, lp_master_field, source="")


def fetch_live_dictionary(api_url: str, timeout: float = 4.0) -> Optional[list[dict]]:
    """Returns the raw alias-groups JSON, or None on any failure (unreachable, non-200, bad JSON) —
    network unavailability is an expected, tolerated condition here, never a crash."""
    url = api_url.rstrip("/") + "/api/field-mapping/alias-groups"
    try:
        req = urllib.request.Request(url, headers={"Accept": "application/json"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            if resp.status != 200:
                return None
            return json.loads(resp.read().decode("utf-8"))
    except (urllib.error.URLError, TimeoutError, OSError, json.JSONDecodeError, ValueError):
        return None


def load_dictionary(api_url: str, offline: bool, refresh: bool, verbose: bool) -> Dictionary:
    """Reach-live -> cache -> bundled-fallback, in that order (skips live entirely if --offline
    unless --refresh-dictionary was also asked for, since refresh explicitly wants a live pull)."""
    def log(msg):
        if verbose:
            print(f"[dictionary] {msg}", file=sys.stderr)

    if not offline or refresh:
        raw = fetch_live_dictionary(api_url)
        if raw is not None:
            REFERENCE_DIR.mkdir(parents=True, exist_ok=True)
            DICTIONARY_CACHE_FILE.write_text(
                json.dumps({"fetched_at": datetime.now().isoformat(timespec="seconds"),
                            "source_url": api_url, "groups": raw}, indent=2),
                encoding="utf-8",
            )
            d = _flatten_alias_groups(raw)
            d.source = f"live:{api_url}"
            log(f"fetched live from {api_url} ({len(d.canonical_fields)} canonical fields); cached to {DICTIONARY_CACHE_FILE}")
            return d
        log(f"live fetch from {api_url} failed or timed out")

    if DICTIONARY_CACHE_FILE.is_file():
        try:
            cached = json.loads(DICTIONARY_CACHE_FILE.read_text(encoding="utf-8"))
            d = _flatten_alias_groups(cached.get("groups", []))
            d.source = f"cache:{DICTIONARY_CACHE_FILE} (fetched_at={cached.get('fetched_at', '?')})"
            log(f"using cached dictionary ({len(d.canonical_fields)} canonical fields)")
            return d
        except (json.JSONDecodeError, OSError):
            log("cache file present but unreadable; falling back to bundled defaults")

    alias_lookup = {_normalize(a): canon for canon, aliases in BUNDLED_FALLBACK_FIELDS.items() for a in aliases}
    for canon in BUNDLED_FALLBACK_FIELDS:
        alias_lookup[_normalize(canon)] = canon
    log(f"using bundled fallback dictionary ({len(BUNDLED_FALLBACK_FIELDS)} fields) — "
        f"this is a small safety net, not the real ~38-field registry; run with --refresh-dictionary "
        f"once pe-sub-api is reachable at {api_url}")
    return Dictionary(list(BUNDLED_FALLBACK_FIELDS), alias_lookup, {}, source="bundled-fallback")


def _cleanup_cache() -> None:
    """Remove the Field Mapping Dictionary cache file and __pycache__ after analysis completes.
    (Regenerable on next run; not kept between invocations.)"""
    if DICTIONARY_CACHE_FILE.is_file():
        try:
            DICTIONARY_CACHE_FILE.unlink()
        except OSError:
            pass  # If deletion fails, continue silently — cache is not load-critical.

    pycache_dir = SCRIPT_DIR / "__pycache__"
    if pycache_dir.is_dir():
        try:
            shutil.rmtree(pycache_dir)
        except OSError:
            pass  # If deletion fails, continue silently — not load-critical.


# ==============================================================================================
#  ExcelAnalyzer — structural + semantic analysis of a raw Agent BB workbook
# ==============================================================================================
class ExcelAnalyzer:
    """Detects LP-grid sheets, header rows, column->canonical mappings, LP-category group banners,
    and skip-row candidates in an uploaded Agent BB workbook. Every heuristic here has a direct
    counterpart in pe-sub-api's TemplateProfiler.java (see inline cross-references) so a human
    switching between this tool and the live /api/bb-templates/profile endpoint sees consistent
    judgment calls. Never raises on a single malformed sheet/row — flags and continues (D10 posture,
    same as lp_db_extract.py: a dirty file never aborts the run)."""

    def __init__(self, dictionary: Dictionary, *, min_header_matches: int = MIN_HEADER_MATCHES_DEFAULT,
                 max_rows_per_sheet: int = MAX_ROWS_PER_SHEET_DEFAULT, color_scan: bool = True,
                 verbose: bool = False):
        self.dictionary = dictionary
        self.min_header_matches = min_header_matches
        self.max_rows_per_sheet = max_rows_per_sheet
        self.color_scan = color_scan
        self.verbose = verbose

    def _log(self, msg: str) -> None:
        if self.verbose:
            print(f"[analyze] {msg}", file=sys.stderr)

    # ── Workbook-level entry point ──────────────────────────────────────────────────────────
    def analyze_workbook(self, path: Path, *, agent_bank_override: Optional[str] = None,
                          slug_override: Optional[str] = None,
                          template_class_override: Optional[str] = None) -> WorkbookAnalysis:
        read_only = not self.color_scan   # color/merged-cell inspection needs style access
        wb = openpyxl.load_workbook(path, read_only=read_only, data_only=True)

        analysis = WorkbookAnalysis(file_name=path.name, analyzed_at=datetime.now().isoformat(timespec="seconds"))
        analysis.dictionary_source = self.dictionary.source
        analysis.dictionary_field_count = len(self.dictionary.canonical_fields)

        first_grid_sheet_rows: Optional[list[list]] = None
        first_grid_header_row = -1
        seen_fingerprints: dict[tuple, str] = {}
        tab_sort = 0

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            rows, truncated = self._materialize_rows(ws)
            if truncated:
                analysis.truncated_sheets.append(sheet_name)
                self._log(f"sheet '{sheet_name}' truncated at {self.max_rows_per_sheet} rows for analysis")

            scan = self._find_header_row(rows)
            if scan is None:
                analysis.non_grid_sheets.append(sheet_name)
                self._log(f"sheet '{sheet_name}': no header row cleared min_header_matches={self.min_header_matches}; not treated as an LP grid")
                continue

            header_row_idx, matched_count = scan
            header_cells = self._non_blank(rows[header_row_idx]) if header_row_idx < len(rows) else []
            matched_fields, unmatched = self._map_columns(header_cells)

            tab_sort += 1
            fingerprint = tuple(sorted(_normalize(c) for c in header_cells))
            duplicate_of = seen_fingerprints.get(fingerprint)
            if duplicate_of is None:
                seen_fingerprints[fingerprint] = sheet_name

            tab = TabAnalysis(
                sheet_name=sheet_name,
                tab_sort=tab_sort,
                header=ProfiledField.of(header_row_idx, "high",
                                         f"alias-scored header row ({matched_count} canonical matches, threshold {self.min_header_matches})"),
                header_row_span=self._detect_header_span(ws, header_row_idx, len(header_cells), read_only),
                columns=header_cells,
                matched_canonical=matched_count,
                matched_fields=matched_fields,
                unmatched_headers=unmatched,
                groups=self._detect_groups(rows, header_row_idx),
                skip_row_keywords=list(DEFAULT_SKIP_KEYWORDS),
                skip_rows_detected=self._detect_skip_rows(rows, header_row_idx),
                structure_fingerprint=fingerprint,
                duplicate_of=duplicate_of,
            )
            if self.color_scan and not read_only:
                tab.color_flags = self._detect_color_flags(ws, header_row_idx, len(rows))
            if duplicate_of:
                tab.notes.append(
                    f"Column structure identical to sheet '{duplicate_of}' — confirm whether this is a "
                    f"genuine continuation/sleeve (keep both, in tabSort order) or an accidental duplicate export."
                )
            if unmatched:
                tab.notes.append(
                    f"{len(unmatched)} header(s) did not match any canonical field: {unmatched}. "
                    f"Add a Field Mapping alias (POST /api/field-mapping/aliases) or confirm these columns "
                    f"are intentionally out of scope for this template."
                )
            analysis.tabs.append(tab)

            if first_grid_sheet_rows is None:
                first_grid_sheet_rows = rows
                first_grid_header_row = header_row_idx

        n_tabs = len(analysis.tabs)
        analysis.workbook_shape = ProfiledField.of(
            "multiple" if n_tabs > 1 else "single",   # BbTemplateRequest has no third state; a 0-tab
                                                        # proposal is caught separately (see main()) —
                                                        # this value is a placeholder in that case, but
                                                        # its LOW confidence below says so honestly.
            "medium" if n_tabs > 1 else ("high" if n_tabs == 1 else "low"),
            f"{n_tabs} grid sheet(s) detected out of {len(wb.sheetnames)} total sheet(s)",
        )
        analysis.agent_name = self._infer_agent(first_grid_sheet_rows, first_grid_header_row, agent_bank_override)
        analysis.title = self._infer_title(first_grid_sheet_rows)
        slug = (slug_override or "").strip() or self._slug_from_filename(path.name)
        analysis.slug = ProfiledField.of(
            slug, "high" if slug_override else ("high" if slug else "low"),
            "operator-provided" if slug_override else "derived from filename",
        )
        analysis.template_class = ProfiledField.of(
            (template_class_override or "A"),
            "high" if template_class_override else "low",
            "operator-provided" if template_class_override
            else "default 'A' — template_class is a business/workflow classification, not a "
                 "structural signal; confirm with the template owner",
        )
        analysis.auto_discover_tabs = ProfiledField.of(
            False, "low",
            "assumed named tabs — confirm if this agent's tab names vary run-to-run (would need auto-discover)",
        )
        analysis.detect_keys = self._detect_keys(slug, analysis.title.value)
        analysis.overall_confidence = (
            "low" if not analysis.tabs
            else "high" if analysis.tabs[0].matched_canonical >= self.min_header_matches
            else "medium"
        )
        return analysis

    # ── Sheet materialization ───────────────────────────────────────────────────────────────
    def _materialize_rows(self, ws) -> tuple[list[list], bool]:
        """Reads all cell values into a plain list of lists (openpyxl read_only worksheets stream
        from XML — a single Python list at BB-workbook scale, hundreds of rows, is trivial after
        that; this trades a little memory for being able to re-scan the sheet freely below without
        juggling generator re-entrancy). Capped at max_rows_per_sheet as a large-file safety net."""
        rows: list[list] = []
        truncated = False
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i >= self.max_rows_per_sheet:
                truncated = True
                break
            rows.append(["" if c is None else c for c in row])
        return rows, truncated

    def _non_blank(self, cells: list) -> list[str]:
        return [str(c).strip() for c in cells if c is not None and str(c).strip() != ""]

    # ── Header detection (TemplateProfiler.findHeaderRow / matchesAlias) ───────────────────
    def _find_header_row(self, rows: list[list]) -> Optional[tuple[int, int]]:
        best_row, best_matched = -1, 0
        limit = min(HEADER_SCAN_ROWS, len(rows))
        for i in range(limit):
            matched = sum(1 for cell in rows[i] if cell and self._matches_alias(_normalize(cell)))
            if matched > best_matched:
                best_matched, best_row = matched, i
        return (best_row, best_matched) if best_matched >= self.min_header_matches else None

    def _matches_alias(self, norm_cell: str) -> bool:
        if not norm_cell:
            return False
        if norm_cell in self.dictionary.alias_lookup:
            return True
        for alias_norm in self.dictionary.alias_lookup:
            if not alias_norm:
                continue
            if len(alias_norm) >= 4 and " " in alias_norm and alias_norm in norm_cell:
                return True
            if len(norm_cell) >= 4 and " " in norm_cell and norm_cell in alias_norm:
                return True
        return False

    def _map_columns(self, header_cells: list[str]) -> tuple[list[dict], list[str]]:
        matched, unmatched = [], []
        for h in header_cells:
            norm = _normalize(h)
            canonical = self.dictionary.alias_lookup.get(norm)
            if canonical is None:
                for alias_norm, canon in self.dictionary.alias_lookup.items():
                    if alias_norm and ((len(alias_norm) >= 4 and " " in alias_norm and alias_norm in norm)
                                        or (len(norm) >= 4 and " " in norm and norm in alias_norm)):
                        canonical = canon
                        break
            if canonical:
                matched.append({"header": h, "canonical": canonical,
                                 "lpMasterField": self.dictionary.lp_master_field.get(canonical, "")})
            else:
                unmatched.append(h)
        return matched, unmatched

    # ── Group / LP-category banner detection (TemplateProfiler.detectGroups/classifyGroup) ──
    def _detect_groups(self, rows: list[list], header_row: int) -> list[ProposedGroup]:
        groups: list[ProposedGroup] = []
        for i in range(header_row + 1, len(rows)):
            if len(groups) >= MAX_GROUPS:
                break
            populated = self._non_blank(rows[i])
            if len(populated) != 1:
                continue
            text = populated[0]
            norm = _normalize(text)
            if not (3 <= len(norm) <= 60) or self._is_numeric(text) or any(b in norm for b in SKIP_BANNER_NORMS):
                continue
            classification, confidence = self._classify_group(norm, text)
            groups.append(ProposedGroup(text, classification, confidence, i))
        return groups

    def _classify_group(self, norm: str, verbatim: str) -> tuple[str, str]:
        if "exclud" in norm or "ineligible" in norm:
            return "Ineligible Investors", "medium"
        if "designated" in norm and "pwm" in norm:
            return "Designated PWM", "medium"
        if "designated" in norm or "institutional" in norm:
            return "Designated Institutional", "medium"
        if ("non" in norm and "rated" in norm) or "unrated" in norm:
            return "Non-Rated Included", "medium"
        if "rated" in norm or "includ" in norm:
            return "Rated Included", "medium"
        return verbatim, "low"   # verbatim self-map (e.g. a sleeve/feeder section header)

    # ── Row-skip detection — broader than banner-only: any row after the header whose FIRST
    #    populated cell contains a skip keyword, so a subtotal line embedded mid-grid (label +
    #    numeric columns still populated) is caught, not just a lone-cell banner. ──────────────
    def _detect_skip_rows(self, rows: list[list], header_row: int) -> list[SkipRowHit]:
        hits: list[SkipRowHit] = []
        for i in range(header_row + 1, len(rows)):
            populated = self._non_blank(rows[i])
            if not populated:
                continue
            first_norm = _normalize(populated[0])
            hit_kw = next((kw for kw in DEFAULT_SKIP_KEYWORDS if _normalize(kw) in first_norm), None)
            if hit_kw:
                hits.append(SkipRowHit(i, hit_kw, " | ".join(populated[:4])))
        return hits

    def _is_numeric(self, s: str) -> bool:
        t = re.sub(r"[,$%\s]", "", s)
        if not t:
            return False
        try:
            float(t)
            return True
        except ValueError:
            return False

    # ── Header row span via vertically-merged cells (openpyxl merged_cells; non-read-only only) ─
    def _detect_header_span(self, ws, header_row_idx0: int, header_col_count: int, read_only: bool) -> int:
        if read_only or header_col_count == 0:
            return 1
        try:
            excel_header_row = header_row_idx0 + 1  # merged-cell ranges are 1-based Excel rows
            span = 1
            for rng in ws.merged_cells.ranges:
                if rng.min_row <= excel_header_row <= rng.max_row and rng.max_row > rng.min_row:
                    if rng.min_col <= header_col_count:  # roughly within the header's column band
                        span = max(span, rng.max_row - rng.min_row + 1)
            return min(span, 3)   # a >3-row header is almost certainly an unrelated merge elsewhere
        except Exception:
            return 1   # style-API quirks must never abort the run — best-effort only

    # ── Color-flag / legend detection — best-effort, bounded to the detected grid's data rows ──
    def _detect_color_flags(self, ws, header_row_idx0: int, total_rows: int) -> list[ColorFlag]:
        try:
            signatures: dict[str, dict] = {}
            excel_start = header_row_idx0 + 2   # first data row, 1-based
            excel_end = min(total_rows, header_row_idx0 + 1 + 2000)  # bound the style scan too
            for r in range(excel_start, excel_end + 1):
                for c in range(1, 40):  # generous column cap; BB grids are rarely wider than this
                    cell = ws.cell(row=r, column=c)
                    if cell.value in (None, ""):
                        continue
                    sig = self._cell_color_signature(cell)
                    if sig is None:
                        continue
                    entry = signatures.setdefault(sig, {"count": 0, "sample": f"{get_column_letter(c)}{r}"})
                    entry["count"] += 1
            flags = [ColorFlag(sig, v["sample"], v["count"]) for sig, v in signatures.items()]
            flags.sort(key=lambda f: -f.occurrences)
            return flags[:8]   # cap — this is a hint list for the Legend sheet, not an audit trail
        except Exception:
            return []   # styling APIs vary by file; never let this abort the analysis

    def _cell_color_signature(self, cell) -> Optional[str]:
        try:
            fill = cell.fill
            if fill is not None and fill.patternType == "solid" and fill.fgColor is not None:
                rgb = getattr(fill.fgColor, "rgb", None)
                if isinstance(rgb, str) and rgb not in ("00000000", "FFFFFFFF"):
                    return f"Fill #{rgb[-6:]}"
        except Exception:
            pass
        try:
            font = cell.font
            if font is not None and font.color is not None:
                rgb = getattr(font.color, "rgb", None)
                if isinstance(rgb, str) and rgb not in ("FF000000", "00000000"):
                    return f"Font #{rgb[-6:]}"
        except Exception:
            pass
        return None

    # ── Agent / title / slug / detect-keys (TemplateProfiler.inferAgent/inferTitle/slugFromFilename) ──
    def _infer_agent(self, rows: Optional[list[list]], header_row: int, override: Optional[str]) -> ProfiledField:
        if override and override.strip():
            return ProfiledField.of(override.strip(), "high", "operator-provided (--agent-bank)")
        if rows:
            limit = header_row if header_row >= 0 else len(rows)
            for i in range(min(limit, len(rows))):
                cells = rows[i]
                for c, cell in enumerate(cells):
                    if _normalize(cell) in AGENT_LABELS:
                        for v in cells[c + 1:]:
                            if v not in (None, ""):
                                return ProfiledField.of(str(v).strip(), "medium",
                                                         f'summary-block label "{str(cells[c]).strip()}"')
        return ProfiledField.of(None, "low", "no agent label found — pass --agent-bank")

    def _infer_title(self, rows: Optional[list[list]]) -> ProfiledField:
        if rows:
            for i in range(min(2, len(rows))):
                for cell in rows[i]:
                    if cell not in (None, "") and len(str(cell).strip()) >= 4:
                        return ProfiledField.of(str(cell).strip(), "medium", f"first non-empty cell in row {i + 1}")
        return ProfiledField.of(None, "low", "no title cell found")

    def _slug_from_filename(self, file_name: str) -> str:
        base = re.sub(r"(?i)\.(xlsx|xls|csv)$", "", file_name or "")
        slug = re.sub(r"[^a-z0-9]+", "-", base.lower())
        slug = re.sub(r"^(agent-bb|bb-template-import|bb-template)-", "", slug)
        slug = slug.strip("-")
        if len(slug) > 50:   # bb_templates.template_slug is VARCHAR(50); trim at a word boundary
            slug = re.sub(r"-+[^-]*$", "", slug[:50])
        return slug

    def _detect_keys(self, slug: str, title: Optional[str]) -> list[str]:
        keys: list[str] = []
        if title and title.strip():
            keys.append(_normalize(title))
        if slug and slug.strip():
            keys.append(slug.replace("-", " "))
        seen = set()
        return [k for k in keys if k and not (k in seen or seen.add(k))]


# ==============================================================================================
#  TemplateBuilder — WorkbookAnalysis -> {"template": BbTemplateRequest-shaped, "analysis": {...}}
# ==============================================================================================
class TemplateBuilder:
    def __init__(self, analysis: WorkbookAnalysis):
        self.a = analysis

    def build_request(self) -> dict:
        """camelCase, exactly BbTemplateRequest's field set — paste as-is into a POST/PUT body."""
        a = self.a
        first = a.tabs[0] if a.tabs else None
        return {
            "templateSlug": a.slug.value,
            "templateName": a.slug.value,   # import convention: no separate display name (see BbTemplateImportService)
            "agentName": a.agent_name.value,
            "templateClass": a.template_class.value,
            "sheetName": first.sheet_name if first else None,
            "headerRowIndex": first.header.value if first else None,
            "autoLearned": True,   # this proposal's origin IS the heuristic tool, not a human author
            "trancheCount": max(1, len(a.tabs)),
            "hasGroupingRows": bool(first and first.groups),
            "hasColorFlags": bool(first and first.color_flags),
            "autoDiscoverTabs": a.auto_discover_tabs.value,
            "summaryRowsAboveHeader": first.header.value if first else 0,
            "summaryRowRange": f"1-{first.header.value}" if first and first.header.value else None,
            "titleRow": 1 if a.title.value else None,
            "titleText": a.title.value,
            "detectKeys": a.detect_keys,
            "legend": [{"style": f.style, "meaning": ""} for f in (first.color_flags if first else [])],
            "notes": [n for tab in a.tabs for n in tab.notes] + self._cross_role_group_caution(),
            "tabs": [self._tab_request(t) for t in a.tabs],
        }

    def _tab_request(self, t: TabAnalysis) -> dict:
        return {
            "tabRole": "LP_GRID",   # the only role this tool infers structurally; others are operator-assigned
            "tabSort": t.tab_sort,
            "sheetName": t.sheet_name,
            "sleeveName": t.sheet_name,
            "headerRowIndex": t.header.value,
            "headerRowSpan": t.header_row_span,
            "skipRowKeywords": t.skip_row_keywords,
            "columns": t.columns,
            "groups": [{"groupSort": i + 1, "headerText": g.header_text, "classification": g.classification}
                       for i, g in enumerate(t.groups)],
        }

    def _cross_role_group_caution(self) -> list[str]:
        """pe-sub-api's BB-Template-Import format (BbTemplateImportService.parseWorkbook) indexes the
        Groups sheet by tab_role ONLY, not by sheet name (`groupsByRole.getOrDefault(role, ...)` is
        attached to every Tabs row sharing that role). This tool only ever proposes tabRole=LP_GRID,
        so a >1-tab proposal where more than one tab carries groups would, if imported as-is, attach
        every LP_GRID tab's groups to every LP_GRID tab — a real characteristic of the production
        import format, not something to silently paper over here. Flag it; let the operator decide
        (keep one tab's groups, or reassign tabRole on the non-primary tabs before staging)."""
        tabs_with_groups = [t for t in self.a.tabs if t.groups]
        if len(self.a.tabs) > 1 and tabs_with_groups:
            return [
                f"{len(self.a.tabs)} tabs share tabRole=LP_GRID and {len(tabs_with_groups)} of them "
                f"have group banners — the BB-Template-Import format keys its Groups sheet by tab_role "
                f"only (see BbTemplateImportService.parseWorkbook), so importing this AS-IS would "
                f"attach every LP_GRID tab's groups to every LP_GRID tab. Confirm before staging, or "
                f"reassign tabRole on the non-primary tab(s)."
            ]
        return []

    def build_review_envelope(self) -> dict:
        a = self.a
        review_required: list[str] = []

        def flag(label: str, pf: ProfiledField):
            if pf.confidence in ("medium", "low"):
                review_required.append(f"[{pf.confidence}] {label}: {pf.value!r} — {pf.evidence}")

        flag("slug", a.slug); flag("agent name", a.agent_name); flag("title", a.title)
        flag("template_class", a.template_class); flag("workbook shape", a.workbook_shape)
        for t in a.tabs:
            for g in t.groups:
                if g.confidence != "high":
                    review_required.append(f"[{g.confidence}] tab '{t.sheet_name}' group \"{g.header_text}\" "
                                            f"-> classification \"{g.classification}\"")
            review_required.extend(f"[note] tab '{t.sheet_name}': {n}" for n in t.notes)
        review_required.extend(f"[note] {n}" for n in self._cross_role_group_caution())
        if a.truncated_sheets:
            review_required.append(f"[low] sheet(s) truncated at the row cap during analysis: {a.truncated_sheets} "
                                    f"— re-run with a higher --max-rows-per-sheet if data was cut off")

        return {
            "_readme": (
                "Two sections: 'template' is camelCase and matches pe-sub-api's BbTemplateRequest "
                "exactly (paste it into POST/PUT /api/bb-templates once reviewed). 'analysis' is this "
                "tool's own diagnostic report — confidence/evidence per field, which headers matched "
                "the Field Mapping Dictionary vs. did not, duplicate/continuation tab detection, and a "
                "flat review_required list of everything below high confidence. See the script's "
                "module docstring (parse_excel_templates.py) for the full format + scoring explanation."
            ),
            "template": self.build_request(),
            "analysis": {
                "source_file": a.file_name,
                "analyzed_at": a.analyzed_at,
                "overall_confidence": a.overall_confidence,
                "dictionary_source": a.dictionary_source,
                "dictionary_field_count": a.dictionary_field_count,
                "slug": asdict(a.slug),
                "agent_name": asdict(a.agent_name),
                "title": asdict(a.title),
                "template_class": asdict(a.template_class),
                "workbook_shape": asdict(a.workbook_shape),
                "auto_discover_tabs": asdict(a.auto_discover_tabs),
                "detect_keys": a.detect_keys,
                "non_grid_sheets": a.non_grid_sheets,
                "truncated_sheets": a.truncated_sheets,
                "tabs": [self._tab_analysis_dict(t) for t in a.tabs],
                "review_required": review_required,
            },
        }

    def _tab_analysis_dict(self, t: TabAnalysis) -> dict:
        return {
            "sheet_name": t.sheet_name, "tab_sort": t.tab_sort, "header": asdict(t.header),
            "header_row_span": t.header_row_span, "columns": t.columns,
            "matched_canonical": t.matched_canonical, "matched_fields": t.matched_fields,
            "unmatched_headers": t.unmatched_headers,
            "groups": [asdict(g) for g in t.groups],
            "skip_row_keywords": t.skip_row_keywords,
            "skip_rows_detected": [asdict(h) for h in t.skip_rows_detected],
            "color_flags": [asdict(c) for c in t.color_flags],
            "structure_fingerprint": list(t.structure_fingerprint),
            "duplicate_of": t.duplicate_of,
            "notes": t.notes,
        }

    # ── Companion output: the REAL BB-Template-Import-<slug>.xlsx meta-workbook, matching the
    #    exact sheet/column format BbTemplateImportService.parseWorkbook expects (verified against
    #    the existing data/bb-templates/BB-Template-Import-aep-vii.xlsx). --stage-import writes this
    #    into the live-watched directory; without that flag it can still be written elsewhere for
    #    a human to inspect/hand-edit before staging it themselves. ──────────────────────────────
    def write_import_workbook(self, path: Path) -> None:
        req = self.build_request()
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        ws = wb.create_sheet("Template")
        ws.append(["template_slug", "agent_bank", "template_class", "sheet_name", "header_row_index",
                   "header_row_span", "auto_learned", "tranche_count", "has_grouping_rows",
                   "has_color_flags", "auto_discover_tabs", "summary_rows_above_header",
                   "summary_row_range", "title_row", "title_text", "detect_keys"])
        first_tab = req["tabs"][0] if req["tabs"] else {}
        ws.append([
            req["templateSlug"], req["agentName"], req["templateClass"], req["sheetName"],
            (req["headerRowIndex"] + 1) if req["headerRowIndex"] is not None else None,   # back to 1-based
            first_tab.get("headerRowSpan", 1), req["autoLearned"], req["trancheCount"],
            req["hasGroupingRows"], req["hasColorFlags"], req["autoDiscoverTabs"],
            req["summaryRowsAboveHeader"], req["summaryRowRange"], req["titleRow"], req["titleText"],
            ", ".join(req["detectKeys"]),
        ])

        ws = wb.create_sheet("Tabs")
        ws.append(["tab_role", "tab_sort", "sheet_name", "sleeve_name", "header_row_index",
                   "header_row_span", "skip_row_keywords", "expected_source_headers_json"])
        for t in req["tabs"]:
            ws.append([t["tabRole"], t["tabSort"], t["sheetName"], t["sleeveName"],
                       (t["headerRowIndex"] + 1) if t["headerRowIndex"] is not None else None,
                       t["headerRowSpan"], ", ".join(t["skipRowKeywords"]), json.dumps(t["columns"])])

        ws = wb.create_sheet("Groups")
        ws.append(["tab_role", "group_sort", "header_text", "classification"])
        for t in req["tabs"]:
            for g in t["groups"]:
                ws.append([t["tabRole"], g["groupSort"], g["headerText"], g["classification"]])

        ws = wb.create_sheet("Columns")
        ws.append(["tab_role", "column_sort", "source_header"])
        for t in req["tabs"]:
            for i, col in enumerate(t["columns"], start=1):
                ws.append([t["tabRole"], i, col])

        ws = wb.create_sheet("Legend")
        ws.append(["legend_sort", "style", "meaning"])
        for i, entry in enumerate(req["legend"], start=1):
            ws.append([i, entry["style"], entry["meaning"]])

        ws = wb.create_sheet("Notes")
        ws.append(["note_sort", "note"])
        for i, note in enumerate(req["notes"], start=1):
            ws.append([i, note])
        if not req["notes"]:
            ws.append([1, "Auto-generated by parse_excel_templates.py — review every 'analysis.review_required' "
                          "entry in the companion JSON before treating this as verified. Set auto_learned=false "
                          "once confirmed."])

        path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(path)


# ==============================================================================================
#  Console / file summary report
# ==============================================================================================
def _kv(label: str, value: str, width: int = 20) -> str:
    return f"{label:<{width}}: {value}"


def render_summary(analysis: WorkbookAnalysis, envelope: dict) -> str:
    lines = [
        "Agent BB workbook analysis — summary",
        _kv("generated", analysis.analyzed_at),
        _kv("source file", analysis.file_name),
        _kv("dictionary source", f"{analysis.dictionary_source} ({analysis.dictionary_field_count} canonical fields)"),
        _kv("overall confidence", analysis.overall_confidence),
        _kv("proposed slug", f"{analysis.slug.value}  [{analysis.slug.confidence}]"),
        _kv("proposed agent name", f"{analysis.agent_name.value}  [{analysis.agent_name.confidence}]"),
        _kv("workbook shape", f"{analysis.workbook_shape.value}  [{analysis.workbook_shape.confidence}]"),
        "",
        _kv("sheets scanned", str(len(analysis.tabs) + len(analysis.non_grid_sheets))),
        _kv("  recognized as LP grids", f"{len(analysis.tabs)}  ({[t.sheet_name for t in analysis.tabs]})"),
        _kv("  not recognized", f"{len(analysis.non_grid_sheets)}  ({analysis.non_grid_sheets})"),
    ]
    for t in analysis.tabs:
        lines.append(
            f"\n  tab '{t.sheet_name}' (sort {t.tab_sort}): header row {t.header.value} "
            f"(0-based), {t.matched_canonical} canonical match(es), {len(t.unmatched_headers)} unmatched, "
            f"{len(t.groups)} group(s), {len(t.skip_rows_detected)} skip-row hit(s), "
            f"{len(t.color_flags)} color signature(s)"
        )
        if t.duplicate_of:
            lines.append(f"    ! duplicate structure of '{t.duplicate_of}' — confirm continuation vs. duplicate")
    review = envelope["analysis"]["review_required"]
    lines.append(f"\nreview_required ({len(review)} item(s)):")
    if review:
        lines.extend(f"  - {r}" for r in review)
    else:
        lines.append("  (none — every field cleared high confidence)")
    return "\n".join(lines)


# ==============================================================================================
#  CLI
# ==============================================================================================
def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="parse_excel_templates.py",
        description="Analyze a raw Agent BB or investor-list workbook and generate a BB template.",
    )
    p.add_argument("input", help="Path to the Excel workbook to analyze")
    return p


def _resolve_input(raw: str) -> Path:
    candidates = [Path(raw)]
    if not Path(raw).is_absolute():
        candidates += [Path.cwd() / raw, DEFAULT_IMPORT_DIR / raw, WATCHED_TEMPLATES_DIR / raw]
    for c in candidates:
        if c.is_file():
            return c
    raise SystemExit(
        f"Input workbook not found: {raw}\nTried: {[str(c) for c in candidates]}\n"
        f"(relative paths are also tried under {DEFAULT_IMPORT_DIR} and {WATCHED_TEMPLATES_DIR})"
    )


def main(argv: Optional[list[str]] = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass

    args = build_arg_parser().parse_args(argv)
    input_path = _resolve_input(args.input)

    print(f"Source: {input_path}")
    dictionary = load_dictionary(DEFAULT_API_URL, offline=False, refresh=False, verbose=False)

    analyzer = ExcelAnalyzer(dictionary)
    try:
        analysis = analyzer.analyze_workbook(input_path)
    except Exception as e:
        raise SystemExit(f"Failed to analyze '{input_path}': {e}") from e

    builder = TemplateBuilder(analysis)
    envelope = builder.build_review_envelope()
    summary = render_summary(analysis, envelope)
    print("\n" + summary)

    if not analysis.tabs:
        print("\nNo LP-grid sheet recognized in this workbook.", file=sys.stderr)
        _cleanup_cache()
        return 1

    slug = analysis.slug.value or "untitled"
    xlsx_path = WATCHED_TEMPLATES_DIR / f"BB-Template-Import-{slug}.xlsx"
    builder.write_import_workbook(xlsx_path)
    print(f"\nTemplate written: {xlsx_path}")
    _cleanup_cache()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
