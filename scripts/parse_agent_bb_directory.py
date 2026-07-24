#!/usr/bin/env python3
r"""
Batch-process Excel files organized by subdirectories.

Walks through a directory hierarchy where each subdirectory represents an Agent Bank and contains
individual Excel files (.xlsx/.xls). For each file, analyzes it and generates a
BB-Template-Import-*.xlsx template. Features robust data validation to skip Legend, Notes, and
metadata rows.

All generated templates are collected to <input_directory>/bb_templates/ along with a
run_summary.csv log file tracking processing results.

USAGE
    python parse_agent_bb_directory.py <input_directory>

Example:
    python parse_agent_bb_directory.py ~/AgentBBs

Input directory structure:
    ~/AgentBBs/
    ├── Wells Fargo/
    │   ├── agent_bb_2026-07-15.xlsx
    ├── JPM/
    │   └── bb-2026-07-20.xlsx
    └── ...

Output:
    ~/AgentBBs/bb_templates/
    ├── BB-Template-Import-agent-bb-2026.xlsx
    ├── BB-Template-Import-bb-2026-07-20.xlsx
    └── run_summary.csv

Log file format (CSV):
    Agent Bank,Source File,Template Generated,Output File,Status,Notes
    Wells Fargo,agent_bb_2026-07-15.xlsx,YES,BB-Template-Import-agent-bb-2026.xlsx,SUCCESS,
    JPM,bb-2026-07-20.xlsx,YES,BB-Template-Import-bb-2026-07-20.xlsx,SUCCESS,
    BoA,CP-VII-2026.xlsx,NO,,SKIPPED,No LP-grid sheet recognized
"""
from __future__ import annotations

import argparse
import csv
import logging
import re
import sys
from pathlib import Path
from typing import Optional

# Import the reusable components from parse_excel_templates.py
try:
    from parse_excel_templates import (
        ExcelAnalyzer,
        TemplateBuilder,
        load_dictionary,
        Dictionary,
    )
except ImportError as e:
    sys.exit(f"Failed to import parse_excel_templates module: {e}\n"
             f"Ensure this script is in the same directory as parse_excel_templates.py")

# ==============================================================================================
#  Logging Setup
# ==============================================================================================
logger = logging.getLogger("parse_agent_bb_directory")


def _setup_logging() -> None:
    """Configure logging to stderr."""
    handler = logging.StreamHandler(sys.stderr)
    handler.setFormatter(
        logging.Formatter("[%(levelname)s] %(message)s")
    )
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)


# ==============================================================================================
#  Filename Slug Cleaning
# ==============================================================================================

# Common large-bank prefixes (all normalized to lowercase for matching)
BANK_PREFIXES = {
    "wells fargo", "wells-fargo",
    "citi", "citigroup",
    "jpmorgan", "jpmorgan chase", "jp morgan chase",
    "bank of america", "bofa", "bofa bank",
    "us bank", "usbank", "u.s. bank",
    "pnc", "pnc bank",
    "td bank", "td", "toronto dominion",
    "truist", "bb&t",
    "capital one",
    "charles schwab",
    "goldman sachs",
    "morgan stanley",
    "bank of new york mellon", "bny mellon",
    "blackrock",
    "vanguard",
    "fidelity",
}


def clean_filename_for_slug(file_name: str) -> str:
    """
    Parse out dates, bank-name prefixes, and 'Borrowing Base' prefixes from filenames
    to derive a clean slug.

    Examples:
        "2026-07-23-Citi-Borrowing-Base-Agent-BB.xlsx" → "Agent-BB"
        "Wells-Fargo-2026-Q3-BB.xlsx" → "Q3"
        "JPMorgan-Chase-Borrowing-Base-2026-07-23-Facility-XYZ.xlsx" → "Facility-XYZ"
        "UBS-Borrowing-Base-2026.xlsx" → "UBS"
        "agent-bb.xlsx" → "agent-bb"

    Strategy:
    1. Remove file extension (.xlsx, .xls, .csv)
    2. Remove dates in flexible formats:
       - Compact: YYYYMMDD (e.g., 20250630)
       - Dot: YYYY.MM.DD (e.g., 2025.06.30)
       - Dash: YYYY-MM-DD (e.g., 2025-06-30)
       - Slash: YYYY/MM/DD or MM/DD/YYYY (e.g., 2025/06/30)
       - Quarters: Q1-2026, Q4-2026, etc.
    3. Remove "Borrowing Base" term anywhere
    4. Remove bank-name prefixes at start (Wells Fargo, Citi, JPMorgan, etc.)
    5. Remove leading "Agent BB" or "BB" prefix (indicator terms, not content)
    6. Normalize to lowercase, collapse separators to dashes
    7. Cap at 50 characters (bb_templates.template_slug VARCHAR(50))
    """
    # Remove extension
    base = re.sub(r"(?i)\.(xlsx|xls|csv)$", "", file_name or "")

    # Step 1: Remove dates in flexible formats
    # Remove compact format (YYYYMMDD, e.g., 20250630) - must be surrounded by separators or at boundaries
    base = re.sub(r"(?:^|[-\s\.])\d{8}(?=[-\s\.]|$)", " ", base)

    # Remove dot format (YYYY.MM.DD, e.g., 2025.06.30)
    base = re.sub(r"\d{4}\.\d{2}\.\d{2}", " ", base)

    # Remove ISO format (YYYY-MM-DD or YYYY/MM/DD)
    base = re.sub(r"(?:^|[-\s])\d{4}(?:[-/]\d{1,2})?(?:[-/]\d{1,2})?(?=[-\s]|$)", " ", base)

    # Remove US format (MM/DD/YYYY)
    base = re.sub(r"(?:^|[-\s])(?:0?[1-9]|1[0-2])/(?:0?[1-9]|[12]\d|3[01])/\d{4}(?=[-\s]|$)", " ", base)

    # Remove quarters (Q1-2026, Q4 2026, etc.)
    base = re.sub(r"(?i)(?:^|[-\s])[Qq]\d(?:[-\s])?\d{4}(?=[-\s]|$)", " ", base)

    # Remove year-only patterns (YYYY at boundaries, e.g., "Fund 2026" → "Fund")
    base = re.sub(r"(?:^|[-\s])\d{4}(?=[-\s]|$)", " ", base)

    # Step 2: Remove all instances of "Borrowing Base" (anywhere in string)
    base = re.sub(r"(?i)borrowing\s*[-]?\s*base", " ", base)

    # Step 2.5: Normalize and trim (collapse multiple spaces/dashes, trim edges)
    # This helps us properly detect prefixes in step 3
    base = re.sub(r"[-_/\s]+", " ", base)
    base = base.strip()

    # Step 3: Remove common bank-name prefixes (at START, followed by separator)
    # Sort by length descending to match longer names first (e.g., "JPMorgan Chase" before "JPMorgan")
    base_lower = base.lower()
    for bank_prefix in sorted(BANK_PREFIXES, key=len, reverse=True):
        escaped_prefix = re.escape(bank_prefix)
        # Match bank prefix at start, followed by space (now normalized)
        pattern = f"^{escaped_prefix}(?:\\s+|$)"
        if re.match(pattern, base_lower):
            base = re.sub(f"(?i)^{escaped_prefix}\\s+", "", base)
            break

    # Step 4: Remove leading "Borrowing Base" indicators ONLY if followed by more content
    # "Agent BB", "Agent-BB", "BB-Template", "BB-Import", etc. — but keep if it's the only content
    base_before_step4 = base
    base = re.sub(r"(?i)^(?:agent\s+bb|bb)\s+", "", base)
    # If we removed everything, restore it
    if not base or base.isspace():
        base = base_before_step4

    # Step 4.5: Remove trailing standalone "BB" (but not if part of "Agent-BB")
    # Only remove if BB is at end and not preceded by "Agent"
    if not re.search(r"(?i)agent\s+bb\s*$", base):
        base = re.sub(r"(?i)\s+bb\s*$", "", base)

    # Step 5: Final normalization
    # Collapse multiple spaces/dashes, then normalize to kebab-case
    slug = base.strip()
    slug = re.sub(r"[-_/\s]+", "-", slug)  # Collapse separators to single dash
    slug = re.sub(r"[^a-z0-9\-]", "", slug.lower())  # Keep only alphanumeric and dash
    slug = re.sub(r"-+", "-", slug)  # Collapse multiple dashes
    slug = slug.strip("-")

    # Cap at 50 characters (bb_templates.template_slug is VARCHAR(50))
    if len(slug) > 50:
        slug = re.sub(r"-+[^-]*$", "", slug[:50])

    return slug or "untitled"


# ==============================================================================================
#  Collision Avoidance
# ==============================================================================================
def resolve_output_path(slug: str, output_dir: Path) -> tuple[Path, str]:
    """
    Resolve output path with collision avoidance.

    If `BB-Template-Import-{slug}.xlsx` already exists in output_dir, appends a numeric
    suffix: `-1`, `-2`, etc. until a unique filename is found.

    Returns:
        (resolved_path, collision_suffix_str) where collision_suffix_str is "" if no collision,
        or "-1", "-2", etc. if collision was detected and resolved
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"BB-Template-Import-{slug}.xlsx"
    output_path = output_dir / base_name

    if not output_path.exists():
        return output_path, ""

    # Collision detected; find a unique suffix
    suffix_num = 1
    while True:
        collision_suffix = f"-{suffix_num}"
        alternate_name = f"BB-Template-Import-{slug}{collision_suffix}.xlsx"
        alternate_path = output_dir / alternate_name
        if not alternate_path.exists():
            return alternate_path, collision_suffix
        suffix_num += 1
        # Safety cap to prevent infinite loops (unlikely but defensive)
        if suffix_num > 100:
            raise SystemExit(f"Too many collisions for slug '{slug}' (tried up to -{suffix_num})")


# ==============================================================================================
#  Directory Walking & Processing
# ==============================================================================================
def process_agent_bb_directory(input_dir: Path) -> tuple[int, int, int]:
    """
    Walk through subdirectories and process each .xlsx file.

    All generated templates are collected to input_dir/bb_templates/. A summary log file
    (run_summary.csv) is written to the same directory.

    Args:
        input_dir: Root directory containing subdirectories with Excel files (each subdirectory
                   is treated as an Agent Bank)

    Returns:
        (files_processed, files_skipped, files_failed): counts of successful, skipped, and failed files
    """
    _setup_logging()

    logger.info(f"Starting batch process of {input_dir}")

    # Validate input directory
    if not input_dir.is_dir():
        raise SystemExit(f"Input directory does not exist: {input_dir}")

    # Create output directory
    output_dir = input_dir / "bb_templates"
    output_dir.mkdir(parents=True, exist_ok=True)
    logger.info(f"Output directory: {output_dir}")

    # Load field mapping dictionary once (reused for all files)
    logger.info("Loading field mappings...")
    dictionary = load_dictionary()

    # Create analyzer once (stateless; reused per file)
    analyzer = ExcelAnalyzer(dictionary)

    files_processed = 0
    files_skipped = 0
    files_failed = 0
    log_records: list[dict] = []

    # Walk subdirectories (each is an Agent Bank name)
    # Skip Python cache and virtual environment directories
    skip_dirs = {"__pycache__", ".pytest_cache", ".venv", "venv", "env", ".env", "node_modules", ".git", "bb_templates"}
    for agent_dir in sorted(input_dir.iterdir()):
        if not agent_dir.is_dir():
            continue
        if agent_dir.name in skip_dirs:
            continue

        agent_name = agent_dir.name
        logger.info(f"\nProcessing Agent Bank: {agent_name}")

        # Find all .xlsx files in this subdirectory
        xlsx_files = sorted(agent_dir.glob("*.xlsx")) + sorted(agent_dir.glob("*.xls"))

        if not xlsx_files:
            logger.warning(f"  No Excel files found in {agent_dir}")
            continue

        logger.info(f"  Found {len(xlsx_files)} file(s)")

        for xlsx_path in xlsx_files:
            source_filename = xlsx_path.name
            log_entry = {
                "Agent Bank": agent_name,
                "Source File": source_filename,
                "Template Generated": "NO",
                "Output File": "",
                "Status": "UNKNOWN",
                "Notes": "",
            }

            try:
                logger.info(f"  Processing: {source_filename}")

                # Analyze the workbook (with agent_bank override set to the subdirectory name)
                analysis = analyzer.analyze_workbook(
                    xlsx_path,
                    agent_bank_override=agent_name,  # Use subdirectory name as Agent Bank
                )

                # Check if any grids were detected
                if not analysis.tabs:
                    logger.warning(f"    → No LP-grid sheet recognized; skipping")
                    log_entry["Status"] = "SKIPPED"
                    log_entry["Notes"] = "No LP-grid sheet recognized"
                    log_records.append(log_entry)
                    files_skipped += 1
                    continue

                # Override the slug with cleaned filename
                cleaned_slug = clean_filename_for_slug(source_filename)
                analysis.slug.value = cleaned_slug
                analysis.slug.confidence = "high"
                analysis.slug.evidence = "cleaned from filename (dates/prefixes removed)"

                # Build template and write output to bb_templates
                builder = TemplateBuilder(analysis)
                output_path, collision_suffix = resolve_output_path(cleaned_slug, output_dir)
                builder.write_import_workbook(output_path)

                output_filename = output_path.name
                if collision_suffix:
                    logger.warning(
                        f"    → Collision avoided: '{cleaned_slug}.xlsx' already exists; "
                        f"using suffix '{collision_suffix}' → {output_filename}"
                    )
                    log_entry["Notes"] = f"Collision resolved with suffix {collision_suffix}"
                else:
                    logger.info(f"    → Generated: {output_filename}")

                log_entry["Template Generated"] = "YES"
                log_entry["Output File"] = output_filename
                log_entry["Status"] = "SUCCESS"
                log_records.append(log_entry)
                files_processed += 1

            except Exception as e:
                logger.error(f"    → Failed: {e}")
                log_entry["Status"] = "FAILED"
                log_entry["Notes"] = str(e)
                log_records.append(log_entry)
                files_failed += 1

    # Write summary log file
    log_file_path = output_dir / "run_summary.csv"
    _write_summary_log(log_file_path, log_records, files_processed, files_skipped, files_failed)

    logger.info(
        f"\nBatch process complete: {files_processed} succeeded, {files_skipped} skipped, {files_failed} failed"
    )
    logger.info(f"Summary log written to: {log_file_path}")
    return files_processed, files_skipped, files_failed


def _write_summary_log(
    log_path: Path,
    records: list[dict],
    processed: int,
    skipped: int,
    failed: int,
) -> None:
    """
    Write processing results to a CSV summary log file.

    Args:
        log_path: Path to the output CSV file
        records: List of processing log entries
        processed: Count of successful templates
        skipped: Count of skipped files
        failed: Count of failed files
    """
    fieldnames = ["Agent Bank", "Source File", "Template Generated", "Output File", "Status", "Notes"]

    try:
        with open(log_path, "w", newline="", encoding="utf-8") as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

            # Write header
            writer.writeheader()

            # Write all records
            writer.writerows(records)

            # Write summary stats as comments (CSV-readable as empty rows with notes in last column)
            csvfile.write("\n")
            csvfile.write(f"# Total Files Processed: {processed}\n")
            csvfile.write(f"# Total Files Skipped: {skipped}\n")
            csvfile.write(f"# Total Files Failed: {failed}\n")
            csvfile.write(f"# Total Files: {processed + skipped + failed}\n")

        logger.info(f"Summary log written: {processed} SUCCESS, {skipped} SKIPPED, {failed} FAILED")

    except Exception as e:
        logger.error(f"Failed to write summary log: {e}")


# ==============================================================================================
#  CLI
# ==============================================================================================
def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="parse_agent_bb_directory.py",
        description="Batch-process Excel files organized by subdirectories. All templates are collected to input_directory/bb_templates/.",
    )
    p.add_argument(
        "input_directory",
        help="Root directory containing subdirectories with Excel files (each subdirectory is treated as an Agent Bank)",
    )
    return p


def _resolve_input_directory(raw: str) -> Path:
    """
    Resolve the input directory path (absolute or relative to current working directory).
    """
    p = Path(raw)
    if p.is_absolute():
        if p.is_dir():
            return p
    else:
        relative = Path.cwd() / p
        if relative.is_dir():
            return relative

    raise SystemExit(f"Input directory not found: {raw}")


def main(argv: Optional[list[str]] = None) -> int:
    """Main entry point."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass

    args = build_arg_parser().parse_args(argv)
    input_dir = _resolve_input_directory(args.input_directory)

    print(f"Input directory: {input_dir}")

    try:
        processed, skipped, failed = process_agent_bb_directory(input_dir)

        # Exit with non-zero if any files failed (partial success is still a failure signal)
        return 1 if failed > 0 else 0

    except Exception as e:
        logger.error(f"Fatal error: {e}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
