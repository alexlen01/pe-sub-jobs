#!/usr/bin/env python3
r"""
Batch-process Agent BB Excel files organized by Agent Bank subdirectories.

Walks through a directory hierarchy (e.g., data/import/agent-bank-batch/), where each
subdirectory name is an Agent Bank name and contains individual Agent BB Excel files (.xlsx).
For each file found, analyzes it using ExcelAnalyzer and generates a BB-Template-Import-*.xlsx
template, with the subdirectory name injected as the Agent Bank value.

USAGE
    python parse_agent_bb_directory.py <input_directory>

Example:
    python parse_agent_bb_directory.py data/import/agent-bank-batch/

Input directory structure:
    data/import/agent-bank-batch/
    ├── Wells Fargo/
    │   ├── agent_bb_2026-07-15.xlsx
    │   └── additional_file.xlsx
    ├── Citi/
    │   └── bb-2026-07-20.xlsx
    └── ...

Output:
    All generated BB-Template-Import-*.xlsx files are written to data/bb-templates/
    (the directory watched by BbTemplateDirectoryImporter in pe-sub-jobs).
"""
from __future__ import annotations

import argparse
import logging
import re
import sys
from pathlib import Path
from typing import Optional

# Import the reusable components from parse_excel_templates.py (treated as a library module).
# This script does NOT duplicate ExcelAnalyzer, TemplateBuilder, or load_dictionary logic.
try:
    from parse_excel_templates import (
        ExcelAnalyzer,
        TemplateBuilder,
        load_dictionary,
        _cleanup_cache,
        WATCHED_TEMPLATES_DIR,
        DEFAULT_API_URL,
    )
except ImportError as e:
    sys.exit(f"Failed to import parse_excel_templates module: {e}\n"
             f"Ensure this script is run from pe-sub-jobs/scripts/ or PYTHONPATH includes it.")

# ==============================================================================================
#  Logging Setup
# ==============================================================================================
logger = logging.getLogger("parse_agent_bb_directory")


def _setup_logging(verbose: bool = False) -> None:
    """Configure logging to stderr."""
    level = logging.DEBUG if verbose else logging.INFO
    handler = logging.StreamHandler(sys.stderr)
    handler.setFormatter(
        logging.Formatter("[%(levelname)s] %(message)s")
    )
    logger.addHandler(handler)
    logger.setLevel(level)


# ==============================================================================================
#  Filename Slug Cleaning
# ==============================================================================================
def clean_filename_for_slug(file_name: str) -> str:
    """
    Parse out dates and 'Borrowing Base' prefixes from filenames to derive a clean slug.

    Examples:
        "2026-07-23-Agent-BB.xlsx" → "Agent-BB"
        "UBS-Borrowing-Base-2026.xlsx" → "UBS"
        "Wells-Fargo-BB-2026-06-15.xlsx" → "Wells-Fargo-BB"
        "agent-bb.xlsx" → "agent-bb"

    Strategy:
    1. Remove file extension (.xlsx, .xls, .csv)
    2. Remove leading dates (e.g., "2026-07-23-")
    3. Remove trailing dates (e.g., "-2026" or "-2026-06-15")
    4. Remove "Borrowing-Base" or "Borrowing Base" prefix/suffix (case-insensitive)
    5. Remove redundant separators, trim, and lowercase
    """
    # Remove extension
    base = re.sub(r"(?i)\.(xlsx|xls|csv)$", "", file_name or "")

    # Remove leading dates (YYYY-MM-DD or YYYY)
    base = re.sub(r"^\d{4}(?:-\d{2})?(?:-\d{2})?[-]?", "", base)

    # Remove trailing dates (e.g., "-2026" or "-2026-06-15" at the end)
    base = re.sub(r"[-]?\d{4}(?:-\d{2})?(?:-\d{2})?$", "", base)

    # Remove "Borrowing Base" or "Borrowing-Base" prefix (case-insensitive)
    base = re.sub(r"(?i)^borrowing[-\s]base[-\s]+", "", base)

    # Remove "Borrowing Base" or "Borrowing-Base" suffix (case-insensitive)
    base = re.sub(r"(?i)[-\s]+borrowing[-\s]base$", "", base)

    # Clean up: normalize to lowercase, collapse non-alphanum to "-", trim
    slug = re.sub(r"[^a-z0-9]+", "-", base.lower())
    slug = slug.strip("-")

    # Cap at 50 characters (bb_templates.template_slug is VARCHAR(50))
    if len(slug) > 50:
        slug = re.sub(r"-+[^-]*$", "", slug[:50])

    return slug or "untitled"


# ==============================================================================================
#  Directory Walking & Processing
# ==============================================================================================
def process_agent_bb_directory(
    input_dir: Path,
    api_url: str = DEFAULT_API_URL,
    offline: bool = False,
    verbose: bool = False,
) -> tuple[int, int]:
    """
    Walk through subdirectories and process each .xlsx file.

    Args:
        input_dir: Root directory containing Agent Bank subdirectories
        api_url: API URL for fetching field mappings
        offline: If True, skip live API calls; use cache or bundled fallback
        verbose: If True, print diagnostic messages

    Returns:
        (files_processed, files_failed): counts of successful and failed files
    """
    _setup_logging(verbose)

    logger.info(f"Starting batch process of {input_dir}")

    # Validate input directory
    if not input_dir.is_dir():
        raise SystemExit(f"Input directory does not exist: {input_dir}")

    # Load field mapping dictionary once (reused for all files)
    logger.info("Loading field mapping dictionary...")
    dictionary = load_dictionary(api_url, offline=offline, refresh=False, verbose=verbose)

    # Create analyzer and builder once (stateless; reused per file)
    analyzer = ExcelAnalyzer(dictionary, verbose=verbose)

    # Ensure output directory exists
    WATCHED_TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
    logger.info(f"Output directory: {WATCHED_TEMPLATES_DIR}")

    files_processed = 0
    files_failed = 0

    # Walk subdirectories (each is an Agent Bank name)
    for agent_dir in sorted(input_dir.iterdir()):
        if not agent_dir.is_dir():
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
            try:
                logger.info(f"  Processing: {xlsx_path.name}")

                # Analyze the workbook (with agent_bank override set to the subdirectory name)
                analysis = analyzer.analyze_workbook(
                    xlsx_path,
                    agent_bank_override=agent_name,  # Use subdirectory name as Agent Bank
                )

                # Check if any grids were detected
                if not analysis.tabs:
                    logger.warning(f"    → No LP-grid sheet recognized; skipping")
                    files_failed += 1
                    continue

                # Override the slug with cleaned filename
                cleaned_slug = clean_filename_for_slug(xlsx_path.name)
                analysis.slug.value = cleaned_slug
                analysis.slug.confidence = "high"
                analysis.slug.evidence = "cleaned from filename (dates/prefixes removed)"

                # Build template and write output
                builder = TemplateBuilder(analysis)
                output_path = WATCHED_TEMPLATES_DIR / f"BB-Template-Import-{cleaned_slug}.xlsx"
                builder.write_import_workbook(output_path)

                logger.info(f"    → Generated: {output_path.name}")
                files_processed += 1

            except Exception as e:
                logger.error(f"    → Failed: {e}", exc_info=verbose)
                files_failed += 1

    # Cleanup
    logger.info("\nCleaning up cache...")
    _cleanup_cache()

    logger.info(f"\nBatch process complete: {files_processed} succeeded, {files_failed} failed")
    return files_processed, files_failed


# ==============================================================================================
#  CLI
# ==============================================================================================
def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="parse_agent_bb_directory.py",
        description="Batch-process Agent BB Excel files organized by Agent Bank subdirectories.",
    )
    p.add_argument(
        "input_directory",
        help="Root directory containing Agent Bank subdirectories (e.g., data/import/agent-bank-batch/)",
    )
    p.add_argument(
        "--offline",
        action="store_true",
        help="Skip live API calls; use cached or bundled field mappings",
    )
    p.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable debug logging",
    )
    return p


def _resolve_input_directory(raw: str) -> Path:
    """
    Resolve the input directory path, trying multiple candidates.

    Candidates (in order):
    1. Absolute path (as given)
    2. Relative to current working directory
    3. Relative to pe-sub-jobs root (inferred from this script's location)
    """
    candidates = []

    # Try as absolute or relative to cwd
    p = Path(raw)
    if p.is_absolute():
        candidates.append(p)
    else:
        candidates.append(Path.cwd() / p)
        # Also try relative to pe-sub-jobs root (this script's parent's parent)
        script_dir = Path(__file__).resolve().parent
        jobs_root = script_dir.parent
        candidates.append(jobs_root / p)

    for candidate in candidates:
        if candidate.is_dir():
            return candidate

    raise SystemExit(
        f"Input directory not found: {raw}\n"
        f"Tried:\n" + "\n".join(f"  {c}" for c in candidates)
    )


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
        processed, failed = process_agent_bb_directory(
            input_dir,
            api_url=DEFAULT_API_URL,
            offline=args.offline,
            verbose=args.verbose,
        )

        # Exit with non-zero if any files failed (partial success is still a failure signal)
        return 1 if failed > 0 else 0

    except Exception as e:
        logger.error(f"Fatal error: {e}", exc_info=args.verbose)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
