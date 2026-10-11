"""
USPTO PATENT COLLECTION PIPELINE  (v15.9)
=========================================

v15.9: PDF links for reissue (RE) and defensive publication (T)
    patents (retry with --pdf --retry-errors-only).

v15.8: HTTP connection pool sized to the worker count (no more
    'Connection pool is full' warnings); API key no longer sent
    with PDF requests.

v15.7: --pdf shows complete/total per section before starting;
    workbook-loading summaries now go to pipeline.log only.

v15.6: APS numbers whose check character is "&" (e.g.
    "US06333061&") now get correct PDF links too.

v15.5: --pdf skips previously FAILED patents; retry them with
    --pdf --retry-errors-only once everything else is done.

v15.4: APS PDFs -- drops the Green Book check digit from APS
    numbers before requesting the PDF (was HTTP 404 for all APS).

v15.3: --pdf accepts --format, --limit, --skip-errors and
    --retry-errors-only, like --database.

v15.2: --filter --format <product> filters one database at a
    time (per-product CSVs, merged into filtered_patents.csv).

v15.1: fixes patents being lost when USPTO writes a record's
    <?xml ...?> declaration on the same line as the previous
    document's closing tag (seen after sequence listings);
    adds --requeue-merged to reparse files already affected.

v15 CHANGES (database stage -- see README_pipeline_v15.md):
    - One SQLite file per product (APS, SGML, XML grants,
      applications) instead of one combined file.
    - Metadata and full text in separate tables; claims and
      description stored zlib-compressed.
    - Workers stream records straight into SQLite in small
      batches instead of returning whole files' worth of
      parsed text to the main process (the memory blow-up
      behind the XML-stage crashes).
    - --filter reads full text only for CPC matches.
    - --migrate-db copies an existing v14 database across.

Purpose:
    1. Find USPTO bulk-data products (four of them, downloaded
       and parsed separately since they're different formats:
         - PTGRAPS  (Patent Grants, APS/Green Book, ~1976-1996)
         - PTGRSGM  (Patent Grants, SGML, ~1996-2001)
         - PTGRXML  (Patent Grants, XML, ~2001-present)
         - APPXML   (Patent Applications, XML, 2001-present)
    2. Download files for each product
    3. Extract downloaded archives
    4. Parse patent metadata/full text (three real parsers:
       parse_xml_file, parse_sgml_file, parse_legacy_file)
    5. Build a local SQLite database
    6. Filter patents using CPC codes (with an IPC fallback for
       pre-2013 records, which never carry real CPC) and
       keywords -- a patent must match BOTH to be selected
    7. Download the actual PDF for each selected patent, via
       USPTO's public Patent Public Search image server (no
       API key needed for this specific step -- see
       get_pdf_url() for the details and known gaps)

IMPORTANT:
    - This program is designed to be restartable at every
      stage. Downloads, database inserts, and PDF fetches all
      skip work that's already done rather than redoing it.
    - Failed downloads/PDFs are logged to their manifest CSVs
      with the actual error, not silently dropped.
    - Bulk data is kept locally under BULK_DATA_ROOT_DIR /
      DATABASE_ROOT_DIR / PDF_ROOT_DIR (can be the same drive,
      or split across separate physical drives -- see the
      storage-locations config near the top of this file).
    - PDFs are only downloaded AFTER filtering -- this stage
      only ever touches the (much smaller) set of patents that
      already passed the CPC/keyword filter.

YOU MUST FILL IN:
    1. USPTO_API_KEY
       (only needed for the bulk-download/database stages --
       the PDF stage doesn't use it)
    2. CPC_EXCEL_INPUT_FOLDER / CPC_COLUMN / CPC_START_ROW
       (every .xlsx file directly inside this folder is
       picked up automatically -- CPC codes are merged and
       de-duplicated across all of them)
    3. KEYWORDS_SHEET_NAME / KEYWORDS_COLUMN / KEYWORDS_START_ROW
       (keywords are read from the "Keywords" sheet of the
       same workbook)

HOW TO RUN (see main() / --help for the full list):
    --test        Confirm the USPTO API key/connection work
    --sample      Download a small, date-spread sample per
                  product -- use this before a full --download
                  to sanity-check the parsers against real data
    --download    Download the FULL corpus for all 4 products
    --extract     Extract downloaded zip archives
    --database    Parse extracted files into the SQLite database
    --filter      Apply CPC + keyword filtering, write
                  filtered_patents.csv
    --pdf         Download PDFs for the filtered patents
    --all         Run download -> extract -> database -> filter
                  -> pdf in one pass (only after --sample and
                  --test have both looked correct)

KNOWN GAPS (not silently ignored -- see the relevant docstrings):
    - get_pdf_url() doesn't yet have a confirmed URL format for
      reissue ("RE") or Statutory Invention Registration ("H")
      patents; those fail cleanly rather than guessing.
    - Documents over ~1000 pages may not be retrievable through
      the PDF endpoint used here (per USPTO's own docs); not
      independently confirmed what the failure looks like, but
      it will fail cleanly (caught by the existing retry/PDF
      validation logic) rather than corrupt anything.

Recommended:
    - Python 3.9+ (tested on 3.9 and 3.12)
    - 8 TB external drive
    - SSD preferred for parsing/database operations
"""

# ============================================================
# IMPORTS
# ============================================================

import os
import sys
import shutil
import re
import csv
import json
import time
import sqlite3
import logging
import argparse
import zipfile
import zlib
import threading
import xml.etree.ElementTree as ET
from html.parser import HTMLParser
from pathlib import Path
from datetime import datetime
from urllib.parse import quote
from concurrent.futures import ThreadPoolExecutor, ProcessPoolExecutor, as_completed
from concurrent.futures.process import BrokenProcessPool

import requests
from tqdm import tqdm
import openpyxl
from openpyxl.utils import column_index_from_string


# ============================================================
# USER CONFIGURATION
# ============================================================

# ------------------------------------------------------------
# !!! YOU ENTER THIS !!!
# ------------------------------------------------------------
# Put your USPTO Open Data Portal API key here.
#
# BETTER OPTION:
# Set it as an environment variable instead:
#
# Windows PowerShell:
#     $env:USPTO_API_KEY="YOUR_KEY"
#
# Then leave this as None.
# ------------------------------------------------------------

USPTO_API_KEY = "YOUR_API_KEY_HERE"
# Example:
# USPTO_API_KEY = "YOUR_API_KEY_HERE"


# ------------------------------------------------------------
# !!! YOU ENTER THIS !!!
# ------------------------------------------------------------
# CRITICAL: all three of these must point at an EXTERNAL
# drive -- never a path under C:\Users\... or anywhere inside
# a OneDrive/Dropbox/Google Drive-synced folder. This pipeline
# writes a genuinely large volume of data (potentially
# multiple TB across the full historical corpus). Pointing it
# at a cloud-synced folder means Windows will try to upload
# ALL of it to the cloud in the background, on top of writing
# it locally -- consuming your internal drive space, your
# cloud storage quota, and a large amount of upload bandwidth,
# all for data you almost certainly don't want backed up to
# the cloud anyway.
#
# Each category below CAN point at a different external drive
# if you have more than one, since different stages have
# different I/O needs:
#
#   BULK_DATA_ROOT_DIR -> RAW (downloaded zips) + EXTRACTED
#                         (parsed-out XML/SGML/text files).
#                         Large, mostly sequential writes/reads
#                         -- a cheap external HDD is fine.
#
#   DATABASE_ROOT_DIR   -> the SQLite database. The ONE stage
#                         that genuinely benefits from an SSD's
#                         random-I/O speed (WAL checkpointing,
#                         index construction) -- everywhere
#                         else, drive type barely matters.
#
#   PDF_ROOT_DIR         -> downloaded patent PDFs. Individually
#                         small files, but download SPEED is
#                         capped by the self-imposed PPUBS rate
#                         limit (~1/second regardless of worker
#                         count -- see
#                         PPUBS_MIN_REQUEST_INTERVAL_SECONDS),
#                         so disk speed is never the bottleneck
#                         here. An HDD keeps up just as well as
#                         an SSD would -- no need to reserve SSD
#                         space for this.
#
# Currently configured for a SINGLE external drive (all three
# point at the same path) -- reasonable for a one-time run
# where a bit of extra time on the database stage is an
# acceptable tradeoff. If you later want the database on a
# separate, faster SSD (worth it mainly if you expect to
# rebuild the database repeatedly while iterating on your
# CPC/keyword matrices), just point DATABASE_ROOT_DIR at a
# different drive letter -- everything else needs no changes.
#
# Update the drive letter below to match whatever your actual
# external drive is assigned once connected (check File
# Explorer).
#
# RESULTS/MANIFEST/LOGS are small either way and always live
# under BULK_DATA_ROOT_DIR.
# ------------------------------------------------------------

BULK_DATA_ROOT_DIR = Path(r"D:\uspto_corpus")

DATABASE_ROOT_DIR = BULK_DATA_ROOT_DIR

PDF_ROOT_DIR = BULK_DATA_ROOT_DIR


# ------------------------------------------------------------
# !!! YOU ENTER THIS !!!
# ------------------------------------------------------------
# CPC codes for your first filtering stage are read directly
# from your polymer-film search matrix Excel workbook(s).
#
# Rather than listing individual files, this points at a
# FOLDER -- every .xlsx file directly inside it is picked up
# automatically. To add a new workbook to the filter, just
# save it into this folder; no code change needed. Codes and
# keywords are merged (de-duplicated) across ALL of them, so
# you can split your search matrix across multiple files (e.g.
# one per technology area) instead of maintaining one giant
# workbook.
#
# Rather than pointing at one fixed sheet within each workbook,
# every worksheet whose TAB NAME looks like a CPC class/
# subclass (a letter, two digits, then a letter -- e.g.
# "B29C", "C08J", optionally with trailing digits like
# "B29C48") is treated as a CPC sheet and its codes are pulled
# in. Sheets that don't match that pattern (e.g. "Notes",
# "Summary", "Keywords", a title/matrix sheet) are ignored.
#
# Change these settings to match your setup:
#
#   CPC_EXCEL_INPUT_FOLDER -> folder containing your .xlsx
#                             workbook(s)
#   CPC_COLUMN             -> the column letter holding CPC
#                             codes on each class sheet (same
#                             across all files -- if a future
#                             file uses a different column,
#                             this would need to become a
#                             per-file setting instead)
#   CPC_START_ROW           -> the first row containing an
#                             actual code (i.e. the row after
#                             your header row)
#
# NOTE: unlike BULK_DATA_ROOT_DIR/DATABASE_ROOT_DIR/
# PDF_ROOT_DIR above, this folder only ever holds your small
# Excel workbooks (tens of KB each) -- not a storage-capacity
# concern either way. Defaulted to the same external drive
# below purely for consistency (so literally nothing writes to
# C:\ or OneDrive), but it's genuinely fine to instead keep
# this alongside the script on your internal drive if you'd
# rather manage/back up your search-matrix workbooks with your
# code instead of on the external drive.
# ------------------------------------------------------------

CPC_EXCEL_INPUT_FOLDER = Path(
    r"D:\uspto_corpus\input"
)

CPC_COLUMN = "E"

CPC_START_ROW = 2


def discover_excel_files(folder):

    """
    Finds every .xlsx file directly inside `folder` (not
    subfolders -- kept non-recursive so an unrelated workbook
    buried somewhere else can't get picked up by accident).

    Skips Excel's temporary lock files (named like
    "~$Some_File.xlsx", created while that workbook is open
    elsewhere) and hidden/dotfiles, neither of which are real
    workbooks and would otherwise fail to parse with a
    confusing error.

    Sorted for deterministic, repeatable ordering across runs.
    """

    if not folder.exists():

        raise RuntimeError(

            "\nCPC EXCEL INPUT FOLDER NOT FOUND:\n"
            f"{folder}\n\n"
            "Update CPC_EXCEL_INPUT_FOLDER near the top of "
            "this program to point at the folder containing "
            "your search-matrix workbook(s).\n"

        )

    files = sorted(

        path

        for path in folder.glob("*.xlsx")

        if not path.name.startswith("~$")
        and not path.name.startswith(".")

    )

    if not files:

        raise RuntimeError(

            "\nNO .xlsx FILES FOUND IN:\n"
            f"{folder}\n\n"
            "Save your CPC/keyword search-matrix workbook(s) "
            "into this folder.\n"

        )

    return files


CPC_EXCEL_FILES = discover_excel_files(
    CPC_EXCEL_INPUT_FOLDER
)

# Pattern used to recognize a CPC class/subclass sheet name,
# e.g. "B29C", "C08J", "B29C48". Non-matching sheet names
# (summary tabs, notes, etc.) are skipped automatically.
CPC_SHEET_NAME_PATTERN = re.compile(
    r"^[A-Z]\d{2}[A-Z]\d*$"
)


# ------------------------------------------------------------
# Reads the CPC code list out of every workbook in
# CPC_EXCEL_FILES, merging and de-duplicating across all of
# them.
#
# For each workbook, scans every worksheet whose tab name
# matches CPC_SHEET_NAME_PATTERN. On each matching sheet, reads
# CPC_COLUMN from CPC_START_ROW through the last used row,
# skipping blank cells. Non-matching ("non-class") sheets are
# ignored entirely. Codes are de-duplicated and upper-cased
# across ALL workbooks combined (e.g. "b29c" -> "B29C"), so the
# same class appearing in two different files (e.g. both a
# general matrix and a material-specific matrix) contributes
# codes from both rather than one shadowing the other.
#
# A workbook contributing zero CPC sheets isn't treated as an
# error on its own (it might be a keywords-only file, or one
# whose class sheets simply happen to be empty placeholders --
# both are real cases seen in practice) -- only the FINAL
# combined code list being empty is an error.
# ------------------------------------------------------------

def load_cpc_codes_from_excel():

    codes = []

    seen = set()

    any_class_sheets_found = False

    for excel_file in CPC_EXCEL_FILES:

        if not excel_file.exists():

            raise RuntimeError(

                "\nCPC EXCEL FILE NOT FOUND:\n"
                f"{excel_file}\n\n"
                "Update CPC_EXCEL_FILES near the top of this "
                "program to point at your search-matrix "
                "workbook(s).\n"

            )

        workbook = openpyxl.load_workbook(
            excel_file,
            data_only=True,
            read_only=True,
        )

        class_sheet_names = [
            name
            for name in workbook.sheetnames
            if CPC_SHEET_NAME_PATTERN.match(
                name.strip().upper()
            )
        ]

        if not class_sheet_names:

            logging.warning(

                "No CPC class sheets found in %s (tab names "
                "found: %s). Continuing with the remaining "
                "workbook(s) in CPC_EXCEL_FILES.",
                excel_file,
                ", ".join(workbook.sheetnames),

            )

            workbook.close()

            continue


        any_class_sheets_found = True

        column_index = column_index_from_string(
            CPC_COLUMN
        )

        sheets_with_codes = 0

        codes_from_this_file = 0

        for sheet_name in class_sheet_names:

            sheet = workbook[sheet_name]

            max_row = sheet.max_row or 0

            found_on_this_sheet = 0

            for row_number in range(
                CPC_START_ROW,
                max_row + 1
            ):

                cell_value = sheet.cell(
                    row=row_number,
                    column=column_index,
                ).value

                if cell_value is None:

                    continue

                code = str(cell_value).strip().upper()

                # Strip ALL internal whitespace too, not just
                # leading/trailing -- Excel entries commonly
                # have a space (regular or non-breaking, e.g.
                # copy-pasted from an official CPC listing PDF)
                # between the class letters and the main-group
                # number (e.g. "B29C 35/00" or "B29C\xa035/00"),
                # but extracted patent CPC/IPC values never have
                # that space (e.g. "B29C35/02"). Python's \s
                # already matches non-breaking spaces (U+00A0)
                # as well as regular ones, so this one line
                # handles both. Since matches_cpc() uses
                # startswith(), even one stray space guarantees
                # every single comparison fails silently -- this
                # bit us for real (0 matches across 15,909 real
                # records) before this fix.

                code = re.sub(r"\s+", "", code)

                if not code:

                    continue

                if code not in seen:

                    seen.add(code)

                    codes.append(code)

                    codes_from_this_file += 1

                found_on_this_sheet += 1

            if found_on_this_sheet:

                sheets_with_codes += 1

        workbook.close()

        _workbook_log.info(

            "%s: %d CPC class sheet(s) detected (%s), %d of "
            "them contained codes, %d new codes added",
            excel_file.name,
            len(class_sheet_names),
            ", ".join(class_sheet_names),
            sheets_with_codes,
            codes_from_this_file,

        )


    if not codes:

        raise RuntimeError(

            "\nNO CPC CODES FOUND.\n\n"
            f"Files checked: "
            f"{', '.join(str(f) for f in CPC_EXCEL_FILES)}\n"
            f"Column: {CPC_COLUMN}\n"
            f"Start row: {CPC_START_ROW}\n\n"
            + (
                "None of the files had a sheet whose tab name "
                "matched the CPC class pattern (letter + 2 "
                "digits + letter, e.g. 'B29C').\n\n"
                if not any_class_sheets_found
                else "Class sheets were found but none of them "
                "contained a code in the configured column/row "
                "-- check CPC_COLUMN and CPC_START_ROW.\n\n"
            )
            + "Check that these settings match your workbook(s)."

        )

    return codes


# ------------------------------------------------------------
# !!! YOU ENTER THIS !!!
# ------------------------------------------------------------
# Keywords for your second filtering stage are read directly
# from EVERY workbook in CPC_EXCEL_FILES, each from a sheet
# named "Keywords" -- merged and de-duplicated across all of
# them, same as the CPC codes.
#
# Change these settings if your layout differs:
#
#   KEYWORDS_SHEET_NAME  -> the worksheet tab holding keywords
#   KEYWORDS_COLUMN      -> the column letter holding keywords
#   KEYWORDS_START_ROW   -> the first row containing an actual
#                           keyword (i.e. the row after your
#                           header row)
# ------------------------------------------------------------

KEYWORDS_SHEET_NAME = "Keywords"

KEYWORDS_COLUMN = "A"

KEYWORDS_START_ROW = 2


# ------------------------------------------------------------
# Reads the keyword list out of every workbook in
# CPC_EXCEL_FILES, sheet KEYWORDS_SHEET_NAME.
#
# For each workbook, scans KEYWORDS_COLUMN from
# KEYWORDS_START_ROW through the last used row, skipping blank
# cells and de-duplicating case-insensitively (across ALL
# workbooks combined) while preserving the first-seen casing
# (matches_keywords() already lower-cases for comparison).
#
# A workbook missing the Keywords sheet entirely is logged as
# a warning and skipped, not a hard failure -- only having zero
# keywords across every workbook combined is an error.
# ------------------------------------------------------------

def load_keywords_from_excel():

    keywords = []

    seen_lowercase = set()

    for excel_file in CPC_EXCEL_FILES:

        if not excel_file.exists():

            raise RuntimeError(

                "\nKEYWORD EXCEL FILE NOT FOUND:\n"
                f"{excel_file}\n\n"
                "Update CPC_EXCEL_FILES near the top of this "
                "program to point at your search-matrix "
                "workbook(s).\n"

            )

        workbook = openpyxl.load_workbook(
            excel_file,
            data_only=True,
            read_only=True,
        )

        if KEYWORDS_SHEET_NAME not in workbook.sheetnames:

            logging.warning(

                "No '%s' sheet found in %s (tab names found: "
                "%s). Continuing with the remaining workbook(s) "
                "in CPC_EXCEL_FILES.",
                KEYWORDS_SHEET_NAME,
                excel_file,
                ", ".join(workbook.sheetnames),

            )

            workbook.close()

            continue

        sheet = workbook[KEYWORDS_SHEET_NAME]

        column_index = column_index_from_string(
            KEYWORDS_COLUMN
        )

        max_row = sheet.max_row or 0

        keywords_from_this_file = 0

        for row_number in range(

            KEYWORDS_START_ROW,
            max_row + 1

        ):

            cell_value = sheet.cell(

                row=row_number,
                column=column_index,

            ).value

            if cell_value is None:

                continue

            keyword = str(cell_value).strip()

            if not keyword:

                continue

            lowercase_keyword = keyword.lower()

            if lowercase_keyword in seen_lowercase:

                continue

            seen_lowercase.add(lowercase_keyword)

            keywords.append(keyword)

            keywords_from_this_file += 1

        workbook.close()

        _workbook_log.info(

            "%s: %d new keyword(s) added from sheet '%s'",
            excel_file.name,
            keywords_from_this_file,
            KEYWORDS_SHEET_NAME,

        )


    if not keywords:

        raise RuntimeError(

            "\nNO KEYWORDS FOUND.\n\n"
            f"Files checked: "
            f"{', '.join(str(f) for f in CPC_EXCEL_FILES)}\n"
            f"Sheet: {KEYWORDS_SHEET_NAME}\n"
            f"Column: {KEYWORDS_COLUMN}\n"
            f"Start row: {KEYWORDS_START_ROW}\n\n"
            "Check that these settings match your workbook(s)."

        )

    return keywords


# ------------------------------------------------------------
# DOWNLOAD SETTINGS
# ------------------------------------------------------------

MAX_RETRIES = 5

RETRY_DELAY_SECONDS = 10

CHUNK_SIZE = 1024 * 1024

REQUEST_TIMEOUT = 300

# HTTP status codes that will NEVER succeed no matter how many
# times we retry (bad URL, not found, forbidden). Fail fast on
# these instead of burning the full retry budget.
PERMANENT_FAILURE_STATUS_CODES = {400, 401, 403, 404, 410}

# HTTP status codes that are worth retrying with an escalating
# backoff (rate limiting / temporary unavailability).
TRANSIENT_FAILURE_STATUS_CODES = {429, 500, 502, 503, 504}

# Number of concurrent workers for bulk-file downloads
# (large files, moderate file count -> keep this modest).
BULK_DOWNLOAD_WORKERS = 4

# Number of concurrent workers for PDF downloads. This is the
# high-volume stage (potentially 500,000+ files), so this is
# the setting most worth tuning for your connection/USPTO's
# rate limits.
PDF_DOWNLOAD_WORKERS = 8

# Number of PROCESSES (not threads) for --database's parsing
# stage. Parsing is CPU-bound -- building an ElementTree DOM
# per XML/SGML record, running keyword regexes over full
# claims/description text -- so threads wouldn't help here
# (the GIL serializes CPU-bound work across threads regardless
# of count); this needs separate processes to actually use
# more than one core.
#
# Fixed at 4 for a 6-core / 12-thread, 16 GB machine: the
# os.cpu_count() default caused WinError 1450 (resource
# exhaustion) against the external drive. Each worker now uses
# roughly one record batch of RAM (tens of MB) rather than a
# whole file's worth, so memory is no longer what limits this --
# heat and drive throughput are. Override per run with
# --workers N (e.g. --workers 3 if the machine runs hot).
#
# WINDOWS NOTE: ProcessPoolExecutor uses the "spawn" start
# method on Windows, which re-imports this module in every
# worker process (including each replacement worker -- see
# DATABASE_WORKER_MAX_TASKS). The Excel loads are guarded so
# workers skip them; the remaining import cost is ~1-2 seconds.
DATABASE_PARSE_WORKERS = 4

# A worker PROCESS dying outright (crashed, or killed by the OS
# under resource pressure -- see WinError 1450 in the module
# notes) breaks the ENTIRE ProcessPoolExecutor it belonged to,
# not just the one file that worker was on: every other future
# still pending in that same pool immediately raises
# BrokenProcessPool too, even though those files never actually
# got a chance to run. build_database() detects this and opens
# a fresh pool to keep going on whatever's left, rather than
# cascading through a wall of false "errors" and requiring a
# whole separate --database invocation to pick them back up.
# This caps how many times it'll do that in a row before giving
# up on the rest for this run (they're simply left unmarked in
# the manifest, so they're picked up normally on the next
# --database run regardless).
DATABASE_MAX_POOL_RESTARTS = 5

# ------------------------------------------------------------
# PPUBS PDF ENDPOINT PACING
# ------------------------------------------------------------
# image-ppubs.uspto.gov (used by get_pdf_url()/download_pdf())
# is NOT part of USPTO's documented, key-authenticated ODP API
# family -- there's no published rate limit to design around,
# unlike the bulk-data endpoints elsewhere in this file. It has
# no robots.txt either (checked directly), which means no
# explicit "keep out" signal, but that's not the same as a
# confirmed-safe volume.
#
# This enforces a minimum interval between requests to this
# specific endpoint, SHARED across all worker threads, so
# PDF_DOWNLOAD_WORKERS running in parallel doesn't turn into a
# burst of simultaneous requests. 1.0 second -> roughly 60
# requests/minute in aggregate, deliberately matching the
# documented rate limit of a closely related USPTO API family
# as a reasonable, conservative default in the absence of
# official guidance for this exact endpoint. Test at small
# scale before trusting a faster setting at 500,000+ file
# volume.

PPUBS_MIN_REQUEST_INTERVAL_SECONDS = 1.0

# Number of concurrent workers for zip extraction. Decompression
# is CPU/IO bound rather than network bound, so keep this modest
# relative to your CPU core count and disk throughput.
# 2 rather than 4 (v15): extraction is disk-bound, and with RAW
# and EXTRACTED on the same spinning drive, more concurrent
# writers mostly add seek thrash rather than throughput.
EXTRACTION_WORKERS = 2

# How often (in completed items) to flush progress to the
# manifest CSV during large parallel runs. Saving after every
# single item is too slow at 500,000+ file scale; saving only
# at the very end risks losing a lot of progress on a crash.
MANIFEST_SAVE_INTERVAL = 200

# --sample mode: how many files to pull per product. These are
# spread evenly across the product's available date range so a
# small sample still spans different USPTO format eras
# (APS/Green Book, SGML, and multiple XML generations) instead
# of just grabbing the N most recent files, which would all be
# the same modern format.
SAMPLE_FILES_PER_PRODUCT = 3


# ------------------------------------------------------------
# PDF SETTINGS
# ------------------------------------------------------------

# We will only download PDFs after filtering.

PDF_DOWNLOAD_ENABLED = True

# Every downloaded PDF is validated by checking for the
# "%PDF-" magic header before it's accepted. This catches
# HTML error/login pages that come back with a 200 status
# but aren't actually PDFs.
PDF_MAGIC_BYTES = b"%PDF-"


# ============================================================
# DIRECTORY STRUCTURE
# ============================================================

RAW_DIR = BULK_DATA_ROOT_DIR / "RAW"

GRANTS_DIR = RAW_DIR / "GRANTS"

# USPTO exposes patent grants as THREE separate bulk-data
# products (not one product spanning all formats/dates), each
# with its own product identifier:
#   PTGRAPS  - APS/Green Book (older, roughly pre-1996)
#   PTGRSGM  - SGML (roughly 1996-2001)
#   PTGRXML  - XML (roughly 2002-present)
# Downloading each into its own subdirectory keeps them
# unambiguous regardless of what file extension (or lack of
# one) a given product actually uses.
GRANTS_APS_DIR = GRANTS_DIR / "APS"

GRANTS_SGML_DIR = GRANTS_DIR / "SGML"

GRANTS_XML_DIR = GRANTS_DIR / "XML"

APPLICATIONS_DIR = RAW_DIR / "APPLICATIONS"

EXTRACTED_DIR = BULK_DATA_ROOT_DIR / "EXTRACTED"

DATABASE_DIR = DATABASE_ROOT_DIR / "DATABASE"

PDF_DIR = PDF_ROOT_DIR / "PDF"

# PDFs are organized into subfolders by document_type, so
# grants and applications never end up mixed together in one
# flat folder. "OTHER" catches anything that isn't cleanly one
# or the other (e.g. Statutory Invention Registrations, which
# are neither a real grant nor a pending application) rather
# than silently forcing it into the wrong bucket.
PDF_GRANTS_DIR = PDF_DIR / "GRANTS"

PDF_APPLICATIONS_DIR = PDF_DIR / "APPLICATIONS"

PDF_OTHER_DIR = PDF_DIR / "OTHER"

RESULTS_DIR = BULK_DATA_ROOT_DIR / "RESULTS"

MANIFEST_DIR = BULK_DATA_ROOT_DIR / "MANIFEST"

LOG_DIR = BULK_DATA_ROOT_DIR / "LOGS"


# ------------------------------------------------------------
# DATABASE FILES -- ONE SQLITE FILE PER PRODUCT (v15)
# ------------------------------------------------------------
# v14 and earlier wrote every product into one ever-growing
# USPTO_PATENTS.sqlite. v15 splits it into one file per
# product. Reasons:
#   - A run only ever touches the file for the product it's
#     working on, so a crash / bad sector / corruption can only
#     damage that one product's data, not all four.
#   - Each file stays a fraction of the size, so checkpoints,
#     index builds, integrity checks, and (optional) backups
#     are proportionally cheaper.
#   - A product can be rebuilt from scratch (delete its file +
#     its manifest rows) without touching the others.
#   - --filter simply reads each file in turn; nothing about
#     the results changes.
#
# The OLD single-file database is only read by --migrate-db
# (see migrate_legacy_database()), which copies its rows into
# these per-product files so already-parsed APS/SGML/XML data
# doesn't have to be re-parsed.
DATABASE_FILES = {
    "aps": DATABASE_DIR / "USPTO_APS.sqlite",
    "sgml": DATABASE_DIR / "USPTO_SGML.sqlite",
    "xml": DATABASE_DIR / "USPTO_XML_GRANTS.sqlite",
    "appxml": DATABASE_DIR / "USPTO_APPLICATIONS.sqlite",
}

# The v14-and-earlier single database. Only used by --migrate-db.
LEGACY_DATABASE_FILE = DATABASE_DIR / "USPTO_PATENTS.sqlite"

# Optional per-product backups (only written with --backup).
DATABASE_BACKUP_DIR = DATABASE_DIR / "BACKUP"

# ------------------------------------------------------------
# TEXT COMPRESSION
# ------------------------------------------------------------
# Claims and description are 95%+ of the database's size and
# are only ever read by --filter, and only for patents that
# already passed the CPC gate. zlib level 3 shrinks this text
# roughly 3-4x at a cost of well under a millisecond per
# patent (tiny next to the XML parse itself). Smaller database
# = far less disk I/O on the external drive, in both --database
# and --filter. Existing uncompressed text (e.g. rows copied by
# --migrate-db with this set to False) is still read correctly
# -- _decompress_text() handles both.
COMPRESS_TEXT = True

TEXT_COMPRESSION_LEVEL = 3

# ------------------------------------------------------------
# INSERT BATCHING / MEMORY BOUNDS
# ------------------------------------------------------------
# Each worker parses records as a STREAM and writes them to
# SQLite every DB_INSERT_BATCH_RECORDS records. This is the
# main memory fix in v15: previously a worker parsed an ENTIRE
# weekly file (a modern XML grant file is ~6,000-8,000 patents
# and several hundred MB of text) into one Python list, pickled
# it, and piped the whole thing back to the main process --
# several GB of RAM per in-flight file across the workers plus
# the main process. Now peak memory per worker is roughly one
# batch (tens of MB), regardless of file size.
DB_INSERT_BATCH_RECORDS = 250

# A worker process is retired and replaced after this many
# files. Long-lived Python processes that churn through GBs of
# short-lived strings fragment their heap and slowly creep up
# in RAM even without a true leak; recycling caps that.
# Only applied on Python 3.11+ (the option doesn't exist
# before that); on 3.9/3.10 workers simply aren't recycled.
DATABASE_WORKER_MAX_TASKS = 20

# How long a worker waits for another worker's write to finish
# before giving up. Writes are short (one batch), so contention
# is brief -- this is a generous safety margin, not an expected
# wait.
DATABASE_BUSY_TIMEOUT_MS = 180_000

# Save the database manifest every N completed files. Each
# worker commits its own rows before reporting back, so the
# manifest never claims a file that isn't safely on disk.
DATABASE_MANIFEST_SAVE_INTERVAL = 10

DOWNLOAD_MANIFEST = (
    MANIFEST_DIR /
    "download_manifest.csv"
)

PDF_MANIFEST = (
    MANIFEST_DIR /
    "pdf_manifest.csv"
)

FILTER_RESULTS = (
    RESULTS_DIR /
    "filtered_patents.csv"
)

DATABASE_MANIFEST = (
    MANIFEST_DIR /
    "database_manifest.csv"
)

# Forward-citation cross-referencing (which OTHER patents cite
# each filtered patent) now lives entirely in the standalone
# uspto_forward_citations_v1.py script, which reads
# filtered_patents.csv above and scans the local extracted
# corpus directly. This replaces an earlier --citations stage
# that queried PatentsView's API for a citation COUNT only --
# PatentsView's live query API was confirmed shut down in
# March 2026, and even before that its citation endpoint was
# never an officially documented one. The offline script gives
# the actual citing patent IDs (not just a count), their dates,
# and an examiner/applicant/other breakdown, from data already
# on disk -- strictly more than PatentsView ever provided, with
# nothing to depend on staying online.


# ============================================================
# CREATE DIRECTORIES
# ============================================================

for directory in [

    BULK_DATA_ROOT_DIR,
    DATABASE_ROOT_DIR,
    PDF_ROOT_DIR,

    RAW_DIR,

    GRANTS_DIR,
    GRANTS_APS_DIR,
    GRANTS_SGML_DIR,
    GRANTS_XML_DIR,

    APPLICATIONS_DIR,

    EXTRACTED_DIR,

    DATABASE_DIR,

    PDF_DIR,

    PDF_GRANTS_DIR,
    PDF_APPLICATIONS_DIR,
    PDF_OTHER_DIR,

    RESULTS_DIR,

    MANIFEST_DIR,

    LOG_DIR,

]:

    directory.mkdir(
        parents=True,
        exist_ok=True
    )


# ============================================================
# LOGGING
# ============================================================

logging.basicConfig(

    filename=LOG_DIR / "pipeline.log",

    level=logging.INFO,

    format=(
        "%(asctime)s | "
        "%(levelname)s | "
        "%(message)s"
    ),

)

console = logging.StreamHandler()

console.setLevel(logging.INFO)

# Workbook-loading summaries (sheets detected, codes/keywords
# added, totals loaded) go to pipeline.log only -- they print on
# every run otherwise. Warnings/errors from loading still show.
WORKBOOK_LOGGER_NAME = "uspto.workbooks"

_workbook_log = logging.getLogger(WORKBOOK_LOGGER_NAME)

console.addFilter(
    lambda record: not (
        record.name == WORKBOOK_LOGGER_NAME
        and record.levelno < logging.WARNING
    )
)

logging.getLogger().addHandler(console)


# ============================================================
# LOAD CPC CODES FROM EXCEL
# ============================================================
#
# Guarded by __name__ == "__main__" rather than running
# unconditionally: ProcessPoolExecutor uses the "spawn" start
# method on Windows (see DATABASE_PARSE_WORKERS above), which
# re-imports this entire module in every worker process at
# pool startup. Without this guard, that means re-running the
# full openpyxl parse of every CPC/keyword workbook once per
# worker for no reason -- the parsing workers (see
# _parse_file_into_database()) never call matches_cpc() or
# matches_keywords() at all; those only run in the main process
# during --filter. So workers get cheap empty defaults instead,
# and only the main process pays for the real Excel parse.
#
# This relies on the same mechanism that already keeps
# workers from re-running main() itself: Python sets __name__
# to "__mp_main__" (not "__main__") in a spawned worker's
# re-imported copy of this module, specifically so guards like
# this one don't fire a second time. If that weren't true,
# ProcessPoolExecutor on Windows would already be recursively
# re-spawning main() on every worker -- which it demonstrably
# isn't, since that guard is what makes ProcessPoolExecutor
# usable here in the first place.

if __name__ == "__main__":

    CPC_CODES = load_cpc_codes_from_excel()

    _workbook_log.info(
        "Loaded %d total CPC codes from %d workbook(s): %s "
        "(column %s, starting row %d)",
        len(CPC_CODES),
        len(CPC_EXCEL_FILES),
        ", ".join(f.name for f in CPC_EXCEL_FILES),
        CPC_COLUMN,
        CPC_START_ROW,
    )

    # Precomputed once for matches_cpc(): a plain tuple lets us
    # use str.startswith(tuple_of_prefixes), which is a single
    # call instead of looping over CPC_CODES in Python for
    # every code on every patent.
    CPC_CODES_UPPER = tuple(
        code.upper()
        for code in CPC_CODES
    )

else:

    # Worker process (or this module imported as a library
    # rather than run directly) -- skip the Excel parse
    # entirely. Never read by parsing workers; matches_cpc()
    # only runs in the main process, where the branch above
    # already set the real values before ProcessPoolExecutor
    # is ever created.
    CPC_CODES = []
    CPC_CODES_UPPER = ()

# Patents typically carry more than one CPC classification.
# When the parsers populate the `cpc` column, multiple codes
# are expected to be joined with one of these delimiters.
CPC_CODE_SPLIT_PATTERN = re.compile(
    r"[;,|]+"
)


# ============================================================
# LOAD KEYWORDS FROM EXCEL
# ============================================================
#
# Same reasoning and same guard as CPC_CODES above.

if __name__ == "__main__":

    KEYWORDS = load_keywords_from_excel()

    _workbook_log.info(
        "Loaded %d total keywords from %d workbook(s): %s "
        "(sheet '%s', column %s, starting row %d)",
        len(KEYWORDS),
        len(CPC_EXCEL_FILES),
        ", ".join(f.name for f in CPC_EXCEL_FILES),
        KEYWORDS_SHEET_NAME,
        KEYWORDS_COLUMN,
        KEYWORDS_START_ROW,
    )

    # Precompiled once, not per row: each keyword becomes a
    # case-insensitive, word-boundary-bound regex so short
    # keywords (e.g. "PET", "EVA") don't false-positive on
    # substrings inside unrelated words (e.g. "carpet",
    # "competing"). Phrases like "polymer film" still match
    # correctly since \b anchors on the whole phrase's edges.
    KEYWORD_PATTERNS = [
        re.compile(
            r"\b" + re.escape(keyword) + r"\b",
            re.IGNORECASE,
        )
        for keyword in KEYWORDS
    ]

    # All keywords as ONE alternation, so matches_keywords()
    # scans each text field once instead of once per keyword.
    # Longest-first so a phrase is preferred over a keyword it
    # contains (only matters for readability of the pattern --
    # the filter just needs any hit).
    KEYWORD_COMBINED_PATTERN = (
        re.compile(
            r"\b(?:"
            + "|".join(
                re.escape(keyword)
                for keyword in sorted(set(KEYWORDS), key=len, reverse=True)
            )
            + r")\b",
            re.IGNORECASE,
        )
        if KEYWORDS
        else None
    )

else:

    KEYWORDS = []
    KEYWORD_PATTERNS = []
    KEYWORD_COMBINED_PATTERN = None


# ============================================================
# USPTO API
# ============================================================

API_BASE = (
    "https://api.uspto.gov/api/v1"
)

PRODUCTS_ENDPOINT = (
    f"{API_BASE}/datasets/products"
)


# ============================================================
# API KEY
# ============================================================

def get_api_key():

    """
    Gets USPTO API key.

    Priority:
        1. API key entered above
        2. USPTO_API_KEY environment variable
    """

    if USPTO_API_KEY:

        return USPTO_API_KEY

    environment_key = os.getenv(
        "USPTO_API_KEY"
    )

    if environment_key:

        return environment_key

    raise RuntimeError(

        "\nUSPTO API KEY NOT FOUND.\n\n"

        "YOU NEED TO ENTER YOUR API KEY.\n\n"

        "Option 1:\n"
        "Set USPTO_API_KEY near the top of this program.\n\n"

        "Option 2 - Windows PowerShell:\n"
        '$env:USPTO_API_KEY="YOUR_KEY"\n'

    )


API_KEY = get_api_key()


# ============================================================
# HTTP SESSION
# ============================================================

SESSION = requests.Session()

# Connection pool sized to the largest number of threads that
# share this session at once. requests' default pool holds 10
# connections per host, so with more than 10 PDF workers the
# extras were opened, used, then thrown away ("Connection pool
# is full, discarding connection") -- harmless, but it meant
# new TLS handshakes instead of reusing connections.
_pool_size = max(PDF_DOWNLOAD_WORKERS, BULK_DOWNLOAD_WORKERS, 10) + 4

_adapter = requests.adapters.HTTPAdapter(
    pool_connections=_pool_size,
    pool_maxsize=_pool_size,
)

SESSION.mount("https://", _adapter)

SESSION.mount("http://", _adapter)

SESSION.headers.update({

    "X-API-KEY": API_KEY,

    "User-Agent":
        "USPTO-Patent-Research-Pipeline/1.0"

})


# ============================================================
# API REQUEST
# ============================================================

def api_get(
    url,
    params=None,
    stream=False
):

    """
    Performs an authenticated USPTO API request.

    Permanent failures (404/403/etc.) raise immediately
    instead of burning the full retry budget. Transient
    failures (429/503/etc.) use an escalating backoff.
    """

    last_error = None

    for attempt in range(
        1,
        MAX_RETRIES + 1
    ):

        try:

            response = SESSION.get(

                url,

                params=params,

                stream=stream,

                timeout=REQUEST_TIMEOUT

            )

            if response.status_code == 200:

                return response


            # ------------------------------------------------
            # PERMANENT FAILURE - do not retry
            # ------------------------------------------------

            if response.status_code in PERMANENT_FAILURE_STATUS_CODES:

                raise RuntimeError(

                    f"Permanent failure HTTP "
                    f"{response.status_code} for {url}"

                )


            logging.warning(

                "USPTO API returned HTTP %s",
                response.status_code

            )


            # ------------------------------------------------
            # TRANSIENT FAILURE - escalating backoff
            # ------------------------------------------------

            if response.status_code in TRANSIENT_FAILURE_STATUS_CODES:

                wait = (
                    RETRY_DELAY_SECONDS *
                    attempt
                )

                logging.warning(

                    "Transient failure (HTTP %s). "
                    "Waiting %s seconds.",
                    response.status_code,
                    wait

                )

                time.sleep(wait)

                continue


            # Any other unexpected status: fixed-delay retry.

            last_error = (
                f"Unexpected HTTP {response.status_code}"
            )

        except RuntimeError:

            # Permanent failure - propagate immediately.

            raise

        except requests.RequestException as error:

            last_error = str(error)

            logging.warning(

                "Request error: %s",
                error

            )


        if attempt < MAX_RETRIES:

            time.sleep(
                RETRY_DELAY_SECONDS
            )


    raise RuntimeError(

        f"USPTO request failed after "
        f"{MAX_RETRIES} attempts:\n{url}\n"
        f"Last error: {last_error}"

    )


# ============================================================
# FIND USPTO PRODUCTS
# ============================================================

def search_products(search_text):

    """
    Search USPTO ODP products.
    """

    url = (
        f"{PRODUCTS_ENDPOINT}/search"
    )

    params = {

        "q": search_text,

        "limit": 100,

        "offset": 0,

    }

    response = api_get(
        url,
        params=params
    )

    return response.json()


# ============================================================
# DISPLAY PRODUCTS
# ============================================================

def display_products(data):

    print()

    print("=" * 80)

    print("USPTO PRODUCTS")

    print("=" * 80)

    products = data.get(
        "bulkDataProductBag",
        []
    )

    for product in products:

        print()

        print(
            "Product ID:",
            product.get(
                "productIdentifier"
            )
        )

        print(
            "Title:",
            product.get(
                "productTitleText"
            )
        )

    print()

    return products


# ============================================================
# FIND PRODUCT
# ============================================================

def find_product(
    search_text,
    expected_title
):

    data = search_products(
        search_text
    )

    products = data.get(
        "bulkDataProductBag",
        []
    )

    for product in products:

        title = product.get(
            "productTitleText",
            ""
        )

        if expected_title.lower() in title.lower():

            return product


    display_products(data)


    raise RuntimeError(

        f"\nCould not automatically identify:\n"
        f"{expected_title}\n\n"

        "Review the products printed above."

    )


# ============================================================
# GET PRODUCT FILE LIST
# ============================================================

def get_product_files(
    product_identifier
):

    url = (
        f"{PRODUCTS_ENDPOINT}/"
        f"{quote(product_identifier, safe='')}"
    )

    params = {

        "includeFiles": "true",

        "limit": 10000,

        "offset": 0,

    }

    response = api_get(

        url,

        params=params

    )

    return response.json()


# ============================================================
# EXTRACT FILE LIST
# ============================================================

def extract_file_list(
    product_data
):
    """
    Extract the actual file records from a USPTO bulk-data
    product response.

    Current USPTO response structure:

        bulkDataProductBag
            -> productFileBag
                -> fileDataBag
                    -> [file records]

    Also supports several alternate response structures.
    """

    # --------------------------------------------------------
    # Locate the product object
    # --------------------------------------------------------

    product = product_data.get(
        "bulkDataProduct",
        {}
    )

    product_bag = product_data.get(
        "bulkDataProductBag"
    )

    # Direct product lookup normally returns the product
    # inside bulkDataProductBag.
    if (
        not isinstance(product, dict)
        and isinstance(product_bag, list)
        and product_bag
    ):
        product = product_bag[0]

    # --------------------------------------------------------
    # Current USPTO format:
    #
    # productFileBag:
    #     {
    #         "count": 1312,
    #         "fileDataBag": [...]
    #     }
    # --------------------------------------------------------

    product_file_bag = None

    if isinstance(product, dict):
        product_file_bag = product.get(
            "productFileBag"
        )

    if isinstance(product_file_bag, dict):

        file_data_bag = product_file_bag.get(
            "fileDataBag"
        )

        if isinstance(file_data_bag, list):
            return file_data_bag

    # --------------------------------------------------------
    # Alternate format: productFiles directly
    # --------------------------------------------------------

    if isinstance(product, dict):

        product_files = product.get(
            "productFiles"
        )

        if isinstance(product_files, list):
            return product_files

    # --------------------------------------------------------
    # Alternate format: productFileBag itself is a list
    # --------------------------------------------------------

    if isinstance(product_file_bag, list):
        return product_file_bag

    # --------------------------------------------------------
    # Fallback: inspect every product in product bag
    # --------------------------------------------------------

    if isinstance(product_bag, list):

        for entry in product_bag:

            if not isinstance(entry, dict):
                continue

            file_bag = entry.get(
                "productFileBag"
            )

            if isinstance(file_bag, dict):

                file_data_bag = file_bag.get(
                    "fileDataBag"
                )

                if isinstance(file_data_bag, list):
                    return file_data_bag

            if isinstance(file_bag, list):
                return file_bag

            product_files = entry.get(
                "productFiles"
            )

            if isinstance(product_files, list):
                return product_files

    # --------------------------------------------------------
    # Nothing recognized
    # --------------------------------------------------------

    return []


def extract_product_totals(
    product_data
):

    """
    Pulls the PRODUCT-LEVEL total size/file count/date range
    out of a product metadata response -- as opposed to
    extract_file_list(), which pulls the individual file
    records. USPTO's own productTotalFileSize field was
    verified against a real product response to match the sum
    of every individually-listed file within 0.0003% (166.17GB
    either way), so this is trustworthy without needing to
    download anything or sum files by hand.

    Handles the same bulkDataProductBag-wrapped envelope
    extract_file_list() already knows how to navigate, since a
    direct product lookup wraps its single result the same way
    the /search endpoint wraps multiple.

    Returns None if the expected fields aren't found anywhere.
    """

    candidates = []

    if isinstance(product_data, dict):

        candidates.append(product_data)

    product_bag = (

        product_data.get("bulkDataProductBag")

        if isinstance(product_data, dict)

        else None

    )

    if isinstance(product_bag, list):

        candidates.extend(

            entry
            for entry in product_bag
            if isinstance(entry, dict)

        )

    bulk_product = (

        product_data.get("bulkDataProduct")

        if isinstance(product_data, dict)

        else None

    )

    if isinstance(bulk_product, dict):

        candidates.append(bulk_product)


    for candidate in candidates:

        if "productTotalFileSize" in candidate:

            return {

                "total_size_bytes": candidate.get(
                    "productTotalFileSize"
                ),

                "total_files": candidate.get(
                    "productFileTotalQuantity"
                ),

                "from_date": candidate.get(
                    "productFromDate"
                ),

                "to_date": candidate.get(
                    "productToDate"
                ),

            }


    return None


def _format_bytes_human(
    num_bytes
):

    if num_bytes is None:

        return "unknown"

    value = float(num_bytes)

    for unit in ["B", "KB", "MB", "GB", "TB"]:

        if value < 1024:

            return f"{value:.2f} {unit}"

        value /= 1024

    return f"{value:.2f} PB"


# ============================================================
# CHECK TOTAL CORPUS SIZE
# ============================================================

def check_total_corpus_size():

    """
    Queries USPTO for each of the 4 known products' METADATA
    ONLY (no bulk files downloaded) and prints their total
    size and file count, so storage can be planned BEFORE
    committing to a full --download.
    """

    products = [

        ("GRANTS_APS", "PTGRAPS"),
        ("GRANTS_SGML", "PTGRSGM"),
        ("GRANTS_XML", "PTGRXML"),
        ("APPLICATIONS_XML", "APPXML"),

    ]

    print()

    print("=" * 80)

    print(
        "USPTO CORPUS SIZE CHECK "
        "(metadata only -- nothing downloaded)"
    )

    print("=" * 80)


    grand_total_bytes = 0

    grand_total_files = 0

    any_totals_missing = False

    for product_type, product_identifier in products:

        print()

        print(f"{product_type} ({product_identifier}):")

        try:

            product_data = get_product_files(
                product_identifier
            )

        except Exception as error:

            print(f"  FAILED to fetch: {error}")

            any_totals_missing = True

            continue


        totals = extract_product_totals(
            product_data
        )

        if totals is None or totals.get(

            "total_size_bytes"

        ) is None:

            print(

                "  Could not find productTotalFileSize in "
                "the response -- USPTO's response shape may "
                "have changed. Raw top-level keys: "
                f"{list(product_data.keys())}"

            )

            any_totals_missing = True

            continue


        size_bytes = totals["total_size_bytes"]

        file_count = totals.get("total_files")

        from_date = totals.get("from_date")

        to_date = totals.get("to_date")


        print(

            f"  Date range: {from_date} to {to_date}"

        )

        print(

            f"  Files: {file_count:,}"

            if file_count is not None

            else "  Files: unknown"

        )

        print(

            f"  Total size: "
            f"{_format_bytes_human(size_bytes)} "
            f"({size_bytes:,} bytes)"

        )


        grand_total_bytes += size_bytes

        if file_count is not None:

            grand_total_files += file_count


    print()

    print("=" * 80)

    print(

        f"GRAND TOTAL: {_format_bytes_human(grand_total_bytes)} "
        f"across {grand_total_files:,} files "
        f"(compressed, as downloaded -- extracted size will "
        f"be larger)"

    )

    if any_totals_missing:

        print(

            "\nNOTE: one or more products' totals could not "
            "be determined -- the grand total above is a "
            "MINIMUM, not the true full total."

        )

    print("=" * 80)


# ============================================================
# GET FILE NAME
# ============================================================

def get_file_name(
    file_info
):

    return (

        file_info.get("fileName")

        or

        file_info.get("filename")

        or

        file_info.get("file_name")

    )


# ============================================================
# SAMPLE FILE SELECTION
# ============================================================

def select_sample_files(

    files,

    sample_size,

):

    """
    Picks a handful of files spread evenly across the
    product's date range, instead of just the first or most
    recent N. USPTO's format has changed multiple times
    across each product's history (APS/Green Book -> SGML ->
    several XML generations), so an even spread is much more
    likely to actually surface one file from each era than
    grabbing whatever happens to sort first.
    """

    if not files or sample_size <= 0:

        return []


    sorted_files = sorted(

        files,

        key=lambda f: (

            f.get("fileDataFromDate")
            or f.get("fileDate")
            or "",
            get_file_name(f) or "",

        ),

    )


    if len(sorted_files) <= sample_size:

        return sorted_files


    indices = set()

    for i in range(sample_size):

        if sample_size > 1:

            position = round(

                i * (len(sorted_files) - 1)
                / (sample_size - 1)

            )

        else:

            position = 0

        indices.add(position)


    return [

        sorted_files[i]

        for i in sorted(indices)

    ]


# ============================================================
# DOWNLOAD MANIFEST
# ============================================================

# ------------------------------------------------------------
# GENERIC MANIFEST HELPERS
# ------------------------------------------------------------
# Both the bulk-download manifest and the PDF manifest are
# CSV files mapping some key -> a status/metadata row, used so
# --download and --pdf can be restarted without redoing work
# that already succeeded. The read/write logic is identical
# between the two; only the field list and how the key is
# built from a row differ, so both are implemented on top of
# these two shared helpers instead of duplicating the file I/O.

def _load_csv_manifest(

    path,

    build_key,

):

    records = {}


    if not path.exists():

        return records


    with open(

        path,

        "r",

        newline="",

        encoding="utf-8"

    ) as file:

        reader = csv.DictReader(file)

        for row in reader:

            records[build_key(row)] = row


    return records


def _save_csv_manifest(

    path,

    records,

    fieldnames,

):

    with open(

        path,

        "w",

        newline="",

        encoding="utf-8"

    ) as file:

        writer = csv.DictWriter(

            file,

            fieldnames=fieldnames

        )

        writer.writeheader()

        for row in records.values():

            writer.writerow(row)


DOWNLOAD_MANIFEST_FIELDS = [

    "product_type",
    "product_identifier",
    "product_title",
    "file_name",
    "file_size",
    "file_date",
    "local_path",
    "status",
    "download_time",
    "error",

]


def load_download_manifest():

    return _load_csv_manifest(

        DOWNLOAD_MANIFEST,

        build_key=lambda row: (

            row["product_identifier"],
            row["file_name"],

        ),

    )


# ============================================================
# SAVE DOWNLOAD MANIFEST
# ============================================================

def save_download_manifest(
    records
):

    _save_csv_manifest(

        DOWNLOAD_MANIFEST,
        records,
        DOWNLOAD_MANIFEST_FIELDS,

    )


# ============================================================
# DOWNLOAD ONE BULK FILE
# ============================================================

def download_bulk_file(

    product_type,

    product_identifier,

    product_title,

    file_info,

    destination,

):

    """
    Downloads a single bulk file.

    Returns (key, record_dict, success_bool). This function
    does NOT mutate the shared manifest directly -- it may be
    called from multiple threads at once, so the caller is
    responsible for writing the result into the manifest
    (under a lock).
    """

    file_name = get_file_name(
        file_info
    )


    if not file_name:

        logging.warning(
            "File has no filename."
        )

        return None, None, False


    key = (

        product_identifier,

        file_name

    )


    expected_size = file_info.get(
        "fileSize"
    )


    destination.parent.mkdir(

        parents=True,

        exist_ok=True

    )


    def build_record(status, download_time="", error=""):

        return {

            "product_type": product_type,
            "product_identifier": product_identifier,
            "product_title": product_title,
            "file_name": file_name,
            "file_size": expected_size,
            "file_date": file_info.get("fileDate"),
            "local_path": str(destination),
            "status": status,
            "download_time": download_time,
            "error": error,

        }


    # --------------------------------------------------------
    # ALREADY COMPLETE
    # --------------------------------------------------------

    if destination.exists():

        actual_size = (
            destination.stat().st_size
        )

        size_ok = (

            expected_size is None

            or

            actual_size == int(expected_size)

        )

        # Bonus integrity check: a .zip that isn't actually
        # a valid zip (truncated/corrupted download) should
        # be re-downloaded rather than silently accepted.
        integrity_ok = True

        if size_ok and destination.suffix.lower() == ".zip":

            integrity_ok = zipfile.is_zipfile(destination)

        if size_ok and integrity_ok:

            logging.info(
                "SKIP %s",
                file_name
            )

            return key, build_record("COMPLETE"), True

        logging.warning(

            "Existing file failed validation, "
            "re-downloading: %s",
            file_name

        )

        destination.unlink(missing_ok=True)


    # --------------------------------------------------------
    # TEMPORARY FILE
    # --------------------------------------------------------

    partial = Path(
        str(destination) + ".part"
    )


    # Prefer the exact download URI supplied by USPTO.
    # This is important because the URI can include the
    # year/path structure used by the current ODP API.
    url = file_info.get(
        "fileDownloadURI"
    )

    # Fallback for older/different response formats.
    if not url:

        url = (

            f"{PRODUCTS_ENDPOINT}/files/"

            f"{quote(product_identifier, safe='')}/"

            f"{quote(file_name, safe='')}"

        )


    last_error = None

    for attempt in range(

        1,

        MAX_RETRIES + 1

    ):

        try:

            existing_size = (

                partial.stat().st_size

                if partial.exists()

                else 0

            )


            headers = {}


            if existing_size > 0:

                headers["Range"] = (

                    f"bytes={existing_size}-"

                )


            response = SESSION.get(

                url,

                headers=headers,

                stream=True,

                timeout=REQUEST_TIMEOUT

            )


            # ----------------------------------------------------
            # PERMANENT FAILURE - stop immediately, don't retry
            # ----------------------------------------------------

            if response.status_code in PERMANENT_FAILURE_STATUS_CODES:

                error_message = (
                    f"Permanent failure HTTP "
                    f"{response.status_code}"
                )

                logging.error(
                    "PERMANENT FAILURE %s: %s",
                    file_name,
                    error_message,
                )

                return (

                    key,
                    build_record("FAILED", error=error_message),
                    False,

                )


            if response.status_code == 206:

                mode = "ab"

            elif response.status_code == 200:

                mode = "wb"

                existing_size = 0

            else:

                raise RuntimeError(

                    f"HTTP {response.status_code}"

                )


            content_length = (

                response.headers.get(
                    "Content-Length"
                )
            )


            total = None


            if content_length:

                total = (
                    int(content_length)
                    +
                    existing_size
                )


            with open(

                partial,

                mode

            ) as output:

                with tqdm(

                    total=total,

                    initial=existing_size,

                    unit="B",

                    unit_scale=True,

                    unit_divisor=1024,

                    desc=file_name[:50],

                    leave=False,

                ) as progress:

                    for chunk in response.iter_content(

                        chunk_size=CHUNK_SIZE

                    ):

                        if chunk:

                            output.write(
                                chunk
                            )

                            progress.update(
                                len(chunk)
                            )


            # Validate final size against what USPTO told us
            # to expect, when that's available.

            final_size = partial.stat().st_size

            if (

                expected_size is not None

                and

                final_size != int(expected_size)

            ):

                raise RuntimeError(

                    f"Size mismatch: expected {expected_size}, "
                    f"got {final_size}"

                )


            # Bonus integrity check before accepting a zip.

            if (

                partial.suffix.lower() == ".zip"

                and not zipfile.is_zipfile(partial)

            ):

                raise RuntimeError(
                    "Downloaded file is not a valid ZIP archive"
                )


            partial.replace(
                destination
            )


            logging.info(
                "Downloaded %s",
                file_name
            )


            return (

                key,

                build_record(
                    "COMPLETE",
                    download_time=datetime.now().isoformat(),
                ),

                True,

            )


        except Exception as error:

            last_error = str(error)

            logging.warning(

                "Failed %s (attempt %d/%d): %s",
                file_name,
                attempt,
                MAX_RETRIES,
                error

            )


            if attempt < MAX_RETRIES:

                time.sleep(
                    RETRY_DELAY_SECONDS * attempt
                )


    return (

        key,

        build_record("FAILED", error=str(last_error)),

        False,

    )


# ============================================================
# DOWNLOAD ENTIRE PRODUCT
# ============================================================

def download_product(

    product_type,

    destination_directory,

    manifest,

    manifest_lock,

    search_text=None,

    expected_title=None,

    product_identifier=None,

    sample=False,

):

    """
    If product_identifier is given, it's used directly and
    search_text/expected_title are ignored entirely. This
    matters because USPTO exposes multiple bulk-data products
    whose titles all contain overlapping text -- e.g.
    "Patent Grant Full-Text Data (No Images) - APS", "- SGML",
    and "- XML" ALL match a fuzzy search for "Patent Grant
    Full-Text". Matching on title substring alone will grab
    whichever of those happens to be listed first, which is
    not necessarily the one you want. Pass exact identifiers
    (e.g. "PTGRXML") whenever you know them.
    """

    print()

    print("=" * 80)

    print(
        f"PRODUCT: {product_type}"
    )

    print("=" * 80)


    if product_identifier:

        product_title = product_type

    else:

        product = find_product(

            search_text,

            expected_title

        )

        product_identifier = (

            product.get(
                "productIdentifier"
            )
        )

        product_title = (

            product.get(
                "productTitleText",
                ""
            )
        )


    print()

    print(
        "Product identifier:",
        product_identifier
    )

    print(
        "Product:",
        product_title
    )


    product_data = get_product_files(

        product_identifier

    )


    # Save response so we don't lose it.

    response_file = (

        LOG_DIR /

        f"{product_type}_product.json"

    )


    with open(

        response_file,

        "w",

        encoding="utf-8"

    ) as file:

        json.dump(

            product_data,

            file,

            indent=2

        )


    files = extract_file_list(

        product_data

    )


    print()

    print(
        f"FILES FOUND: {len(files):,}"
    )

    # Report the count supplied by USPTO as an additional
    # sanity check. The API's product metadata and fileDataBag
    # count can occasionally differ because of revisions or
    # withdrawn/replaced files.
    bulk_product = None

    product_bag = product_data.get(
        "bulkDataProductBag"
    )

    if isinstance(product_bag, list) and product_bag:
        if isinstance(product_bag[0], dict):
            bulk_product = product_bag[0]

    if isinstance(bulk_product, dict):

        product_file_bag = bulk_product.get(
            "productFileBag"
        )

        if isinstance(product_file_bag, dict):

            api_file_count = product_file_bag.get(
                "count"
            )

            if api_file_count is not None:

                print(
                    f"USPTO fileDataBag count: "
                    f"{api_file_count:,}"
                )


    if not files:

        print()

        print(

            f"WARNING: USPTO returned 0 files for "
            f"'{product_identifier}'. This usually means "
            f"either the identifier is wrong, or the "
            f"response is shaped differently than expected."

        )

        print(

            "Top-level response keys:",
            list(product_data.keys()),

        )

        bulk_product = product_data.get(
            "bulkDataProduct"
        )

        if isinstance(bulk_product, dict):

            print(

                "  bulkDataProduct keys:",
                list(bulk_product.keys()),

            )

        product_bag = product_data.get(
            "bulkDataProductBag"
        )

        if isinstance(product_bag, list):

            print(

                f"  bulkDataProductBag: "
                f"{len(product_bag)} entrie(s)",

            )

            for index, entry in enumerate(product_bag):

                if isinstance(entry, dict):

                    print(

                        f"    entry[{index}] keys:",
                        list(entry.keys()),

                    )

        print(

            f"Full raw response saved to: {response_file}"

        )


    if sample:

        original_count = len(files)

        files = select_sample_files(

            files,
            SAMPLE_FILES_PER_PRODUCT,

        )

        print(

            f"SAMPLE MODE: selected {len(files)} of "
            f"{original_count:,} files, spread across the "
            f"available date range (to try to catch "
            f"different USPTO format eras)."

        )

        for file_info in files:

            file_date = (
                file_info.get("fileDataFromDate")
                or file_info.get("fileDate")
                or ""
            )

            print(

                f"  - {get_file_name(file_info)} "
                f"(fileDate: {file_date})"

            )


    print(

        f"Downloading with {BULK_DOWNLOAD_WORKERS} "
        f"parallel workers..."

    )


    completed_since_save = 0

    failures = 0

    with ThreadPoolExecutor(

        max_workers=BULK_DOWNLOAD_WORKERS

    ) as executor:

        future_to_file = {}

        for file_info in files:

            file_name = get_file_name(
                file_info
            )

            destination = (
                destination_directory /
                file_name
            )

            future = executor.submit(

                download_bulk_file,

                product_type,

                product_identifier,

                product_title,

                file_info,

                destination,

            )

            future_to_file[future] = file_name


        for number, future in enumerate(

            as_completed(future_to_file),

            start=1

        ):

            file_name = future_to_file[future]

            try:

                key, record, success = future.result()

            except Exception as error:

                logging.error(

                    "Unhandled error downloading %s: %s",
                    file_name,
                    error,

                )

                key, record, success = None, None, False


            if not success:

                failures += 1


            print(

                f"[{number:,}/{len(files):,}] "
                f"{'OK' if success else 'FAILED'} "
                f"{file_name}"

            )


            if key is not None:

                with manifest_lock:

                    manifest[key] = record

                    completed_since_save += 1

                    if completed_since_save >= MANIFEST_SAVE_INTERVAL:

                        save_download_manifest(
                            manifest
                        )

                        completed_since_save = 0


    # Final save to make sure everything from this product
    # is flushed to disk, even if we didn't hit the interval.

    with manifest_lock:

        save_download_manifest(
            manifest
        )


    print()

    print(

        f"{product_type} complete. "
        f"{len(files) - failures:,} succeeded, "
        f"{failures:,} failed."

    )

    if failures:

        print(

            f"See {DOWNLOAD_MANIFEST} for details "
            f"on failed files (re-running --download "
            f"will retry them)."

        )


# ============================================================
# EXTRACT ZIP FILES
# ============================================================

def _extract_one_archive(zip_path):

    """
    Extracts a single zip archive. Returns (zip_path, success,
    error_message_or_None). Designed to be run from a thread
    pool, so it returns rather than mutating shared state.
    """

    relative = (
        zip_path.relative_to(
            RAW_DIR
        )
    )


    destination = (

        EXTRACTED_DIR /

        relative.parent /

        zip_path.stem

    )


    marker = (

        destination /

        ".EXTRACTION_COMPLETE"

    )


    if marker.exists():

        return zip_path, True, None


    destination.mkdir(

        parents=True,

        exist_ok=True

    )


    try:

        with zipfile.ZipFile(
            zip_path,
            "r"
        ) as archive:

            archive.extractall(
                destination
            )


        marker.touch()

        return zip_path, True, None


    except Exception as error:

        return zip_path, False, str(error)


def extract_all_archives():

    print()

    print("=" * 80)

    print("EXTRACTING BULK FILES")

    print("=" * 80)


    zip_files = list(

        RAW_DIR.rglob(
            "*.zip"
        )

    )


    print(

        f"ZIP files found: "
        f"{len(zip_files):,}"

    )

    print(

        f"Extracting with {EXTRACTION_WORKERS} "
        f"parallel workers..."

    )


    failures = 0

    with ThreadPoolExecutor(

        max_workers=EXTRACTION_WORKERS

    ) as executor:

        futures = [

            executor.submit(
                _extract_one_archive,
                zip_path,
            )

            for zip_path in zip_files

        ]


        for future in tqdm(

            as_completed(futures),

            total=len(zip_files),

            desc="Extracting"

        ):

            zip_path, success, error = future.result()

            if not success:

                failures += 1

                logging.error(

                    "Extraction failed: "
                    "%s | %s",

                    zip_path,

                    error

                )


    print()

    print(

        f"Extraction complete. "
        f"{len(zip_files) - failures:,} succeeded, "
        f"{failures:,} failed."

    )


# ============================================================
# DATABASE
# ============================================================
#
# SCHEMA (v15) -- two tables per product database:
#
#   patents       small, fixed-size bibliographic metadata
#                 (numbers, dates, CPC/IPC, names, title).
#                 ~300-500 bytes per row.
#
#   patent_text   abstract / claims / description, keyed by
#                 the same id as patents.id. Claims and
#                 description are zlib-compressed BLOBs (see
#                 COMPRESS_TEXT).
#
# WHY SPLIT THE TEXT OUT: in v14 cpc/ipc sat AFTER claims and
# description in one wide row. SQLite stores big values in
# "overflow page" chains, and to read a column it has to walk
# every overflow page of every column before it -- so the
# --filter stage's CPC check effectively read the ENTIRE
# corpus's full text off the external drive just to look at a
# 50-character classification string. With the metadata in its
# own narrow table, the CPC pass reads a few GB instead of
# hundreds, and full text is only read for the small fraction
# of patents that pass the CPC gate.
# ------------------------------------------------------------

def get_db_connection(db_path):

    """
    Opens a SQLite connection with pragmas tuned for this
    workload. Every process (main + each worker) opens its own.

    - WAL + synchronous=NORMAL: no fsync per commit; writers
      don't block the (rare) reader.
    - busy_timeout: several worker processes write to the same
      product file; a writer that finds the file locked waits
      for the other's (short) batch to finish instead of
      failing.
    - journal_size_limit: after each checkpoint the -wal file is
      truncated back down to this size, so it can't balloon.
      (v14 held one transaction open across 200 files, which
      on modern XML meant a WAL file tens of GB in size.)
    - cache_size: ~64 MB per connection. Deliberately modest --
      with 4+ worker processes each holding one, a large cache
      multiplies quickly.
    """

    connection = sqlite3.connect(
        db_path,
        timeout=DATABASE_BUSY_TIMEOUT_MS / 1000,
    )

    cursor = connection.cursor()

    cursor.execute("PRAGMA journal_mode = WAL")

    cursor.execute("PRAGMA synchronous = NORMAL")

    cursor.execute(f"PRAGMA busy_timeout = {DATABASE_BUSY_TIMEOUT_MS}")

    cursor.execute("PRAGMA journal_size_limit = 268435456")  # 256 MB

    cursor.execute("PRAGMA temp_store = MEMORY")

    cursor.execute("PRAGMA cache_size = -64000")  # ~64 MB

    return connection


def create_database(db_path):

    """
    Creates the two tables in db_path if they don't exist yet.

    Secondary indexes (idx_patent_number, idx_publication_date)
    are deliberately NOT created here -- see
    create_database_indexes() for why they're built after the
    bulk insert instead.
    """

    connection = get_db_connection(db_path)

    cursor = connection.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS patents (
            id INTEGER PRIMARY KEY,
            patent_number TEXT,
            publication_number TEXT,
            document_type TEXT,
            filing_date TEXT,
            publication_date TEXT,
            grant_date TEXT,
            priority_date TEXT,
            cpc TEXT,
            ipc TEXT,
            title TEXT,
            inventors TEXT,
            applicants TEXT,
            assignees TEXT,
            source_file TEXT,
            source_path TEXT,
            UNIQUE(publication_number)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS patent_text (
            id INTEGER PRIMARY KEY,
            abstract TEXT,
            claims BLOB,
            description BLOB
        )
    """)

    connection.commit()

    connection.close()


def create_database_indexes(db_path):

    """
    Builds idx_patent_number and idx_publication_date AFTER the
    bulk insert has finished (one bulk pass instead of one
    incremental B-tree update per inserted row). IF NOT EXISTS
    makes this a no-op once built; SQLite keeps an existing
    index up to date on later inserts automatically.

    v14's idx_cpc is intentionally gone: the cpc column holds a
    delimited LIST of codes, and --filter matches any code in
    that list by prefix, which a B-tree index on the whole
    string can't help with. It only cost build time and disk.
    Both remaining indexes are on the narrow patents table, so
    they build in minutes rather than hours.
    """

    connection = get_db_connection(db_path)

    connection.execute(
        "CREATE INDEX IF NOT EXISTS idx_patent_number "
        "ON patents(patent_number)"
    )

    connection.execute(
        "CREATE INDEX IF NOT EXISTS idx_publication_date "
        "ON patents(publication_date)"
    )

    connection.commit()

    connection.close()


def _compress_text(text):

    """zlib-compresses a text field for storage (see COMPRESS_TEXT)."""

    if not text:
        return None

    if not COMPRESS_TEXT:
        return text

    return zlib.compress(
        text.encode("utf-8"),
        TEXT_COMPRESSION_LEVEL,
    )


def _decompress_text(value):

    """Inverse of _compress_text(); also accepts plain text."""

    if value is None:
        return ""

    if isinstance(value, (bytes, bytearray, memoryview)):
        return zlib.decompress(bytes(value)).decode("utf-8")

    return value


PATENTS_INSERT_SQL = """
    INSERT OR IGNORE INTO patents (
        patent_number, publication_number, document_type,
        filing_date, publication_date, grant_date, priority_date,
        cpc, ipc, title, inventors, applicants, assignees,
        source_file, source_path
    )
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
"""

PATENT_TEXT_INSERT_SQL = """
    INSERT OR REPLACE INTO patent_text (
        id, abstract, claims, description
    )
    VALUES (?,?,?,?)
"""


def _insert_record_batch(connection, batch, source_path):

    """
    Writes one batch of parsed records in ONE short
    transaction. Returns how many were actually new (INSERT OR
    IGNORE silently skips publication_numbers already present,
    e.g. when a partially-processed file is re-run).

    The text row is only written when the metadata row was
    genuinely new, keyed by that row's id, so the two tables
    always stay in step.

    Called in worker processes. Compression happens before the
    transaction opens, so the write lock is held only for the
    actual inserts.
    """

    prepared = [
        (
            (
                record.get("patent_number"),
                record.get("publication_number"),
                record.get("document_type"),
                record.get("filing_date"),
                record.get("publication_date"),
                record.get("grant_date"),
                record.get("priority_date"),
                record.get("cpc"),
                record.get("ipc"),
                record.get("title"),
                record.get("inventors"),
                record.get("applicants"),
                record.get("assignees"),
                record.get("source_file"),
                source_path,
            ),
            record.get("abstract"),
            _compress_text(record.get("claims")),
            _compress_text(record.get("description")),
        )
        for record in batch
    ]

    inserted = 0

    cursor = connection.cursor()

    # BEGIN IMMEDIATE takes the write lock up front (waiting up
    # to busy_timeout for another worker's batch to finish)
    # rather than failing half-way through the batch.
    cursor.execute("BEGIN IMMEDIATE")

    try:

        for meta_row, abstract, claims, description in prepared:

            cursor.execute(PATENTS_INSERT_SQL, meta_row)

            if cursor.rowcount == 1:

                cursor.execute(
                    PATENT_TEXT_INSERT_SQL,
                    (cursor.lastrowid, abstract, claims, description),
                )

                inserted += 1

        connection.commit()

    except BaseException:

        connection.rollback()

        raise

    return inserted


def backup_database(db_path):

    """
    Copies one product's database to DATABASE/BACKUP/,
    overwriting that product's previous backup. Only runs when
    --backup is passed (v15: opt-in, was automatic in v14).

    Why opt-in now: in v14 every clean --database run copied
    the ENTIRE combined database -- by the XML stage that's
    hundreds of GB of sustained read+write on the same external
    drive, every day, for a copy that sits on the same physical
    disk (so it doesn't protect against the drive itself
    failing). With per-product files, the sensible moment to
    back up is once, when a product is fully finished:
        python ... --database --format sgml --backup
    Ideally point DATABASE_BACKUP_DIR at a DIFFERENT drive.

    Checkpoints WAL into the main file first, then copies to a
    temp file and renames it into place, so an interrupted
    backup never damages the previous one.
    """

    if not db_path.exists():
        return

    DATABASE_BACKUP_DIR.mkdir(parents=True, exist_ok=True)

    backup_path = DATABASE_BACKUP_DIR / db_path.name

    temp_path = DATABASE_BACKUP_DIR / (db_path.name + ".tmp")

    print(f"  Backing up {db_path.name} to {backup_path} ...")

    try:

        connection = get_db_connection(db_path)

        connection.execute("PRAGMA wal_checkpoint(TRUNCATE)")

        connection.close()

        shutil.copy2(db_path, temp_path)

        os.replace(temp_path, backup_path)

        print("  Backup complete.")

        logging.info("Database backed up to %s", backup_path)

    except (OSError, sqlite3.Error) as error:

        print(f"  WARNING: database backup failed: {error}")

        logging.warning("Database backup failed: %s", error)

        if temp_path.exists():

            try:
                temp_path.unlink()
            except OSError:
                pass


# ============================================================
# PARSING PLACEHOLDER
# ============================================================

def iter_patent_records(file_path):

    """
    Detects the file's format (APS / Green Book, SGML, XML) by
    extension and yields one record dict at a time from the
    matching parser. Each record carries: patent_number,
    publication_number, document_type, title, abstract, claims,
    description, inventors, applicants, assignees, the four
    dates, cpc, ipc, source_file.

    A generator, so callers can stream a multi-hundred-MB weekly
    file without ever holding all of its records at once.
    """

    suffix = file_path.suffix.lower()

    if suffix == ".xml":
        return iter_xml_records(file_path)

    if suffix in (".sgml", ".sgm"):
        return iter_sgml_records(file_path)

    # APS/Green Book files don't use a modern extension.
    return iter_legacy_records(file_path)


def parse_patent_file(file_path):

    """List-returning wrapper around iter_patent_records()."""

    return list(iter_patent_records(file_path))


# ============================================================
# XML PARSER
# ============================================================

def _split_glued_declarations(lines):

    """
    Yields the file's lines, but splits any line where an XML
    declaration starts part-way through, so every record
    boundary begins its own line.

    Confirmed against real 2023 application files (e.g.
    ipa230914.xml, 74 occurrences in that one file): after a
    sequence-listing document, USPTO sometimes writes the NEXT
    patent's declaration on the same line as the closing tag:

        </sequence-cwu><?xml version="1.0" encoding="UTF-8"?>

    The splitter only recognised declarations at the start of a
    line, so that next patent was glued onto the sequence
    listing, the chunk failed with "junk after document
    element" (col 15 = the end of </sequence-cwu>), the
    listing was recovered and discarded as non-patent, and the
    patent itself was lost.

    Only a real declaration ("<?xml" + whitespace) with actual
    content before it on the line triggers a split, so
    <?xml-stylesheet ...?> and leading indentation are
    unaffected. A literal "<?xml" can't occur inside XML text
    (it would be escaped as &lt;?xml), so this can't cut a
    record in the middle.
    """

    for line in lines:

        position = line.find("<?xml", 1)

        while position != -1:

            if line[:position].strip() and re.match(r"<\?xml\s", line[position:]):

                yield line[:position] + "\n"

                line = line[position:]

                position = line.find("<?xml", 1)

            else:

                position = line.find("<?xml", position + 1)

        yield line


def iter_xml_records(file_path):

    """
    Parses a USPTO bulk XML file in the "v4.x" DTD family
    (us-patent-grant / us-patent-application root elements,
    dtd-version like "v4.3 2012-12-04" or "v4.4 2013-05-16").

    A single downloaded bulk file is NOT one well-formed XML
    document -- it's many individual XML documents
    concatenated together, each with its own
    <?xml version="1.0"?> declaration. This splits the file on
    those boundaries, parses each record independently, and
    returns a list of record dicts.

    Built and verified against real sample records:
      - a GRANTS design patent (no CPC -- design patents use
        Locarno/US-National classification instead, so cpc
        correctly comes back empty for those)
      - an APPLICATIONS utility patent (full classifications-cpc
        and classifications-ipcr blocks present)

    NOTE: older USPTO XML generations (pre-~2005, before the
    v4.x DTD family) may use different root element names and
    structure entirely. If/when one of those is encountered,
    this logs a warning and skips the record rather than
    guessing -- that generation will need its own handling
    once we have a real sample of it.
    """

    source_file_name = file_path.name

    current_lines = []

    # Matches a genuine XML declaration ("<?xml version=..."),
    # which always has whitespace right after "xml", but NOT
    # other "<?xml"-prefixed processing instructions like
    # "<?xml-stylesheet ...?>" -- confirmed necessary against a
    # real record in pa010315.xml, where one document's
    # <?xml version="1.0"?> declaration is immediately followed
    # on its own line by an <?xml-stylesheet ...?> instruction.
    # A plain startswith("<?xml") check treats THAT line as a
    # second record boundary too, splitting one real document
    # into a bogus empty "record" (just the declaration, which
    # fails to parse) plus the rest of the real document as a
    # separate chunk. The real content still gets captured
    # either way (the second chunk still parses successfully,
    # confirmed by an unchanged total record count against this
    # exact file), so this wasn't losing data -- but it's a
    # needless, confusing "malformed record" warning in the log
    # for something that isn't actually malformed, so it's
    # worth not triggering in the first place.

    _xml_declaration_pattern = re.compile(r"^<\?xml\s")

    with open(

        file_path,

        "r",

        encoding="utf-8",

        errors="replace",

    ) as file:

        for line in _split_glued_declarations(file):

            if (

                _xml_declaration_pattern.match(line.lstrip())

                and current_lines

            ):

                record = _parse_single_xml_record(

                    "".join(current_lines),
                    source_file_name,

                )

                if record:

                    yield record

                current_lines = [line]

            else:

                current_lines.append(line)


    if current_lines:

        record = _parse_single_xml_record(

            "".join(current_lines),
            source_file_name,

        )

        if record:

            yield record


def parse_xml_file(file_path):

    """
    List-returning wrapper around iter_xml_records(), kept for ad-hoc
    testing/inspection. The --database stage uses the generator
    directly so a whole file's records are never held in memory
    at once.
    """

    return list(iter_xml_records(file_path))


def _parse_single_xml_record(

    xml_text,

    source_file_name,

):

    try:

        root = ET.fromstring(xml_text)

    except ET.ParseError as error:

        # A transitional-era PATDOC-as-XML file (see the PATDOC
        # branch below) can be well-formed enough for the
        # lenient SGML-style builder but still fail STRICT XML
        # parsing -- confirmed against a real sample (design
        # patent D0452599): it uses the HTML named entity
        # "&copy;", which isn't one of XML's five predefined
        # entities and isn't declared in this file's internal
        # DOCTYPE subset (which only declares its .TIF image
        # entities), so ET.fromstring() correctly rejects it
        # here before ever reaching the root-tag check below.
        #
        # Before giving up, check cheaply whether this looks
        # like PATDOC content and, if so, fall back to the same
        # lenient extraction the confirmed-PATDOC branch below
        # uses. That builder is built on html.parser with
        # convert_charrefs=True, which resolves the full
        # standard HTML named-entity table (including &copy;),
        # not just XML's five -- so it correctly parses text
        # that strict XML parsing can't.

        # NOTE: searches the FULL xml_text, not just a leading
        # slice. An earlier version capped this at the first
        # 4096 characters, which silently missed real records
        # whose DOCTYPE internal subset runs long -- confirmed
        # against two real records in this format's sample file
        # that each declare dozens of embedded drawing/chemical-
        # structure entities (.TIF/.CDX/.MOL), pushing the
        # actual <patent-application-publication> opening tag
        # past 4096 characters (one as far as ~40,700). This
        # only runs on the ParseError fallback path (rare, not
        # per-record), so searching the full text costs nothing
        # meaningful -- a truncated lookahead's only effect is
        # silently losing real records, which is far worse than
        # the search itself being slightly slower.

        if "<PATDOC" in xml_text:

            return _parse_single_sgml_record(
                xml_text,
                source_file_name,
            )

        if "<patent-application-publication" in xml_text:

            # Same reasoning as the PATDOC fallback just above:
            # this early-PAP format (confirmed against a real
            # sample, doc-number 20010000001) uses old-style
            # named entities ET.fromstring() rejects, so a
            # ParseError here doesn't necessarily mean the
            # record is actually malformed -- try the lenient
            # extraction before giving up.

            return _parse_single_pap_record(
                xml_text,
                source_file_name,
            )

        if "junk after document element" in str(error) and error.position:

            # ET successfully parsed ONE complete root element,
            # then found unexpected trailing content -- almost
            # always because 2+ real records got merged into a
            # single chunk with no boundary this file's splitting
            # logic recognizes (the same underlying class of
            # problem as the <?xml-stylesheet> case above, just
            # a different specific trigger -- confirmed occurring
            # in real weekly grant files: e.g. ipg250930.xml,
            # ipg231114.xml). Losing the WHOLE chunk over this
            # would silently drop every record in it, not just
            # the malformed tail -- so recover at least the
            # FIRST document: ET's error position tells us
            # exactly where its content ends, so slice there and
            # retry parsing just that portion.
            #
            # This is a bounded, single-level recovery (first
            # document only), not full recursive multi-record
            # splitting -- deliberately conservative, since going
            # further without a confirmed real sample of what a
            # 3+-record merged chunk actually looks like risks
            # guessing at boundary logic that could silently
            # mis-slice content rather than just skip it. What's
            # left after the first document is logged clearly so
            # it's visible (not silently dropped) and can inform
            # a fuller fix later if this is common enough to
            # matter for the corpus as a whole.

            line, column = error.position

            text_lines = xml_text.splitlines(keepends=True)

            offset = sum(
                len(text_line)
                for text_line in text_lines[:line - 1]
            ) + column

            first_part = xml_text[:offset]
            remainder = xml_text[offset:].lstrip()

            try:

                root = ET.fromstring(first_part)

                logging.warning(

                    "Recovered the FIRST document from a merged/"
                    "junk-after-document chunk in %s (error at "
                    "line %s, col %s) -- %s additional "
                    "character(s) after it were NOT parsed and "
                    "may contain more real record(s). This file "
                    "is worth a closer look if this recurs often.",

                    source_file_name, line, column, len(remainder),

                )

                # Falls through to the normal root.tag dispatch
                # below using this successfully-recovered `root`
                # -- NOT returning here, unlike the PATDOC/PAP
                # fallbacks above, since a recovered document
                # here could be any known root type (most likely
                # us-patent-grant for this file family).

            except ET.ParseError:

                # Even the trimmed first part doesn't parse --
                # genuinely malformed, not just a merged-chunk
                # boundary issue. Give up on this chunk the same
                # as before.

                logging.warning(

                    "Skipping malformed XML record in %s: %s "
                    "(also failed after attempting merged-chunk "
                    "recovery)",
                    source_file_name,
                    error,

                )

                return None

        else:

            logging.warning(

                "Skipping malformed XML record in %s: %s",
                source_file_name,
                error,

            )

            return None


    # Root elements that are KNOWN and EXPECTED to appear in
    # bulk XML files but aren't patent bibliographic records
    # at all -- e.g. "sequence-cwu" is a WIPO ST.25-era
    # Sequence Listing attachment (DNA/protein sequence data),
    # bundled as its own sub-document alongside the actual
    # patent application record for biotech/pharma filings.
    # These are correctly skipped, just without raising an
    # alarming-sounding "unrecognized format" warning -- the
    # real patent record for the same application should still
    # appear elsewhere in the same file as a normal
    # us-patent-application record.

    _KNOWN_NON_PATENT_ROOT_ELEMENTS = {

        "sequence-cwu",

    }


    if root.tag == "us-patent-grant":

        document_type = "grant"

        bibliographic = root.find(
            "us-bibliographic-data-grant"
        )

    elif root.tag == "us-patent-application":

        document_type = "application"

        bibliographic = root.find(
            "us-bibliographic-data-application"
        )

    elif root.tag == "PATDOC":

        # A brief 2001-2002 transitional format, confirmed
        # against a real sample (design patent D0452599, dated
        # 2002-01-01): the OLD ST.32 tag vocabulary (B100-B748,
        # SDOBI, SDOCL, SDODE -- the same tags this file's own
        # SGML parser already handles) wrapped in valid,
        # well-formed XML syntax (proper closing tags,
        # self-closing elements like <CITED-BY-EXAMINER/>)
        # instead of true SGML's implied tag closing.
        #
        # Rather than duplicating extraction logic for what is
        # structurally the same format, this delegates straight
        # to _parse_single_sgml_record() on the raw text. That
        # function's tree builder is built on html.parser, which
        # already handles BOTH classic SGML's implied closing
        # (via its "search upward for a matching ancestor"
        # recovery) AND this format's explicit self-closing XML
        # tags (via handle_startendtag) -- so it needs no
        # modification to correctly parse this variant too.
        #
        # NOTE: this only fires once ET.fromstring() above has
        # already succeeded (i.e. the file is well-formed enough
        # as XML to reach this point). A transitional-era file
        # that ISN'T well-formed XML would instead hit the
        # ET.ParseError branch above and be skipped as before --
        # not handled here, since that hasn't been confirmed
        # against a real sample. In practice this is unlikely to
        # lose data even then: USPTO appears to have shipped a
        # parallel .sgm file alongside this DTD-2.5 .xml file for
        # at least this date, and that .sgm file is already
        # parsed correctly and independently via the SGML
        # dispatch (by file extension), so the same patent's
        # data would still be captured through that path.

        return _parse_single_sgml_record(
            xml_text,
            source_file_name,
        )

    elif root.tag == "patent-application-publication":

        # Same early-PAP format as the ParseError fallback
        # above, reached here instead on the (less common, but
        # possible) case where the record happens to be strict-
        # XML-well-formed. Delegates to the same extraction
        # either way.

        return _parse_single_pap_record(
            xml_text,
            source_file_name,
        )

    elif root.tag in _KNOWN_NON_PATENT_ROOT_ELEMENTS:

        return None

    else:

        logging.warning(

            "Unrecognized XML root element '%s' in %s -- "
            "this may be an older USPTO XML generation "
            "that hasn't been inspected/supported yet. "
            "Skipping this record.",
            root.tag,
            source_file_name,

        )

        return None


    if bibliographic is None:

        return None


    publication_ref = bibliographic.find(
        "publication-reference/document-id"
    )

    application_ref = bibliographic.find(
        "application-reference/document-id"
    )


    pub_country = (

        _element_text(publication_ref, "country")

        if publication_ref is not None

        else None

    )

    pub_doc_number = (

        _element_text(publication_ref, "doc-number")

        if publication_ref is not None

        else None

    )

    pub_kind = (

        _element_text(publication_ref, "kind")

        if publication_ref is not None

        else None

    )

    publication_date = (

        _normalize_xml_date(
            _element_text(publication_ref, "date")
        )

        if publication_ref is not None

        else None

    )


    publication_number = None

    if pub_country and pub_doc_number:

        publication_number = (

            f"{pub_country}{pub_doc_number}"
            f"{pub_kind or ''}"

        )


    filing_date = (

        _normalize_xml_date(
            _element_text(application_ref, "date")
        )

        if application_ref is not None

        else None

    )


    if document_type == "grant":

        # A grant's own publication IS the patent -- the
        # publication number and patent number are the same
        # value here.

        patent_number = publication_number

        grant_date = publication_date

    else:

        # A pre-grant application publication isn't a patent
        # yet.

        patent_number = None

        grant_date = None


    priority_date = _normalize_xml_date(

        _element_text(

            bibliographic,

            "priority-claims/priority-claim/"
            "document-id/date",

        )

    )


    title = _element_text(

        bibliographic,
        "invention-title",

    )


    abstract = _element_full_text(
        root.find("abstract")
    )

    claims = _element_full_text(
        root.find("claims")
    )

    description = _element_full_text(
        root.find("description")
    )


    inventors = _extract_party_names(

        bibliographic,
        "inventors",
        "inventor",

    )

    applicants = (

        _extract_party_names(
            bibliographic,
            "us-applicants",
            "us-applicant",
        )

        or

        _extract_party_names(
            bibliographic,
            "applicants",
            "applicant",
        )

    )

    assignees = _extract_party_names(

        bibliographic,
        "assignees",
        "assignee",

    )


    cpc = _extract_classification_symbols(

        bibliographic,
        "classification-cpc",

    )

    ipc = _extract_classification_symbols(

        bibliographic,
        "classification-ipcr",

    )


    return {

        "patent_number": patent_number,
        "publication_number": publication_number,
        "document_type": document_type,
        "title": title,
        "abstract": abstract,
        "claims": claims,
        "description": description,
        "inventors": inventors,
        "applicants": applicants,
        "assignees": assignees,
        "filing_date": filing_date,
        "publication_date": publication_date,
        "grant_date": grant_date,
        "priority_date": priority_date,
        "cpc": cpc,
        "ipc": ipc,
        "source_file": source_file_name,

    }


# ------------------------------------------------------------
# XML PARSING HELPERS
# ------------------------------------------------------------

def _element_text(

    element,

    path,

):

    if element is None:

        return None

    found = element.find(path)

    if found is None or found.text is None:

        return None

    text = found.text.strip()

    return text if text else None


def _element_full_text(

    element

):

    """
    Joins ALL text content within an element, including
    nested inline markup (e.g. <b>, <i>, <figref>,
    <claim-ref> inside claims/description/abstract
    paragraphs), collapsed to single spaces. Used for fields
    that are stored as one plain-text blob for keyword
    searching, not for preserving exact formatting.
    """

    if element is None:

        return None

    text = "".join(
        element.itertext()
    )

    text = re.sub(
        r"\s+",
        " ",
        text,
    ).strip()

    return text if text else None


def _normalize_xml_date(

    raw

):

    """
    USPTO XML dates are YYYYMMDD. Normalize to YYYY-MM-DD
    for consistency/readability; leaves anything that
    doesn't match that shape untouched rather than guessing.
    """

    if not raw:

        return None

    raw = raw.strip()

    if len(raw) == 8 and raw.isdigit():

        return f"{raw[0:4]}-{raw[4:6]}-{raw[6:8]}"

    return raw


def _extract_party_names(

    bibliographic,

    container_tag,

    member_tag,

):

    """
    Extracts and joins names from a container like
    <inventors>/<inventor>, <us-applicants>/<us-applicant>,
    or <assignees>/<assignee>. Handles both individual names
    (last-name/first-name) and organizations (orgname).
    """

    if bibliographic is None:

        return None

    container = bibliographic.find(
        f".//{container_tag}"
    )

    if container is None:

        return None


    names = []

    for member in container.findall(member_tag):

        addressbook = member.find("addressbook")

        if addressbook is None:

            continue


        orgname = _element_text(
            addressbook,
            "orgname",
        )

        if orgname:

            names.append(orgname)

            continue


        last_name = (
            _element_text(addressbook, "last-name")
            or ""
        )

        first_name = (
            _element_text(addressbook, "first-name")
            or ""
        )

        full_name = " ".join(

            part
            for part in [first_name, last_name]
            if part

        )

        if full_name:

            names.append(full_name)


    return "; ".join(names) if names else None


def _extract_classification_symbols(

    bibliographic,

    entry_tag,

):

    """
    Extracts CPC or IPC classification symbols. Works for
    both entry_tag="classification-cpc" (nested under
    classifications-cpc/main-cpc or .../further-cpc) and
    entry_tag="classification-ipcr" (nested directly under
    classifications-ipcr) since it searches for the entry
    tag anywhere beneath bibliographic, regardless of the
    intermediate wrapper structure.

    Formats each entry as e.g. "F41H1/04" (section + class +
    subclass + main-group + "/" + subgroup), matching
    standard CPC/IPC symbol notation and the format used in
    the CPC search-matrix Excel workbook.

    Correctly returns None for records with no classification
    of this type at all -- e.g. design patents, which use
    Locarno/US-National classification instead of CPC.
    """

    if bibliographic is None:

        return None

    symbols = []

    for entry in bibliographic.findall(f".//{entry_tag}"):

        section = _element_text(entry, "section") or ""

        klass = _element_text(entry, "class") or ""

        subclass = _element_text(entry, "subclass") or ""

        main_group = (
            _element_text(entry, "main-group") or ""
        )

        subgroup = _element_text(entry, "subgroup") or ""

        symbol = f"{section}{klass}{subclass}{main_group}"

        if subgroup:

            symbol += f"/{subgroup}"

        if symbol.strip():

            symbols.append(symbol)


    return "; ".join(symbols) if symbols else None


# ============================================================
# SGML PARSER
# ============================================================

def iter_sgml_records(file_path):

    """
    Parses a USPTO ST.32 SGML bulk grant file (the "PATDOC"
    DTD family, seen with dtd-version like "2.4 2000-09-20").

    This is NOT valid XML -- a handful of tags (e.g. <EMI>,
    <CITED-BY-EXAMINER>, occasional empty leaf tags like
    <B597US>) are never explicitly closed, which a strict XML
    parser chokes on. This uses a small lenient tree builder
    on top of the standard library's html.parser instead (see
    _LenientSGMLTreeBuilder below), which recovers from an
    unclosed tag by searching up the open-tag stack for the
    nearest ancestor that DOES match the current end tag.

    Records are separated by "<!DOCTYPE PATDOC" declarations.

    Built and verified against real samples spanning DESIGN,
    PLANT, and REISSUE (utility-family) patents (dtd-version
    2.4, 2000-09-20). Confirmed structure:
      - Patent/publication number: SDOBI/B100/B110/DNUM
      - Kind code: SDOBI/B100/B130 ("S..."=design, "P"=plant,
        "E"=reissue, "A"/"B1"/"B2"=utility)
      - Grant date: SDOBI/B100/B140/DATE
      - Filing date: SDOBI/B200/B220/DATE
      - Title: SDOBI/B500/B540/STEXT
      - Inventors: SDOBI/B700/B720/B721/PARTY-US/NAM
      - Assignees: SDOBI/B700/B730/B731/PARTY-US/NAM
      - Abstract: SDOAB (flattened) -- confirmed via reissue
        sample; NOT present at all for design/plant records
      - Description: SDODE (flattened)
      - Claims: SDOCL (flattened)

    NOTE: no CPC exists in this era (pre-2013), same as the
    legacy APS parser. For DESIGN patents (kind starts with
    "S"), classification is Locarno, not IPC, so `ipc` is
    intentionally left None for those -- consistent with how
    design patents are handled in the XML and APS parsers.
    For non-design patents, IPC lives under B500/B510/B51x --
    CONFIRMED against two real samples: a plant patent (single
    "A01H  500" -> "A01H5/00") and a reissue patent that
    carried TWO IPC classifications at once (B511 AND B512),
    which this now correctly gathers both of ("E21B10/60;
    E21B10/46") instead of only the first. B516 (the
    classification edition number, e.g. "7") is explicitly
    excluded even though it matches the same B51<digits>
    naming pattern.

    KNOWN IMPERFECTION: reissue-patent claim/description text
    can contain USPTO's redline markup for what changed during
    reissue (<DEL-S>/<DEL-E> for deleted text, <INS-S>/<INS-E>
    for inserted text). This flattens all of it together
    (e.g. a claim reading "1. A A drill bit, comprising:" --
    the struck-through "A" plus the inserted replacement text
    both end up in the output). This only affects the
    relatively small number of reissue patents specifically,
    and doesn't break keyword search, but the text isn't a
    clean read of "the current claim language only."
    """

    source_file_name = file_path.name

    current_lines = []

    with open(

        file_path,

        "r",

        encoding="utf-8",

        errors="replace",

    ) as file:

        for line in file:

            if (

                line.lstrip().startswith("<!DOCTYPE PATDOC")

                and current_lines

            ):

                record = _parse_single_sgml_record(

                    "".join(current_lines),
                    source_file_name,

                )

                if record:

                    yield record

                current_lines = [line]

            else:

                current_lines.append(line)


    if current_lines:

        record = _parse_single_sgml_record(

            "".join(current_lines),
            source_file_name,

        )

        if record:

            yield record


def parse_sgml_file(file_path):

    """
    List-returning wrapper around iter_sgml_records(), kept for ad-hoc
    testing/inspection. The --database stage uses the generator
    directly so a whole file's records are never held in memory
    at once.
    """

    return list(iter_sgml_records(file_path))


# ------------------------------------------------------------
# SGML PARSING HELPERS
# ------------------------------------------------------------

class _LenientSGMLTreeBuilder(HTMLParser):

    """
    Builds a simple tree out of USPTO ST.32 SGML. NOT a
    general-purpose SGML parser -- just tolerant enough to
    survive this specific format's quirks. DOCTYPE/ENTITY
    declarations are handled natively by html.parser (routed
    to handle_decl, which this doesn't override, so they're
    silently and correctly ignored).

    When an end tag doesn't match the top of the open-tag
    stack (i.e. the matching start tag was never explicitly
    closed), this searches upward for the nearest ancestor
    that DOES match and closes everything back to that point.
    That's exactly the recovery needed for tags like <EMI> or
    <B597US>, which this format never closes.
    """

    def __init__(self):

        super().__init__(convert_charrefs=True)

        self.root = {"tag": "#ROOT", "children": [], "text": ""}

        self.stack = [self.root]


    def handle_starttag(self, tag, attrs):

        node = {

            "tag": tag.upper(),
            "children": [],
            "text": "",

        }

        self.stack[-1]["children"].append(node)

        self.stack.append(node)


    def handle_startendtag(self, tag, attrs):

        node = {

            "tag": tag.upper(),
            "children": [],
            "text": "",

        }

        self.stack[-1]["children"].append(node)


    def handle_endtag(self, tag):

        tag = tag.upper()

        for index in range(len(self.stack) - 1, 0, -1):

            if self.stack[index]["tag"] == tag:

                del self.stack[index:]

                return


    def handle_data(self, data):

        self.stack[-1]["text"] += data


def _sgml_find_first(

    node,

    *tag_path

):

    """
    Descends through node's DIRECT children along tag_path
    (e.g. _sgml_find_first(root, "B100", "B110", "DNUM",
    "PDAT")), returning the final node's text, or None if
    the path doesn't exist anywhere it's tried.
    """

    if node is None:

        return None

    current = node

    for tag in tag_path:

        next_node = None

        for child in current["children"]:

            if child["tag"] == tag:

                next_node = child

                break

        if next_node is None:

            return None

        current = next_node

    text = current["text"].strip()

    return text if text else None


def _sgml_find_all(

    node,

    tag

):

    """
    Recursively finds ALL descendant nodes with the given
    tag, anywhere in the subtree regardless of depth.
    """

    results = []

    for child in node["children"]:

        if child["tag"] == tag:

            results.append(child)

        results.extend(
            _sgml_find_all(child, tag)
        )

    return results


def _sgml_full_text(

    node

):

    """
    Flattens ALL text anywhere within a subtree into one
    string, collapsing whitespace. Mirrors
    _element_full_text() for the XML parser.
    """

    if node is None:

        return None

    parts = [node["text"]]

    for child in node["children"]:

        child_text = _sgml_full_text(child)

        if child_text:

            parts.append(child_text)

    text = " ".join(
        part for part in parts if part
    )

    text = re.sub(r"\s+", " ", text).strip()

    return text if text else None


def _parse_single_sgml_record(

    sgml_text,

    source_file_name,

):

    # Drawings section contains image references (<EMI>) we
    # don't need -- stripped up front to reduce parsing risk
    # for a section that has nothing we use.

    sgml_text = re.sub(

        r"<SDODR.*?</SDODR>",
        "",
        sgml_text,
        flags=re.DOTALL,

    )

    # DOCTYPE/internal ENTITY declarations aren't needed
    # either (image filenames).

    sgml_text = re.sub(

        r"<!DOCTYPE.*?\]>",
        "",
        sgml_text,
        flags=re.DOTALL,

    )


    builder = _LenientSGMLTreeBuilder()

    try:

        builder.feed(sgml_text)

    except Exception as error:

        logging.warning(

            "Skipping malformed SGML record in %s: %s",
            source_file_name,
            error,

        )

        return None


    patdoc = None

    for child in builder.root["children"]:

        if child["tag"] == "PATDOC":

            patdoc = child

            break

    if patdoc is None:

        return None


    doc_number = _sgml_find_first(

        patdoc,
        "SDOBI", "B100", "B110", "DNUM", "PDAT",

    )

    if not doc_number:

        return None


    kind = (

        _sgml_find_first(

            patdoc,
            "SDOBI", "B100", "B130", "PDAT",

        )

        or ""

    )

    grant_date_raw = _sgml_find_first(

        patdoc,
        "SDOBI", "B100", "B140", "DATE", "PDAT",

    )

    filing_date_raw = _sgml_find_first(

        patdoc,
        "SDOBI", "B200", "B220", "DATE", "PDAT",

    )

    title = _sgml_find_first(

        patdoc,
        "SDOBI", "B500", "B540", "STEXT", "PDAT",

    )


    patent_number = f"US{doc_number}{kind}"

    is_design_patent = kind.strip().upper().startswith("S")

    # SGML-era files are grants only (same as APS -- USPTO
    # didn't publish pre-grant applications until AIPA/2001),
    # but "grant" isn't quite accurate for every kind code.
    # A Statutory Invention Registration (kind "H") is
    # explicitly NOT a patent -- the record itself says so
    # ("A statutory invention registration is not a patent.
    # It has the defensive attributes of a patent but does
    # not have the enforceable attributes of a patent.").
    # Tagged distinctly so it can be included or filtered out
    # later without losing the record.
    document_type = (

        "sir"

        if kind.strip().upper().startswith("H")

        else "grant"

    )


    # Classification. Design patents use Locarno here, not
    # IPC -- correctly left as None, same as the XML/APS
    # parsers do for design patents. Non-design IPC extraction
    # confirmed against real samples: a plant patent (B511 ->
    # "A01H5/00") and a reissue patent that carried TWO IPC
    # classifications (B511 AND B512 -- so this walks every
    # direct child of B510 matching "B51<digits>", not just
    # B511, since a record can have more than one).

    ipc = None

    if not is_design_patent:

        b510 = None

        for child in _sgml_find_all(patdoc, "B510"):

            b510 = child

            break


        if b510 is not None:

            symbols = []

            b51x_pattern = re.compile(r"^B51\d+$")

            for child in b510["children"]:

                if not b51x_pattern.match(child["tag"]):

                    continue

                if child["tag"] == "B516":

                    # Edition indicator (e.g. "7"), not a
                    # classification symbol -- exclude
                    # explicitly, since it otherwise matches
                    # the same B51<digits> pattern.

                    continue

                raw_value = _sgml_find_first(
                    child,
                    "PDAT",
                )

                if raw_value:

                    symbol = (
                        _format_aps_classification_symbol(
                            raw_value
                        )
                    )

                    if symbol and symbol not in symbols:

                        symbols.append(symbol)

            ipc = "; ".join(symbols) if symbols else None


    # Inventors (B721) / Assignees (B731) -- both can repeat.

    def _extract_sgml_party_names(container_tag):

        names = []

        for node in _sgml_find_all(patdoc, container_tag):

            org = _sgml_find_first(

                node,
                "PARTY-US", "NAM", "ONM", "STEXT", "PDAT",

            )

            if org:

                names.append(org)

                continue


            first = _sgml_find_first(

                node,
                "PARTY-US", "NAM", "FNM", "PDAT",

            )

            last = _sgml_find_first(

                node,
                "PARTY-US", "NAM", "SNM", "STEXT", "PDAT",

            )

            full_name = " ".join(

                part for part in [first, last] if part

            )

            if full_name:

                names.append(full_name)


        return "; ".join(names) if names else None


    inventors = _extract_sgml_party_names("B721")

    assignees = _extract_sgml_party_names("B731")


    abstract = None

    description = None

    claims = None

    for child in patdoc["children"]:

        if child["tag"] == "SDOAB":

            abstract = _sgml_full_text(child)

        elif child["tag"] == "SDODE":

            description = _sgml_full_text(child)

        elif child["tag"] == "SDOCL":

            claims = _sgml_full_text(child)


    return {

        "patent_number": patent_number,
        "publication_number": patent_number,
        "document_type": document_type,
        "title": title,
        "abstract": abstract,
        "claims": claims,
        "description": description,
        "inventors": inventors,
        "applicants": None,
        "assignees": assignees,
        "filing_date": _normalize_xml_date(filing_date_raw),
        "publication_date": _normalize_xml_date(grant_date_raw),
        "grant_date": _normalize_xml_date(grant_date_raw),
        "priority_date": None,
        "cpc": None,
        "ipc": ipc,
        "source_file": source_file_name,

    }


# ============================================================
# EARLY PATENT APPLICATION PUBLICATION PARSER (PAP, 2001-2002ish)
# ============================================================
#
# USPTO only started publishing pre-grant applications at all
# in March 2001 (under the AIPA). The VERY EARLIEST files (this
# was confirmed against a real sample: doc-number 20010000001,
# dated 2001-03-15) used their own transitional DTD family --
# "pap-v15-2001-01-31.dtd", root element
# <patent-application-publication> -- distinct from BOTH the
# later modern "v4.x" us-patent-application format AND the
# PATDOC/ST.32 grant vocabulary above. Genuinely different tag
# names (document-id, subdoc-bibliographic-information,
# subdoc-abstract, subdoc-description, subdoc-claims, all
# lowercase-hyphenated rather than PATDOC's B1xx codes), so
# this is real new extraction logic, not something that could
# delegate to an existing parser the way the PATDOC case could.
#
# Reuses the SAME lenient tree builder as the SGML/PATDOC path
# (_LenientSGMLTreeBuilder, built on html.parser) rather than
# strict ElementTree, because this format ALSO uses old-style
# named entities ET.fromstring() rejects -- confirmed against
# the real sample, which uses &ldquo; &rdquo; &deg; &Prime;
# &equals; throughout (all correctly resolved by html.parser's
# standard entity table). One NOT fully covered, worth knowing:
# old ISO-8879-style abbreviated Greek-letter entities like
# &agr; (alpha) are left as literal unresolved text -- confirmed
# by direct testing, not assumed. This is a text-fidelity
# footnote for occasional Greek-letter/symbol text (this sample
# uses it for "alpha-pinene" in a chemistry claim), not a
# parsing failure or lost record.
#
# NOTE ON SCOPE: this only extracts what the patents table
# schema stores (bibliographic fields, title, and the three
# full-text sections). The very rich continuity-data
# (continuation/CIP/371 family history) and correspondence-
# address (attorney firm) sections this format carries are
# read but deliberately not captured -- no schema column exists
# for them, matching how continuity/attorney data isn't
# captured for any other format either.

def _parse_single_pap_record(

    pap_text,

    source_file_name,

):

    # Drawings section is image references only -- stripped up
    # front, same reasoning as the PATDOC/SGML path.

    pap_text = re.sub(

        r"<subdoc-drawings.*?</subdoc-drawings>",
        "",
        pap_text,
        flags=re.DOTALL | re.IGNORECASE,

    )

    pap_text = re.sub(

        r"<!DOCTYPE.*?\]>",
        "",
        pap_text,
        flags=re.DOTALL,

    )


    builder = _LenientSGMLTreeBuilder()

    try:

        builder.feed(pap_text)

    except Exception as error:

        logging.warning(

            "Skipping malformed early-PAP record in %s: %s",
            source_file_name,
            error,

        )

        return None


    pub = None

    for child in builder.root["children"]:

        if child["tag"] == "PATENT-APPLICATION-PUBLICATION":

            pub = child

            break

    if pub is None:

        return None


    doc_number = _sgml_find_first(

        pub,
        "SUBDOC-BIBLIOGRAPHIC-INFORMATION", "DOCUMENT-ID",
        "DOC-NUMBER",

    )

    if not doc_number:

        return None


    kind_code = (

        _sgml_find_first(

            pub,
            "SUBDOC-BIBLIOGRAPHIC-INFORMATION", "DOCUMENT-ID",
            "KIND-CODE",

        )

        or ""

    )

    publication_date_raw = _sgml_find_first(

        pub,
        "SUBDOC-BIBLIOGRAPHIC-INFORMATION", "DOCUMENT-ID",
        "DOCUMENT-DATE",

    )

    publication_number = f"US{doc_number}{kind_code}"

    # A pre-grant application publication isn't a patent yet --
    # same convention the modern us-patent-application branch
    # uses (patent_number/grant_date stay None; only granted
    # documents get those filled in).

    patent_number = None

    grant_date = None


    filing_date_raw = _sgml_find_first(

        pub,
        "SUBDOC-BIBLIOGRAPHIC-INFORMATION", "DOMESTIC-FILING-DATA",
        "FILING-DATE",

    )

    title = _sgml_find_first(

        pub,
        "SUBDOC-BIBLIOGRAPHIC-INFORMATION", "TECHNICAL-INFORMATION",
        "TITLE-OF-INVENTION",

    )

    ipc = _sgml_find_first(

        pub,
        "SUBDOC-BIBLIOGRAPHIC-INFORMATION", "TECHNICAL-INFORMATION",
        "CLASSIFICATION-IPC", "CLASSIFICATION-IPC-PRIMARY", "IPC",

    )


    # Foreign priority filing date, when present (confirmed
    # against real data: present in roughly half of a 47-record
    # sample). FOREIGN-PRIORITY-DATA sits as a SIBLING of
    # DOMESTIC-FILING-DATA, and both contain their own
    # FILING-DATE child with the same tag name -- so this scopes
    # to the foreign-priority container first (via _sgml_find_all,
    # which searches the whole subtree by tag) before looking for
    # FILING-DATE within just that node, rather than searching
    # the whole record for "FILING-DATE" and risking a match on
    # the domestic one instead.

    priority_date_raw = None

    for foreign_priority in _sgml_find_all(
        pub, "FOREIGN-PRIORITY-DATA"
    ):

        priority_date_raw = _sgml_find_first(
            foreign_priority, "FILING-DATE"
        )

        if priority_date_raw:
            break


    # Inventors: <first-named-inventor> plus any further
    # <inventor> siblings, each carrying given-name/middle-name/
    # family-name. Found via _sgml_find_all (searches the whole
    # subtree by tag, not a fixed path) since both tags can
    # appear and there's no single fixed parent path shared by
    # both the way B721 covers every inventor in the SGML path.

    inventor_names = []

    for tag in ("FIRST-NAMED-INVENTOR", "INVENTOR"):

        for node in _sgml_find_all(pub, tag):

            given = _sgml_find_first(node, "NAME", "GIVEN-NAME")

            middle = _sgml_find_first(node, "NAME", "MIDDLE-NAME")

            family = _sgml_find_first(node, "NAME", "FAMILY-NAME")

            full_name = " ".join(

                part for part in [given, middle, family] if part

            )

            if full_name:

                inventor_names.append(full_name)

    inventors = "; ".join(inventor_names) if inventor_names else None


    abstract = None

    description = None

    claims = None

    for child in pub["children"]:

        if child["tag"] == "SUBDOC-ABSTRACT":

            abstract = _sgml_full_text(child)

        elif child["tag"] == "SUBDOC-DESCRIPTION":

            description = _sgml_full_text(child)

        elif child["tag"] == "SUBDOC-CLAIMS":

            claims = _sgml_full_text(child)


    return {

        "patent_number": patent_number,
        "publication_number": publication_number,
        "document_type": "application",
        "title": title,
        "abstract": abstract,
        "claims": claims,
        "description": description,
        "inventors": inventors,
        "applicants": None,
        "assignees": None,
        "filing_date": _normalize_xml_date(filing_date_raw),
        "publication_date": _normalize_xml_date(publication_date_raw),
        "grant_date": grant_date,
        "priority_date": _normalize_xml_date(priority_date_raw),
        "cpc": None,
        "ipc": ipc,
        "source_file": source_file_name,

    }


# ============================================================
# LEGACY APS PARSER
# ============================================================

def iter_legacy_records(file_path):

    """
    Parses a USPTO APS/Green Book bulk file (1976-2001ish).

    This is NOT XML or SGML -- it's a tagged flat-text format:
    each field is a line starting with a 2-4 letter uppercase
    tag at column 1 (e.g. "WKU  RE0286710"), and continuation
    lines are indented with no tag. Records are separated by
    a line that is exactly "PATN". A file-level header line
    (e.g. "HHHHHT  APS1  ISSUE - 760106") precedes the first
    record and is skipped.

    Fields are also SECTION-scoped: the same tag (e.g. "NAM",
    "PNO", "APD") means different things depending on which
    section it appears in (top-level bibliographic data vs.
    an inventor block vs. a cited-reference block vs. a
    foreign-priority block, etc.) -- so this tracks the
    current section while scanning and routes each field
    accordingly, rather than treating the file as one flat
    key/value dict.

    IMPORTANT, confirmed from a real 1976 sample file:
      - There is NO CPC classification in this era at all
        (CPC didn't exist until 2013) -- only IPC (the "ICL"
        field) and the deprecated US National Classification
        (OCL/XCL, which this parser does not attempt to
        extract, since we don't use it for filtering).
      - APS-era files are GRANTS ONLY. USPTO didn't publish
        pre-grant applications until AIPA / 2001, so there is
        no "legacy applications" product to worry about here.
      - This means every legacy record's `cpc` field will
        always come back None. matches_cpc() falls back to
        checking `ipc` for records with no CPC (see
        MATCH_IPC_AS_CPC_FALLBACK), so legacy records can
        still be matched via their IPC classification.
    """

    source_file_name = file_path.name

    current_record_lines = []

    in_record = False

    with open(

        file_path,

        "r",

        encoding="utf-8",

        errors="replace",

    ) as file:

        for raw_line in file:

            line = raw_line.rstrip("\r\n")

            if line == "PATN":

                if in_record and current_record_lines:

                    record = _parse_single_legacy_record(

                        current_record_lines,
                        source_file_name,

                    )

                    if record:

                        yield record

                current_record_lines = []

                in_record = True

                continue


            if in_record:

                current_record_lines.append(line)


    if in_record and current_record_lines:

        record = _parse_single_legacy_record(

            current_record_lines,
            source_file_name,

        )

        if record:

            yield record


def parse_legacy_file(file_path):

    """
    List-returning wrapper around iter_legacy_records(), kept for ad-hoc
    testing/inspection. The --database stage uses the generator
    directly so a whole file's records are never held in memory
    at once.
    """

    return list(iter_legacy_records(file_path))


# ------------------------------------------------------------
# LEGACY (APS) PARSING HELPERS
# ------------------------------------------------------------

# Section-marker lines that have no value on the same line
# and change what subsequent tagged fields mean.
_APS_SECTION_MARKERS = {

    "INVT", "ASSG", "REIS", "CLAS", "UREF", "OREF", "FREF",
    "LREP", "PARN", "GOVT", "BSUM", "DETD", "DRWD", "CLMS",
    "ABST", "PRIR", "PCTA", "RLAP",

}

# Sections whose paragraph-style text (PAC/PAR/PA1/STM/PAL,
# plus continuation lines) should be flattened into one of
# our three big free-text fields.
_APS_DESCRIPTION_SECTIONS = {"BSUM", "DETD", "DRWD"}

_APS_TAG_LINE_PATTERN = re.compile(

    r"^([A-Z0-9]{2,4})(?:  (.*))?$"

)


def _parse_single_legacy_record(

    lines,

    source_file_name,

):

    top_fields = {}

    inventors = []

    assignees = []

    priority_date = None

    ipc_symbols = []

    abstract_parts = []

    description_parts = []

    claims_parts = []


    current_section = "PATN"

    current_inventor = None

    current_assignee = None

    # What the next continuation (indented, untagged) line
    # should be appended to.
    continuation_target = None


    def start_continuation_target(getter, setter):

        nonlocal continuation_target

        continuation_target = (getter, setter)


    for line in lines:

        if not line.strip():

            continuation_target = None

            continue


        is_continuation = line[0] in (" ", "\t")


        if is_continuation:

            text = line.strip()

            if continuation_target is not None:

                getter, setter = continuation_target

                setter(
                    (getter() + " " + text).strip()
                )

            elif current_section in _APS_DESCRIPTION_SECTIONS:

                description_parts.append(text)

            elif current_section == "CLMS":

                claims_parts.append(text)

            elif current_section == "ABST":

                abstract_parts.append(text)

            continue


        match = _APS_TAG_LINE_PATTERN.match(line)

        if not match:

            # Doesn't look like a tag line -- treat
            # defensively as a continuation of whatever
            # section we're in.

            if current_section in _APS_DESCRIPTION_SECTIONS:

                description_parts.append(line.strip())

            elif current_section == "CLMS":

                claims_parts.append(line.strip())

            elif current_section == "ABST":

                abstract_parts.append(line.strip())

            continue


        tag = match.group(1)

        value = (match.group(2) or "").strip()


        # ------------------------------------------------
        # SECTION MARKERS (no value on their own line)
        # ------------------------------------------------

        if tag in _APS_SECTION_MARKERS and match.group(2) is None:

            current_section = tag

            continuation_target = None


            if tag == "INVT":

                current_inventor = {"name": ""}

                inventors.append(current_inventor)

            elif tag == "ASSG":

                current_assignee = {"name": ""}

                assignees.append(current_assignee)


            continue


        # ------------------------------------------------
        # TOP-LEVEL BIBLIOGRAPHIC FIELDS
        # ------------------------------------------------

        if current_section == "PATN":

            if tag == "WKU":

                top_fields["wku"] = value

            elif tag == "APD":

                top_fields["filing_date"] = value

            elif tag == "ISD":

                top_fields["grant_date"] = value

            elif tag == "TTL":

                top_fields["title"] = value

                start_continuation_target(

                    lambda: top_fields.get("title", ""),
                    lambda v: top_fields.__setitem__(
                        "title", v
                    ),

                )

            continue


        # ------------------------------------------------
        # INVENTOR BLOCK
        # ------------------------------------------------

        if current_section == "INVT":

            if tag == "NAM" and current_inventor is not None:

                current_inventor["name"] = value

                start_continuation_target(

                    lambda: current_inventor["name"],
                    lambda v: current_inventor.__setitem__(
                        "name", v
                    ),

                )

            continue


        # ------------------------------------------------
        # ASSIGNEE BLOCK
        # ------------------------------------------------

        if current_section == "ASSG":

            if tag == "NAM" and current_assignee is not None:

                current_assignee["name"] = value

                start_continuation_target(

                    lambda: current_assignee["name"],
                    lambda v: current_assignee.__setitem__(
                        "name", v
                    ),

                )

            continue


        # ------------------------------------------------
        # CLASSIFICATION BLOCK (IPC only -- no CPC pre-2013)
        # ------------------------------------------------

        if current_section == "CLAS":

            if tag == "ICL":

                symbol = (
                    _format_aps_classification_symbol(value)
                )

                if symbol:

                    ipc_symbols.append(symbol)

            continue


        # ------------------------------------------------
        # FOREIGN PRIORITY BLOCK (first one wins)
        # ------------------------------------------------

        if current_section == "PRIR":

            if tag == "APD" and priority_date is None:

                priority_date = value

            continue


        # ------------------------------------------------
        # ABSTRACT
        # ------------------------------------------------

        if current_section == "ABST":

            if tag == "PAL":

                abstract_parts.append(value)

                start_continuation_target(

                    lambda: abstract_parts[-1],
                    lambda v: abstract_parts.__setitem__(
                        -1, v
                    ),

                )

            continue


        # ------------------------------------------------
        # DESCRIPTION (background summary / detailed
        # description / brief description of drawings)
        # ------------------------------------------------

        if current_section in _APS_DESCRIPTION_SECTIONS:

            if tag in ("PAC", "PAR", "PA1"):

                description_parts.append(value)

                start_continuation_target(

                    lambda: description_parts[-1],
                    lambda v: description_parts.__setitem__(
                        -1, v
                    ),

                )

            continue


        # ------------------------------------------------
        # CLAIMS
        # ------------------------------------------------

        if current_section == "CLMS":

            if tag in ("STM", "PAR", "PA1"):

                claims_parts.append(value)

                start_continuation_target(

                    lambda: claims_parts[-1],
                    lambda v: claims_parts.__setitem__(
                        -1, v
                    ),

                )

            # NUM (claim number, e.g. "1.") is intentionally
            # skipped -- the following PAR line already
            # repeats the number as part of its own text in
            # this era's convention.

            continue


        # Everything else (REIS, UREF, OREF, FREF, LREP,
        # PARN, GOVT, RLAP, PCTA) isn't needed for our
        # schema and is intentionally ignored.


    wku = top_fields.get("wku")

    if not wku:

        return None


    patent_number = f"US{wku}"


    inventor_names = [

        formatted

        for formatted in (
            _format_aps_name(person["name"])
            for person in inventors
        )

        if formatted

    ]

    assignee_names = [

        formatted

        for formatted in (
            _format_aps_name(person["name"])
            for person in assignees
        )

        if formatted

    ]


    return {

        "patent_number": patent_number,
        "publication_number": patent_number,
        "document_type": "grant",
        "title": top_fields.get("title") or None,
        "abstract": (
            " ".join(abstract_parts) if abstract_parts else None
        ),
        "claims": (
            " ".join(claims_parts) if claims_parts else None
        ),
        "description": (
            " ".join(description_parts)
            if description_parts else None
        ),
        "inventors": (
            "; ".join(inventor_names) if inventor_names else None
        ),
        "applicants": None,
        "assignees": (
            "; ".join(assignee_names) if assignee_names else None
        ),
        "filing_date": _normalize_xml_date(
            top_fields.get("filing_date")
        ),
        "publication_date": _normalize_xml_date(
            top_fields.get("grant_date")
        ),
        "grant_date": _normalize_xml_date(
            top_fields.get("grant_date")
        ),
        "priority_date": _normalize_xml_date(priority_date),
        "cpc": None,
        "ipc": (
            "; ".join(ipc_symbols) if ipc_symbols else None
        ),
        "source_file": source_file_name,

    }


def _format_aps_name(

    raw_name

):

    """
    APS names come as either "Lastname; Firstname" for
    individuals or a plain organization name for assignees
    (no semicolon). Reformats individual names to
    "Firstname Lastname" to match the format used by the XML
    parser, so the `inventors`/`assignees` columns look
    consistent across both eras in the database. Passes
    organization names through unchanged.
    """

    if not raw_name:

        return None

    raw_name = raw_name.strip()

    if not raw_name:

        return None

    if ";" in raw_name:

        last, _, first = raw_name.partition(";")

        last = last.strip()

        first = first.strip()

        if first:

            return f"{first} {last}".strip()

        return last or None

    return raw_name


def _format_aps_classification_symbol(

    raw_value

):

    """
    Converts an APS-format ICL (IPC) value like "B63B 2152"
    or "E05D  700" into standard compact notation
    ("B63B21/52", "E05D7/00"): a 4-character class code,
    followed by the main-group (everything remaining except
    the last 2 characters), a "/", then the subgroup (the
    last 2 characters).

    NOTE: this field-width interpretation was reverse
    engineered from two real examples in the sample file
    rather than from official APS format documentation --
    it matched both examples exactly, but if classification
    symbols start looking wrong once run against a full real
    file, this is the first place to double check.
    """

    if not raw_value:

        return None

    raw_value = raw_value.rstrip()

    if len(raw_value) <= 4:

        return raw_value.strip() or None

    class_code = raw_value[:4].strip()

    remainder = raw_value[4:].strip()

    if len(remainder) < 2:

        return (

            f"{class_code}{remainder}"
            if remainder
            else (class_code or None)

        )

    main_group = remainder[:-2].strip()

    subgroup = remainder[-2:]

    if main_group:

        return f"{class_code}{main_group}/{subgroup}"

    return f"{class_code}/{subgroup}"


# ------------------------------------------------------------
# FILE DISCOVERY
# ------------------------------------------------------------

# Housekeeping / non-patent file types that should never be
# handed to a parser, even though they live under EXTRACTED_DIR.
IGNORED_FILE_SUFFIXES = {

    ".json", ".zip", ".part", ".log", ".csv",
    ".xlsx", ".db", ".sqlite", ".sqlite-wal", ".sqlite-shm",

}

IGNORED_FILE_NAMES = {

    ".EXTRACTION_COMPLETE",

}



def discover_patent_files(search_root=None):

    """
    Single pass over search_root (default: EXTRACTED_DIR),
    classifying every file.

    Anything that isn't XML/SGML and isn't a known
    housekeeping file is treated as a legacy (APS/Green Book)
    candidate and handed to parse_legacy_file() -- those
    older files typically don't carry a modern file extension
    at all, so a suffix-based glob (like *.xml/*.sgml/*.sgm
    alone) would silently skip the entire 1976-early-1990s
    slice of the corpus.

    search_root lets build_database(product_format=...) scope this to
    just one product's extracted files instead of the whole
    corpus -- see DATABASE_FORMAT_TO_RAW_DIR below for how that
    directory is computed.
    """

    if search_root is None:
        search_root = EXTRACTED_DIR

    files = []

    for file_path in search_root.rglob("*"):

        if not file_path.is_file():

            continue

        if file_path.name in IGNORED_FILE_NAMES:

            continue

        if file_path.suffix.lower() in IGNORED_FILE_SUFFIXES:

            continue

        files.append(file_path)

    return files


# ============================================================
# DATABASE MANIFEST (file-level resumability)
# ============================================================
#
# Unlike the download and PDF stages, --database previously had
# no way to know which files it had already parsed -- it
# rewalked and reparsed the ENTIRE EXTRACTED_DIR on every run,
# relying on INSERT OR IGNORE to silently discard rows that
# were already in the database. That means the expensive part
# (parsing) always happened again even for files whose data was
# already fully captured; only the cheap part (the actual row
# insert) was being skipped. At corpus scale, re-running
# --database after adding a new batch of downloaded files meant
# re-parsing your entire historical corpus just to pick up the
# new files.
#
# This manifest tracks path + size + mtime per file, matching
# the pattern already used for DOWNLOAD_MANIFEST/PDF_MANIFEST.
# A file is only re-parsed if it's new, or if its size no
# longer matches what's recorded (re-extracted/replaced file).
# Files that errored are NOT marked successful, so they're
# automatically retried on the next run rather than silently
# skipped forever.
# ------------------------------------------------------------

DATABASE_MANIFEST_FIELDS = [

    "path",
    "size",
    "mtime",
    "records_inserted",
    "status",
    "error",
    "processed_time",

]


def load_database_manifest():

    return _load_csv_manifest(

        DATABASE_MANIFEST,

        build_key=lambda row: row["path"],

    )


def save_database_manifest(
    records
):

    _save_csv_manifest(

        DATABASE_MANIFEST,
        records,
        DATABASE_MANIFEST_FIELDS,

    )


# ============================================================
# PARSE ONE FILE INTO THE DATABASE (runs in a worker PROCESS)
# ============================================================

def _lower_worker_priority():

    """
    ProcessPoolExecutor initializer: drops each parsing worker
    to BELOW_NORMAL priority on Windows so the desktop stays
    responsive during a multi-hour build. This does not make
    the build slower when nothing else is running -- the
    workers still get every idle CPU cycle. (No-op elsewhere.)
    """

    if os.name != "nt":
        return

    try:

        import ctypes

        below_normal_priority_class = 0x00004000

        ctypes.windll.kernel32.SetPriorityClass(
            ctypes.windll.kernel32.GetCurrentProcess(),
            below_normal_priority_class,
        )

    except Exception:

        pass


def _parse_file_into_database(file_path, db_path):

    """
    Streams one extracted file through its parser and writes
    the records straight into db_path in batches of
    DB_INSERT_BATCH_RECORDS -- then returns only a small status
    dict (counts, no record data) to the main process.

    v14 returned every parsed record to the main process
    instead. For a modern weekly XML file that meant building,
    pickling, piping, and un-pickling several hundred MB to a
    few GB of text per file, with up to one file per worker in
    flight plus whatever results were queued in the main
    process -- which is what drove a 16 GB machine into
    memory exhaustion / heavy paging at the XML stage.

    Multiple workers write to the same product database;
    SQLite serializes them (one writer at a time), and each
    batch's write takes a fraction of a second against a parse
    that takes tens of seconds, so they rarely wait on each
    other. The manifest stays single-writer in the main process.

    If a file fails part-way, the batches already committed
    stay in the database and the file is marked "error"; on
    retry, INSERT OR IGNORE skips the records that are already
    there, so nothing is duplicated.

    error_kind distinguishes a normal per-file parse problem
    ("parse") from SQLite failing to write ("database") -- the
    latter stops the whole run (see build_database()).
    """

    result = {
        "path": str(file_path),
        "size": 0,
        "mtime": 0.0,
        "parsed": 0,
        "inserted": 0,
        "error": None,
        "error_kind": None,
    }

    try:

        stat = file_path.stat()

        result["size"] = stat.st_size

        result["mtime"] = stat.st_mtime

    except OSError as error:

        result["error"] = f"{type(error).__name__}: {error}"

        result["error_kind"] = "parse"

        return result

    connection = None

    try:

        connection = get_db_connection(db_path)

        batch = []

        for record in iter_patent_records(file_path):

            batch.append(record)

            result["parsed"] += 1

            if len(batch) >= DB_INSERT_BATCH_RECORDS:

                result["inserted"] += _insert_record_batch(
                    connection, batch, str(file_path)
                )

                batch = []

        if batch:

            result["inserted"] += _insert_record_batch(
                connection, batch, str(file_path)
            )

    except sqlite3.Error as error:

        result["error"] = f"{type(error).__name__}: {error}"

        result["error_kind"] = "database"

    except Exception as error:

        result["error"] = f"{type(error).__name__}: {error}"

        result["error_kind"] = "parse"

    finally:

        if connection is not None:

            try:
                connection.close()
            except sqlite3.Error:
                pass

    return result


# ============================================================
# BUILD DATABASE
# ============================================================

# Maps a --format value to the RAW-side directory download_product()
# uses for that product. The EXTRACTED-side equivalent is computed
# the same way _extract_one_archive() computes it for each archive
# (EXTRACTED_DIR plus that RAW directory's path relative to
# RAW_DIR), so it can't drift out of sync with extraction.
DATABASE_FORMAT_TO_RAW_DIR = {
    "aps": GRANTS_APS_DIR,
    "sgml": GRANTS_SGML_DIR,
    "xml": GRANTS_XML_DIR,
    "appxml": APPLICATIONS_DIR,
}

DATABASE_PRODUCT_ORDER = ["aps", "sgml", "xml", "appxml"]

MIGRATION_STATE_FILE = DATABASE_DIR / "migration_v15_state.json"


def _product_extracted_dir(product):

    return EXTRACTED_DIR / DATABASE_FORMAT_TO_RAW_DIR[product].relative_to(RAW_DIR)


def _load_migration_state():

    if not MIGRATION_STATE_FILE.exists():
        return {}

    try:
        return json.loads(MIGRATION_STATE_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def _save_migration_state(state):

    temp_path = MIGRATION_STATE_FILE.with_suffix(".tmp")

    temp_path.write_text(json.dumps(state, indent=2), encoding="utf-8")

    os.replace(temp_path, MIGRATION_STATE_FILE)


def _legacy_database_needs_migration():

    """
    True if a v14 single-file database exists and hasn't been
    fully copied into the per-product files yet. While that's
    the case, files the manifest already marks "success" have
    their rows ONLY in the old file -- so --database and
    --filter refuse to run rather than silently skip them.
    """

    if not LEGACY_DATABASE_FILE.exists():
        return False

    return not _load_migration_state().get("complete", False)


def _print_migration_required():

    print()
    print(f"  Found the v14 single-file database: {LEGACY_DATABASE_FILE}")
    print("  v15 stores each product in its own file. Before --database or")
    print("  --filter can run, either:")
    print()
    print("    (a) copy the existing rows over (no re-parsing needed):")
    print("          python uspto_patent_pipeline_v15.py --migrate-db")
    print()
    print("    (b) or start clean: delete that file AND")
    print(f"          {DATABASE_MANIFEST}")
    print("        and every product will be re-parsed from scratch.")
    print()


def build_database(
    limit=None,
    skip_errors=False,
    retry_errors_only=False,
    product_format=None,
    backup=False,
    workers=None,
):

    """
    Parses extracted files into the per-product databases.
    With --format, only that product; otherwise every product
    in turn (APS -> SGML -> XML -> APPXML), each into its own
    file. --limit caps the total number of files this run
    across all products.
    """

    if skip_errors and retry_errors_only:

        print(
            "  --skip-errors and --retry-errors-only can't be used "
            "together (that would skip everything). Use one or "
            "the other."
        )

        return

    if _legacy_database_needs_migration():

        _print_migration_required()

        return

    if product_format is not None and product_format not in DATABASE_FILES:

        print(
            f"  Unknown --format '{product_format}'. Choose from: "
            f"{', '.join(DATABASE_FILES)}."
        )

        return

    worker_count = workers or DATABASE_PARSE_WORKERS

    products = [product_format] if product_format else DATABASE_PRODUCT_ORDER

    manifest = load_database_manifest()

    remaining_limit = limit

    for product in products:

        outcome = _build_product_database(
            product=product,
            manifest=manifest,
            limit=remaining_limit,
            skip_errors=skip_errors,
            retry_errors_only=retry_errors_only,
            backup=backup,
            worker_count=worker_count,
        )

        if outcome["fatal"]:

            break

        if remaining_limit is not None:

            remaining_limit -= outcome["files_attempted"]

            if remaining_limit <= 0:

                print(
                    f"  --limit {limit} reached -- stopping. Run "
                    f"--database again to continue."
                )

                break


def _build_product_database(
    product,
    manifest,
    limit,
    skip_errors,
    retry_errors_only,
    backup,
    worker_count,
):

    outcome = {"fatal": False, "files_attempted": 0}

    db_path = DATABASE_FILES[product]

    search_root = _product_extracted_dir(product)

    print()
    print(f"  [{product}] {search_root}  ->  {db_path.name}")

    if not search_root.exists():

        print("    Nothing extracted for this product yet -- skipping.")

        return outcome

    create_database(db_path)

    files = sorted(discover_patent_files(search_root))

    logging.info("[%s] Candidate patent files found: %s", product, len(files))

    # --------------------------------------------------------
    # SKIP FILES ALREADY PROCESSED (manifest-based resumability)
    # --------------------------------------------------------
    # A file needs parsing if it's new, or its recorded size no
    # longer matches (re-extracted/replaced). Previously-errored
    # files aren't "done" and are retried automatically.

    pending_files = []
    skipped_count = 0
    previously_errored_count = 0

    for file_path in files:

        manifest_row = manifest.get(str(file_path))

        row_status = manifest_row.get("status") if manifest_row else None

        if row_status == "success":

            try:
                current_size = file_path.stat().st_size
            except OSError:
                current_size = None

            if str(current_size) == manifest_row.get("size"):
                skipped_count += 1
                continue

        if row_status == "error":

            previously_errored_count += 1

            if skip_errors:
                continue

        elif retry_errors_only:

            continue

        pending_files.append(file_path)

    print(
        f"    Already parsed (skipped): {skipped_count:,}   "
        f"Pending this run: {len(pending_files):,}"
    )

    if previously_errored_count:

        note = (
            "included in the count above"
            if not skip_errors
            else "excluded this run (--skip-errors)"
        )

        print(
            f"    Previously-errored files known: "
            f"{previously_errored_count:,} ({note})"
        )

    total_pending = len(pending_files)

    if limit is not None and limit < total_pending:

        pending_files = pending_files[:limit]

        print(
            f"    --limit: processing {len(pending_files):,} of "
            f"{total_pending:,} pending files this run."
        )

    outcome["files_attempted"] = len(pending_files)

    if not pending_files:

        print("    Nothing new to parse for this product.")

        create_database_indexes(db_path)

        return outcome

    print(
        f"    Parsing with {worker_count} worker process(es), "
        f"streaming {DB_INSERT_BATCH_RECORDS} records per write..."
    )

    total_parsed = 0
    total_inserted = 0
    error_count = 0
    results_since_save = 0
    completed_paths = set()
    pool_restart_count = 0
    fatal_db_error = None

    progress_bar = tqdm(total=len(pending_files), desc=f"Parsing {product}")

    while len(completed_paths) < len(pending_files) and fatal_db_error is None:

        remaining_paths = [
            file_path
            for file_path in pending_files
            if str(file_path) not in completed_paths
        ]

        if pool_restart_count > DATABASE_MAX_POOL_RESTARTS:

            logging.error(
                "[%s] Worker pool broke %s times in a row -- stopping "
                "early. %s file(s) left unparsed (not marked done).",
                product, pool_restart_count, len(remaining_paths),
            )

            print(
                f"\n    Worker pool broke {pool_restart_count} times in a "
                f"row -- stopping early. {len(remaining_paths):,} file(s) "
                f"left for the next --database run."
            )

            outcome["fatal"] = True

            break

        pool_broke = False

        # max_tasks_per_child (worker recycling) only exists on
        # Python 3.11+. On older Pythons the pool simply runs
        # without it -- streaming inserts already keep worker
        # memory bounded, so this is a nice-to-have, not needed.
        pool_kwargs = {
            "max_workers": worker_count,
            "initializer": _lower_worker_priority,
        }

        if sys.version_info >= (3, 11):
            pool_kwargs["max_tasks_per_child"] = DATABASE_WORKER_MAX_TASKS

        with ProcessPoolExecutor(**pool_kwargs) as executor:

            future_to_path = {
                executor.submit(
                    _parse_file_into_database, file_path, db_path
                ): file_path
                for file_path in remaining_paths
            }

            for future in as_completed(future_to_path):

                file_path = future_to_path.pop(future)

                try:

                    result = future.result()

                except BrokenProcessPool as error:

                    # The whole pool is dead, not just this file.
                    # Files already handled are in completed_paths;
                    # everything else goes to a fresh pool. Any
                    # half-written batch from the dead worker was
                    # rolled back by SQLite automatically.

                    logging.warning(
                        "[%s] Worker pool broke while parsing %s: %s -- "
                        "restarting the pool for the remaining file(s).",
                        product, file_path, error,
                    )

                    pool_broke = True

                    break

                except Exception as error:

                    result = {
                        "path": str(file_path),
                        "size": 0,
                        "mtime": 0.0,
                        "parsed": 0,
                        "inserted": 0,
                        "error": f"{type(error).__name__}: {error}",
                        "error_kind": "parse",
                    }

                now = datetime.now().isoformat(timespec="seconds")

                if result["error_kind"] == "database":

                    # SQLite itself couldn't write (e.g. "disk I/O
                    # error"). Not a per-file problem, and retrying
                    # writes against a failing drive can make
                    # things worse -- stop the run cleanly. The
                    # file is NOT marked done.

                    fatal_db_error = result

                    executor.shutdown(wait=True, cancel_futures=True)

                    break

                if result["error"]:

                    logging.error(
                        "Failed to parse %s: %s", result["path"], result["error"]
                    )

                    error_count += 1

                    status = "error"

                else:

                    status = "success"

                total_parsed += result["parsed"]
                total_inserted += result["inserted"]

                manifest[result["path"]] = {
                    "path": result["path"],
                    "size": str(result["size"]),
                    "mtime": str(result["mtime"]),
                    "records_inserted": str(result["inserted"]),
                    "status": status,
                    "error": result["error"] or "",
                    "processed_time": now,
                }

                completed_paths.add(result["path"])

                progress_bar.update(1)

                results_since_save += 1

                if results_since_save >= DATABASE_MANIFEST_SAVE_INTERVAL:

                    save_database_manifest(manifest)

                    results_since_save = 0

        if pool_broke:
            pool_restart_count += 1
        else:
            pool_restart_count = 0

    progress_bar.close()

    save_database_manifest(manifest)

    if fatal_db_error is not None:

        outcome["fatal"] = True

        logging.error(
            "[%s] Database write failed (%s) while inserting records "
            "from %s. Stopping this run. Run PRAGMA integrity_check on "
            "%s before restarting, and check the drive (chkdsk) if this "
            "recurs.",
            product, fatal_db_error["error"], fatal_db_error["path"], db_path,
        )

        print()
        print(f"  DATABASE WRITE FAILED: {fatal_db_error['error']}")
        print("  This looks like a drive/filesystem problem, not a parse")
        print("  error -- stopping now rather than continuing to write.")
        print("  Check the database's integrity before running again:")
        print(
            f'  python -c "import sqlite3; '
            f"conn = sqlite3.connect(r'{db_path}'); "
            f"print(conn.execute('PRAGMA integrity_check').fetchall())\""
        )

        return outcome

    # Merge the WAL into the main file and shrink it, then build
    # indexes (fast -- they're on the narrow metadata table).

    connection = get_db_connection(db_path)

    connection.execute("PRAGMA wal_checkpoint(TRUNCATE)")

    connection.close()

    create_database_indexes(db_path)

    if backup:

        backup_database(db_path)

    logging.info(
        "[%s] Database build complete: %s records parsed, %s new rows "
        "inserted, %s files done of %s pending (%s skipped, %s errors, "
        "%s pool restart(s)).",
        product, total_parsed, total_inserted, len(completed_paths),
        len(pending_files), skipped_count, error_count, pool_restart_count,
    )

    print()
    print(f"    Records parsed this run   : {total_parsed:,}")
    print(f"    New rows inserted         : {total_inserted:,}")
    print(f"    Files parsed this run     : {len(completed_paths):,}")
    print(f"    Files skipped (unchanged) : {skipped_count:,}")
    print(f"    Files with errors         : {error_count:,}")

    if error_count:
        print(
            "    Errored files are NOT marked done and will be "
            "retried automatically on the next --database run."
        )

    return outcome


# ============================================================
# ONE-TIME MIGRATION: v14 single database -> per-product files
# ============================================================

MIGRATION_COMMIT_ROWS = 2000


def migrate_legacy_database():

    """
    Copies every row of the v14 USPTO_PATENTS.sqlite into the
    matching v15 per-product file, routing each row by its
    source_path (which extracted product folder it came from),
    compressing claims/description on the way. One sequential
    read of the old file -- much cheaper than re-parsing.

    Resumable: progress (the last old-row id copied) is saved to
    migration_v15_state.json after each commit, so an
    interrupted migration continues where it stopped, and
    INSERT OR IGNORE makes any overlap harmless.

    The old file is never modified or deleted. Once this prints
    "complete" and the counts look right, delete
    USPTO_PATENTS.sqlite (and its copy under DATABASE/BACKUP/)
    yourself to reclaim the space.
    """

    print()
    print("=" * 80)
    print("MIGRATING v14 DATABASE -> PER-PRODUCT DATABASES")
    print("=" * 80)

    if not LEGACY_DATABASE_FILE.exists():

        print(f"  No v14 database at {LEGACY_DATABASE_FILE} -- nothing to do.")

        return

    state = _load_migration_state()

    if state.get("complete"):

        print("  Migration already completed earlier -- nothing to do.")

        return

    last_id = int(state.get("last_id", 0))

    counts = state.get("counts", {product: 0 for product in DATABASE_FILES})

    unmatched = int(state.get("unmatched", 0))

    prefixes = [
        (str(_product_extracted_dir(product)) + os.sep, product)
        for product in DATABASE_PRODUCT_ORDER
    ]

    targets = {}

    for product, db_path in DATABASE_FILES.items():

        create_database(db_path)

        targets[product] = get_db_connection(db_path)

    old = sqlite3.connect(LEGACY_DATABASE_FILE)

    old.execute("PRAGMA cache_size = -64000")

    total_rows = old.execute("SELECT MAX(id) FROM patents").fetchone()[0] or 0

    print(f"  Resuming after old row id {last_id:,} of ~{total_rows:,}")

    cursor = old.execute(
        """
        SELECT id, patent_number, publication_number, document_type,
               filing_date, publication_date, grant_date, priority_date,
               cpc, ipc, title, inventors, applicants, assignees,
               source_file, source_path, abstract, claims, description
        FROM patents
        WHERE id > ?
        ORDER BY id
        """,
        (last_id,),
    )

    pending = {product: [] for product in DATABASE_FILES}

    rows_since_commit = 0

    unmatched_examples = []

    def flush():

        for product, rows in pending.items():

            if not rows:
                continue

            connection = targets[product]

            connection.execute("BEGIN IMMEDIATE")

            for meta_row, abstract, claims, description in rows:

                cur = connection.execute(PATENTS_INSERT_SQL, meta_row)

                if cur.rowcount == 1:

                    connection.execute(
                        PATENT_TEXT_INSERT_SQL,
                        (cur.lastrowid, abstract, claims, description),
                    )

                    counts[product] = counts.get(product, 0) + 1

            connection.commit()

            rows.clear()

    progress_bar = tqdm(total=max(total_rows - last_id, 0), desc="Migrating")

    try:

        for row in cursor:

            old_id = row[0]

            source_path = row[15] or ""

            product = None

            for prefix, candidate in prefixes:

                if source_path.startswith(prefix):
                    product = candidate
                    break

            if product is None:

                unmatched += 1

                if len(unmatched_examples) < 5:
                    unmatched_examples.append(source_path)

            else:

                pending[product].append(
                    (
                        row[1:16],
                        row[16],
                        _compress_text(row[17]),
                        _compress_text(row[18]),
                    )
                )

            rows_since_commit += 1

            last_id = old_id

            if rows_since_commit >= MIGRATION_COMMIT_ROWS:

                flush()

                _save_migration_state(
                    {"last_id": last_id, "counts": counts, "unmatched": unmatched}
                )

                progress_bar.update(rows_since_commit)

                rows_since_commit = 0

        flush()

        progress_bar.update(rows_since_commit)

    finally:

        progress_bar.close()

        old.close()

    for product, connection in targets.items():

        connection.execute("PRAGMA wal_checkpoint(TRUNCATE)")

        connection.close()

        create_database_indexes(DATABASE_FILES[product])

    _save_migration_state(
        {"last_id": last_id, "counts": counts, "unmatched": unmatched,
         "complete": True}
    )

    print()
    print("  Migration complete. Rows copied per product:")

    for product in DATABASE_PRODUCT_ORDER:
        print(f"    {product:<7} {counts.get(product, 0):>12,}")

    if unmatched:

        print(
            f"  WARNING: {unmatched:,} row(s) had a source_path outside "
            f"all four product folders and were NOT copied, e.g.:"
        )

        for example in unmatched_examples:
            print(f"    {example}")

        print(
            "  Their files' manifest rows still say 'success'; remove "
            "those rows from database_manifest.csv so --database "
            "re-parses them."
        )

    print()
    print("  Once you're happy with the counts, delete the old file(s)")
    print(f"  to reclaim space: {LEGACY_DATABASE_FILE}")
    print(f"  and {DATABASE_BACKUP_DIR / LEGACY_DATABASE_FILE.name} (if present).")


# ============================================================
# ONE-TIME REPAIR: re-queue files hit by the glued-declaration bug
# ============================================================

def mark_merged_chunk_files_for_reparse():

    """
    Finds every XML file that pipeline.log shows was affected
    by the glued-declaration bug (see _split_glued_declarations())
    and marks it "error" in database_manifest.csv, so
    --database --retry-errors-only re-parses just those files.

    Re-parsing is safe: INSERT OR IGNORE skips every record
    already in the database, so only the patents that were lost
    get added. Nothing in any database is changed by this step.

    Looks for both log messages this bug produced:
      "Recovered the FIRST document from a merged/junk-after-
       document chunk in <file>"
      "Skipping malformed XML record in <file>: ... (also failed
       after attempting merged-chunk recovery)"
    """

    print()
    print("=" * 80)
    print("RE-QUEUEING FILES AFFECTED BY THE MERGED-RECORD BUG")
    print("=" * 80)

    log_path = LOG_DIR / "pipeline.log"

    if not log_path.exists():

        print(f"  No log found at {log_path}.")

        return

    patterns = [
        re.compile(r"merged/junk-after-document chunk in (\S+?) \(error"),
        re.compile(r"Skipping malformed XML record in (\S+?): .*merged-chunk recovery"),
    ]

    affected_names = {}

    with open(log_path, "r", encoding="utf-8", errors="replace") as handle:

        for line in handle:

            if "merged" not in line:
                continue

            for pattern in patterns:

                match = pattern.search(line)

                if match:

                    name = match.group(1)

                    affected_names[name] = affected_names.get(name, 0) + 1

                    break

    if not affected_names:

        print("  No affected files found in the log -- nothing to do.")

        return

    manifest = load_database_manifest()

    marked = {product: [] for product in DATABASE_PRODUCT_ORDER}

    not_in_manifest = set(affected_names)

    for path_text, row in manifest.items():

        name = Path(path_text).name

        if name not in affected_names:
            continue

        not_in_manifest.discard(name)

        row["status"] = "error"

        row["error"] = "re-queued: merged-record bug fixed in v15.1"

        product = next(
            (
                candidate
                for candidate in DATABASE_PRODUCT_ORDER
                if str(path_text).startswith(
                    str(_product_extracted_dir(candidate)) + os.sep
                )
            ),
            "other",
        )

        marked.setdefault(product, []).append(name)

    save_database_manifest(manifest)

    total = sum(len(names) for names in marked.values())

    print(f"  Files re-queued for parsing: {total:,}")

    for product, names in marked.items():

        if names:

            warnings = sum(affected_names[name] for name in names)

            print(
                f"    {product:<7} {len(names):>5,} file(s)  "
                f"({warnings:,} logged warning(s) -- roughly one lost "
                f"patent each)"
            )

    if not_in_manifest:

        print(
            f"  {len(not_in_manifest):,} affected file name(s) from the log "
            f"aren't in the manifest (e.g. not parsed yet) -- they'll be "
            f"parsed normally when their turn comes."
        )

    print()
    print("  Next, for each product listed above, run for example:")
    print("    python uspto_patent_pipeline_v15.py --database --format xml --retry-errors-only")
    print("    python uspto_patent_pipeline_v15.py --database --format appxml --retry-errors-only")


# ============================================================
# CPC FILTER
# ============================================================

# ------------------------------------------------------------
# !!! YOU ENTER THIS !!!
# ------------------------------------------------------------
# Patents granted before CPC existed (roughly pre-2013) have
# NO CPC classification at all -- only IPC. If this is True,
# a record with no CPC match falls back to checking its IPC
# codes against the same CPC_CODES prefixes, instead of
# automatically excluding every legacy/historical patent just
# because CPC didn't exist yet when it was granted. CPC was
# built as an extension of IPC, so the same class/subclass/
# main-group prefix (e.g. "B29C48") is generally still a
# meaningful match against an IPC symbol from before CPC
# existed.
#
# Set this to False if you'd rather restrict results strictly
# to patents that actually carry real CPC classifications.
# ------------------------------------------------------------

MATCH_IPC_AS_CPC_FALLBACK = True


def matches_cpc(

    cpc_text,

    ipc_text=None,

):

    """
    Patents commonly carry multiple CPC classifications. If
    cpc_text is a delimited list (e.g. "B29C48/08; H01L21/02"),
    checking cpc_text.startswith(code) against the whole
    string would only ever match when the RELEVANT code
    happens to be first -- silently missing matches. Instead,
    split into individual codes and check each one.

    If cpc_text has no match (including the common case of
    cpc_text being empty entirely, as with ALL pre-2013
    legacy/APS records) and MATCH_IPC_AS_CPC_FALLBACK is True,
    also checks ipc_text against the same CPC_CODES prefixes.
    """

    if _classification_text_matches_cpc_codes(cpc_text):

        return True


    if (

        MATCH_IPC_AS_CPC_FALLBACK

        and _classification_text_matches_cpc_codes(ipc_text)

    ):

        return True


    return False


def _classification_text_matches_cpc_codes(

    classification_text

):

    if not classification_text:

        return False


    individual_codes = (

        re.sub(r"\s+", "", code.strip().upper())

        for code in CPC_CODE_SPLIT_PATTERN.split(
            classification_text
        )

        if code.strip()

    )


    for code in individual_codes:

        if code.startswith(CPC_CODES_UPPER):

            return True


    return False


# ============================================================
# KEYWORD FILTER
# ============================================================

def matches_keywords(
    title,
    abstract,
    claims,
    description,
):

    """
    True if any keyword appears in any of the four fields.

    v15 changes (same results, much less work per patent):
      - One combined regex (KEYWORD_COMBINED_PATTERN) scans each
        field ONCE, instead of one full pass over the text per
        keyword.
      - Fields are checked cheapest-first (title, abstract,
        claims, description) and stop at the first hit, so the
        large description is often never decompressed or
        scanned at all.
      - claims/description may arrive as compressed BLOBs from
        the database; they're decompressed only if reached.
    """

    if KEYWORD_COMBINED_PATTERN is None:
        return False

    for field in (title, abstract, claims, description):

        text = _decompress_text(field)

        if text and KEYWORD_COMBINED_PATTERN.search(text):
            return True

    return False


# ============================================================
# FILTER DATABASE
# ============================================================

FILTER_OUTPUT_FIELDS = [
    "patent_number",
    "publication_number",
    "document_type",
    "title",
    "inventors",
    "applicants",
    "assignees",
    "filing_date",
    "publication_date",
    "grant_date",
    "priority_date",
    "cpc",
    "ipc",
    "source_path",
]


def _product_filter_results_path(product):

    """RESULTS/filtered_patents_<product>.csv -- one per product."""

    return FILTER_RESULTS.with_name(
        f"{FILTER_RESULTS.stem}_{product}{FILTER_RESULTS.suffix}"
    )


def _filter_one_product(product, db_path):

    """
    Filters one product database into its own CSV
    (filtered_patents_<product>.csv). Writes to a .tmp file and
    renames it into place at the end, so an interrupted run never
    leaves a half-written result behind.

    Two-phase, as before: CPC/IPC gate on the narrow `patents`
    table first; full text is fetched only for CPC matches and
    then checked against the keywords.

    Returns the number of matches.
    """

    output_path = _product_filter_results_path(product)

    temp_path = output_path.with_suffix(".tmp")

    connection = get_db_connection(db_path)

    total_patents = connection.execute(
        "SELECT COUNT(*) FROM patents"
    ).fetchone()[0]

    meta_cursor = connection.execute("""
        SELECT id, patent_number, publication_number, document_type,
               title, inventors, applicants, assignees,
               filing_date, publication_date, grant_date,
               priority_date, cpc, ipc, source_path
        FROM patents
        ORDER BY id
    """)

    text_cursor = connection.cursor()

    match_count = 0

    try:

        with open(temp_path, "w", newline="", encoding="utf-8-sig") as output_file:

            writer = csv.DictWriter(output_file, fieldnames=FILTER_OUTPUT_FIELDS)

            writer.writeheader()

            for row in tqdm(meta_cursor, total=total_patents, desc=f"Filtering {product}"):

                (
                    patent_id, patent_number, publication_number,
                    document_type, title, inventors, applicants,
                    assignees, filing_date, publication_date,
                    grant_date, priority_date, cpc, ipc, source_path,
                ) = row

                if not matches_cpc(cpc, ipc):
                    continue

                text_row = text_cursor.execute(
                    "SELECT abstract, claims, description "
                    "FROM patent_text WHERE id = ?",
                    (patent_id,),
                ).fetchone()

                abstract, claims, description = text_row or (None, None, None)

                if not matches_keywords(title, abstract, claims, description):
                    continue

                writer.writerow({
                    "patent_number": patent_number,
                    "publication_number": publication_number,
                    "document_type": document_type,
                    "title": title,
                    "inventors": inventors,
                    "applicants": applicants,
                    "assignees": assignees,
                    "filing_date": filing_date,
                    "publication_date": publication_date,
                    "grant_date": grant_date,
                    "priority_date": priority_date,
                    "cpc": cpc,
                    "ipc": ipc,
                    "source_path": source_path,
                })

                match_count += 1

    finally:

        connection.close()

    os.replace(temp_path, output_path)

    return match_count


def _combine_filter_results():

    """
    Rebuilds RESULTS/filtered_patents.csv (what --pdf and the
    forward-citations script read) by concatenating every
    per-product result file that exists, in APS -> SGML -> XML
    -> APPXML order. Returns (rows written, products included,
    products not yet filtered).
    """

    included = []
    missing = []
    total = 0

    temp_path = FILTER_RESULTS.with_suffix(".tmp")

    with open(temp_path, "w", newline="", encoding="utf-8-sig") as output_file:

        writer = csv.DictWriter(output_file, fieldnames=FILTER_OUTPUT_FIELDS)

        writer.writeheader()

        for product in DATABASE_PRODUCT_ORDER:

            product_path = _product_filter_results_path(product)

            if not product_path.exists():

                if DATABASE_FILES[product].exists():
                    missing.append(product)

                continue

            with open(product_path, "r", newline="", encoding="utf-8-sig") as input_file:

                for row in csv.DictReader(input_file):

                    writer.writerow({field: row.get(field, "") for field in FILTER_OUTPUT_FIELDS})

                    total += 1

            included.append(product)

    os.replace(temp_path, FILTER_RESULTS)

    return total, included, missing


def filter_patents(product_format=None):

    """
    Applies the CPC + keyword filter.

    --filter                 filters all four product databases
    --filter --format aps    filters just one (sgml / xml / appxml
                             likewise), so the filter can be run
                             in sections, like --database

    Each product's matches go to filtered_patents_<product>.csv.
    After every run, filtered_patents.csv is rebuilt from all the
    per-product files that exist, so it always holds every
    section filtered so far.

    NOTE: each section uses the CPC codes / keywords loaded at the
    time it ran. If the workbooks change between sections, re-run
    the earlier sections too so every product uses the same
    criteria.
    """

    print()
    print("=" * 80)
    print("FILTERING PATENTS")
    print("=" * 80)

    if _legacy_database_needs_migration():

        _print_migration_required()

        return

    if product_format is not None and product_format not in DATABASE_FILES:

        print(
            f"  Unknown --format '{product_format}'. Choose from: "
            f"{', '.join(DATABASE_FILES)}."
        )

        return

    products = [product_format] if product_format else DATABASE_PRODUCT_ORDER

    filtered_any = False

    for product in products:

        db_path = DATABASE_FILES[product]

        if not db_path.exists():

            print(f"  {product}: no database yet ({db_path.name}) -- skipping.")

            continue

        matches = _filter_one_product(product, db_path)

        filtered_any = True

        print(f"  {product}: {matches:,} matches  ->  {_product_filter_results_path(product).name}")

    if not filtered_any:

        print("  Nothing was filtered -- run --database first.")

        return

    total, included, missing = _combine_filter_results()

    print()
    print(f"Combined results ({', '.join(included)}): {total:,} patents")
    print(FILTER_RESULTS)

    if missing:

        print(
            f"  Not filtered yet: {', '.join(missing)} -- run "
            f"--filter --format <product> for each to add them."
        )


# ============================================================
# PDF URL
# ============================================================

# ------------------------------------------------------------
# PPUBS RATE LIMITER
# ------------------------------------------------------------
# A shared lock + last-request timestamp, so the minimum
# interval is enforced across the AGGREGATE request rate from
# all PDF_DOWNLOAD_WORKERS threads combined -- not per-thread
# (which wouldn't meaningfully throttle anything, since workers
# would still all fire independently at the same time).

_ppubs_rate_limit_lock = threading.Lock()

_ppubs_last_request_time = [0.0]


def _wait_for_ppubs_rate_limit():

    with _ppubs_rate_limit_lock:

        now = time.monotonic()

        elapsed = now - _ppubs_last_request_time[0]

        wait_time = (
            PPUBS_MIN_REQUEST_INTERVAL_SECONDS - elapsed
        )

        if wait_time > 0:

            time.sleep(wait_time)

        _ppubs_last_request_time[0] = time.monotonic()


def get_pdf_url(
    publication_number,
    legacy_aps=False,
):

    """
    Returns a direct PDF download URL for a patent/publication
    number, using USPTO's public Patent Public Search (PPUBS)
    image server -- NOT the api.uspto.gov ODP endpoints used
    elsewhere in this file.

    VERIFIED (not guessed): fetched a real example of this URL
    directly and confirmed it returns actual patent PDF content,
    with no API key or authentication required at all:

        https://image-ppubs.uspto.gov/dirsearch-public/
            print/downloadPdf/10848629

    Format rules, from USPTO's own "PDF Links QRG" (2024):

        7-digit utility patent   -> zero-padded to 7 digits
        8-digit utility patent   -> as-is (patents > 10,000,000)
        design patent            -> "D" + digits
        plant patent              -> "PP" + 5-digit
        published application    -> "yyyy" + 7-digit (11 total)

    IMPORTANT ADVANTAGE over the api.uspto.gov application-
    documents route: this is driven purely by patent number
    format, not filing date, so unlike that route (which USPTO
    documents as covering only applications filed after
    January 1, 2001), this should also work for pre-2001
    legacy (APS/SGML) records -- i.e. it may cover the WHOLE
    corpus, not just the modern half. That's not yet confirmed
    against an actual pre-2001 patent number, only inferred
    from how the format rules are described.

    KNOWN GAPS (not yet handled, return None):
      - Reissue patents ("RE" prefix) -- no documented format
      - Statutory Invention Registrations ("H" prefix) -- no
        documented format
      - Documents over 1000 pages don't work via this endpoint
        at all (per USPTO's own documentation) -- this isn't
        something we can detect in advance, so a failure here
        should be treated as permanent and moved past, not
        retried.
    """

    normalized = _normalize_number_for_ppubs(
        publication_number,
        legacy_aps=legacy_aps,
    )

    if not normalized:

        return None

    return (

        "https://image-ppubs.uspto.gov/dirsearch-public/"
        f"print/downloadPdf/{normalized}"

    )


def _normalize_number_for_ppubs(
    publication_number,
    legacy_aps=False,
):

    """
    Converts our internally-stored patent/publication number
    (e.g. "USD0707016S1", "US06167569A", "US20140123360A1",
    "USPP011728P") into the digit-string format PPUBS expects
    (e.g. "D0707016", "6167569", "20140123360", "PP011728").

    Reissue (RE) and defensive publication (T) numbers are
    supported; SIR (H) returns None (no working format found).
    """

    if not publication_number:

        return None

    value = publication_number.strip().upper()

    if value.startswith("US"):

        value = value[2:]



    # APS (Green Book) records store the WKU as their number:
    # 9 characters, the LAST of which is a check character, not
    # part of the patent number. Usually a digit ("047948337" is
    # patent 4,794,833 + check 7), but in real APS files it is
    # sometimes "&" (e.g. "WKU  06333061&" in
    # pftaps20011225_wk52 -- 830 such rows in the filter
    # results, all APS). Passing it through asked PPUBS for a
    # nonexistent number (HTTP 404) or was rejected outright.
    # (The first patent of 1976 is 3,930,271, whose WKU is
    # 039302711; PatentsView also converts WKUs by dropping the
    # final character.) Dropping it leaves prefix + zero-padded
    # number, which the rules below already handle
    # ("04794833" -> "4794833", "D0238000" -> "D0238000",
    # "PP003564" -> "PP003564"). Only applied to APS-sourced
    # records -- SGML/XML numbers have no check character.

    if legacy_aps and len(value) == 9:
        value = value[:-1]

    # Real patent/publication numbers only contain letters and
    # digits; drop any other stray character as a safety net so
    # it can't make an otherwise valid number "unsupported".

    value = re.sub(r"[^A-Z0-9]", "", value)


    # Published application: 11 digits, first 4 being a
    # plausible year, optionally followed by a kind code
    # (e.g. "20140123360A1" -> "20140123360").

    application_pub_match = re.match(

        r"^(\d{4}\d{7})[A-Z]?\d*$",
        value,

    )

    if application_pub_match:

        year = int(application_pub_match.group(1)[:4])

        if 1900 <= year <= 2100:

            return application_pub_match.group(1)


    # Design patent: "D" + digits, optionally followed by a
    # kind code (e.g. "D0707016S1" -> "D0707016").

    design_match = re.match(

        r"^D(\d+)[A-Z]?\d*$",
        value,

    )

    if design_match:

        return f"D{design_match.group(1)}"


    # Plant patent: "PP" + digits, optionally followed by a
    # kind code (e.g. "PP011728P" -> "PP011728").

    plant_match = re.match(

        r"^PP(\d+)[A-Z]?\d*$",
        value,

    )

    if plant_match:

        return "PP" + plant_match.group(1).zfill(5)


    # Reissue and SIR: no documented format -- don't guess.

    # Reissue and defensive publication: prefix + the number
    # with NO zero padding. Confirmed in a browser (Oct 2026):
    # downloadPdf/RE38552 and downloadPdf/T105202 work;
    # RE038552 and T0105202 do not. Stored forms handled:
    # XML/SGML "RE038552E1", APS "RE0374423" (check char
    # already dropped above -> "RE037442"), APS "T01052020"
    # (-> "T0105202").

    reissue_match = re.match(r"^RE(\d+)[A-Z]?\d*$", value)

    if reissue_match:
        return "RE" + (reissue_match.group(1).lstrip("0") or "0")

    defensive_match = re.match(r"^T(\d+)[A-Z]?\d*$", value)

    if defensive_match:
        return "T" + (defensive_match.group(1).lstrip("0") or "0")

    # SIR ("H") numbers: no working link format found -- both
    # downloadPdf/H2102 and downloadPdf/H0002102 fail -- so they
    # stay unsupported and need a manual lookup.

    if value.startswith("H"):

        return None


    # Plain utility patent: digits only, optionally followed
    # by a kind code (e.g. "06167569A" -> "6167569",
    # "10123456B2" -> "10123456").

    utility_match = re.match(

        r"^(\d+)[A-Z]?\d*$",
        value,

    )

    if utility_match:

        digits = utility_match.group(1).lstrip("0") or "0"

        if len(digits) <= 7:

            return digits.zfill(7)

        return digits


    return None


# ============================================================
# PDF MANIFEST
# ============================================================

PDF_MANIFEST_FIELDS = [

    "publication_number",
    "local_path",
    "status",
    "download_time",
    "error",

]


def load_pdf_manifest():

    """
    Same pattern as the bulk-download manifest: lets --pdf be
    restarted at 500,000+ file scale without re-attempting
    publications that already succeeded, and without silently
    re-hammering ones that are permanently failing (e.g. no
    PDF available).
    """

    return _load_csv_manifest(

        PDF_MANIFEST,

        build_key=lambda row: row["publication_number"],

    )


def save_pdf_manifest(
    records
):

    _save_csv_manifest(

        PDF_MANIFEST,
        records,
        PDF_MANIFEST_FIELDS,

    )


# ============================================================
# DOWNLOAD PDF
# ============================================================

def _pdf_subdirectory_for_document_type(

    document_type

):

    """
    Routes a patent's PDF into PDF_GRANTS_DIR,
    PDF_APPLICATIONS_DIR, or PDF_OTHER_DIR based on its
    document_type ("grant"/"application"/"sir", set by the
    parsers -- see parse_xml_file()/parse_sgml_file()/
    parse_legacy_file()). Anything not cleanly "grant" or
    "application" (e.g. "sir", or a missing/unexpected value)
    goes to PDF_OTHER_DIR rather than being silently forced
    into the wrong bucket.
    """

    normalized = (

        (document_type or "").strip().lower()

    )

    if normalized == "grant":

        return PDF_GRANTS_DIR

    if normalized == "application":

        return PDF_APPLICATIONS_DIR

    return PDF_OTHER_DIR


def download_pdf(

    publication_number,

    document_type=None,

    legacy_aps=False,

):

    """
    Downloads a single patent PDF.

    Saved under PDF_GRANTS_DIR/PDF_APPLICATIONS_DIR/
    PDF_OTHER_DIR depending on document_type, so grants and
    applications never end up mixed in one flat folder.

    Returns a record dict. This does NOT write to the shared
    manifest directly -- it's called from multiple threads at
    once, so the caller writes the result under a lock.
    """

    def build_record(

        status,
        local_path="",
        download_time="",
        error="",

    ):

        return {

            "publication_number": publication_number,
            "local_path": local_path,
            "status": status,
            "download_time": download_time,
            "error": error,

        }


    url = get_pdf_url(
        publication_number,
        legacy_aps=legacy_aps,
    )


    if not url:

        return build_record(

            "FAILED",

            error=(
                "No PDF URL could be constructed for this "
                "publication number (unsupported format -- "
                "e.g. reissue or SIR patents aren't supported "
                "by get_pdf_url() yet)"
            ),

        )


    safe_name = re.sub(

        r"[^A-Za-z0-9_.-]",

        "_",

        publication_number

    )


    destination_dir = _pdf_subdirectory_for_document_type(
        document_type
    )

    destination = (

        destination_dir /

        f"{safe_name}.pdf"

    )


    if destination.exists():

        return build_record(
            "COMPLETE",
            local_path=str(destination),
        )


    partial = Path(
        str(destination) + ".part"
    )


    last_error = None

    for attempt in range(

        1,

        MAX_RETRIES + 1

    ):

        try:

            # Self-imposed pacing for the undocumented PPUBS
            # endpoint -- see PPUBS_MIN_REQUEST_INTERVAL_SECONDS.
            # Applied on every attempt, including retries.

            _wait_for_ppubs_rate_limit()

            response = SESSION.get(

                url,

                stream=True,

                timeout=REQUEST_TIMEOUT,

                # The PDF server is a public endpoint that doesn't
                # use API keys; don't send the key to it. (None
                # removes the session-wide header for this call.)
                headers={"X-API-KEY": None},

            )


            # ----------------------------------------------------
            # PERMANENT FAILURE - stop immediately, don't retry
            # ----------------------------------------------------

            if response.status_code in PERMANENT_FAILURE_STATUS_CODES:

                error_message = (
                    f"Permanent failure HTTP "
                    f"{response.status_code}"
                )

                logging.error(

                    "PERMANENT FAILURE PDF %s: %s",
                    publication_number,
                    error_message,

                )

                return build_record(
                    "FAILED",
                    error=error_message,
                )


            if response.status_code != 200:

                raise RuntimeError(

                    f"HTTP "
                    f"{response.status_code}"

                )


            with open(

                partial,

                "wb"

            ) as file:

                for chunk in response.iter_content(

                    chunk_size=CHUNK_SIZE

                ):

                    if chunk:

                        file.write(
                            chunk
                        )


            # ----------------------------------------------------
            # VALIDATE: must actually be a PDF.
            #
            # USPTO (like many services) can return HTTP 200 with
            # an HTML error/login page instead of the document.
            # Without this check that page gets silently saved
            # and counted as a successful download.
            # ----------------------------------------------------

            with open(

                partial,

                "rb"

            ) as file:

                header = file.read(
                    len(PDF_MAGIC_BYTES)
                )


            if header != PDF_MAGIC_BYTES:

                raise RuntimeError(

                    "Downloaded file is not a valid PDF "
                    "(missing %PDF- header)"

                )


            partial.replace(
                destination
            )


            return build_record(

                "COMPLETE",

                local_path=str(destination),

                download_time=datetime.now().isoformat(),

            )


        except Exception as error:

            last_error = str(error)

            logging.warning(

                "PDF download failed "
                "%s (attempt %d/%d): %s",

                publication_number,
                attempt,
                MAX_RETRIES,
                error

            )


            if attempt < MAX_RETRIES:

                time.sleep(
                    RETRY_DELAY_SECONDS * attempt
                )


    partial.unlink(missing_ok=True)

    return build_record(
        "FAILED",
        error=str(last_error),
    )


# ============================================================
# DOWNLOAD SELECTED PDFs
# ============================================================

def _print_pdf_progress_summary(manifest):

    """
    Before any download starts, shows how far along each section
    is: PDFs complete / patents selected, per section (from each
    filtered_patents_<product>.csv) plus a total across all
    sections, with the FAILED count (those wait for
    --retry-errors-only).
    """

    labels = {
        "aps": "APS",
        "sgml": "SGML",
        "xml": "XML grants",
        "appxml": "Applications",
    }

    print()
    print("PDF PROGRESS BY SECTION")
    print(f"  {'Section':<18}{'Complete':>12}   {'Total':>10}{'%':>9}{'Failed':>10}")

    sum_complete = 0
    sum_total = 0
    sum_failed = 0

    for product in DATABASE_PRODUCT_ORDER:

        path = _product_filter_results_path(product)

        if not path.exists():

            print(f"  {labels[product]:<18}{'(not filtered yet)':>34}")

            continue

        numbers = set()

        with open(path, "r", encoding="utf-8-sig", newline="") as handle:

            for row in csv.DictReader(handle):

                number = row.get("publication_number")

                if number:
                    numbers.add(number)

        complete = 0
        failed = 0

        for number in numbers:

            status = (manifest.get(number) or {}).get("status")

            if status == "COMPLETE":
                complete += 1

            elif status == "FAILED":
                failed += 1

        total = len(numbers)

        percent = (100.0 * complete / total) if total else 0.0

        print(
            f"  {labels[product]:<18}{complete:>12,} / {total:>10,}"
            f"{percent:>8.1f}%{failed:>10,}"
        )

        sum_complete += complete
        sum_total += total
        sum_failed += failed

    percent = (100.0 * sum_complete / sum_total) if sum_total else 0.0

    print(f"  {'-' * 59}")
    print(
        f"  {'ALL SECTIONS':<18}{sum_complete:>12,} / {sum_total:>10,}"
        f"{percent:>8.1f}%{sum_failed:>10,}"
    )


def _is_aps_row(row):

    """
    True if a filter-results row came from an APS (1976-2000,
    Green Book) file, whose stored number carries a WKU check
    digit (see _normalize_number_for_ppubs()). Decided by the
    row's source_path -- every filter row records which
    extracted folder it was parsed from.
    """

    source_path = row.get("source_path") or ""

    aps_root = str(_product_extracted_dir("aps")) + os.sep

    return source_path.startswith(aps_root)


def download_selected_pdfs(
    product_format=None,
    limit=None,
    skip_errors=False,
    retry_errors_only=False,
):

    """
    Downloads PDFs for the filter results.

    --pdf                        everything in filtered_patents.csv
    --pdf --format aps           only filtered_patents_aps.csv
                                 (sgml / xml / appxml likewise)
    --pdf --limit 500            at most 500 downloads this run
    --pdf --retry-errors-only    retry ONLY previously FAILED ones

    Previously FAILED patents are always skipped by a normal
    --pdf run (v15.5), so no --limit batch is spent on them.
    Retry them all at the end with --retry-errors-only.
    (--skip-errors is accepted but no longer needed here.)

    All of these combine. The PDF manifest is shared across
    sections, so a patent already COMPLETE is skipped whichever
    section it's requested from.
    """

    if not PDF_DOWNLOAD_ENABLED:

        print(
            "PDF downloading disabled."
        )

        return

    if skip_errors and retry_errors_only:

        print(
            "  --skip-errors and --retry-errors-only can't be used "
            "together (that would skip everything). Use one or "
            "the other."
        )

        return

    if product_format is not None:

        if product_format not in DATABASE_FILES:

            print(
                f"  Unknown --format '{product_format}'. Choose from: "
                f"{', '.join(DATABASE_FILES)}."
            )

            return

        source_csv = _product_filter_results_path(product_format)

        if not source_csv.exists():

            print(
                f"  {source_csv.name} does not exist -- run "
                f"--filter --format {product_format} first."
            )

            return

    else:

        source_csv = FILTER_RESULTS

        if not source_csv.exists():

            raise RuntimeError(

                "filtered_patents.csv "
                "does not exist."

            )


    with open(

        source_csv,

        "r",

        encoding="utf-8-sig"

    ) as file:

        reader = csv.DictReader(
            file
        )

        rows = list(reader)


    print()

    print("=" * 80)

    print(
        "DOWNLOADING SELECTED PATENT PDFs"
        + (f"  [{product_format}]" if product_format else "")
    )

    print("=" * 80)

    print(f"  Source: {source_csv.name}")


    manifest = load_pdf_manifest()

    manifest_lock = threading.Lock()

    _print_pdf_progress_summary(manifest)


    # Skip anything the manifest already marks COMPLETE --
    # at 500,000+ files, avoiding even an existence-check
    # round trip for known-done PDFs matters.

    pending = []

    already_complete = 0

    previously_failed = 0

    seen = set()

    for row in rows:

        publication_number = row.get(
            "publication_number"
        )

        if not publication_number or publication_number in seen:

            continue

        seen.add(publication_number)

        existing = manifest.get(
            publication_number
        )

        status = existing.get("status") if existing else None

        if status == "COMPLETE":

            already_complete += 1

            continue

        if status == "FAILED":

            previously_failed += 1

            if not retry_errors_only:
                continue

        elif retry_errors_only:

            continue

        pending.append(

            (

                publication_number,
                row.get("document_type"),
                _is_aps_row(row),

            )

        )

    total_pending = len(pending)

    if limit is not None and limit < total_pending:

        pending = pending[:limit]


    print()

    print(f"Total selected patents: {len(seen):,}")

    print(f"Already downloaded (per manifest): {already_complete:,}")

    if previously_failed:

        if retry_errors_only:

            print(f"Previously failed (retrying now): {previously_failed:,}")

        else:

            print(
                f"Previously failed (skipped -- retry at the end with "
                f"--pdf --retry-errors-only): {previously_failed:,}"
            )

    print(f"Remaining to download: {total_pending:,}")

    if len(pending) < total_pending:

        print(f"--limit: downloading {len(pending):,} this run")

    print(f"Using {PDF_DOWNLOAD_WORKERS} parallel workers.")

    print()

    if not pending:

        print("Nothing to download.")

        return



    completed_since_save = 0

    successes = 0

    failures = 0

    with ThreadPoolExecutor(

        max_workers=PDF_DOWNLOAD_WORKERS

    ) as executor:

        future_to_pub = {

            executor.submit(
                download_pdf,
                publication_number,
                document_type,
                legacy_aps,
            ): publication_number

            for publication_number, document_type, legacy_aps in pending

        }


        for future in tqdm(

            as_completed(future_to_pub),

            total=len(pending),

            desc="PDFs"

        ):

            publication_number = future_to_pub[future]

            try:

                record = future.result()

            except Exception as error:

                logging.error(

                    "Unhandled error downloading "
                    "PDF %s: %s",
                    publication_number,
                    error,

                )

                record = {

                    "publication_number": publication_number,
                    "local_path": "",
                    "status": "FAILED",
                    "download_time": "",
                    "error": str(error),

                }


            if record.get("status") == "COMPLETE":

                successes += 1

            else:

                failures += 1


            with manifest_lock:

                manifest[publication_number] = record

                completed_since_save += 1

                if completed_since_save >= MANIFEST_SAVE_INTERVAL:

                    save_pdf_manifest(
                        manifest
                    )

                    completed_since_save = 0


    with manifest_lock:

        save_pdf_manifest(
            manifest
        )


    print()

    print(

        f"PDF downloads complete. "
        f"{successes:,} succeeded, "
        f"{failures:,} failed."

    )

    if failures:

        print(

            f"See {PDF_MANIFEST} for details on failed "
            f"PDFs. They're skipped by normal --pdf runs; "
            f"retry them at the end with --pdf --retry-errors-only."

        )


# ============================================================
# TEST USPTO CONNECTION
# ============================================================

def test_connection():

    print()

    print(
        "Testing USPTO API connection..."
    )


    data = search_products(
        "Patent Grant Full-Text"
    )


    products = display_products(
        data
    )


    if products:

        print(
            "\nUSPTO API connection successful."
        )

    else:

        print(
            "\nNo products returned."
        )


# ============================================================
# DOWNLOAD EVERYTHING
# ============================================================

def download_everything(sample=False):

    manifest = (
        load_download_manifest()
    )

    manifest_lock = threading.Lock()


    # --------------------------------------------------------
    # GRANTS - three separate products, exact identifiers.
    #
    # Fuzzy title matching (searching for "Patent Grant
    # Full-Text" and taking whichever result matches first)
    # is ambiguous: APS/SGML/XML product titles all contain
    # that phrase, so it could silently grab the wrong one.
    # Using known identifiers avoids that entirely.
    # --------------------------------------------------------

    download_product(

        product_type="GRANTS_APS",

        product_identifier="PTGRAPS",

        destination_directory=
            GRANTS_APS_DIR,

        manifest=
            manifest,

        manifest_lock=
            manifest_lock,

        sample=
            sample,

    )

    download_product(

        product_type="GRANTS_SGML",

        product_identifier="PTGRSGM",

        destination_directory=
            GRANTS_SGML_DIR,

        manifest=
            manifest,

        manifest_lock=
            manifest_lock,

        sample=
            sample,

    )

    download_product(

        product_type="GRANTS_XML",

        product_identifier="PTGRXML",

        destination_directory=
            GRANTS_XML_DIR,

        manifest=
            manifest,

        manifest_lock=
            manifest_lock,

        sample=
            sample,

    )


    # --------------------------------------------------------
    # APPLICATIONS - one product (XML only; USPTO only
    # publishes applications from 2001+, XML era only).
    # --------------------------------------------------------

    download_product(

        product_type="APPLICATIONS_XML",

        product_identifier="APPXML",

        destination_directory=
            APPLICATIONS_DIR,

        manifest=
            manifest,

        manifest_lock=
            manifest_lock,

        sample=
            sample,

    )


# ============================================================
# DOWNLOAD ONE PRODUCT AT A TIME
# ============================================================
#
# download_everything() runs all four products back-to-back in
# one call, in a fixed order (APS -> SGML -> XML grants ->
# APPXML). That's restartable at the individual-FILE level via
# the manifest -- an interrupted run resumes without redoing
# completed files -- but it doesn't give a clean "this product
# is fully done" checkpoint if you're deliberately splitting
# the download across several days (e.g. one product per day,
# so the machine doesn't need to stay on for a multi-day
# continuous run). These four wrappers each run exactly one
# product to completion, so a day's work has a clear boundary:
# run one of these, then --extract and --database (both already
# skip whatever they processed on previous days -- see
# _extract_one_archive()'s completion marker and
# database_manifest.csv), then move to the next product the
# next day. Run --filter only once all four have completed --
# it's meant to query the complete database, so running it
# after only part of the corpus is downloaded just produces an
# incomplete filtered_patents.csv you'd need to redo later.
#
# Each of these is a direct call into the same download_product()
# used inside download_everything() above, with the same
# arguments -- no new download logic, just a narrower entry
# point into what already exists.

def download_aps_only(sample=False):

    manifest = load_download_manifest()

    manifest_lock = threading.Lock()

    download_product(

        product_type="GRANTS_APS",

        product_identifier="PTGRAPS",

        destination_directory=
            GRANTS_APS_DIR,

        manifest=
            manifest,

        manifest_lock=
            manifest_lock,

        sample=
            sample,

    )


def download_sgml_only(sample=False):

    manifest = load_download_manifest()

    manifest_lock = threading.Lock()

    download_product(

        product_type="GRANTS_SGML",

        product_identifier="PTGRSGM",

        destination_directory=
            GRANTS_SGML_DIR,

        manifest=
            manifest,

        manifest_lock=
            manifest_lock,

        sample=
            sample,

    )


def download_xml_grants_only(sample=False):

    manifest = load_download_manifest()

    manifest_lock = threading.Lock()

    download_product(

        product_type="GRANTS_XML",

        product_identifier="PTGRXML",

        destination_directory=
            GRANTS_XML_DIR,

        manifest=
            manifest,

        manifest_lock=
            manifest_lock,

        sample=
            sample,

    )


def download_appxml_only(sample=False):

    manifest = load_download_manifest()

    manifest_lock = threading.Lock()

    download_product(

        product_type="APPLICATIONS_XML",

        product_identifier="APPXML",

        destination_directory=
            APPLICATIONS_DIR,

        manifest=
            manifest,

        manifest_lock=
            manifest_lock,

        sample=
            sample,

    )


# ============================================================
# COMMAND LINE
# ============================================================

def main():

    parser = argparse.ArgumentParser(

        description=
            "USPTO Patent Research Pipeline"

    )


    parser.add_argument(

        "--test",

        action="store_true",

        help=
            "Test USPTO API connection"

    )


    parser.add_argument(

        "--download",

        action="store_true",

        help=
            "Download USPTO bulk files (all four products)"

    )


    parser.add_argument(

        "--download-aps",

        action="store_true",

        help=(
            "Download only the APS/Green Book grants product "
            "(PTGRAPS, ~1976-1996). Use this to split the "
            "download across multiple sessions/days instead of "
            "one long --download run -- --extract and --database "
            "are both already safe to rerun daily without "
            "redoing completed work."
        )

    )


    parser.add_argument(

        "--download-sgml",

        action="store_true",

        help=(
            "Download only the SGML grants product "
            "(PTGRSGM, ~1996-2001)."
        )

    )


    parser.add_argument(

        "--download-xml",

        action="store_true",

        help=(
            "Download only the XML grants product "
            "(PTGRXML, ~2001-present)."
        )

    )


    parser.add_argument(

        "--download-appxml",

        action="store_true",

        help=(
            "Download only the XML applications product "
            "(APPXML, ~2001-present)."
        )

    )


    parser.add_argument(

        "--sample",

        action="store_true",

        help=(
            "Download only a small sample of files "
            f"(spread across each product's date range) "
            "instead of the full corpus. Use this to get "
            "real files for testing the parsers before "
            "committing to a full --download."
        )

    )


    parser.add_argument(

        "--sizecheck",

        action="store_true",

        help=(
            "Query USPTO for the total size/file count of "
            "each of the 4 products WITHOUT downloading "
            "anything -- use this to plan storage before "
            "committing to a full --download."
        )

    )


    parser.add_argument(

        "--extract",

        action="store_true",

        help=
            "Extract downloaded archives"

    )


    parser.add_argument(

        "--database",

        action="store_true",

        help=
            "Build local patent database"

    )


    parser.add_argument(

        "--limit",

        type=int,

        default=None,

        metavar="N",

        help=(
            "With --pdf: download at most N PDFs this run. "
            "With --database: process only the next N pending "
            "files this run, then stop cleanly, instead of the "
            "whole remaining corpus in one long invocation. "
            "Useful for getting real, bounded checkpoints of "
            "progress on a large corpus rather than a single "
            "multi-hour run you have to hope finishes clean -- "
            "e.g. --database --limit 50 processes 50 files and "
            "returns; run it again (with or without --limit) to "
            "continue from exactly where it left off, same as "
            "any other --database run."
        )

    )


    parser.add_argument(

        "--skip-errors",

        action="store_true",

        help=(
            "With --database: don't retry files that already "
            "failed in a previous run -- only process untouched "
            "files. Use this to get through a full first pass of "
            "the corpus without a known-problematic file "
            "repeatedly interrupting progress; come back with "
            "--retry-errors-only afterward to work through what "
            "was set aside."
        )

    )


    parser.add_argument(

        "--retry-errors-only",

        action="store_true",

        help=(
            "With --database: process ONLY files that previously "
            "failed (status=error in the manifest) -- skips both "
            "already-succeeded and never-yet-tried files. Use "
            "this after a --skip-errors first pass to go back and "
            "work through what was set aside, separately from the "
            "rest of the corpus. Combine with --limit to do this "
            "in small batches too. With --pdf: retry ONLY PDFs "
            "that previously failed (normal --pdf runs skip them)."
        )

    )


    parser.add_argument(

        "--format",

        choices=["aps", "sgml", "xml", "appxml"],

        default=None,

        help=(
            "With --pdf: download only that product's filter "
            "results (filtered_patents_<product>.csv). "
            "With --filter: filter only that ONE product's "
            "database (results also merged into filtered_patents.csv). "
            "With --database: only process files from this ONE "
            "product's extracted folder (aps, sgml, xml, or "
            "appxml), instead of the whole EXTRACTED_DIR. Lets "
            "you finish one format completely -- e.g. "
            "--database --format aps -- before moving to the "
            "next, the same way --download-aps/--download-sgml/"
            "etc. let you split the download stage. Combine "
            "freely with --limit, --skip-errors, and "
            "--retry-errors-only."
        )

    )


    parser.add_argument(
        "--backup",
        action="store_true",
        help=(
            "With --database: after the run, copy each processed "
            "product's database to DATABASE/BACKUP/. Off by default "
            "(it's a full copy of a very large file); best used once "
            "per product when it's finished."
        ),
    )


    parser.add_argument(
        "--no-backup",
        action="store_true",
        help="Accepted for compatibility; backups are now opt-in (--backup).",
    )


    parser.add_argument(
        "--workers",
        type=int,
        default=None,
        metavar="N",
        help=(
            f"With --database: number of parsing processes "
            f"(default {DATABASE_PARSE_WORKERS}). Lower it if the "
            f"machine runs hot or needs to stay responsive."
        ),
    )


    parser.add_argument(
        "--requeue-merged",
        action="store_true",
        help=(
            "One-time repair: mark every XML file that pipeline.log shows "
            "lost records to the merged-record bug (fixed in v15.1) for "
            "reparse. Then run --database --format <product> "
            "--retry-errors-only."
        ),
    )


    parser.add_argument(
        "--migrate-db",
        action="store_true",
        help=(
            "One-time: copy the v14 single USPTO_PATENTS.sqlite into "
            "the v15 per-product databases, so nothing already parsed "
            "has to be re-parsed. Resumable; never modifies the old file."
        ),
    )


    parser.add_argument(

        "--filter",

        action="store_true",

        help=
            "Filter database using CPC and keywords"

    )


    parser.add_argument(

        "--pdf",

        action="store_true",

        help=
            "Download PDFs for selected patents"

    )


    parser.add_argument(

        "--all",

        action="store_true",

        help=
            "Run all pipeline stages"

    )


    args = parser.parse_args()


    # --------------------------------------------------------
    # TEST
    # --------------------------------------------------------

    if args.test:

        test_connection()

        return


    # --------------------------------------------------------
    # ALL
    # --------------------------------------------------------

    if args.all:

        download_everything()

        extract_all_archives()

        build_database()

        filter_patents()

        download_selected_pdfs()

        return


    # --------------------------------------------------------
    # SAMPLE
    # --------------------------------------------------------

    if args.sample:

        download_everything(sample=True)

        return


    # --------------------------------------------------------
    # SIZE CHECK
    # --------------------------------------------------------

    if args.sizecheck:

        check_total_corpus_size()

        return


    # --------------------------------------------------------
    # INDIVIDUAL STAGES
    # --------------------------------------------------------

    if args.download:

        download_everything()


    if args.download_aps:

        download_aps_only()


    if args.download_sgml:

        download_sgml_only()


    if args.download_xml:

        download_xml_grants_only()


    if args.download_appxml:

        download_appxml_only()


    if args.extract:

        extract_all_archives()


    if args.migrate_db:

        migrate_legacy_database()


    if args.requeue_merged:

        mark_merged_chunk_files_for_reparse()


    if args.database:

        build_database(

            limit=args.limit,
            skip_errors=args.skip_errors,
            retry_errors_only=args.retry_errors_only,
            product_format=args.format,
            backup=args.backup,
            workers=args.workers,

        )


    if args.filter:

        filter_patents(product_format=args.format)


    if args.pdf:

        download_selected_pdfs(
            product_format=args.format,
            limit=args.limit,
            skip_errors=args.skip_errors,
            retry_errors_only=args.retry_errors_only,
        )


    # --------------------------------------------------------
    # NO ARGUMENTS
    # --------------------------------------------------------

    if not any(vars(args).values()):

        parser.print_help()


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    main()
