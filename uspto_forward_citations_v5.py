"""
USPTO FORWARD-CITATION CROSS-REFERENCE  (v5.1)
==============================================

v5.1: reads RESULTS/FILTERED\filtered_patents.csv and writes to
    RESULTS/FORWARD CITATIONS. Patent type and claim category. The Forward Citations
    sheet gains Patent Type, Claim Category and Independent
    Claim Categories (carried over from filtered_patents.csv --
    pipeline v15.10's --filter or --classify fills them). The
    pairs sheets/CSVs gain the filtered patent's type and claim
    category plus the CITING patent's type (Utility / Design /
    Application ...), read from its number. Run Notes counts
    filtered patents per claim category. No re-scan needed:
    just --export again.

v5: citation pairs split by section (APS / SGML / XML grants /
    Applications). Excel gets one pairs sheet per section, with
    "(2)", "(3)" continuation sheets past 1,000,000 rows -- every
    pair is included, nothing is truncated. Each section is also
    written as forward_citation_pairs_<section>.csv, split into
    "_part2", "_part3" files of the same size so each opens in
    Excel. The combined forward_citation_pairs.csv is unchanged.

v4 (export speed):
    - Cohort Percentile computed with a sorted lookup instead of
      comparing every patent against every other one in its
      cohort (was O(n^2) -- the main cause of slow --export).
      Identical results.
    - --skip-excel-pairs leaves the million-row Citation Pairs
      sheet out of the workbook (the pairs CSV has them all).
    - --export prints how long each step takes.

v3 (speed):
    - Reads the weekly ZIP archives in the RAW folder directly (default),
      decompressing on the fly, instead of the extracted files.
      Same records, roughly 5-10x less data read off the drive,
      and the EXTRACTED folder is no longer needed for this.
      --source extracted keeps the old behaviour.
    - Scans with worker PROCESSES (real parallelism across CPU
      cores) instead of threads.
    - Ignores the pipeline's ".EXTRACTION_COMPLETE" marker files
      (they were being counted as corpus files and logged as
      "unrecognized format" errors in extracted mode).
    - Cheap prefilters on the per-line record-boundary checks.
    - A v2 index can't be resumed in v3 (it tracked extracted
      files, v3 tracks zips): run --reset once.

v2 (for uspto_patent_pipeline_v15.py):
    - APS numbers: the Green Book WKU's 9th character is a check
      character and is now always dropped for APS records -- on
      both the filtered side and the citing side. Fixes APS
      patents whose check character is "&" (e.g. US06333061&),
      which could never match a citation, and stops a patent
      that appears in BOTH the 2001 APS and SGML files from
      being counted twice as a citing patent.
    - Duplicate filtered patents (same patent in APS and SGML)
      are reported once.
    - XML: a record whose <?xml declaration is glued onto the
      previous record's closing tag ("</sequence-cwu><?xml ...")
      is now split correctly (same fix as pipeline v15.1).
    - If you ran v1 before, run --reset once before --scan.

Purpose:
    Takes the patent IDs produced by the main pipeline's
    --filter stage (filtered_patents.csv) and finds, for each
    one, every OTHER patent in the local corpus that CITES it.

    This is the reverse of what's in the raw data. Each patent
    record lists the references IT cites (backward citations).
    Nothing in the raw data says "here is who cites me." So we
    build that by scanning every record in the corpus, reading
    its cited-reference list, and inverting the relationship.

    Output is an Excel workbook with one row per filtered
    patent:

        Filtered Patent ID | Times Cited | Citations Per Year |
        Cohort Percentile | Cited by Applicant | Cited by
        Examiner | Cited by Other/Unknown | Citing Patent IDs...

    plus a complete long-format CSV (one row per citing/cited
    pair -- also carrying the citing patent's own date and the
    examiner/applicant/other category of that specific
    citation), which is the authoritative output: Excel has
    hard cell/row limits that a 500,000-patent corpus can
    exceed.

    WHY RAW COUNT ISN'T ENOUGH, AND WHAT'S DONE ABOUT IT:
    A raw forward-citation total is a weak "value" proxy on
    its own -- it's dominated by how OLD a patent is (more
    years to accumulate citations) and by EXAMINER search
    overlap (which reflects prior-art search practice, not the
    field building on the invention). Two things correct for
    this, computed in export_results():

        Citations Per Year   raw count normalized by years the
                              patent has sat inside the corpus's
                              observable window -- see
                              CORPUS_COVERAGE_END_DATE

        Cohort Percentile    raw count ranked against other
                              FILTERED patents granted in the
                              same multi-year window -- see
                              CITATION_COHORT_YEAR_WINDOW

    and separately, capturing citation category lets you see
    the applicant-cited count, which is a stronger influence
    signal than the raw total (which mixes in examiner search
    hits). See "CITATION RATE NORMALIZATION" further down for
    the full reasoning and caveats on both metrics.

Why this is a SEPARATE script and not another --stage:
    The main pipeline module does real work at import time
    (loads every CPC workbook, builds regex tables, creates
    directories). Keeping this standalone means it can run
    concurrently with other stages, be tested on its own, and
    can't destabilize a 7,500-line file that already works.
    See "WIRING THIS INTO THE MAIN PIPELINE" at the bottom if
    you'd rather have it as `--forward-citations`.

WHERE THE CITED-REFERENCE DATA LIVES (all three eras):

    XML (PTGRXML / APPXML, ~2001-present)
        <us-references-cited>            (DTD v4.3+)
        <references-cited>               (older DTDs)
          <us-citation> / <citation>
            <patcit num="00001">
              <document-id>
                <country>US</country>
                <doc-number>2374201</doc-number>
                <kind>A</kind>
        Only <patcit> with country US is used. <nplcit>
        (non-patent literature) and foreign countries are
        skipped.

    SGML (PTGRSGM, ~1996-2001)
        <B560>
          <B561>                          patent citations
            <PCIT><DOC><DNUM><PDAT>5658432</PDAT>
          <B562>                          non-patent lit -- skipped
        A foreign citation carries <CTRY> inside <DOC>; those
        are skipped. Design cites appear as "D. 254336" and are
        normalized to D254336.

    APS / Green Book (PTGRAPS, ~1976-1996)
        UREF                              US patent references
          PNO  2790186
          ISD  19570400
        OREF (other/non-patent) and FREF (foreign) are skipped.

IMPORTANT LIMITATIONS -- read these, they affect interpretation:

    1. Citations are only ever found in the citing document.
       A patent can only be cited by documents published
       AFTER it. If your extracted corpus stops at some year,
       forward-citation counts are truncated at that year.

    2. Coverage is only as complete as the scanned corpus (RAW zips or EXTRACTED). If you
       ran --sample rather than a full --download/--extract,
       these counts are a sample, not a true total. The script
       prints corpus coverage stats so this is never silent.

    3. Pre-grant application publications (APPXML) frequently
       carry NO references-cited section at publication --
       references are added during examination and appear in
       the granted document. So most citing documents will be
       grants. That's expected, not a bug.

    4. This counts US patent-document citations only. Foreign
       family members citing your patents (EPO/WIPO) are not
       in USPTO bulk data. That needs the EPO OPS / BigQuery
       path.

    5. APS-era patent numbers carry a CHECK DIGIT (see
       _canonical_keys). This is handled, but --verify will
       show you exactly how the matching resolved so you can
       confirm it against real data rather than trust it.

HOW TO RUN:
    --scan        Scan the corpus and build the citation index
                  (restartable -- re-running skips files that
                  were already scanned)
    --export      Write the Excel workbook + CSV from the index
    --all         --scan then --export
    --sample N    Scan only the first N files (do this first)
    --verify      Report how filtered patents matched, plus
                  index/corpus coverage stats -- no scanning
    --reset       Delete the citation index and start over

Recommended first run:
    python uspto_forward_citations_v5.py --sample 25 --all
"""

# ============================================================
# IMPORTS
# ============================================================

import os
import re
import csv
import sys
import time
import sqlite3
import argparse
import bisect
import io
import threading
import zipfile
from pathlib import Path
from datetime import datetime
from collections import defaultdict
from concurrent.futures import ProcessPoolExecutor, as_completed

from tqdm import tqdm
import openpyxl
from openpyxl.utils import get_column_letter


# ============================================================
# USER CONFIGURATION
# ============================================================
#
# These MUST match the values in uspto_patent_pipeline_v15.py.
# Only the root needs changing -- everything else is derived
# the same way the main pipeline derives it.
# ------------------------------------------------------------

BULK_DATA_ROOT_DIR = Path(r"D:\uspto_corpus")

DATABASE_ROOT_DIR = BULK_DATA_ROOT_DIR


# How many files to scan in parallel.
#
# This stage is I/O bound, not CPU bound -- it reads a very
# large volume off disk and does comparatively cheap regex
# work on the first chunk of each record. Threads (not
# processes) are correct here: they overlap disk waits without
# the Windows spawn-reimport problem that ProcessPoolExecutor
# would hit against the main pipeline's import-time work.
#
# On an external HDD, more than ~4 threads usually makes
# things SLOWER (seek thrash). On an SSD, 8 is reasonable.

SCAN_WORKERS = 4

# Where to read the corpus from:
#   "zip"        the weekly archives under RAW\ (default; fastest
#                on a spinning drive -- far fewer bytes to read)
#   "extracted"  the unzipped copies under EXTRACTED\
# Override per run with --source.
SCAN_SOURCE = "zip"


# Only the head of each record is examined. Bibliographic data
# and the cited-reference list always appear before the
# description/claims in all three formats, so there's no reason
# to read a 4 MB description into memory to find a citation
# list that ended 30 KB in. This cap is the safety net for a
# malformed record with no recognizable end marker.

RECORD_HEAD_CAP_BYTES = 512 * 1024


# Citation pairs are split by SECTION (the filtered patent's
# product: APS, SGML, XML grants, Applications). Each section gets
# its own Excel sheet; a section with more than this many pairs
# continues onto "(2)", "(3)", ... sheets. Excel's hard limit is
# 1,048,576 rows per sheet including the header. The per-section
# pairs CSVs are split into parts of the same size, so every part
# opens in Excel without losing rows.

EXCEL_PAIR_ROW_LIMIT = 1_000_000

# The pairs sheets duplicate the pairs CSVs and are the slowest
# part of --export (and make the workbook large). Set False, or
# pass --skip-excel-pairs, to leave them out of the workbook.
EXCEL_INCLUDE_PAIRS_SHEET = True

# Section labels, in output order. Sheet names must be <= 31
# characters.
PAIR_SECTIONS = [
    ("aps", "APS"),
    ("sgml", "SGML"),
    ("xml", "XML grants"),
    ("appxml", "Applications"),
    ("other", "Other"),
]


# Excel's hard per-cell character limit is 32,767. Citing-patent
# lists are split across as many columns as needed, each kept
# under this.

EXCEL_CELL_CHAR_BUDGET = 30_000


# ------------------------------------------------------------
# CITATION RATE NORMALIZATION
# ------------------------------------------------------------
#
# A raw forward-citation count is a weak value signal on its
# own: it's dominated by AGE (a 1985 patent has had 40 years to
# accumulate citations, a 2023 one has had two) and by how
# citation PRACTICE has drifted across the corpus's 1976-present
# span (more liberal IDS citation in later decades, for one).
#
# Two normalized metrics correct for this, computed per
# filtered patent in export_results():
#
#   Citations Per Year
#       raw count / years the patent has sat inside the
#       corpus's OBSERVABLE window (see
#       CORPUS_COVERAGE_END_DATE below) -- turns a flat total
#       into a rate.
#
#   Cohort Percentile
#       each patent's raw count ranked against other FILTERED
#       patents granted in the same multi-year window (see
#       CITATION_COHORT_YEAR_WINDOW) -- corrects for both age
#       and era-level citation-practice drift at once, at the
#       cost of only being meaningful within YOUR filtered set,
#       not the whole corpus.
#
# Neither is a substitute for judgment -- the citation-arrival
# curve is itself skewed (most patents peak 3-6 years after
# grant, not on a straight line), and a citation count still
# doesn't distinguish "genuinely influential" from "in a
# crowded, heavily-cited field." Treat both as one input among
# several (see the Run Notes sheet in the exported workbook).
# ------------------------------------------------------------

# If you know precisely how far forward your extracted corpus
# reaches (e.g. --download/--extract was only run through a
# certain year), set this as "YYYYMMDD" for a more accurate
# rate denominator. Left as None, the script uses the LATEST
# citing_date it actually found across the whole index as a
# proxy for the corpus's forward edge -- reasonable, but
# slightly conservative, since the true corpus edge can be
# later than the last citation that happened to land on one of
# your filtered patents.

CORPUS_COVERAGE_END_DATE = None

# Filtered patents are grouped into buckets of this many grant
# years (e.g. 5 -> 1980-1984, 1985-1989, ...) before computing
# cohort percentile. Wider buckets trade cohort precision for
# more patents per bucket -- on a small filtered set, a 1-year
# bucket can leave a patent being "ranked" against itself, which
# is why a cohort needs at least CITATION_COHORT_MIN_SIZE peers
# before a percentile is reported at all (see below).

CITATION_COHORT_YEAR_WINDOW = 5

# Minimum number of filtered patents that must share a cohort
# bucket before a percentile is computed for it. Below this, a
# percentile is statistically meaningless noise -- the column is
# left blank rather than reporting a number that looks precise
# but isn't.

CITATION_COHORT_MIN_SIZE = 5

# A patent granted very recently hasn't had time to accumulate
# citations, so dividing by its true (near-zero) age would
# produce wildly unstable, misleadingly huge rates. This floor
# keeps the denominator sane for last year's grants.

CITATION_RATE_MIN_YEARS = 1.0


# ------------------------------------------------------------
# DERIVED PATHS (mirrors the main pipeline)
# ------------------------------------------------------------

EXTRACTED_DIR = BULK_DATA_ROOT_DIR / "EXTRACTED"

RAW_DIR = BULK_DATA_ROOT_DIR / "RAW"

DATABASE_DIR = DATABASE_ROOT_DIR / "DATABASE"

RESULTS_DIR = BULK_DATA_ROOT_DIR / "RESULTS"

LOG_DIR = BULK_DATA_ROOT_DIR / "LOGS"


# The main pipeline's --filter output (RESULTS\FILTERED) and
# this script's output (RESULTS\FORWARD CITATIONS).
FILTER_RESULTS_DIR = RESULTS_DIR / "FILTERED"

FORWARD_CITATIONS_DIR = RESULTS_DIR / "FORWARD CITATIONS"

FILTER_RESULTS = FILTER_RESULTS_DIR / "filtered_patents.csv"

CITATION_INDEX_DB = DATABASE_DIR / "FORWARD_CITATIONS.sqlite"

FORWARD_CITATION_XLSX = FORWARD_CITATIONS_DIR / "forward_citations.xlsx"

FORWARD_CITATION_PAIRS_CSV = (
    FORWARD_CITATIONS_DIR / "forward_citation_pairs.csv"
)

FORWARD_CITATION_VERIFY_CSV = (
    FORWARD_CITATIONS_DIR / "forward_citation_match_report.csv"
)


# Same ignore rules the main pipeline's discover_patent_files()
# uses, so both stages agree on what counts as a data file.

IGNORED_FILE_SUFFIXES = {
    ".tif", ".tiff", ".jpg", ".jpeg", ".png", ".gif",
    ".pdf", ".zip", ".gz", ".tar", ".dtd", ".ent",
    ".css", ".xsl", ".xsd", ".log", ".md5", ".sha1",
}

IGNORED_FILE_NAMES = {
    "Thumbs.db", ".DS_Store", "readme.txt", "README.txt",
    ".EXTRACTION_COMPLETE",
}


# ============================================================
# PATENT NUMBER NORMALIZATION
# ============================================================
#
# THE CORE PROBLEM THIS SOLVES:
#
# The same patent is written five different ways depending on
# where you found it:
#
#   filtered_patents.csv (XML era)   US10123456B2
#   filtered_patents.csv (SGML era)  US06167569A
#   filtered_patents.csv (APS era)   US39302711   <- check digit!
#   cited in SGML                    D. 254336
#   cited in APS                     2790186
#   cited in XML                     2374201
#
# Comparing any of those directly gives zero matches. Every
# number on both sides is reduced to one canonical key:
#
#   US-3930271      utility
#   US-D435713      design
#   US-RE28671      reissue
#   US-PP11728      plant
#   US-20140123360  pre-grant application publication
#
# This is the same class of bug as the CPC whitespace issue --
# silent, total, and only visible as "why did nothing match."
# ------------------------------------------------------------

# Kind codes that can trail a patent number. Order matters:
# longest first, so "A1" is stripped before "A".
_KIND_CODE_PATTERN = re.compile(
    r"(?:A1|A2|A9|B1|B2|B9|C1|C2|C3|E1|H1|P1|P2|P3|"
    r"S1|S|A|B|C|E|P|H|T)$"
)

# Type prefixes that are part of the identity of the patent
# (a design patent D123456 is NOT utility patent 123456).
_TYPE_PREFIX_PATTERN = re.compile(
    r"^(RE|PP|SIR|AI|D|H|T|X)"
)

# Junk that shows up inside cited references, especially in
# SGML/APS ("D. 254,336", "US 4,165,542").
_JUNK_CHAR_PATTERN = re.compile(r"[\s.,\-/]+")


def _canonical_keys(
    raw_number,
    aps_check_digit_possible=False,
    aps_wku=False,
):
    """
    Reduces any of the number formats above to canonical keys.

    Returns a LIST, best guess first, because one input can be
    genuinely ambiguous:

    The APS/Green Book WKU field appends a CHECK DIGIT to the
    patent number. The main pipeline stores that raw, so an APS
    record for patent 3,930,271 is stored as "US39302711" --
    eight digits where seven are the real number. A cited
    reference to that same patent in a later patent says
    "3930271". Those never match unless the check digit is
    stripped.

    But blindly stripping the last digit off every 8-digit
    number would break real 8-digit patents (10,000,000+,
    granted from 2018 onward), which is why this returns
    candidates instead of committing to one interpretation --
    and why --verify reports which candidate actually matched.
    """

    if not raw_number:
        return []

    value = str(raw_number).strip().upper()

    if not value:
        return []

    # Strip the country prefix. Careful: "US" is also not a
    # type prefix, so this is done before prefix detection.
    if value.startswith("US"):
        value = value[2:]

    # v2: a number KNOWN to come from an APS record is a Green
    # Book WKU -- always 9 characters, the last being a check
    # character (usually a digit, sometimes "&"). Drop it
    # deterministically instead of offering two guesses.
    # Confirmed against real APS files and the USPTO PDF server
    # (WKU 047948337 -> patent 4794833; WKU 06333061& ->
    # 6333061; RE0374423 -> RE37442; T01052020 -> T105202).
    if aps_wku:
        value = value.strip()

        if len(value) == 9:
            value = value[:-1]

    value = _JUNK_CHAR_PATTERN.sub("", value)

    if not value:
        return []

    # Split off a type prefix (D / RE / PP / H / T / ...).
    prefix = ""
    prefix_match = _TYPE_PREFIX_PATTERN.match(value)

    if prefix_match:
        prefix = prefix_match.group(1)
        value = value[len(prefix):]

    # An 11-digit number beginning with a plausible year is a
    # pre-grant application publication (20140123360), not a
    # patent number. It gets no kind-code stripping and no
    # zero-stripping -- the leading digits are the year.
    application_match = re.match(r"^(\d{11})[A-Z]?\d?$", value)

    if not prefix and application_match:
        digits = application_match.group(1)
        year = int(digits[:4])

        if 1900 <= year <= 2100:
            return [f"US-{digits}"]

    # Strip a trailing kind code, but only if what remains is
    # still digits. "6167569A" -> "6167569". Guards against
    # eating a digit from something already clean.
    kind_match = _KIND_CODE_PATTERN.search(value)

    if kind_match and kind_match.start() > 0:
        stripped = value[:kind_match.start()]

        if stripped.isdigit():
            value = stripped

    if not value.isdigit():
        # Something we don't recognize. Return it normalized
        # rather than dropping it silently -- it'll simply
        # never match anything, and --verify will show it.
        return [f"US-{prefix}{value}"] if value else []

    primary_digits = value.lstrip("0") or "0"

    keys = [f"US-{prefix}{primary_digits}"]

    # APS check-digit alternate. Only offered when the caller
    # says the number came from (or might have come from) an
    # APS-era record, and only when dropping a digit still
    # leaves a plausible patent number.
    if aps_check_digit_possible and len(primary_digits) > 1:
        without_check = primary_digits[:-1].lstrip("0")

        if without_check:
            alternate = f"US-{prefix}{without_check}"

            if alternate not in keys:
                keys.append(alternate)

    return keys


def _canonical_key(raw_number, aps_wku=False):
    """Single best-guess key, for the citing side. aps_wku=True
    for numbers read from an APS record's WKU field."""

    keys = _canonical_keys(raw_number, aps_wku=aps_wku)

    return keys[0] if keys else None


# ============================================================
# PATENT TYPE  (v5.1)
# ============================================================
#
# Same rules as classify_patent_type() in
# uspto_patent_pipeline_v15.py (kept as a copy so this script
# still runs standalone). Used for citing patents, which only
# come with a number, and as a fallback for filtered patents
# when filtered_patents.csv predates the patent_type column.

PATENT_TYPE_UTILITY = "Utility"
PATENT_TYPE_DESIGN = "Design"
PATENT_TYPE_PLANT = "Plant"
PATENT_TYPE_REISSUE = "Reissue"
PATENT_TYPE_SIR = "Statutory Invention Registration (SIR)"
PATENT_TYPE_DEFENSIVE = "Defensive Publication"
PATENT_TYPE_APPLICATION = "Application (utility)"
PATENT_TYPE_PLANT_APPLICATION = "Application (plant)"

_patent_type_cache = {}


def classify_patent_type(number, document_type=None):

    """Utility / Design / Plant / Reissue / SIR / Defensive
    Publication / Application, from the number and kind code."""

    cache_key = (number, document_type)

    cached = _patent_type_cache.get(cache_key)

    if cached is not None:
        return cached

    value = (number or "").strip().upper()

    if value.startswith("US"):
        value = value[2:]

    value = re.sub(r"[^A-Z0-9]", "", value)

    doc_type = (document_type or "").strip().lower()

    application_match = re.match(r"^(\d{11})([A-Z]\d?)?$", value)

    if doc_type == "application" or (
        application_match is not None
        and 1900 <= int(value[:4]) <= 2100
    ):
        kind = application_match.group(2) if application_match else ""
        result = (
            PATENT_TYPE_PLANT_APPLICATION
            if (kind or "").startswith("P")
            else PATENT_TYPE_APPLICATION
        )
    elif value.startswith("PP"):
        result = PATENT_TYPE_PLANT
    elif value.startswith("RE"):
        result = PATENT_TYPE_REISSUE
    elif value.startswith("D"):
        result = PATENT_TYPE_DESIGN
    elif value.startswith("H") or doc_type == "sir":
        result = PATENT_TYPE_SIR
    elif value.startswith("T"):
        result = PATENT_TYPE_DEFENSIVE
    elif value[:1].isdigit():
        result = PATENT_TYPE_UTILITY
    else:
        result = ""

    if len(_patent_type_cache) < 2_000_000:
        _patent_type_cache[cache_key] = result

    return result


# ============================================================
# CITATION RATE HELPERS
# ============================================================

def _parse_year(date_value):
    """
    Extracts a plausible 4-digit year from a date value in any
    of the shapes this project's date fields show up in --
    "YYYYMMDD" (the normal case), a bare "YYYY", or something
    with stray punctuation. Returns None rather than guessing
    if nothing plausible is found, so a malformed or missing
    date quietly drops that patent OUT of rate/percentile
    calculations instead of polluting them with a wrong year.
    """

    if not date_value:
        return None

    digits = re.sub(r"\D", "", str(date_value))

    if len(digits) < 4:
        return None

    try:
        year = int(digits[:4])

    except ValueError:
        return None

    if 1900 <= year <= 2100:
        return year

    return None


def _cohort_bucket(year, window=CITATION_COHORT_YEAR_WINDOW):
    """
    Buckets a grant year into a window-sized cohort, e.g. with
    window=5: 1983 -> 1980, 1987 -> 1985. Returns None if year
    is None, so callers can distinguish "no cohort" from a
    real bucket at year 0 (which can't occur anyway, but keeps
    intent explicit).
    """

    if year is None:
        return None

    return (year // window) * window


# ============================================================
# RECORD STREAMING
# ============================================================
#
# Bulk files concatenate thousands of records. Reading a whole
# file into memory would be 500 MB+ per file, times SCAN_WORKERS
# threads. Instead each record's HEAD is yielded -- everything
# from the record's start marker up to the point where the
# citation data is known to have ended.
# ------------------------------------------------------------

def _split_glued_xml_declarations(lines):
    """
    Yields lines, splitting any line where an XML declaration
    starts part-way through (USPTO sometimes writes
    "</sequence-cwu><?xml version=...?>" on one line). Without
    this, the next patent is glued onto the sequence listing and
    its citations are attributed to the wrong record or lost.
    """

    for line in lines:

        position = line.find(b"<?xml", 1)

        while position != -1:

            if line[:position].strip() and re.match(
                rb"<\?xml\s", line[position:]
            ):
                yield line[:position] + b"\n"
                line = line[position:]
                position = line.find(b"<?xml", 1)

            else:
                position = line.find(b"<?xml", position + 1)

        yield line


def _iter_record_heads(
    handle,
    is_start_line,
    is_stop_line,
    split_glued_xml=False,
):
    """
    Streams (record_head_bytes,) for each record in an open
    binary stream (a plain file, or a member inside a zip).

    `is_start_line(line)` marks a new record boundary.
    `is_stop_line(line)` marks the point past which nothing
    citation-related can appear, so buffering stops (but the
    record isn't emitted until the NEXT start marker -- the
    stop marker only stops accumulating).
    """

    buffer = []
    buffer_size = 0
    in_record = False
    collecting = False

    lines = (
        _split_glued_xml_declarations(handle)
        if split_glued_xml else handle
    )

    for line in lines:

        if is_start_line(line):

            if in_record and buffer:
                yield b"".join(buffer)

            buffer = [line]
            buffer_size = len(line)
            in_record = True
            collecting = True
            continue

        if not in_record or not collecting:
            continue

        if is_stop_line(line):
            collecting = False
            continue

        buffer.append(line)
        buffer_size += len(line)

        if buffer_size >= RECORD_HEAD_CAP_BYTES:
            collecting = False

    if in_record and buffer:
        yield b"".join(buffer)


# ============================================================
# FORMAT DETECTION
# ============================================================

def detect_format(file_path):
    """
    Sniffs the first non-empty bytes of a file to decide which
    extractor to use. Suffix alone is unreliable -- APS-era
    files often carry no extension at all, and USPTO has
    shipped SGML content in .xml-named files.
    """

    try:
        with open(file_path, "rb") as handle:
            head = handle.read(4096)

    except OSError:
        return None

    return detect_format_bytes(head)


def detect_format_bytes(head):
    """detect_format() on the first ~4 KB of a stream."""

    if not head:
        return None

    upper = head.upper()

    if b"<PATDOC" in upper or b"ST.32" in upper:
        return "sgml"

    if b"<?XML" in upper:
        return "xml"

    if b"<!DOCTYPE" in upper and (
        b"US-PATENT" in upper or b"PATENT-GRANT" in upper
    ):
        return "xml"

    # APS/Green Book: record separator is a bare "PATN" line,
    # and files begin with an "HHHHHT  APS1 ..." header.
    if upper.startswith(b"HHHHH") or b"\nPATN" in upper[:4096]:
        return "aps"

    if upper.startswith(b"PATN"):
        return "aps"

    return None


# ============================================================
# XML EXTRACTOR  (PTGRXML / APPXML, ~2001-present)
# ============================================================

_XML_START = re.compile(rb"^\s*<\?xml\s", re.IGNORECASE)

_XML_PUB_REF = re.compile(
    rb"<publication-reference>(.*?)</publication-reference>",
    re.DOTALL | re.IGNORECASE,
)

_XML_DOC_NUMBER = re.compile(
    rb"<doc-number>\s*([^<]+?)\s*</doc-number>",
    re.IGNORECASE,
)

_XML_KIND = re.compile(
    rb"<kind>\s*([^<]+?)\s*</kind>",
    re.IGNORECASE,
)

_XML_DATE = re.compile(
    rb"<date>\s*([^<]+?)\s*</date>",
    re.IGNORECASE,
)

# v4.3+ uses us-references-cited; older DTDs use
# references-cited. Both are matched.
_XML_REFS_BLOCK = re.compile(
    rb"<(us-)?references-cited>(.*?)</(us-)?references-cited>",
    re.DOTALL | re.IGNORECASE,
)

# Each individual reference (patent or non-patent) is wrapped
# in <us-citation>...</us-citation> (v4.3+) or plain
# <citation>...</citation> (older DTDs). <category> -- "cited
# by examiner" / "cited by applicant" / "cited by third party"
# -- is a SIBLING of <patcit> inside this wrapper, not inside
# patcit itself, so the wrapper has to be matched as its own
# unit to keep the category paired with the right reference.
_XML_CITATION_WRAPPER = re.compile(
    rb"<(?:us-)?citation>(.*?)</(?:us-)?citation>",
    re.DOTALL | re.IGNORECASE,
)

_XML_PATCIT = re.compile(
    rb"<patcit\b[^>]*>(.*?)</patcit>",
    re.DOTALL | re.IGNORECASE,
)

_XML_CATEGORY = re.compile(
    rb"<category>\s*([^<]+?)\s*</category>",
    re.IGNORECASE,
)

_XML_COUNTRY = re.compile(
    rb"<country>\s*([^<]+?)\s*</country>",
    re.IGNORECASE,
)


def _xml_is_stop_line(line):
    stripped = line.lstrip()[:24].lower()

    return (
        stripped.startswith(b"<description")
        or stripped.startswith(b"<claims")
        or stripped.startswith(b"<drawings")
    )


def _normalize_citation_category(raw_category):
    """
    Collapses USPTO's category text to one of three labels.
    "cited by examiner" is the common case (~85%+ of US
    citations) and reflects the EXAMINER'S search, not
    necessarily the applicant's or the field's use of the
    prior art. "cited by applicant" is the stronger influence
    signal -- it's the applicant/attorney actively pointing to
    that prior work. Third-party (post-grant proceedings, IDS
    submissions from opposing counsel) is kept distinct rather
    than folded into either.
    """

    if not raw_category:
        return "unknown"

    text = raw_category.decode("ascii", "ignore").strip().lower()

    if "examiner" in text:
        return "examiner"

    if "applicant" in text:
        return "applicant"

    if "third" in text:
        return "third_party"

    return "unknown"


def extract_citations_xml(record_head):
    """
    Returns (citing_number, citing_date, [(cited_number, category), ...])
    for one XML record.

    Only <patcit> entries whose country is US are kept.
    <nplcit> (journal articles, office actions, search reports)
    and foreign patent families are skipped -- they can't match
    a US patent ID and would only add noise.

    citing_date is the citing document's OWN publication/grant
    date (from <publication-reference>), not any date inside
    the citation block -- it's what lets citation counts be
    turned into a year-over-year curve later instead of a flat
    total.
    """

    pub_ref = _XML_PUB_REF.search(record_head)

    if not pub_ref:
        return None, None, []

    pub_block = pub_ref.group(1)

    doc_match = _XML_DOC_NUMBER.search(pub_block)

    if not doc_match:
        return None, None, []

    country_match = _XML_COUNTRY.search(pub_block)
    country = (
        country_match.group(1).decode("ascii", "ignore").strip().upper()
        if country_match else "US"
    )

    if country and country != "US":
        return None, None, []

    kind_match = _XML_KIND.search(pub_block)

    citing_number = doc_match.group(1).decode("ascii", "ignore").strip()

    if kind_match:
        citing_number += kind_match.group(1).decode("ascii", "ignore").strip()

    citing_number = "US" + citing_number

    date_match = _XML_DATE.search(pub_block)

    citing_date = (
        date_match.group(1).decode("ascii", "ignore").strip()
        if date_match else None
    )

    cited = []

    refs_match = _XML_REFS_BLOCK.search(record_head)

    if refs_match:

        refs_block = refs_match.group(2)

        for wrapper_match in _XML_CITATION_WRAPPER.finditer(refs_block):

            wrapper_block = wrapper_match.group(1)

            patcit_match = _XML_PATCIT.search(wrapper_block)

            if not patcit_match:
                # <nplcit> (non-patent literature) -- skipped.
                continue

            patcit_block = patcit_match.group(1)

            cited_country_match = _XML_COUNTRY.search(patcit_block)

            cited_country = (
                cited_country_match.group(1)
                .decode("ascii", "ignore").strip().upper()
                if cited_country_match else ""
            )

            if cited_country and cited_country != "US":
                continue

            cited_doc = _XML_DOC_NUMBER.search(patcit_block)

            if not cited_doc:
                continue

            category_match = _XML_CATEGORY.search(wrapper_block)

            category = _normalize_citation_category(
                category_match.group(1) if category_match else None
            )

            cited.append((
                cited_doc.group(1).decode("ascii", "ignore").strip(),
                category,
            ))

    return citing_number, citing_date, cited


# ============================================================
# SGML EXTRACTOR  (PTGRSGM, ~1996-2001)
# ============================================================

_SGML_START = re.compile(
    rb"^\s*(?:<!DOCTYPE\s+PATDOC|<PATDOC\b)", re.IGNORECASE
)

_SGML_B110 = re.compile(
    rb"<B110>\s*<DNUM>\s*<PDAT>\s*([^<]+?)\s*</PDAT>",
    re.DOTALL | re.IGNORECASE,
)

_SGML_B130 = re.compile(
    rb"<B130>\s*<PDAT>\s*([^<]+?)\s*</PDAT>",
    re.DOTALL | re.IGNORECASE,
)

_SGML_B140 = re.compile(
    rb"<B140>\s*<DATE>\s*<PDAT>\s*([^<]+?)\s*</PDAT>",
    re.DOTALL | re.IGNORECASE,
)

_SGML_B560 = re.compile(
    rb"<B560>(.*?)</B560>",
    re.DOTALL | re.IGNORECASE,
)

_SGML_B561 = re.compile(
    rb"<B561>(.*?)</B561>",
    re.DOTALL | re.IGNORECASE,
)

_SGML_PCIT_DNUM = re.compile(
    rb"<PCIT>.*?<DOC>.*?<DNUM>\s*<PDAT>\s*([^<]+?)\s*</PDAT>",
    re.DOTALL | re.IGNORECASE,
)

_SGML_CTRY = re.compile(
    rb"<CTRY>\s*<PDAT>\s*([^<]+?)\s*</PDAT>",
    re.IGNORECASE,
)

# SGML marks each citation's provenance with a bare tag right
# after </PCIT>. Confirmed against real USPTO sample files:
# CITED-BY-EXAMINER and CITED-BY-OTHER (SGML's era predates
# XML's three-way applicant/examiner/third-party split -- OTHER
# covers everything not examiner-sourced). It's a standalone
# tag with no separate closing tag or content, so this is
# matched as plain text presence rather than a capture group.
_SGML_CITED_BY_EXAMINER = re.compile(
    rb"CITED-BY-EXAMINER", re.IGNORECASE
)

_SGML_CITED_BY_OTHER = re.compile(
    rb"CITED-BY-OTHER", re.IGNORECASE
)


def _sgml_is_stop_line(line):
    stripped = line.lstrip()[:12].upper()

    return (
        stripped.startswith(b"<SDODE")
        or stripped.startswith(b"<SDOCL")
        or stripped.startswith(b"<SDOAB")
    )


def extract_citations_sgml(record_head):
    """
    Returns (citing_number, citing_date, [(cited_number, category), ...])
    for one SGML record.

    B561 holds patent citations; B562 holds non-patent
    literature and is ignored entirely. A foreign citation
    carries <CTRY> inside its <DOC> -- those are skipped, since
    only US numbers can match a US patent ID.

    citing_date comes from B140 (the citing document's own
    grant date), not anything inside the citation block.
    """

    b110 = _SGML_B110.search(record_head)

    if not b110:
        return None, None, []

    doc_number = b110.group(1).decode("ascii", "ignore").strip()

    b130 = _SGML_B130.search(record_head)

    kind = (
        b130.group(1).decode("ascii", "ignore").strip()
        if b130 else ""
    )

    citing_number = f"US{doc_number}{kind}"

    b140 = _SGML_B140.search(record_head)

    citing_date = (
        b140.group(1).decode("ascii", "ignore").strip()
        if b140 else None
    )

    cited = []

    b560 = _SGML_B560.search(record_head)

    if b560:

        for b561 in _SGML_B561.finditer(b560.group(1)):

            block = b561.group(1)

            ctry = _SGML_CTRY.search(block)

            if ctry:
                country = (
                    ctry.group(1).decode("ascii", "ignore").strip().upper()
                )

                if country and country != "US":
                    continue

            dnum = _SGML_PCIT_DNUM.search(block)

            if not dnum:
                continue

            # NOTE: unlike XML's explicit three-way category,
            # SGML only distinguishes "examiner" from "other".
            # "other" is NOT assumed to mean applicant -- it
            # covers applicant IDS submissions, third-party
            # submissions, and anything else non-examiner
            # alike. Don't relabel this to "applicant" without
            # separately confirming what CITED-BY-OTHER means
            # against real prosecution records.
            if _SGML_CITED_BY_EXAMINER.search(block):
                category = "examiner"
            elif _SGML_CITED_BY_OTHER.search(block):
                category = "other"
            else:
                category = "unknown"

            cited.append((
                dnum.group(1).decode("ascii", "ignore").strip(),
                category,
            ))

    return citing_number, citing_date, cited


# ============================================================
# APS / GREEN BOOK EXTRACTOR  (PTGRAPS, ~1976-1996)
# ============================================================
#
# Flat tagged text, not markup. Records separated by a line
# that is exactly "PATN". Fields are SECTION-scoped, which is
# the whole difficulty: "PNO" means a cited US patent inside a
# UREF block and a cited FOREIGN patent inside a FREF block.
# Reading PNO without tracking the current section silently
# mixes foreign patent numbers into US results.
# ------------------------------------------------------------

_APS_SECTION_MARKERS = {
    b"INVT", b"ASSG", b"REIS", b"CLAS", b"UREF", b"OREF",
    b"FREF", b"LREP", b"PARN", b"GOVT", b"BSUM", b"DETD",
    b"DRWD", b"CLMS", b"ABST", b"PRIR", b"PCTA", b"RLAP",
}

# Sections that appear AFTER all reference sections -- once one
# of these starts, there is nothing left worth reading.
_APS_STOP_SECTIONS = {
    b"BSUM", b"DETD", b"DRWD", b"CLMS", b"ABST",
}


def _aps_is_start_line(line):
    return line.rstrip(b"\r\n") == b"PATN"


def _aps_is_stop_line(line):
    return line.rstrip(b"\r\n") in _APS_STOP_SECTIONS


def extract_citations_aps(record_head):
    """
    Returns (citing_number, citing_date, [(cited_number, category), ...])
    for one APS record.

    Only PNO values inside a UREF section are collected. FREF
    (foreign) and OREF (non-patent) are deliberately skipped.

    citing_date is ISD (issue date) from the PATN section --
    the citing patent's own grant date.

    category is always "unknown" for this era: APS/Green Book
    text does not carry a per-reference examiner/applicant
    marker the way SGML (CITED-BY-EXAMINER/OTHER) and XML
    (<category>) do. Reporting "unknown" here is the honest
    answer -- guessing examiner vs. applicant for this era
    would be fabricating a distinction the source data doesn't
    make.

    NOTE: the WKU value returned here still carries its check
    digit. It's left intact rather than guessed at -- the
    check-digit question is resolved on the lookup side by
    _canonical_keys(), where both interpretations can be tried
    and reported.
    """

    citing_number = None
    citing_date = None
    cited = []
    current_section = "PATN"

    for raw_line in record_head.split(b"\n"):

        line = raw_line.rstrip(b"\r")

        if not line:
            continue

        stripped = line.strip()

        if stripped in _APS_SECTION_MARKERS and line[:1] != b" ":
            current_section = stripped.decode("ascii", "ignore")
            continue

        # Tagged field lines are "TAG  value" at column 1.
        if line[:1] == b" ":
            continue

        parts = line.split(b"  ", 1)

        if len(parts) != 2:
            continue

        tag = parts[0].strip()
        value = parts[1].strip()

        if not value:
            continue

        if current_section == "PATN" and tag == b"WKU":
            citing_number = "US" + value.decode("ascii", "ignore")

        elif current_section == "PATN" and tag == b"ISD":
            citing_date = value.decode("ascii", "ignore")

        elif current_section == "UREF" and tag == b"PNO":
            cited.append((
                value.decode("ascii", "ignore"),
                "unknown",
            ))

    return citing_number, citing_date, cited


# ============================================================
# PER-FILE SCAN
# ============================================================

_FORMAT_HANDLERS = {

    "xml": (
        lambda line: (
            line.startswith(b"<?xml")
            or (line[:1].isspace() and bool(_XML_START.match(line)))
        ),
        _xml_is_stop_line,
        extract_citations_xml,
        True,   # citing side can be an APS-style check-digit number?
    ),

    "sgml": (
        lambda line: (
            line[:2] in (b"<!", b"<P", b"<p")
            or line[:1].isspace()
        ) and bool(_SGML_START.match(line)),
        _sgml_is_stop_line,
        extract_citations_sgml,
        False,
    ),

    "aps": (
        _aps_is_start_line,
        _aps_is_stop_line,
        extract_citations_aps,
        True,
    ),

}


# Set once per worker process by _init_scan_worker(), so the
# (large) target-key set is sent to each process once instead of
# being pickled along with every one of ~4,000 tasks.
_WORKER_TARGET_KEYS = None


def _init_scan_worker(target_keys):

    global _WORKER_TARGET_KEYS

    _WORKER_TARGET_KEYS = target_keys

    _lower_process_priority()


def _lower_process_priority():

    """BELOW_NORMAL priority on Windows so the desktop stays
    responsive during a long scan. No-op elsewhere."""

    if os.name != "nt":
        return

    try:
        import ctypes

        ctypes.windll.kernel32.SetPriorityClass(
            ctypes.windll.kernel32.GetCurrentProcess(), 0x00004000
        )

    except Exception:
        pass


def _is_data_member_name(name):

    base = name.replace("\\", "/").rsplit("/", 1)[-1]

    if not base or name.endswith("/"):
        return False

    if base in IGNORED_FILE_NAMES:
        return False

    suffix = os.path.splitext(base)[1].lower()

    return suffix not in IGNORED_FILE_SUFFIXES


def _scan_stream(open_stream, target_keys, result):

    """
    Scans one data stream (a plain file or a zip member).
    `open_stream()` returns a fresh binary stream each call --
    it's called once to sniff the format and once to scan.
    Adds counts and hits into `result`. Returns the detected
    format, or None if the stream isn't patent data.
    """

    with open_stream() as handle:
        file_format = detect_format_bytes(handle.read(4096))

    if file_format is None:
        return None

    (
        is_start_line,
        is_stop_line,
        extractor,
        _,
    ) = _FORMAT_HANDLERS[file_format]

    hits = result["hits"]

    with open_stream() as raw_handle:

        handle = io.BufferedReader(raw_handle, buffer_size=1 << 20)

        for record_head in _iter_record_heads(
            handle,
            is_start_line,
            is_stop_line,
            split_glued_xml=(file_format == "xml"),
        ):

            result["records"] += 1

            citing_number, citing_date, cited_entries = extractor(
                record_head
            )

            if not citing_number or not cited_entries:
                continue

            citing_key = _canonical_key(
                citing_number,
                aps_wku=(file_format == "aps"),
            )

            # Show APS citing patents by their real number
            # (US064000017 -> US6400001), not the raw WKU.
            if file_format == "aps" and citing_key:
                citing_number = "US" + citing_key[3:]

            for cited_number, category in cited_entries:

                result["citations_seen"] += 1

                cited_key = _canonical_key(cited_number)

                if not cited_key:
                    continue

                if cited_key not in target_keys:
                    continue

                # A patent citing itself is a data artifact
                # (usually a continuation referencing its own
                # parent's publication) -- not a forward
                # citation in any meaningful sense.
                if citing_key and citing_key == cited_key:
                    continue

                hits.append((
                    cited_key,
                    citing_number,
                    citing_key or "",
                    citing_date or "",
                    category,
                ))

    return file_format


def scan_one_file(file_path, target_keys=None):

    """
    Scans one corpus unit -- a weekly ZIP archive (every data
    member inside it, decompressed on the fly) or a single
    extracted file -- and returns every citation in it that
    points AT one of the filtered patents.

    Returns a dict:
        {
            "path": str,
            "format": "xml" | "sgml" | "aps" | None,
            "records": int,
            "citations_seen": int,
            "hits": [
                (cited_key, citing_number, citing_key,
                 citing_date, category),
                ...
            ],
            "error": str | None,
        }

    citing_date is the CITING patent's own grant/publication
    date; category is "examiner", "applicant", "other",
    "third_party" or "unknown" (see the extractors).

    Only hits against the target set are returned, so memory
    and the index stay proportional to YOUR result set, not to
    all of USPTO.
    """

    if target_keys is None:
        target_keys = _WORKER_TARGET_KEYS

    file_path = Path(file_path)

    result = {
        "path": str(file_path),
        "format": None,
        "records": 0,
        "citations_seen": 0,
        "hits": [],
        "error": None,
    }

    formats = set()

    try:

        if file_path.suffix.lower() == ".zip":

            with zipfile.ZipFile(file_path) as archive:

                for info in archive.infolist():

                    if info.is_dir() or not _is_data_member_name(info.filename):
                        continue

                    found = _scan_stream(
                        lambda info=info: archive.open(info),
                        target_keys,
                        result,
                    )

                    if found:
                        formats.add(found)

        else:

            found = _scan_stream(
                lambda: open(file_path, "rb"),
                target_keys,
                result,
            )

            if found:
                formats.add(found)

    except Exception as error:

        result["error"] = f"{type(error).__name__}: {error}"

    if formats:

        result["format"] = "+".join(sorted(formats))

    elif not result["error"]:

        result["error"] = "unrecognized format"

    return result


# ============================================================
# CITATION INDEX DATABASE
# ============================================================
#
# SQLite rather than an in-memory dict, for the same reason the
# main pipeline uses one: this has to be restartable. A full
# corpus scan is an hours-to-days job on an external drive, and
# losing it to a power blip or a single malformed file would be
# unacceptable. Every file's results are committed as that file
# finishes, and already-scanned files are skipped on restart.
# ------------------------------------------------------------

def get_index_connection():

    DATABASE_DIR.mkdir(parents=True, exist_ok=True)

    connection = sqlite3.connect(
        CITATION_INDEX_DB,
        timeout=60,
    )

    connection.execute("PRAGMA journal_mode=WAL")
    connection.execute("PRAGMA synchronous=NORMAL")

    return connection


def create_index_database():

    connection = get_index_connection()

    cursor = connection.cursor()

    # One row per (cited patent, citing patent) pair. The
    # PRIMARY KEY makes re-scanning a file idempotent -- an
    # interrupted run that partially wrote a file's results
    # can safely rewrite them.
    cursor.execute("""

        CREATE TABLE IF NOT EXISTS forward_citations (

            cited_key       TEXT NOT NULL,
            citing_key      TEXT NOT NULL,
            citing_number   TEXT,
            citing_date     TEXT,
            category        TEXT,
            source_format   TEXT,
            source_file     TEXT,

            PRIMARY KEY (cited_key, citing_key)
        )

    """)

    # Existing indexes created before citing_date/category
    # existed won't have these columns -- add them if missing
    # so re-running --scan against an older index doesn't
    # crash instead of just picking up where it left off.
    existing_columns = {
        row[1] for row in
        cursor.execute("PRAGMA table_info(forward_citations)")
    }

    if "citing_date" not in existing_columns:
        cursor.execute(
            "ALTER TABLE forward_citations ADD COLUMN citing_date TEXT"
        )

    if "category" not in existing_columns:
        cursor.execute(
            "ALTER TABLE forward_citations ADD COLUMN category TEXT"
        )

    cursor.execute("""

        CREATE INDEX IF NOT EXISTS idx_cited_key
        ON forward_citations(cited_key)

    """)

    # Restart bookkeeping. Size+mtime are stored so that a file
    # that was re-downloaded or re-extracted gets re-scanned
    # rather than wrongly skipped.
    cursor.execute("""

        CREATE TABLE IF NOT EXISTS scanned_files (

            path            TEXT PRIMARY KEY,
            size            INTEGER,
            mtime           REAL,
            source_format   TEXT,
            records         INTEGER,
            citations_seen  INTEGER,
            hits            INTEGER,
            error           TEXT,
            scanned_at      TEXT
        )

    """)

    connection.commit()

    return connection


def load_scanned_files(connection):
    """
    Returns {path: (size, mtime)} for files already scanned
    WITHOUT error. Files that errored are retried on the next
    run rather than silently written off.
    """

    rows = connection.execute("""

        SELECT path, size, mtime
        FROM scanned_files
        WHERE error IS NULL OR error = ''

    """).fetchall()

    return {row[0]: (row[1], row[2]) for row in rows}


# ============================================================
# LOAD THE FILTERED PATENT LIST
# ============================================================

def _section_from_source_path(source_path):

    """
    Which product a filtered patent came from, from the folder
    its record was parsed out of (e.g. ...\\GRANTS\\XML\\...).
    """

    parts = [
        part.upper()
        for part in (source_path or "").replace("\\", "/").split("/")
    ]

    if "APPLICATIONS" in parts:
        return "appxml"

    if "GRANTS" in parts:

        after = parts[parts.index("GRANTS") + 1:parts.index("GRANTS") + 2]

        if after == ["APS"]:
            return "aps"

        if after == ["SGML"]:
            return "sgml"

        if after == ["XML"]:
            return "xml"

    if "APS" in parts:
        return "aps"

    return "other"


def load_filtered_patents():
    """
    Reads filtered_patents.csv (the --filter stage's output)
    and builds the lookup structures the scan needs.

    Returns (records, target_keys) where:

        records      list of dicts, one per filtered patent,
                     carrying the original ID plus every
                     canonical key it might be found under

        target_keys  set of every canonical key across all
                     filtered patents -- this is what the
                     scanner tests each citation against

    The APS check-digit ambiguity is handled by registering
    BOTH candidate keys for any patent that could be affected.
    Registering both is safe: a citation can only match one of
    them in practice, and --verify reports which one hit, so
    the ambiguity is measured rather than assumed away.
    """

    if not FILTER_RESULTS.exists():
        raise SystemExit(
            f"\nfiltered_patents.csv not found at:\n"
            f"  {FILTER_RESULTS}\n\n"
            f"Run the main pipeline's --filter stage first."
        )

    records = []
    target_keys = set()
    seen_primary_keys = set()
    duplicate_count = 0

    with open(
        FILTER_RESULTS,
        "r",
        newline="",
        encoding="utf-8-sig",
    ) as handle:

        reader = csv.DictReader(handle)

        for row in reader:

            patent_number = (
                row.get("patent_number")
                or row.get("publication_number")
                or ""
            ).strip()

            if not patent_number:
                continue

            source_path = (row.get("source_path") or "")

            grant_date = (row.get("grant_date") or "").strip()

            # Is this record from the APS era, where the stored
            # number carries a check digit? Two independent
            # signals, either is enough:
            #   - the file lives under the APS product folder
            #   - the grant date predates the SGML changeover
            from_aps_file = (
                "APS" in source_path.upper().replace("\\", "/").split("/")
            )

            aps_era = from_aps_file or (
                grant_date[:4].isdigit()
                and int(grant_date[:4]) < 1997
            )

            # v2: a record from an APS file gets its WKU check
            # character dropped deterministically (one key).
            # Only an APS-era date WITHOUT an APS source path
            # falls back to offering both candidate keys.
            keys = _canonical_keys(
                patent_number,
                aps_check_digit_possible=(aps_era and not from_aps_file),
                aps_wku=from_aps_file,
            )

            if not keys:
                continue

            # The same patent can appear in both the 2001 APS and
            # SGML files (e.g. US062068660 and US06206866B1). Keep
            # the first and count the rest.
            if keys[0] in seen_primary_keys:
                duplicate_count += 1
                continue

            seen_primary_keys.add(keys[0])

            records.append({
                "patent_number": patent_number,
                "publication_number": (
                    row.get("publication_number") or ""
                ).strip(),
                "title": (row.get("title") or "").strip(),
                "patent_type": (
                    (row.get("patent_type") or "").strip()
                    or classify_patent_type(
                        patent_number, row.get("document_type")
                    )
                ),
                "claim_category": (
                    row.get("claim_category") or ""
                ).strip(),
                "independent_claim_categories": (
                    row.get("independent_claim_categories") or ""
                ).strip(),
                "grant_date": grant_date,
                "publication_date": (
                    row.get("publication_date") or ""
                ).strip(),
                "assignees": (row.get("assignees") or "").strip(),
                "cpc": (row.get("cpc") or "").strip(),
                "keys": keys,
                "aps_era": aps_era,
                "section": _section_from_source_path(source_path),
            })

            target_keys.update(keys)

    if duplicate_count:
        print(
            f"  Duplicate filtered rows skipped (same patent listed "
            f"in two sections, e.g. 2001 APS + SGML): {duplicate_count:,}"
        )

    return records, target_keys


# ============================================================
# SCAN DRIVER
# ============================================================

def discover_corpus_files(source=None):
    """
    The units a scan works through.

    "zip": every weekly .zip archive under RAW_DIR (the data
           members inside are read on the fly).
    "extracted": every data file under EXTRACTED_DIR -- anything
           that isn't an image, archive or housekeeping file,
           since APS-era files carry no useful extension.
    """

    source = source or SCAN_SOURCE

    if source == "zip":
        return sorted(RAW_DIR.rglob("*.zip"))

    files = []

    for file_path in EXTRACTED_DIR.rglob("*"):

        if not file_path.is_file():
            continue

        if file_path.name in IGNORED_FILE_NAMES:
            continue

        if file_path.suffix.lower() in IGNORED_FILE_SUFFIXES:
            continue

        files.append(file_path)

    return sorted(files)


def scan_corpus(sample_limit=None, source=None):

    source = source or SCAN_SOURCE

    print()
    print("=" * 78)
    print("BUILDING FORWARD-CITATION INDEX")
    print("=" * 78)

    records, target_keys = load_filtered_patents()

    print()
    print(f"  Filtered patents loaded : {len(records):,}")
    print(f"  Lookup keys registered  : {len(target_keys):,}")

    if not target_keys:
        raise SystemExit(
            "No usable patent IDs in filtered_patents.csv -- "
            "nothing to cross-reference."
        )

    corpus_root = RAW_DIR if source == "zip" else EXTRACTED_DIR

    if not corpus_root.exists():
        raise SystemExit(
            f"\nCorpus folder does not exist:\n  {corpus_root}\n\n"
            f"Run the main pipeline's --extract stage first."
        )

    connection = create_index_database()

    already_scanned = load_scanned_files(connection)

    print(f"  Previously scanned files: {len(already_scanned):,}")

    all_files = discover_corpus_files(source)

    print(f"  Reading from            : {corpus_root}  ({source})")

    pending = []

    for file_path in all_files:

        try:
            stat = file_path.stat()

        except OSError:
            continue

        previous = already_scanned.get(str(file_path))

        if previous and previous[0] == stat.st_size:
            continue

        pending.append(file_path)

    if sample_limit:
        pending = pending[:sample_limit]
        print(f"  SAMPLE MODE -- limiting to {sample_limit} file(s)")

    print(f"  Corpus files found      : {len(all_files):,}")
    print(f"  Files to scan this run  : {len(pending):,}")
    print()

    if not pending:
        print("  Nothing to scan. Index is already up to date.")
        connection.close()
        return

    cursor = connection.cursor()

    write_lock = threading.Lock()

    totals = {
        "records": 0,
        "citations_seen": 0,
        "hits": 0,
        "errors": 0,
        "formats": defaultdict(int),
    }

    def commit_result(result, stat_size, stat_mtime):

        if result["hits"]:

            cursor.executemany("""

                INSERT INTO forward_citations (
                    cited_key,
                    citing_key,
                    citing_number,
                    citing_date,
                    category,
                    source_format,
                    source_file
                )
                VALUES (?, ?, ?, ?, ?, ?, ?)

                -- The same citing patent can be read twice (2001
                -- APS + SGML overlap). Keep ONE row, preferring
                -- the copy that knows the citation category
                -- (SGML/XML) over APS's "unknown" -- independent
                -- of which file happened to finish first.
                ON CONFLICT(cited_key, citing_key) DO UPDATE SET
                    citing_number = excluded.citing_number,
                    citing_date   = excluded.citing_date,
                    category      = excluded.category,
                    source_format = excluded.source_format,
                    source_file   = excluded.source_file
                WHERE COALESCE(forward_citations.category, 'unknown')
                          IN ('unknown', '')
                  AND COALESCE(excluded.category, 'unknown')
                          NOT IN ('unknown', '')

            """, [
                (
                    cited_key,
                    citing_key,
                    citing_number,
                    citing_date,
                    category,
                    result["format"],
                    os.path.basename(result["path"]),
                )
                for cited_key, citing_number, citing_key,
                    citing_date, category
                in result["hits"]
            ])

        cursor.execute("""

            INSERT OR REPLACE INTO scanned_files (
                path, size, mtime, source_format,
                records, citations_seen, hits, error, scanned_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)

        """, (
            result["path"],
            stat_size,
            stat_mtime,
            result["format"],
            result["records"],
            result["citations_seen"],
            len(result["hits"]),
            result["error"],
            datetime.now().isoformat(timespec="seconds"),
        ))

        connection.commit()

    # Worker PROCESSES (not threads): scanning is mostly Python
    # work once the drive isn't the bottleneck, and threads can't
    # run Python code in parallel. This script doesn't import the
    # main pipeline, so Windows "spawn" re-import is cheap here.
    # The target-key set goes to each worker once, at startup.
    with ProcessPoolExecutor(
        max_workers=SCAN_WORKERS,
        initializer=_init_scan_worker,
        initargs=(target_keys,),
    ) as executor:

        future_to_path = {}

        for file_path in pending:

            future = executor.submit(scan_one_file, str(file_path))

            future_to_path[future] = file_path

        for future in tqdm(
            as_completed(future_to_path),
            total=len(future_to_path),
            desc="Scanning",
            unit="file",
        ):

            file_path = future_to_path[future]

            try:
                result = future.result()

            except Exception as error:
                result = {
                    "path": str(file_path),
                    "format": None,
                    "records": 0,
                    "citations_seen": 0,
                    "hits": [],
                    "error": f"{type(error).__name__}: {error}",
                }

            try:
                stat = file_path.stat()
                stat_size, stat_mtime = stat.st_size, stat.st_mtime

            except OSError:
                stat_size, stat_mtime = 0, 0.0

            with write_lock:
                commit_result(result, stat_size, stat_mtime)

            totals["records"] += result["records"]
            totals["citations_seen"] += result["citations_seen"]
            totals["hits"] += len(result["hits"])

            if result["error"]:
                totals["errors"] += 1

            if result["format"]:
                totals["formats"][result["format"]] += 1

    connection.close()

    print()
    print("-" * 78)
    print(f"  Patent records read     : {totals['records']:,}")
    print(f"  Cited references read   : {totals['citations_seen']:,}")
    print(f"  Hits on filtered set    : {totals['hits']:,}")
    print(f"  Files with errors       : {totals['errors']:,}")

    for file_format, count in sorted(totals["formats"].items()):
        print(f"    {file_format:<6} files scanned  : {count:,}")

    if totals["errors"]:
        print()
        print("  Files that errored are NOT marked complete and")
        print("  will be retried automatically on the next --scan.")

    print("-" * 78)


# ============================================================
# EXPORT
# ============================================================

def _chunk_citing_list(citing_numbers):
    """
    Excel refuses any cell over 32,767 characters. A heavily
    cited patent can easily blow past that (a foundational
    patent with 3,000 forward citations is ~36,000 characters
    of comma-separated IDs), and Excel's failure mode is
    truncation, not an error -- data would be silently lost.

    So the list is split across as many columns as it needs.
    """

    chunks = []
    current = []
    current_length = 0

    for number in citing_numbers:

        addition = len(number) + 2

        if current and current_length + addition > EXCEL_CELL_CHAR_BUDGET:
            chunks.append(", ".join(current))
            current = []
            current_length = 0

        current.append(number)
        current_length += addition

    if current:
        chunks.append(", ".join(current))

    return chunks


def _unique_citations(record, citing_by_key):

    """
    Yields (key, citing_number, citing_date, category) for every
    patent citing `record`, each citing patent once, across all
    of the record's lookup keys.
    """

    seen = set()

    for key in record["keys"]:

        for citing_number, citing_date, category in (
            citing_by_key.get(key, [])
        ):

            if citing_number in seen:
                continue

            seen.add(citing_number)

            yield key, citing_number, citing_date, category


class _SectionCsvWriters:

    """
    Writes forward_citation_pairs_<section>.csv files, starting a
    new "_part2", "_part3" ... file whenever one reaches
    EXCEL_PAIR_ROW_LIMIT rows, so each file opens in Excel
    without losing rows. close() returns [(path, rows), ...].
    """

    def __init__(self, header):

        self.header = header
        self.open_files = {}
        self.finished = []

    def _path(self, section, part):

        suffix = f"_{section}" if part == 1 else f"_{section}_part{part}"

        return FORWARD_CITATION_PAIRS_CSV.with_name(
            FORWARD_CITATION_PAIRS_CSV.stem + suffix
            + FORWARD_CITATION_PAIRS_CSV.suffix
        )

    def _start(self, section, part):

        path = self._path(section, part)

        handle = open(path, "w", newline="", encoding="utf-8-sig")

        writer = csv.writer(handle)

        writer.writerow(self.header)

        self.open_files[section] = [handle, writer, path, part, 0]

    def write(self, section, row):

        state = self.open_files.get(section)

        if state is None:

            self._start(section, 1)

        elif state[4] >= EXCEL_PAIR_ROW_LIMIT:

            state[0].close()

            self.finished.append((state[2], state[4]))

            self._start(section, state[3] + 1)

        state = self.open_files[section]

        state[1].writerow(row)

        state[4] += 1

    def close(self):

        for handle, _writer, path, _part, rows in self.open_files.values():

            handle.close()

            self.finished.append((path, rows))

        self.open_files = {}

        order = {name: index for index, (name, _) in enumerate(PAIR_SECTIONS)}

        def sort_key(item):
            stem = item[0].stem[len(FORWARD_CITATION_PAIRS_CSV.stem) + 1:]
            section, _, part = stem.partition("_part")
            return (order.get(section, 99), int(part or 1))

        return sorted(self.finished, key=sort_key)


def export_results():


    export_started = time.time()

    step_started = [export_started]

    def _step(label):

        now = time.time()

        print(f"    [{now - step_started[0]:6.1f}s] {label}")

        step_started[0] = now

    print()
    print("=" * 78)
    print("EXPORTING FORWARD-CITATION RESULTS")
    print("=" * 78)

    if not CITATION_INDEX_DB.exists():
        raise SystemExit(
            f"\nNo citation index found at:\n  {CITATION_INDEX_DB}\n\n"
            f"Run --scan first."
        )

    records, _ = load_filtered_patents()

    connection = get_index_connection()

    # Pull the whole index grouped by cited patent. This is
    # bounded by YOUR result set, not the corpus -- even a
    # 500,000-patent filter with an average of 10 forward
    # citations each is ~5M short strings, which is fine in
    # memory. The corpus-wide graph, which would not be, is
    # never materialized.
    citing_by_key = defaultdict(list)

    rows = connection.execute("""

        SELECT cited_key, citing_number, citing_key,
               citing_date, category
        FROM forward_citations
        ORDER BY cited_key, citing_key

    """)

    for cited_key, citing_number, citing_key, citing_date, category in rows:
        citing_by_key[cited_key].append((
            citing_number or citing_key,
            citing_date or "",
            category or "unknown",
        ))

    # --------------------------------------------------------
    # CORPUS COVERAGE END YEAR (for Citations Per Year)
    # --------------------------------------------------------
    #
    # The rate denominator needs to know how far FORWARD the
    # corpus actually reaches, so a patent's "years observed"
    # doesn't extend past the point where the corpus could no
    # longer have shown a citation even if one existed.
    #
    # An explicit CORPUS_COVERAGE_END_DATE is trusted first,
    # since you know your own --download/--extract range better
    # than anything derivable from the index. Otherwise, the
    # latest citing_date actually found anywhere in the index is
    # used as a proxy. This is queried against the FULL index,
    # not just hits on the target set, so a sparse filtered set
    # still gets a reasonable coverage estimate.

    if CORPUS_COVERAGE_END_DATE:

        coverage_end_year = _parse_year(CORPUS_COVERAGE_END_DATE)

    else:

        max_date_row = connection.execute(
            "SELECT MAX(citing_date) FROM forward_citations"
        ).fetchone()

        coverage_end_year = (
            _parse_year(max_date_row[0]) if max_date_row else None
        )

    if coverage_end_year is None:
        # Index is empty or every citing_date was unparseable --
        # fall back to today's year rather than leaving rates
        # entirely blank. This will UNDERSTATE how many years a
        # patent has actually been observable if the corpus is
        # smaller than "everything through today", so it's a
        # last resort, not the normal path.
        coverage_end_year = datetime.now().year

    connection.close()

    print()
    print(f"  Cited patents in index  : {len(citing_by_key):,}")
    print(f"  Corpus coverage through : {coverage_end_year}")

    if not CORPUS_COVERAGE_END_DATE:
        print(
            "    (auto-detected from latest citing_date in index -- "
            "set CORPUS_COVERAGE_END_DATE if you know the true "
            "extracted-corpus range for a more accurate rate)"
        )

    # --------------------------------------------------------
    # LONG-FORMAT CSV  (authoritative, unlimited)
    # --------------------------------------------------------

    FORWARD_CITATIONS_DIR.mkdir(parents=True, exist_ok=True)

    pair_count = 0

    pairs_header = [
        "filtered_patent_number",
        "filtered_patent_type",
        "filtered_claim_category",
        "canonical_key",
        "citing_patent_number",
        "citing_patent_type",
        "citing_patent_date",
        "citation_category",
        "filtered_patent_title",
    ]

    # Remove per-section files from an earlier export first, so a
    # stale "_part3" from a bigger previous run can't linger.
    for old_file in FORWARD_CITATIONS_DIR.glob(
        FORWARD_CITATION_PAIRS_CSV.stem + "_*" + FORWARD_CITATION_PAIRS_CSV.suffix
    ):
        try:
            old_file.unlink()
        except OSError:
            pass

    section_writers = _SectionCsvWriters(pairs_header)

    with open(
        FORWARD_CITATION_PAIRS_CSV,
        "w",
        newline="",
        encoding="utf-8-sig",
    ) as handle:

        writer = csv.writer(handle)

        writer.writerow(pairs_header)

        for record in records:

            for key, citing_number, citing_date, category in (
                _unique_citations(record, citing_by_key)
            ):

                row = [
                    record["patent_number"],
                    record["patent_type"],
                    record["claim_category"],
                    key,
                    citing_number,
                    classify_patent_type(citing_number),
                    citing_date,
                    category,
                    record["title"],
                ]

                writer.writerow(row)

                section_writers.write(record["section"], row)

                pair_count += 1

    section_csv_files = section_writers.close()

    print(f"  Citation pairs written  : {pair_count:,}")

    _step("index loaded + pairs CSVs written")

    # --------------------------------------------------------
    # EXCEL WORKBOOK
    # --------------------------------------------------------
    #
    # write_only mode: openpyxl's normal mode holds every cell
    # object in memory, which at this scale is gigabytes.
    # write_only streams rows straight to the file.
    # --------------------------------------------------------

    workbook = openpyxl.Workbook(write_only=True)

    summary = workbook.create_sheet("Forward Citations")

    max_citing_columns = 1

    summary_rows = []

    total_cited = 0
    total_uncited = 0

    for record in records:

        citing_numbers = []
        seen = set()

        # "applicant" is the strongest influence signal --
        # actively pointed to by whoever wrote the citing
        # patent, not just surfaced by an examiner's search.
        # "examiner" is the common case and mostly reflects
        # search-tool overlap. "other_unknown" folds together
        # SGML's non-examiner "other" (which may or may not be
        # applicant-sourced -- not assumed either way) and
        # every APS-era citation (that era carries no
        # per-reference provenance at all).
        examiner_count = 0
        applicant_count = 0
        other_unknown_count = 0

        for key in record["keys"]:
            for citing_number, citing_date, category in (
                citing_by_key.get(key, [])
            ):
                if citing_number not in seen:
                    seen.add(citing_number)
                    citing_numbers.append(citing_number)

                    if category == "examiner":
                        examiner_count += 1
                    elif category == "applicant":
                        applicant_count += 1
                    else:
                        other_unknown_count += 1

        citing_numbers.sort()

        count = len(citing_numbers)

        if count:
            total_cited += 1
        else:
            total_uncited += 1

        chunks = _chunk_citing_list(citing_numbers)

        max_citing_columns = max(max_citing_columns, len(chunks) or 1)

        # --------------------------------------------------
        # RATE NORMALIZATION
        # --------------------------------------------------
        #
        # grant_year comes from grant_date, falling back to
        # publication_date for record types (design, plant)
        # where the CSV may only carry one of the two.
        # Anything unparseable leaves rate/percentile blank
        # rather than silently computing against a wrong year.

        grant_year = _parse_year(
            record["grant_date"] or record["publication_date"]
        )

        years_observed = None
        citations_per_year = None

        if grant_year is not None:

            years_observed = max(
                CITATION_RATE_MIN_YEARS,
                coverage_end_year - grant_year,
            )

            citations_per_year = (
                round(count / years_observed, 3)
                if years_observed
                else None
            )

        summary_rows.append({
            "record": record,
            "count": count,
            "chunks": chunks,
            "examiner_count": examiner_count,
            "applicant_count": applicant_count,
            "other_unknown_count": other_unknown_count,
            "grant_year": grant_year,
            "cohort": _cohort_bucket(grant_year),
            "years_observed": years_observed,
            "citations_per_year": citations_per_year,
            "cohort_percentile": None,  # filled in below
        })

    # --------------------------------------------------------
    # COHORT PERCENTILE
    # --------------------------------------------------------
    #
    # Ranks each patent's raw citation count against other
    # FILTERED patents sharing its grant-year cohort bucket.
    # This is meaningful only WITHIN your filtered set, not
    # against the whole corpus -- building a corpus-wide
    # percentile would require indexing every citation edge in
    # USPTO, not just the ones landing on your target set,
    # which is the exact scale problem the target-key filter
    # in scan_one_file() exists to avoid.
    #
    # A cohort smaller than CITATION_COHORT_MIN_SIZE is left
    # blank -- a percentile computed against 2 or 3 peers reads
    # as precise but isn't.

    cohort_members = defaultdict(list)

    for row in summary_rows:
        if row["cohort"] is not None:
            cohort_members[row["cohort"]].append(row)

    for bucket, members in cohort_members.items():

        if len(members) < CITATION_COHORT_MIN_SIZE:
            continue

        sorted_counts = sorted(m["count"] for m in members)
        cohort_size = len(sorted_counts)

        # rank = how many patents in the cohort have a count <=
        # this one. bisect_right finds that in one O(log n)
        # lookup; v1-v3 counted it with a full pass per patent,
        # which is O(n^2) per cohort -- billions of steps for a
        # large filtered set. Results are identical.

        for row in members:
            # Percentile = share of the cohort at or below this
            # patent's count. Using "<=" (not "<") means a
            # patent tied with the whole cohort still lands at
            # the 100th percentile rather than being pushed
            # below patents it's actually tied with.
            rank = bisect.bisect_right(sorted_counts, row["count"])
            row["cohort_percentile"] = round(
                100 * rank / cohort_size, 1
            )

    header = [
        "Filtered Patent ID",
        "Title",
        "Patent Type",
        "Claim Category",
        "Independent Claim Categories",
        "Grant Date",
        "Assignee",
        "CPC",
        "Times Cited",
        "Citations Per Year",
        "Cohort Percentile",
        "Cited by Applicant",
        "Cited by Examiner",
        "Cited by Other/Unknown",
    ]

    for index in range(max_citing_columns):
        header.append(
            "Citing Patent IDs"
            if max_citing_columns == 1
            else f"Citing Patent IDs ({index + 1})"
        )

    _step("counts, rates and cohort percentiles")

    summary.append(header)

    # Most-cited first: the whole point of this output is to
    # surface which of your filtered patents the field actually
    # built on, so the interesting rows should be at the top
    # rather than buried alphabetically.
    summary_rows.sort(key=lambda item: -item["count"])

    for row in summary_rows:

        record = row["record"]

        out = [
            record["patent_number"],
            record["title"],
            record["patent_type"],
            record["claim_category"],
            record["independent_claim_categories"],
            record["grant_date"] or record["publication_date"],
            record["assignees"],
            record["cpc"],
            row["count"],
            (
                row["citations_per_year"]
                if row["citations_per_year"] is not None
                else ""
            ),
            (
                row["cohort_percentile"]
                if row["cohort_percentile"] is not None
                else ""
            ),
            row["applicant_count"],
            row["examiner_count"],
            row["other_unknown_count"],
        ]

        for index in range(max_citing_columns):
            chunks = row["chunks"]
            out.append(chunks[index] if index < len(chunks) else "")

        summary.append(out)

    # --------------------------------------------------------
    # PAIRS SHEET
    # --------------------------------------------------------

    _step("Excel Forward Citations sheet")

    pairs_written = 0

    # (section label, sheet name, rows) for each pairs sheet, for
    # the Run Notes sheet and the console summary.
    pairs_sheets_written = []

    if EXCEL_INCLUDE_PAIRS_SHEET:

        pairs_header = [
            "Filtered Patent ID",
            "Filtered Patent Type",
            "Filtered Claim Category",
            "Citing Patent ID",
            "Citing Patent Type",
            "Citing Patent Date",
            "Citation Category",
        ]

        for section, label in PAIR_SECTIONS:

            sheet = None
            sheet_number = 0
            sheet_rows = 0

            for row_data in summary_rows:

                record, count = row_data["record"], row_data["count"]

                if not count or record["section"] != section:
                    continue

                for _key, citing_number, citing_date, category in (
                    _unique_citations(record, citing_by_key)
                ):

                    if sheet is None or sheet_rows >= EXCEL_PAIR_ROW_LIMIT:

                        if sheet is not None:
                            pairs_sheets_written.append(
                                (label, sheet_name, sheet_rows)
                            )

                        sheet_number += 1

                        sheet_name = (
                            f"Pairs - {label}"
                            if sheet_number == 1
                            else f"Pairs - {label} ({sheet_number})"
                        )

                        sheet = workbook.create_sheet(sheet_name)

                        sheet.append(pairs_header)

                        sheet_rows = 0

                    sheet.append([
                        record["patent_number"],
                        record["patent_type"],
                        record["claim_category"],
                        citing_number,
                        classify_patent_type(citing_number),
                        citing_date,
                        category,
                    ])

                    sheet_rows += 1
                    pairs_written += 1

            if sheet is not None:
                pairs_sheets_written.append((label, sheet_name, sheet_rows))

        _step(f"Excel pairs sheets ({len(pairs_sheets_written)})")

    # --------------------------------------------------------
    # RUN NOTES SHEET
    # --------------------------------------------------------
    #
    # A results file that doesn't say what it covers is a
    # results file someone will misread six months from now.
    # --------------------------------------------------------

    notes = workbook.create_sheet("Run Notes")

    notes.append(["Generated", datetime.now().isoformat(timespec="seconds")])
    notes.append(["Source filter file", str(FILTER_RESULTS)])
    notes.append(["Citation index", str(CITATION_INDEX_DB)])
    notes.append(["Filtered patents", len(records)])
    notes.append(["Patents cited at least once", total_cited])
    notes.append(["Patents with zero citations found", total_uncited])
    notes.append(["Total citation pairs", pair_count])
    notes.append([
        "Corpus coverage through (year)", coverage_end_year
    ])
    notes.append([
        "Coverage year source",
        (
            "CORPUS_COVERAGE_END_DATE (manually set)"
            if CORPUS_COVERAGE_END_DATE
            else "auto-detected: latest citing_date found in index"
        ),
    ])
    notes.append([
        "Cohort bucket width (years)", CITATION_COHORT_YEAR_WINDOW
    ])
    notes.append([
        "Minimum cohort size to report a percentile",
        CITATION_COHORT_MIN_SIZE,
    ])
    notes.append([])
    notes.append(["IMPORTANT COVERAGE NOTES"])
    notes.append([
        "Counts reflect only the corpus scanned (RAW zips by default, or EXTRACTED)."
    ])
    notes.append([
        "A patent can only be cited by documents published after it."
    ])
    notes.append([
        "US patent-document citations only; foreign citing families "
        "are not in USPTO bulk data."
    ])
    notes.append([
        "Pre-grant applications often carry no references-cited "
        "section, so citing documents are mostly grants."
    ])
    notes.append([
        "'Cited by Applicant' is the strongest influence signal -- "
        "the citing document actively pointed to that prior art, "
        "rather than an examiner's search surfacing it."
    ])
    notes.append([
        "'Cited by Other/Unknown' includes: SGML-era 'other' "
        "citations (may or may not be applicant-sourced -- not "
        "assumed either way), and ALL APS-era (pre-1996) citations, "
        "since that format carries no per-reference examiner/"
        "applicant marker at all."
    ])
    notes.append([])
    notes.append(["CITATION RATE NOTES"])
    notes.append([
        "'Citations Per Year' = Times Cited / years the patent has "
        "sat inside the corpus's observable window (grant year to "
        "the coverage year above, floored at "
        f"{CITATION_RATE_MIN_YEARS} year so very recent grants "
        "don't produce inflated rates)."
    ])
    notes.append([
        "'Cohort Percentile' ranks each patent's raw citation count "
        "against OTHER FILTERED patents granted in the same "
        f"{CITATION_COHORT_YEAR_WINDOW}-year window -- meaningful "
        "only within this filtered set, not against all of USPTO, "
        "since only citations landing on your filtered patents are "
        "indexed (by design, to keep the index proportional to "
        "your results rather than the whole corpus)."
    ])
    notes.append([
        "Both columns are blank where a patent's grant/publication "
        "date couldn't be parsed, or (Cohort Percentile only) where "
        f"fewer than {CITATION_COHORT_MIN_SIZE} filtered patents "
        "share that grant-year cohort."
    ])
    notes.append([
        "The citation-arrival curve is itself skewed -- most "
        "patents peak roughly 3-6 years after grant, not on a "
        "straight line -- so a flat per-year rate is an "
        "approximation, not a true accumulation model."
    ])
    notes.append([
        "Citing Patent Date (Citation Pairs sheet / pairs CSV) is "
        "the CITING patent's own grant/publication date -- it's "
        "the raw input the rate above is computed from, and is "
        "there directly if you want to build your own curve "
        "(e.g. citations binned by year-since-grant) instead."
    ])

    notes.append([])
    notes.append(["PATENT TYPE AND CLAIM CATEGORY"])
    notes.append([
        "'Patent Type' is the USPTO document type, read from the "
        "number/kind code (Utility, Design, Plant, Reissue, SIR, "
        "Defensive Publication, Application)."
    ])
    notes.append([
        "'Claim Category' is what claim 1 claims -- Process / Method, "
        "Apparatus / System, Product / Article, Composition, "
        "Product-by-process, Design, Plant -- read from the claim's "
        "preamble by the main pipeline (a heuristic: USPTO does not "
        "record it). 'Independent Claim Categories' lists every "
        "category across the independent claims."
    ])
    notes.append([
        "'Citing Patent Type' (pairs) is read from the citing "
        "patent's number only."
    ])

    category_counts = defaultdict(int)
    category_cited = defaultdict(int)

    for row in summary_rows:
        label = row["record"]["claim_category"] or "(blank)"
        category_counts[label] += 1
        category_cited[label] += row["count"]

    if set(category_counts) == {"(blank)"}:
        notes.append([
            "Claim Category is blank: filtered_patents.csv predates "
            "pipeline v15.10. Run the pipeline's --classify, then "
            "--export here again."
        ])
    else:
        notes.append(["Claim category", "Filtered patents", "Total times cited"])
        for label, count in sorted(
            category_counts.items(), key=lambda item: -item[1]
        ):
            notes.append([label, count, category_cited[label]])

    if not EXCEL_INCLUDE_PAIRS_SHEET:
        notes.append([])
        notes.append([
            "Citation Pairs sheet omitted (EXCEL_INCLUDE_PAIRS_SHEET / "
            "--skip-excel-pairs). The pairs CSV has the complete set."
        ])

    if pairs_sheets_written:
        notes.append([])
        notes.append(["CITATION PAIRS SHEETS (split by section)"])
        notes.append([
            f"Each sheet holds up to {EXCEL_PAIR_ROW_LIMIT:,} pairs; a "
            f"larger section continues on a '(2)', '(3)' sheet. Within "
            f"a section, patents are listed most-cited first. Every "
            f"pair is included."
        ])
        for label, sheet_name, rows in pairs_sheets_written:
            notes.append([sheet_name, rows])

    _step("Excel summary + notes rows")

    workbook.save(FORWARD_CITATION_XLSX)

    _step("Excel workbook saved")

    print(f"  Export took {time.time() - export_started:,.1f}s")

    rated_count = sum(
        1 for row in summary_rows if row["citations_per_year"] is not None
    )

    percentiled_count = sum(
        1 for row in summary_rows if row["cohort_percentile"] is not None
    )

    print(f"  Patents with >=1 citation: {total_cited:,}")
    print(f"  Patents with 0 citations : {total_uncited:,}")
    print(f"  Patents with a computed rate       : {rated_count:,}")
    print(f"  Patents with a cohort percentile   : {percentiled_count:,}")
    print()
    print("  Written:")
    print(f"    {FORWARD_CITATION_XLSX}")
    print(f"    {FORWARD_CITATION_PAIRS_CSV}  (all sections)")

    for path, rows in section_csv_files:
        print(f"    {path.name}  ({rows:,} pairs)")

    if pairs_sheets_written:
        print()
        print("  Excel pairs sheets:")
        for label, sheet_name, rows in pairs_sheets_written:
            print(f"    {sheet_name:<30}{rows:>12,} pairs")


# ============================================================
# VERIFY
# ============================================================

def verify():
    """
    Reports how the matching actually resolved, so the APS
    check-digit handling (and coverage generally) is something
    you can confirm against real data rather than take on
    faith.
    """

    print()
    print("=" * 78)
    print("FORWARD-CITATION MATCH REPORT")
    print("=" * 78)

    records, target_keys = load_filtered_patents()

    if not CITATION_INDEX_DB.exists():
        raise SystemExit(
            f"\nNo citation index found at:\n  {CITATION_INDEX_DB}\n\n"
            f"Run --scan first."
        )

    connection = get_index_connection()

    counts_by_key = dict(connection.execute("""

        SELECT cited_key, COUNT(*)
        FROM forward_citations
        GROUP BY cited_key

    """).fetchall())

    scan_stats = connection.execute("""

        SELECT
            COUNT(*),
            SUM(records),
            SUM(citations_seen),
            SUM(hits),
            SUM(CASE WHEN error IS NOT NULL AND error != ''
                     THEN 1 ELSE 0 END)
        FROM scanned_files

    """).fetchone()

    format_stats = connection.execute("""

        SELECT source_format, COUNT(*), SUM(records)
        FROM scanned_files
        GROUP BY source_format

    """).fetchall()

    category_stats = connection.execute("""

        SELECT category, COUNT(*)
        FROM forward_citations
        GROUP BY category
        ORDER BY COUNT(*) DESC

    """).fetchall()

    errors = connection.execute("""

        SELECT path, error
        FROM scanned_files
        WHERE error IS NOT NULL AND error != ''
        LIMIT 20

    """).fetchall()

    connection.close()

    print()
    print("CORPUS COVERAGE")
    print(f"  Files scanned           : {scan_stats[0] or 0:,}")
    print(f"  Patent records read     : {scan_stats[1] or 0:,}")
    print(f"  Cited references read   : {scan_stats[2] or 0:,}")
    print(f"  Hits on filtered set    : {scan_stats[3] or 0:,}")
    print(f"  Files with errors       : {scan_stats[4] or 0:,}")

    print()
    print("BY FORMAT")

    for source_format, file_count, record_count in format_stats:
        print(
            f"  {str(source_format or 'unrecognized'):<14} "
            f"{file_count:>8,} files  "
            f"{record_count or 0:>12,} records"
        )

    print()
    print("CITATIONS BY CATEGORY (across all hits on your filtered set)")

    total_category_hits = sum(count for _, count in category_stats) or 1

    for category, count in category_stats:
        pct = 100 * count / total_category_hits
        print(f"  {str(category or 'unknown'):<12} {count:>10,}  ({pct:.1f}%)")

    print()
    print("  'examiner' citations mostly reflect the examiner's own")
    print("  search overlap, not the field's use of the invention.")
    print("  'applicant' citations are the stronger influence signal.")
    print("  All APS-era (pre-1996) citations report as 'unknown' --")
    print("  that format has no per-reference provenance marker.")

    # --------------------------------------------------------
    # KEY RESOLUTION -- the APS check-digit question
    # --------------------------------------------------------

    primary_hits = 0
    alternate_hits = 0
    both_hits = 0
    no_hits = 0
    ambiguous = []

    FORWARD_CITATIONS_DIR.mkdir(parents=True, exist_ok=True)

    with open(
        FORWARD_CITATION_VERIFY_CSV,
        "w",
        newline="",
        encoding="utf-8-sig",
    ) as handle:

        writer = csv.writer(handle)

        writer.writerow([
            "filtered_patent_number",
            "aps_era",
            "primary_key",
            "primary_key_hits",
            "alternate_key",
            "alternate_key_hits",
            "resolution",
        ])

        for record in records:

            keys = record["keys"]

            primary = keys[0]
            alternate = keys[1] if len(keys) > 1 else ""

            primary_count = counts_by_key.get(primary, 0)
            alternate_count = (
                counts_by_key.get(alternate, 0) if alternate else 0
            )

            if primary_count and alternate_count:
                resolution = "BOTH (ambiguous)"
                both_hits += 1
                ambiguous.append(record["patent_number"])

            elif primary_count:
                resolution = "primary"
                primary_hits += 1

            elif alternate_count:
                resolution = "alternate (check digit stripped)"
                alternate_hits += 1

            else:
                resolution = "no citations found"
                no_hits += 1

            writer.writerow([
                record["patent_number"],
                "yes" if record["aps_era"] else "no",
                primary,
                primary_count,
                alternate,
                alternate_count,
                resolution,
            ])

    print()
    print("KEY RESOLUTION")
    print(f"  Matched on primary key  : {primary_hits:,}")
    print(f"  Matched on alternate key: {alternate_hits:,}")
    print(f"  Matched on BOTH         : {both_hits:,}")
    print(f"  No citations found      : {no_hits:,}")

    if alternate_hits:
        print()
        print("  The alternate-key matches are APS-era patents whose")
        print("  stored WKU number carries a check digit. That the")
        print("  stripped form matches real citations is the evidence")
        print("  the check-digit handling is correct.")

    if both_hits:
        print()
        print(f"  {both_hits:,} patents matched BOTH keys -- these are")
        print("  genuinely ambiguous and their counts may be inflated.")
        print("  Spot-check a few against Google Patents:")

        for number in ambiguous[:5]:
            print(f"    {number}")

    if errors:
        print()
        print("FILES WITH ERRORS (first 20)")

        for path, error in errors:
            print(f"  {os.path.basename(path)}: {error}")

    print()
    print(f"  Full report: {FORWARD_CITATION_VERIFY_CSV}")


# ============================================================
# RESET
# ============================================================

def reset_index():

    if not CITATION_INDEX_DB.exists():
        print("No citation index to delete.")
        return

    answer = input(
        f"Delete {CITATION_INDEX_DB}?\n"
        f"The next --scan will re-read the entire corpus. [y/N] "
    )

    if answer.strip().lower() != "y":
        print("Cancelled.")
        return

    for suffix in ["", "-wal", "-shm"]:

        target = Path(str(CITATION_INDEX_DB) + suffix)

        if target.exists():
            target.unlink()

    print("Citation index deleted.")


# ============================================================
# COMMAND LINE
# ============================================================

def main():

    global SCAN_WORKERS, EXCEL_INCLUDE_PAIRS_SHEET

    parser = argparse.ArgumentParser(
        description=(
            "Find which patents cite each patent in "
            "filtered_patents.csv"
        )
    )

    parser.add_argument(
        "--scan",
        action="store_true",
        help=(
            "Scan the corpus and build the citation index "
            "(restartable)"
        ),
    )

    parser.add_argument(
        "--export",
        action="store_true",
        help="Write the Excel workbook and CSV from the index",
    )

    parser.add_argument(
        "--all",
        action="store_true",
        help="--scan then --export",
    )

    parser.add_argument(
        "--sample",
        type=int,
        metavar="N",
        help=(
            "Scan only the first N unscanned files. Do this "
            "before committing to a full scan."
        ),
    )

    parser.add_argument(
        "--verify",
        action="store_true",
        help=(
            "Report coverage and how patent-number matching "
            "resolved. No scanning."
        ),
    )

    parser.add_argument(
        "--source",
        choices=["zip", "extracted"],
        default=None,
        help=(
            f"Read the weekly zip archives (zip, default: "
            f"{SCAN_SOURCE}) or the extracted files (extracted)."
        ),
    )

    parser.add_argument(
        "--workers",
        type=int,
        default=None,
        metavar="N",
        help=f"Scanning processes (default {SCAN_WORKERS}).",
    )

    parser.add_argument(
        "--skip-excel-pairs",
        action="store_true",
        help=(
            "Leave the Citation Pairs sheet out of the Excel "
            "workbook (the pairs CSV still has every pair). "
            "Fastest --export."
        ),
    )

    parser.add_argument(
        "--reset",
        action="store_true",
        help="Delete the citation index and start over",
    )

    args = parser.parse_args()

    if args.skip_excel_pairs:
        EXCEL_INCLUDE_PAIRS_SHEET = False

    if args.workers:
        SCAN_WORKERS = args.workers

    if args.reset:
        reset_index()
        return

    if args.verify:
        verify()
        return

    if args.all:
        scan_corpus(sample_limit=args.sample, source=args.source)
        export_results()
        return

    if args.scan or args.sample:
        scan_corpus(sample_limit=args.sample, source=args.source)

    if args.export:
        export_results()

    if not any([
        args.scan, args.export, args.all,
        args.sample, args.verify, args.reset,
    ]):
        parser.print_help()


if __name__ == "__main__":
    main()


# ============================================================
# WIRING THIS INTO THE MAIN PIPELINE (optional)
# ============================================================
#
# If you'd rather run this as a stage of
# uspto_patent_pipeline_v15.py than as a separate script, add
# to that file's main():
#
#     parser.add_argument(
#         "--forward-citations",
#         action="store_true",
#         help=(
#             "Cross-reference filtered patents against the "
#             "whole corpus to find which patents cite them"
#         ),
#     )
#
#     if args.forward_citations:
#         import uspto_forward_citations_v5
#         uspto_forward_citations_v5.scan_corpus()
#         uspto_forward_citations_v5.export_results()
#
# Put both files in the same folder. Note this is a DIFFERENT
# thing from the existing --citations stage: that one asks
# PatentsView's API for a total citation count (all of USPTO,
# but an unconfirmed endpoint and no list of who). This one
# reads your local raw data, is fully offline and verifiable,
# and gives you the actual citing patent IDs -- limited to
# whatever you've extracted locally. They complement each
# other: --citations for a true total, this for the identities.
#
# In --all, place it after filter_patents() and it can run
# either before or after download_selected_pdfs() -- it doesn't
# touch PDFs.
