"""
RESEARCH OVERVIEW REPORT  (v1)
==============================

Builds one overview folder that sits OUTSIDE both corpora and
summarizes them:

    D:\\OVERVIEW\\
        overview_report.xlsx       the data-analysis workbook
        FORWARD CITATIONS\\         forward_citations.xlsx plus the
                                   per-section pairs CSVs (each one
                                   fits in Excel -- the combined
                                   forward_citation_pairs.csv, which
                                   can pass Excel's row limit, is
                                   left out)
        FILTERED PATENTS\\          filtered_patents_<product>.csv
                                   (aps / sgml / xml / appxml) --
                                   the combined filtered_patents.csv
                                   is left out

overview_report.xlsx sheets:

    Patents   bar chart: patents by decade, stacked by claim
              category; bar chart: patents by claim category; bar
              chart: patents by file type (APS / SGML / XML grants /
              Applications); table: top 10 patents by forward
              citations
    Crawl     vendor data-sheet crawl: headline numbers; charts of
              documents by material category, by document type,
              top 15 vendors, and test-method families
    Vendors   one row per vendor: pages crawled, PDF links found,
              documents kept by type, translations, rejected, size
    Notes     what was read, when, and how the numbers were counted

Reads only -- nothing in D:\\uspto_corpus or D:\\tds_corpus is
changed. Safe to re-run any time: the report is rebuilt and only
new or changed files are copied.

Usage:
    python overview_report_v1.py              report + copy files
    python overview_report_v1.py --no-copy    report only
    python overview_report_v1.py --copy-only  copy files only
"""

import argparse
import csv
import re
import shutil
import sqlite3
import sys
import time
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# ============================================================
# CONFIG
# ============================================================

USPTO_ROOT = Path(r"D:\uspto_corpus")

TDS_ROOT = Path(r"D:\tds_corpus")

OVERVIEW_ROOT = Path(r"D:\OVERVIEW")

REPORT_NAME = "overview_report.xlsx"


# Inputs (match uspto_patent_pipeline_v15.py / forward citations v5.1)
FILTERED_DIR = USPTO_ROOT / "RESULTS" / "FILTERED"
FORWARD_DIR = USPTO_ROOT / "RESULTS" / "FORWARD CITATIONS"

TDS_MANIFEST = TDS_ROOT / "tds_manifest.csv"
TDS_DB = TDS_ROOT / "tds_crawler.db"

# Outputs
OVERVIEW_FORWARD_DIR = OVERVIEW_ROOT / "FORWARD CITATIONS"
OVERVIEW_FILTERED_DIR = OVERVIEW_ROOT / "FILTERED PATENTS"


PRODUCTS = [
    ("aps", "APS grants (1976-2001)"),
    ("sgml", "SGML grants (2001)"),
    ("xml", "XML grants (2002-present)"),
    ("appxml", "Applications (2001-present)"),
]

CLAIM_CATEGORY_ORDER = [
    "Process / Method",
    "Apparatus / System",
    "Product / Article",
    "Composition",
    "Product-by-process",
    "Design",
    "Plant",
    "Unclassified",
    "No claims text",
]

NOT_CLASSIFIED = "Not classified (run --classify)"

TOP_N_PATENTS = 10

TOP_N_VENDORS = 15

DOC_TYPE_LABEL = {
    "tds": "Technical data sheet",
    "sds": "Safety data sheet",
    "brochure": "Brochure / selector guide",
    "certificate": "Certificate / statement",
    "corporate": "Corporate document",
    "commercial": "Commercial / supplier document",
    "other": "Technical document",
    "unreadable": "Scanned / image-only PDF",
}


csv.field_size_limit(min(sys.maxsize, 2**31 - 1))


# ============================================================
# EXCEL HELPERS
# ============================================================

HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="2F4F6F")
TITLE_FONT = Font(bold=True, size=14)
SECTION_FONT = Font(bold=True, size=12)


def _write_title(sheet, row, text, font=SECTION_FONT):

    sheet.cell(row=row, column=1, value=text).font = font


def _write_table(sheet, top, left, header, rows, number_format="#,##0"):

    """Writes header + rows starting at (top, left). Returns the
    last row written."""

    for offset, label in enumerate(header):
        cell = sheet.cell(row=top, column=left + offset, value=label)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    for row_index, values in enumerate(rows, start=1):
        for offset, value in enumerate(values):
            cell = sheet.cell(row=top + row_index, column=left + offset, value=value)
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                cell.number_format = number_format

    return top + len(rows)


def _bar_chart(
    sheet, title, data_top, data_left, data_bottom, series_count,
    anchor, x_title=None, y_title="Count", stacked=False,
    horizontal=False, width=22, height=10,
):

    """Bar chart over a table whose first column holds the
    categories and whose next `series_count` columns hold the
    series (header row = series names)."""

    chart = BarChart()
    chart.type = "bar" if horizontal else "col"
    chart.title = title
    chart.width = width
    chart.height = height
    chart.y_axis.title = y_title
    if x_title:
        chart.x_axis.title = x_title

    if stacked:
        chart.grouping = "stacked"
        chart.overlap = 100
    else:
        chart.legend = None if series_count == 1 else chart.legend

    data = Reference(
        sheet,
        min_col=data_left + 1,
        max_col=data_left + series_count,
        min_row=data_top,
        max_row=data_bottom,
    )
    categories = Reference(
        sheet,
        min_col=data_left,
        min_row=data_top + 1,
        max_row=data_bottom,
    )

    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    # openpyxl defaults hide the axes in recent Excel builds.
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.scaling.min = 0  # bars start at zero

    if horizontal:
        chart.x_axis.scaling.orientation = "maxMin"

    sheet.add_chart(chart, anchor)


def _set_widths(sheet, widths):

    for column, width in widths.items():
        sheet.column_dimensions[column].width = width


# ============================================================
# PATENTS
# ============================================================

def _year(*values):

    for value in values:
        digits = (value or "").strip()[:4]
        if digits.isdigit() and 1790 <= int(digits) <= 2100:
            return int(digits)

    return None


def _grant_key(number, product):

    """
    Normalized grant number so the same patent listed in both the
    2001 APS and SGML files is counted once. Applications keep
    their own key (an application and its later grant are two
    documents).
    """

    value = (number or "").strip().upper()

    if value.startswith("US"):
        value = value[2:]

    if product == "aps" and len(value) == 9:
        value = value[:-1]  # Green Book check character

    value = re.sub(r"[^A-Z0-9]", "", value)

    match = re.match(r"^([A-Z]*)(\d+)", value)

    if not match:
        return f"{product}:{value}"

    prefix, digits = match.group(1), match.group(2)

    if product == "appxml":
        return f"app:{digits}"

    return f"{prefix}{digits.lstrip('0') or '0'}"


def load_filtered_patents():

    """
    Reads every filtered_patents_<product>.csv. Returns
    (patents_by_number, file_type_counts, duplicates_removed,
     classified) where patents_by_number maps patent_number (as
    stored) -> row summary, de-duplicated across products.
    """

    patents = {}
    seen_keys = set()
    file_counts = []
    duplicates = 0
    classified = False

    for product, label in PRODUCTS:

        path = FILTERED_DIR / f"filtered_patents_{product}.csv"

        if not path.exists():
            file_counts.append((label, 0))
            print(f"  {path.name}: not found")
            continue

        count = 0

        with open(path, "r", newline="", encoding="utf-8-sig") as handle:

            reader = csv.DictReader(handle)

            if "claim_category" in (reader.fieldnames or []):
                classified = True

            for row in reader:

                count += 1

                number = (row.get("patent_number") or row.get("publication_number") or "").strip()

                key = _grant_key(number, product)

                if key in seen_keys:
                    duplicates += 1
                    continue

                seen_keys.add(key)

                patents[number] = {
                    "number": number,
                    "product": product,
                    "title": (row.get("title") or "").strip(),
                    "assignee": (row.get("assignees") or row.get("applicants") or "").strip(),
                    "year": _year(
                        row.get("grant_date"),
                        row.get("publication_date"),
                        row.get("filing_date"),
                    ),
                    "patent_type": (row.get("patent_type") or "").strip(),
                    "claim_category": (row.get("claim_category") or "").strip() or NOT_CLASSIFIED,
                }

        file_counts.append((label, count))

        print(f"  {path.name}: {count:,} rows")

    return patents, file_counts, duplicates, classified


def load_forward_citation_counts():

    """
    Counts citing patents per filtered patent from the pairs CSV
    (one row per unique citing/cited pair). Uses the combined
    forward_citation_pairs.csv, or the per-section files if the
    combined one isn't there.
    """

    combined = FORWARD_DIR / "forward_citation_pairs.csv"

    if combined.exists():
        files = [combined]
    else:
        files = sorted(FORWARD_DIR.glob("forward_citation_pairs_*.csv"))

    totals = Counter()
    by_category = defaultdict(Counter)

    for path in files:

        print(f"  reading {path.name} ...")

        with open(path, "r", newline="", encoding="utf-8-sig") as handle:

            reader = csv.reader(handle)
            header = next(reader, [])

            try:
                cited_col = header.index("filtered_patent_number")
            except ValueError:
                continue

            category_col = (
                header.index("citation_category")
                if "citation_category" in header
                else None
            )

            for row in reader:

                if len(row) <= cited_col:
                    continue

                number = row[cited_col]

                totals[number] += 1

                if category_col is not None and len(row) > category_col:
                    by_category[number][row[category_col]] += 1

    return totals, by_category, [path.name for path in files]


def build_patents_sheet(workbook, notes):

    print("\nPATENTS")

    patents, file_counts, duplicates, classified = load_filtered_patents()

    sheet = workbook.create_sheet("Patents")

    _write_title(sheet, 1, "USPTO filtered patent corpus", TITLE_FONT)

    sheet.cell(
        row=2, column=1,
        value=(
            f"{len(patents):,} unique patents and applications "
            f"({sum(count for _, count in file_counts):,} filter-result rows; "
            f"{duplicates:,} listed in two files counted once)"
        ),
    )

    # --- categories present, in a fixed order -------------------

    category_counts = Counter(p["claim_category"] for p in patents.values())

    categories = [c for c in CLAIM_CATEGORY_ORDER if category_counts.get(c)]
    categories += sorted(c for c in category_counts if c not in categories)

    # --- 1. patents by decade, stacked by claim category -------

    decade_counts = defaultdict(Counter)
    undated = 0

    for patent in patents.values():
        if patent["year"] is None:
            undated += 1
            continue
        decade_counts[patent["year"] // 10 * 10][patent["claim_category"]] += 1

    row = 4
    _write_title(sheet, row, "Patents by decade and claim category")

    table_top = row + 1
    decade_rows = [
        [f"{decade}s"] + [decade_counts[decade].get(c, 0) for c in categories]
        for decade in sorted(decade_counts)
    ]
    table_bottom = _write_table(sheet, table_top, 1, ["Decade"] + categories, decade_rows)

    _bar_chart(
        sheet, "Patents by decade (by claim category)",
        table_top, 1, table_bottom, len(categories),
        anchor=f"{get_column_letter(len(categories) + 3)}{table_top}",
        x_title="Decade (grant date; publication date for applications)",
        y_title="Patents", stacked=True, width=26, height=12,
    )

    row = max(table_bottom, table_top + 24) + 2

    # --- 2. patents by claim category ---------------------------

    _write_title(sheet, row, "Patents by claim category")

    table_top = row + 1
    table_bottom = _write_table(
        sheet, table_top, 1, ["Claim category", "Patents"],
        [[c, category_counts[c]] for c in categories],
    )

    _bar_chart(
        sheet, "Patents by claim category",
        table_top, 1, table_bottom, 1,
        anchor=f"D{table_top}", y_title="Patents", width=22, height=10,
    )

    row = max(table_bottom, table_top + 20) + 2

    # --- 3. patents by file type --------------------------------

    _write_title(sheet, row, "Filter results by file type")

    table_top = row + 1
    table_bottom = _write_table(
        sheet, table_top, 1, ["File type", "Patents"],
        [[label, count] for label, count in file_counts],
    )

    _bar_chart(
        sheet, "Patents by file type",
        table_top, 1, table_bottom, 1,
        anchor=f"D{table_top}", y_title="Patents", width=22, height=10,
    )

    row = max(table_bottom, table_top + 20) + 2

    # --- 4. top 10 by forward citations -------------------------

    print("\nFORWARD CITATIONS")

    totals, by_category, pair_files = load_forward_citation_counts()

    _write_title(sheet, row, f"Top {TOP_N_PATENTS} patents by forward citations")

    header = [
        "Rank", "Patent", "Title", "Assignee", "Year", "Patent type",
        "Claim category", "Times cited", "Cited by applicant",
        "Cited by examiner", "Cited by other / unknown",
    ]

    top_rows = []

    for rank, (number, count) in enumerate(totals.most_common(TOP_N_PATENTS), start=1):

        patent = patents.get(number, {})
        split = by_category.get(number, Counter())

        top_rows.append([
            rank,
            number,
            patent.get("title", ""),
            patent.get("assignee", ""),
            patent.get("year"),
            patent.get("patent_type", ""),
            patent.get("claim_category", ""),
            count,
            split.get("applicant", 0),
            split.get("examiner", 0),
            count - split.get("applicant", 0) - split.get("examiner", 0),
        ])

    if top_rows:
        table_bottom = _write_table(sheet, row + 1, 1, header, top_rows)
        for r in range(row + 2, table_bottom + 1):
            sheet.cell(row=r, column=5).number_format = "0"
            sheet.cell(row=r, column=3).alignment = Alignment(wrap_text=True, vertical="top")
    else:
        sheet.cell(row=row + 1, column=1, value="No forward-citation pairs found -- run the forward-citations --export first.")

    _set_widths(sheet, {
        "A": 30, "B": 18, "C": 48, "D": 30, "E": 14, "F": 20,
        "G": 20, "H": 14, "I": 14, "J": 14, "K": 14,
    })

    sheet.freeze_panes = "A4"

    notes.append(["PATENTS"])
    notes.append(["Filter results read from", str(FILTERED_DIR)])
    notes.append(["Unique patents/applications", len(patents)])
    notes.append(["Rows listed in two product files (2001 APS + SGML), counted once", duplicates])
    notes.append(["Patents with no parseable date (left out of the decade chart)", undated])
    notes.append(["Decade is taken from the grant date (publication date for applications)."])
    notes.append([
        "Applications and grants are separate documents, so an invention "
        "published as an application and later granted counts in both."
    ])
    if not classified:
        notes.append([
            "Claim category is missing from the filter results -- run "
            "uspto_patent_pipeline_v15.py --classify, then re-run this report."
        ])
    notes.append(["Forward citations counted from", ", ".join(pair_files) or "(none found)"])
    notes.append(["Filtered patents cited at least once", len(totals)])
    notes.append([])


# ============================================================
# CRAWL
# ============================================================

def _open_tds_db():

    if not TDS_DB.exists():
        return None

    try:
        connection = sqlite3.connect(
            f"file:{TDS_DB.as_posix()}?mode=ro", uri=True, timeout=30
        )
        connection.execute("PRAGMA busy_timeout=30000")
        return connection
    except sqlite3.Error as error:
        print(f"  could not open {TDS_DB.name}: {error}")
        return None


def _db_counts(connection, sql):

    """{vendor: {state: count}} from a 'vendor, state, count' query;
    empty if the table isn't there."""

    result = defaultdict(Counter)

    if connection is None:
        return result

    try:
        for vendor, state, count in connection.execute(sql):
            result[vendor][state or ""] += count
    except sqlite3.Error as error:
        print(f"  skipped a crawl-database query: {error}")

    return result


def build_crawl_sheets(workbook, notes):

    print("\nCRAWL")

    if not TDS_MANIFEST.exists():
        sheet = workbook.create_sheet("Crawl")
        sheet.cell(row=1, column=1, value=f"{TDS_MANIFEST} not found -- run tds_crawler_v1.py --export first.")
        return

    with open(TDS_MANIFEST, "r", newline="", encoding="utf-8-sig") as handle:
        documents = list(csv.DictReader(handle))

    print(f"  {TDS_MANIFEST.name}: {len(documents):,} files")

    connection = _open_tds_db()

    pages = _db_counts(connection, "SELECT vendor, state, COUNT(*) FROM pages GROUP BY vendor, state")
    links = _db_counts(connection, "SELECT vendor, state, COUNT(*) FROM pdf_links GROUP BY vendor, state")

    vendor_info = {}

    if connection is not None:
        try:
            for vendor, categories, status in connection.execute(
                "SELECT vendor, categories, status FROM vendors"
            ):
                vendor_info[vendor] = (categories or "", status or "")
        except sqlite3.Error as error:
            print(f"  skipped vendors table: {error}")
        connection.close()

    # --- per-document tallies -----------------------------------

    def kind(doc):
        if (doc.get("rejected_as") or "").strip():
            return "rejected"
        if (doc.get("is_translated_copy") or "").strip():
            return "translation"
        return "kept"

    kept = [d for d in documents if kind(d) == "kept"]
    translations = [d for d in documents if kind(d) == "translation"]
    rejected = [d for d in documents if kind(d) == "rejected"]

    def as_int(value):
        try:
            return int(float(value or 0))
        except ValueError:
            return 0

    category_counts = Counter(
        (d.get("category_folder") or "Unassigned").replace("_", " ") for d in kept
    )

    type_counts = Counter(
        DOC_TYPE_LABEL.get((d.get("doc_type_from_text") or "").strip().lower(),
                           (d.get("doc_type_from_text") or "Unknown").strip() or "Unknown")
        for d in kept
    )

    family_counts = Counter()
    for d in kept:
        for family in re.split(r"[;,]", d.get("method_families") or ""):
            family = family.strip()
            if family:
                family_counts[family] += 1
    with_methods = sum(1 for d in kept if (d.get("test_methods") or "").strip())

    language_counts = Counter((d.get("language") or "unknown").strip() or "unknown" for d in documents if kind(d) != "rejected")

    # --- per-vendor table ---------------------------------------

    vendors = defaultdict(lambda: {
        "category": "", "kept": 0, "tds": 0, "sds": 0, "other_types": 0,
        "translations": 0, "rejected": 0, "pages": 0, "bytes": 0,
        "latest": "",
    })

    for d in documents:
        v = vendors[d.get("vendor") or "(unknown)"]
        v["category"] = v["category"] or (d.get("category_folder") or "").replace("_", " ")
        k = kind(d)
        if k == "kept":
            v["kept"] += 1
            doc_type = (d.get("doc_type_from_text") or "").strip().lower()
            if doc_type == "tds":
                v["tds"] += 1
            elif doc_type == "sds":
                v["sds"] += 1
            else:
                v["other_types"] += 1
        elif k == "translation":
            v["translations"] += 1
        else:
            v["rejected"] += 1
        if k != "rejected":
            v["pages"] += as_int(d.get("pages"))
            v["bytes"] += as_int(d.get("bytes"))
        v["latest"] = max(v["latest"], (d.get("downloaded_at") or "")[:10])

    for name in set(vendor_info) | set(pages) | set(links):
        vendors[name]  # every crawled vendor gets a row, even with 0 files

    vendor_rows = []

    for name, v in vendors.items():
        categories, status = vendor_info.get(name, ("", ""))
        vendor_rows.append([
            name,
            v["category"] or categories,
            status,
            pages[name].get("done", 0),
            sum(pages[name].values()),
            sum(links[name].values()),
            sum(c for state, c in links[name].items() if state in ("downloaded", "duplicate", "unchanged")),
            links[name].get("error", 0),
            v["kept"], v["tds"], v["sds"], v["other_types"],
            v["translations"], v["rejected"],
            v["pages"], round(v["bytes"] / 1_048_576, 1),
            v["latest"],
        ])

    vendor_rows.sort(key=lambda r: (-r[8], r[0].lower()))

    # --- Crawl sheet --------------------------------------------

    sheet = workbook.create_sheet("Crawl")

    _write_title(sheet, 1, "Vendor technical data sheet crawl", TITLE_FONT)

    total_pages_crawled = sum(c.get("done", 0) for c in pages.values())
    total_links = sum(sum(c.values()) for c in links.values())
    link_states = Counter()
    for c in links.values():
        link_states.update(c)

    headline = [
        ["Vendors in crawl", len(vendor_rows)],
        ["Vendors with at least one kept document", sum(1 for r in vendor_rows if r[8])],
        ["Web pages crawled", total_pages_crawled],
        ["PDF links found", total_links],
        ["Unique PDF files downloaded", len(documents)],
        ["Documents kept (English / original)", len(kept)],
        ["Foreign-language copies (_translations)", len(translations)],
        ["Rejected (_rejected)", len(rejected)],
        ["Kept documents citing a test method", with_methods],
        ["Total pages in kept + translated files", sum(as_int(d.get("pages")) for d in kept + translations)],
        ["Total size of kept + translated files (GB)", round(sum(as_int(d.get("bytes")) for d in kept + translations) / 1_073_741_824, 2)],
    ]

    _write_title(sheet, 3, "Headline numbers")
    bottom = _write_table(sheet, 4, 1, ["Measure", "Value"], headline)
    sheet.cell(row=bottom, column=2).number_format = "0.00"

    row = bottom + 2
    _write_title(sheet, row, "PDF links by download outcome")
    bottom = _write_table(
        sheet, row + 1, 1, ["Outcome", "Links"],
        [[state or "(blank)", count] for state, count in link_states.most_common()],
    )

    def chart_block(start_row, title, header, rows, horizontal=False, height=10):
        _write_title(sheet, start_row, title)
        top = start_row + 1
        end = _write_table(sheet, top, 1, header, rows)
        if rows:
            _bar_chart(
                sheet, title, top, 1, end, 1,
                anchor=f"D{top}", y_title=header[1],
                horizontal=horizontal, width=22, height=height,
            )
        return max(end, top + int(height * 2)) + 2

    row = max(bottom + 2, 18)

    row = chart_block(
        row, "Documents by material category", ["Material category", "Documents"],
        sorted(category_counts.items()),
    )
    row = chart_block(
        row, "Documents by type", ["Document type", "Documents"],
        type_counts.most_common(),
    )
    row = chart_block(
        row, f"Top {TOP_N_VENDORS} vendors by documents kept", ["Vendor", "Documents"],
        [[r[0], r[8]] for r in vendor_rows[:TOP_N_VENDORS] if r[8]],
        horizontal=True, height=12,
    )
    row = chart_block(
        row, "Test-method families cited (kept documents)", ["Method family", "Documents"],
        family_counts.most_common(),
    )
    row = chart_block(
        row, "Documents by language (kept + translations)", ["Language", "Documents"],
        language_counts.most_common(),
    )

    _set_widths(sheet, {"A": 44, "B": 14, "C": 4})

    # --- Vendors sheet ------------------------------------------

    vendor_sheet = workbook.create_sheet("Vendors")

    header = [
        "Vendor", "Material category", "Seed status", "Pages crawled",
        "Pages queued/seen", "PDF links found", "PDF links downloaded",
        "PDF links errored", "Documents kept", "Technical data sheets",
        "Safety data sheets", "Other document types", "Translations",
        "Rejected", "Total PDF pages", "Size (MB)", "Last download",
    ]

    bottom = _write_table(vendor_sheet, 1, 1, header, vendor_rows)

    for r in range(2, bottom + 1):
        vendor_sheet.cell(row=r, column=16).number_format = "#,##0.0"

    vendor_sheet.auto_filter.ref = f"A1:{get_column_letter(len(header))}{bottom}"
    vendor_sheet.freeze_panes = "B2"
    vendor_sheet.row_dimensions[1].height = 32

    _set_widths(vendor_sheet, {get_column_letter(i): 13 for i in range(3, len(header) + 1)})
    _set_widths(vendor_sheet, {"A": 36, "B": 24})

    notes.append(["CRAWL"])
    notes.append(["Manifest read", str(TDS_MANIFEST)])
    notes.append(["Crawl database read (read-only)", str(TDS_DB) if connection is not None else "not found / not readable"])
    notes.append([
        "'Documents kept' = files not in _rejected and not a foreign-language "
        "copy; type is the type read from the PDF text."
    ])
    notes.append([
        "'PDF links downloaded' includes links whose file was a duplicate of "
        "one already saved or unchanged since the last run."
    ])
    notes.append([])


# ============================================================
# COPY FILES
# ============================================================

def _copy_set(source_files, destination, label, notes):

    destination.mkdir(parents=True, exist_ok=True)

    wanted = {path.name for path in source_files}

    copied = skipped = 0

    for source in source_files:

        target = destination / source.name

        stat = source.stat()

        if (
            target.exists()
            and target.stat().st_size == stat.st_size
            and target.stat().st_mtime >= stat.st_mtime
        ):
            skipped += 1
            continue

        print(f"  copying {source.name} ({stat.st_size / 1_048_576:,.1f} MB)")

        shutil.copy2(source, target)

        copied += 1

    # A stale "_part3" from an earlier, bigger run shouldn't linger.
    removed = 0
    for old in destination.iterdir():
        if old.is_file() and old.suffix.lower() in (".csv", ".xlsx") and old.name not in wanted:
            old.unlink()
            removed += 1

    print(f"  {label}: {copied} copied, {skipped} already up to date, {removed} stale removed")

    notes.append([label, f"{len(source_files)} files -> {destination}"])

    for source in source_files:
        notes.append(["", source.name])


def copy_files(notes):

    print("\nCOPYING FILES")

    notes.append(["COPIED FILES"])

    forward_files = []

    workbook_path = FORWARD_DIR / "forward_citations.xlsx"
    if workbook_path.exists():
        forward_files.append(workbook_path)

    # Per-section pairs files (and their _partN continuations) --
    # each one fits within Excel's row limit. The combined
    # forward_citation_pairs.csv does not, so it stays behind.
    forward_files += sorted(
        path for path in FORWARD_DIR.glob("forward_citation_pairs_*.csv")
    )

    _copy_set(forward_files, OVERVIEW_FORWARD_DIR, "Forward citations", notes)

    filtered_files = [
        FILTERED_DIR / f"filtered_patents_{product}.csv"
        for product, _ in PRODUCTS
        if (FILTERED_DIR / f"filtered_patents_{product}.csv").exists()
    ]

    _copy_set(filtered_files, OVERVIEW_FILTERED_DIR, "Filtered patents (by file type)", notes)

    notes.append([])


# ============================================================
# MAIN
# ============================================================

def main():

    parser = argparse.ArgumentParser(description="Build the research overview folder and report.")
    parser.add_argument("--no-copy", action="store_true", help="build the report only")
    parser.add_argument("--copy-only", action="store_true", help="copy the result files only")
    args = parser.parse_args()

    started = time.time()

    OVERVIEW_ROOT.mkdir(parents=True, exist_ok=True)

    notes = [
        ["Generated", datetime.now().isoformat(timespec="seconds")],
        ["USPTO corpus", str(USPTO_ROOT)],
        ["Data-sheet corpus", str(TDS_ROOT)],
        [],
    ]

    if not args.copy_only:

        workbook = openpyxl.Workbook()
        workbook.remove(workbook.active)

        build_patents_sheet(workbook, notes)
        build_crawl_sheets(workbook, notes)

    if not args.no_copy:
        copy_files(notes)

    if not args.copy_only:

        notes_sheet = workbook.create_sheet("Notes")
        for line in notes:
            notes_sheet.append(line)
        notes_sheet.column_dimensions["A"].width = 70
        notes_sheet.column_dimensions["B"].width = 70

        report_path = OVERVIEW_ROOT / REPORT_NAME

        try:
            workbook.save(report_path)
        except PermissionError:
            report_path = OVERVIEW_ROOT / f"overview_report_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            workbook.save(report_path)
            print(f"\n  {REPORT_NAME} is open in Excel -- wrote a time-stamped copy instead.")

        print(f"\nReport: {report_path}")

    print(f"Done in {time.time() - started:,.0f}s")


if __name__ == "__main__":
    main()
