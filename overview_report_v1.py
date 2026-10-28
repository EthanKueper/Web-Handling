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
              citations; line chart: patents per year (grants vs
              applications); top 15 CPC subclasses
    Citations forward-citation distribution, top 10 by citations
              per year, average citations by claim category, and
              what kinds of documents cite the filtered patents
    Crawl     vendor data-sheet crawl: headline numbers; charts of
              documents by material category, by document type,
              top 15 vendors, test-method families and language;
              material category x document type table; vendors with
              no kept documents and why
    Vendors   one row per vendor: pages crawled, PDF links found,
              documents kept by type, translations, rejected, size
    Vendors & Patents
              data-sheet vendors matched to patent assignees: patents,
              citations and data sheets side by side
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
from openpyxl.chart import BarChart, LineChart, Reference
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


# Short titles for the CPC/IPC subclasses most likely to appear
# in a web-handling / films / nonwovens / foil / paper corpus.
# A subclass not listed here is still counted -- its title cell
# is just left blank.
CPC_SUBCLASS_TITLES = {
    "A41D": "Outerwear; protective garments",
    "A47K": "Sanitary equipment; toilet accessories",
    "A47L": "Domestic washing or cleaning",
    "A61B": "Diagnosis; surgery",
    "A61F": "Absorbent articles, bandages, diapers; prostheses",
    "A61K": "Medical, dental or toiletry preparations",
    "A61L": "Sterilising; materials for dressings and absorbent pads",
    "A61M": "Devices for introducing media into the body",
    "B01D": "Separation; filtration",
    "B01J": "Chemical or physical processes; catalysis",
    "B05B": "Spraying apparatus; nozzles",
    "B05C": "Apparatus for applying liquids to surfaces (coaters)",
    "B05D": "Processes for applying liquids to surfaces (coating)",
    "B08B": "Cleaning in general",
    "B21B": "Rolling of metal",
    "B21C": "Making metal sheet, wire, rods, tubes (non-rolling)",
    "B21D": "Working sheet metal",
    "B22F": "Working metallic powder",
    "B23K": "Soldering, welding, cladding",
    "B26D": "Cutting; perforating; punching",
    "B26F": "Perforating, punching, cutting-out by other means",
    "B27N": "Articles from wood particles or fibres",
    "B29B": "Preparing plastics (pellets, preforms)",
    "B29C": "Shaping or joining plastics (extrusion, moulding, stretching)",
    "B29D": "Producing particular articles from plastics",
    "B29K": "Plastics materials (indexing)",
    "B29L": "Plastics articles (indexing)",
    "B30B": "Presses",
    "B31B": "Making paper or cardboard containers",
    "B31C": "Making wound paper articles (tubes)",
    "B31D": "Making other paper articles",
    "B31F": "Working paper (creping, embossing, corrugating)",
    "B32B": "Layered products (laminates)",
    "B41F": "Printing machines",
    "B41J": "Selective printing mechanisms (inkjet)",
    "B41M": "Printing, duplicating, marking processes",
    "B41N": "Printing plates and blankets",
    "B42D": "Books; security documents and cards",
    "B44C": "Decorative effects; decalcomania",
    "B65B": "Packaging machines and methods",
    "B65C": "Labelling machines",
    "B65D": "Containers, packaging elements, packages",
    "B65G": "Conveying; storage",
    "B65H": "Handling webs, sheets and filaments (winding, unwinding)",
    "B82Y": "Nanotechnology",
    "C01B": "Non-metallic elements and compounds",
    "C03C": "Glass composition; glass surface treatment",
    "C08B": "Polysaccharides; cellulose derivatives",
    "C08C": "Rubber treatment",
    "C08F": "Addition polymers (polyolefins, acrylics)",
    "C08G": "Condensation polymers (polyesters, polyamides, PU)",
    "C08H": "Derivatives of natural polymers",
    "C08J": "Working-up of polymers (films, foams, compounding)",
    "C08K": "Polymer additives and fillers",
    "C08L": "Polymer compositions",
    "C09B": "Organic dyes",
    "C09C": "Pigments and fillers (treatment)",
    "C09D": "Coating compositions; inks; varnishes",
    "C09J": "Adhesives",
    "C09K": "Materials for miscellaneous applications",
    "C11D": "Detergents",
    "C22C": "Alloys",
    "C22F": "Heat treatment of non-ferrous metals and alloys",
    "C23C": "Coating metals; vapour deposition",
    "C25D": "Electroplating; electrolytic coating",
    "D01D": "Making artificial filaments (spinning apparatus)",
    "D01F": "Artificial filaments (chemical features)",
    "D02G": "Crimping fibres; yarns",
    "D02J": "Finishing filaments and yarns",
    "D03D": "Woven fabrics; weaving",
    "D04B": "Knitting",
    "D04H": "Nonwoven fabrics",
    "D06B": "Treating textiles with liquids or gases",
    "D06C": "Finishing and stretching textile fabrics",
    "D06M": "Chemical treatment of fibres and fabrics",
    "D06N": "Coated textile coverings (floor, wall)",
    "D06P": "Dyeing or printing textiles",
    "D21B": "Fibrous raw materials for paper",
    "D21C": "Pulp production",
    "D21D": "Stock preparation before the paper machine",
    "D21F": "Paper-making machines and methods",
    "D21G": "Calenders; paper-machine accessories",
    "D21H": "Paper compositions, coating and impregnation",
    "D21J": "Fibreboard; moulded pulp",
    "F16C": "Shafts; bearings",
    "F16L": "Pipes and hoses",
    "F26B": "Drying",
    "G01B": "Measuring length, thickness, area",
    "G01N": "Analysing materials (testing, inspection)",
    "G03F": "Photolithography",
    "G05B": "Control systems",
    "G06F": "Digital data processing",
    "G06T": "Image processing",
    "H01B": "Cables; conductors; insulators",
    "H01G": "Capacitors",
    "H01L": "Semiconductor devices",
    "H01M": "Batteries; fuel cells (incl. separators, foils)",
    "H05K": "Printed circuits",
    "Y10T": "Former US classification cross-reference art",
}

TOP_N_CPC = 15

# Forward-citation rate: same floor as the forward-citations
# script (a patent granted last year isn't divided by ~0), and a
# minimum count so a patent with 3 citations in its first year
# doesn't top the "fastest-cited" table.
CITATION_RATE_MIN_YEARS = 1.0
CITATION_RATE_MIN_CITES = 10

CITATION_BUCKETS = [
    ("0", 0, 0),
    ("1-5", 1, 5),
    ("6-20", 6, 20),
    ("21-100", 21, 100),
    ("101-500", 101, 500),
    ("Over 500", 501, None),
]


def classify_patent_type(number):

    """Utility / Design / Plant / Reissue / SIR / Defensive
    Publication / Application, from the number (same rules as the
    pipeline's classify_patent_type)."""

    value = re.sub(r"[^A-Z0-9]", "", (number or "").strip().upper())

    if value.startswith("US"):
        value = value[2:]

    application = re.match(r"^(\d{11})([A-Z]\d?)?$", value)

    if application and 1900 <= int(value[:4]) <= 2100:
        return "Application (plant)" if (application.group(2) or "").startswith("P") else "Application (utility)"
    if value.startswith("PP"):
        return "Plant"
    if value.startswith("RE"):
        return "Reissue"
    if value.startswith("D"):
        return "Design"
    if value.startswith("H"):
        return "Statutory Invention Registration (SIR)"
    if value.startswith("T"):
        return "Defensive Publication"
    if value[:1].isdigit():
        return "Utility"
    return "Unknown"


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


def _subclasses(cpc, ipc):

    """Distinct 4-character subclasses (e.g. B65H) from the CPC
    codes, falling back to IPC for records without CPC (pre-2013),
    the same fallback the filter uses."""

    found = []

    for code in re.split(r"[;,]", cpc or "") or []:
        match = re.match(r"\s*([A-HY]\d{2}[A-Z])", code.upper())
        if match and match.group(1) not in found:
            found.append(match.group(1))

    if not found:
        for code in re.split(r"[;,]", ipc or ""):
            match = re.match(r"\s*([A-H]\d{2}[A-Z])", code.upper())
            if match and match.group(1) not in found:
                found.append(match.group(1))

    return tuple(found)


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
                    "is_application": product == "appxml",
                    "subclasses": _subclasses(row.get("cpc"), row.get("ipc")),
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

    Returns a dict:
        totals         Counter  filtered patent -> times cited
        by_category    {patent: Counter(applicant/examiner/...)}
        citing_types   Counter  type -> distinct citing documents
        coverage_year  latest citing-patent year seen
        files          names of the files read
    """

    combined = FORWARD_DIR / "forward_citation_pairs.csv"

    if combined.exists():
        files = [combined]
    else:
        files = sorted(FORWARD_DIR.glob("forward_citation_pairs_*.csv"))

    totals = Counter()
    by_category = defaultdict(Counter)
    citing_seen = set()
    citing_types = Counter()
    coverage_year = None

    for path in files:

        print(f"  reading {path.name} ...")

        with open(path, "r", newline="", encoding="utf-8-sig") as handle:

            reader = csv.reader(handle)
            header = next(reader, [])

            try:
                cited_col = header.index("filtered_patent_number")
            except ValueError:
                continue

            def column(name):
                return header.index(name) if name in header else None

            category_col = column("citation_category")
            citing_col = column("citing_patent_number")
            citing_type_col = column("citing_patent_type")
            citing_date_col = column("citing_patent_date")

            width = len(header)

            for row in reader:

                if len(row) < width:
                    continue

                number = row[cited_col]

                totals[number] += 1

                if category_col is not None:
                    by_category[number][row[category_col]] += 1

                if citing_col is not None:
                    citing = row[citing_col]
                    if citing not in citing_seen:
                        citing_seen.add(citing)
                        citing_types[
                            (row[citing_type_col] if citing_type_col is not None else "")
                            or classify_patent_type(citing)
                        ] += 1

                if citing_date_col is not None:
                    year = _year(row[citing_date_col])
                    if year and (coverage_year is None or year > coverage_year):
                        coverage_year = year

    return {
        "totals": totals,
        "by_category": by_category,
        "citing_types": citing_types,
        "coverage_year": coverage_year or datetime.now().year,
        "files": [path.name for path in files],
    }


def _patent_table(sheet, row, title, ranked, cites, extra_header=None):

    """Top-N patent table. `ranked` is a list of
    (patent_number, extra_value_or_None). Returns the next free
    row."""

    patents, totals, by_category = cites["patents"], cites["totals"], cites["by_category"]

    _write_title(sheet, row, title)

    header = [
        "Rank", "Patent", "Title", "Assignee", "Year", "Patent type",
        "Claim category", "Times cited", "Cited by applicant",
        "Cited by examiner", "Cited by other / unknown",
    ]
    if extra_header:
        header.append(extra_header)

    rows = []

    for rank, (number, extra) in enumerate(ranked, start=1):

        patent = patents.get(number, {})
        split = by_category.get(number, Counter())
        count = totals.get(number, 0)

        values = [
            rank, number, patent.get("title", ""), patent.get("assignee", ""),
            patent.get("year"), patent.get("patent_type", ""),
            patent.get("claim_category", ""), count,
            split.get("applicant", 0), split.get("examiner", 0),
            count - split.get("applicant", 0) - split.get("examiner", 0),
        ]
        if extra_header:
            values.append(extra)

        rows.append(values)

    if not rows:
        sheet.cell(row=row + 1, column=1, value="No forward-citation pairs found -- run the forward-citations --export first.")
        return row + 3

    bottom = _write_table(sheet, row + 1, 1, header, rows)

    for r in range(row + 2, bottom + 1):
        sheet.cell(row=r, column=5).number_format = "0"
        sheet.cell(row=r, column=3).alignment = Alignment(wrap_text=True, vertical="top")
        if extra_header:
            sheet.cell(row=r, column=len(header)).number_format = "0.00"

    return bottom + 2


def build_patents_sheet(workbook, notes, data, cites):

    patents, file_counts, duplicates, classified = data

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

    row = _patent_table(
        sheet, row, f"Top {TOP_N_PATENTS} patents by forward citations",
        [(number, None) for number, _ in cites["totals"].most_common(TOP_N_PATENTS)],
        cites,
    )

    # --- 5. patents per year: grants vs applications ------------

    year_counts = defaultdict(Counter)

    for patent in patents.values():
        if patent["year"] is not None:
            year_counts[patent["year"]]["Applications" if patent["is_application"] else "Grants"] += 1

    _write_title(sheet, row, "Patents per year (grants vs applications)")

    table_top = row + 1
    years = range(min(year_counts), max(year_counts) + 1) if year_counts else []
    table_bottom = _write_table(
        sheet, table_top, 1, ["Year", "Grants", "Applications"],
        [[str(year), year_counts[year]["Grants"], year_counts[year]["Applications"]] for year in years],
    )

    if year_counts:
        chart = LineChart()
        chart.title = "Patents per year"
        chart.width, chart.height = 26, 11
        chart.y_axis.title = "Documents"
        chart.x_axis.title = "Year (grant date; publication date for applications)"
        chart.add_data(Reference(sheet, min_col=2, max_col=3, min_row=table_top, max_row=table_bottom), titles_from_data=True)
        chart.set_categories(Reference(sheet, min_col=1, min_row=table_top + 1, max_row=table_bottom))
        chart.x_axis.delete = False
        chart.y_axis.delete = False
        chart.y_axis.scaling.min = 0
        chart.x_axis.tickLblSkip = 5
        sheet.add_chart(chart, f"E{table_top}")

    row = max(table_bottom, table_top + 24) + 2

    # --- 6. top CPC subclasses ----------------------------------

    subclass_counts = Counter()
    subclass_cites = Counter()

    for number, patent in patents.items():
        for subclass in patent["subclasses"]:
            subclass_counts[subclass] += 1
            subclass_cites[subclass] += cites["totals"].get(number, 0)

    _write_title(sheet, row, f"Top {TOP_N_CPC} CPC subclasses")

    table_top = row + 1
    top_subclasses = subclass_counts.most_common(TOP_N_CPC)
    table_bottom = _write_table(
        sheet, table_top, 1,
        ["Subclass", "Patents", "Share of corpus", "Forward citations", "Title"],
        [
            [
                subclass, count,
                round(count / len(patents), 4) if patents else 0,
                subclass_cites[subclass],
                CPC_SUBCLASS_TITLES.get(subclass, ""),
            ]
            for subclass, count in top_subclasses
        ],
    )

    for r in range(table_top + 1, table_bottom + 1):
        sheet.cell(row=r, column=3).number_format = "0.0%"

    if top_subclasses:
        _bar_chart(
            sheet, f"Top {TOP_N_CPC} CPC subclasses", table_top, 1, table_bottom, 1,
            anchor=f"G{table_top}", y_title="Patents", horizontal=True, width=20, height=11,
        )

    row = max(table_bottom, table_top + 22) + 2

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
    notes.append([
        "CPC subclasses: a patent counts once in each subclass it carries "
        "(IPC used for records with no CPC); shares add to more than 100%."
    ])
    notes.append([])


def build_citations_sheet(workbook, notes, data, cites):

    patents = data[0]
    totals = cites["totals"]
    coverage_year = cites["coverage_year"]

    sheet = workbook.create_sheet("Citations")

    _write_title(sheet, 1, "Forward citations to the filtered patents", TITLE_FONT)

    sheet.cell(
        row=2, column=1,
        value=(
            f"{sum(totals.values()):,} citations to {len(totals):,} of "
            f"{len(patents):,} filtered patents; citing patents through {coverage_year}"
        ),
    )

    # --- 1. distribution ----------------------------------------

    bucket_counts = Counter()

    for number in patents:
        count = totals.get(number, 0)
        for label, low, high in CITATION_BUCKETS:
            if count >= low and (high is None or count <= high):
                bucket_counts[label] += 1
                break

    row = 4
    _write_title(sheet, row, "How often filtered patents are cited")

    table_top = row + 1
    table_bottom = _write_table(
        sheet, table_top, 1, ["Times cited", "Patents", "Share"],
        [
            [label, bucket_counts[label], round(bucket_counts[label] / len(patents), 4) if patents else 0]
            for label, _, _ in CITATION_BUCKETS
        ],
    )

    for r in range(table_top + 1, table_bottom + 1):
        sheet.cell(row=r, column=3).number_format = "0.0%"

    _bar_chart(
        sheet, "Filtered patents by times cited", table_top, 1, table_bottom, 1,
        anchor=f"E{table_top}", x_title="Times cited", y_title="Patents", width=20, height=10,
    )

    row = max(table_bottom, table_top + 20) + 2

    # --- 2. top 10 by citations per year ------------------------

    rates = []

    for number, count in totals.items():
        patent = patents.get(number)
        if not patent or patent["year"] is None or count < CITATION_RATE_MIN_CITES:
            continue
        years = max(CITATION_RATE_MIN_YEARS, coverage_year - patent["year"])
        rates.append((count / years, number))

    rates.sort(reverse=True)

    row = _patent_table(
        sheet, row,
        f"Top {TOP_N_PATENTS} patents by citations per year "
        f"(at least {CITATION_RATE_MIN_CITES} citations)",
        [(number, round(rate, 2)) for rate, number in rates[:TOP_N_PATENTS]],
        cites, extra_header="Citations per year",
    )

    # --- 3. average citations by claim category -----------------

    category_patents = Counter()
    category_cites = Counter()
    category_cited = Counter()

    for number, patent in patents.items():
        category = patent["claim_category"]
        count = totals.get(number, 0)
        category_patents[category] += 1
        category_cites[category] += count
        if count:
            category_cited[category] += 1

    categories = [c for c in CLAIM_CATEGORY_ORDER if category_patents.get(c)]
    categories += sorted(c for c in category_patents if c not in categories)

    _write_title(sheet, row, "Average forward citations by claim category")

    table_top = row + 1
    table_bottom = _write_table(
        sheet, table_top, 1,
        ["Claim category", "Average citations per patent", "Patents", "Cited at least once", "Total citations"],
        [
            [
                c,
                round(category_cites[c] / category_patents[c], 2),
                category_patents[c], category_cited[c], category_cites[c],
            ]
            for c in categories
        ],
    )

    for r in range(table_top + 1, table_bottom + 1):
        sheet.cell(row=r, column=2).number_format = "0.00"

    _bar_chart(
        sheet, "Average forward citations per patent, by claim category",
        table_top, 1, table_bottom, 1,
        anchor=f"G{table_top}", y_title="Citations per patent", width=22, height=10,
    )

    row = max(table_bottom, table_top + 20) + 2

    # --- 4. who cites them --------------------------------------

    _write_title(sheet, row, "Documents citing the filtered patents, by type")

    citing = cites["citing_types"]
    total_citing = sum(citing.values())

    table_top = row + 1
    table_bottom = _write_table(
        sheet, table_top, 1, ["Citing document type", "Citing documents", "Share"],
        [
            [label, count, round(count / total_citing, 4) if total_citing else 0]
            for label, count in citing.most_common()
        ],
    )

    for r in range(table_top + 1, table_bottom + 1):
        sheet.cell(row=r, column=3).number_format = "0.0%"

    _set_widths(sheet, {
        "A": 34, "B": 18, "C": 48, "D": 30, "E": 14, "F": 20,
        "G": 20, "H": 14, "I": 14, "J": 14, "K": 14, "L": 14,
    })

    notes.append(["CITATIONS"])
    notes.append(["Forward citations counted from", ", ".join(cites["files"]) or "(none found)"])
    notes.append(["Filtered patents cited at least once", len(totals)])
    notes.append([
        "Citations per year = times cited / years from grant to "
        f"{coverage_year} (the latest citing-patent year), at least "
        f"{CITATION_RATE_MIN_YEARS:g} year; patents with fewer than "
        f"{CITATION_RATE_MIN_CITES} citations are left out of that table."
    ])
    notes.append([
        "Citing documents are counted once each, however many filtered "
        "patents they cite."
    ])
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
        return {}

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

    # --- material category x document type ----------------------

    type_order = [label for label, _ in type_counts.most_common()]
    matrix = defaultdict(Counter)

    for d in kept:
        doc_type = (d.get("doc_type_from_text") or "").strip()
        matrix[(d.get("category_folder") or "Unassigned").replace("_", " ")][
            DOC_TYPE_LABEL.get(doc_type.lower(), doc_type or "Unknown")
        ] += 1

    _write_title(sheet, row, "Documents by material category and type (kept documents)")

    matrix_rows = [
        [category] + [matrix[category].get(t, 0) for t in type_order] + [sum(matrix[category].values())]
        for category in sorted(matrix)
    ]
    matrix_rows.append(
        ["Total"] + [type_counts[t] for t in type_order] + [len(kept)]
    )

    bottom = _write_table(sheet, row + 1, 1, ["Material category"] + type_order + ["Total"], matrix_rows)

    for c in range(1, len(type_order) + 3):
        sheet.cell(row=bottom, column=c).font = Font(bold=True)

    row = bottom + 3

    # --- vendors with no kept documents --------------------------

    def no_document_reason(r):
        name, status, crawled, seen, link_count, downloaded = r[0], r[2], r[3], r[4], r[5], r[6]
        page_states = pages.get(name, Counter())
        if status == "skip":
            return "Skipped: status=skip in vendor_seeds.csv"
        if seen == 0:
            return "Not crawled yet: no pages fetched"
        if crawled == 0:
            return (
                f"Site blocked or erroring: {page_states.get('blocked', 0)} blocked, "
                f"{page_states.get('error', 0)} errors, {page_states.get('queued', 0)} still queued"
            )
        if link_count == 0:
            return f"No PDF links found on {crawled:,} pages crawled"
        if downloaded == 0:
            outcomes = ", ".join(
                f"{state or 'blank'} {count}" for state, count in links.get(name, Counter()).most_common(3)
            )
            return f"PDF links found but none downloaded ({outcomes})"
        if r[12] and not r[13]:
            return f"Only foreign-language copies ({r[12]}), no English version kept"
        if r[13]:
            return f"All downloaded files rejected ({r[13]})"
        return "Downloads not kept (off-topic files are deleted -- see logs\\deleted_off_topic_*.csv)"

    empty = [r for r in vendor_rows if not r[8]]

    _write_title(sheet, row, f"Vendors with no kept documents ({len(empty)})")

    bottom = _write_table(
        sheet, row + 1, 1,
        ["Vendor", "Material category", "Seed status", "Pages crawled",
         "PDF links found", "PDF links downloaded", "Translations", "Rejected", "Reason"],
        [
            [r[0], r[1], r[2], r[3], r[5], r[6], r[12], r[13], no_document_reason(r)]
            for r in sorted(empty, key=lambda r: r[0].lower())
        ],
    )

    if not empty:
        sheet.cell(row=row + 2, column=1, value="Every vendor has at least one kept document.")

    _set_widths(sheet, {"A": 44, "B": 22, "C": 16, "D": 14, "E": 14, "F": 14, "G": 14, "H": 14, "I": 60})

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

    return {
        r[0]: {"category": r[1], "kept": r[8], "tds": r[9], "sds": r[10]}
        for r in vendor_rows
    }


# ============================================================
# VENDORS & PATENTS
# ============================================================

# Words that don't identify a company. Stripped from both vendor
# names and assignee names before matching, so "3M Company" and
# "3M Innovative Properties Company" both reduce to start "3m".
_CORPORATE_WORDS = {
    "inc", "incorporated", "corp", "corporation", "company", "co",
    "ltd", "limited", "llc", "lp", "llp", "plc", "gmbh", "mbh", "ag",
    "kg", "kgaa", "se", "sa", "sas", "spa", "srl", "bv", "nv", "oy",
    "oyj", "ab", "as", "asa", "aps", "kk", "the", "and", "of",
    "holding", "holdings", "group", "international", "intl",
}

# A one-word vendor name this short or generic would match too
# many unrelated assignees; such names need a second word.
_MIN_SINGLE_WORD_LENGTH = 3


def _company_words(name):

    text = (name or "").lower().replace("&", " ")
    text = re.sub(r"[^\w\s]", " ", text)

    return [w for w in text.split() if w and w not in _CORPORATE_WORDS]


def build_vendor_patent_sheet(workbook, notes, data, cites, vendor_summary):

    print("\nVENDORS & PATENTS")

    patents = data[0]
    totals = cites["totals"]

    sheet = workbook.create_sheet("Vendors & Patents")

    _write_title(sheet, 1, "Data-sheet vendors that also hold filtered patents", TITLE_FONT)

    if not vendor_summary:
        sheet.cell(row=3, column=1, value="No crawl manifest found.")
        return

    # Index vendors by their first identifying word.
    by_first_word = defaultdict(list)

    for vendor in vendor_summary:
        words = _company_words(vendor)
        if not words:
            continue
        if len(words) == 1 and len(words[0]) < _MIN_SINGLE_WORD_LENGTH and not words[0][0].isdigit():
            continue
        by_first_word[words[0]].append((vendor, words))

    matched = defaultdict(list)          # vendor -> [patent numbers]
    assignee_names = defaultdict(Counter)

    for number, patent in patents.items():

        hit_vendors = set()

        for assignee in re.split(r";", patent["assignee"]):

            words = _company_words(assignee)

            if not words:
                continue

            for vendor, vendor_words in by_first_word.get(words[0], []):
                if words[:len(vendor_words)] == vendor_words:
                    hit_vendors.add(vendor)
                    assignee_names[vendor][assignee.strip()] += 1

        for vendor in hit_vendors:
            matched[vendor].append(number)

    rows = []

    for vendor, numbers in matched.items():

        info = vendor_summary[vendor]
        cited = [n for n in numbers if totals.get(n)]
        top = max(numbers, key=lambda n: totals.get(n, 0))

        rows.append([
            vendor,
            info["category"],
            info["kept"],
            info["tds"],
            len(numbers),
            sum(1 for n in numbers if not patents[n]["is_application"]),
            sum(1 for n in numbers if patents[n]["is_application"]),
            len(cited),
            sum(totals.get(n, 0) for n in numbers),
            top if totals.get(top) else "",
            patents[top]["title"] if totals.get(top) else "",
            "; ".join(name for name, _ in assignee_names[vendor].most_common(3)),
        ])

    rows.sort(key=lambda r: (-r[4], r[0].lower()))

    sheet.cell(
        row=2, column=1,
        value=(
            f"{len(rows)} of {len(vendor_summary)} crawled vendors matched to "
            f"patent assignees by name"
        ),
    )

    _write_title(sheet, 4, f"Top {TOP_N_VENDORS} vendors by filtered patents")

    chart_top = 5
    chart_bottom = _write_table(
        sheet, chart_top, 1, ["Vendor", "Filtered patents"],
        [[r[0], r[4]] for r in rows[:TOP_N_VENDORS]],
    )

    if rows:
        _bar_chart(
            sheet, f"Top {TOP_N_VENDORS} vendors by filtered patents",
            chart_top, 1, chart_bottom, 1, anchor=f"D{chart_top}",
            y_title="Patents", horizontal=True, width=22, height=12,
        )

    row = max(chart_bottom, chart_top + 24) + 2

    _write_title(sheet, row, "All matched vendors")

    header = [
        "Vendor", "Material category", "Documents kept", "Technical data sheets",
        "Filtered patents", "Grants", "Applications", "Patents cited at least once",
        "Total forward citations", "Most-cited patent", "Most-cited patent title",
        "Assignee names matched (top 3)",
    ]

    bottom = _write_table(sheet, row + 1, 1, header, rows)

    if rows:
        sheet.auto_filter.ref = f"A{row + 1}:{get_column_letter(len(header))}{bottom}"

    sheet.row_dimensions[row + 1].height = 32

    _set_widths(sheet, {
        "A": 34, "B": 24, "C": 12, "D": 12, "E": 12, "F": 12, "G": 12,
        "H": 12, "I": 12, "J": 16, "K": 48, "L": 60,
    })

    print(f"  {len(rows)} vendors matched to assignees")

    notes.append(["VENDORS & PATENTS"])
    notes.append([
        "Vendors are matched to patent assignees by name: corporate words "
        "(Inc, Corp, GmbH, Company ...) are dropped and the assignee must "
        "begin with the vendor's remaining words, so '3M Company' matches "
        "'3M Innovative Properties Company'."
    ])
    notes.append([
        "It's approximate: a subsidiary under a different name is missed, "
        "and a shared first name (e.g. 'Dow' / 'Dow Corning') can match. "
        "Check the 'Assignee names matched' column."
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

        print("\nPATENTS")
        data = load_filtered_patents()

        print("\nFORWARD CITATIONS")
        cites = load_forward_citation_counts()
        cites["patents"] = data[0]

        build_patents_sheet(workbook, notes, data, cites)
        build_citations_sheet(workbook, notes, data, cites)

        vendor_summary = build_crawl_sheets(workbook, notes)

        build_vendor_patent_sheet(workbook, notes, data, cites, vendor_summary)

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
