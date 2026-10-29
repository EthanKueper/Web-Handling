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

overview_report.xlsx sheets (each section says in plain words what
it shows, with a one-line takeaway computed from the data):

    Start Here          key numbers, a guide to every sheet (with
                        links), a glossary of terms, and the files
                        in the folder
    Patents             patents by decade; by decade split by claim
                        category; by claim category; by file type;
                        top 10 by forward citations; patents per
                        year (grants vs applications); top 15 CPC
                        subclasses
    Citations           how often patents are cited; top 10 by
                        citations per year; average citations by
                        claim category; what kinds of documents cite
                        them
    Crawl               headline numbers; PDF links by outcome;
                        documents by material category, type,
                        vendor, test-method family and language;
                        category x type table; vendors with no kept
                        documents and why
    Vendors             one row per vendor (filterable)
    Vendors & Patents   data-sheet vendors matched to patent
                        assignees
    How It Was Counted  sources read and counting rules

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
import math
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
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint
from openpyxl.comments import Comment
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
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

LINK_STATE_LABEL = {
    "downloaded": "Downloaded (new file)",
    "duplicate": "Same file as another link",
    "unchanged": "Unchanged since last run",
    "not_pdf": "Link was not a PDF",
    "error": "Download error",
    "too_large": "File too large",
    "robots": "Blocked by the site (robots.txt)",
    "pending": "Not downloaded yet",
    "": "(no status)",
}

LANGUAGE_NAMES = {
    "en": "English", "de": "German", "fr": "French", "es": "Spanish",
    "it": "Italian", "pt": "Portuguese", "nl": "Dutch", "sv": "Swedish",
    "fi": "Finnish", "da": "Danish", "no": "Norwegian", "pl": "Polish",
    "cs": "Czech", "hu": "Hungarian", "ru": "Russian", "tr": "Turkish",
    "ja": "Japanese", "zh": "Chinese", "ko": "Korean", "th": "Thai",
    "vi": "Vietnamese", "id": "Indonesian", "ar": "Arabic",
    "unknown": "Unknown",
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
    ("Never cited", 0, 0),
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
# LOOK AND FEEL
# ============================================================
#
# Every analysis sheet follows the same pattern so a first-time
# reader always knows where to look:
#
#   Sheet title, one-sentence summary, "Back to Start Here" link
#   For each section:
#       blue band with the section title
#       grey line: what this shows, in plain words
#       yellow line: the takeaway, computed from the data
#       table on the left, chart on the right (or below)

START_SHEET = "Start Here"

NAVY = "1F3864"
ACCENT = "2E75B6"
MUTED = "595959"
SECTION_FILL = "DDEBF7"
BAND_FILL = "F3F7FC"
TAKEAWAY_FILL = "FFF2CC"
TAKEAWAY_TEXT = "7F6000"
TILE_FILL = "EEF4FB"
WARNING_FILL = "FCE4D6"
LINK_COLOR = "0563C1"

# Series colours, in order, for charts without a fixed meaning.
PALETTE = ["2E75B6", "ED7D31", "70AD47", "7030A0", "FFC000", "5B9BD5", "C00000", "A5A5A5"]

# Fixed colour per claim category, so a category looks the same
# in every chart.
CATEGORY_COLORS = {
    "Process / Method": "2E75B6",
    "Apparatus / System": "ED7D31",
    "Product / Article": "70AD47",
    "Composition": "7030A0",
    "Product-by-process": "FFC000",
    "Design": "5B9BD5",
    "Plant": "548235",
    "Unclassified": "A5A5A5",
    "No claims text": "D9D9D9",
    NOT_CLASSIFIED: "BFBFBF",
}

_THIN = Side(style="thin", color="BFBFBF")
CELL_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
TOTAL_BORDER = Border(left=_THIN, right=_THIN, top=Side(style="medium", color=NAVY), bottom=_THIN)

ROW_HEIGHT_CM = 0.53   # default 15-pt Excel row


def _category_color(category):

    return CATEGORY_COLORS.get(category, "8EA9DB")


def _pct(part, whole):

    return f"{part / whole:.0%}" if whole else "0%"


def _setup_sheet(ws, widths, landscape=True):

    """Column widths, no gridlines, print to one page wide."""

    for column, width in widths.items():
        ws.column_dimensions[column].width = width

    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def _span_width(ws, first_col, last_col):

    return sum(
        ws.column_dimensions[get_column_letter(c)].width or 8.43
        for c in range(first_col, last_col + 1)
    )


def _text_row(ws, row, text, first_col, last_col, font=None, fill=None, min_height=15):

    """One merged, wrapped line of text across first_col..last_col,
    with the row height set so it all shows (Excel doesn't
    auto-fit merged cells)."""

    if last_col > first_col:
        ws.merge_cells(start_row=row, start_column=first_col, end_row=row, end_column=last_col)

    cell = ws.cell(row=row, column=first_col, value=text)
    cell.alignment = Alignment(wrap_text=True, vertical="center")

    if font:
        cell.font = font

    if fill:
        for c in range(first_col, last_col + 1):
            ws.cell(row=row, column=c).fill = fill

    chars_per_line = max(10, _span_width(ws, first_col, last_col) * 1.1)
    size = (font.size if font and font.size else 11) / 11
    lines = sum(max(1, math.ceil(len(part) / chars_per_line)) for part in str(text).split("\n"))

    ws.row_dimensions[row].height = max(min_height, lines * 15 * size + 4)


def _sheet_header(ws, title, summary, span):

    """Title, summary line and a link back to Start Here. Returns
    the first free row."""

    ws.cell(row=1, column=1, value=title).font = Font(bold=True, size=18, color=NAVY)
    ws.row_dimensions[1].height = 30

    _text_row(ws, 2, summary, 1, span, font=Font(size=11, italic=True, color=MUTED))

    link = ws.cell(row=3, column=1, value="\u2190 Back to Start Here")
    link.hyperlink = f"#'{START_SHEET}'!A1"
    link.font = Font(color=LINK_COLOR, underline="single")

    return 5


def _section(ws, row, title, what, span, takeaway=None):

    """Section band + plain-words explanation + takeaway. Returns
    the row where the section's table should start."""

    band = PatternFill("solid", fgColor=SECTION_FILL)

    for c in range(1, span + 1):
        ws.cell(row=row, column=c).fill = band

    ws.cell(row=row, column=1, value=title).font = Font(bold=True, size=13, color=NAVY)
    ws.row_dimensions[row].height = 22

    row += 1
    _text_row(ws, row, what, 1, span, font=Font(size=10, color=MUTED))

    if takeaway:
        row += 1
        _text_row(
            ws, row, f"Takeaway: {takeaway}", 1, span,
            font=Font(size=10, bold=True, color=TAKEAWAY_TEXT),
            fill=PatternFill("solid", fgColor=TAKEAWAY_FILL),
        )

    return row + 2


def _write_table(
    ws, top, header, rows, left=1, formats=None, wrap_columns=(),
    total_row=False, comments=None, spans=None,
):

    """
    Header + rows at (top, left), with borders, banded rows and
    number formats. `formats` maps a column offset (0-based) to a
    number format; otherwise ints get thousands separators and
    floats two decimals. `spans` maps a column offset to how many
    sheet columns it covers (merged), so a long text field such as
    a title can be wide without widening the whole sheet column.
    Returns the last row written.
    """

    formats = formats or {}
    comments = comments or {}
    spans = spans or {}

    positions = []
    column = left
    for offset in range(len(header)):
        positions.append(column)
        column += spans.get(offset, 1)

    def merge(row, offset):
        width = spans.get(offset, 1)
        if width > 1:
            first = positions[offset]
            ws.merge_cells(start_row=row, start_column=first, end_row=row, end_column=first + width - 1)
            for c in range(first + 1, first + width):
                ws.cell(row=row, column=c).border = CELL_BORDER

    header_fill = PatternFill("solid", fgColor=NAVY)
    band = PatternFill("solid", fgColor=BAND_FILL)

    tallest = 1

    for offset, label in enumerate(header):

        column = positions[offset]
        merge(top, offset)
        cell = ws.cell(row=top, column=column, value=label)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
        cell.border = CELL_BORDER

        width = _span_width(ws, column, column + spans.get(offset, 1) - 1)
        tallest = max(tallest, math.ceil(len(str(label)) / max(4, width - 1)))

        if label in comments:
            note = Comment(comments[label], "Overview report")
            note.width, note.height = 300, 110
            cell.comment = note

    ws.row_dimensions[top].height = max(20, 15 * tallest + 4)

    for index, values in enumerate(rows, start=1):

        row = top + index
        is_total = total_row and index == len(rows)

        for offset, value in enumerate(values):

            merge(row, offset)
            cell = ws.cell(row=row, column=positions[offset], value=value)
            cell.border = TOTAL_BORDER if is_total else CELL_BORDER

            if index % 2 == 0:
                for c in range(positions[offset], positions[offset] + spans.get(offset, 1)):
                    ws.cell(row=row, column=c).fill = band

            if is_total:
                cell.font = Font(bold=True)

            if offset in formats:
                cell.number_format = formats[offset]
            elif isinstance(value, bool):
                pass
            elif isinstance(value, int):
                cell.number_format = "#,##0"
            elif isinstance(value, float):
                cell.number_format = "#,##0.00"

            if offset in wrap_columns:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                if spans.get(offset, 1) > 1 and value:
                    chars = max(10, _span_width(ws, positions[offset], positions[offset] + spans[offset] - 1) * 1.1)
                    needed = 15 * math.ceil(len(str(value)) / chars) + 3
                    ws.row_dimensions[row].height = max(ws.row_dimensions[row].height or 15, needed)
            elif isinstance(value, (int, float)) and not isinstance(value, bool):
                cell.alignment = Alignment(horizontal="right", vertical="top")
            else:
                cell.alignment = Alignment(vertical="top")

    return top + len(rows)


def _empty_note(ws, row, text):

    ws.cell(row=row, column=1, value=text).font = Font(italic=True, color=MUTED)

    return row + 2


def _chart_rows(height_cm):

    return math.ceil(height_cm / ROW_HEIGHT_CM) + 1


def _bar_chart(
    ws, title, top, bottom, anchor, left=1, series_count=1,
    x_title=None, y_title=None, stacked=False, horizontal=False,
    width=17, height=9, labels=True, colors=None, point_colors=None,
    number_format="#,##0",
):

    """
    Bar chart over a table: categories in column `left`, series in
    the next `series_count` columns, header row = series names.
    Returns how many rows the chart covers.
    """

    chart = BarChart()
    chart.type = "bar" if horizontal else "col"
    chart.title = title
    chart.width = width
    chart.height = height
    chart.gapWidth = 50 if stacked else 70

    chart.add_data(
        Reference(ws, min_col=left + 1, max_col=left + series_count, min_row=top, max_row=bottom),
        titles_from_data=True,
    )
    chart.set_categories(Reference(ws, min_col=left, min_row=top + 1, max_row=bottom))

    colors = colors or PALETTE

    for index, series in enumerate(chart.series):
        color = colors[index % len(colors)]
        series.graphicalProperties.solidFill = color
        series.graphicalProperties.line.solidFill = color

    if point_colors and series_count == 1 and chart.series:
        for index, color in enumerate(point_colors):
            point = DataPoint(idx=index)
            point.graphicalProperties.solidFill = color
            point.graphicalProperties.line.solidFill = color
            chart.series[0].dPt.append(point)

    if stacked:
        chart.grouping = "stacked"
        chart.overlap = 100

    if series_count == 1:
        chart.legend = None
    else:
        chart.legend.position = "b"

    if labels and not stacked:
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showVal = True
        chart.dataLabels.showSerName = False
        chart.dataLabels.showCatName = False
        chart.dataLabels.showLegendKey = False
        chart.dataLabels.showPercent = False
        chart.dataLabels.position = "outEnd"
        chart.dataLabels.numFmt = number_format

    if x_title:
        chart.x_axis.title = x_title
    if y_title:
        chart.y_axis.title = y_title

    # openpyxl's defaults hide the axes in recent Excel builds.
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.scaling.min = 0
    chart.y_axis.number_format = number_format

    if horizontal:
        # Biggest at the top, value axis along the bottom.
        chart.x_axis.scaling.orientation = "maxMin"
        chart.y_axis.crosses = "max"

    ws.add_chart(chart, anchor)

    return _chart_rows(height)


def _line_chart(ws, title, top, bottom, anchor, series_count, x_title=None, y_title=None, width=24, height=10, tick_skip=5):

    chart = LineChart()
    chart.title = title
    chart.width = width
    chart.height = height

    chart.add_data(
        Reference(ws, min_col=2, max_col=1 + series_count, min_row=top, max_row=bottom),
        titles_from_data=True,
    )
    chart.set_categories(Reference(ws, min_col=1, min_row=top + 1, max_row=bottom))

    for index, series in enumerate(chart.series):
        series.graphicalProperties.line.solidFill = PALETTE[index % len(PALETTE)]
        series.graphicalProperties.line.width = 28575  # 2.25 pt
        series.smooth = False
        series.marker.symbol = "none"

    chart.legend.position = "b"

    if x_title:
        chart.x_axis.title = x_title
    if y_title:
        chart.y_axis.title = y_title

    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.scaling.min = 0
    chart.y_axis.number_format = "#,##0"
    chart.x_axis.tickLblSkip = tick_skip
    chart.x_axis.tickMarkSkip = tick_skip

    ws.add_chart(chart, anchor)

    return _chart_rows(height)


def _after(table_bottom, table_top, chart_rows):

    """First row of the next section: below the table or the
    chart, whichever ends lower, plus a gap."""

    return max(table_bottom, table_top + chart_rows) + 3


# ============================================================
# LOAD DATA
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


# ============================================================
# PATENTS SHEET
# ============================================================

PATENT_WIDTHS = {"A": 30, **{get_column_letter(c): 13 for c in range(2, 15)}}
PATENT_SPAN = 14

TOP_TABLE_COMMENTS = {
    "Times cited": "Number of later US patents and applications that list this patent as prior art (forward citations).",
    "Cited by applicant": "Citations the later document's own applicant made -- the strongest sign the field built on this patent.",
    "Cited by examiner": "Citations the patent examiner added during the search. Common, and partly reflect search-tool overlap.",
    "Cited by other / unknown": "SGML-era 'other' citations and all pre-1996 (APS) citations, which don't record who made them.",
    "Claim category": "What claim 1 claims: a process, an apparatus, a product, a composition ... Read from the claim wording.",
    "Patent type": "USPTO document type from the number: Utility, Design, Plant, Reissue, Application ...",
    "Citations per year": "Times cited divided by the years since grant (up to the latest citing year in the data).",
}


def _patent_table(ws, row, ranked, cites, extra_header=None):

    """Top-N patent table at `row`. `ranked` is a list of
    (patent_number, extra_value_or_None). Returns the last row."""

    patents, totals, by_category = cites["patents"], cites["totals"], cites["by_category"]

    # Column A holds "1.  US1234567B2"; Title spans 3 sheet
    # columns and Assignee 2, so the sheet's other tables can keep
    # narrow, even columns.
    header = [
        "Patent", "Year", "Title", "Assignee", "Patent type",
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
            f"{rank}.  {number}", patent.get("year"), patent.get("title", ""),
            patent.get("assignee", ""), patent.get("patent_type", ""),
            patent.get("claim_category", ""), count,
            split.get("applicant", 0), split.get("examiner", 0),
            count - split.get("applicant", 0) - split.get("examiner", 0),
        ]
        if extra_header:
            values.append(extra)

        rows.append(values)

    formats = {1: "0"}
    if extra_header:
        formats[len(header) - 1] = "0.00"

    return _write_table(
        ws, row, header, rows, formats=formats, wrap_columns=(2, 3, 4, 5),
        comments=TOP_TABLE_COMMENTS, spans={2: 3, 3: 2},
    )


def build_patents_sheet(workbook, notes, data, cites, stats):

    patents, file_counts, duplicates, classified = data
    totals = cites["totals"]
    total = len(patents)

    ws = workbook.create_sheet("Patents")
    ws.sheet_properties.tabColor = ACCENT
    _setup_sheet(ws, PATENT_WIDTHS)

    years = [p["year"] for p in patents.values() if p["year"] is not None]
    applications = sum(1 for p in patents.values() if p["is_application"])

    stats.update({
        "patents": total,
        "grants": total - applications,
        "applications": applications,
        "classified": classified,
        "first_year": min(years) if years else None,
        "last_year": max(years) if years else None,
    })

    span_text = f" from {min(years)} to {max(years)}" if years else ""

    row = _sheet_header(
        ws, "Patents",
        f"The {total:,} US patents and published applications{span_text} that "
        f"matched both the CPC codes and the keywords in the search matrix "
        f"({total - applications:,} granted patents, {applications:,} applications).",
        PATENT_SPAN,
    )

    # Claim categories present, in a fixed order.
    category_counts = Counter(p["claim_category"] for p in patents.values())
    categories = [c for c in CLAIM_CATEGORY_ORDER if category_counts.get(c)]
    categories += sorted(c for c in category_counts if c not in categories)

    decade_counts = defaultdict(Counter)
    undated = 0

    for patent in patents.values():
        if patent["year"] is None:
            undated += 1
            continue
        decade_counts[patent["year"] // 10 * 10][patent["claim_category"]] += 1

    decades = sorted(decade_counts)
    decade_totals = {d: sum(decade_counts[d].values()) for d in decades}

    # --- 1. patents by decade -----------------------------------

    takeaway = None
    if decades:
        busiest = max(decades, key=lambda d: decade_totals[d])
        takeaway = (
            f"The {busiest}s have the most patents: {decade_totals[busiest]:,} "
            f"({_pct(decade_totals[busiest], total)} of the corpus)."
        )

    top = _section(
        ws, row, "Patents by decade",
        "How many filtered patents fall in each decade, by grant date (publication date "
        "for applications). USPTO started publishing applications in 2001, so from the "
        "2000s on both granted patents and applications are counted.",
        PATENT_SPAN, takeaway,
    )

    bottom = _write_table(
        ws, top, ["Decade", "Patents", "Share of total"],
        [[f"{d}s", decade_totals[d], decade_totals[d] / total if total else 0] for d in decades],
        formats={2: "0%"},
    )

    chart_rows = 0
    if decades:
        chart_rows = _bar_chart(
            ws, "Patents by decade", top, bottom, f"E{top}",
            x_title="Decade", y_title="Patents", width=18, height=9,
        )
    else:
        bottom = _empty_note(ws, top + 1, "No dated patents found.")

    row = _after(bottom, top, chart_rows)

    # --- 2. by decade, split by claim category -------------------

    takeaway = None
    if decades and classified:
        latest = decades[-1]
        leader, count = decade_counts[latest].most_common(1)[0]
        takeaway = (
            f"In the {latest}s the most common claim type is {leader} "
            f"({_pct(count, decade_totals[latest])} of that decade)."
        )

    top = _section(
        ws, row, "Patents by decade, split by claim category",
        "The same decades, with each bar split by what the patent's first claim covers: "
        "a process or method, an apparatus or machine, a product or article, a "
        "composition, and so on. 'Unclassified' means the claim's wording didn't match a "
        "known category. See Start Here for definitions.",
        PATENT_SPAN, takeaway,
    )

    bottom = _write_table(
        ws, top, ["Decade"] + categories,
        [[f"{d}s"] + [decade_counts[d].get(c, 0) for c in categories] for d in decades],
    )

    if decades:
        chart_top = bottom + 2
        chart_rows = _bar_chart(
            ws, "Patents by decade and claim category", top, bottom, f"A{chart_top}",
            series_count=len(categories), stacked=True, x_title="Decade",
            y_title="Patents", width=28, height=11,
            colors=[_category_color(c) for c in categories],
        )
        row = chart_top + chart_rows + 2
    else:
        row = bottom + 3

    # --- 3. by claim category -----------------------------------

    takeaway = None
    if categories and classified:
        leader = max(categories, key=lambda c: category_counts[c])
        takeaway = f"{leader} is the most common claim type: {category_counts[leader]:,} patents ({_pct(category_counts[leader], total)})."
        if category_counts.get("Unclassified"):
            takeaway += (
                f" {_pct(category_counts['Unclassified'], total)} couldn't be classified "
                f"from the claim wording."
            )
    elif not classified:
        takeaway = "Claim categories are missing: run uspto_patent_pipeline_v15.py --classify, then re-run this report."

    top = _section(
        ws, row, "Patents by claim category",
        "What the patents claim, read from the opening words of claim 1 (e.g. 'A method "
        "of...' is a Process / Method, 'A multilayer film...' is a Product / Article). "
        "USPTO doesn't record this, so it's worked out from the claim text.",
        PATENT_SPAN, takeaway,
    )

    bottom = _write_table(
        ws, top, ["Claim category", "Patents", "Share of total"],
        [[c, category_counts[c], category_counts[c] / total if total else 0] for c in categories],
        formats={2: "0%"},
    )

    chart_rows = 0
    if categories:
        chart_rows = _bar_chart(
            ws, "Patents by claim category", top, bottom, f"E{top}",
            y_title="Patents", width=18, height=9,
            point_colors=[_category_color(c) for c in categories],
        )

    row = _after(bottom, top, chart_rows)

    # --- 4. by file type ----------------------------------------

    rows_total = sum(count for _, count in file_counts)
    biggest = max(file_counts, key=lambda item: item[1]) if file_counts else None

    top = _section(
        ws, row, "Patents by file type",
        "USPTO publishes its data in a different format for each era: APS text files for "
        "1976-2001 grants, SGML for 2001 grants, XML for grants from 2002 on, and XML for "
        "published applications from 2001 on. This shows how many filter results came "
        f"from each. {duplicates:,} patents appear in both the 2001 APS and SGML files; "
        "they are counted once everywhere else in this report.",
        PATENT_SPAN,
        (
            f"Most results come from {biggest[0]}: {biggest[1]:,} ({_pct(biggest[1], rows_total)})."
            if biggest and biggest[1] else None
        ),
    )

    bottom = _write_table(
        ws, top, ["File type", "Patents", "Share of total"],
        [[label, count, count / rows_total if rows_total else 0] for label, count in file_counts],
        formats={2: "0%"},
    )

    chart_rows = _bar_chart(
        ws, "Patents by file type", top, bottom, f"E{top}",
        y_title="Patents", width=18, height=9,
    )

    row = _after(bottom, top, chart_rows)

    # --- 5. top 10 by forward citations -------------------------

    ranked = totals.most_common(TOP_N_PATENTS)

    takeaway = None
    if ranked:
        number, count = ranked[0]
        title = patents.get(number, {}).get("title", "")
        takeaway = f"The most-cited patent is {number}" + (f", \"{title}\"" if title else "") + f", cited {count:,} times."

    top = _section(
        ws, row, f"Top {TOP_N_PATENTS} patents by forward citations",
        "The filtered patents that later US patents and applications cite most. A forward "
        "citation means a later document lists this patent as prior art, so a high count "
        "suggests the field built on it. Hover over a column heading for what it means. "
        "Older patents have had longer to collect citations; the Citations sheet adjusts "
        "for age.",
        PATENT_SPAN, takeaway,
    )

    if ranked:
        bottom = _patent_table(ws, top, [(number, None) for number, _ in ranked], cites)
        row = bottom + 3
    else:
        row = _empty_note(ws, top, "No forward-citation pairs found -- run uspto_forward_citations_v5.py --export first.") + 1

    # --- 6. per year: grants vs applications --------------------

    year_counts = defaultdict(Counter)

    for patent in patents.values():
        if patent["year"] is not None:
            year_counts[patent["year"]]["Applications" if patent["is_application"] else "Grants"] += 1

    takeaway = None
    if year_counts:
        peak = max(year_counts, key=lambda y: sum(year_counts[y].values()))
        takeaway = f"The busiest year is {peak}, with {sum(year_counts[peak].values()):,} documents."

    top = _section(
        ws, row, "Patents per year: grants and applications",
        "Year-by-year detail behind the decade charts, with granted patents and published "
        "applications as separate lines. An invention published as an application and "
        "later granted appears once on each line.",
        PATENT_SPAN, takeaway,
    )

    year_range = range(min(year_counts), max(year_counts) + 1) if year_counts else []

    bottom = _write_table(
        ws, top, ["Year", "Grants", "Applications"],
        [[str(year), year_counts[year]["Grants"], year_counts[year]["Applications"]] for year in year_range],
    )

    chart_rows = 0
    if year_counts:
        chart_rows = _line_chart(
            ws, "Patents per year", top, bottom, f"E{top}", series_count=2,
            x_title="Year", y_title="Documents", width=26, height=11,
        )

    row = _after(bottom, top, chart_rows)

    # --- 7. top CPC subclasses ----------------------------------

    subclass_counts = Counter()
    subclass_cites = Counter()

    for number, patent in patents.items():
        for subclass in patent["subclasses"]:
            subclass_counts[subclass] += 1
            subclass_cites[subclass] += totals.get(number, 0)

    top_subclasses = subclass_counts.most_common(TOP_N_CPC)

    takeaway = None
    if top_subclasses:
        code, count = top_subclasses[0]
        name = CPC_SUBCLASS_TITLES.get(code)
        takeaway = f"{code}" + (f" ({name})" if name else "") + f" appears on {count:,} patents ({_pct(count, total)})."

    top = _section(
        ws, row, f"Top {TOP_N_CPC} CPC subclasses",
        "CPC (Cooperative Patent Classification) sorts patents by technology. A subclass is "
        "a 4-character code such as B65H, covering a broad technical area. Most patents "
        "carry several subclasses, so the shares add up to more than 100%. Patents from "
        "before 2013 that have no CPC codes use their IPC codes instead.",
        PATENT_SPAN, takeaway,
    )

    bottom = _write_table(
        ws, top, ["Subclass", "Patents", "Share of patents", "Forward citations", "What it covers"],
        [
            [
                code, count, count / total if total else 0, subclass_cites[code],
                CPC_SUBCLASS_TITLES.get(code, ""),
            ]
            for code, count in top_subclasses
        ],
        formats={2: "0%"}, wrap_columns=(4,), spans={4: 4},
    )

    chart_rows = 0
    if top_subclasses:
        chart_rows = _bar_chart(
            ws, f"Top {TOP_N_CPC} CPC subclasses", top, bottom, f"J{top}",
            y_title="Patents", horizontal=True, width=16, height=12,
        )

    ws.freeze_panes = "A4"

    notes.append(["PATENTS"])
    notes.append(["Filter results read from", str(FILTERED_DIR)])
    notes.append(["Unique patents/applications", total])
    notes.append(["Rows listed in two product files (2001 APS + SGML), counted once", duplicates])
    notes.append(["Patents with no parseable date (left out of the decade and year charts)", undated])
    notes.append(["Decade and year are taken from the grant date (publication date for applications).", ""])
    notes.append([
        "Applications and grants are separate documents, so an invention published "
        "as an application and later granted counts in both.", "",
    ])
    if not classified:
        notes.append([
            "Claim category is missing from the filter results -- run "
            "uspto_patent_pipeline_v15.py --classify, then re-run this report.", "",
        ])
    notes.append([
        "CPC subclasses: a patent counts once in each subclass it carries (IPC used "
        "for records with no CPC); shares add to more than 100%.", "",
    ])
    notes.append([])


# ============================================================
# CITATIONS SHEET
# ============================================================

def build_citations_sheet(workbook, notes, data, cites, stats):

    patents = data[0]
    totals = cites["totals"]
    coverage_year = cites["coverage_year"]
    total = len(patents)
    cited = sum(1 for n in patents if totals.get(n))

    stats.update({
        "citations": sum(totals.values()),
        "cited_patents": cited,
    })

    ws = workbook.create_sheet("Citations")
    ws.sheet_properties.tabColor = ACCENT
    _setup_sheet(ws, PATENT_WIDTHS)

    row = _sheet_header(
        ws, "Forward citations",
        f"How often later US patents and applications cite the filtered patents: "
        f"{sum(totals.values()):,} citations to {cited:,} of the {total:,} filtered "
        f"patents, counting citing documents up to {coverage_year}.",
        PATENT_SPAN,
    )

    # --- 1. distribution ----------------------------------------

    bucket_counts = Counter()

    for number in patents:
        count = totals.get(number, 0)
        for label, low, high in CITATION_BUCKETS:
            if count >= low and (high is None or count <= high):
                bucket_counts[label] += 1
                break

    heavy = bucket_counts["101-500"] + bucket_counts["Over 500"]

    top = _section(
        ws, row, "How often the filtered patents are cited",
        "Each filtered patent is placed in a group by how many later documents cite it. "
        "Most patents are cited a few times or not at all; a small number are cited "
        "heavily. Recent patents haven't had time to be cited yet.",
        PATENT_SPAN,
        (
            f"{_pct(cited, total)} have been cited at least once, and {heavy:,} have "
            f"been cited more than 100 times." if total else None
        ),
    )

    bottom = _write_table(
        ws, top, ["Times cited", "Patents", "Share of total"],
        [
            [label, bucket_counts[label], bucket_counts[label] / total if total else 0]
            for label, _, _ in CITATION_BUCKETS
        ],
        formats={2: "0%"},
    )

    chart_rows = _bar_chart(
        ws, "Filtered patents by times cited", top, bottom, f"E{top}",
        x_title="Times cited", y_title="Patents", width=18, height=9,
    )

    row = _after(bottom, top, chart_rows)

    # --- 2. top 10 by citations per year ------------------------

    rates = []

    for number, count in totals.items():
        patent = patents.get(number)
        if not patent or patent["year"] is None or count < CITATION_RATE_MIN_CITES:
            continue
        years = max(CITATION_RATE_MIN_YEARS, coverage_year - patent["year"])
        rates.append((count / years, number))

    rates.sort(reverse=True)

    takeaway = None
    if rates:
        rate, number = rates[0]
        takeaway = (
            f"{number} is being cited fastest: about {rate:,.1f} times a year since "
            f"{patents[number]['year']}."
        )

    top = _section(
        ws, row, f"Top {TOP_N_PATENTS} fastest-cited patents (citations per year)",
        "Raw citation counts favour old patents, which have had decades to collect "
        "citations. This table divides each patent's citations by the years since it was "
        f"granted, which brings newer, quickly adopted patents to the top. Only patents "
        f"with at least {CITATION_RATE_MIN_CITES} citations are included.",
        PATENT_SPAN, takeaway,
    )

    if rates:
        bottom = _patent_table(
            ws, top, [(number, round(rate, 2)) for rate, number in rates[:TOP_N_PATENTS]],
            cites, extra_header="Citations per year",
        )
        row = bottom + 3
    else:
        row = _empty_note(ws, top, "No patents with enough citations yet.") + 1

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

    takeaway = None
    real = [c for c in categories if c not in ("Unclassified", "No claims text", NOT_CLASSIFIED)]
    if real:
        leader = max(real, key=lambda c: category_cites[c] / category_patents[c])
        takeaway = (
            f"{leader} patents are cited most on average: "
            f"{category_cites[leader] / category_patents[leader]:.1f} citations per patent."
        )

    top = _section(
        ws, row, "Average citations by claim category",
        "Whether process, apparatus, product or composition patents get cited more in "
        "this field. Average = total citations to patents in the category divided by "
        "the number of patents in it.",
        PATENT_SPAN, takeaway,
    )

    bottom = _write_table(
        ws, top,
        ["Claim category", "Average citations per patent", "Patents", "Cited at least once", "Total citations"],
        [
            [
                c, category_cites[c] / category_patents[c],
                category_patents[c], category_cited[c], category_cites[c],
            ]
            for c in categories
        ],
        formats={1: "0.0"},
    )

    chart_rows = 0
    if categories:
        chart_rows = _bar_chart(
            ws, "Average citations per patent", top, bottom, f"G{top}",
            y_title="Citations per patent", width=18, height=9,
            point_colors=[_category_color(c) for c in categories], number_format="0.0",
        )

    row = _after(bottom, top, chart_rows)

    # --- 4. who cites them --------------------------------------

    citing = cites["citing_types"]
    total_citing = sum(citing.values())

    takeaway = None
    if citing:
        label, count = citing.most_common(1)[0]
        takeaway = f"Most citing documents are {label.lower()} documents ({_pct(count, total_citing)})."

    top = _section(
        ws, row, "What kinds of documents cite them",
        "The later documents that cite the filtered patents, by type. Each citing "
        "document is counted once, however many filtered patents it cites. Published "
        "applications often carry no citation list, so most citing documents are "
        "granted patents.",
        PATENT_SPAN, takeaway,
    )

    bottom = _write_table(
        ws, top, ["Citing document type", "Citing documents", "Share of total"],
        [[label, count, count / total_citing if total_citing else 0] for label, count in citing.most_common()],
        formats={2: "0%"},
    )

    chart_rows = 0
    if citing:
        chart_rows = _bar_chart(
            ws, "Citing documents by type", top, bottom, f"E{top}",
            y_title="Documents", width=18, height=9,
        )

    ws.freeze_panes = "A4"

    notes.append(["CITATIONS"])
    notes.append(["Forward citations counted from", ", ".join(cites["files"]) or "(none found)"])
    notes.append(["Filtered patents cited at least once", cited])
    notes.append([
        "Citations per year = times cited / years from grant to "
        f"{coverage_year} (the latest citing-patent year), at least "
        f"{CITATION_RATE_MIN_YEARS:g} year; patents with fewer than "
        f"{CITATION_RATE_MIN_CITES} citations are left out of that table.", "",
    ])
    notes.append(["Citing documents are counted once each, however many filtered patents they cite.", ""])
    notes.append([])


# ============================================================
# CRAWL SHEETS
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


def _category_name(folder):

    """'4_Polymer_Films' -> 'Polymer Films'."""

    return re.sub(r"^\d+\s*", "", (folder or "").replace("_", " ")).strip()


def _language_name(code):

    parts = [p.strip().lower() for p in (code or "unknown").split("+") if p.strip()]

    return " + ".join(LANGUAGE_NAMES.get(p, p) for p in parts) or "Unknown"


CRAWL_WIDTHS = {
    "A": 40, "B": 22, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14,
    "H": 14, "I": 60,
}
CRAWL_SPAN = 9

VENDOR_COMMENTS = {
    "Seed status": "From vendor_seeds.csv: ok / verify are crawled, skip is not.",
    "Pages crawled": "Web pages on the vendor's site fetched successfully.",
    "Pages queued or seen": "All pages the crawler knows about, including ones still queued, blocked or errored.",
    "PDF links found": "Links to PDF files found on the crawled pages.",
    "PDF links downloaded": "Links whose file was saved, was the same file as another link, or was unchanged since the last run.",
    "Documents kept": "PDFs kept in the vendor's folder: not rejected and not a foreign-language copy.",
    "Translations": "Foreign-language copies kept in the _translations folder.",
    "Rejected": "Files moved to _rejected (unreadable, or judged not useful).",
    "Total PDF pages": "Pages across the kept and translated files.",
}


def build_crawl_sheets(workbook, notes, stats):

    print("\nCRAWL")

    if not TDS_MANIFEST.exists():
        ws = workbook.create_sheet("Crawl")
        _setup_sheet(ws, CRAWL_WIDTHS)
        _sheet_header(ws, "Data-sheet crawl", f"{TDS_MANIFEST} not found -- run tds_crawler_v1.py --export first.", CRAWL_SPAN)
        return {}

    with open(TDS_MANIFEST, "r", newline="", encoding="utf-8-sig") as handle:
        documents = list(csv.DictReader(handle))

    print(f"  {TDS_MANIFEST.name}: {len(documents):,} files")

    connection = _open_tds_db()

    pages = _db_counts(connection, "SELECT vendor, state, COUNT(*) FROM pages GROUP BY vendor, state")
    links = _db_counts(connection, "SELECT vendor, state, COUNT(*) FROM pdf_links GROUP BY vendor, state")

    vendor_info = {}

    db_read = connection is not None

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

    def doc_type_label(doc):
        raw = (doc.get("doc_type_from_text") or "").strip()
        return DOC_TYPE_LABEL.get(raw.lower(), raw or "Unknown")

    def as_int(value):
        try:
            return int(float(value or 0))
        except ValueError:
            return 0

    def category_label(doc):
        return _category_name(doc.get("category_folder")) or "Unassigned"

    kept = [d for d in documents if kind(d) == "kept"]
    translations = [d for d in documents if kind(d) == "translation"]
    rejected = [d for d in documents if kind(d) == "rejected"]

    category_counts = Counter(category_label(d) for d in kept)
    type_counts = Counter(doc_type_label(d) for d in kept)

    family_counts = Counter()
    for d in kept:
        for family in re.split(r"[;,]", d.get("method_families") or ""):
            family = family.strip()
            if family:
                family_counts[family] += 1

    with_methods = sum(1 for d in kept if (d.get("test_methods") or "").strip())

    language_counts = Counter(
        _language_name((d.get("language") or "").strip()) for d in kept + translations
    )

    # --- per-vendor table ---------------------------------------

    vendors = defaultdict(lambda: {
        "category": "", "kept": 0, "tds": 0, "sds": 0, "other_types": 0,
        "translations": 0, "rejected": 0, "pages": 0, "bytes": 0,
        "latest": "",
    })

    for d in documents:
        v = vendors[d.get("vendor") or "(unknown)"]
        v["category"] = v["category"] or _category_name(d.get("category_folder"))
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

    total_pages_crawled = sum(c.get("done", 0) for c in pages.values())
    total_links = sum(sum(c.values()) for c in links.values())
    link_states = Counter()
    for c in links.values():
        link_states.update(c)

    vendors_with_docs = sum(1 for r in vendor_rows if r[8])

    stats.update({
        "vendors": len(vendor_rows),
        "vendors_with_docs": vendors_with_docs,
        "docs_kept": len(kept),
        "data_sheets": type_counts.get(DOC_TYPE_LABEL["tds"], 0),
    })

    # --- Crawl sheet --------------------------------------------

    ws = workbook.create_sheet("Crawl")
    ws.sheet_properties.tabColor = "70AD47"
    _setup_sheet(ws, CRAWL_WIDTHS)

    row = _sheet_header(
        ws, "Data-sheet crawl",
        f"Supplier websites crawled for technical data sheets and related documents: "
        f"{len(vendor_rows):,} vendors, {len(kept):,} documents kept in English or their "
        f"original language, plus {len(translations):,} foreign-language copies.",
        CRAWL_SPAN,
    )

    # --- 1. headline numbers ------------------------------------

    top = _section(
        ws, row, "Headline numbers",
        "The crawl from start to finish: pages visited on supplier websites, PDF links "
        "found on them, files downloaded, and what was kept. Duplicate files (the same "
        "PDF at two links) are stored once.",
        CRAWL_SPAN,
        f"{vendors_with_docs:,} of {len(vendor_rows):,} vendors have at least one kept document.",
    )

    headline = [
        ["Vendors in the crawl", len(vendor_rows)],
        ["Vendors with at least one kept document", vendors_with_docs],
        ["Web pages crawled", total_pages_crawled],
        ["PDF links found", total_links],
        ["Unique PDF files downloaded", len(documents)],
        ["Documents kept (English or original language)", len(kept)],
        ["Foreign-language copies (in _translations folders)", len(translations)],
        ["Rejected files (in _rejected folders)", len(rejected)],
        ["Kept documents that cite a test method", with_methods],
        ["Pages in kept and translated files", sum(as_int(d.get("pages")) for d in kept + translations)],
        ["Size of kept and translated files (GB)", round(sum(as_int(d.get("bytes")) for d in kept + translations) / 1_073_741_824, 2)],
    ]

    bottom = _write_table(ws, top, ["Measure", "Value"], headline)
    ws.cell(row=bottom, column=2).number_format = "#,##0.00"

    row = bottom + 3

    # --- 2. link outcomes ---------------------------------------

    top = _section(
        ws, row, "What happened to each PDF link",
        "Every PDF link the crawler found, by outcome. 'Same file as another link' means "
        "the file was already saved from a different address.",
        CRAWL_SPAN,
    )

    outcome_rows = [
        [LINK_STATE_LABEL.get(state, state), count, count / total_links if total_links else 0]
        for state, count in link_states.most_common()
    ]

    bottom = _write_table(ws, top, ["Outcome", "Links", "Share of links"], outcome_rows, formats={2: "0%"})

    chart_rows = 0
    if outcome_rows:
        chart_rows = _bar_chart(
            ws, "PDF links by outcome", top, bottom, f"E{top}",
            y_title="Links", horizontal=True, width=17, height=8,
        )

    row = _after(bottom, top, chart_rows)

    # --- chart blocks -------------------------------------------

    def block(start_row, title, what, takeaway, header, rows, chart_title, share=True, horizontal=False, height=9):

        top = _section(ws, start_row, title, what, CRAWL_SPAN, takeaway)

        whole = sum(r[1] for r in rows)

        table_rows = [list(r) + ([r[1] / whole if whole else 0] if share else []) for r in rows]

        bottom = _write_table(
            ws, top, header + (["Share"] if share else []), table_rows,
            formats={2: "0%"} if share else {},
        )

        if not rows:
            return _empty_note(ws, top + 1, "Nothing to show yet.") + 1

        chart_rows = _bar_chart(
            ws, chart_title, top, bottom, f"E{top}", y_title=header[1],
            horizontal=horizontal, width=18, height=height,
        )

        return _after(bottom, top, chart_rows)

    leader = category_counts.most_common(1)
    row = block(
        row, "Documents by material category",
        "Kept documents by the material category the vendor is listed under in the "
        "vendor workbook (foil, nonwovens, paper, films, converting, equipment, process).",
        f"{leader[0][0]} has the most documents: {leader[0][1]:,}." if leader else None,
        ["Material category", "Documents"], sorted(category_counts.items()),
        "Documents by material category",
    )

    tds = type_counts.get(DOC_TYPE_LABEL["tds"], 0)
    row = block(
        row, "Documents by type",
        "What kind of document each kept file is, read from its first pages. Technical "
        "data sheets list material properties; safety data sheets cover handling and "
        "hazards; brochures and selector guides compare product ranges.",
        f"{_pct(tds, len(kept))} of kept documents are technical data sheets ({tds:,})." if kept else None,
        ["Document type", "Documents"], type_counts.most_common(),
        "Documents by type", horizontal=True,
    )

    top_vendors = [[r[0], r[8]] for r in vendor_rows[:TOP_N_VENDORS] if r[8]]
    row = block(
        row, f"Top {TOP_N_VENDORS} vendors by documents kept",
        "The suppliers with the most kept documents. The full list, with crawl details "
        "for every vendor, is on the Vendors sheet.",
        f"{top_vendors[0][0]} has the most: {top_vendors[0][1]:,} documents." if top_vendors else None,
        ["Vendor", "Documents"], top_vendors,
        f"Top {TOP_N_VENDORS} vendors by documents kept", share=False, horizontal=True, height=12,
    )

    family = family_counts.most_common(1)
    row = block(
        row, "Test-method families cited",
        "Data sheets that cite standard test methods (ASTM, ISO, TAPPI, DIN ...) let you "
        "compare property values across suppliers. A document citing both ASTM and ISO "
        "methods counts under both, so shares add to more than 100%.",
        (
            f"{_pct(with_methods, len(kept))} of kept documents cite at least one test "
            f"method; {family[0][0]} is the most common ({family[0][1]:,} documents)."
            if family else None
        ),
        ["Method family", "Documents"], family_counts.most_common(),
        "Test-method families cited", share=False,
    )

    english = sum(count for name, count in language_counts.items() if "English" in name)
    row = block(
        row, "Documents by language",
        "Language of the kept documents and foreign-language copies, read from the PDF "
        "text. A foreign-language copy is kept aside in _translations when the English "
        "version is in the folder.",
        f"{_pct(english, sum(language_counts.values()))} include English." if language_counts else None,
        ["Language", "Documents"], language_counts.most_common(),
        "Documents by language",
    )

    # --- material category x document type ----------------------

    type_order = [label for label, _ in type_counts.most_common()]
    matrix = defaultdict(Counter)

    for d in kept:
        matrix[category_label(d)][doc_type_label(d)] += 1

    top = _section(
        ws, row, "Documents by material category and type",
        "How each material category's documents break down by type. Read across a row "
        "for one category; the bottom row is the total for each type.",
        CRAWL_SPAN,
    )

    matrix_rows = [
        [category] + [matrix[category].get(t, 0) for t in type_order] + [sum(matrix[category].values())]
        for category in sorted(matrix)
    ]
    matrix_rows.append(["Total"] + [type_counts[t] for t in type_order] + [len(kept)])

    bottom = _write_table(
        ws, top, ["Material category"] + type_order + ["Total"], matrix_rows, total_row=True,
    )

    row = bottom + 4

    # --- vendors with no kept documents -------------------------

    def no_document_reason(r):
        name, status, crawled, seen, link_count, downloaded = r[0], r[2], r[3], r[4], r[5], r[6]
        page_states = pages.get(name, Counter())
        if status == "skip":
            return "Skipped: status is 'skip' in vendor_seeds.csv"
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
                f"{LINK_STATE_LABEL.get(state, state).lower()}: {count}"
                for state, count in links.get(name, Counter()).most_common(3)
            )
            return f"PDF links found but none downloaded ({outcomes})"
        if r[12] and not r[13]:
            return f"Only foreign-language copies ({r[12]}), no English version kept"
        if r[13]:
            return f"All downloaded files rejected ({r[13]})"
        return "Downloads weren't kept: off-topic files are deleted (see logs\\deleted_off_topic_*.csv)"

    empty = sorted((r for r in vendor_rows if not r[8]), key=lambda r: r[0].lower())

    top = _section(
        ws, row, f"Vendors with no kept documents ({len(empty)})",
        "Suppliers the crawl hasn't collected anything from yet, with the most likely "
        "reason. Use this as a to-do list: a blocked site may need a different start "
        "page in vendor_seeds.csv, and 'no PDF links' often means the data sheets sit "
        "behind a search or login.",
        CRAWL_SPAN,
        "Every vendor has at least one kept document." if not empty else None,
    )

    if empty:
        _write_table(
            ws, top,
            ["Vendor", "Material category", "Seed status", "Pages crawled",
             "PDF links found", "PDF links downloaded", "Translations", "Rejected", "Likely reason"],
            [[r[0], r[1], r[2], r[3], r[5], r[6], r[12], r[13], no_document_reason(r)] for r in empty],
            wrap_columns=(8,), comments=VENDOR_COMMENTS,
        )

    ws.freeze_panes = "A4"

    # --- Vendors sheet ------------------------------------------

    vs = workbook.create_sheet("Vendors")
    vs.sheet_properties.tabColor = "70AD47"

    header = [
        "Vendor", "Material category", "Seed status", "Pages crawled",
        "Pages queued or seen", "PDF links found", "PDF links downloaded",
        "PDF links errored", "Documents kept", "Technical data sheets",
        "Safety data sheets", "Other document types", "Translations",
        "Rejected", "Total PDF pages", "Size (MB)", "Last download",
    ]

    widths = {get_column_letter(i): 12 for i in range(3, len(header) + 1)}
    widths.update({"A": 36, "B": 24, "Q": 14})
    _setup_sheet(vs, widths)

    _sheet_header(
        vs, "Vendors",
        "One row per supplier in the crawl, most documents first. Use the filter arrows "
        "in the header row to sort or narrow the list; hover over a heading to see what "
        "it counts.",
        len(header),
    )

    table_top = 5
    bottom = _write_table(
        vs, table_top, header, vendor_rows, formats={15: "#,##0.0"}, comments=VENDOR_COMMENTS,
    )

    if vendor_rows:
        vs.auto_filter.ref = f"A{table_top}:{get_column_letter(len(header))}{bottom}"
        vs.conditional_formatting.add(
            f"I{table_top + 1}:I{bottom}",
            DataBarRule(start_type="num", start_value=0, end_type="max", color="5B9BD5"),
        )

    vs.freeze_panes = f"B{table_top + 1}"

    notes.append(["CRAWL"])
    notes.append(["Manifest read", str(TDS_MANIFEST)])
    notes.append(["Crawl database read (read-only)", str(TDS_DB) if db_read else "not found / not readable"])
    notes.append([
        "'Documents kept' = files not in _rejected and not a foreign-language copy; "
        "type is the type read from the PDF text.", "",
    ])
    notes.append([
        "'PDF links downloaded' includes links whose file was a duplicate of one "
        "already saved or unchanged since the last run.", "",
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

# A one-word vendor name this short would match too many unrelated
# assignees; such names need a second word.
_MIN_SINGLE_WORD_LENGTH = 3


def _company_words(name):

    text = (name or "").lower().replace("&", " ")
    text = re.sub(r"[^\w\s]", " ", text)

    return [w for w in text.split() if w and w not in _CORPORATE_WORDS]


def build_vendor_patent_sheet(workbook, notes, data, cites, vendor_summary, stats):

    print("\nVENDORS & PATENTS")

    patents = data[0]
    totals = cites["totals"]

    ws = workbook.create_sheet("Vendors & Patents")
    ws.sheet_properties.tabColor = "7030A0"
    _setup_sheet(ws, {
        "A": 34, "B": 24, "C": 12, "D": 12, "E": 12, "F": 12, "G": 12,
        "H": 12, "I": 12, "J": 16, "K": 46, "L": 50,
    })
    span = 12

    if not vendor_summary:
        _sheet_header(ws, "Vendors & Patents", "No crawl manifest found, so there are no vendors to match.", span)
        stats["vendors_with_patents"] = 0
        return

    by_first_word = defaultdict(list)

    for vendor in vendor_summary:
        words = _company_words(vendor)
        if not words:
            continue
        if len(words) == 1 and len(words[0]) < _MIN_SINGLE_WORD_LENGTH and not words[0][0].isdigit():
            continue
        by_first_word[words[0]].append((vendor, words))

    matched = defaultdict(list)
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
        top_patent = max(numbers, key=lambda n: totals.get(n, 0))

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
            top_patent if totals.get(top_patent) else "",
            patents[top_patent]["title"] if totals.get(top_patent) else "",
            "; ".join(name for name, _ in assignee_names[vendor].most_common(3)),
        ])

    rows.sort(key=lambda r: (-r[4], r[0].lower()))

    stats["vendors_with_patents"] = len(rows)

    row = _sheet_header(
        ws, "Vendors & Patents",
        f"Links the two halves of the research: {len(rows)} of the {len(vendor_summary)} "
        f"suppliers in the data-sheet crawl also own patents in the filtered patent set.",
        span,
    )

    top = _section(
        ws, row, f"Top {TOP_N_VENDORS} vendors by filtered patents",
        "Suppliers that both publish data sheets and hold patents in this field, ranked "
        "by number of filtered patents (grants and applications).",
        span,
        f"{rows[0][0]} holds the most: {rows[0][4]:,} filtered patents." if rows else None,
    )

    bottom = _write_table(
        ws, top, ["Vendor", "Filtered patents"], [[r[0], r[4]] for r in rows[:TOP_N_VENDORS]],
    )

    chart_rows = 0
    if rows:
        chart_rows = _bar_chart(
            ws, f"Top {TOP_N_VENDORS} vendors by filtered patents", top, bottom, f"D{top}",
            y_title="Patents", horizontal=True, width=20, height=12,
        )
    else:
        bottom = _empty_note(ws, top + 1, "No vendor names matched any patent assignee.")

    row = _after(bottom, top, chart_rows)

    top = _section(
        ws, row, "All matched vendors",
        "Every supplier matched to a patent assignee, with its data sheets and patents "
        "side by side. Vendors are matched to assignees by name, which is approximate: "
        "check the last column to confirm each match (for example, 'Dow' also matches "
        "'Dow Corning'). A subsidiary under a different name won't be matched.",
        span,
    )

    header = [
        "Vendor", "Material category", "Documents kept", "Technical data sheets",
        "Filtered patents", "Grants", "Applications", "Patents cited at least once",
        "Total forward citations", "Most-cited patent", "Most-cited patent title",
        "Assignee names matched (top 3)",
    ]

    bottom = _write_table(ws, top, header, rows, wrap_columns=(10, 11))

    if rows:
        ws.auto_filter.ref = f"A{top}:{get_column_letter(len(header))}{bottom}"

    ws.freeze_panes = "A4"

    print(f"  {len(rows)} vendors matched to assignees")

    notes.append(["VENDORS & PATENTS"])
    notes.append([
        "Vendors are matched to patent assignees by name: corporate words (Inc, Corp, "
        "GmbH, Company ...) are dropped and the assignee must begin with the vendor's "
        "remaining words, so '3M Company' matches '3M Innovative Properties Company'.", "",
    ])
    notes.append([
        "It's approximate: a subsidiary under a different name is missed, and a shared "
        "first name (e.g. 'Dow' / 'Dow Corning') can match. Check the 'Assignee names "
        "matched' column.", "",
    ])
    notes.append([])


# ============================================================
# START HERE + NOTES
# ============================================================

SHEET_GUIDE = [
    ("Patents", "The filtered patent corpus: patents by decade (in total and by claim category), by claim category, by file type, the 10 most-cited patents, patents per year, and the most common CPC technology areas."),
    ("Citations", "How often later patents cite the filtered ones: the spread of citation counts, the fastest-cited patents (adjusted for age), which claim types get cited most, and what kinds of documents do the citing."),
    ("Crawl", "The supplier data-sheet crawl: headline numbers, what happened to each PDF link, documents by material category, type, vendor, test method and language, and the vendors with nothing collected yet."),
    ("Vendors", "One filterable row per supplier with its crawl and document counts."),
    ("Vendors & Patents", "Suppliers in the crawl that also own patents in the filtered set, with data sheets and patents side by side."),
    ("How It Was Counted", "Where every number comes from and the counting rules, for checking or repeating the analysis."),
]

GLOSSARY = [
    ("Filtered patent", "A US patent or published application that matched both a CPC code and a keyword from the search matrix workbooks."),
    ("Granted patent vs. application", "An application is published about 18 months after filing, before it is examined. A granted patent has passed examination. The same invention can appear as both."),
    ("Forward citation", "A later patent or application that lists this patent as prior art. More forward citations generally means more influence on later work."),
    ("Cited by applicant / examiner", "Who added the citation to the later document: the applicant (a deliberate reference, the stronger signal) or the patent examiner (found during the search)."),
    ("Patent type", "The USPTO document type: Utility (most patents), Design (appearance), Plant, Reissue (a corrected re-issue of an earlier patent), or Application."),
    ("Claim category", "What the patent's first claim protects: a Process / Method, an Apparatus / System, a Product / Article, a Composition (a material or mixture), or a Product-by-process. Read from the claim wording, so 'Unclassified' means the wording wasn't recognised."),
    ("CPC subclass", "Cooperative Patent Classification: a code such as B65H (web handling) or D04H (nonwovens) that places a patent in a technology area."),
    ("APS / SGML / XML", "The formats USPTO used to publish its bulk data in different eras: APS text (1976-2001), SGML (2001), XML (2002 on, and applications from 2001)."),
    ("Technical / safety data sheet", "A technical data sheet (TDS) lists a material's properties and test results; a safety data sheet (SDS) covers handling, hazards and storage."),
    ("Kept / translation / rejected", "Kept: the document is in the vendor's folder. Translation: a foreign-language copy set aside because the English version is kept. Rejected: unreadable or not useful, moved to _rejected."),
]


def build_start_sheet(ws, stats):

    _setup_sheet(ws, {"A": 2, **{get_column_letter(c): 19 for c in range(2, 10)}}, landscape=False)
    ws.sheet_properties.tabColor = NAVY

    span_first, span_last = 2, 9

    ws.cell(row=1, column=2, value="Web Films research overview").font = Font(bold=True, size=22, color=NAVY)
    ws.row_dimensions[1].height = 36

    _text_row(
        ws, 2,
        "A summary of the USPTO patent corpus (web handling, converting, films, foil, "
        "nonwovens and paper) and the supplier data-sheet crawl. Start with the key "
        f"numbers below, then follow the links to each sheet. Generated {datetime.now():%B %d, %Y at %H:%M}.",
        span_first, span_last, font=Font(size=11, italic=True, color=MUTED),
    )

    row = 4

    if not stats.get("classified", True):
        _text_row(
            ws, row,
            "Claim categories are missing from the filter results. Run "
            "uspto_patent_pipeline_v15.py --classify, then run this report again.",
            span_first, span_last, font=Font(bold=True, color="C00000"),
            fill=PatternFill("solid", fgColor=WARNING_FILL),
        )
        row += 2

    def band(r, title):
        fill = PatternFill("solid", fgColor=SECTION_FILL)
        for c in range(span_first, span_last + 1):
            ws.cell(row=r, column=c).fill = fill
        ws.cell(row=r, column=span_first, value=title).font = Font(bold=True, size=13, color=NAVY)
        ws.row_dimensions[r].height = 22

    def tile(r, c, value, label, number_format="#,##0"):
        gap = Side(style="thick", color="FFFFFF")
        fill = PatternFill("solid", fgColor=TILE_FILL)
        for rr in (r, r + 1):
            ws.merge_cells(start_row=rr, start_column=c, end_row=rr, end_column=c + 1)
            for cc in (c, c + 1):
                cell = ws.cell(row=rr, column=cc)
                cell.fill = fill
                cell.border = Border(left=gap if cc == c else None, right=gap if cc == c + 1 else None)
        value_cell = ws.cell(row=r, column=c, value=value)
        value_cell.font = Font(bold=True, size=22, color=ACCENT)
        value_cell.alignment = Alignment(horizontal="center", vertical="bottom")
        value_cell.number_format = number_format
        label_cell = ws.cell(row=r + 1, column=c, value=label)
        label_cell.font = Font(size=10, color=MUTED)
        label_cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
        ws.row_dimensions[r].height = 38
        ws.row_dimensions[r + 1].height = 32

    patents = stats.get("patents", 0)

    band(row, "Patents at a glance")
    row += 1
    tile(row, 2, patents, "Filtered patents and applications")
    tile(row, 4, stats.get("grants", 0), "Granted patents")
    tile(row, 6, stats.get("applications", 0), "Published applications")
    tile(
        row, 8, stats.get("cited_patents", 0) / patents if patents else 0,
        "Cited at least once by a later patent", number_format="0%",
    )
    row += 3

    band(row, "Data-sheet crawl at a glance")
    row += 1
    tile(row, 2, stats.get("vendors", 0), "Supplier websites crawled")
    tile(row, 4, stats.get("docs_kept", 0), "Documents kept")
    tile(row, 6, stats.get("data_sheets", 0), "Technical data sheets")
    tile(row, 8, stats.get("vendors_with_patents", 0), "Suppliers that also hold filtered patents")
    row += 3

    band(row, "What's in this workbook")
    row += 1

    header_fill = PatternFill("solid", fgColor=NAVY)
    for c, label in ((2, "Sheet"), (3, "What you'll find (click a sheet name to go there)")):
        cell = ws.cell(row=row, column=c, value=label)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
    ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=span_last)
    for c in range(4, span_last + 1):
        ws.cell(row=row, column=c).fill = header_fill
    row += 1

    for name, description in SHEET_GUIDE:
        link = ws.cell(row=row, column=2, value=name)
        link.hyperlink = f"#'{name}'!A1"
        link.font = Font(color=LINK_COLOR, underline="single", bold=True)
        link.alignment = Alignment(vertical="center", wrap_text=True)
        _text_row(ws, row, description, 3, span_last, font=Font(size=10))
        row += 1

    row += 1
    band(row, "Words used in this report")
    row += 1

    for term, meaning in GLOSSARY:
        cell = ws.cell(row=row, column=2, value=term)
        cell.font = Font(bold=True, size=10)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        _text_row(ws, row, meaning, 3, span_last, font=Font(size=10), min_height=30 if len(term) > 20 else 15)
        row += 1

    row += 1
    band(row, "Files in this folder")
    row += 1

    for name, description in (
        ("FILTERED PATENTS", "The filtered patents, one CSV per file type (filtered_patents_aps / _sgml / _xml / _appxml). Each opens in Excel; the columns include patent type and claim category."),
        ("FORWARD CITATIONS", "forward_citations.xlsx (one row per filtered patent with its citation counts) and the citing / cited pairs split by file type, each small enough to open in Excel."),
    ):
        cell = ws.cell(row=row, column=2, value=name)
        cell.font = Font(bold=True, size=10)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        _text_row(ws, row, description, 3, span_last, font=Font(size=10), min_height=30)
        row += 1


def build_notes_sheet(workbook, notes):

    ws = workbook.create_sheet("How It Was Counted")
    ws.sheet_properties.tabColor = "A5A5A5"
    _setup_sheet(ws, {"A": 80, "B": 70})

    row = _sheet_header(
        ws, "How it was counted",
        "Where each number in this workbook comes from, and the rules used to count it.",
        2,
    )

    band = PatternFill("solid", fgColor=SECTION_FILL)

    for line in notes:

        if not line:
            row += 1
            continue

        if len(line) == 1 and str(line[0]).isupper():
            for c in (1, 2):
                ws.cell(row=row, column=c).fill = band
            ws.cell(row=row, column=1, value=line[0].title()).font = Font(bold=True, size=12, color=NAVY)
            ws.row_dimensions[row].height = 20
            row += 1
            continue

        label = line[0]
        value = line[1] if len(line) > 1 else ""

        if value == "" and len(str(label)) > 70:
            _text_row(ws, row, label, 1, 2, font=Font(size=10))
        else:
            a = ws.cell(row=row, column=1, value=label)
            a.alignment = Alignment(wrap_text=True, vertical="top")
            a.font = Font(size=10)
            b = ws.cell(row=row, column=2, value=value)
            b.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
            b.font = Font(size=10)
            if isinstance(value, int):
                b.number_format = "#,##0"

        row += 1


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
    forward_files += sorted(FORWARD_DIR.glob("forward_citation_pairs_*.csv"))

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
        ["SOURCES"],
        ["Generated", datetime.now().strftime("%Y-%m-%d %H:%M")],
        ["USPTO corpus", str(USPTO_ROOT)],
        ["Data-sheet corpus", str(TDS_ROOT)],
        [],
    ]

    stats = {}

    if not args.copy_only:

        workbook = openpyxl.Workbook()
        start_sheet = workbook.active
        start_sheet.title = START_SHEET

        print("\nPATENTS")
        data = load_filtered_patents()

        print("\nFORWARD CITATIONS")
        cites = load_forward_citation_counts()
        cites["patents"] = data[0]

        build_patents_sheet(workbook, notes, data, cites, stats)
        build_citations_sheet(workbook, notes, data, cites, stats)

        vendor_summary = build_crawl_sheets(workbook, notes, stats)

        build_vendor_patent_sheet(workbook, notes, data, cites, vendor_summary, stats)

    if not args.no_copy:
        copy_files(notes)

    if not args.copy_only:

        build_notes_sheet(workbook, notes)
        build_start_sheet(start_sheet, stats)
        workbook.active = 0

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
