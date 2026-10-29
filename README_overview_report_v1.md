# Overview Report (`overview_report_v1.py`)

Reads the results of the other three programs and builds **one folder that summarises the whole project**: an Excel report a newcomer can follow, plus copies of the result files that open in Excel.

| | |
|---|---|
| **Input** | Patent filter results and forward citations (`D:\uspto_corpus\RESULTS\...`), crawl manifest and database (`D:\tds_corpus\...`) |
| **Output** | `D:\OVERVIEW\overview_report.xlsx`, `D:\OVERVIEW\FILTERED PATENTS\`, `D:\OVERVIEW\FORWARD CITATIONS\` |
| **Run time** | A few minutes (mostly reading the 520 MB citation file; the first run also copies about 1 GB) |
| **Run after** | Pipeline `--filter` and `--classify`, forward citations `--export`, crawler `--export` |

---

## Why it exists

The raw results are large (hundreds of thousands of rows) and spread across two drives' worth of folders. This report answers the obvious first questions in one place (how many patents, from when, what kind, which matter most, what the suppliers publish, which suppliers also hold patents), with each answer explained in plain words. It's meant for someone opening the data for the first time.

It **only reads** from `D:\uspto_corpus` and `D:\tds_corpus`; nothing there is changed. The crawl database is opened read-only, so it's safe to run during a crawl.

## Quick start

```
python overview_report_v1.py              # build the report and copy the result files
python overview_report_v1.py --no-copy    # report only
python overview_report_v1.py --copy-only  # copy files only
```

Then open `D:\OVERVIEW\overview_report.xlsx` and start on the **Start Here** tab.

## Before you run

- Settings at the top of the script: `USPTO_ROOT` (`D:\uspto_corpus`), `TDS_ROOT` (`D:\tds_corpus`), `OVERVIEW_ROOT` (`D:\OVERVIEW`). Change them if your drive letter or folders differ.
- The folders it reads must match the other programs: `RESULTS\FILTERED\` and `RESULTS\FORWARD CITATIONS\` under `USPTO_ROOT`, and `tds_manifest.csv` / `tds_crawler.db` under `TDS_ROOT`.
- Only `openpyxl` is needed beyond Python itself.
- A missing input doesn't stop the report: the affected sections show a note or zeros, and the console says which file wasn't found.

---

## What you get

```
D:\OVERVIEW\
  overview_report.xlsx
  FILTERED PATENTS\     filtered_patents_aps.csv, _sgml, _xml, _appxml
  FORWARD CITATIONS\    forward_citations.xlsx + forward_citation_pairs_<format>.csv (and _part2 files)
```

The combined `filtered_patents.csv` and `forward_citation_pairs.csv` are **left out on purpose**: they exceed Excel's 1,048,576-row limit, while each per-format file fits. Re-running copies only new or changed files and removes stale ones (e.g. an old `_part3`).

### The workbook

Every analysis tab uses the same layout, so a reader always knows where to look:

- Title, one-line summary and a link back to Start Here.
- For each section: a **blue** title bar, a **grey** line explaining what it shows, a **yellow Takeaway** line computed from the data, then the table with its chart.
- Hover over a column heading with a small red corner to see what it counts. Each claim category has the same colour in every chart.

| Tab | Contents |
|---|---|
| Start Here | Key numbers, a guide to every tab (click to go there), a glossary, and what's in the two copied folders |
| Patents | Patents by decade; by decade split by claim category; by claim category; by file type; top 10 by forward citations; patents per year (grants vs applications); top 15 CPC subclasses |
| Citations | How often patents are cited; top 10 fastest-cited (citations per year, at least 10 citations); average citations by claim category; what kinds of documents cite them |
| Crawl | Headline numbers; what happened to each PDF link; documents by material category, type, top 15 suppliers, test-method family and language; category × type table; suppliers with no kept documents and the likely reason |
| Vendors | One filterable row per supplier with crawl and document counts |
| Vendors & Patents | Suppliers matched to patent owners: top 15 chart, then every match with data sheets and patents side by side and the owner names matched |
| How It Was Counted | Sources, timestamps and counting rules |

If the report is open in Excel when you run the script, a time-stamped copy is written instead.

---

## Checking a run

- The console prints the rows read from each input file and what was copied.
- The **How It Was Counted** tab records exactly which files were read and when.
- If claim categories show as "Not classified (run --classify)", run the pipeline's `--classify`, then this script again.

## Need to know: how things are counted

- **A patent in both the 2001 APS and SGML files counts once** (4,317 such patents). The "by file type" chart counts rows per file, so it includes them in both.
- **An application and its later grant are separate documents** and both count.
- **Decade and year** come from the grant date (publication date for applications).
- **Forward citations** are counted from the pairs CSV (one row per unique citing/cited pair), so they match Times Cited in `forward_citations.xlsx`.
- **Citations per year** = times cited ÷ years from grant to the latest citing year in the data (at least 1 year).
- **CPC subclasses:** a patent counts once in each subclass it carries (IPC for records without CPC), so shares add to more than 100%. Subclass names come from a built-in list; an unlisted subclass shows a blank name.
- **Suppliers are matched to patent owners by name:** corporate words (Inc, Corp, GmbH, Company...) are dropped and the owner must begin with the supplier's remaining words, so "3M Company" matches "3M Innovative Properties Company". It's approximate: a subsidiary under another name is missed, and a shared first word can over-match ("Dow" also matches "Dow Corning"). The matched owner names are listed so you can check.
- **Crawl "documents kept"** = files not in `_rejected` and not a foreign-language copy; document type is the type read from the PDF text.
- **Suppliers with no kept documents:** the "likely reason" is inferred from crawl counts (skipped in seeds, not crawled, site blocked, no PDF links, nothing downloaded, all rejected, only translations, or off-topic deletions). Check `--status` in the crawler for detail.

## Settings (top of the script)

| Setting | Default | Controls |
|---|---|---|
| `USPTO_ROOT`, `TDS_ROOT`, `OVERVIEW_ROOT` | `D:\uspto_corpus`, `D:\tds_corpus`, `D:\OVERVIEW` | Where to read and write |
| `TOP_N_PATENTS`, `TOP_N_VENDORS`, `TOP_N_CPC` | 10, 15, 15 | Length of the top lists |
| `CITATION_RATE_MIN_CITES` | 10 | Minimum citations for the fastest-cited table |
| `CPC_SUBCLASS_TITLES` | built-in list | Short names shown for CPC subclasses; add any that show blank |
