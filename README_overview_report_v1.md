# Research Overview Report (`overview_report_v1.py`)

Builds one folder outside both corpora that summarizes the USPTO patent
pipeline and the vendor data-sheet crawler, and holds the result files that
open cleanly in Excel. It only reads from `D:\uspto_corpus` and
`D:\tds_corpus`; nothing there is changed.

```
D:\OVERVIEW\
  overview_report.xlsx
  FORWARD CITATIONS\    forward_citations.xlsx + forward_citation_pairs_<section>.csv
                        (and _part2, _part3 ... continuations)
  FILTERED PATENTS\     filtered_patents_aps.csv, _sgml, _xml, _appxml
```

The combined `forward_citation_pairs.csv` and `filtered_patents.csv` are left
out on purpose: the per-section / per-file-type files are the ones that each fit
within Excel's 1,048,576-row limit.

## Running it

```
python overview_report_v1.py              report + copy files
python overview_report_v1.py --no-copy    report only
python overview_report_v1.py --copy-only  copy files only
```

Re-run any time. The report is rebuilt; files are copied only when new or
changed, and a stale file (e.g. an old `_part3`) is removed from the overview
folder. If `overview_report.xlsx` is open in Excel, a time-stamped copy is
written instead. Paths are set in the CONFIG block at the top
(`USPTO_ROOT`, `TDS_ROOT`, `OVERVIEW_ROOT`).

Run it after the pipeline's `--classify` (for claim categories) and the
forward-citations `--export`.

## `overview_report.xlsx`

| Sheet | Contents |
|---|---|
| Patents | Bar chart of patents by decade, stacked by claim category; bar chart of patents by claim category; bar chart of filter results by file type (APS / SGML / XML grants / Applications); table of the top 10 patents by forward citations (title, assignee, year, type, claim category, times cited, applicant / examiner / other split) |
| Crawl | Headline numbers (vendors, pages crawled, PDF links, files downloaded, kept / translations / rejected, pages, size), PDF links by download outcome, and charts of documents by material category, by document type, top 15 vendors, test-method families, and language |
| Vendors | One row per vendor: seed status, pages crawled, PDF links found / downloaded / errored, documents kept by type (TDS / SDS / other), translations, rejected, total PDF pages, size, last download. Filterable. |
| Notes | Sources read, when, and how each number was counted |

How things are counted:

- A patent listed in both the 2001 APS and SGML files counts once. An
  application and its later grant are separate documents and both count.
- Decade comes from the grant date (publication date for applications).
- Forward citations are counted from the pairs CSV (one row per unique
  citing/cited pair), matching Times Cited in `forward_citations.xlsx`.
- Crawl "documents kept" are files not in `_rejected` and not a
  foreign-language copy; document type is the type read from the PDF text.
  The crawl database is opened read-only, so it's safe to run while a crawl is
  going.
