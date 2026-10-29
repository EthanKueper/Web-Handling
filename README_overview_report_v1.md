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

Written for someone opening it for the first time. Every analysis sheet has the
same layout: a title and one-line summary, then sections that each have a blue
title band, a grey line saying in plain words what the section shows, a yellow
**Takeaway** line computed from the data, and the table on the left with its
chart beside or below it. Hover over a column heading marked with a red corner
for what that column counts. Each sheet links back to Start Here.

| Sheet | Contents |
|---|---|
| Start Here | Key numbers as tiles, a guide to every sheet (click a name to go there), a glossary of the terms used, and what's in the two copied folders |
| Patents | Patents by decade; patents by decade split by claim category; patents by claim category; patents by file type (APS / SGML / XML grants / Applications); top 10 patents by forward citations; patents per year, grants vs applications; top 15 CPC subclasses |
| Citations | How often filtered patents are cited; top 10 fastest-cited patents (citations per year, at least 10 citations); average citations per patent by claim category; what kinds of documents cite them |
| Crawl | Headline numbers; what happened to each PDF link; documents by material category, type, top 15 vendors, test-method family and language; material category × document type table; vendors with no kept documents and the likely reason |
| Vendors | One filterable row per vendor with crawl and document counts (data bar on Documents kept) |
| Vendors & Patents | Crawled vendors matched to patent assignees by name: top 15 chart, then every match with data sheets and patents side by side and the assignee names matched |
| How It Was Counted | Sources read and the counting rules |

Claim categories keep the same colour in every chart.

How things are counted:

- A patent listed in both the 2001 APS and SGML files counts once. An
  application and its later grant are separate documents and both count.
- Decade comes from the grant date (publication date for applications).
- Forward citations are counted from the pairs CSV (one row per unique
  citing/cited pair), matching Times Cited in `forward_citations.xlsx`.
- Citations per year = times cited ÷ years from grant to the latest citing
  year in the data (at least 1 year).
- CPC subclasses: a patent counts once in each subclass it carries (IPC for
  records without CPC), so shares add to more than 100%.
- Vendor ↔ assignee matching drops corporate words (Inc, Corp, GmbH,
  Company …) and requires the assignee to begin with the vendor's remaining
  words: "3M Company" matches "3M Innovative Properties Company". It's
  approximate — a subsidiary under another name is missed, and a shared first
  name ("Dow" / "Dow Corning") can match — so the matched assignee names are
  listed for checking.
- Crawl "documents kept" are files not in `_rejected` and not a
  foreign-language copy; document type is the type read from the PDF text.
  The crawl database is opened read-only, so it's safe to run while a crawl is
  going.
