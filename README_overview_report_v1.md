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
| Patents | Bar chart of patents by decade, stacked by claim category; bar chart of patents by claim category; bar chart of filter results by file type (APS / SGML / XML grants / Applications); table of the top 10 patents by forward citations (title, assignee, year, type, claim category, times cited, applicant / examiner / other split); line chart of patents per year, grants vs applications; top 15 CPC subclasses with share of the corpus, forward citations and a short title |
| Citations | How often filtered patents are cited (0, 1–5, 6–20, 21–100, 101–500, over 500); top 10 by citations per year (patents with at least 10 citations, so new patents moving fast show up); average citations per patent by claim category; the documents citing your patents, by type (utility, design, application …) |
| Crawl | Headline numbers (vendors, pages crawled, PDF links, files downloaded, kept / translations / rejected, pages, size), PDF links by download outcome, charts of documents by material category, by document type, top 15 vendors, test-method families and language; a material category × document type table; every vendor with no kept documents and the likely reason (skipped in seeds, not crawled, site blocked, no PDF links, downloads failed, all rejected, only translations, off-topic) |
| Vendors | One row per vendor: seed status, pages crawled, PDF links found / downloaded / errored, documents kept by type (TDS / SDS / other), translations, rejected, total PDF pages, size, last download. Filterable. |
| Vendors & Patents | Crawled vendors matched to patent assignees by name: top 15 by patents (chart), then every match with documents kept, patents (grants / applications), patents cited, total citations, most-cited patent, and the assignee names that matched |
| Notes | Sources read, when, and how each number was counted |

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
