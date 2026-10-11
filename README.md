# Web Films — Patent & Technical Data Corpus Tools

Python tools for building a research library on **web handling, web converting, polymer films, foil, nonwoven fabrics and paper processes**. The repository contains three scripts:

| Script | What it builds | Detailed docs |
|---|---|---|
| [`uspto_patent_pipeline_v15.py`](uspto_patent_pipeline_v15.py) | A CPC- and keyword-filtered library of USPTO patent PDFs (grants and applications, 1976 to present, 500,000+ patents scanned) | [`README_uspto_patent_pipeline_v15.md`](README_uspto_patent_pipeline_v15.md) |
| [`uspto_forward_citations_v1.py`](uspto_forward_citations_v1.py) | Forward-citation counts and normalised importance metrics for the filtered patents, computed fully offline | [`README_uspto_forward_citations_v1.md`](README_uspto_forward_citations_v1.md) |
| [`tds_crawler_v1.py`](tds_crawler_v1.py) | A classified library of vendor technical data sheets (TDS) for the same material categories | [`README_tds_crawler_v1.md`](README_tds_crawler_v1.md) |

This page is an overview. Each script's own README has the full command reference, options and output layout.

---

## How the pieces fit together

```
                 Domain workbooks (CPC codes + keywords)
                                  │
USPTO bulk data ──► uspto_patent_pipeline_v15.py ──► SQLite DBs ──► candidate CSV ──► patent PDFs
                                  │
                                  ▼
                   uspto_forward_citations_v1.py ──► citation counts, Citations Per Year,
                                                     Cohort Percentile

Vendor workbook + vendor_seeds.csv ──► tds_crawler_v1.py ──► technical data sheet PDFs + manifest
```

The two patent scripts are one workflow: the pipeline builds the parsed patent database and the filtered corpus, and the citations script ranks those patents by how often later patents cite them. The TDS crawler is independent; it covers the same material categories from the supplier side.

---

## 1. `uspto_patent_pipeline_v15.py` — patent corpus pipeline

Downloads USPTO's official bulk data products, parses them into SQLite, filters by CPC classification **and** keywords, and retrieves the PDFs of the matching patents.

**Data products**

| Product | Format | Coverage |
|---|---|---|
| `PTGRAPS` | APS / Green Book | ~1976–1996 grants |
| `PTGRSGM` | SGML | ~1996–2001 grants |
| `PTGRXML` | XML | ~2001–present grants |
| `APPXML` | XML | ~2001–present applications |

**Stages:** Download → Extract → Parse → SQLite database → CPC/IPC filter → Keyword filter → Candidate CSV → PDF retrieval

**Key behaviour**

- **Precision over breadth.** A patent must match both a CPC code *and* a keyword (AND logic).
- **Excel-driven search terms.** CPC codes and keywords load from a folder of domain workbooks (web processes, web converting, web handling, nonwovens, foil, paper). Codes are normalised on load, including regular and non-breaking spaces, which otherwise cause silent filter misses.
- **IPC fallback.** Pre-2013 patents often lack CPC codes; `MATCH_IPC_AS_CPC_FALLBACK = True` lets their IPC codes stand in.
- **Per-product databases.** Each bulk product gets its own SQLite file (WAL mode, deferred index creation).
- **Split downloads.** `--download-aps`, `--download-sgml`, `--download-xml` and `--download-appxml` let the multi-day bulk download be spread over several sessions.
- **Parallel and resumable.** Parsing uses `ProcessPoolExecutor` with resumability through `database_manifest.csv`; PDF downloads use `ThreadPoolExecutor` with a download manifest.
- **Split storage roots.** Bulk archives (re-downloadable, suited to HDD), databases, and the curated PDF library (suited to SSD) can live on different drives. Default root: `D:\uspto_corpus`.

See [`README_uspto_patent_pipeline_v15.md`](README_uspto_patent_pipeline_v15.md) for the commands and configuration.

---

## 2. `uspto_forward_citations_v1.py` — forward citations

Counts how many times each patent is cited by later patents, using the locally parsed bulk data only. It replaces the PatentsView API, which shut down in March 2026.

- Three format-specific citation extractors: XML, SGML and APS.
- Records the citation category (examiner, applicant, other) and the citing patent's date.
- Two normalised metrics to correct for the recency bias of raw counts:
  - **Citations Per Year** — citations divided by years since grant.
  - **Cohort Percentile** — the patent's rank against patents of the same cohort.

Forward citations are a stronger signal of a patent's importance than backward citations, so this script is the main tool for prioritising reading within the filtered corpus.

See [`README_uspto_forward_citations_v1.md`](README_uspto_forward_citations_v1.md) for usage.

---

## 3. `tds_crawler_v1.py` — vendor technical data sheet crawler

Crawls every vendor listed in `Vendors_With_Technical_Sheets_By_Material_Category.xlsx`, finds the PDF links on each site, downloads the data sheets, and tags each one by document type and by the test methods it cites (ASTM, ISO, TAPPI, NWSP, IPC-TM-650, DIN, JIS).

- **Seed-driven.** Domains, start URLs, path limits, JavaScript rendering and brand filters are set per vendor in `vendor_seeds.csv`.
- **Organised output.** PDFs are filed by material category and vendor (`pdfs\<Category>\<Vendor>\`), de-duplicated by SHA-256, and listed in `tds_manifest.csv`.
- **English first.** Translated copies are detected and the English version is preferred; foreign-only sheets are kept and flagged.
- **Junk filtering.** Investor-relations, sustainability, press and report PDFs are filtered at crawl, link and content level; rejected files are moved, not deleted.
- **Polite and resumable.** Obeys robots.txt, rate-limits per vendor, backs off on 429/503, and resumes after Ctrl-C.

Default root: `D:\tds_corpus`, with the workbook and seed file in `D:\tds_corpus\input\`.

Quick start:

```
pip install -r requirements.txt
python tds_crawler_v1.py --list-vendors
python tds_crawler_v1.py --discover --vendors "rogers,camden,novelis" --max-pages 60 --contact you@example.com
python tds_crawler_v1.py --status
python tds_crawler_v1.py --download --classify --export --vendors "rogers,camden,novelis"
python tds_crawler_v1.py --all --contact you@example.com
```

See [`README_tds_crawler_v1.md`](README_tds_crawler_v1.md) for all options, the seed file columns and the output layout.

---

## Requirements

- Python 3.9+
- Windows is the tested platform (paths default to `D:\`); scripts use `if __name__ == "__main__":` guards for Windows `spawn`-mode multiprocessing.
- Enough disk for the bulk USPTO archives and databases; an external drive is recommended.
- Dependencies are listed in `requirements.txt`. Playwright is optional, for JavaScript-heavy vendor sites in the TDS crawler.

## Working practice

- **Stage and validate.** Run small samples (a single product, a few vendors, `--discover` dry runs) before full runs.
- **Workbooks are authoritative.** CPC codes, keywords and the vendor list come from Excel workbooks, not from the scripts.
- **Versioned filenames.** Each script carries its version in the filename (`_v15`, `_v1`), and its README is named to match.

## Data sources

- USPTO Bulk Data (`bulkdata.uspto.gov`) — patent full text and metadata
- USPTO PPUBS (`image-ppubs.uspto.gov`) — patent PDFs
- Vendor websites listed in the vendor workbook — technical data sheets

Respect each source's terms of use. Some vendor portals forbid automated download; check before running the crawler at high volume.

## Roadmap

- International patent coverage through EPO OPS (INPADOC) and Google Patents Public Data on BigQuery
- Directory-count mode to get exact file counts per USPTO product before large downloads
