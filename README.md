# Web Films Research Corpus

Python tools that build an offline research library for **web handling, web converting, polymer films, foil, nonwovens and paper**:

- **Every relevant US patent and published application from 1976 to today**, selected from USPTO's own bulk data, with the PDFs, a measure of how influential each patent is, and what each one claims (a process, a machine, a product...).
- **The technical data sheets published by suppliers** of those materials and machines, collected from their websites and sorted by material, supplier, document type and language.
- **One Excel overview** that summarises both, written for someone seeing the data for the first time.

Everything runs on a Windows PC and writes to an external hard drive. Nothing depends on a paid service or a third-party search site.

> **New here?** If you only want to *use* the data, you don't need to run any code: open `D:\OVERVIEW\overview_report.xlsx` (start on the **Start Here** tab) and read `D:\OVERVIEW\Web_Films_Project_Summary.docx`. This README is for running or changing the code.

---

## The four programs

| # | Program | What it does | README |
|---|---|---|---|
| 1 | [`uspto_patent_pipeline_v15.py`](uspto_patent_pipeline_v15.py) | Downloads USPTO bulk patent data, builds a local database, selects the patents matching a CPC-code + keyword search matrix, labels what each one claims, and downloads their PDFs. | [README_pipeline_v15.md](README_pipeline_v15.md) |
| 2 | [`uspto_forward_citations_v5.py`](uspto_forward_citations_v5.py) | Finds every later patent that cites each selected patent, so you can rank them by influence. | [README_forward_citations_v1.md](README_forward_citations_v1.md) |
| 3 | [`tds_crawler_v1.py`](tds_crawler_v1.py) | Crawls supplier websites and collects their technical data sheets and related documents. | [README_tds_crawler_v1.md](README_tds_crawler_v1.md) |
| 4 | [`overview_report_v1.py`](overview_report_v1.py) | Reads the results of 1–3 and builds the overview workbook and folder. | [README_overview_report_v1.md](README_overview_report_v1.md) |

```
 Search-matrix workbooks ─┐
 (CPC codes + keywords)   ▼
 USPTO bulk data ──► 1. Patent pipeline ──► filtered patents + PDFs ──► 2. Forward citations ─┐
                                                                                             ├─► 4. Overview report
 Vendor workbook ─┐                                                                          │     (D:\OVERVIEW)
 vendor_seeds.csv ┴──► 3. Data-sheet crawler ──► supplier documents + catalog ───────────────┘
```

Programs 1 → 2 → 4 must run in that order. Program 3 is independent and can run any time before 4.

---

## Quick start

1. **Install Python 3.9 or newer** (developed in PyCharm on Windows) and the packages:
   ```
   python -m pip install -r requirements.txt
   ```
2. **Connect the hard drive as `D:`**, or change the folder settings at the top of each script (each README says which ones).
3. **Get a free USPTO API key** from the [USPTO Open Data Portal](https://data.uspto.gov/) and set it (only the patent download step needs it):
   ```
   $env:USPTO_API_KEY="your-key"        # PowerShell, current window only
   ```
4. **Try each program small before running it in full.** Every README has a "Quick start" with a small trial run. All four can be stopped (Ctrl-C) and restarted: finished work is skipped.

| To... | Run |
|---|---|
| Check the USPTO connection | `python uspto_patent_pipeline_v15.py --test` |
| Build the patent corpus (days) | `python uspto_patent_pipeline_v15.py --all` |
| Add influence data (hours) | `python uspto_forward_citations_v5.py --all` |
| Crawl suppliers (days, in batches) | `python tds_crawler_v1.py --all --contact you@example.com` |
| Refresh the overview (minutes) | `python overview_report_v1.py` |

---

## What's on the hard drive

| Folder | Made by | Contents |
|---|---|---|
| `D:\uspto_corpus\input\` | you | The 8 search-matrix workbooks (CPC codes and keywords) |
| `D:\uspto_corpus\RAW\`, `EXTRACTED\` | 1 | USPTO weekly archives, zipped and unzipped (re-downloadable) |
| `D:\uspto_corpus\DATABASE\` | 1, 2 | Every parsed patent (one SQLite file per USPTO format) and the citation index |
| `D:\uspto_corpus\RESULTS\FILTERED\` | 1 | **The selected patents** (`filtered_patents*.csv`) |
| `D:\uspto_corpus\RESULTS\FORWARD CITATIONS\` | 2 | **Citation counts** (`forward_citations.xlsx`) and citing/cited pairs |
| `D:\uspto_corpus\PDF\` | 1 | Patent PDFs (GRANTS, APPLICATIONS, OTHER) |
| `D:\uspto_corpus\MANIFEST\`, `LOGS\` | 1 | Progress records and run log |
| `D:\tds_corpus\` | 3 | Supplier documents (`pdfs\`), catalog (`tds_catalog.xlsx`), manifest, logs, audit records |
| `D:\OVERVIEW\` | 4 | `overview_report.xlsx`, Excel-sized copies of the results, sample PDFs |

Size: the patent database alone is about 310 GB, and a full rebuild needs well over 500 GB free. The supplier documents are about 14 GB.

---

## Things everyone should know

These are the design decisions and limits that shape every number in the outputs. Each program's README has its own detailed list.

- **US patents only.** Foreign patents and citations from foreign patents aren't included. (EPO and Google Patents coverage is planned; see Roadmap.)
- **A patent is selected only if it matches a CPC code *and* a keyword** from the search matrix. That's deliberate: it keeps precision high, at the cost of missing patents with unusual wording or classification. To change what's selected, edit the workbooks, not the code.
- **Patents before 2013 have no CPC codes**, so their older IPC codes are used instead. **Design patents have neither**, so they are almost never selected.
- **"Process vs. product" is not recorded by USPTO.** The claim category is worked out from the wording of claim 1; about 10% can't be classified.
- **Granted patents and published applications are separate documents.** One invention can appear twice (as its application and its grant).
- **The nonwoven search terms pull in medical products** (wound dressings, absorbent articles, CPC A61F). They dominate the most-cited lists; filter out A61F if you only care about industrial webs.
- **The crawler only collects public PDFs.** Data sheets behind logins, search forms or request forms are missed, and some supplier sites block crawlers. Some sites' terms forbid automated download; check before running at high volume.
- **The workbooks are the source of truth** for CPC codes, keywords and the supplier list. Every CPC code in a workbook should have its definition sheet in the same workbook.

---

## Requirements

- Windows (the scripts default to `D:\` paths; multiprocessing is written for Windows' `spawn` mode). Python 3.9 or newer.
- Packages in [`requirements.txt`](requirements.txt). Optional: `playwright` for the crawler's `--js` mode.
- A USPTO Open Data Portal API key for downloading patent data.
- An external drive (an SSD helps the database stage; anything works for the rest). **Never** point the output at a OneDrive/Dropbox-synced folder: it will try to upload hundreds of GB.

## Working practice

- **Stage and validate:** run a small sample first (`--sample`, `--format`, `--vendors`), look at the output, then run in full.
- **Versions are in the file names** (`_v15`, `_v5`). The first line of each run's log or console output shows the exact version that ran.
- **Check the logs after a run** (each README's "Checking a run" section lists what to look at).

## Data sources

- [USPTO Open Data Portal](https://data.uspto.gov/) bulk data: patent grants (APS, SGML, XML) and applications (XML), 1976 to present.
- USPTO Patent Public Search image server: patent PDFs.
- Supplier websites listed in the vendor workbook.

## Roadmap

- Non-US patents through the European Patent Office's OPS service (INPADOC, about 100 patent offices) and Google Patents Public Data (BigQuery).
- Better claim-category word lists, using the patents currently labelled Unclassified.
- Pull property values (thickness, tensile strength...) out of the technical data sheets into a comparison table.
