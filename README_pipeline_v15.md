# USPTO Patent Pipeline (`uspto_patent_pipeline_v15.py`, v15.10)

Builds a local copy of USPTO's patent data from 1976 to today, then picks out the patents relevant to web films and related materials, labels what each one claims, and downloads their PDFs.

| | |
|---|---|
| **Input** | USPTO bulk data (downloaded by the script) + your search-matrix workbooks in `D:\uspto_corpus\input\` |
| **Main output** | `D:\uspto_corpus\RESULTS\FILTERED\filtered_patents.csv` (the selected patents) and `D:\uspto_corpus\PDF\` (their PDFs) |
| **Run time** | Days for a full build (mostly downloading and parsing); minutes to re-filter an existing database |
| **Disk** | Well over 500 GB for a full build (the database alone is about 310 GB) |
| **Next step** | [Forward citations](README_forward_citations_v1.md), then the [overview report](README_overview_report_v1.md) |

---

## Why it works this way

- **USPTO's own bulk data, not a search website.** Every result can be traced to a source file, nothing depends on a third-party service, and the whole history can be searched offline as often as needed. (The PatentsView API this project once used shut down in March 2026.)
- **Download everything once, filter as often as you like.** The database holds every parsed patent, not just the matches, so changing the search only means re-running `--filter`, not re-downloading.
- **Precision over breadth.** A patent is selected only if it has a CPC code from the search matrix **and** contains a keyword. A patent that mentions "film" but is classified somewhere unrelated (real example: a marine-hardware patent, IPC B63B) is correctly left out.
- **Every stage can be stopped and restarted.** Progress is recorded in manifest files, so an interrupted run carries on where it stopped.

## Quick start

```
python -m pip install -r requirements.txt
$env:USPTO_API_KEY="your-key"                         # free key from data.uspto.gov

python uspto_patent_pipeline_v15.py --test            # 1. check the key and connection
python uspto_patent_pipeline_v15.py --sizecheck       # 2. see how much data a full download is
python uspto_patent_pipeline_v15.py --sample          # 3. download ~12 files spread across all eras
python uspto_patent_pipeline_v15.py --extract         # 4. then run the rest on that sample
python uspto_patent_pipeline_v15.py --database
python uspto_patent_pipeline_v15.py --filter
```

Open `RESULTS\FILTERED\filtered_patents.csv` and check the results look sensible. Then run the full build:

```
python uspto_patent_pipeline_v15.py --all             # download -> extract -> database -> filter -> pdf
```

> **Don't combine `--sample` with `--all`.** `--all` always runs the full download. For a small end-to-end test, run the stages one at a time as above.

---

## Before you run

### 1. Settings at the top of the script

| Setting | What to set |
|---|---|
| `USPTO_API_KEY` | Leave as is and set the `USPTO_API_KEY` environment variable instead (keeps the key out of the code). Only the download stages need it. |
| `BULK_DATA_ROOT_DIR`, `DATABASE_ROOT_DIR`, `PDF_ROOT_DIR` | Where everything is written. Default `D:\uspto_corpus` for all three; they can be on different drives. **Use an external drive, never `C:\Users\...` or a OneDrive/Dropbox folder** (it would try to upload hundreds of GB). |
| `CPC_EXCEL_INPUT_FOLDER` | Folder holding the search-matrix workbooks. Default `D:\uspto_corpus\input`. Every `.xlsx` directly inside it is read and merged. |

### 2. The search matrix (what gets selected)

The workbooks in the input folder decide which patents are selected. There's one per topic (currently Foil, Nonwoven, Nonwoven Processes, Paper, Polymer Film, Web Converting, Web Handling, Web Processes: 1,332 CPC codes and 317 keywords in total).

- **CPC codes:** every sheet whose tab name looks like a CPC class (letter, two digits, letter, e.g. `B65H`, `C08J`, or `B29C48`) is a CPC sheet. Codes are read from column **E**, from row **2** (`CPC_COLUMN`, `CPC_START_ROW`). Other sheets ("Notes", "Summary") are ignored.
- **Keywords:** the sheet named **Keywords**, column **A**, from row **2** (`KEYWORDS_SHEET_NAME`, `KEYWORDS_COLUMN`, `KEYWORDS_START_ROW`).
- Codes are cleaned on load (including invisible non-breaking spaces from Excel, which once caused thousands of silent misses).
- Each workbook should be self-contained: every CPC code in its search matrix should have a definition sheet in the same file.

The first lines of each run's log list every workbook, sheet and count loaded. **Check them:** if a workbook or sheet is missing there, its codes weren't used.

### 3. Assumptions

- Windows, Python 3.9+ (tested on 3.9 and 3.12), run from PyCharm's terminal or any terminal.
- The drive is connected before you start; paths are hard-coded to it.
- An SSD helps the `--database` stage; drive type barely matters elsewhere.

---

## What it does

| Stage | Command | What happens |
|---|---|---|
| Download | `--download` (or `--download-aps`, `--download-sgml`, `--download-xml`, `--download-appxml`) | Downloads every weekly archive for the four USPTO products (below) into `RAW\`. |
| Extract | `--extract` | Unzips into `EXTRACTED\`. |
| Database | `--database [--format aps\|sgml\|xml\|appxml]` | Reads every patent record into SQLite: numbers, dates, title, CPC/IPC, inventors, applicants, assignees, abstract, claims, description. One database file per product. |
| Filter | `--filter [--format ...]` | Selects patents with a matching CPC code (or IPC code before 2013) **and** a keyword in the title, abstract, claims or description. Also labels patent type and claim category. |
| Classify | `--classify [--format ...]` | Adds the patent type / claim category labels to results filtered before v15.10, without re-filtering. |
| PDFs | `--pdf [--format ...] [--limit N]` | Downloads the PDF of every selected patent from USPTO's Patent Public Search server. |

**The four USPTO products** (USPTO publishes each era in a different format, so each has its own parser):

| Product | Format | Covers |
|---|---|---|
| `PTGRAPS` | APS "Green Book" tagged text | Granted patents 1976–2001 |
| `PTGRSGM` | SGML | Granted patents 2001 |
| `PTGRXML` | XML | Granted patents 2002–present |
| `APPXML` | XML | Published applications 2001–present |

For the current corpus: 4,113 weekly files, about 18.3 million records parsed, 474,183 unique patents and applications selected.

---

## What you get

| File | What it is |
|---|---|
| `RESULTS\FILTERED\filtered_patents.csv` | **The selected patents**, all products together. Read by `--pdf` and the forward-citations script. Too large for Excel; use Power Query or code. |
| `RESULTS\FILTERED\filtered_patents_<product>.csv` | The same, one file per product (`aps`, `sgml`, `xml`, `appxml`). Each opens in Excel. |
| `PDF\GRANTS\`, `PDF\APPLICATIONS\`, `PDF\OTHER\` | The PDFs, named by patent number. `OTHER` holds statutory invention registrations and anything unexpected. |
| `DATABASE\USPTO_APS.sqlite`, `USPTO_SGML.sqlite`, `USPTO_XML_GRANTS.sqlite`, `USPTO_APPLICATIONS.sqlite` | Every parsed patent (not just matches). Open with DB Browser for SQLite. See [Database layout](#database-layout). |

### Reading `filtered_patents*.csv`

One row per patent.

| Column | Meaning |
|---|---|
| `patent_number` | Granted patent number, e.g. `US10123456B2`. Blank for applications. **APS-era numbers (1976–2001) end in an extra check character**: `US039302711` is patent 3,930,271. |
| `publication_number` | Number as published; applications look like `US20140123360A1` (year + 7 digits + kind code). |
| `document_type` | `grant`, `application`, or `sir` (statutory invention registration). |
| `patent_type` | Utility, Design, Plant, Reissue, Statutory Invention Registration (SIR), Defensive Publication, Application (utility), Application (plant). Read from the number and kind code, so it's exact. |
| `claim_category` | What claim 1 protects: Process / Method, Apparatus / System, Product / Article, Composition, Product-by-process, Design, Plant. **Unclassified** = the wording wasn't recognised; **No claims text** = the record has no claims. See [How claim category is decided](#how-claim-category-is-decided). |
| `independent_claim_categories` | Every category across the independent claims, e.g. `Process / Method; Apparatus / System` for a patent claiming both the method and the machine. |
| `title`, `inventors`, `applicants`, `assignees` | As published; several names are separated by `;`. The assignee is the owner when published, not necessarily today. |
| `filing_date`, `publication_date`, `grant_date`, `priority_date` | `YYYY-MM-DD`. Applications have no grant date. |
| `cpc`, `ipc` | Classification codes separated by `;`. |
| `source_path` | The extracted USPTO file the record came from. |

---

## Checking a run

| File | What to look at |
|---|---|
| `LOGS\pipeline.log` | Time-stamped log of every run (`date time \| LEVEL \| message`). Check the opening lines (workbooks, CPC codes and keywords loaded), then search for `ERROR` and `WARNING`. |
| `MANIFEST\download_manifest.csv` | One row per archive: `file_name`, `file_size`, `local_path`, `status`, `error`. Every row should be `COMPLETE`. |
| `MANIFEST\database_manifest.csv` | One row per file parsed: `path`, `records_inserted`, `status`, `error`. Every row should be `success`; `error` rows are retried on the next `--database`. |
| `MANIFEST\pdf_manifest.csv` | One row per selected patent: `publication_number`, `local_path`, `status` (`COMPLETE`/`FAILED`), `error`. Use it to find a patent's PDF. |
| `LOGS\*_product.json` | USPTO's own listing of each product's files, saved for reference. |
| Console | Counts per product after `--filter` (and per category after `--classify`). |

Current state: all downloads complete, 4,101 files parsed with no errors, 477,688 PDFs downloaded and 812 failed (see Limits).

---

## Need to know: assumptions and limits

- **US only, 1976 to present.** Pre-1976 patents aren't in USPTO's full-text bulk data.
- **CPC + keyword, both required.** High precision; patents with unusual wording or classification are missed. Broaden the workbooks to widen the net.
- **CPC didn't exist before 2013.** Older patents are matched on their IPC codes instead (`MATCH_IPC_AS_CPC_FALLBACK = True`), which are less specific.
- **Design patents are almost never selected.** They carry no CPC or IPC codes at all (they use Locarno classification), so nothing matches. Only 5 are in the current results.
- **The claim category is inferred**, not recorded by USPTO. About 10% of patents are Unclassified, and a few may be mislabelled.
- **One invention can appear twice**, as its published application and as its granted patent. Patents that appear in both the 2001 APS and SGML files are kept in both product files (4,317 of them); the overview report counts them once.
- **The nonwoven CPC groups (A61F13, A61L15) bring in medical products** such as wound dressings. Filter on `cpc` if you don't want them.
- **Some PDFs can't be downloaded automatically (812 currently):**
  - Statutory invention registrations (`H` numbers): no working PDF link format exists.
  - A few hundred utility patents return "not found" (HTTP 404) from USPTO's server.
  - Documents over about 1,000 pages may not be served (per USPTO's documentation).

  They are left as `FAILED` in `pdf_manifest.csv` for manual download (Google Patents or Espacenet). Reissue (`RE`) and defensive publication (`T`) PDFs work since v15.9.
- **Changing the workbooks only affects new filter runs.** Re-run `--filter` for every product (and then forward citations with `--reset`) so all results use the same criteria.

---

## Common tasks

**Spread the download over several days.** Each `--download-<product>` flag downloads one product to completion. `--extract` and `--database` skip finished work, so it's safe to run them every day:

```
Day 1: --download-aps    --extract --database --format aps
Day 2: --download-sgml   --extract --database --format sgml
Day 3: --download-xml    --extract --database --format xml
Day 4: --download-appxml --extract --database --format appxml
Day 5: --filter
```

Run `--filter` only once all four products are in the database; earlier, the results would be incomplete.

**Change the search.** Edit the workbooks, then:

```
python uspto_patent_pipeline_v15.py --filter          # all products (or --format one at a time)
python uspto_patent_pipeline_v15.py --pdf             # PDFs for any newly selected patents
```

**Download PDFs in batches.** `--pdf --format xml --limit 500` stops after 500 downloads; run again to continue. Failed PDFs are skipped by normal runs. Retry them at the end with `--pdf --retry-errors-only`. `--pdf` is optional: nothing else reads the PDFs.

**Add claim categories to older results.** `--classify` (or `--classify --format xml`) labels results filtered before v15.10 in minutes, without re-filtering. Run it before the forward-citations `--export` so the labels carry through.

**Machine getting hot or unresponsive during `--database`.** Use `--workers 3` or `--workers 2`. A sudden power-off (not a freeze) usually means overheating or power supply: check Windows Event Viewer for "Kernel-Power 41" and a temperature monitor such as HWiNFO.

---

## Command reference

| Command | Does |
|---|---|
| `--test` | Check the API key and connection |
| `--sizecheck` | Report each product's total size and file count without downloading |
| `--sample` | Download a few files per product, spread across its dates |
| `--download`, `--download-aps`, `--download-sgml`, `--download-xml`, `--download-appxml` | Download all products, or one |
| `--extract` | Unzip downloaded archives |
| `--database` | Parse into SQLite. Options: `--format`, `--limit N`, `--workers N`, `--skip-errors`, `--retry-errors-only`, `--backup` |
| `--filter` | Select patents. Option: `--format` |
| `--classify` | Add patent type / claim category to existing results. Option: `--format` |
| `--pdf` | Download PDFs. Options: `--format`, `--limit N`, `--retry-errors-only` |
| `--all` | Download → extract → database → filter → PDFs |
| `--migrate-db` | One-time move of a v14 database into the v15 layout |
| `--requeue-merged` | One-time re-parse of files affected by the merged-record bug fixed in v15.1 |

### Main settings (top of the script)

| Setting | Default | Controls |
|---|---|---|
| `DATABASE_PARSE_WORKERS` | 4 | Parallel parsing processes (`--workers` overrides) |
| `PDF_DOWNLOAD_WORKERS` | 8 | Parallel PDF downloads |
| `PPUBS_MIN_REQUEST_INTERVAL_SECONDS` | 1.0 | Minimum gap between PDF requests across all workers (politeness to USPTO) |
| `SAMPLE_FILES_PER_PRODUCT` | 3 | Files per product for `--sample` |
| `MATCH_IPC_AS_CPC_FALLBACK` | True | Use IPC codes for patents without CPC |
| `COMPRESS_TEXT` | True | Store claims/description compressed (3–4× smaller) |

---

## Technical details

### How claim category is decided

USPTO doesn't record whether a patent protects a product or a process, so the script reads claim 1. The main noun of the claim's opening phrase decides it: "A film winding **apparatus**..." → Apparatus / System; "A multilayer **film**..." → Product / Article; "A **method** of..." → Process / Method; "A polyethylene **composition**..." → Composition; "...obtained by the process of..." → Product-by-process. If the noun isn't in the word lists it's left **Unclassified** rather than guessed. The word lists are in the "PATENT TYPE + CLAIM CATEGORY" section of the script; add nouns there to reduce Unclassified results.

A claim that refers to another claim but opens a new category ("An apparatus for carrying out the method of claim 1") counts as independent. Plain dependent claims and bracketed (cancelled) reissue claims don't.

### Database layout

One SQLite file per product, each with two tables:

| Table | Contents |
|---|---|
| `patents` | Numbers, document type, dates, `cpc`, `ipc`, title, inventors, applicants, assignees, source file. Indexed on `patent_number`, `publication_number`, `publication_date`. |
| `patent_text` | `abstract` (text), `claims` and `description` (zlib-compressed). Same `id` as `patents`. To read: `zlib.decompress(blob).decode("utf-8")`. |

Splitting the text out keeps the CPC check fast (it only reads the narrow table), and per-product files mean a crash or drive fault can only affect one product.

### How `--filter` works

It first checks CPC/IPC on the narrow table, then reads text only for CPC matches. Keywords are combined into one case-insensitive whole-word pattern, and fields are checked title → abstract → claims → description, stopping at the first hit.

### `--database` performance

- Each worker streams a file and writes every 250 records, so memory stays low regardless of file size (v14 held whole files in memory and ran out of RAM).
- Workers run at below-normal priority on Windows so the PC stays usable, and on Python 3.11+ are recycled periodically.
- A file that fails part-way is marked `error` and retried next run; already-stored rows are skipped, so nothing is duplicated.
- If a worker crashes (often with `WinError 1450`), a fresh pool continues. A disk write error stops the run cleanly rather than writing to a failing drive.
- Indexes are built after loading. `--backup` copies each finished product's database to `DATABASE\BACKUP\` (opt-in).
- Windows re-imports the script in every worker process; the workbook loading is behind `if __name__ == "__main__":` so workers skip it.

### Upgrading from v14

If the old single `DATABASE\USPTO_PATENTS.sqlite` exists, `--database` and `--filter` stop until you run `--migrate-db`, which copies every row into the per-product files without re-parsing (resumable; never changes the old file). When the counts it prints look right, delete the old file yourself.

### Merged-record fix (v15.1)

Some weekly XML files put a patent's `<?xml ...?>` declaration on the same line as the previous document's closing tag; before v15.1 that patent was lost. To recover patents from files parsed earlier: `--requeue-merged`, then `--database --format xml --retry-errors-only` and the same for `appxml`.

### Forward citations

An older `--citations` stage queried PatentsView, which shut down in March 2026. Citations are now handled offline by [`uspto_forward_citations_v5.py`](README_forward_citations_v1.md), which needs this pipeline run through at least `--filter` (PDFs not required).

---

## Version history

| Version | Change |
|---|---|
| v15.10 | Patent type and claim category columns; `--classify`; results read/written in `RESULTS\FILTERED` |
| v15.9 | PDF links for reissue (RE) and defensive publication (T) patents |
| v15.8 | Connection pool sized to worker count; API key no longer sent with PDF requests |
| v15.7 | `--pdf` shows progress per section; workbook summaries go to the log only |
| v15.6 | APS numbers ending in `&` get correct PDF links |
| v15.5 | Failed PDFs skipped until `--retry-errors-only` |
| v15.4 | APS check digit dropped before requesting PDFs |
| v15.3 | `--pdf` accepts `--format`, `--limit`, `--retry-errors-only` |
| v15.2 | `--filter --format` filters one product at a time |
| v15.1 | Merged-record fix; `--requeue-merged` |
| v15 | One database per product, compressed text, streaming parse, `--migrate-db` |
