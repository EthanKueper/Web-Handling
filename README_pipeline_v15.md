# USPTO Patent Collection Pipeline (`uspto_patent_pipeline_v15.py`)

Downloads, parses, filters, and retrieves PDFs for USPTO patents matching a
CPC + keyword search matrix, across the full historical corpus (1976–present).

## What it does, in order

1. **Download** — pulls USPTO bulk-data archives for four separate products,
   since USPTO ships each era in a different format:
   - `PTGRAPS` — Patent Grants, APS/Green Book (flat tagged text, ~1976–1996)
   - `PTGRSGM` — Patent Grants, SGML (~1996–2001)
   - `PTGRXML` — Patent Grants, XML (~2001–present)
   - `APPXML` — Patent Applications, XML (2001–present)
2. **Extract** — unpacks the downloaded archives.
3. **Database** — parses every record (auto-detecting format) into a local
   SQLite database with a consistent schema regardless of source era.
4. **Filter** — a patent must match **both** a CPC/IPC code from your search
   matrix **and** a keyword to be selected. Writes `filtered_patents.csv`.
5. **PDF** — downloads the actual PDF for every filtered patent only (not
   the whole corpus), via USPTO's Patent Public Search image server.

Every stage is restartable — downloads, database inserts, and PDF fetches
all skip work that's already done rather than redoing it.

## Before running: fill in the config

Three things at the top of the file need your values:

| Constant | What it is |
|---|---|
| `USPTO_API_KEY` | USPTO Open Data Portal key (or set the `USPTO_API_KEY` env var instead and leave this as-is) |
| `BULK_DATA_ROOT_DIR` / `DATABASE_ROOT_DIR` / `PDF_ROOT_DIR` | **Must point at an external drive**, never `C:\Users\...` or a OneDrive/Dropbox-synced folder — this pipeline writes a genuinely large volume of data and a cloud-synced folder will try to upload all of it in the background. Default: `D:\uspto_corpus` for all three; they can point at different physical drives if you have more than one. |
| `CPC_EXCEL_INPUT_FOLDER` | Folder containing your CPC/keyword search-matrix workbook(s). Every `.xlsx` directly inside it is picked up automatically and merged — no code change needed to add a new workbook. Default: `D:\uspto_corpus\input` |

### How the search matrix is read

- Every worksheet whose **tab name** looks like a CPC class/subclass (a
  letter, two digits, a letter — e.g. `B29C`, `C08J`, optionally with
  trailing digits like `B29C48`) is treated as a CPC sheet. Sheets named
  things like "Notes" or "Summary" are ignored.
- CPC codes are read from column `CPC_COLUMN` (default `E`), starting at row
  `CPC_START_ROW` (default `2`).
- Keywords are read from a sheet named `KEYWORDS_SHEET_NAME` (default
  `"Keywords"`), column `KEYWORDS_COLUMN` (default `A`), starting at row
  `KEYWORDS_START_ROW` (default `2`).
- Every CPC code in the matrix should resolve to a definition sheet within
  the **same workbook** — each domain file is meant to be self-contained.

## Running it

```
python uspto_patent_pipeline_v15.py --test        # confirm API key/connection work
python uspto_patent_pipeline_v15.py --sizecheck    # check corpus size before downloading anything
python uspto_patent_pipeline_v15.py --sample       # small date-spread sample — sanity-check parsers first
python uspto_patent_pipeline_v15.py --all          # download -> extract -> database -> filter -> pdf
```

Individual stages can also be run on their own: `--download`, `--extract`,
`--database`, `--filter`, `--pdf`. Recommended order the first time:
`--test`, then `--sizecheck`, then `--sample` to confirm the parsers look
right against real files, **then** `--all`.

**Watch out:** `--sample` and `--all` don't combine the way you'd expect.
`--all` calls the full-corpus download regardless of whether `--sample` is
also passed — it checks `args.all` first and returns before ever looking at
`args.sample`. To do a small end-to-end run, run each stage on its own
instead of using `--all`:

```
python uspto_patent_pipeline_v15.py --sample     # ~12 files across all 4 products/eras
python uspto_patent_pipeline_v15.py --extract
python uspto_patent_pipeline_v15.py --database
python uspto_patent_pipeline_v15.py --filter
python uspto_patent_pipeline_v15.py --pdf        # optional — see below
```

**`--pdf` is optional and independent.** Nothing downstream (including the
forward-citations script) reads the PDFs back in — they're for your own
records. You can stop after `--filter` and come back to `--pdf` whenever
you want; it always covers every row currently in `filtered_patents.csv`,
regardless of when you run it.

## Splitting the download across multiple days

`--download` pulls all four products (APS, SGML, XML grants, APPXML)
back-to-back in one call. It's restartable at the individual-file level —
an interrupted run resumes without redoing completed files — but that
doesn't give a clean "this product is fully done" checkpoint if you're
deliberately spreading the download across several sessions rather than
leaving the machine running continuously for a multi-day pull.

Four narrower flags exist for exactly this — each downloads exactly one
product to completion:

```
--download-aps       PTGRAPS,  ~1976-1996
--download-sgml       PTGRSGM,  ~1996-2001
--download-xml         PTGRXML,  ~2001-present
--download-appxml   APPXML,   ~2001-present (applications)
```

A one-product-per-day rhythm looks like:

```
Day 1: --download-aps    --extract --database --format aps
Day 2: --download-sgml   --extract --database --format sgml
Day 3: --download-xml    --extract --database --format xml
Day 4: --download-appxml --extract --database --format appxml
Day 5: --filter
```

`--extract` and `--database` are both already safe to rerun every day
without redoing prior work — `--extract` checks a `.EXTRACTION_COMPLETE`
marker per archive before touching it again, and `--database` checks
`MANIFEST/database_manifest.csv` before re-parsing a file — so running
both after each day's download only costs real time on what's actually
new. **Run `--filter` only once all four products have completed** — it's
meant to query the complete database, so running it early just produces an
incomplete `filtered_patents.csv` you'd need to redo later.

## Output

| File | Contents |
|---|---|
| `RESULTS/filtered_patents.csv` | The patents that matched both CPC and keyword — patent number, patent type, claim category, title, dates, CPC/IPC, inventors, assignees, source path |
| `PDF/GRANTS/`, `PDF/APPLICATIONS/`, `PDF/OTHER/` | Downloaded PDFs, split by document type |
| `RESULTS/filtered_patents_<product>.csv` | Per-product filter results (`aps`, `sgml`, `xml`, `appxml`), merged into `filtered_patents.csv` |
| `DATABASE/USPTO_APS.sqlite`, `USPTO_SGML.sqlite`, `USPTO_XML_GRANTS.sqlite`, `USPTO_APPLICATIONS.sqlite` | The full parsed corpus, one file per product (not just filtered matches) |
| `MANIFEST/download_manifest.csv`, `MANIFEST/database_manifest.csv`, `MANIFEST/pdf_manifest.csv` | Restart bookkeeping — what's already been downloaded/parsed/fetched, and why anything failed |
| `LOGS/pipeline.log` | Full run log |

## Database layout (v15)

**One SQLite file per product** under `DATABASE/`:
`USPTO_APS.sqlite`, `USPTO_SGML.sqlite`, `USPTO_XML_GRANTS.sqlite`,
`USPTO_APPLICATIONS.sqlite`. A run only writes to the file for the product
it's working on, so a crash or drive fault can only affect that product,
and each file stays small enough that checkpoints, index builds and
integrity checks are quick. `--filter` reads all four in turn; results are
identical to the old single-file layout.

**Two tables per file:**

| Table | Contents |
|---|---|
| `patents` | Narrow metadata: numbers, document type, dates, `cpc`, `ipc`, title, inventors, applicants, assignees, source file/path. Indexed on `patent_number` and `publication_date`. |
| `patent_text` | `abstract` (plain text), `claims` and `description` (zlib-compressed BLOBs). Same `id` as the `patents` row. |

Why: the CPC check only needs the narrow table, and full text is read only
for patents that already passed CPC. Compression shrinks claims/description
roughly 3–4×, which directly cuts disk I/O on the external drive. To read
text yourself: `zlib.decompress(blob).decode("utf-8")`. Set
`COMPRESS_TEXT = False` to store plain text instead (both are read
correctly).

### Upgrading from v14 (one-time)

If `DATABASE/USPTO_PATENTS.sqlite` exists, `--database` and `--filter`
refuse to run until you do one of:

```
python uspto_patent_pipeline_v15.py --migrate-db
```

This copies every existing row into the per-product files (routed by
`source_path`) without re-parsing anything. It's resumable (progress saved
in `DATABASE/migration_v15_state.json`) and never modifies the old file.
When it prints per-product counts that look right, delete
`USPTO_PATENTS.sqlite` and `DATABASE/BACKUP/USPTO_PATENTS.sqlite` yourself.
The existing `database_manifest.csv` stays valid, so `--database --format
xml` then picks up exactly where the XML stage left off.

Alternatively, delete the old database **and** `MANIFEST/database_manifest.csv`
to re-parse everything from scratch.

## `--database` performance and machine load

- **Streaming, bounded memory.** Each worker process parses a file as a
  stream and writes records to SQLite every `DB_INSERT_BATCH_RECORDS`
  (default 250) records, returning only counts to the main process. Peak
  RAM per worker is one batch (tens of MB) regardless of file size. v14
  instead returned every parsed record of a whole weekly file to the main
  process — several GB per in-flight file at the XML stage, which is what
  exhausted 16 GB of RAM.
- **Parallel parsing.** `DATABASE_PARSE_WORKERS` processes (default 4;
  override per run with `--workers N`). Workers run at below-normal
  priority on Windows so the desktop stays usable, and on Python 3.11+ each
  is recycled after `DATABASE_WORKER_MAX_TASKS` files to stop slow heap growth.
- **Short write transactions.** Each batch commits on its own (SQLite
  serializes the workers' writes; each takes a fraction of a second). The
  WAL file is capped by `journal_size_limit`, instead of growing across
  200 files as in v14.
- **File-level resumability.** `MANIFEST/database_manifest.csv` tracks
  finished files (path + size). A file that failed part-way keeps the rows
  it committed, is marked `error`, and is retried next run; `INSERT OR
  IGNORE` skips anything already present, so nothing duplicates.
- **Deferred indexes.** `idx_patent_number` and `idx_publication_date` are
  built after the insert, on the narrow table, so they're fast. v14's
  `idx_cpc` was removed: `cpc` holds a delimited list matched by prefix,
  which that index couldn't serve.
- **Worker-crash recovery.** If a worker process dies (`BrokenProcessPool`,
  often alongside `WinError 1450`), a fresh pool continues on the remaining
  files, up to `DATABASE_MAX_POOL_RESTARTS` times in a row. Anything
  unfinished is left for the next run.
- **Drive errors stop the run.** A SQLite write failure (e.g. `disk I/O
  error`) stops the run cleanly with an integrity-check command to run,
  rather than continuing to write to a failing drive.
- **Backups are opt-in.** `--backup` copies each processed product's file to
  `DATABASE/BACKUP/`. Best used once per product when it's complete, ideally
  with `DATABASE_BACKUP_DIR` on a different drive. (v14 copied the entire
  combined database after every run.)

**If the machine still gets hot or unresponsive:** run with `--workers 3` or
`--workers 2`. An abrupt power-off (rather than a freeze or a Python error)
usually points to overheating or power delivery rather than RAM: Windows
Event Viewer will show a Kernel-Power 41 event, and a temperature monitor
such as HWiNFO will show whether the CPU is hitting its thermal limit.

**Windows note:** `ProcessPoolExecutor` uses `spawn` on Windows, which
re-imports this module in every worker. The CPC/keyword Excel loading is
guarded behind `if __name__ == "__main__":` so workers skip it.

## Merged-record fix and `--requeue-merged` (v15.1)

In some weekly XML files (confirmed in 2023 applications, e.g.
`ipa230914.xml`, 74 times in that one file) USPTO writes a patent's
`<?xml ...?>` declaration on the same line as the closing tag of the
sequence listing before it (`</sequence-cwu><?xml ...`). Before v15.1 the
splitter only recognised declarations at the start of a line, so that patent
was lost, logged as "Recovered the FIRST document from a merged/junk-after-
document chunk ... col 15". v15.1 splits such lines first; other files parse
exactly as before.

To recover patents lost in files parsed before the fix (one-time):

```
python uspto_patent_pipeline_v15.py --requeue-merged
python uspto_patent_pipeline_v15.py --database --format xml --retry-errors-only
python uspto_patent_pipeline_v15.py --database --format appxml --retry-errors-only
```

`--requeue-merged` reads `LOGS/pipeline.log` for those warnings and marks the
affected files for reparse. Already-stored records are skipped, so only the
missing patents are added.

## Filtering in sections (v15.2)

`--filter` can run one product database at a time, like `--database`:

```
python uspto_patent_pipeline_v15.py --filter --format aps
python uspto_patent_pipeline_v15.py --filter --format sgml
python uspto_patent_pipeline_v15.py --filter --format xml
python uspto_patent_pipeline_v15.py --filter --format appxml
```

Each section writes `RESULTS/filtered_patents_<product>.csv`, and after every
run `RESULTS/filtered_patents.csv` (what `--pdf` and the forward-citations
script read) is rebuilt from all the per-product files filtered so far. A
plain `--filter` still does all four in one go. Re-running a section replaces
only that section's results. If you change the CPC/keyword workbooks partway
through, re-run the sections already done so every product uses the same
criteria.

## Downloading PDFs in sections (v15.3)

`--pdf` takes the same controls as `--database`:

```
python uspto_patent_pipeline_v15.py --pdf --format aps --limit 500
python uspto_patent_pipeline_v15.py --pdf --format appxml
python uspto_patent_pipeline_v15.py --pdf --retry-errors-only
```

- `--format <product>` downloads only that section's filter results
  (`filtered_patents_<product>.csv`); without it, `--pdf` uses the combined
  `filtered_patents.csv`.
- `--limit N` stops after N downloads this run; run again to continue.
- Previously `FAILED` PDFs are skipped by normal `--pdf` runs, so no `--limit`
  batch is spent on them. Once everything else is downloaded, retry them with
  `--pdf --retry-errors-only` (add `--format`/`--limit` as usual).
- The PDF manifest is shared, so anything already `COMPLETE` is skipped no
  matter which section requests it.

## Patent type and claim category (v15.10)

Every filter results file carries three classification columns, right after
`document_type`:

| Column | Where it comes from | Values |
|---|---|---|
| `patent_type` | The number / kind code (official USPTO document types) | Utility, Design, Plant, Reissue, Statutory Invention Registration (SIR), Defensive Publication, Application (utility), Application (plant) |
| `claim_category` | Claim 1's preamble — the statutory category of what's claimed (35 U.S.C. 101) | Process / Method, Apparatus / System, Product / Article, Composition, Product-by-process, Design, Plant, Unclassified, No claims text |
| `independent_claim_categories` | Every category found across all independent claims, in claim order | e.g. `Process / Method; Apparatus / System` for a patent claiming both the method and the machine |

`patent_type` is exact. `claim_category` is read from the claim wording, since
USPTO doesn't record "product vs process" as a field: the head noun of the
preamble decides it ("A film winding **apparatus**…" → Apparatus / System,
"A multilayer **film**…" → Product / Article, "A **method** of…" →
Process / Method, "A polyethylene **composition**…" → Composition; "…obtained
by the process of…" → Product-by-process). A preamble it can't place is left
`Unclassified` rather than guessed — the word lists in the "PATENT TYPE +
CLAIM CATEGORY" section of the script are where to extend it. A claim that
refers back to another claim but opens a new category ("An apparatus for
carrying out the method of claim 1") counts as independent; plain dependent
claims and bracketed (cancelled) reissue claims don't.

New `--filter` runs fill the columns automatically. To add them to results
you've **already** filtered without re-running the filter:

```
python uspto_patent_pipeline_v15.py --classify
python uspto_patent_pipeline_v15.py --classify --format xml
```

`--classify` looks each filtered patent's claims up in its product database
(by `publication_number`, which is indexed), rewrites
`filtered_patents_<product>.csv`, rebuilds `filtered_patents.csv`, and prints
a count per category. It only touches the filtered patents, so it takes
minutes rather than a full filter pass. Run it before the forward-citations
`--export` so those columns carry through to `forward_citations.xlsx`.

Design patents rarely show up in the filter results at all — they carry no
CPC/IPC (see Known gaps) — so expect `Design` to be uncommon.

## `--filter` performance

Two-phase: it scans only the narrow `patents` table for the CPC/IPC gate,
then fetches text for CPC matches only. Keywords are compiled into one
combined, case-insensitive, word-bounded regex, and fields are checked
title → abstract → claims → description, stopping at the first hit (so the
description is often never decompressed).

## Known gaps (not silently ignored — the code raises/logs these)

- **`get_pdf_url()`** doesn't have a confirmed URL format for reissue
  (`RE`) or Statutory Invention Registration (`H`) patents — these fail
  cleanly rather than guessing at a URL.
- **Documents over ~1000 pages** may not be retrievable through the PDF
  endpoint (per USPTO's own docs) — not independently confirmed what the
  failure looks like, but the existing retry/validation logic catches it
  rather than corrupting anything.
- **CPC didn't exist before 2013.** `MATCH_IPC_AS_CPC_FALLBACK = True`
  makes pre-2013 records matchable via their IPC classification instead.
- **Design patents carry no CPC or IPC at all** in USPTO's XML — they use
  Locarno classification instead. This means `matches_cpc()` will exclude
  *every* design-patent grant from `--filter`, regardless of subject
  matter, since there's nothing in the CPC/IPC fields to match against.
  This is deliberate, documented parser behavior, not a bug — but worth
  knowing if design patents matter to your project. Confirmed against a
  real granted design patent: zero `classifications-cpc`/`classifications-
  ipcr` blocks present in its XML.
- **A patent must match BOTH CPC and keyword** to be selected — a patent
  can use on-topic vocabulary (e.g. mentioning "film") while being
  classified in a completely unrelated CPC section, and will correctly be
  excluded. Confirmed against real data: a marine-hardware patent (IPC
  `B63B`) whose abstract repeatedly says "flexible synthetic plastic film"
  is excluded by the CPC gate despite the keyword match.

## Forward citations

An earlier `--citations` stage (which queried USPTO's PatentsView API for
a citation *count*) has been removed — PatentsView's live query API was
confirmed shut down in March 2026, and even before that its citation
endpoint was never officially documented. Forward-citation cross-referencing
now lives in a separate script, **`uspto_forward_citations_v1.py`** — see
its own README. It reads `filtered_patents.csv` (produced by `--filter`
above) and scans the local extracted corpus directly, so **run this
pipeline through at least `--filter` first**, then run the forward-citations
script — `--pdf` is not a prerequisite for it.

## Environment

- Python 3.9+ (tested on 3.9 and 3.12; worker recycling needs 3.11+ and is skipped on older versions)
- Windows, run from PyCharm's built-in terminal (or any terminal)
- 8 TB external drive recommended for a full historical-corpus run
- SSD preferred for the database stage specifically (WAL checkpointing,
  index construction, and parallel-process I/O) — everywhere else, drive
  type barely matters
