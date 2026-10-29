# USPTO Forward Citations (`uspto_forward_citations_v5.py`, v5.1)

For every patent selected by the [patent pipeline](README_pipeline_v15.md), finds the **later** patents that cite it, so you can see which patents the field actually built on.

| | |
|---|---|
| **Input** | `D:\uspto_corpus\RESULTS\FILTERED\filtered_patents.csv` (from the pipeline's `--filter`) + the weekly archives in `D:\uspto_corpus\RAW\` |
| **Main output** | `D:\uspto_corpus\RESULTS\FORWARD CITATIONS\forward_citations.xlsx` |
| **Run time** | Hours for the full corpus (scan); a few minutes to re-export |
| **Needs** | The pipeline run through `--filter` (and ideally `--classify`). PDFs are not needed. |
| **Next step** | The [overview report](README_overview_report_v1.md) |

*(This README keeps its original file name so existing links still work; it documents script v5.1.)*

---

## Why it exists

Every patent lists the earlier patents it cites ("references cited"). A **forward citation** is the reverse: a later patent listing *this* one. Forward citations are the standard measure of a patent's influence, but nothing in USPTO's data says "who cites me". The script works it out by reading the reference list of every patent in the corpus and keeping the ones that point at a selected patent.

It runs entirely offline. It replaced the PatentsView API, which shut down in March 2026.

## Quick start

```
python uspto_forward_citations_v5.py --sample 25 --all   # trial: scan 25 archives, then export
python uspto_forward_citations_v5.py --all               # full scan + export (restartable)
python uspto_forward_citations_v5.py --verify            # check how patent numbers matched
```

Open `forward_citations.xlsx`: the most-cited patents are at the top, and the **Run Notes** sheet says what the run covered.

> **If the filter results change** (you edited the search workbooks and re-ran `--filter`), run `--reset` before scanning again. The index remembers which archives it has read, not which patents it was looking for, so it would otherwise skip archives that contain citations to the new patents.

---

## Before you run

- **One setting must match the pipeline:** `BULK_DATA_ROOT_DIR = Path(r"D:\uspto_corpus")` at the top of the script. Every other path is derived from it.
- **Run the pipeline first**, through at least `--filter`. Run `--classify` too if you want the patent type and claim category columns filled in.
- **Keep `RAW\`** (the downloaded zip archives): the scan reads them directly. (`--source extracted` reads the unzipped copies instead.)
- This script never writes to the pipeline's database, so running one can't break the other.

---

## What it does

| Stage | Command | What happens |
|---|---|---|
| Scan | `--scan` | Reads every weekly archive (4,095 files, about 20 million records, 223 million cited references), extracts each record's US patent citations, and stores the ones that point to a selected patent in a citation index. Restartable: finished archives are skipped. |
| Export | `--export` | Counts citations per selected patent, splits them by who added them, computes rates and percentiles, and writes the Excel and CSV results. |
| Verify | `--verify` | Reports how patent numbers were matched and writes a per-patent match report. |
| Both | `--all` | Scan, then export. |

**Where citations come from in each format:**

| Format | Citations read from | Who added the citation? |
|---|---|---|
| XML (2002–present grants, applications) | `<us-references-cited>` / `<references-cited>` | Recorded: examiner, applicant or third party |
| SGML (2001 grants) | `B561` patent-citation blocks | Partly: examiner or "other" (not assumed to mean applicant) |
| APS (1976–2001 grants) | `UREF` sections | Not recorded: reported as unknown |

Only US patent citations count; foreign patents and non-patent literature can't match a US patent number. Only the first part of each record is read (bibliographic data and citations always come before the long description), which keeps the scan fast.

**Number matching.** The script normalises every number format the pipeline stores (`US10123456B2`, `US06167569A`, `USD0435713S`, `USPP011728P`, `USRE038552E1`...). APS-era numbers carry an extra check character (`US039302711` is patent 3,930,271); it's removed for APS records, and both readings are tried where the source isn't certain. `--verify` shows which reading matched.

---

## What you get

All in `D:\uspto_corpus\RESULTS\FORWARD CITATIONS\` unless noted.

| File | What it is |
|---|---|
| `forward_citations.xlsx` | **Forward Citations** sheet: one row per selected patent, most-cited first. **Pairs** sheets: every citing/cited pair, one sheet per format, continuing onto "(2)" sheets past 1 million rows. **Run Notes**: what the run covered, the coverage year, caveats and counts per claim category. |
| `forward_citation_pairs.csv` | Every citing/cited pair in one file: the complete record (3.5 million rows; too large for Excel). |
| `forward_citation_pairs_<format>.csv` | The pairs split by format (`aps`, `sgml`, `xml`, `appxml`), with `_part2`, `_part3`... files so each one opens in Excel. |
| `forward_citation_match_report.csv` | From `--verify`: how each patent's number matched. |
| `D:\uspto_corpus\DATABASE\FORWARD_CITATIONS.sqlite` | The citation index (restart state). Delete it only with `--reset`. |

### Reading the Forward Citations sheet

| Column | Meaning and how to use it |
|---|---|
| Filtered Patent ID, Title, Grant Date, Assignee, CPC | The selected patent, from the filter results. |
| Patent Type, Claim Category, Independent Claim Categories | From the pipeline's `--filter` / `--classify` (blank claim category = results filtered before v15.10; run `--classify`, then `--export` again). |
| **Times Cited** | Number of later US documents citing it. Higher means more influence, but older patents have had longer to collect citations. |
| **Citations Per Year** | Times Cited ÷ years since grant. Compares patents of different ages. |
| **Cohort Percentile** | Rank against the other selected patents granted in the same 5-year window: 95 means cited more than 95% of its peers. The fairest single score. Blank when the date didn't parse or fewer than 5 patents share the window. |
| Cited by Applicant | Added by the later patent's own applicant: a deliberate reference and the strongest signal. |
| Cited by Examiner | Added by the patent examiner during the search; common and less meaningful. |
| Cited by Other/Unknown | SGML-era "other" citations, plus all APS-era citations (which don't record who added them). |
| Citing Patent IDs | The citing patents, split over several columns when the list exceeds Excel's 32,767-character cell limit. |

### Reading the pairs files

One row per citation: `filtered_patent_number`, `filtered_patent_type`, `filtered_claim_category`, `canonical_key` (normalised number used for matching), `citing_patent_number`, `citing_patent_type`, `citing_patent_date`, `citation_category` (applicant / examiner / other / unknown), `filtered_patent_title`. Use them to see *who* cites a patent, or to build your own year-by-year citation curves.

---

## Checking a run

- **Console:** the scan prints records read, references read, hits and files with errors; the export prints the time per step and how many patents got a rate and a percentile.
- **`FORWARD_CITATIONS.sqlite` → `scanned_files` table:** one row per archive with records, citations seen, hits and any error. Files with an error are retried automatically on the next `--scan`.
- **Run Notes sheet:** confirms the input file, coverage year and totals.
- **`--verify`:** run it after a full scan to confirm number matching worked.

Current state: 4,095 archives scanned with no errors; 3,512,368 citations to 292,706 of 474,183 selected patents.

---

## Need to know: assumptions and limits

- **US citations only.** Citations from foreign patents aren't in USPTO data.
- **Only the corpus on the drive counts.** A `--sample` run gives sample-sized numbers, not real totals.
- **Recent patents look less cited** because they've had little time. Use Citations Per Year or Cohort Percentile for them.
- **Published applications rarely carry a citation list**, so almost all citing documents are granted patents.
- **The per-year rate is an approximation.** Most patents collect citations fastest 3–6 years after grant, not at a steady rate.
- **Percentiles are relative to the selected patents only**, not to all of USPTO. Indexing every citation in the corpus would be far larger.
- **Rates use the latest citing year in the data as "now"** (auto-detected). Set `CORPUS_COVERAGE_END_DATE = "YYYYMMDD"` if you know the exact end of your downloaded data.
- **Re-run `--reset` + `--scan` whenever the filter results change** (see Quick start).

---

## Command reference

| Command | Does |
|---|---|
| `--scan` | Build or continue the citation index |
| `--export` | Write the Excel and CSV results from the index |
| `--all` | Scan + export |
| `--sample N` | Scan only the next N unscanned archives (combine with `--all` for a trial) |
| `--verify` | Report how numbers matched; write the match report |
| `--reset` | Delete the citation index and start over |
| `--source zip\|extracted` | Read the zip archives (default, much faster on a spinning drive) or the unzipped files |
| `--workers N` | Scan processes (default 4) |
| `--skip-excel-pairs` | Leave the pairs sheets out of the workbook: the fastest export (the CSVs still have every pair) |

### Settings (top of the script)

| Setting | Default | Controls |
|---|---|---|
| `SCAN_WORKERS` | 4 | Scan processes. More than about 4 on a spinning drive is usually slower; 8 is reasonable on an SSD. |
| `SCAN_SOURCE` | `"zip"` | Read archives directly (`zip`) or the unzipped files (`extracted`) |
| `EXCEL_PAIR_ROW_LIMIT` | 1,000,000 | Rows per pairs sheet / pairs CSV part (Excel's limit is 1,048,576) |
| `EXCEL_INCLUDE_PAIRS_SHEET` | True | Include pairs sheets in the workbook |
| `EXCEL_CELL_CHAR_BUDGET` | 30,000 | Characters per Citing Patent IDs cell before starting another column |
| `CORPUS_COVERAGE_END_DATE` | None (auto) | End of the data, for the per-year rate |
| `CITATION_COHORT_YEAR_WINDOW` | 5 | Width of the grant-year windows for Cohort Percentile |
| `CITATION_COHORT_MIN_SIZE` | 5 | Smallest window that gets a percentile |
| `CITATION_RATE_MIN_YEARS` | 1.0 | Minimum years used in the per-year rate (stops very new patents getting huge rates) |

---

## How it was checked

- **SGML:** two real grant records checked by hand; every citation and its examiner/other tag matched.
- **APS:** a real 1976 reissue patent checked by hand; all 8 citations found, and the reissued patent's own number (a different field) correctly not counted. The check-digit handling was confirmed on the same file.
- **XML:** a real design patent with 21 citations checked by hand; all 21 found, with the correct category.
- **Rates and percentiles:** checked against a synthetic set with known answers.
- Not yet confirmed on real data: what SGML "other" citations really mean, and real XML "applicant" or "third party" tags in a hand-checked record (the full run does report many applicant citations).

## Version history

| Version | Change |
|---|---|
| v5.1 | Patent type, claim category and citing patent type columns; reads `RESULTS\FILTERED`, writes `RESULTS\FORWARD CITATIONS` |
| v5 | Pairs split by format into Excel sheets and CSV parts, nothing truncated |
| v4 | Much faster Cohort Percentile; `--skip-excel-pairs`; step timings |
| v3 | Reads zip archives directly (5–10× less disk reading); worker processes; ignores extraction marker files |
| v2 | APS check character always removed; duplicate APS/SGML patents reported once; merged XML records split correctly |
| v1 | First offline version, replacing PatentsView |
