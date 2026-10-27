# USPTO Forward-Citation Cross-Reference (`uspto_forward_citations_v1.py`)

Takes the patent IDs in `filtered_patents.csv` (produced by the main
pipeline's `--filter` stage) and finds every OTHER patent in the local
corpus that **cites** each one — the reverse of what's in the raw data,
since each patent record only lists what *it* cites, never who cites it.

## Relationship to `uspto_patent_pipeline_v12.py`

This is a separate, standalone script — **run the main pipeline first**,
through at least `--filter` (`--pdf` is NOT required — this script never
reads PDFs at all, it only reads the raw extracted XML/SGML/APS files and
`filtered_patents.csv`). It never touches the main pipeline's SQLite
database, so there's nothing to reconcile between the two and nothing you
can break in one by running the other.

```
python uspto_patent_pipeline_v12.py --filter        # produces filtered_patents.csv (--pdf optional, not needed here)
python uspto_forward_citations_v1.py --all           # then this
```

## Before running

Only one setting needs to match the main pipeline:

```python
BULK_DATA_ROOT_DIR = Path(r"D:\uspto_corpus")   # same value as in v12
```

Everything else (`EXTRACTED_DIR`, `RESULTS_DIR`, etc.) is derived from it
the same way the main pipeline derives its own paths.

## Running it

```
python uspto_forward_citations_v1.py --sample 25 --all   # start here — scan+export on a small slice
python uspto_forward_citations_v1.py --all                 # full corpus scan + export
python uspto_forward_citations_v1.py --verify                # coverage stats + match-quality report
python uspto_forward_citations_v1.py --reset                 # delete the citation index and start over
```

Unlike the main pipeline, `--sample` and `--all` **do** combine correctly
here — `python uspto_forward_citations_v1.py --sample 25 --all` really
does limit the scan to ~25 files.

`--scan` builds a restartable SQLite index of citation hits; `--export`
writes the Excel/CSV output from whatever's in that index; `--all` runs
both. Scanning is safe to interrupt and restart — already-scanned files
(matched by size) are skipped, and files that errored are automatically
retried on the next run rather than silently marked done.

**One thing `--reset` and restartability don't cover automatically:** the
index tracks "already scanned" per *file*, not per *target set*. If you
rerun `--filter` in the main pipeline with a different CPC/keyword matrix,
producing a different `filtered_patents.csv`, files already scanned under
the old target set will be skipped — even if they'd contain a real hit
against your new filtered patents. Run `--reset` before `--scan` whenever
`filtered_patents.csv`'s *content* has changed, not just when re-running
against the same one.

## What it reads and how it matches

Scans every file under `EXTRACTED_DIR`, auto-detecting format (APS/SGML/XML
— same three eras the main pipeline handles), and reads each record's
**cited-reference list**:

| Format | Where citations live | Category available? |
|---|---|---|
| XML (grants, ~2001–present) | `<us-references-cited>` / `<references-cited>`, one `<patcit>` per `<us-citation>` | Yes — `<category>`: examiner / applicant / third-party |
| SGML (~1996–2001) | `B560` → `B561` blocks (patent cites); `B562` (non-patent lit) is skipped | Partial — `CITED-BY-EXAMINER` / `CITED-BY-OTHER` only (this era predates the three-way split; `OTHER` is **not** assumed to mean "applicant") |
| APS/Green Book (~1976–1996) | `UREF` sections, `PNO` field; `FREF` (foreign) and `OREF` (non-patent) skipped; `REIS` section's own `PNO` (the original patent being reissued) is correctly excluded | No — this era carries no per-reference provenance marker at all; always reports `unknown` |

Only US patent-document citations count — foreign families and non-patent
literature are excluded, since they can't match a US patent ID anyway.

**Patent-number normalization** handles every format the main pipeline
produces (`US10123456B2`, `US06167569A`, `USD0435713S`, `D. 254336`,
`2790186`, `USPP011728P`, ...) plus the APS-era **check-digit ambiguity**:
APS numbers carry a trailing check digit the main pipeline stores raw, so
`US39302711` is really patent 3,930,271. Both interpretations are tried,
and `--verify` reports which one actually matched real citations.

## Output

| File | Contents |
|---|---|
| `RESULTS/FORWARD CITATIONS/forward_citations.xlsx` | One row per filtered patent — see columns below |
| `RESULTS/FORWARD CITATIONS/forward_citation_pairs.csv` | **Authoritative, unlimited** long-format version: one row per citing/cited pair, with the citing patent's own date and that specific citation's category. Also carries the filtered patent's type and claim category and the citing patent's type (v5.1). Excel has hard row/cell limits a large corpus can exceed; this CSV never truncates. |
| `RESULTS/FORWARD CITATIONS/forward_citation_match_report.csv` | Per-patent detail on how number matching resolved (from `--verify`) |
| `DATABASE/FORWARD_CITATIONS.sqlite` | The restartable citation index itself |

### `forward_citations.xlsx` columns

| Column | Meaning |
|---|---|
| Patent Type | USPTO document type from the number/kind code — Utility, Design, Plant, Reissue, SIR, Defensive Publication, Application (utility/plant) |
| Claim Category | What claim 1 claims — Process / Method, Apparatus / System, Product / Article, Composition, Product-by-process, Design, Plant (from pipeline v15.10's `--filter` or `--classify`; blank if `filtered_patents.csv` is older) |
| Independent Claim Categories | Every category across the patent's independent claims, e.g. `Process / Method; Apparatus / System` |
| Times Cited | Raw forward-citation count |
| Citations Per Year | Times Cited ÷ years the patent has sat inside the corpus's *observable* window — see "Rate normalization" below |
| Cohort Percentile | Times Cited ranked against other **filtered** patents granted in the same multi-year window — blank if the patent's date didn't parse or its cohort has too few members |
| Cited by Applicant | Citations where the citing document explicitly pointed to this prior art — the stronger influence signal |
| Cited by Examiner | Citations that mostly reflect the examiner's own search overlap, not necessarily the field's use of the invention |
| Cited by Other/Unknown | SGML's non-examiner "other" (may or may not be applicant-sourced) plus all APS-era citations, which carry no provenance marker |
| Citing Patent IDs | Split across multiple columns if the list would exceed Excel's ~32,767-character cell limit |

Rows are sorted most-cited first.

**Patent type / claim category (v5.1).** The type and claim-category columns
come straight from `filtered_patents.csv`. If that file was filtered before
pipeline v15.10, run `uspto_patent_pipeline_v15.py --classify` once, then
`--export` here again — no re-scan needed (Patent Type is filled from the
number either way). The pairs sheets add Filtered Patent Type, Filtered
Claim Category and Citing Patent Type, and Run Notes lists how many filtered
patents fall in each claim category and how often each group was cited. A "Run Notes" sheet records exactly what
the run covered, what coverage year was used, and all the caveats below.

## Rate normalization — what it means and what it doesn't

A raw citation count is a weak value signal on its own: it's dominated by
how *old* a patent is (more years to accumulate citations) and by how
citation *practice* has drifted across a 1976–present corpus.

**Citations Per Year.** The denominator is capped at the corpus's own
forward edge, not "years since grant to today" — a patent shouldn't get
credit for years the corpus can't actually show citations for. By default
this edge is auto-detected as the latest `citing_date` found anywhere in
the index; set `CORPUS_COVERAGE_END_DATE = "YYYYMMDD"` near the top of the
file if you know your real `--download`/`--extract` range precisely. A
floor (`CITATION_RATE_MIN_YEARS`, default `1.0`) prevents a division blowup
for patents granted very recently.

**Cohort Percentile.** Ranks each patent only against *other filtered
patents* sharing its grant-year cohort (`CITATION_COHORT_YEAR_WINDOW`,
default 5 years) — not against the whole USPTO corpus. Building a
corpus-wide percentile would require indexing every citation edge in
existence rather than just the ones landing on your target set, which is
the exact scale problem the target-key filtering is designed to avoid.
Cohorts smaller than `CITATION_COHORT_MIN_SIZE` (default 5) are left
blank rather than reporting a percentile computed against 2–3 peers.

**Caveats that don't go away:**
- The citation-arrival curve is itself skewed — most patents peak roughly
  3–6 years after grant, not on a straight line — so a flat per-year rate
  is an approximation, not a true accumulation model.
- Counts reflect only the corpus present in `EXTRACTED_DIR`. A `--sample`
  run gives sample-scale numbers, not true totals.
- A patent can only be cited by documents published *after* it.
- Pre-grant applications usually carry no references-cited section at
  publication, so citing documents are mostly grants.

## Config reference

| Constant | Default | What it controls |
|---|---|---|
| `SCAN_WORKERS` | `4` | Parallel file-scan threads. More than ~4 on an external HDD is usually slower (seek thrash); 8 is reasonable on an SSD. |
| `EXCEL_PAIR_ROW_LIMIT` | `1,000,000` | Row cap on the "Citation Pairs" Excel sheet (Excel's hard limit is 1,048,576). The CSV never truncates. |
| `EXCEL_CELL_CHAR_BUDGET` | `30,000` | Character budget per "Citing Patent IDs" cell before splitting into another column. |
| `CORPUS_COVERAGE_END_DATE` | `None` (auto-detect) | Set to `"YYYYMMDD"` for a precise Citations Per Year denominator. |
| `CITATION_COHORT_YEAR_WINDOW` | `5` | Grant-year bucket width for Cohort Percentile. |
| `CITATION_COHORT_MIN_SIZE` | `5` | Minimum patents sharing a cohort before a percentile is reported. |
| `CITATION_RATE_MIN_YEARS` | `1.0` | Floor on the Citations Per Year denominator. |

## Verified against real USPTO data

- **SGML**: extraction hand-checked against two genuine uploaded grant
  records — every citation, and the real `CITED-BY-EXAMINER`/
  `CITED-BY-OTHER` tags, matched exactly.
- **APS/Green Book**: hand-checked against a real 1976 reissue patent
  (`RE0286710`) — correctly extracted all 8 genuine `UREF`/`PNO` citations
  (including a real duplicate) while correctly *excluding* the `REIS`
  section's own `PNO` (the original patent being reissued), which is a
  different field in a different section, not a cited reference. Also
  confirmed the APS check-digit round trip end-to-end against this same
  file.
- **XML**: hand-checked against a real granted design patent
  (`USD0707016S1`) with 21 genuine citations — all 21 extracted correctly,
  including design-patent (`D`-prefixed) cite numbers and the real
  `cited by examiner` category tag.
- **Rate/percentile math**: verified against a controlled synthetic set
  with known citation counts (0/1/2/3/4/5 across one cohort) — resulting
  percentiles matched hand-calculated values exactly (16.7/33.3/50.0/
  66.7/83.3/100.0).

Not yet confirmed against real data: an SGML `CITED-BY-OTHER` citation's
true real-world meaning, and an XML citation actually carrying
`cited by applicant` or `cited by third party` (every real record seen so
far uses `cited by examiner`).
