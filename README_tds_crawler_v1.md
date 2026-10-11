# tds_crawler_v1 — vendor technical data sheet harvester

Crawls every vendor in `Vendors_With_Technical_Sheets_By_Material_Category.xlsx`, collects the PDF links on each site, downloads the data sheets, and tags each one by document type and by the test methods it cites.

## Files

| File | Purpose |
|---|---|
| `tds_crawler_v1.py` | The crawler (Python 3.9+) |
| `vendor_seeds.csv` | Domains and start URLs for each vendor. **Edit this file**, not the script |
| `README_tds_crawler_v1.md` | This file |
| `requirements.txt` | Dependencies |

The script looks for the workbook and seed file in `D:\tds_corpus\input\` (`DEFAULT_WORKBOOK` and `DEFAULT_SEEDS` near the top of the script). Override them with `--workbook PATH` and `--seeds PATH`.

## Install

```
pip install -r requirements.txt
# optional, for JavaScript-heavy sites (Dow, ExxonMobil, SABIC, 3M, Henkel...):
pip install playwright && playwright install chromium
```

## Recommended run order (stage and validate)

```
python tds_crawler_v1.py --list-vendors                       # check seed status per vendor
python tds_crawler_v1.py --discover --vendors "rogers,camden,novelis" --max-pages 60 --contact you@example.com
python tds_crawler_v1.py --status                              # links found per vendor and type, nothing downloaded yet
python tds_crawler_v1.py --download --classify --export --vendors "rogers,camden,novelis"
python tds_crawler_v1.py --all --contact you@example.com       # full run, all crawlable vendors
python tds_crawler_v1.py --all --js                            # add Playwright for render_js=1 vendors
```

Running `--discover` on its own is the dry run: it lists every PDF link and how each was classified, without downloading anything.

## Output (default root `D:\tds_corpus`, change it with `--root`)

```
D:\tds_corpus\
  input\                   workbook + vendor_seeds.csv
  pdfs\<Category>\<Vendor>\*.pdf   e.g. pdfs\4_Polymer_Films\Toray_Industries\  (SHA-256 de-duplicated)
  pdfs\<Category>\<Vendor>\_translations\   foreign-language copies whose English version is in the folder above
  pdfs\<Category>\<Vendor>\_rejected\       downloaded files whose text shows junk: reports, IR, press releases... (delete when happy)
  tds_crawler.db            SQLite (WAL): pages, pdf_links, files, vendors
  tds_manifest.csv          one row per file: vendor, categories, title, doc type, test methods, source URL...
  tds_catalog.xlsx          Excel catalogue: Category | Vendor | Document (click to open) | Description | Type |
                            Test methods | Pages | Language | Source URL  (sheets: Catalog, Translations, Rejected)
  logs\tds_crawler_YYYYMMDD.log
```

Category folders come from the workbook sheet names: `1_Foil_Materials`, `2_Nonwoven_Fabrics`, `3_Specialty_Paper`, `4_Polymer_Films`, `5_Web_Converting`, `6_Equipment_Machinery`, `7_Process_Technology`. A company listed on more than one sheet (Dow, ExxonMobil, Oerlikon, Brückner) goes in the first sheet's folder, and `tds_manifest.csv` lists all of its categories. To choose a different folder for a company, fill in its `folder` column in `vendor_seeds.csv`. Files downloaded in an older layout are moved automatically on the next `--download`, or by running `--reorganize` on its own.

## Excel catalogue

`--export` (and `--all`) also writes `tds_catalog.xlsx` in the corpus root. It has one row per PDF, with the material category, supplier and file name; the file name is a link, so clicking it opens the PDF. Each row also has a one-line description built from the document's opening text (for example *Technical data sheet: DuPont Pyralux AG - All-Polyimide Double-Sided Copper-Clad Laminate*), plus the type, test methods cited, page count, language and source URL. Kept documents are on the **Catalog** sheet. Translated copies are on **Translations**, with their English twin named, and junk is on **Rejected**, with the reason. Keep the workbook in the corpus root so the links work. If it's open in Excel during a run, a time-stamped copy is written instead.

## Foreign-language documents (English first)

Many vendors post translated copies of the same sheet, such as `RO1200 Data Sheet - CN.pdf`. The crawler recognises these from a language code at the end of the filename (`- CN`, `_JP`, `(KR)`, `-DE`...), a language word (`_Chinese`), or a locale folder in the URL (`/zh-cn/`, `/ja/`), and always tries to get the English version first:

1. If the English version is among the links found, only the English one is downloaded.
2. If not, it tries likely English URLs: the same name without the code, or with `-EN` / `_EN`, or the `/en/` folder instead of `/ja/`.
3. If no English version exists, the foreign copy is downloaded anyway so no data is lost, and the manifest flags it.

Translated copies already on disk whose English version later turns up are moved to `_translations\`, not deleted. Add `--keep-translations` to download every language version.

After `--classify`, the manifest's `language` column is read from the PDF text: `en`, `zh`, `ja`, `ko`, `de`..., or `ja+en` for bilingual sheets. `english_version` gives the path of the English file for a translated copy, or `NOT FOUND` when a non-English sheet has no English counterpart. Filter on `NOT FOUND` to see which ones need a manual search or a request to the vendor.

## Junk filtering (corporate, IR, press)

Junk is filtered at three levels:

1. **While crawling.** Pages in investor-relations, sustainability/CSR, news, press, careers and finance sections are deprioritised and rarely fetched.
2. **Before download.** A PDF whose URL, or the page it was found on, sits in one of those sections (`/ir/`, `/sustainability/`, `/newsrelease/`...) is tagged `corporate` and skipped.
3. **After download.** `--classify` reads each file's first pages. Files whose text shows an unwanted type are moved to `_rejected\`, not deleted. Unwanted types are integrated or annual reports, management plans, balance sheets and press releases, in English, Japanese, Chinese and Korean, plus anything else not in `--doc-types`. The manifest's `rejected_as` column says why. Use `--keep-rejected` to switch this off. Besides the page text, the check also reads each PDF's built-in title and filename (all Unicode-normalised, so full-width or decomposed Japanese characters still match). A file that can't be judged, because its text is garbled or it's image-only, is treated as junk when most of the same supplier's files with the same name pattern are junk; for example `3003_nmm.pdf` next to `2903_nmm.pdf`, `201903_nmm.pdf` and so on. These rules apply to every supplier. The log line `Text extraction: pypdf x.y` shows which PDF library version did the reading. Keep it current with `pip install -U pypdf`.

Running `--reclassify` re-checks every file already on disk with the current rules. It also pairs foreign files with their English twin when only the English file is marked, for example `C1990_data.pdf` (Japanese) and `C1990_data_english.pdf`.

## Distributors and brand filters

When a supplier's sheets are hosted on a distributor's site (Camden Copper and JX Metals on insulectro.com), the supplier's `pdf_filter` seed column keeps only its own branded files, for example `insulectro.com:camden|denkai|\bMHT\b|\bHP2\b`. Every other PDF on that site goes to the distributor's own row (Insulectro). Files already downloaded under the wrong supplier are re-assigned and moved automatically on the next `--download` or `--reorganize`.

## How it works

**Discovery.** The crawler reads robots.txt and the sitemaps, then follows links from the start URLs in order of a relevance score. Words like datasheet, TDS, technical, download and products in the URL or link text raise the score; news, careers, investor and login lower it. It stays inside each vendor's `domains` and optional `path_prefixes`, and skips non-English locale folders such as `/de/` and `/zh-cn/`. It catches PDFs in `<a>`, iframe and embed tags, data-attributes, and URLs embedded in script/JSON, plus download endpoints that return `application/pdf` from a URL that doesn't end in .pdf.

**Download.** Each file is streamed and kept only if it starts with a `%PDF` header, which rejects HTML error pages saved as .pdf. Files larger than `--max-pdf-mb` are skipped, and a PDF reachable from several URLs is stored once. The filename comes from the Content-Disposition header when the server provides one.

**Link types.** Each link is classified from its URL and link text as `tds`, `sds`, `brochure`, `certificate`, `corporate` or `other`. By default the crawler downloads `tds,brochure,other`; change that with `--doc-types tds` or `--doc-types all`. The `other` type is kept because many data sheets are named only by product code, such as `MHT.pdf`.

**Classification.** After download, pypdf reads the text of pages 1–2 and records the document type from the text, the title, the page count, and the test methods cited: ASTM D882, ISO 527, TAPPI T410, NWSP 100.1, IPC-TM-650 2.4.8, DIN, JIS. The `method_families` column lets you separate film, paper and nonwoven data before comparing values. A file whose name says "data sheet", or whose first pages cite two or more test methods, is classed as a data sheet unless it is a safety data sheet.

**Resumable.** Ctrl-C saves state, and running the same command again continues from there. A vendor that hits `--max-pages` keeps its queue, so another `--discover` run goes deeper.

**Revisions.** `--download --refresh` re-checks each downloaded PDF with ETag and If-Modified-Since. A changed sheet is saved as a new file and the old one is kept.

## Search seeding (optional, best for 3M / Dow / ExxonMobil)

Set `SERPAPI_KEY`, or `GOOGLE_CSE_KEY` + `GOOGLE_CSE_CX`, then add `--search-seed`. For each domain the crawler queries `site:<domain> filetype:pdf ("technical data sheet" OR datasheet ...)` and adds the hits directly.

## Editing vendor_seeds.csv

| Column | Meaning |
|---|---|
| status | `ok` crawl · `verify` crawl, but domain or relevance unconfirmed · `skip` don't crawl |
| domains | `;`-separated; subdomains are included automatically |
| start_urls | `;`-separated entry points |
| path_prefixes | optional, e.g. `/fibers/`, to stay inside one section of a huge site |
| render_js | `1` = render with Playwright when `--js` is given |
| folder | optional, overrides the category folder, e.g. `5_Web_Converting` |
| pdf_filter | optional, `[domain:]regex`. On a shared site, keep only PDFs whose filename or link text matches |

Translated duplicates are skipped and junk such as ethics/hotline pages, licence agreements, testimonials and success stories is filtered out at link level.

When a vendor appears in the workbook but has no seed row, the script adds one automatically with `status=skip`.

## Politeness defaults

The crawler waits 2 s between requests to each vendor, or longer if robots.txt sets a Crawl-delay. It processes 4 vendors in parallel, backs off on 429 and 503 responses, follows robots.txt rules, and sends an identifying User-Agent; add your contact with `--contact`. Site terms vary, and some vendors (UL Prospector-style portals, 3M) forbid automated download, so check before running high volumes.
