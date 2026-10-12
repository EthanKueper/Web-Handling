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

## Which version ran?

The first log line of every run shows the build, the rules version and the full path of the script that ran, for example `tds_crawler_v1 build 2026-10-11 r8 (rules v8) - C:\...\tds_crawler_v1.py`. Browsers often save a re-downloaded file as `tds_crawler_v1 (1).py`, so check that line after replacing the script. Whenever the sorting rules change, `--classify` (or `--all`) re-checks every file automatically. A run without `--classify` warns if files were sorted by older rules.

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
  tds_catalog.xlsx          Excel catalogue: Summary + one sheet per category (Vendor | Document (click to open) |
                            Description | Type | Test methods | Pages | Language | Source URL) + Translations + Rejected
  logs\tds_crawler_YYYYMMDD.log
```

Category folders come from the workbook sheet names: `1_Foil_Materials`, `2_Nonwoven_Fabrics`, `3_Specialty_Paper`, `4_Polymer_Films`, `5_Web_Converting`, `6_Equipment_Machinery`, `7_Process_Technology`. A company listed on more than one sheet (Dow, ExxonMobil, Oerlikon, Brückner) goes in the first sheet's folder, and `tds_manifest.csv` lists all of its categories. To choose a different folder for a company, fill in its `folder` column in `vendor_seeds.csv`. Files downloaded in an older layout are moved automatically on the next `--download`, or by running `--reorganize` on its own.

## Excel catalogue

`--export` (and `--all`) also writes `tds_catalog.xlsx` in the corpus root. It has these sheets:

- **Summary:** one row per material category, with counts of suppliers, documents, translations and rejected files. Click a category name to jump to its sheet.
- **One sheet per category** (`1 Foil Materials`, `2 Nonwoven Fabrics`, `4 Polymer Films`, ... `8 Multi Category`): one row per kept PDF, with Vendor, Document, Description, Type, Test methods, Pages, Language and Source URL. The file name is a link, so clicking it opens the PDF. The description is built from the document's opening text, for example *Technical data sheet: DuPont Pyralux AG - All-Polyimide Double-Sided Copper-Clad Laminate*.
- **Translations:** every non-English copy, with its category and the English version's file name, or `NOT FOUND`.
- **Rejected:** junk files, with their category and the reason.

Keep the workbook in the corpus root so the links work. If it's open in Excel during a run, a time-stamped copy is written instead.

## Foreign-language documents (English first)

The main supplier folder holds English documents only. For every non-English document, the crawler tries to find the English version first:

1. **During the crawl.** A translated copy is recognised by a language code at the end of its filename (`- CN`, `_JP`, `(KR)`, `-DE`...), a language word (`_Chinese`), or a locale folder in its URL (`/zh-cn/`, `/ja/`). The English version is downloaded first.
2. **Guessed URLs.** If the English version wasn't linked, likely English URLs are tried: the name without the code, `-EN`, `_EN`, `_english`, or the `/en/` folder instead of `/ja/`.
3. **After `--classify`.** Documents found to be non-English from their text, even without a code in the name (for example `C1990_data.pdf` in Japanese), get the same URL guesses. Every non-English file is then matched against English files on disk by name.

Where the files go:
- **English version found:** the English file goes in the main folder and the non-English copy in `_translations\`.
- **No English version found:** the non-English copy still goes in `_translations\`, and the crawler moves on.

A file named like a translation counts as non-English unless its text is purely English. The language code can appear anywhere in the name: `X - CN`, `X_CN_200706`, `X-CHIN-0521-web`, `X D_DE_2025`, `X D_rd` vs `X E_rd`. English twins are matched with the code removed wherever it sits, so `Extruder D_DE_2025_Web.pdf` pairs with `Extruder D_EN_2025_Web.pdf`. Bilingual sheets, such as a Japanese sheet with full English text, stay in the main folder.

In `tds_manifest.csv` and on the catalogue's **Translations** sheet, `english_version` names the English file, or says `NOT FOUND`. The manifest's `language` column is read from the PDF text: `en`, `zh`, `ja`, `ko`, `de`..., or `ja+en` for bilingual sheets. `--keep-translations` turns language sorting off.

## Junk filtering: allowlist (the content has to prove it's technical)

Every downloaded file is read (up to the first 6 pages) and must show **real technical content**, or it goes to `_rejected\`. Nothing is kept just because it didn't look like junk, and technical-sounding words on their own aren't enough.

**1. Rejected outright, whatever else is in the file:**
- **Commercial or supplier paperwork** (English, German, Chinese, Japanese): invoicing, credit notes, order confirmations, purchase orders, purchasing or sales terms, delivery guidelines and addresses, Incoterms, goods receipt, supplier inquiry forms and portals, dangerous-goods shipping forms (ADR/IATA/IMDG), freight and cargo registrations, commodity-group lists, human-rights and supply-chain due-diligence statements.
- **Corporate material:** annual, integrated or sustainability reports, management plans, balance sheets, press releases, maps.
- **Non-technical file names:** logos, CVs, business principles, codes of conduct, policies, privacy, AGM or board papers, org charts, invitations, news, ESG, ethics, careers, PFAS, conflict minerals or GHG letters, declarations, disclosures, testimonials, success stories or case studies, invoicing, purchasing, delivery, supplier, cargo or logistics. These are also skipped at crawl time, so they're never downloaded.

**2. Kept only with substance.** A file counts as technical when it has at least one of these:
- a test method (ASTM, ISO, TAPPI, NWSP, IPC, DIN, JIS);
- three or more property names (tensile, elongation, basis weight, peel, dielectric… in English, Chinese, Japanese, Korean, German or French);
- two property names plus units or a data-sheet heading;
- machine specifications (m/min, kg/h, kW, rpm, L/D, working width, throughput…);
- equipment terms (extruder, flat die, spinneret, winder, slitter, coater…) together with units, specs or properties.

It also needs a total evidence score of 3 or more. Stray units or single words (a "%", a "mm", "temperature") don't count, and property words only match as whole words, so "temper" can't match inside "temperature".

**Equipment makers are judged by their own standard.** For suppliers on the Equipment & Machinery or Process Technology sheets, the brochure is the technical document. So any file that names their machines or components (extruders, dies, screws, spinnerets, lines, winders…), gives machine specs, or has a technical or language-coded file name is kept, as long as it has already passed the commercial, corporate and junk-name checks. Their supplier paperwork and service flyers without product content are still rejected.

**3. File names only help when the text can't decide.** A technical name (data sheet, TDS, spec, grade, selection guide, processing guidelines, brochure, catalogue, application note…), taken from the file name or the PDF's built-in title but never from the first line of text, keeps a file only if one of these is true:
- the PDF is image-only or has almost no text;
- the text also shows some technical content;
- the text is in a language the property words don't fully cover.

The non-English copy of a kept English document is always kept, and filed under `_translations\`.

The manifest's `rejected_as` column and the catalogue's **Rejected** sheet give the reason: `commercial`, `corporate`, `not technical (name)`, `no technical content`, `no readable text`, `sds`… For kept files, the **Evidence** column shows why each was kept, for example `12 test method(s); units: %, N/width, mm…; data-sheet heading` or `machine specs: l/d; equipment: extruder, screw`. Nothing is deleted. Use `--keep-rejected` to switch rejection off.

Also handled before checking: crawl-time skipping of investor-relations, sustainability, news, press, careers and finance pages; Unicode normalisation; decoding of shifted-character (`.BSLFU`) and character-code (`/MT78`) text; and treating garbled files as corporate when their same-named siblings are. `--classify` scores any file that hasn't been checked under the current rules, and `--reclassify` re-checks everything.

## Distributors and brand filters

When a supplier's sheets are hosted on a distributor's site (Camden Copper and JX Metals on insulectro.com), the supplier's `pdf_filter` seed column keeps only its own branded files, for example `insulectro.com:camden|denkai|\bMHT\b|\bHP2\b`. Every other PDF on that site goes to the distributor's own row (Insulectro). Files already downloaded under the wrong supplier are re-assigned and moved automatically on the next `--download` or `--reorganize`.

## How it works

**Discovery.** The crawler reads robots.txt and the sitemaps, then follows links from the start URLs in order of a relevance score. Words like datasheet, TDS, technical, download and products in the URL or link text raise the score; news, careers, investor and login lower it. It stays inside each vendor's `domains` and optional `path_prefixes`, and skips non-English locale folders such as `/de/` and `/zh-cn/`. It catches PDFs in `<a>`, iframe and embed tags, data-attributes, and URLs embedded in script/JSON, plus download endpoints that return `application/pdf` from a URL that doesn't end in .pdf.

**Download.** Each file is streamed and kept only if it starts with a `%PDF` header, which rejects HTML error pages saved as .pdf. Files larger than `--max-pdf-mb` are skipped, and a PDF reachable from several URLs is stored once. The filename comes from the Content-Disposition header when the server provides one.

**Link types.** Each link is classified from its URL and link text as `tds`, `sds`, `brochure`, `certificate`, `corporate` or `other`. By default the crawler downloads `tds,brochure,other`; change that with `--doc-types tds` or `--doc-types all`. The `other` type is kept because many data sheets are named only by product code, such as `MHT.pdf`.

**Classification.** After download, pypdf reads the text of pages 1–2 and records the document type from the text, the title, the page count, and the test methods cited: ASTM D882, ISO 527, TAPPI T410, NWSP 100.1, IPC-TM-650 2.4.8, DIN, JIS. The `method_families` column lets you separate film, paper and nonwoven data before comparing values. A file whose name says "data sheet", or whose first pages cite two or more test methods, is classed as a data sheet unless it is a safety data sheet.

**Resumable.** Ctrl-C saves state, and running the same command again continues from there. A vendor that hits `--max-pages` keeps its queue, so another `--discover` run goes deeper.

**Files you move or delete yourself.** If you move a PDF by hand, the crawler finds it again under `pdfs\` and records the new location. If you delete one, it's dropped from the database once and not downloaded again. If a large share of files seem missing, for example because the drive isn't connected, nothing is changed.

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
