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

The first log line of every run shows the build, the rules version and the full path of the script that ran, for example `tds_crawler_v1 build 2026-10-25 r22 (rules v12) - C:\...\tds_crawler_v1.py`. Browsers often save a re-downloaded file as `tds_crawler_v1 (1).py`, so check that line after replacing the script. Whenever the sorting rules change, `--classify` (or `--all`) re-checks every file automatically. A run without `--classify` warns if files were sorted by older rules.

## Install

```
pip install -r requirements.txt
python -m pip install cryptography    # needed to read encrypted PDFs (many Sika sheets)
# optional, for JavaScript-heavy sites (Dow, ExxonMobil, SABIC, 3M, Henkel...):
python -m pip install playwright
```

That is all `--js` needs: the script uses Microsoft Edge or Google Chrome from your PC. `python -m playwright install chromium` is optional, and only needed if neither browser is installed. Use `python -m ...` rather than plain `pip` or `playwright`, so the package goes into the same Python that runs the script. A run with `--js` logs `JavaScript rendering: using Microsoft Edge` (or similar) when it works. If it can't start a browser, it says why and carries on without JavaScript.

## Long runs: is it frozen?

Every 5 minutes the log prints a "Still working" block naming what each worker is doing and for how long (a page, a PDF, a file being classified). A quiet stretch with those lines is a slow site, not a freeze; sites that ask for long pauses in robots.txt (Barmag: 60 s) are the usual cause. A single PDF download is abandoned after 5 minutes and retried on the next run, and guessed English URLs are skipped on such slow sites.

Ctrl-C works at once: the first press finishes the current requests and saves state (seconds, at most about a minute); a second press quits immediately. Either way the database stays consistent, and re-running the same command continues where it stopped.

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

Every vendor in the workbook gets its folder under `pdfs\<Category>\` on every run, even before anything is downloaded, so an empty folder means that vendor has no files yet (see `--status` for why).

```
D:\tds_corpus\
  input\                   workbook + vendor_seeds.csv
  pdfs\<Category>\<Vendor>\*.pdf   e.g. pdfs\4_Polymer_Films\Toray_Industries\  (SHA-256 de-duplicated)
  pdfs\<Category>\<Vendor>\_translations\   foreign-language copies whose English version is in the folder above
  pdfs\<Category>\<Vendor>\_rejected\       files that could not be judged, or off-topic files when --keep-off-topic is used
  tds_crawler.db            SQLite (WAL): pages, pdf_links, files, vendors
  tds_manifest.csv          one row per file: vendor, categories, title, doc type, test methods, source URL...
  tds_catalog.xlsx          Excel catalogue: Summary + one sheet per category (Vendor | Document (click to open) |
                            Description | Type | Test methods | Pages | Language | Source URL) + Translations + Rejected
  logs\tds_crawler_YYYYMMDD.log
  logs\deleted_off_topic_YYYYMMDD.csv   every off-topic file deleted, with its source URL
```

Category folders come from the workbook sheet names: `1_Foil_Materials`, `2_Nonwoven_Fabrics`, `3_Specialty_Paper`, `4_Polymer_Films`, `5_Web_Converting`, `6_Equipment_Machinery`, `7_Process_Technology`. A company listed on more than one sheet (Dow, ExxonMobil, Oerlikon, Brückner) goes in the first sheet's folder, and `tds_manifest.csv` lists all of its categories. To choose a different folder for a company, fill in its `folder` column in `vendor_seeds.csv`. Files downloaded in an older layout are moved automatically on the next `--download`, or by running `--reorganize` on its own.

## Excel catalogue

`--export` (and `--all`) also writes `tds_catalog.xlsx` in the corpus root. It has these sheets:

- **Summary:** one row per material category, with counts of suppliers, documents, translations and rejected files. Click a category name to jump to its sheet.
- **One sheet per category** (`1 Foil Materials`, `2 Nonwoven Fabrics`, `4 Polymer Films`, ... `8 Multi Category`): one row per kept PDF, with Vendor, Document, Description, Type, Test methods, Pages, Language and Source URL. The file name is a link, so clicking it opens the PDF. The description is built from the document's opening text, for example *Technical data sheet: DuPont Pyralux AG - All-Polyimide Double-Sided Copper-Clad Laminate*.
- **Translations:** every non-English copy, with its category and the English version's file name, or `NOT FOUND`.
- **Rejected:** files kept aside (unreadable, or off-topic with `--keep-off-topic`), with the reason.

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

Language is read from the text with stopword lists for English, German, French, Spanish, Italian, Portuguese, Dutch, Swedish, Polish, Turkish, Finnish, Romanian, Hungarian and Malay, plus script detection for Chinese, Japanese, Korean and Cyrillic.

In `tds_manifest.csv` and on the catalogue's **Translations** sheet, `english_version` names the English file, or says `NOT FOUND`. The manifest's `language` column is read from the PDF text: `en`, `zh`, `ja`, `ko`, `de`..., or `ja+en` for bilingual sheets. `--keep-translations` turns language sorting off.

## What is kept: the topic decides

The corpus covers the topics in the *Patent Search List for Corpus Additions*: **webs** (polymer films, foil, paper, nonwovens, fabric), the **web processes** that make them (casting, extrusion, blown and cast film, orientation, calendering, papermaking, spunbond, meltblown, needle punching, foil rolling...), **web converting** (coating, printing, adhesives and PSA, hot-melt, release and barrier coatings, laminating, surface treatment, metallising, curing, embossing, die cutting, slitting, rewinding, sealing, labels, tapes, liners...) and **web handling** equipment (unwinds, winders, guides, tension control, accumulators, splicers, rollers, static control, inspection...).

Every downloaded file is read (first 6 pages). **Anything that relates to these topics is kept**: data sheets, safety data sheets (SDS), brochures, bulletins, processing guides, selection guides, product certificates. A file does not have to be a data sheet. Non-English documents stay under `_translations\`.

**Off-topic files are deleted** and their links are marked `deleted`, so they are never downloaded again. Each one is listed in `logs\deleted_off_topic_YYYYMMDD.csv` with its source URL. Use `--keep-off-topic` to move them to `_rejected\` instead. Off-topic means:

- **Company paperwork, for every vendor:** integrated, annual and sustainability reports, management plans, balance sheets, proxy statements, codes of conduct, business principles, policies, whistleblower and ethics notices, conflict-minerals, GHG and carbon disclosures, supplier forms and portals, purchasing and sales terms, invoicing, order confirmations, claim guides, plant certificates (ISO 9001 / 14001 / 45001, IATF 16949, AS9100), chain-of-custody certificates, registration notices, licence agreements, CVs. Sales terms only count when the file *is* the terms (file name or first line), never a "see our terms and conditions" footer on a product sheet.
- **Company paperwork by category (r14):** stock-exchange and statutory filings, purchasing / supplier paperwork (purchase conditions, supplier training and portals, Coupa / Ariba guides, contractor rules), site rules / safety / EHS notices, HR documents, compliance and position papers, plant or chain-of-custody certificates (FSC, PEFC, FSSC 22000, ISCC, BRCGS, halal, HACCP, GMP...), and (r21) news / HR / corporate leftovers: press releases, job ads, obituaries, open-source license texts, due-diligence and anti-bribery documents, supplier-portal FAQs. The phrase lists are in `PAPERWORK_RULES` in the script and cover the main European languages. Filings and purchasing paperwork are matched on the name and opening text; the other categories only on the name / first lines and only when the document has few web terms, so a product sheet mentioning "ISO 9001" or a safety note stays. Product certificates (EU type-examination, declarations of performance / conformity, BBA product approvals) are kept.
- **Sika: industrial only.** Kept: SikaForce, SikaPower, SikaMelt, SikaFast, SikaTack, SikaBiresin / Advanced Resins, SikaSense, SikaDamp, SikaBlock, Sika Aktivator, Sikaflex 2xx / 5xx / WS / P2G, Sika Primer 2xx, and documents about automotive, marine, appliance, composite, laminating, hot-melt or pressure-sensitive applications. Off-topic: construction product lines (Sikafloor, Sikacrete, SikaGrout, SikaTile, Sarnafil, Sikaplan, SikaProof, Sikalastic, Sikagard, SikaTop, Sikadur, Parex, LaHabra, Merkrete, Schönox, Rmax...) and anything about concrete, mortar, flooring, roofing, waterproofing, tile, stucco or facades. Sika's row in `vendor_seeds.csv` also limits the crawl to its `/en/industry` pages.
- **Per-vendor scope** (`scope_keep` / `scope_exclude` columns in `vendor_seeds.csv`) for companies whose sites mix on- and off-topic product lines. Both are case-insensitive regexes, checked against the PDF's URL, the page it was found on, its file name and its title. An exclude match on the URL or found-on page wins (a whole site section is out), unless the keep pattern matches the file's own name or title; otherwise a keep match on the name or title keeps the file, an exclude match there removes it, and a keep match on the URL alone keeps it. Set today:
  - **Johns Manville:** nonwovens, fiberglass (chopped strands, rovings, glass mats, microfiber), filtration and OEM / aerospace insulation kept; commercial roofing and building insulation off-topic.
  - **DuPont:** Tyvek and Tychem (flash-spun nonwovens), Cyrel flexo plates and platemaking, Tedlar, Kapton and other films kept; water treatment (FilmTec, AmberLite, AmberLyst...), Liveo pharma, Styrofoam and foam sealants off-topic.
  - **BASF:** polyurethane systems, coatings, adhesives, film resins (ecovio, ecoflex, Ultramid, Ultrason) and engineering-plastics / injection-molded part documents kept; gas treating and LNG (Durasorb, OASE, Claus/SRU), AdBlue, oilfield, animal nutrition, pharma intermediates, Cavipor / Enertite / Elastospray insulation, Elastopan / Infinergy footwear and Basotect off-topic.
  - **3M:** respirator leaflets, electronic connectors and camera cables off-topic.
  - **International Paper:** third-party industry reports (glass / steel / aluminium recycling, food surveys, legislation) off-topic.
  - **Kruger:** newsletters, energy and corporate reports off-topic.
  - KraussMaffei, Akzo Nobel and Schlenk are kept whole (injection molding, aerospace coatings and pigments included).
  Add the same columns to any other broad vendor; the next `--classify` applies them to new files (run `--reclassify` to re-check existing ones).
- **Avery Dennison medical:** sprays, creams, swabs and lotions (not webs). Dressings, tapes, films and labels stay.
- **ExxonMobil:** every product document counts as on-topic (`TOPIC_KEEP_ALL_VENDORS` near the top of the script).
- All other vendors in the workbook supply webs, web materials or web machinery, so their product documents are kept even when the wording is terse (`EM-370Z LAM`) or not in English.

Obvious cases are skipped before download, from the file name and link text alone (company paperwork, Sika construction lines, Avery sprays/creams); the database marks those links `off_topic` so they aren't tried again.

A PDF that can't be opened (for example an encrypted file when `cryptography` isn't installed) is never deleted: its topic is unknown, so it falls back to the older technical-evidence rules (test methods, property names with units, data-sheet headings, machine specs) and may go to `_rejected\` with `no technical content` or `no readable text`. Install `cryptography` and run `--classify` again to judge it properly.

Also handled before checking: crawl-time skipping of investor-relations, sustainability, news, press, careers and finance pages; Unicode normalisation; decoding of shifted-character (`.BSLFU`) and character-code (`/MT78`) text. `--classify` judges any file that hasn't been checked under the current rules, `--reclassify` re-checks everything, and `--keep-rejected` turns all sorting off.

## Distributors and brand filters

When a supplier's sheets are hosted on a distributor's site (Camden Copper and JX Metals on insulectro.com), the supplier's `pdf_filter` seed column keeps only its own branded files, for example `insulectro.com:camden|denkai|\bMHT\b|\bHP2\b`. Every other PDF on that site goes to the distributor's own row (Insulectro). Files already downloaded under the wrong supplier are re-assigned and moved automatically on the next `--download` or `--reorganize`.

## How it works

**Discovery.** The crawler reads robots.txt and the sitemaps, then follows links from the start URLs in order of a relevance score. Words like datasheet, TDS, technical, download and products in the URL or link text raise the score; news, careers, investor and login lower it. It stays inside each vendor's `domains` and optional `path_prefixes`, and skips non-English locale folders such as `/de/` and `/zh-cn/`. It catches PDFs in `<a>`, iframe and embed tags, data-attributes, and URLs embedded in script/JSON, plus download endpoints that return `application/pdf` from a URL that doesn't end in .pdf.

**Download.** Each file is streamed and kept only if it starts with a `%PDF` header, which rejects HTML error pages saved as .pdf. Files larger than `--max-pdf-mb` are skipped, and a PDF reachable from several URLs is stored once. The filename comes from the Content-Disposition header when the server provides one.

**Link types.** Each link is classified from its URL and link text as `tds`, `sds`, `brochure`, `certificate`, `corporate` or `other`. By default the crawler downloads `tds,sds,brochure,certificate,other` (everything except corporate links); change that with `--doc-types tds` or `--doc-types all`. The `other` type is kept because many data sheets are named only by product code, such as `MHT.pdf`.

**Classification.** After download, pypdf reads the text of pages 1–2 and records the document type from the text, the title, the page count, and the test methods cited: ASTM D882, ISO 527, TAPPI T410, NWSP 100.1, IPC-TM-650 2.4.8, DIN, JIS. The `method_families` column lets you separate film, paper and nonwoven data before comparing values. A file whose name says "data sheet", or whose first pages cite two or more test methods, is classed as a data sheet unless it is a safety data sheet.

**Resumable.** Ctrl-C saves state, and running the same command again continues from there. A vendor that hits `--max-pages` keeps its queue, so another `--discover` run goes deeper.

**Files you move or delete yourself.** If you move a PDF by hand, the crawler finds it again under `pdfs\` and records the new location. If you delete one, it's dropped from the database once and not downloaded again. If a large share of files seem missing, for example because the drive isn't connected, nothing is changed.

**Revisions.** `--download --refresh` re-checks each downloaded PDF with ETag and If-Modified-Since. A changed sheet is saved as a new file and the old one is kept.

## Document APIs (Qnity, WordPress, Dow)

Some document libraries are JavaScript apps that load their file list from an API, and the files sit on another host with no `.pdf` in the URL, so crawling the pages finds nothing. The crawler reads these APIs directly at the start of `--discover`:

- **Qnity Electronics** (former DuPont Electronics): the Resource Center API, for technical data sheets, product data and information sheets, technical information, brochures, overviews and application notes. Kapton, Pyralux, Oasis, Temprion, Riston and other film, laminate, flex-circuit and thermal-tape documents are kept; Kalrez seals, EKC and other semiconductor chemicals and display materials are left out (`QNITY_KEEP_TAG_RE` / `QNITY_DROP_TAG_RE` in the script). Qnity's SDSs are in a separate login-only system and aren't collected.

- **WordPress sites** (Mylar, Arclin, Amut, Mitsubishi Polyester Film, Gneuss, Trexel, Meaf...): the public media-library API (`/wp-json/wp/v2/media`) lists every PDF the site has uploaded, including ones only reachable through filters or accordions. It is read for every vendor at the start of `--discover` (sites that aren't WordPress simply answer 404); `--no-wp-media` turns it off. The usual topic, paperwork and vendor-scope rules then decide what is kept.
- **Dow:** `dow.com/documents/...pdf` links open an HTML viewer; the crawler requests them with `?iframe=true`, which returns the PDF itself.

Product detail pages (`/p/d/...`, `/products/<name>`) are visited before more listing and filter pages, because that is where data-sheet links are (3M's product pages each link their Tech Data Sheet).

robots.txt rules are read the way sites and Google intend (the longest matching rule wins), so `Disallow: /` plus `Allow: /api/public/` allows `/api/public/...`.

## Search seeding (optional, best for 3M / Dow / ExxonMobil)

Set `SERPAPI_KEY`, or `GOOGLE_CSE_KEY` + `GOOGLE_CSE_CX`, then add `--search-seed`. For each domain the crawler queries `site:<domain> filetype:pdf ("technical data sheet" OR datasheet ...)` and adds the hits directly.

## Editing vendor_seeds.csv

`max_pages` sets a page cap for one vendor that overrides `--max-pages` (set today: 3M 3000, because every product page has to be visited; Oerlikon/Barmag 20, because its robots.txt asks for 60 s between requests). Leave it empty to use `--max-pages`.

| Column | Meaning |
|---|---|
| status | `ok` crawl · `verify` crawl, but domain or relevance unconfirmed · `skip` don't crawl |
| domains | `;`-separated; subdomains are included automatically |
| start_urls | `;`-separated entry points |
| path_prefixes | optional, e.g. `/fibers/`, to stay inside one section of a huge site |
| render_js | `1` = render with Playwright when `--js` is given |
| folder | optional, overrides the category folder, e.g. `5_Web_Converting` |
| pdf_filter | optional, `[domain:]regex`. On a shared site, keep only PDFs whose filename or link text matches |
| scope_keep | optional regex: documents of this vendor that are on-topic (see *What is kept*) |
| scope_exclude | optional regex: this vendor's off-topic product lines or site sections, e.g. a site section such as `/commercial-roofing/` or a product word such as `roof` |

Translated duplicates are skipped and junk such as ethics/hotline pages, licence agreements, testimonials and success stories is filtered out at link level.

When a vendor appears in the workbook but has no seed row, the script adds one automatically with `status=skip`.

## Politeness defaults

The crawler waits 2 s between requests to each vendor, or longer if robots.txt sets a Crawl-delay. It processes 4 vendors in parallel, backs off on 429 and 503 responses, follows robots.txt rules, and sends an identifying User-Agent; add your contact with `--contact`. Site terms vary, and some vendors (UL Prospector-style portals, 3M) forbid automated download, so check before running high volumes.
