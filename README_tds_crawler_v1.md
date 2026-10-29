# Supplier Data-Sheet Crawler (`tds_crawler_v1.py`)

Visits the websites of the suppliers listed in the vendor workbook, finds the PDFs they publish (technical data sheets, safety data sheets, brochures, processing guides...), downloads them, and sorts them by material category, supplier, document type and language. The result is a browsable library of what suppliers say about their films, foils, nonwovens, papers, converting materials and machines.

| | |
|---|---|
| **Input** | `D:\tds_corpus\input\Vendors_With_Technical_Sheets_By_Material_Category.xlsx` (the supplier list) and `D:\tds_corpus\input\vendor_seeds.csv` (how to crawl each one) |
| **Main output** | `D:\tds_corpus\pdfs\<Category>\<Vendor>\` and the clickable catalog `D:\tds_corpus\tds_catalog.xlsx` |
| **Run time** | Days for every supplier; run it in batches with `--vendors` |
| **Depends on** | Nothing else in the project. The [overview report](README_overview_report_v1.md) reads its results. |

---

## Why it works this way

- **Suppliers, not search engines.** Data sheets live on each supplier's own site, in very different layouts, so the crawler follows links from each supplier's start pages and also reads document libraries' APIs where a site loads its files with JavaScript.
- **Keep everything on-topic, not just data sheets.** Safety data sheets, brochures, selector guides and processing guides are kept because they describe the materials. Company paperwork (annual reports, codes of conduct, purchase terms, plant certificates) is removed.
- **English first.** Where a supplier publishes the same sheet in several languages, the English one goes in the main folder and the others are set aside in `_translations`.
- **Nothing is lost silently.** Deleted off-topic files are logged with their source address; unreadable files go to `_rejected`; the same file reached from several links is stored once.
- **Polite.** It obeys robots.txt, waits between requests, backs off when a site is busy, and identifies itself.

## Quick start

```
python -m pip install -r requirements.txt                     # includes cryptography for encrypted PDFs
python tds_crawler_v1.py --list-vendors                       # suppliers, categories and seed status
python tds_crawler_v1.py --discover --vendors "rogers,camden,novelis" --max-pages 60 --contact you@example.com
python tds_crawler_v1.py --status                              # links found per supplier and type (nothing downloaded yet)
python tds_crawler_v1.py --download --classify --export --vendors "rogers,camden,novelis"
```

Open `D:\tds_corpus\tds_catalog.xlsx` and check the documents look right. Then crawl everyone:

```
python tds_crawler_v1.py --all --contact you@example.com       # discover + download + classify + export
python tds_crawler_v1.py --all --js --contact you@example.com  # also render JavaScript-heavy sites (render_js=1)
```

`--discover` on its own is a dry run: it lists every PDF link and how it was classified without downloading anything. Ctrl-C is safe at any time (see [Long runs](#long-runs-is-it-frozen)); the same command continues where it stopped.

---

## Before you run

### Files

| File | Where | Purpose |
|---|---|---|
| `tds_crawler_v1.py` | this repository | The crawler |
| `Vendors_With_Technical_Sheets_By_Material_Category.xlsx` | `D:\tds_corpus\input\` | The supplier list: one sheet per material category, plus a reference guide |
| `vendor_seeds.csv` | `D:\tds_corpus\input\` | One row per supplier: where and how to crawl. **Edit this file, not the script**, to change a supplier ([columns](#editing-vendor_seedscsv)) |
| `requirements.txt` | this repository | Packages to install |

The script looks for the workbook and seed file at `DEFAULT_WORKBOOK` and `DEFAULT_SEEDS` near the top (both in `D:\tds_corpus\input\`), and writes to `DEFAULT_ROOT` (`D:\tds_corpus`). Override with `--workbook`, `--seeds` and `--root`. A supplier in the workbook without a seed row gets one added automatically with `status=skip` until you fill in its website.

### Install

```
python -m pip install -r requirements.txt
python -m pip install playwright        # optional: only for --js
```

- **`cryptography`** (in requirements.txt) lets the crawler read encrypted PDFs, which many suppliers publish (many Sika sheets). Without it they can't be read and go to `_rejected`.
- **`--js`** uses the Microsoft Edge or Google Chrome already on the PC; `python -m playwright install chromium` is only needed if neither is installed. A working run logs `JavaScript rendering: using Microsoft Edge`; if no browser starts, it says why and carries on without JavaScript.
- Use `python -m pip ...` rather than plain `pip`, so packages go into the Python that runs the script.

### Which version ran?

The first log line of every run shows the build, the rules version and the path of the script, e.g. `tds_crawler_v1 build 2026-10-25 r22 (rules v12) - C:\...\tds_crawler_v1.py`. Browsers often save a re-downloaded file as `tds_crawler_v1 (1).py`, so check that line after replacing the script. When the sorting rules change (the rules version goes up), `--classify` re-checks every file automatically; a run without `--classify` warns that files were sorted by older rules.

---

## What it does

| Stage | Command | What happens |
|---|---|---|
| Discover | `--discover` | Reads each supplier's robots.txt and sitemaps, reads document APIs where available, then follows links from the start pages (most promising first) and records every PDF link with a type guessed from its address and link text. |
| Download | `--download` | Downloads the PDF links of the wanted types (by default everything except corporate documents). Keeps only real PDFs, skips files over 50 MB, and stores identical files once (SHA-256). |
| Classify | `--classify` | Reads the first pages of each new file: document type, title, page count, language, test methods cited. Deletes off-topic files (logged), moves foreign copies to `_translations` and unreadable files to `_rejected`. |
| Export | `--export` | Writes `tds_manifest.csv` and `tds_catalog.xlsx`. |
| Everything | `--all` | Discover → download → classify → export. |

Current corpus: 65 suppliers crawled (75,802 pages, 97,249 PDF links), 12,905 documents kept (9,062 technical data sheets), 1,950 foreign-language copies, 132 rejected; 13.8 GB.

---

## What you get

```
D:\tds_corpus\
  input\                                     workbook + vendor_seeds.csv
  pdfs\<Category>\<Vendor>\*.pdf             kept documents, e.g. pdfs\4_Polymer_Films\Toray_Industries\
  pdfs\<Category>\<Vendor>\_translations\    foreign-language copies (English version in the folder above, if found)
  pdfs\<Category>\<Vendor>\_rejected\        files that couldn't be judged (or off-topic, with --keep-off-topic)
  tds_catalog.xlsx                           clickable Excel catalog
  tds_manifest.csv                           one row per file, every detail recorded
  tds_crawler.db                             crawl state (pages, PDF links, files, suppliers)
  logs\tds_crawler_YYYYMMDD.log              daily run log
  logs\deleted_off_topic_YYYYMMDD.csv        every off-topic file deleted, with its source address
```

Category folders come from the workbook's sheet names (`1_Foil_Materials` ... `7_Process_Technology`; DuPont's seed row sends it to `8_Multi_Category` through the `folder` column). A company on several sheets (Dow, ExxonMobil, Oerlikon, Brückner) goes in its first sheet's folder; the manifest lists all its categories. Set the `folder` seed column to choose a different one. Every supplier gets a folder even before anything is downloaded, so an empty folder means nothing collected yet (`--status` says why). Files in an older layout are moved automatically by `--download` or `--reorganize`.

### The catalog (`tds_catalog.xlsx`)

- **Summary:** per material category, the number of suppliers, documents, translations and rejected files. Click a category to jump to its sheet.
- **One sheet per category:** one row per kept PDF with Vendor, Document (click to open the PDF), Description (built from the document's opening text, e.g. *Technical data sheet: DuPont Pyralux AG - All-Polyimide Double-Sided Copper-Clad Laminate*), Type, Test methods, Evidence, Pages, Language and Source URL.
- **Translations:** every foreign-language copy and its English version's file name, or `NOT FOUND`.
- **Rejected:** files set aside, with the reason.

Keep the catalog in `D:\tds_corpus` so the links work. If it's open in Excel during a run, a time-stamped copy is written instead.

### The manifest (`tds_manifest.csv`)

| Column | Meaning |
|---|---|
| `vendor`, `categories`, `category_folder`, `path` | Supplier, its material categories, and where the file is |
| `bytes`, `pages`, `title`, `language` | File details read from the PDF. Language is `en`, `de`, `zh`, `ja`..., or e.g. `ja+en` for bilingual sheets |
| `is_translated_copy`, `english_version` | `yes` for a foreign copy, with the English file's name or `NOT FOUND` |
| `rejected_as` | Why a file is in `_rejected` (blank if kept) |
| `tech_score`, `evidence` | How technical the document looks and the evidence found (units, property names, test methods, headings) |
| `doc_type_from_text`, `doc_type_from_link` | Type read from the PDF text (`tds`, `sds`, `brochure`, `certificate`, `corporate`, `commercial`, `other`, `unreadable`) and guessed from the link |
| `test_methods`, `method_families` | Standards cited, e.g. `ASTM D882; ISO 527`, and their families. Use the families to separate film, paper and nonwoven data before comparing values. |
| `source_url`, `anchor_text`, `found_on_page` | Where the file came from |
| `server_last_modified`, `downloaded_at` | Dates, for refreshing |
| `urls_with_same_file`, `sha256` | How many links led to this file, and its fingerprint |

---

## Checking a run

| Where | What to look at |
|---|---|
| `--status` | Per supplier: pages crawled, links by type and state, files kept. The first place to look when a supplier has nothing. |
| `logs\tds_crawler_YYYYMMDD.log` | The first line (version). `WARNING`/`ERROR` lines. "Still working" blocks every 5 minutes on long runs. |
| `logs\deleted_off_topic_YYYYMMDD.csv` | Every deleted file (`deleted_at`, `vendor`, `reason`, `path`, `source_url`, `sha256`). Skim it after a classify run to be sure nothing useful was removed. |
| `tds_catalog.xlsx` → Rejected / Translations | Files set aside and why; translations with no English version. |
| `tds_crawler.db` | Full state. Table `pdf_links` has every link and its outcome (below). |

**What happens to PDF links** (the `state` column of `pdf_links`):

| State | Meaning |
|---|---|
| `downloaded` | Saved as a new file |
| `duplicate` | Same file as one already saved from another link |
| `unchanged` | Re-checked with `--refresh`; not changed |
| `pending` | Found but not downloaded (type not wanted, or not reached yet) |
| `off_topic` | Judged off-topic from its address or link text; never downloaded |
| `deleted` | Downloaded, then deleted as off-topic or out of scope (logged) |
| `translation` | A foreign-language copy |
| `guess_failed` | An English-version address the crawler guessed didn't exist (or the site was too slow to try) |
| `error`, `robots`, `not_pdf`, `too_large` | Download failed, blocked by robots.txt, not a PDF, or over the size limit |

### Long runs: is it frozen?

Every 5 minutes the log prints a "Still working" block naming what each worker is doing and for how long (a page, a PDF, a file being classified). A quiet stretch with those lines is a slow site, not a freeze; sites whose robots.txt asks for long pauses (Barmag: 60 s) are the usual cause. A single PDF download is abandoned after 5 minutes and retried next run, and guessed English addresses are skipped on such slow sites.

Ctrl-C works at once: the first press finishes the current requests and saves state (seconds, at most about a minute); a second press quits immediately. Either way the database stays consistent.

---

## Need to know: assumptions and limits

- **Public PDFs only.** Documents behind logins, request forms or search boxes aren't collected (e.g. Qnity's safety data sheets). Some sites block crawlers entirely. In the current run 8 suppliers yielded nothing: Asahi Kasei, IMS Technologies, Sappi, Stora Enso (blocked or erroring), Huntsman, Innovatec (links found, nothing kept), Krahn Chemie, Schuler (no reachable PDF links).
- **Check site terms.** Some suppliers' terms forbid automated download (Prospector-style portals, 3M). Use `--contact` so site owners can reach you, and don't raise the request rate.
- **"On-topic" is decided by rules**, built from the project's topic list and three manual audits (October 2026). They're strict about company paperwork and lenient about product documents; read [What is kept](#what-is-kept-the-topic-decides) before relying on completeness. Use `--keep-off-topic` to move rather than delete.
- **Document type and language are read from the PDF text.** Scanned (image-only) PDFs can't be read and are labelled `unreadable`; encrypted PDFs need `cryptography`.
- **The first 2 pages decide the document type; the first 6 decide the topic.** A data sheet with a long cover section may be typed as a general technical document.
- **Websites change.** Re-running finds new documents; `--download --refresh` re-checks existing ones (a changed sheet is saved as a new file and the old one kept).
- **Your manual changes are respected.** Move a PDF yourself and the crawler finds its new location; delete one and it isn't downloaded again. If a large share of files seem missing (e.g. the drive isn't connected), nothing is changed.

---

## Command reference

| Command / option | Does |
|---|---|
| `--list-vendors` | Show suppliers, categories and seed status |
| `--discover` | Crawl sites and collect PDF links (dry run) |
| `--download` | Download pending links (`--refresh` also re-checks downloaded files for changes) |
| `--classify` | Sort new files (`--reclassify` re-checks all files) |
| `--export` | Write the manifest and catalog |
| `--all` | Discover + download + classify + export |
| `--status` | Print per-supplier counts |
| `--reorganize` | Move existing files into the category/vendor layout |
| `--vendors "a,b"` / `--categories "foil,paper"` / `--only-ok` | Limit which suppliers run (name substring / workbook sheet / seed status `ok` only) |
| `--max-pages N` (400), `--max-depth N` (4) | Crawl limits per supplier per run (a seed's `max_pages` overrides) |
| `--delay S` (2.0), `--workers N` (4) | Seconds between requests per supplier; suppliers crawled in parallel |
| `--doc-types` | Link types to download: `tds,sds,brochure,certificate,other` by default, or `all` |
| `--max-pdf-mb N` (50), `--max-attempts N` (3) | Size limit; attempts before giving up on a PDF |
| `--recrawl` | Redo discovery for suppliers already finished |
| `--js` | Render `render_js=1` suppliers in a browser |
| `--search-seed` | Add search-engine results (needs `SERPAPI_KEY`, or `GOOGLE_CSE_KEY` + `GOOGLE_CSE_CX`) |
| `--no-wp-media`, `--no-sitemaps`, `--no-offsite-pdfs` | Turn off the WordPress media API, sitemaps, or PDFs hosted on other domains |
| `--keep-translations`, `--keep-off-topic`, `--keep-rejected` | Turn off language sorting / delete off-topic (move to `_rejected` instead) / all sorting |
| `--contact EMAIL` | Added to the User-Agent (recommended) |
| `--workbook`, `--seeds`, `--root` | Override input and output locations |

---

# Detailed behaviour

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
