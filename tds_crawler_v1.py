#!/usr/bin/env python3
"""
tds_crawler_v1.py - Technical data sheet (TDS) harvester for vendor websites.

Reads the vendor workbook (Vendors_With_Technical_Sheets_By_Material_Category.xlsx)
plus an editable seed file (vendor_seeds.csv) that maps each vendor to its
domains / start URLs, then:

  1. DISCOVER  - polite, robots-aware crawl of each vendor site (sitemaps +
                 relevance-ranked link following) collecting PDF links.
                 Optional: Playwright rendering for JS sites, search-API seeding.
  2. DOWNLOAD  - streams each PDF, validates the %PDF header, de-duplicates by
                 SHA-256, stores under <root>/pdfs/<Vendor>/.
  3. CLASSIFY  - reads page 1-2 text (pypdf) to tag document type and the test
                 methods used (ASTM / ISO / TAPPI / NWSP / IPC / DIN / JIS).

Everything is tracked in SQLite (<root>/tds_crawler.db, WAL mode) so every
stage is resumable. Python 3.9+.

Typical use:
    python tds_crawler_v1.py --discover --vendors "rogers,novelis" --max-pages 50
    python tds_crawler_v1.py --status
    python tds_crawler_v1.py --download --classify
    python tds_crawler_v1.py --all                 # everything, all vendors
    python tds_crawler_v1.py --export              # tds_manifest.csv
"""
from __future__ import annotations

import argparse
import html as html_lib
import csv
import gzip
import hashlib
import heapq
import itertools
import logging
import os
import re
import sqlite3
import sys
import threading
import time
import unicodedata
from concurrent.futures import ThreadPoolExecutor, as_completed, wait, FIRST_COMPLETED
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from urllib import robotparser
from urllib.parse import quote, unquote, urldefrag, urljoin, urlparse

import warnings

import requests
from bs4 import BeautifulSoup

try:  # silence "XML parsed as HTML" - PDF links are still found
    from bs4 import XMLParsedAsHTMLWarning
    warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)
except ImportError:  # older bs4
    pass

try:
    import openpyxl
except ImportError:  # pragma: no cover
    openpyxl = None

# =============================================================================
# CONFIGURATION (edit here or override on the command line)
# =============================================================================
SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_ROOT = r"D:\tds_corpus" if os.name == "nt" else str(Path.cwd() / "tds_corpus")
DEFAULT_WORKBOOK = Path(r"D:\tds_corpus\input\Vendors_With_Technical_Sheets_By_Material_Category.xlsx")
DEFAULT_SEEDS = Path(r"D:\tds_corpus\input\vendor_seeds.csv")

USER_AGENT_BASE = "TDSCollector/1.0 (technical datasheet research)"
BUILD = "2026-10-25 r22"        # shown at startup - tells you which version of the script actually ran
RULES_VERSION = 12              # bump when classification rules change: files are re-scored automatically
                                # (r14 paperwork rules were applied to the corpus by the 2026-10-24 audit, so
                                #  no bump - they apply to new downloads; bump to re-check every file)
DEFAULT_DELAY_S = 2.0          # min seconds between requests to the same vendor
DEFAULT_MAX_PAGES = 400        # HTML pages fetched per vendor per discovery run
DEFAULT_MAX_DEPTH = 4
SLOW_HOST_DELAY = 10           # seconds: hosts asking for longer pauses get no guessed-URL downloads
DEFAULT_WORKERS = 4            # vendors crawled in parallel (one thread each)
DEFAULT_MAX_PDF_MB = 50
DEFAULT_DOC_TYPES = "tds,sds,brochure,certificate,other"   # corporate skipped; SDSs kept (material safety data)
MIN_LINK_SCORE = -2            # links scoring below this are not followed
MAX_SITEMAP_FILES = 60
MAX_SITEMAP_URLS = 60000
REQUEST_TIMEOUT = (15, 45)
MAX_RETRIES = 2

# First path segment that looks like a locale (de, fr-fr, zh_cn...) is skipped
# unless listed here - avoids crawling the same site in 20 languages.
LOCALE_ALLOW = {"en", "en-us", "en-gb", "en_us", "en_gb", "us", "global", "int",
                "ww", "english", "en-global", "en-int", "gb", "uk", "na"}
LOCALE_RE = re.compile(r"^[a-z]{2}([-_][a-z]{2})?$")
# only real language / country codes count as a locale folder ('/jm/' on jm.com is a brand folder, not a language)
LOCALE_CODES = {"ar", "bg", "ca", "cs", "da", "de", "el", "es", "et", "fa", "fi", "fr", "he", "hi", "hr", "hu", "id",
                "is", "it", "ja", "ko", "lt", "lv", "ms", "nl", "no", "pl", "pt", "ro", "ru", "sk", "sl", "sr", "sv",
                "th", "tr", "uk", "vi", "zh", "cn", "jp", "kr", "tw", "br", "mx", "cz", "se", "dk", "gr", "at", "ch",
                "be", "ar", "cl", "co", "pe", "ve", "ua", "il", "sa", "ae", "eg", "in", "my", "sg", "ph", "au", "nz",
                "ca", "za", "es", "fr", "de", "it", "nl", "pl", "pt", "ru", "us", "gb", "en"}


def is_locale(seg: str) -> bool:
    seg = seg.lower()
    return bool(LOCALE_RE.match(seg)) and seg[:2] in LOCALE_CODES

SKIP_EXT = {
    ".jpg", ".jpeg", ".png", ".gif", ".svg", ".webp", ".bmp", ".ico", ".tif", ".tiff",
    ".mp4", ".mov", ".avi", ".wmv", ".webm", ".mp3", ".wav", ".zip", ".rar", ".7z",
    ".gz", ".tar", ".css", ".js", ".json", ".woff", ".woff2", ".ttf", ".eot",
    ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".exe", ".dmg", ".msi",
    ".rss", ".atom", ".vcf", ".ics", ".step", ".stp", ".dwg", ".dxf", ".igs",
}

# --- link relevance scoring (URL path + anchor text) -------------------------
POS_TERMS = {
    "datasheet": 6, "data-sheet": 6, "data_sheet": 6, "data sheet": 6, "tds": 6,
    "datenblatt": 6, "fiche technique": 5, "technical": 3, "specification": 3,
    "specs": 2, "spec": 2, "download": 3, "documents": 3, "document": 2,
    "literature": 3, "library": 3, "resources": 2, "resource": 2, "products": 2,
    "product": 2, "grade": 2, "grades": 2, "portfolio": 1, "brochure": 2,
    "catalog": 2, "catalogue": 2, "film": 1, "films": 1, "foil": 1, "foils": 1,
    "nonwoven": 1, "nonwovens": 1, "paper": 1, "papers": 1, "media": 1,
    "resin": 1, "resins": 1, "polymer": 1, "polymers": 1, "adhesive": 1,
    "adhesives": 1, "coating": 1, "coatings": 1, "technische": 3, "produkte": 2,
}
NEG_TERMS = {
    "news": -4, "press": -4, "career": -8, "careers": -8, "job": -8, "jobs": -8,
    "investor": -8, "investors": -8, "event": -3, "events": -3, "blog": -3,
    "privacy": -8, "cookie": -8, "cookies": -8, "login": -8, "signin": -8,
    "register": -6, "cart": -8, "checkout": -8, "contact": -2, "imprint": -8,
    "impressum": -8, "legal": -5, "terms": -5, "sitemap": -1, "search": -3,
    "tag": -3, "author": -4, "webinar": -3, "video": -3, "videos": -3,
    "annual-report": -8, "governance": -6, "esg": -3, "people": -3, "story": -2,
    "stories": -2, "mailto": -10, "javascript": -10, "share": -4, "print": -2,
    "ir": -8, "csr": -8, "sustainability": -6, "recruit": -8, "recruitment": -8, "newsrelease": -8,
    "news_summary": -8, "finance": -6, "financial": -6, "shareholder": -8, "shareholders": -8,
    "stock": -4, "integrated": -4, "disclosure": -6, "materiality": -6, "backnumber": -6,
}
# PDFs whose URL - or the page they were found on - sits in one of these sections are corporate
CORPORATE_SEGMENTS = {"ir", "investor", "investors", "investor-relations", "csr", "sustainability", "esg",
                      "news", "newsrelease", "news_summary", "press", "press-releases", "pressroom", "media-center",
                      "recruit", "careers", "career", "jobs", "governance", "finance", "financial", "shareholders",
                      "integrated", "annual-report", "annual-reports", "backnumber", "materiality", "company-profile"}

# --- PDF document-type classification ----------------------------------------
DOC_TYPE_PATTERNS = [
    ("sds", re.compile(r"\b(m?sds|safety[\s_-]*data[\s_-]*sheet|sicherheitsdatenblatt|"
                       r"fiche de donn[ée]es de s[ée]curit[ée]|hoja de seguridad)\b", re.I)),
    ("certificate", re.compile(r"\b(certificate|certification|iso[\s_-]?9001|iso[\s_-]?14001|"
                               r"declaration|conformity|compliance|reach|rohs|fda[\s_-]letter|"
                               r"halal|kosher|statement)\b", re.I)),
    ("corporate", re.compile(r"\b(annual[\s_-]*report|sustainability[\s_-]*report|"
                             r"investor|quarterly|q[1-4][\s_-]?20\d\d|code[\s_-]of[\s_-]conduct|"
                             r"terms[\s_-]*(and|&)[\s_-]*conditions|privacy|general[\s_-]terms|"
                             r"purchase[\s_-]*order|press[\s_-]*release|proxy|ethics|hotline|eula|end[\s_-]*user|"
                             r"licen[cs]e[\s_-]*agreement|testimonial|success[\s_-]*story|"
                             r"corporate[\s_-]*fact[\s_-]*sheet)\b", re.I)),
    ("tds", re.compile(r"(\btds\b|\bpds\b|technical[\s_-]*data|data[\s_-]*sheet|datasheet|"
                       r"product[\s_-]*data|technisches[\s_-]*datenblatt|fiche[\s_-]*technique|"
                       r"spec[\s_-]*sheet|specification|properties)", re.I)),
    ("brochure", re.compile(r"\b(brochure|catalog|catalogue|flyer|leaflet|overview|"
                            r"selector|selection[\s_-]guide|portfolio|line[\s_-]?card)\b", re.I)),
]
ALL_DOC_TYPES = ["tds", "sds", "brochure", "certificate", "corporate", "commercial", "other"]
# Stricter certificate test for PDF *text* (TDS bodies often mention REACH/RoHS/FDA compliance)
TEXT_CERT_RE = re.compile(r"(certificate of (analysis|conformance|conformity|compliance)|declaration of|"
                          r"\bstatement (of|on|regarding)\b|conflict minerals|to whom it may concern|"
                          r"dear (valued )?customer|certification letter)", re.I)

# --- translated copies (English-first) ----------------------------------------
# "X - CN.pdf", "X_JP.pdf", "X (KR).pdf", "X-DE-2.pdf"... (upper-case codes only, to avoid product codes)
LANG_CODES = {"CHIN": "zh", "CN": "zh", "ZH": "zh", "TW": "zh", "CHS": "zh", "CHT": "zh", "JP": "ja", "JA": "ja",
              "KR": "ko", "KO": "ko", "DE": "de", "FR": "fr", "ES": "es", "IT": "it", "PT": "pt",
              "BR": "pt", "RU": "ru", "PL": "pl", "TR": "tr", "CZ": "cs", "NL": "nl", "SE": "sv",
              "HU": "hu"}
LANG_WORDS = {"chinese": "zh", "japanese": "ja", "korean": "ko", "german": "de", "deutsch": "de",
              "french": "fr", "francais": "fr", "spanish": "es", "espanol": "es", "italian": "it",
              "italiano": "it", "portuguese": "pt", "russian": "ru", "polish": "pl", "turkish": "tr"}
_TRAIL = r"(?:[\s_\-.]+(?i:\d{1,8}|web|final|rd|print|lowres|v\d+))*"
TRANSLATION_CODE_RE = re.compile(
    r"^(?P<base>.+?)(?:[\s_\-.]+\(?|\()(?P<code>" + "|".join(sorted(LANG_CODES, key=len, reverse=True) + ["CHIN"]) +
    r")\)?(?:[\s_\-]?\d{1,2})?" + _TRAIL + "$")
TRANSLATION_WORD_RE = re.compile(
    r"^(?P<base>.+?)[\s_\-(]+(?P<word>" + "|".join(LANG_WORDS) + r")\)?$", re.I)
ENGLISH_MARK_RE = re.compile(r"[\s_\-.]+\(?(EN|ENG|[Ee]nglish|en|eng|US|GB|UK)\)?(?:[\s_\-]?\d{1,2})?$")
COLLISION_SUFFIX_RE = re.compile(r"_[0-9a-f]{8}$")
GOOD_STATES = ("downloaded", "duplicate", "unchanged")
# Supplier / purchasing / logistics paperwork - never technical documents
COMMERCIAL_TEXT_RE = re.compile(
    r"(invoic(e|es|ing)|credit notes?|order confirmations?|purchase orders?|purchasing (terms|conditions)|"
    r"general (terms|conditions) of (purchase|sale|delivery)|terms and conditions|delivery (guidelines|address|"
    r"addresses|conditions|instructions|note)|goods receipt|incoterms|supplier (inquiry|portal|code of conduct|"
    r"self[- ]assessment|registration|questionnaire)|dangerous goods|\bADR\b.{0,40}\bIATA|IMDG|cargomarket|"
    r"cargo market|freight|forwarder|unloading|opening hours|map directions|commodity groups?|human rights|"
    r"due diligence|supply chain act|"
    r"rechnungsstellung|gutschrift|auftragsbest[äa]tigung|einkaufsbedingungen|lieferbedingungen|"
    r"anlieferung|lieferanschrift|warengruppe|grundsatzerkl[äa]rung|sorgfaltspflicht|lieferkettengesetz|"
    r"发票|订单确认|采购条款|采购条件|請求書|注文書|購買条件)", re.I)
# data-sheet boilerplate that names sales terms without being sales paperwork, e.g. "The quality of our
# products is covered by the terms of the General Conditions of Sale" - removed before COMMERCIAL_TEXT_RE runs
DISCLAIMER_RE = re.compile(
    r"[^.\n]{0,160}\b(covered by|subject to|governed by|according to|in accordance with|refer to|see)\b"
    r"[^.]{0,80}\b(terms and conditions|(general )?(terms|conditions) of (sale|delivery|purchase))[^.]{0,80}"
    r"|[^.\n]{0,160}\b(unterliegt|unterliegen|gemäß|gemaess|gelten)\b[^.]{0,80}"
    r"(verkaufs|liefer)bedingungen[^.]{0,80}", re.I)
CORPORATE_TEXT_RE = re.compile(
    r"(integrated report|annual report|sustainability report|medium\s*-?\s*term\s+management|"
    r"long\s*-?\s*term\s+vision|management plan|balance sheet|financial results|financial statements|securities report|shareholders|"
    r"corporate governance report|esg data|access map|booth map|floor plan|site map|notice of|"
    r"貸借対照表|損益計算書|中期経営計画|長期ビジョン|統合報告|決算|株主|有価証券報告書|各\s*位|ステークホルダー|"
    r"ニュースリリース|プレスリリース|会社概要|アクセスマップ|ブースマップ|計算書類|附属明細書|"
    r"委託生産|お知らせ|経営計画|事業報告|"
    r"资产负债表|年度报告|可持续发展报告|中期经营计划|重大课题|公司治理|"
    r"재무상태표|지속가능경영보고서|사업보고서)", re.I)

# --- test-method extraction (post-download classification) ------------------
TEST_METHOD_PATTERNS = [
    ("ASTM", re.compile(r"\bASTM[\s\-]*([A-Z])[\s\-]*(\d{1,5})(?:[\-/]\d+)?\b")),
    ("ISO", re.compile(r"\b(?:EN[\s\-]+)?ISO[\s\-]*(\d{3,5})(?:[\-:]\d+)?\b")),
    ("TAPPI", re.compile(r"\bTAPPI[\s\-]*T[\s\-]*(\d{2,4})\b")),
    ("NWSP", re.compile(r"\b(?:NWSP|WSP|ERT)[\s\-]*(\d{1,3}\.\d{1,2})\b")),
    ("IPC", re.compile(r"\bIPC[\s\-]*TM[\s\-]*650(?:[\s\-]+(\d\.\d{1,2}\.\d{1,2}[A-Za-z]?))?")),
    ("DIN", re.compile(r"\bDIN[\s\-]+(?:EN[\s\-]+)?(\d{3,5})\b")),
    ("JIS", re.compile(r"\bJIS[\s\-]*([A-Z])[\s\-]*(\d{4})\b")),
]

log = logging.getLogger("tds")
STOP = threading.Event()
ACTIVITY: dict[int, tuple[float, str]] = {}   # thread id -> (started, what it is doing), for the heartbeat log
MAX_DOWNLOAD_SECONDS = 300     # a single PDF download is abandoned (and retried next run) after this long
HEARTBEAT_SECONDS = 300        # every 5 minutes the log says what each worker is busy with


def doing(what: str) -> None:
    ACTIVITY[threading.get_ident()] = (time.time(), what)


def heartbeat() -> None:
    """Background thread: every few minutes, log what each worker is doing, so a long quiet stretch
    (a slow site, a big PDF, a 60 s crawl-delay) is visibly not a freeze."""
    while not STOP.wait(HEARTBEAT_SECONDS):
        now = time.time()
        busy = sorted(ACTIVITY.values())
        if not busy:
            continue
        log.info("Still working - %d worker(s) busy:", len(busy))
        for started, what in busy:
            age = int(now - started)
            log.info("   %s (%s)", what[:160], f"{age // 60} min {age % 60} s" if age >= 60 else f"{age} s")


# =============================================================================
# Helpers
# =============================================================================
def utcnow() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def norm_name(name: str) -> str:
    """Normalize a vendor name so 'Dow Chemical (Coatings)' == 'Dow Chemical'."""
    s = unicodedata.normalize("NFKD", str(name or "")).encode("ascii", "ignore").decode()
    s = re.sub(r"\(.*?\)", " ", s).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    stop = {"inc", "ltd", "llc", "ag", "se", "sa", "co", "corp", "corporation",
            "company", "group", "limited", "plc", "gmbh", "the"}
    return " ".join(t for t in s.split() if t not in stop).strip()


def slug(name: str, maxlen: int = 60) -> str:
    s = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode()
    s = re.sub(r"\(.*?\)", "", s)
    s = re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_")
    return s[:maxlen] or "vendor"


def safe_filename(name: str, maxlen: int = 150) -> str:
    name = unquote(name).strip().replace("\x00", "")
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name)
    name = re.sub(r"\s+", " ", name).strip(" .")
    if not name.lower().endswith(".pdf"):
        name = (name or "document") + ".pdf"
    stem, ext = name[:-4], name[-4:]
    if len(name) > maxlen:
        stem = stem[: maxlen - 4]
    if stem.upper() in {"CON", "PRN", "AUX", "NUL"} or re.fullmatch(r"(COM|LPT)\d", stem.upper()):
        stem = "_" + stem
    return stem + ext


def canon_url(url: str) -> str:
    url, _ = urldefrag(url.strip())
    p = urlparse(url)
    if p.scheme not in ("http", "https"):
        return ""
    netloc = p.netloc.lower()
    if netloc.endswith(":80") and p.scheme == "http":
        netloc = netloc[:-3]
    if netloc.endswith(":443") and p.scheme == "https":
        netloc = netloc[:-4]
    path = re.sub(r"/{2,}", "/", p.path or "/")
    # drop common tracking params
    query = "&".join(q for q in p.query.split("&")
                     if q and not re.match(r"(utm_[a-z]+|gclid|fbclid|mc_[a-z]+|hsctatracking)=", q, re.I))
    return f"{p.scheme}://{netloc}{path}" + (f"?{query}" if query else "")


def host_of(url: str) -> str:
    h = urlparse(url).hostname or ""
    return h.lower()


def host_allowed(host: str, domains: list[str]) -> bool:
    return any(host == d or host.endswith("." + d) for d in domains)


def looks_like_pdf_url(url: str) -> bool:
    low = url.lower()
    path = urlparse(low).path
    return path.endswith(".pdf") or ".pdf?" in low or "format=pdf" in low or "/pdf/" in path and "." not in path.rsplit("/", 1)[-1]


def ext_of(url: str) -> str:
    last = urlparse(url).path.rsplit("/", 1)[-1]
    return ("." + last.rsplit(".", 1)[-1].lower()) if "." in last else ""


def score_link(url: str, anchor: str = "") -> int:
    p = urlparse(url)
    text = f"{unquote(p.path)} {unquote(p.query)} {anchor}".lower()
    tokens = set(re.findall(r"[a-z0-9]+", text))
    s = 0
    for term, w in itertools.chain(POS_TERMS.items(), NEG_TERMS.items()):
        if (" " in term or "-" in term or "_" in term or len(term) >= 7):
            if term in text:
                s += w
        elif term in tokens:
            s += w
    # product detail pages carry the data-sheet links: visit them before more listing / filter pages
    if PRODUCT_PAGE_RE.search(p.path):
        s += 4
    return s


PRODUCT_PAGE_RE = re.compile(r"/p/dc?/[^/]+|/products?/[^/]+/?$|/product-detail", re.I)


def classify_doc_type(text: str) -> str:
    for dtype, rx in DOC_TYPE_PATTERNS:
        if rx.search(text):
            return dtype
    return "other"


def stem_of(url: str) -> str:
    name = unquote(urlparse(url).path.rstrip("/").rsplit("/", 1)[-1])
    return name[:-4] if name.lower().endswith(".pdf") else name


LANG_TOKEN_RE = re.compile(r"(?:(?<=^)|(?<=[\s_\-.(]))(?:EN|ENG|ENGL|ENGLISH|ENGLISCH|DE|GER|DEU|DEUTSCH|FR|ES|IT|"
                           r"CN|CHN|CHIN|CHINESE|ZH|JP|JA|KR|KO|RU|PL|PT|TR|NL|D|E|GB|US|UK)(?=$|[\s_\-.)])", re.I)


def match_key(stem: str) -> str:
    """Comparable key for a filename stem: case/punctuation-free, language markers removed wherever they
    appear ('Extruder D_EN_2025' == 'Extruder D_DE_2025', 'Duese Platte E_rd' == 'Duese Platte D_rd')."""
    stem = COLLISION_SUFFIX_RE.sub("", stem.strip())
    stem = ENGLISH_MARK_RE.sub("", stem)
    stem = LANG_TOKEN_RE.sub(" ", stem)
    return re.sub(r"[^a-z0-9]+", "", stem.lower())


def translation_info(url: str) -> tuple[str, str, list[str]]:
    """('zh', key_of_english_name, [guessed English URLs]) for a translated copy, else ('', '', [])."""
    p = urlparse(url)
    stem = stem_of(url)
    lang, base = "", ""
    m = TRANSLATION_CODE_RE.match(stem)
    if m:
        lang, base = LANG_CODES[m.group("code")], m.group("base")
    else:
        m = TRANSLATION_WORD_RE.match(stem)
        if m:
            lang, base = LANG_WORDS[m.group("word").lower()], m.group("base")
    segs = p.path.split("/")
    loc_idx = next((i for i, sg in enumerate(segs[:5]) if is_locale(sg)
                    and sg.lower()[:2] != "en" and sg.lower() not in LOCALE_ALLOW), None)
    if not lang and loc_idx is not None:
        lang, base = segs[loc_idx].lower()[:2], stem
    if not lang:
        return "", "", []
    base = base.rstrip(" _-.")
    guesses: list[str] = []
    folder = p.path.rsplit("/", 1)[0]
    if base != stem:
        for name in (base, base + "-EN", base + " - EN"):
            guesses.append(f"{p.scheme}://{p.netloc}{folder}/{quote(name)}.pdf")
    if loc_idx is not None:
        for en in ("en", "en-us"):
            sg = list(segs)
            sg[loc_idx] = en
            guesses.append(f"{p.scheme}://{p.netloc}{'/'.join(sg)}")
    return lang, match_key(stem), list(dict.fromkeys(g for g in guesses if canon_url(g) != canon_url(url)))


ENGLISH_NAME_SUFFIXES = ("_english", "_en", "-en", "_EN", "-EN")
ENGLISH_E_SUFFIX_RE = re.compile(r"(?<=[A-Z0-9])e(?:[-_]\d)?$")   # 'MHSe', 'PCR-WO3N-CSRBe-1'


def _name_url(path: str) -> str:
    return "file://local/" + quote(Path(path).name)


def name_lang(path: str) -> str:
    """Language code implied by a file name ('X - CN.pdf' -> 'zh'), else ''."""
    return translation_info(_name_url(path))[0]


def foreign_key(path: str) -> str:
    """Match key of the English twin we'd expect for a (possibly suffixed) foreign file name."""
    return match_key(Path(path).stem)       # language tokens are removed wherever they appear


def english_url_guesses(url: str) -> list[str]:
    """Likely URLs of the English version of a foreign document: suffix removed / -EN / _english /
    /en/ folder instead of /zh/ ... (works for names with or without a language suffix)."""
    lang, _, guesses = translation_info(url)
    p = urlparse(url)
    stem = stem_of(url)
    base = stem
    m = TRANSLATION_CODE_RE.match(stem) or TRANSLATION_WORD_RE.match(stem)
    if m:
        base = m.group("base").rstrip(" _-.")
    folder = p.path.rsplit("/", 1)[0]
    out = list(guesses)
    for sfx in ENGLISH_NAME_SUFFIXES:
        out.append(f"{p.scheme}://{p.netloc}{folder}/{quote(base + sfx)}.pdf")
    if re.search(r"[A-Z0-9]$", base):
        out.append(f"{p.scheme}://{p.netloc}{folder}/{quote(base + 'e')}.pdf")   # 'MHS.pdf' -> 'MHSe.pdf'
    if base != stem:
        out.append(f"{p.scheme}://{p.netloc}{folder}/{quote(base)}.pdf")
    me = canon_url(url)
    return list(dict.fromkeys(canon_url(g) for g in out if canon_url(g) and canon_url(g) != me))


ENGLISH_STOPWORDS = {"the", "and", "of", "for", "with", "to", "in", "is", "are", "this", "that", "by",
                     "on", "from", "or", "be", "as", "at", "properties", "typical", "value", "test"}
OTHER_STOPWORDS = {
    "de": {"und", "der", "die", "das", "mit", "für", "ist", "von", "den", "eine", "nicht", "werden"},
    "fr": {"et", "les", "des", "pour", "avec", "est", "une", "dans", "sur", "par", "aux", "sont"},
    "es": {"y", "los", "las", "para", "con", "es", "una", "del", "por", "como", "más", "son"},
    "it": {"e", "il", "della", "per", "con", "sono", "una", "delle", "degli", "nel", "che", "alla"},
    "pt": {"e", "os", "as", "para", "com", "uma", "dos", "das", "não", "são", "pelo", "pela"},
    "nl": {"een", "het", "van", "voor", "zijn", "met", "niet", "wordt", "deze", "ook", "bij", "naar"},
    "sv": {"och", "att", "för", "av", "på", "är", "med", "som", "till", "eller", "inte", "vid"},
    "pl": {"jest", "się", "nie", "oraz", "dla", "przez", "lub", "jak", "który", "które", "przy", "jako"},
    "tr": {"ve", "bir", "için", "ile", "bu", "olarak", "veya", "daha", "çok", "gibi", "olan", "değil"},
    "fi": {"ja", "on", "tai", "kanssa", "jos", "että", "ei", "tämä", "joka", "ovat", "kun", "sekä"},
    "ro": {"și", "în", "cu", "pentru", "sau", "este", "nu", "din", "care", "mai", "fi", "sunt"},
    "hu": {"és", "az", "egy", "hogy", "nem", "vagy", "csak", "meg", "ezt", "kell", "ha", "mint"},
    "ms": {"dan", "yang", "untuk", "dengan", "ini", "atau", "pada", "dalam", "tidak", "boleh", "adalah", "jika"},
}


def detect_language(text: str) -> str:
    """Rough language of PDF text: en, zh, ja, ko, ru, de, fr, es, it, pt, nl, sv, pl, tr, fi, ro, hu, ms, '<xx>+en' (bilingual) or ''."""
    if len(text.strip()) < 40:
        return ""
    words = re.findall(r"[A-Za-zÀ-ÿ]+", text.lower())
    en_hits = sum(1 for w in words if w in ENGLISH_STOPWORDS)
    counts = {"ja": len(re.findall(r"[\u3040-\u30ff]", text)), "ko": len(re.findall(r"[\uac00-\ud7af]", text)),
              "zh": len(re.findall(r"[\u4e00-\u9fff]", text)), "ru": len(re.findall(r"[\u0400-\u04ff]", text))}
    top = max(counts.values())
    if counts["ja"] > 20 or (counts["ja"] >= 8 and en_hits < 5):
        script = "ja"
    else:
        # plenty of CJK/Cyrillic, or a little of it with hardly any English words around it
        script = max(counts, key=counts.get) if top > 30 or (top >= 8 and en_hits < 5) else ""
    if script:
        bilingual = en_hits >= 10 or (len(words) >= 40 and en_hits >= 5)
        return f"{script}+en" if bilingual else script
    other = {k: sum(1 for w in words if w in v) for k, v in OTHER_STOPWORDS.items()}
    best = max(other, key=other.get)
    if other[best] > en_hits * 1.5 and other[best] >= 8:
        return f"{best}+en" if en_hits >= 25 else best
    return "en" if en_hits >= 3 or words else ""


def locale_blocked(url: str, vendor_prefix_locales: set[str]) -> bool:
    segs = [s for s in urlparse(url).path.lower().split("/") if s]
    if not segs:
        return False
    first = segs[0]
    return is_locale(first) and first not in LOCALE_ALLOW and first not in vendor_prefix_locales


# =============================================================================
# Vendor model / loading
# =============================================================================
@dataclass
class Vendor:
    name: str
    key: str
    categories: list[str] = field(default_factory=list)
    category_folders: list[str] = field(default_factory=list)   # e.g. "4_Polymer_Films"
    folder_override: str = ""                                   # seeds.csv "folder" column
    pdf_filter_domain: str = ""                                 # seeds.csv "pdf_filter": [domain:]regex
    pdf_filter: object = None                                   # compiled regex or None
    product_lines: list[str] = field(default_factory=list)
    status: str = "skip"
    domains: list[str] = field(default_factory=list)
    start_urls: list[str] = field(default_factory=list)
    path_prefixes: list[str] = field(default_factory=list)
    render_js: bool = False
    notes: str = ""
    scope_keep: object = None                                   # seeds.csv "scope_keep": compiled regex or None
    scope_exclude: object = None                                # seeds.csv "scope_exclude": compiled regex or None
    max_pages: int = 0                                          # seeds.csv "max_pages": per-vendor page cap (0 = --max-pages)

    @property
    def slug(self) -> str:
        return slug(self.name)

    @property
    def category_folder(self) -> str:
        """Top-level folder under pdfs/. First workbook sheet the vendor appears on,
        unless the seeds 'folder' column overrides it."""
        if self.folder_override:
            return slug(self.folder_override)
        return self.category_folders[0] if self.category_folders else "0_Uncategorized"

    def pdf_dir(self, root: Path) -> Path:
        return root / "pdfs" / self.category_folder / self.slug

    def filter_for(self, host: str):
        """Brand filter that applies to PDFs on this host (shared distributor sites), or None."""
        if self.pdf_filter is None:
            return None
        if self.pdf_filter_domain and not host_allowed(host, [self.pdf_filter_domain]):
            return None
        return self.pdf_filter


def load_workbook_vendors(path: Path) -> dict[str, Vendor]:
    if openpyxl is None:
        sys.exit("openpyxl is required: pip install openpyxl")
    if not path.exists():
        sys.exit(f"Workbook not found: {path}  (use --workbook)")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    vendors: dict[str, Vendor] = {}
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        hdr_idx = None
        for i, r in enumerate(rows[:10]):
            cells = [str(c).strip().lower() if c is not None else "" for c in r]
            # header = a row with several filled cells, one of them a short "vendor..." label
            if sum(1 for c in cells if c) >= 3 and any(c.startswith("vendor") and len(c) < 30 for c in cells):
                hdr_idx = i
                break
        if hdr_idx is None:
            continue  # e.g. REFERENCE GUIDE
        hdr = [str(c).strip().lower() if c is not None else "" for c in rows[hdr_idx]]
        col_v = next(i for i, c in enumerate(hdr) if c.startswith("vendor"))
        col_p = next((i for i, c in enumerate(hdr) if "product" in c), None)
        category = re.sub(r"^\s*\d+\s*", "", ws.title).strip().title()
        cat_folder = slug(ws.title.strip().title())          # "4 POLYMER FILMS" -> "4_Polymer_Films"
        for r in rows[hdr_idx + 1:]:
            if r is None or col_v >= len(r) or not r[col_v]:
                continue
            raw = str(r[col_v]).strip()
            key = norm_name(raw)
            if not key:
                continue
            v = vendors.get(key)
            if v is None:
                v = vendors[key] = Vendor(name=raw, key=key)
            if category not in v.categories:
                v.categories.append(category)
                v.category_folders.append(cat_folder)
            if col_p is not None and col_p < len(r) and r[col_p]:
                pl = str(r[col_p]).strip()
                if pl not in v.product_lines:
                    v.product_lines.append(pl)
    log.info("Workbook: %d unique vendors across %d sheets", len(vendors), len(wb.worksheets))
    return vendors


SEED_FIELDS = ["vendor", "status", "domains", "start_urls", "path_prefixes", "render_js", "folder",
               "pdf_filter", "scope_keep", "scope_exclude", "max_pages", "notes"]


def merge_seeds(vendors: dict[str, Vendor], seeds_path: Path) -> None:
    """Attach seed info; append workbook vendors missing from the seed file."""
    seen: set[str] = set()
    rows: list[dict] = []
    if seeds_path.exists():
        with open(seeds_path, newline="", encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
    split = lambda s: [x.strip() for x in (s or "").split(";") if x.strip()]
    for row in rows:
        key = norm_name(row.get("vendor", ""))
        seen.add(key)
        v = vendors.get(key)
        if v is None:
            continue  # seed row for a vendor not in workbook - ignored
        v.status = (row.get("status") or "skip").strip().lower()
        v.domains = [d.lower().lstrip(".") for d in split(row.get("domains"))]
        v.start_urls = split(row.get("start_urls"))
        v.path_prefixes = split(row.get("path_prefixes"))
        v.render_js = str(row.get("render_js", "0")).strip() in ("1", "true", "yes", "y")
        v.notes = row.get("notes", "") or ""
        v.folder_override = (row.get("folder") or "").strip()
        mp = (row.get("max_pages") or "").strip()
        v.max_pages = int(mp) if mp.isdigit() else 0
        pf = (row.get("pdf_filter") or "").strip()
        if pf:
            m = re.match(r"^([a-z0-9.-]+\.[a-z]{2,}):(.+)$", pf, re.I)
            dom, rx = (m.group(1).lower(), m.group(2)) if m else ("", pf)
            try:
                v.pdf_filter, v.pdf_filter_domain = re.compile(rx, re.I), dom
            except re.error as e:
                log.error("Bad pdf_filter regex for %s: %s", v.name, e)
        for col in ("scope_keep", "scope_exclude"):
            rx = (row.get(col) or "").strip()
            if rx:
                try:
                    setattr(v, col, re.compile(rx, re.I))
                except re.error as e:
                    log.error("Bad %s regex for %s: %s", col, v.name, e)
        if v.status != "skip" and not v.domains and v.start_urls:
            v.domains = sorted({host_of(u).removeprefix("www.") for u in v.start_urls})
    missing = [v for k, v in vendors.items() if k not in seen]
    if missing:
        write_header = not seeds_path.exists()
        fieldnames = SEED_FIELDS
        if not write_header:  # follow the existing file's column order
            with open(seeds_path, newline="", encoding="utf-8-sig") as f:
                fieldnames = next(csv.reader(f), SEED_FIELDS) or SEED_FIELDS
        with open(seeds_path, "a", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
            if write_header:
                w.writeheader()
            for v in missing:
                w.writerow({"vendor": v.name, "status": "skip", "domains": "", "start_urls": "",
                            "path_prefixes": "", "render_js": "0", "folder": "",
                            "notes": "Added from workbook - fill domains/start_urls and set status=ok"})
        log.warning("Appended %d new workbook vendor(s) to %s (status=skip until you add URLs): %s",
                    len(missing), seeds_path.name, ", ".join(v.name for v in missing))


# =============================================================================
# Database
# =============================================================================
SCHEMA = """
CREATE TABLE IF NOT EXISTS vendors (
    vendor TEXT PRIMARY KEY, categories TEXT, status TEXT, domains TEXT,
    discovery_done INTEGER DEFAULT 0, last_discovery TEXT
);
CREATE TABLE IF NOT EXISTS pages (
    url TEXT, vendor TEXT, depth INTEGER, score INTEGER,
    state TEXT,            -- queued | done | error | blocked
    http_status INTEGER, fetched_at TEXT, error TEXT,
    PRIMARY KEY (vendor, url)
);
CREATE INDEX IF NOT EXISTS ix_pages_vendor_state ON pages(vendor, state);
CREATE TABLE IF NOT EXISTS pdf_links (
    url TEXT PRIMARY KEY, vendor TEXT, found_on TEXT, anchor TEXT, doc_type TEXT,
    first_seen TEXT,
    state TEXT DEFAULT 'pending', -- pending | downloaded | duplicate | not_pdf | error | too_large | robots | unchanged
    http_status INTEGER, sha256 TEXT, path TEXT, bytes INTEGER, etag TEXT,
    last_modified TEXT, downloaded_at TEXT, attempts INTEGER DEFAULT 0, error TEXT
);
CREATE INDEX IF NOT EXISTS ix_pdf_vendor_state ON pdf_links(vendor, state);
CREATE TABLE IF NOT EXISTS files (
    sha256 TEXT PRIMARY KEY, path TEXT, bytes INTEGER, vendor TEXT, first_url TEXT,
    added_at TEXT, pages INTEGER, title TEXT, text_doc_type TEXT,
    test_methods TEXT, method_families TEXT, classified_at TEXT
);
"""


class DB:
    def __init__(self, path: Path):
        self.path = path
        self.lock = threading.Lock()
        self._local = threading.local()
        with self.conn() as c:
            c.executescript(SCHEMA)
            self._migrate(c)

    @staticmethod
    def _migrate(c: sqlite3.Connection):
        """Add columns introduced after the first release (keeps existing databases working)."""
        new_cols = {"pdf_links": [("lang", "TEXT"), ("english_url", "TEXT"), ("english_of", "TEXT")],
                    "files": [("language", "TEXT"), ("is_translation", "INTEGER DEFAULT 0"),
                              ("english_path", "TEXT")]}
        for table, cols in new_cols.items():
            have = {r[1] for r in c.execute(f"PRAGMA table_info({table})")}
            for name, typ in cols:
                if name not in have:
                    c.execute(f"ALTER TABLE {table} ADD COLUMN {name} {typ}")
        c.execute("CREATE INDEX IF NOT EXISTS ix_pdf_english_of ON pdf_links(english_of)")
        have_files = {r[1] for r in c.execute("PRAGMA table_info(files)")}
        if "rejected" not in have_files:
            c.execute("ALTER TABLE files ADD COLUMN rejected TEXT")
        if "summary" not in have_files:
            c.execute("ALTER TABLE files ADD COLUMN summary TEXT")
        for col, typ in (("tech_score", "INTEGER"), ("evidence", "TEXT"), ("pdf_title", "TEXT"),
                         ("text_chars", "INTEGER"), ("rules_version", "INTEGER"), ("topic", "TEXT")):
            if col not in have_files:
                c.execute(f"ALTER TABLE files ADD COLUMN {col} {typ}")
        # pages used to be keyed by url alone, which mixed up vendors sharing a site
        pk = [r[1] for r in c.execute("PRAGMA table_info(pages)") if r[5]]
        if pk == ["url"]:
            c.executescript("""
                CREATE TABLE pages_new (url TEXT, vendor TEXT, depth INTEGER, score INTEGER, state TEXT,
                    http_status INTEGER, fetched_at TEXT, error TEXT, PRIMARY KEY (vendor, url));
                INSERT OR IGNORE INTO pages_new SELECT url, vendor, depth, score, state, http_status,
                    fetched_at, error FROM pages;
                DROP TABLE pages;
                ALTER TABLE pages_new RENAME TO pages;
                CREATE INDEX IF NOT EXISTS ix_pages_vendor_state ON pages(vendor, state);""")

    def conn(self) -> sqlite3.Connection:
        c = getattr(self._local, "c", None)
        if c is None:
            c = sqlite3.connect(self.path, timeout=60, check_same_thread=False)
            c.execute("PRAGMA journal_mode=WAL")
            c.execute("PRAGMA synchronous=NORMAL")
            c.execute("PRAGMA busy_timeout=60000")
            c.row_factory = sqlite3.Row
            self._local.c = c
        return c

    def execute(self, sql: str, params=()) -> sqlite3.Cursor:
        with self.lock:
            c = self.conn()
            cur = c.execute(sql, params)
            c.commit()
            return cur

    def executemany(self, sql: str, seq) -> None:
        with self.lock:
            c = self.conn()
            c.executemany(sql, seq)
            c.commit()

    def query(self, sql: str, params=()) -> list[sqlite3.Row]:
        return self.conn().execute(sql, params).fetchall()


# =============================================================================
# HTTP / robots
# =============================================================================
def robots_longest_match(lines, url: str) -> bool:
    """Python's robotparser applies the first matching rule, so "Disallow: /" followed by
    "Allow: /api/public/" blocks everything. Sites (and Google) mean the longest match: re-check the
    rules for "User-agent: *" that way. Returns True when the most specific matching rule is Allow."""
    if not lines:
        return False
    p = urlparse(url)
    path = (p.path or "/") + (f"?{p.query}" if p.query else "")
    rules, in_star, prev_ua = [], False, False
    for raw in lines:
        line = raw.split("#", 1)[0].strip()
        if ":" not in line:
            continue
        key, val = (x.strip() for x in line.split(":", 1))
        key = key.lower()
        if key == "user-agent":
            in_star = (val == "*") or (in_star and prev_ua)
            prev_ua = True
            continue
        prev_ua = False
        if in_star and key in ("allow", "disallow") and val:
            rx = re.escape(val).replace(r"\*", ".*")
            if rx.endswith(r"\$"):
                rx = rx[:-2] + "$"
            rules.append((len(val), key == "allow", re.compile(rx)))
    best = None
    for ln, allow, rx in rules:
        if rx.match(path) and (best is None or ln > best[0] or (ln == best[0] and allow)):
            best = (ln, allow)
    return bool(best and best[1])


class Fetcher:
    """One per worker thread: session, robots cache, per-host pacing."""

    def __init__(self, args):
        self.args = args
        self.s = requests.Session()
        ua = USER_AGENT_BASE + (f" contact: {args.contact}" if args.contact else "")
        self.ua = ua
        self.s.headers.update({
            "User-Agent": ua,
            "Accept": "text/html,application/xhtml+xml,application/pdf;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.8",
        })
        self.robots: dict[str, robotparser.RobotFileParser | None] = {}
        self.robot_sitemaps: dict[str, list[str]] = {}
        self.last_hit: dict[str, float] = {}
        self.delay = args.delay

    def _robots_for(self, url: str):
        p = urlparse(url)
        base = f"{p.scheme}://{p.netloc}"
        if base in self.robots:
            return self.robots[base]
        rp = robotparser.RobotFileParser()
        sitemaps: list[str] = []
        try:
            r = self.s.get(base + "/robots.txt", timeout=REQUEST_TIMEOUT)
            if r.status_code == 200 and "html" not in r.headers.get("Content-Type", ""):
                lines = r.text.splitlines()
                rp.parse(lines)
                rp._tds_lines = lines            # kept for longest-match Allow rules (see allowed())
                sitemaps = [l.split(":", 1)[1].strip() for l in lines
                            if l.lower().startswith("sitemap:")]
            elif r.status_code in (401, 403):
                rp.parse(["User-agent: *", "Disallow: /"])
            else:
                rp.parse([])  # no robots -> allow all
        except requests.RequestException:
            rp.parse([])
        self.robots[base] = rp
        self.robot_sitemaps[base] = sitemaps
        return rp

    def allowed(self, url: str) -> bool:
        if self.args.ignore_robots:
            return True
        rp = self._robots_for(url)
        if not rp or rp.can_fetch(self.ua, url):
            return True
        return robots_longest_match(getattr(rp, "_tds_lines", None), url)

    def crawl_delay(self, url: str) -> float:
        rp = self._robots_for(url)
        cd = None
        try:
            cd = rp.crawl_delay(self.ua) if rp else None
        except Exception:
            cd = None
        return max(self.delay, float(cd or 0))

    def _pace(self, url: str):
        host = host_of(url)
        wait = self.crawl_delay(url) - (time.time() - self.last_hit.get(host, 0))
        if wait > 0:
            STOP.wait(min(wait, 60))
        self.last_hit[host] = time.time()

    def get(self, url: str, stream=False, headers=None) -> requests.Response:
        last_exc = None
        for attempt in range(MAX_RETRIES + 1):
            if STOP.is_set():
                raise requests.RequestException("stopped")
            self._pace(url)
            try:
                r = self.s.get(url, timeout=REQUEST_TIMEOUT, stream=stream,
                               headers=headers or {}, allow_redirects=True)
                if r.status_code in (429, 503) and attempt < MAX_RETRIES:
                    ra = r.headers.get("Retry-After", "")
                    wait = int(ra) if ra.isdigit() else 30 * (attempt + 1)
                    r.close()
                    log.info("  %s -> %d, backing off %ss", url, r.status_code, wait)
                    STOP.wait(min(wait, 300))
                    continue
                return r
            except requests.RequestException as e:
                last_exc = e
                STOP.wait(5 * (attempt + 1))
        raise last_exc or requests.RequestException("failed")


# =============================================================================
# Optional: Playwright renderer and search-API seeding
# =============================================================================
class Renderer:
    def __init__(self):
        from playwright.sync_api import sync_playwright  # noqa
        self._pw = sync_playwright().start()
        self.browser, self.name = self._launch()
        self.ctx = self.browser.new_context(user_agent=USER_AGENT_BASE)

    def _launch(self):
        """Playwright's own Chromium if it was downloaded ('playwright install chromium'); otherwise the
        Microsoft Edge or Google Chrome already installed on this PC - so 'pip install playwright' is enough."""
        tries = []
        if Path("/opt/pw-browsers/chromium").exists():
            tries.append(("Chromium", {"executable_path": "/opt/pw-browsers/chromium"}))
        tries += [("Playwright Chromium", {}), ("Microsoft Edge", {"channel": "msedge"}),
                  ("Google Chrome", {"channel": "chrome"})]
        errors = []
        for name, kw in tries:
            try:
                return self._pw.chromium.launch(headless=True, **kw), name
            except Exception as e:
                errors.append(f"{name}: {str(e).splitlines()[0][:120]}")
        self._pw.stop()
        raise RuntimeError("no browser could be started - " + " | ".join(errors))

    def html(self, url: str) -> tuple[str, str]:
        page = self.ctx.new_page()
        try:
            page.goto(url, wait_until="networkidle", timeout=45000)
            page.wait_for_timeout(1500)
            return page.url, page.content()
        finally:
            page.close()

    def close(self):
        try:
            self.browser.close()
            self._pw.stop()
        except Exception:
            pass


def search_seed_urls(domain: str, max_results: int = 50) -> list[str]:
    """PDF URLs from a search API. Needs SERPAPI_KEY, or GOOGLE_CSE_KEY + GOOGLE_CSE_CX."""
    q = (f'site:{domain} filetype:pdf ("technical data sheet" OR datasheet OR "data sheet" '
         f'OR "product data" OR specification)')
    out: list[str] = []
    try:
        if os.environ.get("SERPAPI_KEY"):
            for start in range(0, max_results, 10):
                r = requests.get("https://serpapi.com/search.json", timeout=30, params={
                    "engine": "google", "q": q, "start": start, "num": 10,
                    "api_key": os.environ["SERPAPI_KEY"]})
                items = r.json().get("organic_results", []) if r.ok else []
                out += [i.get("link") for i in items if i.get("link")]
                if len(items) < 10:
                    break
        elif os.environ.get("GOOGLE_CSE_KEY") and os.environ.get("GOOGLE_CSE_CX"):
            for start in range(1, min(max_results, 100), 10):
                r = requests.get("https://www.googleapis.com/customsearch/v1", timeout=30, params={
                    "key": os.environ["GOOGLE_CSE_KEY"], "cx": os.environ["GOOGLE_CSE_CX"],
                    "q": q, "start": start})
                items = r.json().get("items", []) if r.ok else []
                out += [i.get("link") for i in items if i.get("link")]
                if len(items) < 10:
                    break
        else:
            log.warning("--search-seed set but no SERPAPI_KEY or GOOGLE_CSE_KEY/GOOGLE_CSE_CX in environment")
    except Exception as e:
        log.warning("search seeding failed for %s: %s", domain, e)
    return [u for u in out if u]


# =============================================================================
# DISCOVERY
# =============================================================================
class VendorCrawler:
    def __init__(self, vendor: Vendor, db: DB, args):
        self.v = vendor
        self.db = db
        self.args = args
        self.f = Fetcher(args)
        self.renderer: Renderer | None = None
        self.heap: list = []
        self.counter = itertools.count()
        self.known_pages: set[str] = set()
        self.known_pdfs: set[str] = set()
        self.prefix_locales = set()
        for u in vendor.start_urls + vendor.path_prefixes:
            segs = [s for s in urlparse(u if "://" in u else "http://x" + u).path.lower().split("/") if s]
            if segs and is_locale(segs[0]):
                self.prefix_locales.add(segs[0])

    # ---- scope -------------------------------------------------------------
    def in_scope_page(self, url: str) -> bool:
        if not host_allowed(host_of(url), self.v.domains):
            return False
        if ext_of(url) in SKIP_EXT:
            return False
        path = urlparse(url).path
        if self.v.path_prefixes and not any(path.startswith(p) for p in self.v.path_prefixes):
            return False
        if locale_blocked(url, self.prefix_locales):
            return False
        return True

    def pdf_allowed_host(self, url: str) -> bool:
        return self.args.offsite_pdfs or host_allowed(host_of(url), self.v.domains)

    # ---- queue persistence --------------------------------------------------
    def enqueue(self, url: str, depth: int, score: int):
        url = canon_url(url)
        if not url or url in self.known_pages or depth > self.args.max_depth:
            return
        self.known_pages.add(url)
        heapq.heappush(self.heap, (-(score) + depth, next(self.counter), url, depth, score))
        self.db.execute("INSERT OR IGNORE INTO pages(url, vendor, depth, score, state) VALUES (?,?,?,?, 'queued')",
                        (url, self.v.name, depth, score))

    def add_pdf(self, url: str, found_on: str, anchor: str, trusted: bool = False):
        """trusted: the link comes from the vendor's own document API (any host, no brand filter)."""
        url = canon_url(url)
        if not url or url in self.known_pdfs or not (trusted or self.pdf_allowed_host(url)):
            return False
        self.known_pdfs.add(url)
        text = f"{unquote(urlparse(url).path)} {anchor}"
        filt = None if trusted else self.v.filter_for(host_of(url))
        if filt is not None and not filt.search(text):
            return False  # another company's sheet on a shared (distributor) site
        dtype = classify_doc_type(text)
        segs = {sg.lower() for u in (url, found_on or "") for sg in urlparse(u).path.split("/") if sg}
        if segs & CORPORATE_SEGMENTS or NAME_JUNK_RE.search(stem_of(url)):
            dtype = "corporate"
        self.db.execute(
            "INSERT OR IGNORE INTO pdf_links(url, vendor, found_on, anchor, doc_type, first_seen, lang) "
            "VALUES (?,?,?,?,?,?,?)",
            (url, self.v.name, found_on, (anchor or "")[:300], dtype, utcnow(), translation_info(url)[0]))
        return True

    def load_state(self) -> bool:
        """Restore queue from DB. Returns True if there was prior state."""
        for r in self.db.query("SELECT url FROM pdf_links WHERE vendor=?", (self.v.name,)):
            self.known_pdfs.add(r["url"])
        rows = self.db.query("SELECT url, depth, score, state FROM pages WHERE vendor=?", (self.v.name,))
        for r in rows:
            self.known_pages.add(r["url"])
            if r["state"] == "queued":
                heapq.heappush(self.heap, (-(r["score"] or 0) + r["depth"], next(self.counter),
                                           r["url"], r["depth"], r["score"] or 0))
        return bool(rows)

    # ---- sitemaps ------------------------------------------------------------
    def sitemap_seed(self):
        bases = sorted({f"{urlparse(u).scheme}://{urlparse(u).netloc}" for u in self.v.start_urls})
        todo: list[str] = []
        for b in bases:
            self.f._robots_for(b + "/")
            todo += self.f.robot_sitemaps.get(b, []) or [b + "/sitemap.xml", b + "/sitemap_index.xml"]
        seen, n_urls, files = set(), 0, 0
        while todo and files < MAX_SITEMAP_FILES and n_urls < MAX_SITEMAP_URLS and not STOP.is_set():
            sm = todo.pop(0)
            if sm in seen or not host_allowed(host_of(sm), self.v.domains):
                continue
            seen.add(sm)
            if not self.f.allowed(sm):
                continue
            try:
                r = self.f.get(sm)
                files += 1
                if r.status_code != 200:
                    continue
                data = r.content
                if sm.endswith(".gz") or data[:2] == b"\x1f\x8b":
                    data = gzip.decompress(data)
                locs = re.findall(rb"<loc>\s*(.*?)\s*</loc>", data, re.S)
            except Exception as e:
                log.debug("sitemap %s: %s", sm, e)
                continue
            is_index = b"<sitemapindex" in data[:2000]
            for raw in locs:
                u = raw.decode("utf-8", "ignore").replace("&amp;", "&").strip()
                if is_index or u.lower().endswith((".xml", ".xml.gz")):
                    # prioritise sitemaps that look product/document related
                    (todo.insert(0, u) if score_link(u) > 0 else todo.append(u))
                    continue
                n_urls += 1
                if looks_like_pdf_url(u):
                    self.add_pdf(u, sm, "")
                elif self.in_scope_page(u):
                    sc = score_link(u)
                    if sc >= 3:
                        self.enqueue(u, 1, sc)
        log.info("[%s] sitemaps: %d files, %d URLs, %d queued, %d PDFs so far",
                 self.v.name, files, n_urls, len(self.heap), len(self.known_pdfs))

    # ---- page processing -------------------------------------------------------
    def extract(self, base_url: str, html: str, depth: int):
        soup = BeautifulSoup(html, "html.parser")
        base_tag = soup.find("base", href=True)
        base = urljoin(base_url, base_tag["href"]) if base_tag else base_url
        found: list[tuple[str, str]] = []
        for a in soup.find_all("a", href=True):
            found.append((a["href"], a.get_text(" ", strip=True) or a.get("title", "") or a.get("aria-label", "")))
        for tag, attr in (("iframe", "src"), ("embed", "src"), ("object", "data"), ("link", "href")):
            for t in soup.find_all(tag):
                if t.get(attr):
                    found.append((t[attr], ""))
        for t in soup.find_all(attrs={"data-href": True}):
            found.append((t["data-href"], t.get_text(" ", strip=True)))
        for t in soup.find_all(attrs={"data-url": True}):
            found.append((t["data-url"], t.get_text(" ", strip=True)))
        # PDF URLs embedded in scripts / JSON blobs
        for m in re.finditer(r"""(?:https?:)?(?:\\?/){1,2}[^\s"'<>()]+?\.pdf(?:\?[^\s"'<>()]*)?""", html, re.I):
            found.append((m.group(0).replace("\\/", "/"), ""))
        for href, text in found:
            href = href.strip()
            if not href or href.startswith(("mailto:", "tel:", "javascript:", "#", "data:")):
                continue
            u = canon_url(urljoin(base, href))
            if not u:
                continue
            if looks_like_pdf_url(u):
                self.add_pdf(u, base_url, text)
            elif self.in_scope_page(u):
                sc = score_link(u, text)
                if sc >= MIN_LINK_SCORE:
                    self.enqueue(u, depth + 1, sc)

    def fetch_page(self, url: str, depth: int) -> tuple[str, int | None, str]:
        doing(f"[{self.v.name}] reading page {url}")
        if not self.f.allowed(url):
            return "blocked", None, "robots.txt"
        use_js = self.v.render_js and self.args.js
        try:
            r = self.f.get(url, stream=True)
        except requests.RequestException as e:
            return "error", None, str(e)[:300]
        try:
            ctype = r.headers.get("Content-Type", "").lower()
            final = canon_url(r.url) or url
            if r.status_code >= 400:
                return "error", r.status_code, f"HTTP {r.status_code}"
            if "application/pdf" in ctype:
                self.add_pdf(final, url, "")
                return "done", r.status_code, "pdf"
            if ("html" not in ctype and "xml" not in ctype) or "rss" in ctype or "atom" in ctype:
                return "done", r.status_code, f"skipped {ctype[:40]}"
            if not host_allowed(host_of(final), self.v.domains):
                return "done", r.status_code, f"redirected off-site to {host_of(final)}"
            html = r.raw.read(6_000_000, decode_content=True).decode(r.encoding or "utf-8", "ignore")
        finally:
            r.close()
        if use_js:
            try:
                if self.renderer is None:
                    self.renderer = Renderer()
                final2, html2 = self.renderer.html(final)
                final, html = canon_url(final2) or final, html2
            except Exception as e:
                log.debug("render failed %s: %s", url, e)
        self.extract(final, html, depth)
        return "done", r.status_code, ""

    # ---- document APIs (sites whose document library is loaded by JavaScript) -----------------
    def api_seed(self):
        for r in self.db.query("SELECT url FROM pdf_links WHERE vendor=?", (self.v.name,)):
            self.known_pdfs.add(r["url"])
        if any(host_allowed(d, ["qnityelectronics.com"]) for d in self.v.domains):
            n = qnity_resource_center(self)
            log.info("[%s] resource-center API: %d document link(s)", self.v.name, n)
        if not self.args.no_wp_media:
            n = wordpress_media_pdfs(self)
            if n:
                log.info("[%s] WordPress media library: %d PDF link(s)", self.v.name, n)

    def run(self):
        v = self.v
        try:
            self.api_seed()
        except Exception as e:
            log.warning("[%s] document API failed: %s", v.name, e)
        if not self.args.recrawl:
            row = self.db.query("SELECT discovery_done FROM vendors WHERE vendor=?", (v.name,))
            if row and row[0]["discovery_done"]:
                log.info("[%s] discovery already complete (use --recrawl to redo)", v.name)
                return
        else:
            self.db.execute("DELETE FROM pages WHERE vendor=?", (v.name,))
        self.db.execute(
            "INSERT INTO vendors(vendor, categories, status, domains) VALUES (?,?,?,?) "
            "ON CONFLICT(vendor) DO UPDATE SET categories=excluded.categories, status=excluded.status, "
            "domains=excluded.domains, discovery_done=0",
            (v.name, "; ".join(v.categories), v.status, ";".join(v.domains)))
        resumed = self.load_state()
        if not resumed or not self.heap:
            for u in v.start_urls:
                self.enqueue(u, 0, 100)
            if not self.args.no_sitemaps:
                self.sitemap_seed()
            if self.args.search_seed:
                for d in v.domains:
                    hits = search_seed_urls(d)
                    for u in hits:
                        if looks_like_pdf_url(u):
                            self.add_pdf(u, "search", "")
                        elif self.in_scope_page(u):
                            self.enqueue(u, 1, 10)
                    log.info("[%s] search seed %s: %d hits", v.name, d, len(hits))
        else:
            log.info("[%s] resuming: %d queued pages", v.name, len(self.heap))

        cap = v.max_pages or self.args.max_pages
        fetched = 0
        while self.heap and fetched < cap and not STOP.is_set():
            _, _, url, depth, score = heapq.heappop(self.heap)
            state, code, err = self.fetch_page(url, depth)
            fetched += 1
            self.db.execute("UPDATE pages SET state=?, http_status=?, fetched_at=?, error=? WHERE url=? AND vendor=?",
                            (state, code, utcnow(), err, url, v.name))
            if fetched % 25 == 0:
                log.info("[%s] %d pages fetched, %d queued, %d PDF links",
                         v.name, fetched, len(self.heap), len(self.known_pdfs))
        done = 0 if (STOP.is_set() or self.heap) else 1
        self.db.execute("UPDATE vendors SET discovery_done=?, last_discovery=? WHERE vendor=?",
                        (done, utcnow(), v.name))
        if self.renderer:
            self.renderer.close()
        log.info("[%s] discovery %s: %d pages fetched, %d PDF links known%s",
                 v.name, "complete" if done else ("interrupted" if STOP.is_set() else "paused"), fetched, len(self.known_pdfs),
                 f" (page cap {cap} reached, {len(self.heap)} left queued - re-run --discover to continue)"
                 if self.heap and fetched >= cap else "")


# WordPress sites list every uploaded file through their public REST API. Many vendor sites only link a
# few of their PDFs from crawlable pages (the rest sit behind filters, accordions or forms), so the
# media library is asked for all PDFs. robots.txt is honoured for the API address.
WP_MEDIA_MAX_PAGES = 30          # x 100 files


def wordpress_media_pdfs(crawler: "VendorCrawler") -> int:
    bases = sorted({f"{urlparse(u).scheme}://{urlparse(u).netloc}" for u in crawler.v.start_urls if u})
    n = 0
    for base in bases:
        api = f"{base}/wp-json/wp/v2/media"
        for page in range(1, WP_MEDIA_MAX_PAGES + 1):
            if STOP.is_set():
                return n
            url = f"{api}?mime_type=application/pdf&per_page=100&page={page}&_fields=source_url,title"
            if not crawler.f.allowed(url):
                break
            try:
                r = crawler.f.get(url)
                if r.status_code != 200 or "json" not in r.headers.get("Content-Type", ""):
                    break
                items = r.json()
            except (requests.RequestException, ValueError):
                break
            if not isinstance(items, list) or not items:
                break
            for it in items:
                src = it.get("source_url") or ""
                title = it.get("title") or {}
                title = html_lib.unescape(title.get("rendered", "") if isinstance(title, dict) else str(title))
                if src.lower().split("?")[0].endswith(".pdf") and \
                        crawler.add_pdf(src, f"{base}/wp-json/wp/v2/media", title):
                    n += 1
            if len(items) < 100 or page >= int(r.headers.get("X-WP-TotalPages", page) or page):
                break
    return n


# Qnity Electronics (former DuPont Electronics): the Resource Center is a JavaScript app that loads its
# document list from POST /api/resource-center; the files live on a Sitecore Content Hub host with no
# ".pdf" in the URL, so ordinary crawling finds nothing.
QNITY_LITERATURE_TYPES = ("technical-data-sheet", "product-data-sheet", "product-information-sheet",
                          "technical-information", "brochure", "product-overview", "application",
                          "safety-data-sheet-(sds)")
# product groups that are webs / web converting (films, laminates, flex circuits, dry-film resists,
# insulation papers and composites, thermal tapes); Kalrez seals, semiconductor chemicals and display
# materials are left out
QNITY_KEEP_TAG_RE = re.compile(r"printed-circuit-board|thermal-management|electrical-insulation|polyimide|"
                               r"laminate|circuit-imaging|flexible|kapton|pyralux|oasis|temprion|riston|interra|"
                               r"coverlay|bondply|film", re.I)
QNITY_DROP_TAG_RE = re.compile(r"perfluoroelastomer|kalrez|semiconductor|ekc|display-and-lighting|"
                               r"industrial-finishing", re.I)


def qnity_resource_center(crawler: "VendorCrawler") -> int:
    base = "https://www.qnityelectronics.com"
    found_on = f"{base}/resource-center"
    seen, n = set(), 0
    for lt in QNITY_LITERATURE_TYPES:
        if STOP.is_set():
            break
        crawler.f._pace(base + "/api/resource-center")
        try:
            r = crawler.f.s.post(base + "/api/resource-center", timeout=120,
                                 json={"tagIds": [f"electronics:literature-type/{lt}"], "tagNames": [],
                                       "articleTagNames": [], "excludeTagIds": [], "excludeTagNames": [],
                                       "folderCategoryName": "", "language": "en"})
            results = r.json().get("results", []) if r.status_code == 200 else []
        except (requests.RequestException, ValueError) as e:
            log.debug("qnity %s: %s", lt, e)
            continue
        for it in results:
            url = it.get("url") or ""
            if not url or url in seen or it.get("isSecure") or it.get("isLocked"):
                continue
            if (it.get("fileType") or "PDF").upper() != "PDF":
                continue
            tags = " ".join(it.get("matchedTags") or [])
            title = it.get("title") or ""
            text = f"{title} {tags}"
            if QNITY_DROP_TAG_RE.search(text) or not QNITY_KEEP_TAG_RE.search(text):
                continue
            seen.add(url)
            crawler.add_pdf(url, f"{found_on}?tagId=electronics:literature-type/{lt}", title, trusted=True)
            n += 1
    return n


# =============================================================================
# DOWNLOAD
# =============================================================================
_file_lock = threading.Lock()


def filename_from_response(r: requests.Response, url: str) -> str:
    cd = r.headers.get("Content-Disposition", "")
    m = re.search(r"filename\*\s*=\s*[^']*''([^;]+)", cd) or re.search(r'filename\s*=\s*"?([^";]+)"?', cd)
    if m:
        name = unquote(m.group(1))
        if re.search(r"%[0-9A-Fa-f]{2}", name):      # double-encoded ('%2520' -> '%20' -> ' ')
            name = unquote(name)
        return safe_filename(name)
    last = urlparse(r.url or url).path.rstrip("/").rsplit("/", 1)[-1]
    return safe_filename(last or "document")


def fetch_url(url: str) -> str:
    """The address to request for a PDF link. Dow sends plain /documents/...pdf requests to its HTML
    document viewer; the same address with ?iframe=true returns the PDF itself."""
    p = urlparse(url)
    if host_allowed(p.hostname or "", ["dow.com"]) and "/documents/" in p.path and \
            p.path.lower().endswith(".pdf") and "iframe=" not in p.query:
        return url + ("&" if p.query else "?") + "iframe=true"
    return url


def _download_one(row, f: Fetcher, db: DB, args, root: Path, vendor: Vendor, out_dir: Path) -> str:
    """Download one pdf_links row. Returns the resulting state."""
    url = row["url"]
    doing(f"[{vendor.name}] downloading {url}")
    max_bytes = args.max_pdf_mb * 1024 * 1024
    if not f.allowed(url):
        db.execute("UPDATE pdf_links SET state='robots' WHERE url=?", (url,))
        return "robots"
    hdrs = {}
    if args.refresh and row["state"] in GOOD_STATES:
        if row["etag"]:
            hdrs["If-None-Match"] = row["etag"]
        if row["last_modified"]:
            hdrs["If-Modified-Since"] = row["last_modified"]
    out_dir.mkdir(parents=True, exist_ok=True)
    tmp = out_dir / f".part_{threading.get_ident()}"
    try:
        r = f.get(fetch_url(url), stream=True, headers=hdrs)
    except requests.RequestException as e:
        db.execute("UPDATE pdf_links SET state='error', error=?, attempts=attempts+1 WHERE url=?",
                   (str(e)[:300], url))
        return "error"
    try:
        if r.status_code == 304:
            db.execute("UPDATE pdf_links SET state='unchanged', downloaded_at=? WHERE url=?", (utcnow(), url))
            return "unchanged"
        if r.status_code >= 400:
            db.execute("UPDATE pdf_links SET state='error', http_status=?, error=?, attempts=attempts+1 WHERE url=?",
                       (r.status_code, f"HTTP {r.status_code}", url))
            return "error"
        clen = int(r.headers.get("Content-Length") or 0)
        if clen > max_bytes:
            db.execute("UPDATE pdf_links SET state='too_large', bytes=? WHERE url=?", (clen, url))
            return "too_large"
        h = hashlib.sha256()
        size, head, ok = 0, b"", True
        t0, too_slow = time.time(), False
        with open(tmp, "wb") as fh:
            for chunk in r.iter_content(65536):
                if STOP.is_set():
                    ok = False
                    break
                if time.time() - t0 > MAX_DOWNLOAD_SECONDS:
                    ok, too_slow = False, True
                    break
                if len(head) < 1024:
                    head += chunk[:1024 - len(head)]
                    if len(head) >= 8 and b"%PDF" not in head:
                        ok = False
                        break
                size += len(chunk)
                if size > max_bytes:
                    ok = False
                    break
                h.update(chunk)
                fh.write(chunk)
        if not ok or b"%PDF" not in head:
            tmp.unlink(missing_ok=True)
            if STOP.is_set():
                return "stopped"
            if too_slow:
                db.execute("UPDATE pdf_links SET state='error', error=?, attempts=attempts+1 WHERE url=?",
                           (f"download took over {MAX_DOWNLOAD_SECONDS} s - will retry next run", url))
                log.info("  %s: gave up after %d s (slow server), will retry next run", url, MAX_DOWNLOAD_SECONDS)
                return "error"
            state = "too_large" if size > max_bytes else "not_pdf"
            db.execute("UPDATE pdf_links SET state=?, http_status=?, attempts=attempts+1 WHERE url=?",
                       (state, r.status_code, url))
            return state
        sha = h.hexdigest()
        etag, lm = r.headers.get("ETag"), r.headers.get("Last-Modified")
        with _file_lock:
            existing = db.query("SELECT path FROM files WHERE sha256=?", (sha,))
            if existing:
                tmp.unlink(missing_ok=True)
                st = "unchanged" if row["sha256"] == sha else "duplicate"
                db.execute("UPDATE pdf_links SET state=?, sha256=?, path=?, bytes=?, http_status=?, etag=?, "
                           "last_modified=?, downloaded_at=?, attempts=attempts+1 WHERE url=?",
                           (st, sha, existing[0]["path"], size, r.status_code, etag, lm, utcnow(), url))
                return st
            name = filename_from_response(r, url)
            dest = out_dir / name
            if dest.exists():
                dest = out_dir / f"{dest.stem}_{sha[:8]}.pdf"
            os.replace(tmp, dest)
            rel = str(dest.relative_to(root))
            db.execute("INSERT INTO files(sha256, path, bytes, vendor, first_url, added_at) VALUES (?,?,?,?,?,?)",
                       (sha, rel, size, vendor.name, url, utcnow()))
            db.execute("UPDATE pdf_links SET state='downloaded', sha256=?, path=?, bytes=?, http_status=?, etag=?, "
                       "last_modified=?, downloaded_at=?, attempts=attempts+1, error=NULL WHERE url=?",
                       (sha, rel, size, r.status_code, etag, lm, utcnow(), url))
            return "downloaded"
    except Exception as e:
        tmp.unlink(missing_ok=True)
        db.execute("UPDATE pdf_links SET state='error', error=?, attempts=attempts+1 WHERE url=?",
                   (str(e)[:300], url))
        return "error"
    finally:
        r.close()


def _english_links(db: DB, vendor: Vendor) -> dict[str, sqlite3.Row]:
    """match_key -> English (non-translated) link row for this vendor."""
    out: dict[str, sqlite3.Row] = {}
    for r in db.query("SELECT * FROM pdf_links WHERE vendor=? AND COALESCE(lang,'')='' "
                      "AND state NOT IN ('guess_failed','translation')", (vendor.name,)):
        k = match_key(stem_of(r["url"]))
        if k and (k not in out or r["state"] in GOOD_STATES):
            out[k] = r
    return out


def _add_english_guesses(db: DB, vendor: Vendor, row) -> list[str]:
    guesses = english_url_guesses(row["url"])
    for g in guesses:
        db.execute("INSERT OR IGNORE INTO pdf_links(url, vendor, found_on, anchor, doc_type, first_seen, lang, "
                   "english_of) VALUES (?,?,?,?,?,?, '', ?)",
                   (canon_url(g), vendor.name, "english-guess", "", row["doc_type"], utcnow(), row["url"]))
    return [canon_url(g) for g in guesses]


def _english_for(db: DB, english: dict, row):
    """A successfully downloaded English counterpart of a translated link, or None."""
    _, key, _ = translation_info(row["url"])
    cand = english.get(key)
    if cand is not None:
        cur = db.query("SELECT * FROM pdf_links WHERE url=?", (cand["url"],))
        if cur and cur[0]["state"] in GOOD_STATES:
            return cur[0]
    g = db.query(f"SELECT * FROM pdf_links WHERE english_of=? AND state IN ({','.join('?' * len(GOOD_STATES))}) "
                 "LIMIT 1", (row["url"], *GOOD_STATES))
    return g[0] if g else None


def _guess_already_found(db: DB, row) -> bool:
    return bool(db.query(f"SELECT 1 FROM pdf_links WHERE english_of=? AND url<>? AND state IN "
                         f"({','.join('?' * len(GOOD_STATES))}) LIMIT 1", (row["english_of"], row["url"], *GOOD_STATES)))


def tidy_translations(db: DB, vendor: Vendor, make_guesses: bool) -> int:
    """Already-downloaded translated copies: if the English version is on disk, mark the copy as a
    translation (it moves to _translations/); otherwise queue guessed English URLs to try."""
    english = _english_links(db, vendor)
    marked = 0
    for r in db.query("SELECT * FROM pdf_links WHERE vendor=? AND COALESCE(lang,'')<>'' AND state='downloaded'",
                      (vendor.name,)):
        en = _english_for(db, english, r)
        if en is not None:
            db.execute("UPDATE pdf_links SET state='translation', english_url=? WHERE url=?", (en["url"], r["url"]))
            db.execute("UPDATE files SET is_translation=1, english_path=? WHERE sha256=?", (en["path"], r["sha256"]))
            marked += 1
        elif make_guesses and not db.query("SELECT 1 FROM pdf_links WHERE english_of=? LIMIT 1", (r["url"],)):
            _add_english_guesses(db, vendor, r)
    return marked


def ensure_vendor_folders(vendors: dict, root: Path) -> int:
    """One folder per workbook vendor under pdfs/<Category>/<Vendor>/, created even before anything is
    downloaded, so a vendor with no files yet is easy to spot."""
    n = 0
    for v in vendors.values():
        d = v.pdf_dir(root)
        if not d.exists():
            d.mkdir(parents=True, exist_ok=True)
            n += 1
    if n:
        log.info("Created %d empty vendor folder(s) under pdfs/", n)
    return n


def download_vendor(vendor: Vendor, db: DB, args, root: Path) -> dict:
    f = Fetcher(args)
    types = set(args.doc_types)
    out_dir = vendor.pdf_dir(root)
    stats = {"downloaded": 0, "duplicate": 0, "unchanged": 0, "failed": 0, "skipped_type": 0,
             "translations_with_english": 0, "english_found_by_guess": 0, "foreign_no_english": 0}
    english_first = not args.keep_translations
    if english_first:
        tidy_translations(db, vendor, make_guesses=True)
    states = ("pending", "error") if not args.refresh else ("pending", "error", "downloaded", "unchanged")
    rows = db.query(
        f"SELECT * FROM pdf_links WHERE vendor=? AND state IN ({','.join('?' * len(states))}) "
        f"AND attempts < ? ORDER BY url", (vendor.name, *states, args.max_attempts if not args.refresh else 999))
    # guessed English URLs are mostly misses; on hosts that ask for long pauses between requests (Barmag:
    # 60 s) trying them would take hours, so they are skipped there
    slow = [r for r in rows if r["found_on"] == "english-guess" and f.crawl_delay(r["url"]) > SLOW_HOST_DELAY]
    if slow:
        for r in slow:
            db.execute("UPDATE pdf_links SET state='guess_failed', error='skipped: slow host' WHERE url=?", (r["url"],))
        log.info("[%s] skipped %d guessed English URL(s) on a slow host (robots.txt crawl-delay > %ds)",
                 vendor.name, len(slow), SLOW_HOST_DELAY)
        rows = [r for r in rows if r not in slow]

    def tally(state: str):
        if state in ("downloaded", "duplicate", "unchanged"):
            stats[state] += 1
        elif state not in ("robots", "stopped"):
            stats["failed"] += 1

    # Pass 1: English documents (and guessed English URLs); translated copies wait
    foreign = []
    english = _english_links(db, vendor) if english_first else {}
    for row in rows:
        if STOP.is_set():
            break
        if row["doc_type"] not in types:
            stats["skipped_type"] += 1
            continue
        if link_off_topic(vendor.name, row["url"], row["anchor"] or "", (vendor.scope_keep, vendor.scope_exclude),
                          row["found_on"] or ""):
            db.execute("UPDATE pdf_links SET state='off_topic' WHERE url=?", (row["url"],))
            stats["skipped_off_topic"] = stats.get("skipped_off_topic", 0) + 1
            continue
        if english_first and row["lang"]:
            foreign.append(row)
            _, key, _ = translation_info(row["url"])
            if key not in english:
                _add_english_guesses(db, vendor, row)
            continue
        if row["english_of"] and _guess_already_found(db, row):
            db.execute("UPDATE pdf_links SET state='guess_failed' WHERE url=?", (row["url"],))
            continue
        st = _download_one(row, f, db, args, root, vendor, out_dir)
        if row["english_of"]:
            if st in GOOD_STATES:
                stats["english_found_by_guess"] += 1
            else:
                db.execute("UPDATE pdf_links SET state='guess_failed' WHERE url=?", (row["url"],))
            continue
        tally(st)
    # guesses created during this pass
    if english_first and foreign and not STOP.is_set():
        guess_rows = db.query("SELECT * FROM pdf_links WHERE vendor=? AND english_of IS NOT NULL AND state='pending'",
                              (vendor.name,))
        for row in guess_rows:
            if STOP.is_set():
                break
            if _guess_already_found(db, row):
                db.execute("UPDATE pdf_links SET state='guess_failed' WHERE url=?", (row["url"],))
                continue
            st = _download_one(row, f, db, args, root, vendor, out_dir)
            if st in GOOD_STATES:
                stats["english_found_by_guess"] += 1
            else:
                db.execute("UPDATE pdf_links SET state='guess_failed' WHERE url=?", (row["url"],))
    # Pass 2: translated copies - skip if the English version exists, otherwise keep the copy
    english = _english_links(db, vendor) if english_first else {}
    for row in foreign:
        if STOP.is_set():
            break
        en = _english_for(db, english, row)
        st = _download_one(row, f, db, args, root, vendor, out_dir)
        if en is not None:
            # English version is in the main folder; keep this copy too, filed under _translations/
            if st == "downloaded":
                cur = db.query("SELECT sha256 FROM pdf_links WHERE url=?", (row["url"],))
                db.execute("UPDATE pdf_links SET state='translation', english_url=? WHERE url=?", (en["url"], row["url"]))
                if cur and cur[0]["sha256"]:
                    db.execute("UPDATE files SET is_translation=1, english_path=? WHERE sha256=?",
                               (en["path"], cur[0]["sha256"]))
                stats["translations_with_english"] += 1
            else:
                tally(st)
            continue
        tally(st)
        if st == "downloaded":
            stats["foreign_no_english"] += 1
    if english_first:
        tidy_translations(db, vendor, make_guesses=False)
    log.info("[%s] download: %s", vendor.name, ", ".join(f"{k}={v}" for k, v in stats.items() if v))
    return stats


def backfill_link_lang(db: DB) -> None:
    """Links recorded before translation handling existed get their language tag filled in."""
    # tags from folders that only looked like a locale (e.g. '/jm/' on jm.com) are re-derived
    for r in db.query("SELECT DISTINCT lang FROM pdf_links WHERE lang IS NOT NULL AND lang <> ''"):
        if r["lang"] not in LANG_CODES.values() and r["lang"] not in LOCALE_CODES and r["lang"] not in LANG_WORDS.values():
            db.execute("UPDATE pdf_links SET lang=NULL WHERE lang=?", (r["lang"],))
    rows = db.query("SELECT url FROM pdf_links WHERE lang IS NULL")
    if rows:
        db.executemany("UPDATE pdf_links SET lang=? WHERE url=?", [(translation_info(r["url"])[0], r["url"]) for r in rows])
    corp = []
    for r in db.query("SELECT url, found_on FROM pdf_links WHERE doc_type <> 'corporate'"):
        segs = {sg.lower() for u in (r["url"], r["found_on"] or "") for sg in urlparse(u).path.split("/") if sg}
        if segs & CORPORATE_SEGMENTS:
            corp.append((r["url"],))
    if corp:
        db.executemany("UPDATE pdf_links SET doc_type='corporate' WHERE url=?", corp)


def apply_pdf_filters(db: DB, vendors: dict[str, Vendor]) -> int:
    """On sites shared by several vendors (e.g. a distributor), give each PDF to the vendor whose
    brand filter matches it, else to a vendor on that site without a filter (the distributor)."""
    active = [v for v in vendors.values() if v.domains]
    if not any(v.pdf_filter is not None for v in active):
        return 0
    moved = 0
    for r in db.query("SELECT url, vendor, anchor FROM pdf_links"):
        host = host_of(r["url"])
        on_host = [v for v in active if host_allowed(host, v.domains)]
        if len(on_host) < 2:
            continue
        text = f"{unquote(urlparse(r['url']).path)} {r['anchor'] or ''}"
        specific = [v for v in on_host if v.filter_for(host) is not None and v.filter_for(host).search(text)]
        generic = [v for v in on_host if v.filter_for(host) is None]
        pool = specific or generic
        if not pool or any(v.name == r["vendor"] for v in pool):
            continue
        new = pool[0].name
        db.execute("UPDATE pdf_links SET vendor=? WHERE url=?", (new, r["url"]))
        moved += 1
    if moved:
        db.execute("UPDATE files SET vendor = (SELECT l.vendor FROM pdf_links l WHERE l.url = files.first_url) "
                   "WHERE EXISTS (SELECT 1 FROM pdf_links l WHERE l.url = files.first_url AND l.vendor <> files.vendor)")
        log.info("Re-assigned %d PDF link(s) between vendors sharing a site (brand filters)", moved)
    return moved


def reorganize_files(db: DB, vendors: dict[str, Vendor], root: Path) -> int:
    """Move already-downloaded PDFs into pdfs/<Category>/<Vendor>/ (idempotent)."""
    by_name = {v.name: v for v in vendors.values()}
    moved = 0
    old_dirs: set[Path] = set()
    rows = db.query("SELECT sha256, path, vendor, is_translation, rejected, bytes FROM files")
    gone: list = []
    for row in rows:
        v = by_name.get(row["vendor"])
        if v is None:
            continue
        old = root / row["path"]
        sub = "_rejected" if row["rejected"] else ("_translations" if row["is_translation"] else "")
        target_dir = v.pdf_dir(root) / sub
        if not old.exists():
            found = _find_moved_file(root, old.name, row["bytes"], v)
            if found is None:
                gone.append(row)
                continue
            old = found        # moved by hand: pick it up from where it is now
            rel = str(old.relative_to(root))
            db.execute("UPDATE files SET path=? WHERE sha256=?", (rel, row["sha256"]))
            db.execute("UPDATE pdf_links SET path=? WHERE sha256=?", (rel, row["sha256"]))
        if old.parent == target_dir:
            continue
        target_dir.mkdir(parents=True, exist_ok=True)
        dest = target_dir / old.name
        if dest.exists():
            dest = target_dir / f"{old.stem}_{row['sha256'][:8]}.pdf"
        os.replace(old, dest)
        old_dirs.add(old.parent)
        new_rel = str(dest.relative_to(root))
        db.execute("UPDATE files SET path=? WHERE sha256=?", (new_rel, row["sha256"]))
        db.execute("UPDATE pdf_links SET path=? WHERE sha256=?", (new_rel, row["sha256"]))
        moved += 1
    keep_dirs = {v.pdf_dir(root) for v in vendors.values()}      # every vendor keeps its folder, even empty
    for d in sorted(old_dirs, key=lambda p: len(p.parts), reverse=True):
        try:
            if d != root / "pdfs" and d not in keep_dirs and not any(d.iterdir()):
                d.rmdir()
        except OSError:
            pass
    if moved:
        log.info("Moved %d PDF(s) into pdfs/<Category>/<Vendor>/ (translations -> _translations/, "
                 "junk -> _rejected/)", moved)
    if gone:
        if not (root / "pdfs").exists() or len(gone) > max(20, 0.3 * len(rows)):
            log.warning("%d file(s) listed in the database are not on disk - is the drive connected? "
                        "Nothing was changed.", len(gone))
        else:
            for row in gone:
                db.execute("UPDATE pdf_links SET state='deleted' WHERE sha256=?", (row["sha256"],))
                db.execute("DELETE FROM files WHERE sha256=?", (row["sha256"],))
                db.execute("UPDATE files SET english_path=NULL WHERE english_path=?", (row["path"],))
            names = ", ".join(Path(r["path"]).name for r in gone[:3]) + (" ..." if len(gone) > 3 else "")
            log.info("%d file(s) were deleted outside the crawler (%s) - removed from the database and "
                     "won't be downloaded again", len(gone), names)
    return moved


def _find_moved_file(root: Path, name: str, size, vendor: Vendor):
    """A file the user moved by hand: look in the vendor's folders first, then anywhere under pdfs/."""
    for base in (vendor.pdf_dir(root), root / "pdfs"):
        if not base.exists():
            continue
        for cand in base.rglob(name):
            try:
                if cand.is_file() and (not size or cand.stat().st_size == size):
                    return cand
            except OSError:
                continue
    return None


# =============================================================================
# CLASSIFY (pypdf)
# =============================================================================
def extract_methods(text: str) -> tuple[list[str], list[str]]:
    text = unicodedata.normalize("NFKC", text or "")
    found: dict = {}
    fams: dict = {}
    for fam, rx in TEST_METHOD_PATTERNS:
        for m in rx.finditer(text):
            g = [x for x in m.groups() if x]
            if fam in ("ASTM", "JIS"):
                label = f"{fam} {g[0]}{g[1]}"
            elif fam == "TAPPI":
                label = f"TAPPI T{g[0]}"
            elif fam == "IPC":
                label = "IPC-TM-650" + (f" {g[0]}" if g else "")
            else:
                label = f"{fam} {g[0]}"
            found[label] = None
            fams[fam] = None
    return list(found)[:40], list(fams)


def classify_text_doc_type(text: str, link_type: str = "", n_methods: int = 0, hints: str = "") -> str:
    """Classify from extracted PDF text: header first, body as fallback. A filename that already says
    'data sheet', or a page citing 2+ test methods, counts as a data sheet unless it is an SDS."""
    # NFKC: full-width letters, compatibility ideographs and decomposed kana (フ+゛ -> ブ) match the patterns
    text = unicodedata.normalize("NFKC", text or "")
    hints = unicodedata.normalize("NFKC", hints or "")
    if link_type == "corporate" or (hints and CORPORATE_TEXT_RE.search(hints)):
        return "corporate"
    if COMMERCIAL_TEXT_RE.search(hints or "") or COMMERCIAL_TEXT_RE.search(DISCLAIMER_RE.sub(" ", text[:2500])):
        return "commercial"
    if not text.strip():
        return "unreadable"
    rx = dict(DOC_TYPE_PATTERNS)
    head = text[:800]
    if rx["sds"].search(head) and not rx["tds"].search(head[:200]):
        return "sds"
    if CORPORATE_TEXT_RE.search(text[:1500]):
        return "corporate"
    if link_type == "tds" or n_methods >= 2:
        return "tds"
    if CORPORATE_TEXT_RE.search(text[:5000]):
        return "corporate"
    if rx["corporate"].search(head):
        return "corporate"
    for t in ("tds", "brochure"):
        if rx[t].search(head):
            return t
    if TEXT_CERT_RE.search(head):
        return "certificate"
    for t in ("tds", "brochure"):
        if rx[t].search(text[:6000]):
            return t
    if TEXT_CERT_RE.search(text[:6000]):
        return "certificate"
    return "other"


def name_pattern(path: str) -> str:
    """'201903_nmm.pdf' -> '#_nmm' : digits collapsed, our _xxxxxxxx collision suffix removed."""
    stem = COLLISION_SUFFIX_RE.sub("", Path(path).stem.lower())
    return re.sub(r"\d+", "#", stem)


def infer_from_siblings(db: DB) -> int:
    """Files whose text can't be judged (garbled / image-only / 'other') inherit 'corporate' when most
    of their same-vendor siblings with the same name pattern (e.g. ####_nmm) are corporate."""
    groups: dict = {}
    for r in db.query("SELECT sha256, path, vendor, text_doc_type FROM files WHERE text_doc_type IS NOT NULL"):
        pat = name_pattern(r["path"])
        if not re.search(r"[a-z]", pat):
            continue  # pure dates/numbers are too generic to group on
        groups.setdefault((r["vendor"], pat), []).append(r)
    n = 0
    for members in groups.values():
        if len(members) < 3:
            continue
        judged = [m for m in members if m["text_doc_type"] not in ("other", "unreadable")]
        corp = [m for m in judged if m["text_doc_type"] == "corporate"]
        if len(corp) >= 2 and len(corp) >= 0.75 * len(judged):
            for m in members:
                if m["text_doc_type"] in ("other", "unreadable"):
                    db.execute("UPDATE files SET text_doc_type='corporate' WHERE sha256=?", (m["sha256"],))
                    n += 1
    if n:
        log.info("Marked %d unreadable/unclear file(s) corporate because their same-named siblings are", n)
    return n


# =============================================================================
# TOPIC SCOPE - what the corpus is about
# =============================================================================
# Scope comes from the "Patent Search List for Corpus Additions": a document is on-topic when it is about
# webs (film, foil, paper, nonwovens, fabric), the processes that make them, web converting (coating,
# laminating, adhesives, printing, cutting, labels, tapes...) or web handling equipment. Anything that relates
# to any of these is kept, data sheet or not (SDSs included); everything else is off-topic and is deleted
# (--keep-off-topic moves it to _rejected/ instead). Company paperwork is off-topic for every vendor.
# Vendors whose every product document counts as on-topic:
TOPIC_KEEP_ALL_VENDORS = {"ExxonMobil Chemical"}

_TOPIC_TERMS = [
    # webs
    r"webs?", r"films?", r"foils?", r"papers?", r"paperboards?", r"cardboard", r"board stock", r"nonwovens?",
    r"non[- ]wovens?", r"fabrics?", r"textiles?", r"scrims?", r"felts?", r"fleeces?", r"laminates?", r"laminat\w+",
    r"sheet(ing)? goods", r"roll goods", r"membranes?", r"substrates?", r"facestocks?", r"liners?", r"tapes?",
    r"labels?", r"labelstock", r"wraps?", r"pouch(es)?", r"bags?", r"sacks?", r"packaging", r"overlaminates?",
    r"prepregs?", r"bondply", r"coverlays?", r"copper[- ]clad", r"ccl", r"strip",
    # polymers named in the list
    r"ldpe", r"lldpe", r"hdpe", r"mdpe", r"uhmw(-?pe)?", r"bope", r"mdo(-pe)?", r"polyethylene", r"polyolefins?",
    r"bopp", r"cpp", r"opp", r"polypropylene", r"bopet", r"pet", r"pen", r"pbt", r"petg", r"polyester",
    r"polyamide", r"nylon", r"bopa", r"pa ?6(,?6)?", r"pvc", r"pvdc", r"evoh", r"eva", r"polystyrene", r"hips",
    r"polycarbonate", r"pmma", r"acrylic", r"polyimide", r"kapton", r"ptfe", r"fep", r"pfa", r"etfe", r"pvdf",
    r"ectfe", r"fluoropolymers?", r"peek", r"pps", r"pei", r"ultem", r"psu", r"ppsu", r"tpu", r"tpe", r"tpo",
    r"sebs", r"pvoh", r"pva", r"pla", r"pha", r"cellophane", r"cellulose", r"ionomers?", r"surlyn", r"coc",
    r"lcp", r"elastomers?", r"epdm", r"resins?", r"polymers?", r"copolymers?", r"silicones?", r"polyurethanes?",
    r"aluminium", r"aluminum", r"copper", r"metalli[sz]ed", r"metalli[sz]ing", r"vinyl", r"foams?",
    # nonwoven types / end uses
    r"carded", r"air[- ]?laid", r"wet[- ]?laid", r"dry[- ]?laid", r"spunbond(ed)?", r"melt[- ]?blown", r"spunlace",
    r"hydroentangl\w*", r"needle[- ]?punch\w*", r"thermal(ly)? bond\w*", r"sms", r"electret", r"filter media",
    r"filtration", r"filters?", r"geotextiles?", r"hygiene", r"wipes?", r"topsheets?", r"backsheets?",
    r"absorbent", r"diapers?", r"sterili[sz]ation wrap", r"surgical (gowns?|drapes?)", r"masks?", r"respirators?",
    r"wound dressings?", r"dressings?", r"battery separators?", r"separators?", r"interlinings?", r"insulation",
    r"acoustic", r"fibers?", r"fibres?", r"bicomponent", r"viscose", r"lyocell", r"pulp",
    # web processes
    r"cast(ing)? film", r"casting", r"extru\w+", r"co-?extru\w+", r"tenter\w*", r"calend[ae]r\w*",
    r"papermaking", r"paper machines?", r"blown[- ]film", r"sheet extrusion", r"orient\w+", r"biaxial\w*",
    r"stretch\w*", r"heat[- ]set\w*", r"anneal\w*", r"rolling mill", r"rolled", r"electrodeposited", r"spinnerets?",
    r"spinning", r"dies?", r"melt", r"granules?", r"pellets?",
    # converting
    r"coat(ing|ings|ed|er|ers)", r"printing", r"printable", r"print(ed|er|ers)?", r"flexo\w*", r"gravure",
    r"offset", r"inkjet", r"digital print\w*", r"screen print\w*", r"inks?", r"primers?", r"tie[- ]layers?",
    r"adhesives?", r"pressure[- ]sensitive", r"psa", r"hot[- ]?melts?", r"release (coatings?|liners?|films?|papers?)",
    r"barrier", r"heat[- ]seal\w*", r"sealants?", r"anti[- ]?fog", r"slip", r"lamination", r"laminating",
    r"laminators?", r"extrusion lamination", r"corona", r"plasma", r"flame treat\w*", r"surface treat\w*",
    r"vacuum (coating|metalli[sz]\w+|deposition)", r"vapou?r deposition", r"sputter\w*", r"curing", r"cure",
    r"uv", r"electron beam", r"drying", r"dryers?", r"ovens?", r"emboss\w*", r"hot stamping", r"foil stamping",
    r"cold foil", r"transfer", r"sheeters?", r"die[- ]cut\w*", r"kiss[- ]cut\w*", r"laser cut\w*", r"perforat\w*",
    r"punch\w*", r"slit\w*", r"rewind\w*", r"winding", r"winders?", r"unwind\w*", r"spool\w*", r"festoon\w*",
    r"matrix strip\w*", r"trim", r"inspection", r"registration", r"static", r"ioni[sz]\w+", r"sealing",
    r"welding", r"bag[- ]making", r"form[- ]fill[- ]seal", r"converting", r"converters?", r"bonding",
    r"thermoform\w*", r"wrapping", r"graphics", r"wrap films?", r"window films?",
    # web handling
    r"web guid\w*", r"tension\w*", r"dancers?", r"accumulators?", r"splic\w+", r"idlers?", r"rollers?",
    r"rolls", r"nip", r"spreaders?", r"chill rolls?", r"cores?", r"air shafts?", r"chucks?", r"turret",
    r"line speed", r"web width", r"working width", r"machines?", r"lines?",
    # other languages (Spanish, German, French, Chinese, Japanese, Korean)
    r"pel[ií]culas?", r"l[aá]minas?", r"papel", r"no tejidos?", r"tela no tejida", r"adhesivos?",
    r"recubrimientos?", r"laminaci[oó]n", r"etiquetas?", r"cintas?", r"impresi[oó]n", r"extrusi[oó]n",
    r"envases?", r"embalajes?", r"folien?", r"vlies\w*", r"papier\w*", r"klebstoff\w*", r"klebeb[aä]nd\w*",
    r"beschichtung\w*", r"kaschier\w*", r"etikett\w*", r"verpackung\w*", r"extrusion", r"films? plastiques?",
    r"薄膜", r"箔", r"无纺布", r"不织布", r"胶粘", r"涂布", r"复合", r"包装", r"标签", r"胶带",
    r"フィルム", r"不織布", r"粘着", r"接着", r"塗工", r"ラミネート", r"包装", r"ラベル", r"テープ", r"銅箔", r"圧延",
    r"필름", r"부직포", r"점착", r"접착",
]
TOPIC_RE = re.compile(r"(?<![a-z0-9])(" + "|".join(_TOPIC_TERMS) + r")(?![a-z0-9])", re.I)

# Documents about the company rather than its products - not wanted even when they name films or foils.
# Checked against the file name, PDF title and the opening of the text only (product sheets often end
# with "subject to our terms and conditions of sale", which must not count).
CORPORATE_ONLY_RE = re.compile(
    r"(integrated (sustainability )?report|annual report|sustainability report|esg (data|report)|"
    r"financial (results|statements)|balance sheet|securities report|corporate governance|"
    r"(medium|long)[- ]term (management plan|vision)|proxy statement|code of (business )?conduct|"
    r"business principles|guiding principles|values and principles|human rights|modern slavery|"
    r"privacy (policy|notice)|purchasing conditions|credit card payment|claim entry guide|"
    r"document retrieval guide|transit damage claim|"
    r"supplier (code|portal|registration|questionnaire|inquiry form|quality (requirements|manual))|cargo ?market|due diligence obligations|"
    r"procurement report|conflict minerals|whistleblow\w*|ethics (hotline|point)|ethicspoint|carbon markets|"
    r"co2 reduction strategy|ghg inventory|verification opinion|end user licen[cs]e|software licen[cs]e|"
    r"curriculum vitae|notification of registration|certificate of approval|certificate number|"
    r"attestation (statement|of conformance)|chain of custody|fsc.{0,3}coc|stakeholder|my ?mondi|"
    r"verpflichtungserkl\w+|grundsatzerkl\w+|"
    r"统合报告|統合報告|統合報告書|有価証券報告書|株主|决算|決算|貸借対照表|中期経営計画|長期ビジョン|"
    r"委託生産|各\s*位|ステークホルダー|ポリシー|方針)", re.I)

# Sika: industrial only (bonding, sealing, laminating, composites, automotive / marine / appliance assembly).
# Construction product lines - including roofing and waterproofing membranes - are off-topic.
SIKA_INDUSTRIAL_RE = re.compile(
    r"\b(automotive|aftermarket|auto ?glass|windshields?|windscreens?|direct glazing|body ?shop|vehicles?|"
    r"car bodies|bus(es)? and coach|truck bodies|trailers?|caravans?|recreational vehicles?|rv|"
    r"marine (deck|vessel|industry)|boat ?building|shipbuilding|yachts?|wind (blades?|turbines?)|"
    r"composite (parts|tooling|manufacturing|molds?)|tooling boards?|mold making|infusion|"
    r"laminating adhesives?|flexible packaging|textile lamination|hot[- ]?melts?|pressure[- ]sensitive|"
    r"industrial (assembly|bonding|manufacturing)|appliances?|white goods|oem|e[- ]?mobility|battery packs?|"
    r"panel bonding|sandwich panels?)\b", re.I)
SIKA_CONSTRUCTION_RE = re.compile(
    r"(concrete|mortar|grout|screed|flooring|\bfloors?\b|masonry|admixture|cementitious|cement|stucco|plaster|"
    r"\beifs\b|facade|fa[cç]ade|pavement|\bpavers?\b|basement|tunnel|bridge|parking|roofing|\broofs?\b|"
    r"waterproofing|tiles?|tiling|anchor\w*|rebar|joint seal\w*|expansion joints?|construction|building|"
    r"curtain wall|glazing|window|deck|sidewalk|driveway|countertop|swimming pool|hormig[oó]n|mortero|"
    r"pavimento|lechada|techo|cubierta|impermeabiliz\w*|construcci[oó]n|azulejo|baldosa|"
    r"sika\s*floor|sikafloor|sikacrete|sika\s*grout|sikagrout|sikatile|sika\s*tile|sarnafil|sarnacol|"
    r"sarnafast|sikaplan|sika\s*roof|sikaroof|sikalastic|sikagard|sikatop|sika\s*top|sikarepair|"
    r"sika\s*repair|sikaquick|sika\s*quick|sikaset|sikacem|sikaemaco|emaco|sika\s*level|sikalevel|"
    r"sikadur|armatec|monotop|viscocrete|plastiment|sikament|sikaplast|sika\s*rapid|sikacontrol|"
    r"sika\s*control|sika\s*air|sika\s*aer|sikafiber|sika\s*fiber|sikafume|sikapaver|separol|antisol|"
    r"sikaswell|sikafuko|waterbar|sika\s*inject|sikamur|sikawall|sika\s*wall|sikabond|sika\s*boom|"
    r"sikaflex\W{0,3}(1a|1 ?csl|2c|11 ?fc|15 ?lm|pro|construction|crack|concrete|mortar|self[- ]?level|"
    r"multi|textured|window|joint)|sikasil\W{0,3}(sg|ig|wt|ws|wr|c|gs|n)|parex|lahabra|la habra|merkrete|"
    r"senergy|sch[oö]nox|kingpak|king[- ]ms|king[- ]rs|\bbmi\b|everbuild|rmax|thermasheath|ecomaxci|emseal|"
    r"thorolastic|thoro|masterseal|mastertop|masterprotect|masteremaco|mbrace|ucrete|purcem|epocem|"
    r"thermabase|nailable base|water main|rut-mix|merlex|variance specialty)", re.I)

SIKA_INDUSTRIAL_LINE_RE = re.compile(
    r"(sika\W{0,2}force|sika\W{0,2}power|sika\W{0,2}melt|sikafast|sika\W{0,2}fast\W{0,3}\d|sika\W{0,2}tack|"
    r"biresin|sika\W{0,2}sense|sika\W{0,2}reinforcer|sika\W{0,2}baffle|sika\W{0,2}damp|sika\W{0,2}aktivator|"
    r"sika\W{0,2}activator|sika\W{0,2}transfer|sika\W{0,2}block|advanced resins|sikaflex\W{0,3}[25]\d\d|"
    r"sikaflex\W{0,3}(ws|p2g)|sikasil\W{0,3}[67]\d\d|primer\W{0,3}2\d\d|sika\W{0,2}handclean)", re.I)
SIKA_CONSTRUCTION_LINE_RE = re.compile(
    r"(sika\s*floor|sikafloor|sikacrete|sika\s*grout|sikagrout|sikatile|sika\s*tile|sarnacol|sarnafast|"
    r"sika\s*roof\s*fastener|sikalastic|sikagard|sikatop|sika\s*top|sikarepair|sika\s*repair|sikaquick|"
    r"sika\s*quick|sikaset|sikacem|sikaemaco|emaco|sika\s*level|sikalevel|sikadur|armatec|monotop|viscocrete|"
    r"plastiment|sikament|sikaplast|sika\s*rapid|sikacontrol|sika\s*control|sika\s*air\b|sika\s*aer|sikafiber|"
    r"sika\s*fiber|sikafume|sikapaver|separol|antisol|sikaswell|sikafuko|waterbar|sika\s*inject|sikamur|"
    r"sikawall|sika\s*wall|sikabond|sika\s*boom|sikaflex\W{0,3}(1a|1 ?csl|2c|11 ?fc|15 ?lm|pro\b|construction|"
    r"crack|concrete|mortar|self[- ]?level|multi|textured|window|joint)|sikasil\W{0,3}(sg|ig|wt|ws|wr|c\b|gs|n\b)|"
    r"parex|lahabra|la habra|merkrete|senergy|sch[oö]nox|kingpak|king[- ]ms|king[- ]rs|\bbmi\b|everbuild|"
    r"rmax|thermasheath|ecomaxci|emseal|thorolastic|thoro|masterseal|mastertop|masterprotect|masteremaco|"
    r"mbrace|ucrete|purcem|epocem|thermabase|nailable base|water main|rut-mix|merlex|variance specialty|"
    r"sarnafil|sarna\w+|sikaplan|sika\s*trocal|sikaproof|sika\s*roof|sikaroof|sikashield|sikawrap|carbodur|"
    r"sikagrind|sika\s*one\s*shot|loadflex|pebbletex|lithochrome|pulastic|dritac|drylok|sikaflow|ferrogard|"
    r"fibercast|hydrotite|sikacor|sikaseal|sika\s*film|sika\s*fleece|reemat|sikaflex\W{0,3}[79]\d\d|"
    r"sikaflex\W{0,3}1c|ultra-?max|onedek|rhinobond|retrodriller|\b(ms|rs|hc|fa|hp|c3)-[a-z]?\d)", re.I)

# sales paperwork: only when the document IS one (file name or its very first line), because product
# sheets often carry "see our terms and conditions" in a header or footer
PAPERWORK_TITLE_RE = re.compile(
    r"(terms (and|&) conditions|terms of (use|sale)|general (terms|conditions) of (purchase|sale|delivery)|"
    r"purchase orders?|invoic(e|es|ing)|order confirmations?)", re.I)

# plant / quality-system certificates, judged by file name only (data sheets often say "ISO 9001 certified")
CERT_NAME_RE = re.compile(r"(?<![a-z0-9])(as|iso|iatf)([\s_/-]?iec)?[\s_-]?(9100|9001|14001|27001|45001|50001|16949)(?![0-9])", re.I)

# Avery Dennison medical: liquids and creams are not webs (dressings, tapes and films are)
AVERY_NOT_WEB_RE = re.compile(r"(barrier film spray|\bspray\b|barrier cream|\bcream\b|\bswab\b|lotion|ointment)", re.I)


def topic_hits(text: str) -> int:
    """Number of distinct on-topic terms in the text."""
    return len({m.group(0).lower() for m in TOPIC_RE.finditer(text or "")})


def vendor_scope_decision(scope, names: str, where: str = "") -> str | None:
    """Per-vendor scope from vendor_seeds.csv (scope_keep / scope_exclude regexes, case-insensitive).
    Returns "on", an off-topic reason, or None when the scope says nothing about this file.
    names: file name + PDF title + first line(s); where: source URL and the page it was found on.
    1. scope_exclude matching the URL / found-on page path wins (a whole site section is out of scope),
       unless scope_keep matches the file's own name or title;
    2. scope_keep matching the name / title keeps the file;
    3. scope_exclude matching the name / title rejects it;
    4. scope_keep matching only the URL / found-on page keeps it."""
    if not scope:
        return None
    keep, exclude = scope
    if keep is None and exclude is None:
        return None
    paths = " ".join(re.sub(r"^https?://[^/]+", "", u) for u in (where or "").split())
    if exclude is not None and exclude.search(paths) and not (keep is not None and keep.search(names)):
        return "off-topic (vendor scope)"
    if keep is not None and keep.search(names):
        return "on"
    if exclude is not None and exclude.search(names):
        return "off-topic (vendor scope)"
    if keep is not None and keep.search(paths):
        return "on"
    return None


# ---- Company paperwork (audit 2026-10-24) -------------------------------------------------------
# Stock-exchange filings, purchasing / supplier paperwork, site rules, HR, policy papers and plant /
# chain-of-custody certificates are not about webs even when a web vendor publishes them.
# The two "strong" categories are matched on the name + first ~600 characters; the others only on the
# name + first ~200 characters and only when the document has few on-topic terms (<= 6), so product
# sheets that merely mention "ISO 9001" or a safety note are not caught.
PAPERWORK_RULES = {
 'stock-exchange / statutory filings': r"stock exchange|listing (department|regulations)|scrip code|shareholding pattern|shareholders? whose|corporate governance report|board meeting|outcome of (the )?board|trading window|postal ballot|scrutini[sz]er|agm proceedings|voting results|unclaimed dividend|unpaid dividend|\biepf\b|form (no\.? )?(mgt|isr|inc|sh|15g|15h|10f|iv-a)|loss of share|insider trading|demerge[dr]|scheme of arrangement|resulting company|credit rating|memorandum and articles|related party transactions|sebi|newspaper (publication|advertisement)|securities and exchange commission|form f-4|prospectus|תשקיף|דוח (הדירקטוריון|תקופתי|כספי|שנתי)|רבעון|demonstra[cç][oõ]es financeiras|balan[cç]o patrimonial|country[- ]by[- ]country|cbcr|income tax report",
 'purchasing / supplier paperwork': r"conditions of purchase|general (contract )?(terms|conditions) of purchase|conditions? g[ée]n[ée]rales? d.achats?|condiciones generales de compra|condi[cç][oõ]es gerais de compra|allgemeine (einkaufs|vertrags|liefer)\w*|einkaufsbedingungen|ostoehdot|købsbetingelser|kjøpsvilkår|köpvillkor|inkoopvoorwaarden|warunki zakup|nákupní podmínky|obchodné podmienky nákupu|beszerzési feltételei|achiziție|nabavki|satın alma|закупок|покупка|pirkimo|iepirkuma|ostutingimused|supplier (diversity|manual|training|performance training|instructions|information request|portal|standards|requirements|code)|coupa|ariba|source to pay|\bsnc\b|supply network collaboration|purchase order|government portal|vetting (platform|requirements)|change order protocol|master supplier agreement|tax[- ]exempt|pack list|credit card payment|returns processing|cyber security addendum|ai addendum|data act addendum|standardsoftware|bedingungen für die ausführung|contractor (rules|handbook|manual|requirements|presentation|safety)|kontraktoren|entrepren[oö]rsföreskrifter|vertragspartner|ayuda de trabajo|jobhilfe|pomoc stanowiskowa|supporto lavoro|ジョブエイド|サプライヤー",
 'site rules / safety / EHS': r"site regulations|standort-?ordnung|werkstra(ss|ß)enverkehrsordnung|betreten und befahren|werkschutz|safety (rules|training|regulations|guidance note)|säkerhetsinformation|safety information (for|when)|life saving rules|nine safety rules|task risk assessment|risk assessment (form|example)|\blockout|burner management|process safety management|emergency response|fire water|scaffold|angle grinding|high pressure (water )?(washing|cleaning)|she (plan|standards)|ehs (requirements|visitor)|sicherheitsunterweisung|säkerhet\w*|skydds|ordnings|råvaruhantering|bulkfartyg|anmälningsplikt|instrukcja w sprawie|szkolenie bhp|störfallverordnung|seveso|umweltinspektion|informationen für unsere nachbarn|information till allmänheten|stoppextra|selbstabholer|customer-?organi[sz]ed collections|adr kontrolle",
 'HR / people': r"gender pay|pay (gap|transparency)|wage report|דוח פומבי|דיווח פומבי|égalité hommes|index égalité|mental health|burnout|substance abuse|inclusive workplace|sexual harassment|retirees|pensioners|emeryt|employee (training|benefits)|values (and|&) leadership|principles about teamwork|whistleblow\w*|procedura zgłoszeń|complaints procedure|beschwerdeverfahren|cybersecurity all employees",
 'compliance / policy / position papers': r"code of conduct|c[oó]digo de conduct?a|modell?o (di organizzazione|of organi[sz]ation|organizzativo)|modello 231|compliance guideline|conflicts? of interest|human rights|menschenrecht\w*|statement of principles|stewardship framework|position (on|paper)|key messages|green deal|emission trading|eu ai act|industrial emission directive|transition pathway|biodiversity position|recommendations (on|for) (a )?circular economy|california transparency|deforestation regulation|eudr|politique (qhse|environnementale)|environmental policy|fibre procurement policy|polityka (zakupu|jakości)|politique qhse|qhse policy|politika společnosti|wood (sourcing|procurement)|program pracy sieci|sieci (ciepłowniczej|dystrybucyjnej)|taryfa|struktura paliw|czas trwania przerw|autoconsumo|co2 reduction|climate transition|plano de transi[cç][aã]o",
 'certificates (plant / chain-of-custody)': r"certificate of (registration|approval|conformance)|certificado de cadena de custodia|chain of custody|\bfsc\b.{0,40}certif|pefc|sfi certif|fiber sourcing|rainforest alliance|greencircle|kosher certificate|tisax|letter of conformance|responsible care|oea|authori[sz]ed economic operator|fssc 22000|c-?gmp certificate|organic content standard|cradle to cradle certif|vo[cç] emission test certificate|endorsement letter|preferred reed supplier|zertifikat|certifikát|osvedčenie|certyfikat|ecovadis|lab report|this certificate|certificate holder|certification (body|decision|scheme)|brcgs|\bbrc\b.{0,30}certif|halal|haccp|\bgmp\b.{0,20}(certif|aib)|\bmsc\b.{0,20}certif|iso.?certification|attestazione soa|ghg verification|温室气体核查",
 'news / HR / corporate misc': r"press release|presseerkl[aä]rung|news release|obituary|appoints new|team updates|^stelle\b|stellenangebot|job (offer|opening|posting)|gender (gap|pay)|due diligence|anti-?bribery|anti-?corruption|business gifts|hospitality procedures|global directory|request for information|eu entities|european entities|cartel (idi|publicidad)|descripcion ?proyecto|open source|license conditions|apache license|bsd-[23]-clause|\bmit\b.{0,20}license|postgresql licen[cs]e|sqlite copyright|how to find|ship notice|bank details|credit memo|forwarding agency|classroom activity|family activity|scholarship|employee assistance|supplemental insurance",
}
PAPERWORK_RES = {k: re.compile(v, re.I) for k, v in PAPERWORK_RULES.items()}
PAPERWORK_STRONG = {"stock-exchange / statutory filings", "purchasing / supplier paperwork"}
# product certificates are kept: EU type-examination (PPE), declarations of performance / conformity
PAPERWORK_CERT_EXEMPT_RE = re.compile(r"type[- ]examination|declaration of (performance|conformity)", re.I)


def paperwork_category(name: str, text: str = "") -> str | None:
    """Category of company paperwork the file belongs to, or None."""
    ns = re.sub(r"[_\-]+", " ", name or "")
    t = text or ""
    strong_zone, zone = f"{ns} {t[:600]}", f"{ns} {t[:200]}"
    hits = topic_hits(t[:15000])
    for cat, rx in PAPERWORK_RES.items():
        if cat in PAPERWORK_STRONG:
            if rx.search(strong_zone):
                return cat
        elif hits <= 6 and rx.search(zone):
            if cat.startswith("certificates") and PAPERWORK_CERT_EXEMPT_RE.search(strong_zone):
                continue
            return cat
    return None


def topic_decision(vendor: str, name: str, text: str, keep_all_vendors: set | None = None,
                   scope=None, where: str = "") -> str | None:
    """None = on-topic (keep). Otherwise a short reason the file is off-topic.
    vendor: vendor name; name: file stem + PDF title; text: first pages of extracted text;
    scope: (scope_keep, scope_exclude) from vendor_seeds.csv; where: source URL + found-on page."""
    t = (text or "")[:15000]
    head = f"{name} {t[:1500]}"
    first_line = next((ln for ln in t.splitlines() if ln.strip()), "")[:120]
    name_sp = re.sub(r"[_\-]+", " ", name)            # 'Supplier-Quality-Manual' -> 'Supplier Quality Manual'
    if CORPORATE_ONLY_RE.search(f"{name_sp} {t[:400]}") or CERT_NAME_RE.search(name) or \
            PAPERWORK_TITLE_RE.search(f"{name_sp} {first_line}") or \
            re.search(r"noti\s*fi\s*cati\s*on of regi\s*strati\s*on", t[:400], re.I):
        return "off-topic (corporate)"
    pc = paperwork_category(name, t)
    if pc:
        return f"off-topic (paperwork: {pc})"
    vs = vendor_scope_decision(scope, f"{name} {t[:300]}", where)
    if vs == "on":
        return None
    if vs:
        return vs
    if keep_all_vendors and vendor in keep_all_vendors:
        return None
    v = (vendor or "").lower()
    if v.startswith("sika"):
        # Industrial Sika only (laminating / hot-melt / PSA / bonding and sealing for vehicles, marine,
        # appliances, composites). Construction - including roofing and waterproofing membranes - is out.
        title = f"{name} {t[:300]}"
        if SIKA_INDUSTRIAL_LINE_RE.search(title):
            return None                                   # SikaForce, SikaMelt, Sikaflex-2xx ... by name
        if SIKA_CONSTRUCTION_LINE_RE.search(title):
            return "off-topic (construction)"             # Sikafloor, Sikacrete, Sarnafil, Parex ... by name
        ind = {m.group(0).lower() for m in SIKA_INDUSTRIAL_RE.finditer(t[:6000] + " " + name)}
        con = len(SIKA_CONSTRUCTION_RE.findall(head))
        if ind and (len(ind) >= 2 or con == 0):
            return None
        return "off-topic (construction)" if con else "off-topic"
    if v.startswith("avery") and AVERY_NOT_WEB_RE.search(name) and not re.search(
            r"(dressing|tape|film(?! spray)|foam|nonwoven|label)", name, re.I):
        return "off-topic (not a web product)"
    # every other vendor in the workbook supplies webs, web materials or web machinery: their product
    # documents relate to the topics even when the wording is terse ("EM-370Z LAM") or not in English
    return None


def link_off_topic(vendor: str, url: str, anchor: str = "", scope=None, found_on: str = "") -> bool:
    """Before download: skip links whose file name / link text alone shows they are off-topic
    (company paperwork, plant certificates, Sika construction product lines). Anything unclear is
    downloaded and judged on its content."""
    name = f"{Path(unquote(urlparse(url).path)).stem} {anchor}"
    name_sp = re.sub(r"[_\-]+", " ", name)
    if CORPORATE_ONLY_RE.search(name_sp) or CERT_NAME_RE.search(name) or PAPERWORK_TITLE_RE.search(name_sp) \
            or paperwork_category(name):
        return True
    vs = vendor_scope_decision(scope, name, f"{url} {found_on}")
    if vs:
        return vs != "on"
    v = (vendor or "").lower()
    if v.startswith("sika"):
        return bool(SIKA_CONSTRUCTION_LINE_RE.search(name)) and not SIKA_INDUSTRIAL_LINE_RE.search(name)
    if v.startswith("avery"):
        return bool(AVERY_NOT_WEB_RE.search(name)) and not re.search(
            r"(dressing|tape|film(?! spray)|foam|nonwoven|label)", name, re.I)
    return False


# ---- Allowlist: a file is kept only with positive evidence of technical content ----------------
UNIT_PATTERNS = {
    "g/m²": r"\bg\s*/\s*m\s*[²2]|\bgsm\b|\bg/sq\.?\s*m",
    "µm": r"\d\s*(µm|μm|um|microns?)\b",
    "mm": r"\d\s*mm\b",
    "mil": r"\d\s*mils?\b",
    "MPa/psi": r"\b(MPa|GPa|kPa|psi|ksi|kgf\s*/\s*cm)\b",
    "N/width": r"\bN\s*/\s*(\d+\s*)?(mm|cm|in|m)\b|\blbf?\s*/\s*in\b|\bpli\b|\bgf\s*/",
    "force": r"\d\s*(N|cN|kN|lbf)\b",
    "°C/°F": r"°\s*[CF]\b|\bdeg\.?\s*[CF]\b",
    "%": r"\d\s*%",
    "density": r"(g\s*/\s*cm\s*[³3]|kg\s*/\s*m\s*[³3]|g/cc\b|lb\s*/\s*ft\s*[³3])",
    "melt flow": r"(g\s*/\s*10\s*min|cm\s*[³3]\s*/\s*10\s*min)",
    "electrical": r"\b(GHz|MHz|kV|V/mil|kV/mm|ohm|Ω|ohm-cm|Dk|Df)\b",
    "thermal": r"\bW\s*/\s*m\s*[·.]?\s*K\b|\bppm\s*/\s*°?\s*C\b|\bCTE\b|\bTg\b",
    "barrier": r"\b(cc|cm3|cm³|g)\s*/\s*(m\s*[²2]|100\s*in\s*[²2])\s*[·.]?\s*(day|24\s*h)",
    "air perm.": r"(l\s*/\s*m\s*[²2]\s*/?\s*s\b|\bcfm\b|Gurley|s\s*/\s*100\s*ml)",
    "oz/inch": r"\d\s*(oz|ounces?)\b|\d\s*(\"|in\.|inch(es)?)(\s|$)",
}
UNIT_RES = {k: re.compile(v) for k, v in UNIT_PATTERNS.items()}
PROPERTY_RE = re.compile(
    r"(tensile|elongation|thickness|caliper|basis weight|grammage|areal weight|density|modulus|tear|burst|"
    r"porosity|permeability|peel|adhesion|dielectric|dissipation|loss tangent|thermal conductivity|viscosity|"
    r"melt (flow|index)|haze|gloss|shrink|absorb|coefficient of friction|\bcof\b|puncture|elmendorf|opacity|"
    r"brightness|moisture|wvtr|otr|heat seal|flammab|conductivity|hardness|shore|resistivity|"
    r"yield (strength|stress|point)|impact (strength|resistance)|\belastic modulus|stiffness|softness|"
    r"bulk density|filtration efficiency|pressure drop|penetration|surface resistance|\btack\b|"
    r"shear|release force|coat weight|roughness|annealing|\btemper(s|ed)?\b|purity|roll length|construction|resin content|"
    r"glass style|copper weight|prepreg|laminate|\bply\b|compression|durometer|cell size|fib(er|re)|"
    r"softening point|melting (point|temperature)|\bvicat\b|deflection temperature|\bhdt\b|glass transition|"
    r"crystallization|\bcolou?r\b|yellowness|specific gravity|flash point|molecular weight|mooney|\bash\b|"
    r"volatiles?|comonomer|izod|charpy|flexural|dart (impact|drop)|secant|melt strength|\bmfr\b|\bmfi\b|"
    r"拉伸|伸长|厚度|密度|介电|损耗|导热|热导|剥离|强度|硬度|性能|物性|测试方法|试验方法|典型值|规格|規格|特性|"
    r"引張|伸び|厚さ|比誘電率|誘電正接|熱伝導|はく離|剥離|強度|硬さ|代表値|인장|두께|밀도|유전|강도|물성|"
    r"zugfestigkeit|dicke|dichte|dehnung|flächengewicht|épaisseur|densité|allongement)", re.I)
HEADER_RE = re.compile(
    r"(technical data|data sheet|datasheet|typical (values|properties)|product data|material properties|"
    r"test method|test based on|typical value|physical properties|mechanical properties|electrical properties|thermal properties|"
    r"grade\s+thickness|technische daten|技术数据|技術データ|数据表|データシート|物性表|典型值|代表値|"
    r"technisches datenblatt|fiche technique)", re.I)
_B, _E = r"(?:^|[\W_])", r"(?:$|[\W_])"
# strong: the name alone says it is a technical document
NAME_STRONG_RE = re.compile(_B + r"(tds|pds|data[\s_-]*sheets?|datasheets?|spec(ification)?s?|grades?|properties|"
                            r"technical|product[\s_-]*(data|guide|overview|selector|information)|selector|"
                            r"selection[\s_-]*guide|fabrication|processing|guidelines?|bulletin|brochure|flyer|"
                            r"catalog(ue)?|leaflet|sell[\s_-]*sheet|application[\s_-]*note|line[\s_-]*card|fact[\s_-]*sheet|"
                            r"factsheet|brosch(ü|ue|u)re|datenblatt|prospekt|technische[\s_-]*daten|"
                            r"fiche[\s_-]*technique|ficha[\s_-]*t[ée]cnica)" + _E, re.I)
# weak: a material word - only counts together with some technical text
NAME_MATERIAL_RE = re.compile(_B + r"(laminates?|lam|prepregs?|films?|foils?|nonwovens?|papers?|tapes?|"
                              r"adhesives?|resins?|fab(ric)?s?|webs?|membranes?|coatings?)" + _E, re.I)
NAME_JUNK_RE = re.compile(_B + r"(logos?|curriculum|cv|resume|vcard|business[\s_-]*principles|"
                          r"code[\s_-]*of[\s_-]*conduct|polic(y|ies)|privacy|terms[\s_-]*(and|&)[\s_-]*conditions|"
                          r"agm|general[\s_-]*meeting|board[\s_-]*of[\s_-]*directors|remuneration|nomination|"
                          r"articles[\s_-]*of[\s_-]*association|org(ani[sz]ation(al)?)?[\s_-]*chart|calendar|"
                          r"invitation|agenda|press[\s_-]*releases?|news|annual[\s_-]*report|sustainab\w*|esg|"
                          r"ethics|hotline|careers?|jobs?|pfas|conflict[\s_-]*minerals|ghg|greenhouse|lca|life[\s_-]*cycle[\s_-]*(assessment|analysis)|"
                          r"verification[\s_-]*opinion|letter|declaration|disclosure|testimonials?|"
                          r"success[\s_-]*stor(y|ies)|case[\s_-]*stud(y|ies)|environmental[\s_-]*(bulletin|statement|compliance)|"
                          r"invoic\w*|order[\s_-]*confirmations?|purchas\w*|delivery[\s_-]*(guidelines|conditions|"
                          r"instructions)|supplier[\s_-]*(inquiry|form|code|portal)|cargo\w*|logistics|shipping|"
                          r"registration|human[\s_-]*rights|grundsatzerkl\w*|warengruppe\w*|einkauf\w*|liefer\w*|"
                          r"terms[\s_-]*(of|and)[\s_-]*\w+)" + _E, re.I)
MACHINE_RE = re.compile(
    r"(\d\s*(m/min|mpm|ft/min|kg/h|kg/hr|lb/h|t/h|kW|rpm|1/min|min-1)\b|line speed|working width|web width|"
    r"net width|throughput|screw diameter|\bL/D\b|die width|winding (speed|diameter)|max\.? (speed|width)|"
    r"roll diameter|unwind|rewind|slitting width|production speed|output rate)", re.I)
# equipment / component vocabulary - product documents of machine makers
EQUIP_RE = re.compile(
    r"(extruders?|screws?|barrels?|flat dies?|die heads?|die lips?|feedblocks?|coextrusion|spinnerets?|"
    r"spinning beam|meltblown|spunbond|spunlace|winders?|rewinders?|unwinders?|slitters?|laminators?|coaters?|"
    r"coating heads?|calenders?|tenter|stretching|chill rolls?|corona|web guid\w*|tension control|"
    r"cast (film|sheet) lines?|blown film lines?|sheet lines?|nonwoven lines?|mdo|gear pumps?|screen changers?|"
    r"melt pumps?|thickness gauges?|printing (units?|press)|flexo|gravure|slot[- ]die|"
    r"extruder|schnecke|zylinder|düse|duese|wickler|kaschier|beschicht|"
    r"挤出机|螺杆|模头|喷丝板|收卷|分切|涂布|复合机|流延|吹膜)", re.I)
MIN_TECH_SCORE = 3
STRONG_TECH_SCORE = 15      # this much evidence = a real data sheet, even if its small print names sales terms
LOW_TEXT_CHARS = 400        # less text than this = mostly images; the name may decide


def decode_mt(text: str) -> str:
    """Some PDFs extract as '/MT78/MT101/MT119...' (decimal character codes) - decode them."""
    if not text or text.count("/MT") < 20:
        return text
    return re.sub(r"/MT(\d{2,3})", lambda m: chr(int(m.group(1))) if 32 <= int(m.group(1)) < 1200 else " ", text)


def technical_evidence(text: str) -> tuple[int, str]:
    """Score how much the text looks like a technical document, plus a short human-readable reason."""
    t = unicodedata.normalize("NFKC", text or "")[:15000]
    methods, _ = extract_methods(t)
    units = sorted(k for k, rx in UNIT_RES.items() if rx.search(t))
    props = sorted({m.group(0).lower() for m in PROPERTY_RE.finditer(t)})
    head = bool(HEADER_RE.search(t))
    machine = sorted({m.group(0).lower() for m in MACHINE_RE.finditer(t)})
    equip = sorted({m.group(0).lower() for m in EQUIP_RE.finditer(t)})
    score = (2 * len(methods) + len(units) + min(len(props), 6) + (2 if head else 0) + min(len(machine), 4)
             + min(len(equip), 3))
    # substance: real data, not just a stray unit or a technical-sounding word
    substantive = (bool(methods) or len(props) >= 3 or (len(props) >= 2 and (units or head))
                   or len(machine) >= 2 or (len(props) >= 1 and len(units) >= 2)
                   or (head and len(props) >= 1 and bool(units))            # short sheet: heading + a value table
                   or (len(equip) >= 3 and (units or machine or props)))   # machine / component product docs
    if not substantive:
        score = min(score, MIN_TECH_SCORE - 1)
    parts = []
    if methods:
        parts.append(f"{len(methods)} test method(s)")
    if units:
        parts.append("units: " + ", ".join(units[:6]))
    if props:
        parts.append("properties: " + ", ".join(props[:5]))
    if machine:
        parts.append("machine specs: " + ", ".join(machine[:4]))
    if equip:
        parts.append("equipment: " + ", ".join(equip[:4]))
    if head:
        parts.append("data-sheet heading")
    return score, "; ".join(parts)


def decide_keep(f, doc_types: set, equipment_vendor: bool = False) -> str | None:
    """None = keep, otherwise the reason it is rejected. The TOPIC decides first: anything that relates to
    the corpus topics is kept (data sheet, SDS, brochure, bulletin...), anything off-topic is rejected.
    Files whose topic could not be judged (PDF unreadable) fall back to the technical-evidence rules below,
    where a technical file name only helps when the PDF has (almost) no text. f needs path, pdf_title,
    text_doc_type, tech_score, text_chars, language, english_path, topic."""
    topic = f["topic"]
    if topic == "on":
        return None
    if topic:
        return topic                                    # "off-topic", "off-topic (construction)" ...
    dtype = f["text_doc_type"] or "other"
    score = f["tech_score"] or 0
    stem = Path(f["path"]).stem
    names = f"{stem} {f['pdf_title'] or ''}"
    if NAME_JUNK_RE.search(stem) and score < 8:
        return "not technical (name)"
    if dtype == "commercial" and score >= STRONG_TECH_SCORE:
        dtype = "tds"          # a full property table outweighs one sales phrase in the small print
    if dtype in ("corporate", "commercial"):
        return dtype
    if dtype not in doc_types and dtype != "unreadable":
        return dtype                                    # e.g. sds / certificate, or types left out of --doc-types
    twin = bool(f["english_path"])                      # non-English copy of a kept English document
    if twin:
        return None
    low_text = dtype == "unreadable" or (f["text_chars"] or 0) < LOW_TEXT_CHARS
    if score >= MIN_TECH_SCORE:
        return None
    strong = bool(NAME_STRONG_RE.search(names))
    if equipment_vendor:
        # Machine makers: the brochure IS the technical document. Anything that talks about their
        # machines/components (and already passed the commercial / corporate / junk-name checks) stays.
        ev = f["evidence"] or ""
        if "equipment:" in ev or "machine specs:" in ev or strong or name_lang(f["path"]):
            return None
    if low_text:
        return None if strong else "no readable text"
    foreign = bool(f["language"]) and not _is_english(f["language"])
    # commercial papers and junk names were already rejected above, so a technical name plus some
    # technical text (or CJK text the evidence words may not cover) is enough
    if strong and (score >= 1 or foreign):
        return None
    return "no technical content"


def vendor_scopes(vendors: dict) -> dict:
    """{vendor name: (scope_keep, scope_exclude)} for vendors whose seed row sets a scope."""
    return {v.name: (v.scope_keep, v.scope_exclude) for v in vendors.values()
            if v.scope_keep is not None or v.scope_exclude is not None}


def equipment_vendors(vendors: dict) -> set:
    """Vendors listed on the Equipment & Machinery or Process Technology sheets (machine makers)."""
    out = set()
    for v in vendors.values():
        cats = " ".join(v.categories).lower() + " " + " ".join(v.category_folders).lower()
        if "equipment" in cats or "machinery" in cats or "process technology" in cats or "process_technology" in cats:
            out.add(v.name)
    return out


def evidence_label(f) -> str:
    """What the catalog shows in the Evidence column."""
    if f["evidence"]:
        return f["evidence"]
    if f["rejected"]:
        return ""
    if f["english_path"]:
        return "non-English copy of a kept English document"
    names = f"{Path(f['path']).stem} {f['pdf_title'] or ''}"
    if NAME_STRONG_RE.search(names):
        return ("image-only PDF; " if f["text_doc_type"] == "unreadable" else "") + "technical file name"
    return ""


def reject_unwanted(db: DB, doc_types: list[str], equipment_vendors: set | None = None) -> int:
    """ALLOWLIST: every downloaded file must show technical evidence (test methods, property values with
    units, property names, data-sheet headings - or a technical file name when the text is unreadable).
    Everything else goes to _rejected/ with the reason - never deleted."""
    wanted = set(doc_types)
    n = 0
    for f in db.query("SELECT sha256, path, title, pdf_title, text_doc_type, tech_score, text_chars, language, "
                      "english_path, rejected, evidence, vendor, topic FROM files WHERE text_doc_type IS NOT NULL"):
        reason = decide_keep(f, wanted, equipment_vendor=f["vendor"] in (equipment_vendors or set()))
        if (f["rejected"] or None) != reason:
            db.execute("UPDATE files SET rejected=? WHERE sha256=?", (reason, f["sha256"]))
            n += 1 if reason else 0
    if n:
        log.info("Rejected %d file(s) that are off-topic or have no usable content", n)
    return n


def delete_off_topic(db: DB, root: Path) -> int:
    """Delete files judged off-topic and remember their links as deleted, so they are never downloaded
    again. Each deletion is logged to logs/deleted_off_topic_YYYYMMDD.csv with its source URL."""
    rows = db.query("SELECT sha256, path, vendor, first_url, topic FROM files WHERE topic LIKE 'off-topic%'")
    if not rows:
        return 0
    if not (root / "pdfs").exists():
        log.warning("pdfs/ folder not found - is the drive connected? Nothing was deleted.")
        return 0
    out = root / "logs" / f"deleted_off_topic_{datetime.now():%Y%m%d}.csv"
    out.parent.mkdir(parents=True, exist_ok=True)
    new = not out.exists()
    n = 0
    with open(out, "a", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        if new:
            w.writerow(["deleted_at", "vendor", "reason", "path", "source_url", "sha256"])
        for r in rows:
            p = root / r["path"]
            try:
                if p.exists():
                    p.unlink()
            except OSError as e:
                log.warning("Could not delete %s: %s", p, e)
                continue
            db.execute("UPDATE pdf_links SET state='deleted' WHERE sha256=?", (r["sha256"],))
            db.execute("DELETE FROM files WHERE sha256=?", (r["sha256"],))
            db.execute("UPDATE files SET english_path=NULL WHERE english_path=?", (r["path"],))
            w.writerow([utcnow(), r["vendor"], r["topic"], r["path"], r["first_url"], r["sha256"]])
            n += 1
    log.info("Deleted %d off-topic file(s) - listed in %s (use --keep-off-topic to keep them in _rejected/)",
             n, out.name)
    return n


def _is_english(lang: str) -> bool:
    return lang == "en" or lang.endswith("+en")


def _file_is_foreign(f) -> bool:
    lang = f["language"] or ""
    if name_lang(f["path"]) and lang != "en":
        return True    # named '- CN' etc. and not purely English text (or no readable text)
    return bool(lang) and not _is_english(lang)


def pair_translations_by_text(db: DB) -> int:
    """Pair every non-English file with its English twin on disk (same vendor, matching name once the
    language marker is removed): 'X - CN.pdf' + 'X.pdf', 'C1990_data.pdf' + 'C1990_data_english.pdf'."""
    n = 0
    files = db.query("SELECT sha256, path, vendor, language, is_translation, english_path, rejected FROM files")
    english: dict = {}
    for f in files:
        if _is_english(f["language"] or "") and not name_lang(f["path"]):
            stem = COLLISION_SUFFIX_RE.sub("", Path(f["path"]).stem)
            english.setdefault((f["vendor"], match_key(stem)), f)
            # 'RNK_Ce.pdf' / 'MHSe.pdf' = English copy of 'RNK-C.pdf' / 'MHS.pdf' (lower-case e after a code)
            m = ENGLISH_E_SUFFIX_RE.search(stem)
            if m:
                english.setdefault((f["vendor"], match_key(stem[:m.start()])), f)
    for f in files:
        if f["english_path"] or not _file_is_foreign(f):
            continue
        en = english.get((f["vendor"], foreign_key(f["path"])))
        if en is not None and en["sha256"] != f["sha256"]:
            db.execute("UPDATE files SET is_translation=1, english_path=? WHERE sha256=?", (en["path"], f["sha256"]))
            db.execute("UPDATE pdf_links SET state='translation', english_url=(SELECT first_url FROM files WHERE "
                       "sha256=?) WHERE sha256=? AND state='downloaded'", (en["sha256"], f["sha256"]))
            n += 1
    if n:
        log.info("Paired %d non-English file(s) with their English version", n)
    return n


def mark_non_english(db: DB) -> tuple[int, int]:
    """Main folder = English only. Every non-English document goes to _translations/, whether or not an
    English version was found (english_path tells which). Returns (non-English, with English version)."""
    total = with_en = 0
    for f in db.query("SELECT sha256, path, language, is_translation, english_path, rejected FROM files"):
        foreign = _file_is_foreign(f)
        if foreign and not f["is_translation"]:
            db.execute("UPDATE files SET is_translation=1 WHERE sha256=?", (f["sha256"],))
        elif not foreign and f["is_translation"] and _is_english(f["language"] or ""):
            db.execute("UPDATE files SET is_translation=0, english_path=NULL WHERE sha256=?", (f["sha256"],))
        if foreign and not f["rejected"]:
            total += 1
            with_en += 1 if f["english_path"] else 0
    return total, with_en


def queue_english_guesses(db: DB, vendors: dict[str, Vendor], only: set | None = None) -> list[Vendor]:
    """For non-English files still without an English twin, queue guessed English URLs to try.
    Returns the vendors that got new guesses."""
    by_name = {v.name: v for v in vendors.values()}
    hit: dict[str, Vendor] = {}
    for f in db.query("SELECT f.vendor, f.first_url, l.doc_type FROM files f LEFT JOIN pdf_links l ON "
                      "l.url = f.first_url WHERE f.is_translation=1 AND f.english_path IS NULL AND "
                      "(f.rejected IS NULL OR f.rejected IN ('no technical content', 'no readable text'))"):
        url, v = f["first_url"], by_name.get(f["vendor"])
        if only is not None and f["vendor"] not in only:
            continue
        if not url or v is None or db.query("SELECT 1 FROM pdf_links WHERE english_of=? LIMIT 1", (url,)):
            continue
        for g in english_url_guesses(url):
            db.execute("INSERT OR IGNORE INTO pdf_links(url, vendor, found_on, anchor, doc_type, first_seen, lang, "
                       "english_of) VALUES (?,?,?,?,?,?, '', ?)",
                       (g, v.name, "english-guess", "", f["doc_type"] or "other", utcnow(), url))
        hit[v.name] = v
    return list(hit.values())


CONTROL_CHARS_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")   # not allowed in Excel cells


def unshift_text(text: str) -> str:
    """Some PDFs use fonts whose character codes are shifted by a constant (often 29 or 31), so the
    extracted text reads '.BSLFU0WFSWJFX' instead of 'Market Overview'. If shifting it back gives far
    more English words, use the shifted-back text."""
    if not text or re.search(r"\b(the|and|of|for|with)\b", text[:3000], re.I):
        return text
    # real German / French / Spanish / Italian / Dutch text has no English words but is NOT shifted
    if re.search(r"\b(der|die|und|ist|mit|von|f[üu]r|eine?|les?|la|des|et|est|pour|el|los|las|y|con|para|"
                 r"il|di|che|per|het|een|van|voor)\b", text[:3000], re.I):
        return text
    sample = re.sub(r"\s", "", text[:4000])
    if not sample or sum(1 for c in sample if ord(c) < 128) < 0.95 * len(sample):
        return text  # CJK / accented text - shifted-font garbage is plain ASCII
    letters = [c for c in sample if c.isalpha()]
    lower_share = sum(c.islower() for c in letters) / max(1, len(letters))

    def hits(t: str) -> int:
        return sum(1 for w in re.findall(r"[a-z]+", t[:4000].lower()) if w in ENGLISH_STOPWORDS)

    base, best, best_txt = hits(text), 0, text
    for k in (29, 31, 3, 1):
        # a 29/31 shift turns lowercase letters into capitals and punctuation, so genuine garbage of that
        # kind has almost no lowercase letters - normal text in any Latin language is mostly lowercase
        if k in (29, 31) and lower_share > 0.3:
            continue
        cand = "".join(chr(ord(c) + k) if 33 <= ord(c) <= 126 - k else c for c in text)
        h = hits(cand)
        if h > best:
            best, best_txt = h, cand
    return best_txt if best >= 5 and best > 3 * base else text


SUMMARY_SKIP_RE = re.compile(
    r"(copyright|©|all rights reserved|www\.|https?://|@|\btel\b|\bfax\b|phone|page \d+ of \d+|^\d+ of \d+$|"
    r"^\d[\d\s.,/-]*$|street|road|avenue|\bsuite\b|\bzip\b|confidential|the information contained|"
    r"disclaimer|effective date|^ver\.?\s*\d|^rev\.?\s*\d|^version|^updated|^issued|^date\b|"
    r"^[※*†]|data in this|representative value|not guaranteed|typical values?\b.*\bnot\b|for reference only|"
    r"subject to change|without notice)", re.I)


def make_summary(text: str, filename: str) -> str:
    """One-line description from the first meaningful lines of the PDF (or the filename)."""
    picked: list[str] = []
    for line in text.splitlines():
        ln = re.sub(r"\s+", " ", line).strip(" •·●○■□◆▪-–|*")
        if len(ln) < 12 or SUMMARY_SKIP_RE.search(ln):
            continue
        if sum(ch.isalpha() for ch in ln) < len(ln) * 0.5 and not re.search(r"[\u3040-\u9fff\uac00-\ud7af]", ln):
            continue  # mostly numbers / table cells
        if picked and ln.lower() in picked[-1].lower():
            continue
        if picked and ln[:1].islower() and not picked[-1].endswith((".", ":", ";")):
            picked[-1] = f"{picked[-1]} {ln}"   # same sentence wrapped onto the next line
        else:
            if len(picked) == 2:
                break
            picked.append(ln)
        if len(" - ".join(picked)) > 110:
            break
    out = " - ".join(picked)
    if not out:
        out = re.sub(r"[_\-]+", " ", Path(filename).stem).strip()
    return CONTROL_CHARS_RE.sub(" ", out)[:180]


def classify_files(db: DB, root: Path, redo: bool = False, doc_types=None, keep_rejected: bool = False,
                   equipment_vendors: set | None = None, keep_off_topic: bool = False,
                   scopes: dict | None = None) -> int:
    try:
        from pypdf import PdfReader
    except ImportError:
        log.warning("pypdf not installed - skipping classification (pip install pypdf)")
        return 0
    logging.getLogger("pypdf").setLevel(logging.ERROR)
    try:
        import pypdf
        log.info("Text extraction: pypdf %s", getattr(pypdf, "__version__", "?"))
    except Exception:
        pass
    scopes = scopes or {}
    rows = db.query("SELECT f.sha256, f.path, f.vendor, f.first_url, l.found_on, l.doc_type AS link_type FROM files f "
                    "LEFT JOIN pdf_links l ON l.url = f.first_url" +
                    ("" if redo else " WHERE f.classified_at IS NULL OR f.language IS NULL OR f.summary IS NULL"
                                    " OR f.tech_score IS NULL OR f.text_chars IS NULL"
                                    f" OR COALESCE(f.rules_version, 0) <> {RULES_VERSION}"))
    n = 0
    unread: list[str] = []
    for row in rows:
        if STOP.is_set():
            break
        p = root / row["path"]
        doing(f"classifying {row['path']} ({n + 1} of {len(rows)})")
        title, pages, text = "", None, ""
        rd = None
        read_ok = True
        try:
            rd = PdfReader(str(p))
            pages = len(rd.pages)
            for pg in rd.pages[:6]:
                text += (pg.extract_text() or "") + "\n"
            text = unshift_text(decode_mt(CONTROL_CHARS_RE.sub(" ", text)))
            meta_title = (rd.metadata.title if rd.metadata else None) or ""
            first_line = next((l.strip() for l in text.splitlines() if len(l.strip()) > 3), "")
            title = (meta_title if len(meta_title.strip()) > 3 else first_line)[:200]
        except Exception as e:
            log.debug("classify %s: %s", p, e)
            read_ok = False
            unread.append(str(e))
        methods, fams = extract_methods(text)
        meta_t = ""
        try:
            meta_t = (rd.metadata.title or "") if rd.metadata else ""
        except Exception:
            pass
        dtype = classify_text_doc_type(text, row["link_type"] or "", len(methods),
                                       hints=f"{meta_t} {Path(row['path']).stem}")
        lang = detect_language(text)
        summary = make_summary(text, row["path"])
        tscore, evid = technical_evidence(text)
        # topic: "on", an off-topic reason, or NULL when the PDF could not be opened (never delete what we
        # could not read - e.g. encrypted PDFs need the 'cryptography' package)
        topic = None
        if read_ok:
            topic = topic_decision(row["vendor"], f"{Path(row['path']).stem} {meta_t}",
                                   unicodedata.normalize("NFKC", text), TOPIC_KEEP_ALL_VENDORS,
                                   scope=scopes.get(row["vendor"]),
                                   where=f"{row['first_url'] or ''} {row['found_on'] or ''}") or "on"
        db.execute("UPDATE files SET pages=?, title=?, text_doc_type=?, test_methods=?, method_families=?, "
                   "language=?, summary=?, tech_score=?, evidence=?, pdf_title=?, text_chars=?, rules_version=?, "
                   "topic=?, classified_at=? WHERE sha256=?",
                   (pages, title, dtype, "; ".join(methods), "; ".join(fams), lang, summary, tscore, evid,
                    meta_t[:200], len(re.sub(r"\s", "", text)), RULES_VERSION, topic, utcnow(), row["sha256"]))
        n += 1
    ACTIVITY.pop(threading.get_ident(), None)
    log.info("Classified %d file(s)", n)
    if unread:
        crypto = sum("cryptography" in e for e in unread)
        log.warning("%d PDF(s) could not be read%s - their topic is unknown, so they were not deleted",
                    len(unread), f" ({crypto} encrypted: run \"{sys.executable}\" -m pip install cryptography, "
                    "then --classify again)" if crypto else "")
    infer_from_siblings(db)
    pair_translations_by_text(db)
    if doc_types is not None and not keep_rejected:
        reject_unwanted(db, doc_types, equipment_vendors)
        if not keep_off_topic:
            delete_off_topic(db, root)
    total, with_en = mark_non_english(db)
    if total:
        log.info("Non-English documents: %d (all in _translations/) - %d with an English version in the main "
                 "folder, %d without (english_version = 'NOT FOUND')", total, with_en, total - with_en)
    return n


# =============================================================================
# Reporting
# =============================================================================
def print_status(db: DB, vendors: list[Vendor]):
    print(f"\n{'Vendor':38} {'Seed':7} {'Pages':>6} {'Queue':>6} {'PDFlinks':>8} {'TDS':>5} "
          f"{'Brch':>5} {'Other':>5} {'SDS/Cert/Corp':>13} {'Got':>5} {'Dup':>4} {'Err':>4}")
    print("-" * 122)
    tot = [0] * 10
    for v in sorted(vendors, key=lambda x: x.name.lower()):
        pg = db.query("SELECT state, COUNT(*) n FROM pages WHERE vendor=? GROUP BY state", (v.name,))
        pgd = {r["state"]: r["n"] for r in pg}
        lk = db.query("SELECT doc_type, state, COUNT(*) n FROM pdf_links WHERE vendor=? GROUP BY doc_type, state", (v.name,))
        dt, stt, total = {}, {}, 0
        for r in lk:
            dt[r["doc_type"]] = dt.get(r["doc_type"], 0) + r["n"]
            stt[r["state"]] = stt.get(r["state"], 0) + r["n"]
            total += r["n"]
        vals = [pgd.get("done", 0) + pgd.get("error", 0), pgd.get("queued", 0), total, dt.get("tds", 0),
                dt.get("brochure", 0), dt.get("other", 0),
                dt.get("sds", 0) + dt.get("certificate", 0) + dt.get("corporate", 0),
                stt.get("downloaded", 0), stt.get("duplicate", 0) + stt.get("unchanged", 0),
                stt.get("error", 0) + stt.get("not_pdf", 0)]
        tot = [a + b for a, b in zip(tot, vals)]
        print(f"{v.name[:38]:38} {v.status:7} {vals[0]:>6} {vals[1]:>6} {vals[2]:>8} {vals[3]:>5} {vals[4]:>5} "
              f"{vals[5]:>5} {vals[6]:>13} {vals[7]:>5} {vals[8]:>4} {vals[9]:>4}")
    print("-" * 122)
    print(f"{'TOTAL':38} {'':7} {tot[0]:>6} {tot[1]:>6} {tot[2]:>8} {tot[3]:>5} {tot[4]:>5} {tot[5]:>5} "
          f"{tot[6]:>13} {tot[7]:>5} {tot[8]:>4} {tot[9]:>4}\n")
    sz = db.query("SELECT COUNT(*) n, COALESCE(SUM(bytes),0) b FROM files")[0]
    print(f"Unique PDFs on disk: {sz['n']}  ({sz['b'] / 1e6:.1f} MB)")
    names = [v.name for v in vendors] or [""]
    q = ",".join("?" * len(names))
    tr = db.query(f"SELECT COUNT(*) n FROM pdf_links WHERE state='translation' AND vendor IN ({q})", names)[0]["n"]
    nf = db.query(f"SELECT COUNT(*) n FROM pdf_links WHERE state='downloaded' AND COALESCE(lang,'')<>'' "
                  f"AND vendor IN ({q})", names)[0]["n"]
    if tr or nf:
        print(f"Non-English copies: {tr} with an English version in the main folder; "
              f"{nf} without an English version (all filed under _translations/)")
    skipped = [v.name for v in vendors if v.status == "skip"]
    if skipped:
        print(f"Skipped (status=skip in seeds): {', '.join(skipped)}")


def export_manifest(db: DB, vendors: dict[str, Vendor], root: Path) -> Path:
    cats = {v.name: "; ".join(v.categories) for v in vendors.values()}
    folders = {v.name: v.category_folder for v in vendors.values()}
    out = root / "tds_manifest.csv"
    rows = db.query("""
        SELECT f.vendor, f.path, f.bytes, f.pages, f.title, f.text_doc_type, f.test_methods,
               f.language, f.is_translation, f.english_path, f.rejected, f.tech_score, f.evidence,
               l.lang AS link_lang,
               f.method_families, f.first_url, f.added_at, l.doc_type AS link_doc_type, l.anchor,
               l.found_on, l.last_modified,
               (SELECT COUNT(*) FROM pdf_links x WHERE x.sha256 = f.sha256) AS url_count, f.sha256
        FROM files f LEFT JOIN pdf_links l ON l.url = f.first_url
        ORDER BY f.vendor, f.path""")
    with open(out, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        clean = lambda row: [CONTROL_CHARS_RE.sub(" ", x) if isinstance(x, str) else x for x in row]
        w.writerow(["vendor", "categories", "category_folder", "path", "bytes", "pages", "title", "language",
                    "is_translated_copy", "english_version", "rejected_as", "tech_score", "evidence", "doc_type_from_text",
                    "doc_type_from_link", "test_methods", "method_families", "source_url", "anchor_text",
                    "found_on_page", "server_last_modified", "downloaded_at", "urls_with_same_file", "sha256"])
        for r in rows:
            lang = r["language"] or (r["link_lang"] or "")
            if r["rejected"]:
                eng = ""
            elif r["is_translation"]:
                eng = r["english_path"] or "NOT FOUND"
            else:
                eng = ""
            w.writerow(clean([r["vendor"], cats.get(r["vendor"], ""), folders.get(r["vendor"], ""), r["path"], r["bytes"],
                        r["pages"], r["title"], lang, "yes" if r["is_translation"] else "", eng, r["rejected"] or "", r["tech_score"], r["evidence"] or "",
                        r["text_doc_type"], r["link_doc_type"], r["test_methods"], r["method_families"],
                        r["first_url"], r["anchor"], r["found_on"], r["last_modified"], r["added_at"],
                        r["url_count"], r["sha256"]]))
    log.info("Manifest written: %s (%d rows)", out, len(rows))
    return out


DOC_TYPE_LABEL = {"tds": "Technical data sheet", "brochure": "Brochure / selector guide", "sds": "Safety data sheet",
                  "certificate": "Certificate / statement", "corporate": "Corporate document",
                  "commercial": "Commercial / supplier document",
                  "other": "Technical document", "unreadable": "Scanned / image-only PDF"}


def export_catalog_xlsx(db: DB, vendors: dict[str, Vendor], root: Path) -> Path | None:
    """tds_catalog.xlsx: a Summary sheet, one sheet per material category (Vendor | Document (clickable) |
    Description | Type | Test methods | Pages | Language | Source URL), then Translations and Rejected."""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
        from openpyxl.worksheet.hyperlink import Hyperlink
    except ImportError:
        log.warning("openpyxl missing - catalog not written")
        return None
    by_name = {v.name: v for v in vendors.values()}
    rows = db.query("""SELECT f.*, l.doc_type AS link_type FROM files f
                       LEFT JOIN pdf_links l ON l.url = f.first_url""")

    def category(vname: str) -> tuple[str, str]:
        v = by_name.get(vname)
        folder = v.category_folder if v else "0_Uncategorized"
        return folder, re.sub(r"^\d+_", "", folder).replace("_", " ")

    def sheet_name(folder: str) -> str:
        name = folder.replace("_", " ")                       # '4 Polymer Films'
        return re.sub(r"[\\/?*\[\]:]", "-", name)[:31]

    def type_label(r) -> str:
        return DOC_TYPE_LABEL.get(r["text_doc_type"] or r["link_type"] or "other", "Technical document")

    def describe(r) -> str:
        summ = r["summary"] or re.sub(r"[_\-]+", " ", Path(r["path"]).stem)
        return f"{type_label(r)}: {summ}"

    kept: dict[str, list] = {}
    translations, rejected = [], []
    for r in rows:
        if r["rejected"]:
            rejected.append(r)
        elif r["is_translation"]:
            translations.append(r)
        else:
            kept.setdefault(category(r["vendor"])[0], []).append(r)

    wb = Workbook()
    hdr_fill = PatternFill("solid", fgColor="1F3864")
    link_font = Font(color="0563C1", underline="single")

    def header(ws, cols):
        for c, (name, width) in enumerate(cols, 1):
            cell = ws.cell(1, c, name)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = hdr_fill
            cell.alignment = Alignment(vertical="center")
            ws.column_dimensions[get_column_letter(c)].width = width
        ws.freeze_panes = "A2"

    def put(ws, i, c, v, wrap=False):
        if isinstance(v, str):
            v = CONTROL_CHARS_RE.sub(" ", v)
        cell = ws.cell(i, c, v)
        cell.alignment = Alignment(vertical="top", wrap_text=wrap)
        return cell

    doc_cols = [("Vendor", 30), ("Document", 46), ("Description", 70), ("Type", 22),
                ("Test methods", 34), ("Evidence", 44), ("Pages", 7), ("Language", 9), ("Source URL", 50)]

    def write_docs(ws, items, lead_category: bool, extra=None):
        cols = ([("Category", 20)] if lead_category else []) + doc_cols + ([extra[0]] if extra else [])
        header(ws, cols)
        off = 1 if lead_category else 0
        items.sort(key=lambda r: (category(r["vendor"])[0], r["vendor"].lower(), Path(r["path"]).name.lower()))
        for i, r in enumerate(items, 2):
            vals = ([category(r["vendor"])[1]] if lead_category else []) + [
                r["vendor"], Path(r["path"]).name, describe(r), type_label(r), r["test_methods"] or "",
                evidence_label(r), r["pages"], r["language"] or "", r["first_url"]]
            if extra:
                vals.append(extra[1](r))
            for c, v in enumerate(vals, 1):
                put(ws, i, c, v, wrap=(c == 3 + off))
            doc = ws.cell(i, 2 + off)
            doc.hyperlink = r["path"]                 # relative to the workbook (saved in the corpus root)
            doc.font = link_font
            if r["first_url"]:
                ws.cell(i, 9 + off).hyperlink = r["first_url"]
        ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{max(1, len(items) + 1)}"

    # Summary
    ws = wb.active
    ws.title = "Summary"
    header(ws, [("Category", 30), ("Vendors", 9), ("Documents", 11), ("Translations", 13), ("Rejected", 10)])
    folders = sorted({v.category_folder for v in vendors.values() if v.status != "skip"} | set(kept))
    tr_by = {}
    rj_by = {}
    for r in translations:
        tr_by[category(r["vendor"])[0]] = tr_by.get(category(r["vendor"])[0], 0) + 1
    for r in rejected:
        rj_by[category(r["vendor"])[0]] = rj_by.get(category(r["vendor"])[0], 0) + 1
    i = 1
    for folder in folders:
        items = kept.get(folder, [])
        i += 1
        put(ws, i, 1, re.sub(r"^\d+_", "", folder).replace("_", " "))
        put(ws, i, 2, len({r["vendor"] for r in items}))
        put(ws, i, 3, len(items))
        put(ws, i, 4, tr_by.get(folder, 0))
        put(ws, i, 5, rj_by.get(folder, 0))
        if items:
            ws.cell(i, 1).hyperlink = Hyperlink(ref=f"A{i}", location=f"'{sheet_name(folder)}'!A1")
            ws.cell(i, 1).font = link_font
    i += 1
    put(ws, i, 1, "Total").font = Font(bold=True)
    for c, n in ((3, sum(len(v) for v in kept.values())), (4, len(translations)), (5, len(rejected))):
        put(ws, i, c, n).font = Font(bold=True)

    # One sheet per category
    for folder in sorted(kept):
        write_docs(wb.create_sheet(sheet_name(folder)), kept[folder], lead_category=False)
    # Translations / Rejected (span all categories)
    write_docs(wb.create_sheet("Translations"), translations, lead_category=True,
               extra=(("English version", 46),
                      lambda r: Path(r["english_path"]).name if r["english_path"] else "NOT FOUND"))
    write_docs(wb.create_sheet("Rejected"), rejected, lead_category=True,
               extra=(("Rejected as", 16), lambda r: r["rejected"] or ""))

    out = root / "tds_catalog.xlsx"
    try:
        wb.save(out)
    except PermissionError:
        out = root / f"tds_catalog_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        wb.save(out)
        log.warning("tds_catalog.xlsx is open in Excel - wrote %s instead", out.name)
    log.info("Catalog written: %s (%d documents in %d category sheets, %d translations, %d rejected)", out,
             sum(len(v) for v in kept.values()), len(kept), len(translations), len(rejected))
    return out


# =============================================================================
# Main
# =============================================================================
def setup_logging(root: Path, verbose: bool):
    (root / "logs").mkdir(parents=True, exist_ok=True)
    fmt = logging.Formatter("%(asctime)s %(levelname)-7s %(message)s", "%H:%M:%S")
    log.setLevel(logging.DEBUG)
    sh = logging.StreamHandler(sys.stdout)
    sh.setLevel(logging.DEBUG if verbose else logging.INFO)
    sh.setFormatter(fmt)
    fh = logging.FileHandler(root / "logs" / f"tds_crawler_{datetime.now():%Y%m%d}.log", encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)
    log.handlers[:] = [sh, fh]


def parse_args(argv=None):
    ap = argparse.ArgumentParser(description="Harvest technical data sheet PDFs from vendor websites.",
                                 formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    g = ap.add_argument_group("stages")
    g.add_argument("--discover", action="store_true", help="crawl sites and collect PDF links")
    g.add_argument("--download", action="store_true", help="download pending PDF links")
    g.add_argument("--classify", action="store_true", help="tag downloaded PDFs (doc type, test methods)")
    g.add_argument("--all", action="store_true", help="discover + download + classify + export")
    g.add_argument("--status", action="store_true", help="print per-vendor counts")
    g.add_argument("--export", action="store_true", help="write tds_manifest.csv and tds_catalog.xlsx")
    g.add_argument("--list-vendors", action="store_true", help="show vendors, categories and seed status")
    g.add_argument("--reorganize", action="store_true",
                   help="move existing PDFs into pdfs/<Category>/<Vendor>/ (also runs automatically with --download)")
    ap.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK)
    ap.add_argument("--seeds", type=Path, default=DEFAULT_SEEDS)
    ap.add_argument("--root", type=Path, default=Path(DEFAULT_ROOT), help="output root folder")
    ap.add_argument("--vendors", default="", help="comma list, case-insensitive substring match")
    ap.add_argument("--categories", default="", help="comma list of sheet-name substrings, e.g. 'foil,paper'")
    ap.add_argument("--only-ok", action="store_true", help="crawl only status=ok vendors")
    ap.add_argument("--max-pages", type=int, default=DEFAULT_MAX_PAGES)
    ap.add_argument("--max-depth", type=int, default=DEFAULT_MAX_DEPTH)
    ap.add_argument("--delay", type=float, default=DEFAULT_DELAY_S, help="seconds between requests per vendor")
    ap.add_argument("--workers", type=int, default=DEFAULT_WORKERS, help="vendors processed in parallel")
    ap.add_argument("--doc-types", default=DEFAULT_DOC_TYPES,
                    help=f"which link types to download: {','.join(ALL_DOC_TYPES)} or 'all'")
    ap.add_argument("--max-pdf-mb", type=int, default=DEFAULT_MAX_PDF_MB)
    ap.add_argument("--max-attempts", type=int, default=3, help="give up on a PDF after N failed attempts")
    ap.add_argument("--no-wp-media", action="store_true",
                    help="don't ask WordPress sites' media library API for their PDF list during --discover")
    ap.add_argument("--recrawl", action="store_true", help="redo discovery for vendors already completed")
    ap.add_argument("--refresh", action="store_true",
                    help="re-check downloaded PDFs with ETag/If-Modified-Since to catch revised sheets")
    ap.add_argument("--reclassify", action="store_true", help="re-run classification on all files")
    ap.add_argument("--keep-rejected", action="store_true",
                    help="turn off sorting: don't reject or delete anything (off-topic, corporate...)")
    ap.add_argument("--keep-off-topic", action="store_true",
                    help="don't delete off-topic files (outside the web / film / foil / paper / nonwoven / converting "
                         "scope); move them to _rejected/ instead")
    ap.add_argument("--keep-translations", action="store_true",
                    help="turn off language sorting: keep non-English documents in the main folder and don't hunt for "
                         "English versions")
    ap.add_argument("--js", action="store_true", help="render render_js=1 vendors with Playwright")
    ap.add_argument("--search-seed", action="store_true",
                    help="seed PDF URLs from SerpAPI or Google CSE (env SERPAPI_KEY or GOOGLE_CSE_KEY+GOOGLE_CSE_CX)")
    ap.add_argument("--no-sitemaps", action="store_true")
    ap.add_argument("--offsite-pdfs", action=argparse.BooleanOptionalAction, default=True,
                    help="accept PDF links hosted on other domains (CDNs, distributors)")
    ap.add_argument("--ignore-robots", action="store_true", help=argparse.SUPPRESS)
    ap.add_argument("--contact", default="", help="email or URL appended to the User-Agent (recommended)")
    ap.add_argument("-v", "--verbose", action="store_true")
    a = ap.parse_args(argv)
    a.doc_types = ALL_DOC_TYPES if a.doc_types.strip().lower() == "all" else \
        [t.strip().lower() for t in a.doc_types.split(",") if t.strip()]
    bad = set(a.doc_types) - set(ALL_DOC_TYPES)
    if bad:
        ap.error(f"unknown doc types: {bad}")
    if a.reclassify:
        a.classify = True
    if a.all:
        a.discover = a.download = a.classify = a.export = True
    if not any([a.discover, a.download, a.classify, a.status, a.export, a.list_vendors, a.reorganize]):
        ap.print_help()
        sys.exit(0)
    return a


def select_vendors(all_v: dict[str, Vendor], args) -> list[Vendor]:
    vs = list(all_v.values())
    if args.vendors:
        pats = [norm_name(p) for p in args.vendors.split(",") if p.strip()]
        vs = [v for v in vs if any(p in v.key for p in pats)]
    if args.categories:
        pats = [p.strip().lower() for p in args.categories.split(",") if p.strip()]
        vs = [v for v in vs if any(p in c.lower() for c in v.categories for p in pats)]
    return vs


def run_parallel(fn, vendors: list[Vendor], workers: int):
    """Run fn(vendor) on a few vendors at a time. Waits in 1-second steps so Ctrl-C is noticed at once
    on Windows: the first Ctrl-C finishes the current requests and saves state (seconds, at most about a
    minute); a second Ctrl-C quits immediately (the database stays consistent either way)."""
    def run_one(v):
        try:
            return fn(v)
        finally:
            ACTIVITY.pop(threading.get_ident(), None)

    ex = ThreadPoolExecutor(max_workers=max(1, workers))
    futs = {ex.submit(run_one, v): v for v in vendors}
    pending = set(futs)
    try:
        while pending:
            done, pending = wait(pending, timeout=1, return_when=FIRST_COMPLETED)
            for fut in done:
                try:
                    fut.result()
                except Exception as e:
                    log.exception("[%s] failed: %s", futs[fut].name, e)
    except KeyboardInterrupt:
        log.warning("Ctrl-C: finishing current requests and saving state (press Ctrl-C again to quit now)...")
        STOP.set()
        for f_ in pending:
            f_.cancel()
        try:
            while any(not f_.done() for f_ in futs):
                wait(list(futs), timeout=1)
        except KeyboardInterrupt:
            log.warning("Second Ctrl-C: quitting now. Re-run the same command to continue.")
            logging.shutdown()
            os._exit(130)
        raise
    finally:
        ex.shutdown(wait=False, cancel_futures=True)

def main(argv=None):
    args = parse_args(argv)
    root = args.root
    root.mkdir(parents=True, exist_ok=True)
    setup_logging(root, args.verbose)
    log.info("tds_crawler_v1 build %s (rules v%d) - %s", BUILD, RULES_VERSION, Path(__file__).resolve())
    threading.Thread(target=heartbeat, daemon=True, name="heartbeat").start()
    vendors_all = load_workbook_vendors(args.workbook)
    merge_seeds(vendors_all, args.seeds)
    db = DB(root / "tds_crawler.db")
    ensure_vendor_folders(vendors_all, root)
    chosen = select_vendors(vendors_all, args)
    allowed_status = {"ok"} if args.only_ok else {"ok", "verify"}
    active = [v for v in chosen if v.status in allowed_status and v.domains and v.start_urls]

    if args.list_vendors:
        for v in sorted(chosen, key=lambda x: x.name.lower()):
            print(f"{v.status:7} {v.name:38} {v.category_folder:24} {';'.join(v.domains)}"
                  f"{'  [JS]' if v.render_js else ''}")
        print(f"\n{len(chosen)} vendors selected, {len(active)} crawlable")
    try:
        backfill_link_lang(db)
        apply_pdf_filters(db, vendors_all)
        if args.reorganize and not args.keep_translations:
            for v in vendors_all.values():
                tidy_translations(db, v, make_guesses=False)
            pair_translations_by_text(db)
            mark_non_english(db)
        if args.reorganize or args.download:
            reorganize_files(db, vendors_all, root)
        if args.discover:
            log.info("DISCOVER: %d vendor(s), max %d pages each, %.1fs delay, %d parallel",
                     len(active), args.max_pages, args.delay, args.workers)
            if args.js:
                try:
                    import playwright  # noqa
                except ImportError:
                    log.warning("--js requested but Playwright is not installed for this Python (%s). Run: "
                                "\"%s\" -m pip install playwright  - continuing without JavaScript rendering",
                                sys.executable, sys.executable)
                    args.js = False
            if args.js:
                try:
                    r = Renderer()
                    log.info("JavaScript rendering: using %s", r.name)
                    r.close()
                except Exception as e:
                    log.warning("--js: %s. Install Microsoft Edge or Google Chrome, or run \"%s\" -m playwright "
                                "install chromium - continuing without JavaScript rendering", e, sys.executable)
                    args.js = False
            run_parallel(lambda v: VendorCrawler(v, db, args).run(), active, args.workers)
        if args.download and not STOP.is_set():
            log.info("DOWNLOAD: types=%s", ",".join(args.doc_types))
            backfill_link_lang(db)
            apply_pdf_filters(db, vendors_all)   # discovery may have handed shared-site PDFs to the distributor
            dl_vendors = [v for v in chosen if v.status in allowed_status]
            run_parallel(lambda v: download_vendor(v, db, args, root), dl_vendors, args.workers)
            if not args.keep_translations:
                pair_translations_by_text(db)
                mark_non_english(db)
            reorganize_files(db, vendors_all, root)   # files into place; non-English copies -> _translations/
        if args.classify and not STOP.is_set():
            backfill_link_lang(db)
            classify_files(db, root, redo=args.reclassify, doc_types=args.doc_types, keep_rejected=args.keep_rejected,
                           equipment_vendors=equipment_vendors(vendors_all), keep_off_topic=args.keep_off_topic,
                           scopes=vendor_scopes(vendors_all))
            if args.download and not args.keep_translations and not STOP.is_set():
                more = queue_english_guesses(db, vendors_all, only={v.name for v in chosen})
                if more:
                    log.info("Looking for English versions of non-English documents (%d vendor(s))", len(more))
                    run_parallel(lambda v: download_vendor(v, db, args, root), more, args.workers)
                    classify_files(db, root, redo=False, doc_types=args.doc_types, keep_rejected=args.keep_rejected,
                                   equipment_vendors=equipment_vendors(vendors_all),
                                   keep_off_topic=args.keep_off_topic, scopes=vendor_scopes(vendors_all))
            reorganize_files(db, vendors_all, root)
        if args.export:
            export_manifest(db, vendors_all, root)
            export_catalog_xlsx(db, vendors_all, root)
    except KeyboardInterrupt:
        log.warning("Interrupted - state saved; re-run the same command to resume.")
    stale = db.query(f"SELECT COUNT(*) n FROM files WHERE COALESCE(rules_version, 0) <> {RULES_VERSION}")[0]["n"]
    if stale and not args.classify:
        log.warning("%d file(s) were sorted by older rules - run with --classify to re-check them", stale)
    if args.status or args.discover or args.download:
        print_status(db, chosen)


if __name__ == "__main__":
    main()
