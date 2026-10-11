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
from concurrent.futures import ThreadPoolExecutor, as_completed
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
DEFAULT_DELAY_S = 2.0          # min seconds between requests to the same vendor
DEFAULT_MAX_PAGES = 400        # HTML pages fetched per vendor per discovery run
DEFAULT_MAX_DEPTH = 4
DEFAULT_WORKERS = 4            # vendors crawled in parallel (one thread each)
DEFAULT_MAX_PDF_MB = 50
DEFAULT_DOC_TYPES = "tds,brochure,other"   # sds / certificate / corporate skipped
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
ALL_DOC_TYPES = ["tds", "sds", "brochure", "certificate", "corporate", "other"]
# Stricter certificate test for PDF *text* (TDS bodies often mention REACH/RoHS/FDA compliance)
TEXT_CERT_RE = re.compile(r"(certificate of (analysis|conformance|conformity|compliance)|declaration of|"
                          r"\bstatement (of|on|regarding)\b|conflict minerals|to whom it may concern|"
                          r"dear (valued )?customer|certification letter)", re.I)

# --- translated copies (English-first) ----------------------------------------
# "X - CN.pdf", "X_JP.pdf", "X (KR).pdf", "X-DE-2.pdf"... (upper-case codes only, to avoid product codes)
LANG_CODES = {"CN": "zh", "ZH": "zh", "TW": "zh", "CHS": "zh", "CHT": "zh", "JP": "ja", "JA": "ja",
              "KR": "ko", "KO": "ko", "DE": "de", "FR": "fr", "ES": "es", "IT": "it", "PT": "pt",
              "BR": "pt", "RU": "ru", "PL": "pl", "TR": "tr", "CZ": "cs", "NL": "nl", "SE": "sv",
              "HU": "hu"}
LANG_WORDS = {"chinese": "zh", "japanese": "ja", "korean": "ko", "german": "de", "deutsch": "de",
              "french": "fr", "francais": "fr", "spanish": "es", "espanol": "es", "italian": "it",
              "italiano": "it", "portuguese": "pt", "russian": "ru", "polish": "pl", "turkish": "tr"}
TRANSLATION_CODE_RE = re.compile(
    r"^(?P<base>.+?)(?:[\s_\-.]+\(?|\()(?P<code>" + "|".join(sorted(LANG_CODES, key=len, reverse=True)) +
    r")\)?(?:[\s_\-]?\d{1,2})?$")
TRANSLATION_WORD_RE = re.compile(
    r"^(?P<base>.+?)[\s_\-(]+(?P<word>" + "|".join(LANG_WORDS) + r")\)?$", re.I)
ENGLISH_MARK_RE = re.compile(r"[\s_\-.]+\(?(EN|ENG|[Ee]nglish|en|eng|US|GB|UK)\)?(?:[\s_\-]?\d{1,2})?$")
COLLISION_SUFFIX_RE = re.compile(r"_[0-9a-f]{8}$")
GOOD_STATES = ("downloaded", "duplicate", "unchanged")
CORPORATE_TEXT_RE = re.compile(
    r"(integrated report|annual report|sustainability report|medium\s*-?\s*term\s+management|"
    r"long\s*-?\s*term\s+vision|management plan|balance sheet|financial results|financial statements|securities report|shareholders|"
    r"corporate governance report|esg data|"
    r"貸借対照表|損益計算書|中期経営計画|長期ビジョン|統合報告|決算|株主|有価証券報告書|各\s*位|ステークホルダー|"
    r"お問い合わせ先|ニュースリリース|プレスリリース|会社概要|"
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
    return s


def classify_doc_type(text: str) -> str:
    for dtype, rx in DOC_TYPE_PATTERNS:
        if rx.search(text):
            return dtype
    return "other"


def stem_of(url: str) -> str:
    name = unquote(urlparse(url).path.rstrip("/").rsplit("/", 1)[-1])
    return name[:-4] if name.lower().endswith(".pdf") else name


def match_key(stem: str) -> str:
    """Comparable key for a filename stem: case/punctuation-free, English markers removed."""
    stem = COLLISION_SUFFIX_RE.sub("", stem.strip())
    stem = ENGLISH_MARK_RE.sub("", stem)
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
    loc_idx = next((i for i, sg in enumerate(segs[:5]) if LOCALE_RE.match(sg.lower())
                    and sg.lower()[:2] != "en" and sg.lower() not in LOCALE_ALLOW), None)
    if not lang and loc_idx is not None:
        lang, base = segs[loc_idx].lower()[:2], stem
    if not lang:
        return "", "", []
    base = base.rstrip(" _-.")
    guesses: list[str] = []
    folder = p.path.rsplit("/", 1)[0]
    if base != stem:
        for name in (base, base + "-EN", base + "_EN", base + " - EN", base + " EN"):
            guesses.append(f"{p.scheme}://{p.netloc}{folder}/{quote(name)}.pdf")
    if loc_idx is not None:
        for en in ("en", "en-us", "en-gb"):
            sg = list(segs)
            sg[loc_idx] = en
            guesses.append(f"{p.scheme}://{p.netloc}{'/'.join(sg)}")
    return lang, match_key(base), list(dict.fromkeys(g for g in guesses if canon_url(g) != canon_url(url)))


ENGLISH_STOPWORDS = {"the", "and", "of", "for", "with", "to", "in", "is", "are", "this", "that", "by",
                     "on", "from", "or", "be", "as", "at", "properties", "typical", "value", "test"}
OTHER_STOPWORDS = {
    "de": {"und", "der", "die", "das", "mit", "für", "ist", "von", "den", "eine", "nicht", "werden"},
    "fr": {"et", "les", "des", "pour", "avec", "est", "une", "dans", "sur", "par", "aux", "sont"},
    "es": {"y", "los", "las", "para", "con", "es", "una", "del", "por", "como", "más", "son"},
    "it": {"e", "il", "della", "per", "con", "sono", "una", "delle", "degli", "nel", "che", "alla"},
    "pt": {"e", "os", "as", "para", "com", "uma", "dos", "das", "não", "são", "pelo", "pela"},
}


def detect_language(text: str) -> str:
    """Rough language of PDF text: en, zh, ja, ko, ru, de, fr, es, it, pt, '<xx>+en' (bilingual) or ''."""
    if len(text.strip()) < 40:
        return ""
    words = re.findall(r"[A-Za-zÀ-ÿ]+", text.lower())
    en_hits = sum(1 for w in words if w in ENGLISH_STOPWORDS)
    counts = {"ja": len(re.findall(r"[\u3040-\u30ff]", text)), "ko": len(re.findall(r"[\uac00-\ud7af]", text)),
              "zh": len(re.findall(r"[\u4e00-\u9fff]", text)), "ru": len(re.findall(r"[\u0400-\u04ff]", text))}
    if counts["ja"] > 20:
        script = "ja"
    else:
        script = max(counts, key=counts.get) if max(counts.values()) > 30 else ""
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
    return bool(LOCALE_RE.match(first)) and first not in LOCALE_ALLOW and first not in vendor_prefix_locales


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
               "pdf_filter", "notes"]


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
        pf = (row.get("pdf_filter") or "").strip()
        if pf:
            m = re.match(r"^([a-z0-9.-]+\.[a-z]{2,}):(.+)$", pf, re.I)
            dom, rx = (m.group(1).lower(), m.group(2)) if m else ("", pf)
            try:
                v.pdf_filter, v.pdf_filter_domain = re.compile(rx, re.I), dom
            except re.error as e:
                log.error("Bad pdf_filter regex for %s: %s", v.name, e)
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
        if "rejected" not in {r[1] for r in c.execute("PRAGMA table_info(files)")}:
            c.execute("ALTER TABLE files ADD COLUMN rejected TEXT")
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
        return rp.can_fetch(self.ua, url) if rp else True

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
        exe = "/opt/pw-browsers/chromium" if Path("/opt/pw-browsers/chromium").exists() else None
        self.browser = self._pw.chromium.launch(headless=True, **({"executable_path": exe} if exe else {}))
        self.ctx = self.browser.new_context(user_agent=USER_AGENT_BASE)

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
            if segs and LOCALE_RE.match(segs[0]):
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

    def add_pdf(self, url: str, found_on: str, anchor: str):
        url = canon_url(url)
        if not url or url in self.known_pdfs or not self.pdf_allowed_host(url):
            return
        self.known_pdfs.add(url)
        text = f"{unquote(urlparse(url).path)} {anchor}"
        filt = self.v.filter_for(host_of(url))
        if filt is not None and not filt.search(text):
            return  # another company's sheet on a shared (distributor) site
        dtype = classify_doc_type(text)
        segs = {sg.lower() for u in (url, found_on or "") for sg in urlparse(u).path.split("/") if sg}
        if segs & CORPORATE_SEGMENTS:
            dtype = "corporate"
        self.db.execute(
            "INSERT OR IGNORE INTO pdf_links(url, vendor, found_on, anchor, doc_type, first_seen, lang) "
            "VALUES (?,?,?,?,?,?,?)",
            (url, self.v.name, found_on, (anchor or "")[:300], dtype, utcnow(), translation_info(url)[0]))

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

    def run(self):
        v = self.v
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

        fetched = 0
        while self.heap and fetched < self.args.max_pages and not STOP.is_set():
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
                 f" (page cap {self.args.max_pages} reached, {len(self.heap)} left queued - re-run --discover to continue)"
                 if self.heap and fetched >= self.args.max_pages else "")


# =============================================================================
# DOWNLOAD
# =============================================================================
_file_lock = threading.Lock()


def filename_from_response(r: requests.Response, url: str) -> str:
    cd = r.headers.get("Content-Disposition", "")
    m = re.search(r"filename\*\s*=\s*[^']*''([^;]+)", cd) or re.search(r'filename\s*=\s*"?([^";]+)"?', cd)
    if m:
        return safe_filename(m.group(1))
    last = urlparse(r.url or url).path.rstrip("/").rsplit("/", 1)[-1]
    return safe_filename(last or "document")


def _download_one(row, f: Fetcher, db: DB, args, root: Path, vendor: Vendor, out_dir: Path) -> str:
    """Download one pdf_links row. Returns the resulting state."""
    url = row["url"]
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
        r = f.get(url, stream=True, headers=hdrs)
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
        with open(tmp, "wb") as fh:
            for chunk in r.iter_content(65536):
                if STOP.is_set():
                    ok = False
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
    _, _, guesses = translation_info(row["url"])
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


def download_vendor(vendor: Vendor, db: DB, args, root: Path) -> dict:
    f = Fetcher(args)
    types = set(args.doc_types)
    out_dir = vendor.pdf_dir(root)
    stats = {"downloaded": 0, "duplicate": 0, "unchanged": 0, "failed": 0, "skipped_type": 0,
             "translations_skipped": 0, "english_found_by_guess": 0, "foreign_kept_no_english": 0}
    english_first = not args.keep_translations
    if english_first:
        tidy_translations(db, vendor, make_guesses=True)
    states = ("pending", "error") if not args.refresh else ("pending", "error", "downloaded", "unchanged")
    rows = db.query(
        f"SELECT * FROM pdf_links WHERE vendor=? AND state IN ({','.join('?' * len(states))}) "
        f"AND attempts < ? ORDER BY url", (vendor.name, *states, args.max_attempts if not args.refresh else 999))

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
        if en is not None:
            db.execute("UPDATE pdf_links SET state='translation', english_url=? WHERE url=?", (en["url"], row["url"]))
            stats["translations_skipped"] += 1
            continue
        st = _download_one(row, f, db, args, root, vendor, out_dir)
        tally(st)
        if st == "downloaded":
            stats["foreign_kept_no_english"] += 1
    if english_first:
        tidy_translations(db, vendor, make_guesses=False)
    log.info("[%s] download: %s", vendor.name, ", ".join(f"{k}={v}" for k, v in stats.items() if v))
    return stats


def backfill_link_lang(db: DB) -> None:
    """Links recorded before translation handling existed get their language tag filled in."""
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
    for row in db.query("SELECT sha256, path, vendor, is_translation, rejected FROM files"):
        v = by_name.get(row["vendor"])
        if v is None:
            continue
        old = root / row["path"]
        sub = "_rejected" if row["rejected"] else ("_translations" if row["is_translation"] else "")
        target_dir = v.pdf_dir(root) / sub
        if old.parent == target_dir:
            continue
        if not old.exists():
            log.warning("reorganize: missing file %s", old)
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
    for d in sorted(old_dirs, key=lambda p: len(p.parts), reverse=True):
        try:
            if d != root / "pdfs" and not any(d.iterdir()):
                d.rmdir()
        except OSError:
            pass
    if moved:
        log.info("Moved %d PDF(s) into pdfs/<Category>/<Vendor>/ (translations -> _translations/, "
                 "junk -> _rejected/)", moved)
    return moved


# =============================================================================
# CLASSIFY (pypdf)
# =============================================================================
def extract_methods(text: str) -> tuple[list[str], list[str]]:
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


def classify_text_doc_type(text: str, link_type: str = "", n_methods: int = 0) -> str:
    """Classify from extracted PDF text: header first, body as fallback. A filename that already says
    'data sheet', or a page citing 2+ test methods, counts as a data sheet unless it is an SDS."""
    if link_type == "corporate":
        return "corporate"
    if not text.strip():
        return "unreadable"
    rx = dict(DOC_TYPE_PATTERNS)
    head = text[:800]
    if rx["sds"].search(head) and not rx["tds"].search(head[:200]):
        return "sds"
    if CORPORATE_TEXT_RE.search(head):
        return "corporate"
    if link_type == "tds" or n_methods >= 2:
        return "tds"
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


def reject_unwanted(db: DB, doc_types: list[str]) -> int:
    """Files whose *text* shows a type you don't collect (corporate, SDS, certificate...) are flagged
    rejected and moved to _rejected/ - never deleted. Unreadable files are kept."""
    keep = set(doc_types) | {"unreadable"}
    n = 0
    for r in db.query("SELECT sha256, text_doc_type, rejected FROM files WHERE text_doc_type IS NOT NULL"):
        want = r["text_doc_type"] if r["text_doc_type"] not in keep else None
        if (r["rejected"] or None) != want:
            db.execute("UPDATE files SET rejected=? WHERE sha256=?", (want, r["sha256"]))
            n += 1 if want else 0
    if n:
        log.info("Flagged %d downloaded file(s) as not wanted by their text (moved to _rejected/)", n)
    return n


def pair_translations_by_text(db: DB) -> int:
    """Non-English files (by text language) whose English twin is on disk under a matching name,
    e.g. 'C1990_data.pdf' (Japanese) + 'C1990_data_english.pdf' -> mark the first as a translation."""
    n = 0
    files = db.query("SELECT sha256, path, vendor, language, is_translation FROM files")
    english: dict = {}
    for f in files:
        if f["language"] in ("en",) or (f["language"] or "").endswith("+en"):
            english.setdefault((f["vendor"], match_key(Path(f["path"]).stem)), f)
    for f in files:
        lang = f["language"] or ""
        if f["is_translation"] or not lang or lang == "en" or lang.endswith("+en"):
            continue
        en = english.get((f["vendor"], match_key(Path(f["path"]).stem)))
        if en is not None and en["sha256"] != f["sha256"]:
            db.execute("UPDATE files SET is_translation=1, english_path=? WHERE sha256=?", (en["path"], f["sha256"]))
            db.execute("UPDATE pdf_links SET state='translation', english_url=(SELECT first_url FROM files WHERE "
                       "sha256=?) WHERE sha256=? AND state='downloaded'", (en["sha256"], f["sha256"]))
            n += 1
    if n:
        log.info("Paired %d non-English file(s) with their English version by name (moved to _translations/)", n)
    return n


def classify_files(db: DB, root: Path, redo: bool = False, doc_types=None, keep_rejected: bool = False) -> int:
    try:
        from pypdf import PdfReader
    except ImportError:
        log.warning("pypdf not installed - skipping classification (pip install pypdf)")
        return 0
    logging.getLogger("pypdf").setLevel(logging.ERROR)
    rows = db.query("SELECT f.sha256, f.path, l.doc_type AS link_type FROM files f "
                    "LEFT JOIN pdf_links l ON l.url = f.first_url" +
                    ("" if redo else " WHERE f.classified_at IS NULL OR f.language IS NULL"))
    n = 0
    for row in rows:
        if STOP.is_set():
            break
        p = root / row["path"]
        title, pages, text = "", None, ""
        try:
            rd = PdfReader(str(p))
            pages = len(rd.pages)
            for pg in rd.pages[:2]:
                text += (pg.extract_text() or "") + "\n"
            meta_title = (rd.metadata.title if rd.metadata else None) or ""
            first_line = next((l.strip() for l in text.splitlines() if len(l.strip()) > 3), "")
            title = (meta_title if len(meta_title.strip()) > 3 else first_line)[:200]
        except Exception as e:
            log.debug("classify %s: %s", p, e)
        methods, fams = extract_methods(text)
        dtype = classify_text_doc_type(text, row["link_type"] or "", len(methods))
        lang = detect_language(text)
        db.execute("UPDATE files SET pages=?, title=?, text_doc_type=?, test_methods=?, method_families=?, "
                   "language=?, classified_at=? WHERE sha256=?",
                   (pages, title, dtype, "; ".join(methods), "; ".join(fams), lang, utcnow(), row["sha256"]))
        n += 1
    log.info("Classified %d file(s)", n)
    if doc_types is not None and not keep_rejected:
        reject_unwanted(db, doc_types)
    pair_translations_by_text(db)
    lone = db.query("SELECT COUNT(*) n FROM files WHERE COALESCE(is_translation,0)=0 AND rejected IS NULL AND "
                    "language NOT IN ('', 'en') AND language NOT LIKE '%+en'")[0]["n"]
    if lone:
        log.info("%d non-English file(s) have no English version on disk (manifest column english_version = "
                 "'NOT FOUND')", lone)
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
        print(f"Translated copies: {tr} skipped/moved because the English version exists; "
              f"{nf} kept because no English version was found")
    skipped = [v.name for v in vendors if v.status == "skip"]
    if skipped:
        print(f"Skipped (status=skip in seeds): {', '.join(skipped)}")


def export_manifest(db: DB, vendors: dict[str, Vendor], root: Path) -> Path:
    cats = {v.name: "; ".join(v.categories) for v in vendors.values()}
    folders = {v.name: v.category_folder for v in vendors.values()}
    out = root / "tds_manifest.csv"
    rows = db.query("""
        SELECT f.vendor, f.path, f.bytes, f.pages, f.title, f.text_doc_type, f.test_methods,
               f.language, f.is_translation, f.english_path, f.rejected, l.lang AS link_lang,
               f.method_families, f.first_url, f.added_at, l.doc_type AS link_doc_type, l.anchor,
               l.found_on, l.last_modified,
               (SELECT COUNT(*) FROM pdf_links x WHERE x.sha256 = f.sha256) AS url_count, f.sha256
        FROM files f LEFT JOIN pdf_links l ON l.url = f.first_url
        ORDER BY f.vendor, f.path""")
    with open(out, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["vendor", "categories", "category_folder", "path", "bytes", "pages", "title", "language",
                    "is_translated_copy", "english_version", "rejected_as", "doc_type_from_text",
                    "doc_type_from_link", "test_methods", "method_families", "source_url", "anchor_text",
                    "found_on_page", "server_last_modified", "downloaded_at", "urls_with_same_file", "sha256"])
        for r in rows:
            lang = r["language"] or (r["link_lang"] or "")
            if r["rejected"]:
                eng = ""
            elif r["is_translation"]:
                eng = r["english_path"] or ""
            elif lang and lang != "en" and not lang.endswith("+en"):
                eng = "NOT FOUND"
            else:
                eng = ""
            w.writerow([r["vendor"], cats.get(r["vendor"], ""), folders.get(r["vendor"], ""), r["path"], r["bytes"],
                        r["pages"], r["title"], lang, "yes" if r["is_translation"] else "", eng, r["rejected"] or "",
                        r["text_doc_type"], r["link_doc_type"], r["test_methods"], r["method_families"],
                        r["first_url"], r["anchor"], r["found_on"], r["last_modified"], r["added_at"],
                        r["url_count"], r["sha256"]])
    log.info("Manifest written: %s (%d rows)", out, len(rows))
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
    g.add_argument("--export", action="store_true", help="write tds_manifest.csv")
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
    ap.add_argument("--recrawl", action="store_true", help="redo discovery for vendors already completed")
    ap.add_argument("--refresh", action="store_true",
                    help="re-check downloaded PDFs with ETag/If-Modified-Since to catch revised sheets")
    ap.add_argument("--reclassify", action="store_true", help="re-run classification on all files")
    ap.add_argument("--keep-rejected", action="store_true",
                    help="don't move files whose text shows an unwanted type (corporate, SDS...) to _rejected/")
    ap.add_argument("--keep-translations", action="store_true",
                    help="download translated copies (e.g. '- CN') even when the English version exists")
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
    with ThreadPoolExecutor(max_workers=max(1, workers)) as ex:
        futs = {ex.submit(fn, v): v for v in vendors}
        try:
            for fut in as_completed(futs):
                try:
                    fut.result()
                except Exception as e:
                    log.exception("[%s] failed: %s", futs[fut].name, e)
        except KeyboardInterrupt:
            log.warning("Ctrl-C: finishing current requests and saving state...")
            STOP.set()
            raise


def main(argv=None):
    args = parse_args(argv)
    root = args.root
    root.mkdir(parents=True, exist_ok=True)
    setup_logging(root, args.verbose)
    vendors_all = load_workbook_vendors(args.workbook)
    merge_seeds(vendors_all, args.seeds)
    db = DB(root / "tds_crawler.db")
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
        if args.reorganize or args.download:
            reorganize_files(db, vendors_all, root)
        if args.discover:
            log.info("DISCOVER: %d vendor(s), max %d pages each, %.1fs delay, %d parallel",
                     len(active), args.max_pages, args.delay, args.workers)
            if args.js:
                try:
                    import playwright  # noqa
                except ImportError:
                    log.warning("--js requested but playwright not installed (pip install playwright)")
                    args.js = False
            run_parallel(lambda v: VendorCrawler(v, db, args).run(), active, args.workers)
        if args.download and not STOP.is_set():
            log.info("DOWNLOAD: types=%s", ",".join(args.doc_types))
            backfill_link_lang(db)
            apply_pdf_filters(db, vendors_all)   # discovery may have handed shared-site PDFs to the distributor
            dl_vendors = [v for v in chosen if v.status in allowed_status]
            run_parallel(lambda v: download_vendor(v, db, args, root), dl_vendors, args.workers)
            reorganize_files(db, vendors_all, root)   # files into place; translated copies -> _translations/
        if args.classify and not STOP.is_set():
            backfill_link_lang(db)
            classify_files(db, root, redo=args.reclassify, doc_types=args.doc_types, keep_rejected=args.keep_rejected)
            reorganize_files(db, vendors_all, root)
        if args.export:
            export_manifest(db, vendors_all, root)
    except KeyboardInterrupt:
        log.warning("Interrupted - state saved; re-run the same command to resume.")
    if args.status or args.discover or args.download:
        print_status(db, chosen)


if __name__ == "__main__":
    main()
