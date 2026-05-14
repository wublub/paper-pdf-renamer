"""PDF 文献一键命名工具

特点:
  - 自动从 PDF 中提取 DOI, 通过 CrossRef API 获取准确元数据
  - 自动从爱科学查询期刊影响因子、JCR 分区和中科院分区
  - 支持字段: Title / Journal / Journal-Abbr / Year / Accepted-Date /
              Published-Date / First-Author / DOI / IF / JCR / 中科院分区
  - 命名顺序、分隔符可在 GUI 中自由调整
  - 配置自动保存到 ~/.pdf_renamer_config.json

依赖: pypdf (若未安装会自动提示安装)
"""

import json
import re
import subprocess
import sys
import threading
import tkinter as tk
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from html import escape
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Optional

try:
    from pypdf import PdfReader
except ImportError:
    try:
        from PyPDF2 import PdfReader  # type: ignore
    except ImportError:
        PdfReader = None  # type: ignore


CONFIG_PATH = Path.home() / ".pdf_renamer_config.json"
IIKX_API_URL = "https://www.iikx.com/api/search/"
IIKX_ALL_CLASSIDS = "125,124,126,127,128,129,130,131,132,133,134,135,123,136"
JOURNAL_QUERY_STOPWORDS = {"a", "an", "and", "for", "in", "of", "on", "the"}

# 视觉主题:苹果风格 (iOS / macOS Big Sur 配色)
THEME = {
    "bg":            "#f5f5f7",   # 窗口浅灰底
    "panel":         "#ffffff",   # 卡片白底
    "border":        "#d2d2d7",   # 极浅灰边框
    "text":          "#1d1d1f",   # 主文字 (Apple 近黑)
    "muted":         "#86868b",   # 二级文字
    "primary":       "#0071e3",   # Apple 蓝
    "primary_hover": "#0077ed",
    "primary_press": "#0066cc",
    "accent":        "#34c759",   # iOS 绿
    "accent_hover":  "#30b850",
    "danger":        "#ff3b30",   # iOS 红
    "danger_hover":  "#e0342a",
    "stripe":        "#fafafa",   # 表格斑马纹
    "select_bg":     "#e1eeff",   # 浅蓝选中
    "select_fg":     "#0049a7",
    "heading_bg":    "#f5f5f7",
    "btn_bg":        "#e8e8ed",   # 次按钮浅灰
    "btn_hover":     "#dcdce1",
    "btn_press":     "#cfcfd5",
}
UI_FONT = ("Microsoft YaHei UI", 10)
UI_FONT_BOLD = ("Microsoft YaHei UI", 10, "bold")
UI_FONT_SMALL = ("Microsoft YaHei UI", 9)
UI_FONT_TITLE = ("Microsoft YaHei UI", 16, "bold")
UI_FONT_SUBTITLE = ("Microsoft YaHei UI", 11, "bold")
MONO_FONT = ("Cascadia Mono", 10)


def resource_path(rel: str) -> Path:
    """获取资源路径，兼容源码运行和 PyInstaller 打包后的运行环境。"""
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return Path(base) / rel
    return Path(__file__).resolve().parent / rel

FIELDS = {
    "title":          "Title  标题",
    "journal":        "Journal  期刊全名",
    "journal_short":  "Journal-Abbr  期刊缩写",
    "issn":           "ISSN",
    "eissn":          "E-ISSN",
    "year":           "Year  年份",
    "accepted_date":  "Accepted-Date  接受日期",
    "published_date": "Published-Date  出版日期",
    "first_author":   "First-Author  第一作者",
    "doi":            "DOI",
    "impact_factor":  "Impact-Factor  影响因子",
    "impact_factor_year": "IF-Year  影响因子年份",
    "jcr_quartile":   "JCR  分区",
    "cas_category":   "CAS-Category  中科院大类",
    "cas_quartile":   "CAS-Quartile  中科院区",
    "cas_partition":  "CAS  中科院分区",
    "journal_subject": "Subject  小类学科",
    "jcr_category":   "JCR-Category  学科分类",
    "publisher":      "Publisher  出版商",
    "five_year_if":   "5-Year-IF  五年影响因子",
    "oa_status":      "OA  开放访问",
    "journal_info_url": "Journal-URL  期刊信息页",
}

DEFAULT_TOKENS = [
    {"type": "field", "value": "year"},
    {"type": "sep",   "value": "_"},
    {"type": "field", "value": "journal"},
    {"type": "sep",   "value": "_"},
    {"type": "field", "value": "title"},
]

DEFAULT_PREVIEW_FIELDS = [
    "title",
    "journal",
    "impact_factor",
    "jcr_quartile",
    "cas_partition",
    "journal_subject",
    "year",
    "accepted_date",
]

PREVIEW_FIXED_LEFT = ("orig",)
PREVIEW_FIXED_RIGHT = ("newname",)
PREVIEW_LABELS = {
    "orig": "原文件名",
    "newname": "新文件名 (预览)",
    **FIELDS,
}
PREVIEW_WIDTHS = {
    "orig": 170,
    "title": 230,
    "journal": 150,
    "journal_short": 110,
    "issn": 85,
    "eissn": 85,
    "year": 55,
    "accepted_date": 90,
    "published_date": 95,
    "first_author": 95,
    "doi": 170,
    "impact_factor": 70,
    "impact_factor_year": 70,
    "jcr_quartile": 55,
    "cas_category": 95,
    "cas_quartile": 75,
    "cas_partition": 100,
    "journal_subject": 160,
    "jcr_category": 210,
    "publisher": 160,
    "five_year_if": 80,
    "oa_status": 75,
    "journal_info_url": 220,
    "newname": 330,
}


# ---------- 元数据 ----------

@dataclass
class PaperMeta:
    title: str = ""
    journal: str = ""
    journal_short: str = ""
    issn: str = ""
    eissn: str = ""
    year: str = ""
    accepted_date: str = ""
    published_date: str = ""
    first_author: str = ""
    doi: str = ""
    impact_factor: str = ""
    impact_factor_year: str = ""
    jcr_quartile: str = ""
    cas_category: str = ""
    cas_quartile: str = ""
    cas_partition: str = ""
    journal_subject: str = ""
    jcr_category: str = ""
    publisher: str = ""
    five_year_if: str = ""
    oa_status: str = ""
    journal_info_url: str = ""
    source_file: str = ""

    def get(self, name: str) -> str:
        return getattr(self, name, "") or "Unknown"


@dataclass
class JournalInfo:
    title: str = ""
    short_title: str = ""
    issn: str = ""
    eissn: str = ""
    impact_factor: str = ""
    impact_factor_year: str = ""
    jcr_quartile: str = ""
    cas_category: str = ""
    cas_quartile: str = ""
    cas_partition: str = ""
    journal_subject: str = ""
    jcr_category: str = ""
    publisher: str = ""
    five_year_if: str = ""
    oa_status: str = ""
    url: str = ""


DOI_RE = re.compile(r"10\.\d{4,9}/[-._;()/:A-Z0-9]+", re.IGNORECASE)
ACCEPTED_RE = re.compile(
    r"Accepted[^\d\n]{0,30}?"
    r"(\d{1,2}\s+[A-Za-z]+\s+\d{4}"
    r"|[A-Za-z]+\s+\d{1,2},?\s+\d{4}"
    r"|\d{4}[-/.]\d{1,2}[-/.]\d{1,2})",
    re.IGNORECASE,
)
MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dec": 12,
    "january": 1, "february": 2, "march": 3, "april": 4, "june": 6, "july": 7,
    "august": 8, "september": 9, "october": 10, "november": 11, "december": 12,
}


def extract_text(pdf: Path, pages: int = 3) -> str:
    if PdfReader is None:
        return ""
    try:
        r = PdfReader(str(pdf))
        n = min(len(r.pages), pages)
        return "\n".join((r.pages[i].extract_text() or "") for i in range(n))
    except Exception:
        return ""


def find_doi(text: str) -> str:
    if not text:
        return ""
    m = DOI_RE.search(text)
    if not m:
        return ""
    return m.group(0).rstrip(".,;)]}>")


def query_crossref(doi: str, timeout: int = 10) -> Optional[dict]:
    url = f"https://api.crossref.org/works/{urllib.parse.quote(doi)}"
    req = urllib.request.Request(
        url, headers={"User-Agent": "PDF-Renamer/1.0 (academic use)"}
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.load(resp).get("message")
    except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, ValueError):
        return None


def fetch_json(url: str, timeout: int = 10) -> Optional[dict]:
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": "PDF-Renamer/1.1 (academic use)",
            "Accept": "application/json,text/plain,*/*",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.load(resp)
    except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, ValueError):
        return None


def fetch_text(url: str, timeout: int = 10) -> str:
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "PDF-Renamer/1.1 (academic use)"},
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return resp.read().decode("utf-8", "ignore")
    except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, ValueError):
        return ""


def date_parts_to_str(parts) -> str:
    if not parts:
        return ""
    p = parts[0] if isinstance(parts, list) and parts and isinstance(parts[0], list) else parts
    if not p:
        return ""
    try:
        if len(p) >= 3:
            return f"{int(p[0]):04d}-{int(p[1]):02d}-{int(p[2]):02d}"
        if len(p) == 2:
            return f"{int(p[0]):04d}-{int(p[1]):02d}"
        return f"{int(p[0]):04d}"
    except (TypeError, ValueError):
        return ""


def normalize_date(s: str) -> str:
    if not s:
        return ""
    s = s.strip()
    if re.match(r"^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}", s):
        m = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", s)
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    m = re.match(r"^(\d{1,2})[-\s]+([A-Za-z]+)[-\s,]+(\d{4})", s)
    if m and MONTHS.get(m.group(2).lower()):
        return (
            f"{int(m.group(3)):04d}-{MONTHS[m.group(2).lower()]:02d}-{int(m.group(1)):02d}"
        )
    m = re.match(r"^([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})", s)
    if m and MONTHS.get(m.group(1).lower()):
        return (
            f"{int(m.group(3)):04d}-{MONTHS[m.group(1).lower()]:02d}-{int(m.group(2)):02d}"
        )
    return s


def clean_html(s: str) -> str:
    if not s:
        return ""
    s = re.sub(r"<br\s*/?>", " ", str(s), flags=re.IGNORECASE)
    s = re.sub(r"<[^>]+>", "", s)
    s = s.replace("&nbsp;", " ")
    return re.sub(r"\s+", " ", s).strip()


def normalize_journal_key(s: str) -> str:
    s = clean_html(s).lower()
    s = s.replace("&", " and ")
    return re.sub(r"[^a-z0-9]+", "", s)


def normalize_issn(s: str) -> str:
    s = (s or "").upper().strip()
    m = re.search(r"\b\d{4}-?\d{3}[\dX]\b", s)
    if not m:
        return ""
    raw = m.group(0).replace("-", "")
    return f"{raw[:4]}-{raw[4:]}"


def first_nonempty(*values: str) -> str:
    for value in values:
        if value:
            return value
    return ""


def format_cas_partition(category: str, quartile: str) -> str:
    category = clean_html(category)
    quartile = clean_html(quartile)
    if category and quartile:
        return f"{category}{quartile}"
    return category or quartile


def journal_query_candidates(journal: str, issn: str = "", eissn: str = "") -> list[str]:
    seen = set()
    candidates = [
        normalize_issn(issn),
        normalize_issn(eissn),
        clean_html(journal),
    ]
    if journal:
        swapped = re.sub(r"\band\b", "&", journal, flags=re.IGNORECASE)
        expanded = journal.replace("&", " and ")
        candidates.extend([swapped, expanded])
        words = re.findall(r"[A-Za-z0-9]+", journal)
        useful = [w for w in words if w.lower() not in JOURNAL_QUERY_STOPWORDS]
        if useful:
            candidates.append(" ".join(useful[:2]))
            candidates.append(useful[-1])
            candidates.extend(useful)

    cleaned = []
    for candidate in candidates:
        candidate = clean_html(candidate)
        if not candidate or candidate.lower() == "n/a":
            continue
        key = candidate.lower()
        if key not in seen:
            cleaned.append(candidate)
            seen.add(key)
    return cleaned


def latest_if_from_item(item: dict) -> tuple[str, str]:
    pairs = []
    for year, value in re.findall(r"(\d{4})\s*:\s*([^,]+)", item.get("if_multi", "")):
        value = clean_html(value)
        if value and value.upper() != "N/A":
            pairs.append((year, value))
    if pairs:
        year, value = max(pairs, key=lambda x: int(x[0]))
        return value, year

    candidates = []
    for key, value in item.items():
        m = re.fullmatch(r"IF(\d{4})", str(key))
        value = clean_html(str(value)) if value is not None else ""
        if m and value and value.upper() not in {"N/A", "0.0", "0.00"}:
            candidates.append((m.group(1), value))
    if candidates:
        year, value = max(candidates, key=lambda x: int(x[0]))
        return value, year
    return "", ""


def journal_info_from_item(item: dict) -> JournalInfo:
    impact_factor, impact_factor_year = latest_if_from_item(item)
    cas_category = clean_html(item.get("jcr11", ""))
    cas_quartile = clean_html(item.get("jcr12", ""))
    url = clean_html(item.get("titleurl", ""))
    if url.startswith("/"):
        url = "https://www.iikx.com" + url
    return JournalInfo(
        title=clean_html(item.get("title", "")),
        short_title=clean_html(item.get("smalltitle", "")),
        issn=normalize_issn(item.get("issn", "")),
        eissn=normalize_issn(item.get("eissn", "")),
        impact_factor=impact_factor,
        impact_factor_year=impact_factor_year,
        jcr_quartile=clean_html(item.get("zky2020", "")),
        cas_category=cas_category,
        cas_quartile=cas_quartile,
        cas_partition=format_cas_partition(cas_category, cas_quartile),
        journal_subject=clean_html(item.get("jcr21", "")),
        jcr_category=clean_html(item.get("category", "")),
        publisher=clean_html(item.get("onlinetime", "")),
        url=url,
    )


def parse_journal_detail(info: JournalInfo, timeout: int = 10) -> JournalInfo:
    if not info.url:
        return info
    html = fetch_text(info.url, timeout=timeout)
    if not html:
        return info

    def item_value(label: str) -> str:
        pattern = rf"<li[^>]*>\s*<strong>{re.escape(label)}：</strong>(.*?)</li>"
        m = re.search(pattern, html, flags=re.IGNORECASE | re.DOTALL)
        return clean_html(m.group(1)) if m else ""

    info.five_year_if = item_value("5年平均影响因子")
    info.oa_status = item_value("是否OA开放访问")
    info.publisher = first_nonempty(info.publisher, item_value("出版商"))
    info.jcr_category = first_nonempty(info.jcr_category, item_value("学科分类与版本"))
    cas_info = parse_latest_cas_table(html)
    if cas_info:
        info.cas_category = cas_info.get("cas_category", info.cas_category)
        info.cas_quartile = cas_info.get("cas_quartile", info.cas_quartile)
        info.cas_partition = format_cas_partition(info.cas_category, info.cas_quartile)
        info.journal_subject = cas_info.get("journal_subject", info.journal_subject)
    return info


def parse_latest_cas_table(html: str) -> dict[str, str]:
    idx = html.find("2025年3月升级版")
    if idx < 0:
        idx = html.find("中科院JCR分区")
    if idx < 0:
        return {}
    next_idx = html.find("2023年12月升级版", idx + 1)
    end_idx = next_idx if next_idx > idx else idx + 3000
    block = html[idx:end_idx]
    if "未收录" in clean_html(block):
        return {}

    category = ""
    quartile = ""
    m = re.search(r"<td>\s*([^<>]+?)\s*<span>\s*([^<>]+?)\s*</span>", block, re.DOTALL)
    if m:
        category = clean_html(m.group(1))
        quartile = clean_html(m.group(2))

    subjects = []
    subject_matches = re.findall(
        r"<td>\s*([A-Z][A-Z\s&/,\-]+?)\s*<br\s*/?>\s*([^<>]+?)\s*</td>\s*"
        r"<td>\s*<span>\s*([^<>]+?)\s*</span>",
        block,
        flags=re.IGNORECASE | re.DOTALL,
    )
    for _english, chinese, sub_quartile in subject_matches:
        chinese = clean_html(chinese)
        sub_quartile = clean_html(sub_quartile)
        if chinese and sub_quartile:
            subjects.append(f"{chinese}{sub_quartile}")
        elif chinese:
            subjects.append(chinese)

    return {
        "cas_category": category,
        "cas_quartile": quartile,
        "journal_subject": "; ".join(subjects),
    }


def score_journal_item(item: dict, journal: str, issn: str, eissn: str) -> int:
    score = 0
    item_issn = normalize_issn(item.get("issn", ""))
    item_eissn = normalize_issn(item.get("eissn", ""))
    wanted_issns = {x for x in (normalize_issn(issn), normalize_issn(eissn)) if x}
    if wanted_issns and (item_issn in wanted_issns or item_eissn in wanted_issns):
        score += 100

    q = normalize_journal_key(journal)
    title = normalize_journal_key(item.get("title", ""))
    short_title = normalize_journal_key(item.get("smalltitle", ""))
    if q and title == q:
        score += 60
    elif q and short_title == q:
        score += 50
    elif q and (q in title or title in q):
        score += 20
    elif q and (q in short_title or short_title in q):
        score += 15
    return score


def query_iikx_journal(
    journal: str = "", issn: str = "", eissn: str = "", timeout: int = 10
) -> Optional[JournalInfo]:
    queries = journal_query_candidates(journal, issn, eissn)
    if not queries:
        return None

    best_item = None
    best_score = 0
    for query in queries:
        params = {
            "page": "0",
            "title": query,
            "smalltitle": query,
            "issn": query,
            "eissn": query,
            "andor": "or",
            "classid": IIKX_ALL_CLASSIDS,
            "orderby": "IF2024",
            "ph": "1",
        }
        data = fetch_json(IIKX_API_URL + "?" + urllib.parse.urlencode(params), timeout=timeout)
        if not data or data.get("code") != 200:
            continue
        for item in ((data.get("result") or {}).get("data") or []):
            score = score_journal_item(item, journal, issn, eissn)
            if clean_html(item.get("jcr11", "")) and clean_html(item.get("jcr21", "")):
                score += 3
            if score > best_score:
                best_item = item
                best_score = score
        if best_score >= 100:
            break

    if not best_item or best_score <= 0:
        return None
    return parse_journal_detail(journal_info_from_item(best_item), timeout=timeout)


def apply_journal_info(meta: PaperMeta, info: JournalInfo):
    meta.journal = first_nonempty(meta.journal, info.title)
    meta.journal_short = first_nonempty(meta.journal_short, info.short_title)
    meta.issn = first_nonempty(meta.issn, info.issn)
    meta.eissn = first_nonempty(meta.eissn, info.eissn)
    meta.impact_factor = info.impact_factor
    meta.impact_factor_year = info.impact_factor_year
    meta.jcr_quartile = info.jcr_quartile
    meta.cas_category = info.cas_category
    meta.cas_quartile = info.cas_quartile
    meta.cas_partition = info.cas_partition
    meta.journal_subject = info.journal_subject
    meta.jcr_category = info.jcr_category
    meta.publisher = info.publisher
    meta.five_year_if = info.five_year_if
    meta.oa_status = info.oa_status
    meta.journal_info_url = info.url


def extract_metadata(pdf: Path) -> PaperMeta:
    meta = PaperMeta(source_file=str(pdf))
    text = extract_text(pdf, pages=3)

    # 1) DOI -> CrossRef (最准确)
    doi = find_doi(text)
    if doi:
        meta.doi = doi
        cr = query_crossref(doi)
        if cr:
            if cr.get("title"):
                meta.title = cr["title"][0]
            ct = [t for t in (cr.get("container-title") or []) if t]
            sct = [t for t in (cr.get("short-container-title") or []) if t]
            if ct:
                # CrossRef 有时给多条, 最长的通常是完整全称
                meta.journal = max(ct, key=len)
            if sct:
                meta.journal_short = sct[0]
            elif len(ct) > 1:
                meta.journal_short = min(ct, key=len)
            issns = [normalize_issn(x) for x in (cr.get("ISSN") or [])]
            issns = [x for x in issns if x]
            if issns:
                meta.issn = issns[0]
                if len(issns) > 1:
                    meta.eissn = issns[1]
            for item in cr.get("issn-type") or []:
                value = normalize_issn(item.get("value", ""))
                if not value:
                    continue
                if item.get("type") == "print":
                    meta.issn = value
                elif item.get("type") == "electronic":
                    meta.eissn = value

            for key in ("published-print", "published-online", "issued"):
                d = cr.get(key, {}).get("date-parts")
                if d:
                    meta.published_date = date_parts_to_str(d)
                    meta.year = meta.published_date[:4]
                    break

            acc = cr.get("accepted", {}).get("date-parts")
            if acc:
                meta.accepted_date = date_parts_to_str(acc)
            else:
                for a in cr.get("assertion") or []:
                    label = (str(a.get("label", "")) + str(a.get("name", ""))).lower()
                    if "accept" in label:
                        v = a.get("value", "")
                        if v:
                            meta.accepted_date = normalize_date(v)
                            break

            authors = cr.get("author") or []
            if authors:
                fam = authors[0].get("family") or ""
                giv = authors[0].get("given") or ""
                meta.first_author = fam or giv

    info = query_iikx_journal(meta.journal, meta.issn, meta.eissn)
    if info:
        apply_journal_info(meta, info)

    # 2) 文本兜底找 Accepted
    if not meta.accepted_date:
        m = ACCEPTED_RE.search(text)
        if m:
            meta.accepted_date = normalize_date(m.group(1))

    # 3) PDF 元数据兜底找标题
    if not meta.title and PdfReader is not None:
        try:
            pdf_info = PdfReader(str(pdf)).metadata or {}
            t = pdf_info.get("/Title", "") if hasattr(pdf_info, "get") else ""
            if t and len(str(t)) > 5:
                meta.title = str(t)
        except Exception:
            pass

    return meta


# ---------- 文件名构造 ----------

def sanitize(s: str, max_len: int = 100) -> str:
    if not s:
        return ""
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", str(s))
    s = re.sub(r"\s+", " ", s).strip()
    return s[:max_len].rstrip()


def build_filename(meta: PaperMeta, tokens: list, ext: str = ".pdf") -> str:
    parts = []
    for t in tokens:
        if t["type"] == "field":
            parts.append(sanitize(meta.get(t["value"])))
        else:
            parts.append(t["value"])
    name = "".join(parts).strip(" _-.") or "Unnamed"
    return name + ext


def preview_value(field: str, path: str, meta: PaperMeta, tokens: list) -> str:
    if field == "orig":
        return Path(path).name
    if field == "newname":
        return build_filename(meta, tokens)
    if field == "title":
        return meta.title[:80]
    return meta.get(field)


def tree_values(path: str, meta: PaperMeta, tokens: list, preview_fields: list) -> tuple:
    fields = list(PREVIEW_FIXED_LEFT) + preview_fields + list(PREVIEW_FIXED_RIGHT)
    return tuple(preview_value(field, path, meta, tokens) for field in fields)


def sort_key_for_value(value: str):
    value = "" if value is None else str(value).strip()
    if not value or value == "Unknown":
        return (3, "")
    m = re.match(r"^(\d{4})(?:[-/.](\d{1,2}))?(?:[-/.](\d{1,2}))?", value)
    if m:
        year = int(m.group(1))
        month = int(m.group(2) or 0)
        day = int(m.group(3) or 0)
        return (1, (year, month, day))
    normalized = value.replace(",", "")
    m = re.match(r"^-?\d+(?:\.\d+)?", normalized)
    if m:
        try:
            return (0, float(m.group(0)))
        except ValueError:
            pass
    return (2, value.casefold())


def excel_col_name(index: int) -> str:
    name = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        name = chr(65 + rem) + name
    return name


def xlsx_cell(value, row: int, col: int) -> str:
    ref = f"{excel_col_name(col)}{row}"
    value = "" if value is None else str(value)
    if value == "":
        return f'<c r="{ref}"/>'
    numeric = re.fullmatch(r"-?\d+(?:\.\d+)?", value.replace(",", ""))
    if numeric and not re.search(r"^0\d", value):
        return f'<c r="{ref}"><v>{value.replace(",", "")}</v></c>'
    return (
        f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">'
        f"{escape(value)}"
        f"</t></is></c>"
    )


def write_xlsx(path: Path, headers: list[str], rows: list[list[str]]):
    created = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    sheet_rows = []
    all_rows = [headers] + rows
    for r_idx, row in enumerate(all_rows, 1):
        cells = "".join(xlsx_cell(value, r_idx, c_idx) for c_idx, value in enumerate(row))
        sheet_rows.append(f'<row r="{r_idx}">{cells}</row>')

    sheet_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetViews><sheetView workbookViewId="0"/></sheetViews>
<sheetFormatPr defaultRowHeight="15"/>
<sheetData>{''.join(sheet_rows)}</sheetData>
<autoFilter ref="A1:{excel_col_name(max(len(headers) - 1, 0))}{max(len(all_rows), 1)}"/>
</worksheet>'''
    workbook_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="文件预览" sheetId="1" r:id="rId1"/></sheets>
</workbook>'''
    rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''
    workbook_rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>'''
    content_types_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''
    core_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:dcmitype="http://purl.org/dc/dcmitype/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:creator>PDF Renamer</dc:creator>
<cp:lastModifiedBy>PDF Renamer</cp:lastModifiedBy>
<dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">{created}</dcterms:modified>
</cp:coreProperties>'''
    app_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>PDF Renamer</Application>
</Properties>'''

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml)
        zf.writestr("_rels/.rels", rels_xml)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        zf.writestr("docProps/core.xml", core_xml)
        zf.writestr("docProps/app.xml", app_xml)


# ---------- GUI ----------

class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("PDF 文献一键命名工具")
        root.geometry("1360x820")
        root.minsize(1100, 760)
        root.configure(background=THEME["bg"])

        self._setup_icon()
        self._setup_style()

        self.folder = tk.StringVar()
        self.tokens, self.preview_fields = self._load_config()
        self.metas: dict[str, PaperMeta] = {}
        self.sort_field = ""
        self.sort_reverse = False
        self._scan_thread: Optional[threading.Thread] = None

        self._build_ui()
        self._refresh_pattern_display()
        self._refresh_preview_fields_display()
        self._configure_preview_columns()

    # ---------- 样式 ----------
    def _setup_icon(self):
        icon = resource_path("assets/pdf_renamer.ico")
        if icon.exists():
            try:
                self.root.iconbitmap(default=str(icon))
            except tk.TclError:
                pass

    def _setup_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        bg = THEME["bg"]
        panel = THEME["panel"]
        text = THEME["text"]
        muted = THEME["muted"]
        border = THEME["border"]
        primary = THEME["primary"]
        primary_hover = THEME["primary_hover"]
        primary_press = THEME["primary_press"]
        accent = THEME["accent"]
        accent_hover = THEME["accent_hover"]
        danger = THEME["danger"]
        danger_hover = THEME["danger_hover"]
        btn_bg = THEME["btn_bg"]
        btn_hover = THEME["btn_hover"]
        btn_press = THEME["btn_press"]

        style.configure(".", background=bg, foreground=text, font=UI_FONT)

        # Frame: 窗口背景 & 卡片背景
        style.configure("TFrame", background=bg)
        style.configure("Card.TFrame", background=panel)

        # Label: 几种字号 + 卡片背景变体
        style.configure("TLabel", background=bg, foreground=text, font=UI_FONT)
        style.configure("Muted.TLabel", background=bg, foreground=muted, font=UI_FONT_SMALL)
        style.configure("Title.TLabel", background=bg, foreground=text, font=UI_FONT_TITLE)
        style.configure("Card.TLabel", background=panel, foreground=text, font=UI_FONT)
        style.configure("CardMuted.TLabel", background=panel, foreground=muted,
                        font=UI_FONT_SMALL)
        style.configure("CardSubtitle.TLabel", background=panel, foreground=text,
                        font=UI_FONT_SUBTITLE)
        style.configure("Pattern.TLabel",
                        background=panel, foreground=primary,
                        font=MONO_FONT, padding=(10, 6),
                        borderwidth=0, relief="flat")
        style.configure("Status.TLabel",
                        background=bg, foreground=muted,
                        font=UI_FONT_SMALL, padding=(0, 4))

        # LabelFrame: 仍然提供但默认极简
        style.configure("TLabelframe", background=bg, bordercolor=border,
                        borderwidth=0, relief="flat")
        style.configure("TLabelframe.Label",
                        background=bg, foreground=text, font=UI_FONT_BOLD)

        # 按钮:
        # - 次按钮:浅灰底,文字深色
        style.configure("TButton",
                        background=btn_bg, foreground=text,
                        font=UI_FONT, padding=(14, 7),
                        borderwidth=0, focusthickness=0, relief="flat")
        style.map("TButton",
                  background=[("active", btn_hover), ("pressed", btn_press),
                              ("disabled", btn_bg)],
                  foreground=[("disabled", muted)])

        # - 主按钮:Apple 蓝
        style.configure("Accent.TButton",
                        background=primary, foreground="#ffffff",
                        font=UI_FONT_BOLD, padding=(16, 8),
                        borderwidth=0, focusthickness=0, relief="flat")
        style.map("Accent.TButton",
                  background=[("active", primary_hover),
                              ("pressed", primary_press),
                              ("disabled", "#a8c9f0")])

        # - 成功按钮:iOS 绿 (Excel 等)
        style.configure("Success.TButton",
                        background=accent, foreground="#ffffff",
                        font=UI_FONT_BOLD, padding=(16, 8),
                        borderwidth=0, focusthickness=0, relief="flat")
        style.map("Success.TButton",
                  background=[("active", accent_hover), ("pressed", accent_hover)])

        # - 危险按钮:iOS 红
        style.configure("Danger.TButton",
                        background=danger, foreground="#ffffff",
                        font=UI_FONT_BOLD, padding=(14, 7),
                        borderwidth=0, focusthickness=0, relief="flat")
        style.map("Danger.TButton",
                  background=[("active", danger_hover), ("pressed", danger_hover)])

        # 输入框:扁平,聚焦时蓝色边框
        style.configure("TEntry",
                        fieldbackground=panel, foreground=text,
                        bordercolor=border, lightcolor=border, darkcolor=border,
                        borderwidth=1, padding=8, relief="flat")
        style.map("TEntry",
                  bordercolor=[("focus", primary)],
                  lightcolor=[("focus", primary)],
                  darkcolor=[("focus", primary)])

        # 分隔线
        style.configure("TSeparator", background=border)

        # 滚动条:极简
        style.configure("Vertical.TScrollbar",
                        background=bg, troughcolor=bg, bordercolor=bg,
                        arrowcolor=muted, gripcount=0, borderwidth=0,
                        relief="flat", arrowsize=14)
        style.configure("Horizontal.TScrollbar",
                        background=bg, troughcolor=bg, bordercolor=bg,
                        arrowcolor=muted, gripcount=0, borderwidth=0,
                        relief="flat", arrowsize=14)

        # 表格
        style.configure("Treeview",
                        background=panel, fieldbackground=panel, foreground=text,
                        rowheight=28, font=UI_FONT, borderwidth=0, relief="flat")
        style.configure("Treeview.Heading",
                        background=THEME["heading_bg"], foreground=muted,
                        font=UI_FONT_BOLD, padding=(8, 8), relief="flat",
                        borderwidth=0)
        style.map("Treeview.Heading",
                  background=[("active", btn_hover)])
        style.map("Treeview",
                  background=[("selected", THEME["select_bg"])],
                  foreground=[("selected", THEME["select_fg"])])

    def _styled_listbox(self, parent, **kw) -> tk.Listbox:
        defaults = dict(
            bg=THEME["panel"],
            fg=THEME["text"],
            selectbackground=THEME["select_bg"],
            selectforeground=THEME["select_fg"],
            highlightthickness=1,
            highlightbackground=THEME["border"],
            highlightcolor=THEME["primary"],
            borderwidth=0,
            relief="flat",
            font=UI_FONT,
            activestyle="none",
            exportselection=False,
        )
        defaults.update(kw)
        return tk.Listbox(parent, **defaults)

    def _enable_dnd_reorder(self, listbox: tk.Listbox, items: list, on_change):
        """在 Listbox 上启用拖拽重排,items 是和 listbox 一一对应的列表。
        on_change() 在每次顺序变化后调用以重绘 listbox 并保存。
        """
        state = {"idx": None}

        def press(event):
            i = listbox.nearest(event.y)
            if 0 <= i < len(items):
                state["idx"] = i
                listbox.configure(cursor="fleur")

        def motion(event):
            src = state["idx"]
            if src is None:
                return
            i = listbox.nearest(event.y)
            if i < 0 or i >= len(items) or i == src:
                return
            item = items.pop(src)
            items.insert(i, item)
            state["idx"] = i
            on_change()
            listbox.selection_clear(0, "end")
            listbox.selection_set(i)
            listbox.activate(i)

        def release(_e):
            state["idx"] = None
            listbox.configure(cursor="")

        listbox.bind("<ButtonPress-1>", press, add="+")
        listbox.bind("<B1-Motion>", motion, add="+")
        listbox.bind("<ButtonRelease-1>", release, add="+")

    def _on_tokens_reordered(self):
        self._refresh_pattern_display()
        self._save_config()
        self._refresh_preview()

    def _on_preview_reordered(self):
        self._refresh_preview_fields_display()
        self._configure_preview_columns()
        self._save_config()

    # config
    def _load_config(self) -> tuple[list, list]:
        if CONFIG_PATH.exists():
            try:
                data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
                if isinstance(data, list):
                    return data, DEFAULT_PREVIEW_FIELDS.copy()
                if isinstance(data, dict):
                    tokens = data.get("tokens")
                    preview_fields = data.get("preview_fields")
                    if not isinstance(tokens, list):
                        tokens = DEFAULT_TOKENS.copy()
                    if not isinstance(preview_fields, list):
                        preview_fields = DEFAULT_PREVIEW_FIELDS.copy()
                    preview_fields = [
                        f for f in preview_fields
                        if isinstance(f, str) and f in FIELDS
                    ]
                    if not preview_fields:
                        preview_fields = DEFAULT_PREVIEW_FIELDS.copy()
                    return tokens, preview_fields
            except Exception:
                pass
        return DEFAULT_TOKENS.copy(), DEFAULT_PREVIEW_FIELDS.copy()

    def _save_config(self):
        try:
            data = {
                "tokens": self.tokens,
                "preview_fields": self.preview_fields,
            }
            CONFIG_PATH.write_text(
                json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
            )
        except Exception:
            pass

    # build UI
    def _build_ui(self):
        # ===== 顶部:标题 + 文件夹选择 =====
        header = ttk.Frame(self.root, padding=(20, 16, 20, 4))
        header.pack(fill="x")
        ttk.Label(header, text="PDF 文献一键命名工具",
                  style="Title.TLabel").pack(side="left")

        folder_row = ttk.Frame(self.root, padding=(20, 4, 20, 12))
        folder_row.pack(fill="x")
        ttk.Entry(folder_row, textvariable=self.folder).pack(
            side="left", fill="x", expand=True, padx=(0, 8)
        )
        ttk.Button(folder_row, text="选择文件夹", style="Accent.TButton",
                   command=self._choose_folder).pack(side="left")

        # ===== 底部:状态栏 + 操作 (先 pack 确保不被挤掉) =====
        footer = ttk.Frame(self.root, padding=(20, 6, 20, 14))
        footer.pack(side="bottom", fill="x")
        self.status = ttk.Label(footer, text="就绪", style="Status.TLabel")
        self.status.pack(side="left", fill="x", expand=True)
        ttk.Button(footer, text="导出 Excel",
                   command=self._export_excel).pack(side="right", padx=(8, 0))
        ttk.Button(footer, text="执行重命名", style="Accent.TButton",
                   command=self._do_rename).pack(side="right")

        # ===== 卡片 1:命名规则 =====
        card1 = ttk.Frame(self.root, style="Card.TFrame", padding=10)
        card1.pack(fill="x", padx=20, pady=(0, 8))

        head1 = ttk.Frame(card1, style="Card.TFrame")
        head1.pack(fill="x", pady=(0, 8))
        ttk.Label(head1, text="命名规则",
                  style="CardSubtitle.TLabel").pack(side="left")
        ttk.Label(head1,
                  text="  双击左侧添加 · 拖动右侧重新排列",
                  style="CardMuted.TLabel").pack(side="left", padx=(6, 0))
        # 规则预览即时显示 (与标题同一行,极简)
        self.pattern_lbl = ttk.Label(head1, text="", style="Pattern.TLabel")
        self.pattern_lbl.pack(side="right")
        ttk.Label(head1, text="规则预览",
                  style="CardMuted.TLabel").pack(side="right", padx=(0, 6))

        body1 = ttk.Frame(card1, style="Card.TFrame")
        body1.pack(fill="x")

        # 左:可用字段
        left1 = ttk.Frame(body1, style="Card.TFrame")
        left1.pack(side="left", fill="y")
        ttk.Label(left1, text="可用字段", style="CardMuted.TLabel").pack(anchor="w")
        self.field_box = self._styled_listbox(left1, height=7, width=30)
        for label in FIELDS.values():
            self.field_box.insert("end", label)
        self.field_box.pack(pady=(4, 0))
        self.field_box.bind("<Double-Button-1>", lambda _e: self._add_field())

        # 中:操作按钮 (紧凑布局)
        mid1 = ttk.Frame(body1, style="Card.TFrame")
        mid1.pack(side="left", padx=12, fill="y")
        ttk.Button(mid1, text="＋ 添加字段", style="Accent.TButton",
                   command=self._add_field).pack(fill="x", pady=(8, 4))

        sep_row = ttk.Frame(mid1, style="Card.TFrame")
        sep_row.pack(fill="x", pady=2)
        ttk.Label(sep_row, text="分隔符", style="CardMuted.TLabel").pack(side="left")
        self.sep_var = tk.StringVar(value="_")
        ttk.Entry(sep_row, textvariable=self.sep_var, width=6).pack(
            side="right", fill="x", expand=True, padx=(8, 0)
        )

        ttk.Button(mid1, text="＋ 添加分隔符",
                   command=self._add_sep).pack(fill="x", pady=(2, 8))
        ttk.Button(mid1, text="× 删除", style="Danger.TButton",
                   command=self._remove).pack(fill="x", pady=1)
        ttk.Button(mid1, text="清空",
                   command=self._clear).pack(fill="x", pady=1)

        # 右:当前命名顺序 (拖动重排)
        right1 = ttk.Frame(body1, style="Card.TFrame")
        right1.pack(side="left", fill="both", expand=True)
        ttk.Label(right1, text="当前命名顺序",
                  style="CardMuted.TLabel").pack(anchor="w")
        self.order_box = self._styled_listbox(right1, height=7)
        self.order_box.pack(fill="both", expand=True, pady=(4, 0))
        self._enable_dnd_reorder(self.order_box, self.tokens, self._on_tokens_reordered)

        # ===== 卡片 2:预览信息 =====
        card2 = ttk.Frame(self.root, style="Card.TFrame", padding=10)
        card2.pack(fill="x", padx=20, pady=(0, 8))

        head2 = ttk.Frame(card2, style="Card.TFrame")
        head2.pack(fill="x", pady=(0, 8))
        ttk.Label(head2, text="预览信息",
                  style="CardSubtitle.TLabel").pack(side="left")
        ttk.Label(head2,
                  text="  双击左侧添加 · 拖动右侧重新排列",
                  style="CardMuted.TLabel").pack(side="left", padx=(6, 0))

        body2 = ttk.Frame(card2, style="Card.TFrame")
        body2.pack(fill="x")

        left2 = ttk.Frame(body2, style="Card.TFrame")
        left2.pack(side="left", fill="y")
        ttk.Label(left2, text="可添加信息", style="CardMuted.TLabel").pack(anchor="w")
        self.preview_field_box = self._styled_listbox(left2, height=4, width=30)
        for label in FIELDS.values():
            self.preview_field_box.insert("end", label)
        self.preview_field_box.pack(pady=(4, 0))
        self.preview_field_box.bind(
            "<Double-Button-1>", lambda _e: self._add_preview_field()
        )

        mid2 = ttk.Frame(body2, style="Card.TFrame")
        mid2.pack(side="left", padx=12, fill="y")
        ttk.Button(mid2, text="＋ 添加", style="Accent.TButton",
                   command=self._add_preview_field).pack(fill="x", pady=(8, 4))
        ttk.Button(mid2, text="× 删除", style="Danger.TButton",
                   command=self._remove_preview_field).pack(fill="x", pady=1)
        ttk.Button(mid2, text="↺ 默认",
                   command=self._reset_preview_fields).pack(fill="x", pady=1)

        right2 = ttk.Frame(body2, style="Card.TFrame")
        right2.pack(side="left", fill="both", expand=True)
        ttk.Label(right2, text="当前预览顺序",
                  style="CardMuted.TLabel").pack(anchor="w")
        self.preview_order_box = self._styled_listbox(right2, height=4)
        self.preview_order_box.pack(fill="both", expand=True, pady=(4, 0))
        self._enable_dnd_reorder(
            self.preview_order_box, self.preview_fields, self._on_preview_reordered
        )

        # ===== 卡片 3:文件预览表 =====
        table_card = ttk.Frame(self.root, style="Card.TFrame", padding=6)
        table_card.pack(fill="both", expand=True, padx=20, pady=(0, 8))

        self.tree = ttk.Treeview(table_card, show="headings")
        self.tree.tag_configure("odd", background=THEME["stripe"])
        self.tree.tag_configure("even", background=THEME["panel"])

        vsb = ttk.Scrollbar(table_card, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_card, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        table_card.rowconfigure(0, weight=1)
        table_card.columnconfigure(0, weight=1)

    # pattern ops
    def _refresh_pattern_display(self):
        self.order_box.delete(0, "end")
        for t in self.tokens:
            if t["type"] == "field":
                self.order_box.insert("end", f"[字段] {FIELDS.get(t['value'], t['value'])}")
            else:
                self.order_box.insert("end", f"[分隔] \"{t['value']}\"")
        text = "".join(
            "{" + t["value"] + "}" if t["type"] == "field" else t["value"]
            for t in self.tokens
        ) or "(空)"
        self.pattern_lbl.config(text=text)

    def _add_field(self):
        sel = self.field_box.curselection()
        if not sel:
            messagebox.showinfo("提示", "请先在左侧选择一个字段")
            return
        key = list(FIELDS.keys())[sel[0]]
        self.tokens.append({"type": "field", "value": key})
        self._refresh_pattern_display()
        self._save_config()
        self._refresh_preview()

    def _add_sep(self):
        s = self.sep_var.get()
        if not s:
            return
        self.tokens.append({"type": "sep", "value": s})
        self._refresh_pattern_display()
        self._save_config()
        self._refresh_preview()

    def _remove(self):
        sel = self.order_box.curselection()
        if not sel:
            return
        del self.tokens[sel[0]]
        self._refresh_pattern_display()
        self._save_config()
        self._refresh_preview()

    def _clear(self):
        if messagebox.askyesno("确认", "清空全部命名规则?"):
            self.tokens = []
            self._refresh_pattern_display()
            self._save_config()
            self._refresh_preview()

    # preview field ops
    def _preview_columns(self) -> list:
        return list(PREVIEW_FIXED_LEFT) + self.preview_fields + list(PREVIEW_FIXED_RIGHT)

    def _refresh_preview_fields_display(self):
        self.preview_order_box.delete(0, "end")
        for field in self.preview_fields:
            self.preview_order_box.insert("end", PREVIEW_LABELS.get(field, field))

    def _configure_preview_columns(self):
        cols = self._preview_columns()
        self.tree.configure(columns=cols)
        for c in cols:
            self._set_heading(c)
            self.tree.column(c, width=PREVIEW_WIDTHS.get(c, 120), stretch=True)
        self._refresh_preview()

    def _set_heading(self, field: str):
        label = PREVIEW_LABELS.get(field, field)
        if field == self.sort_field:
            label += " ↓" if self.sort_reverse else " ↑"
        self.tree.heading(field, text=label, command=lambda f=field: self._sort_by(f))

    def _sort_by(self, field: str):
        if self.sort_field == field:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_field = field
            self.sort_reverse = False
        for c in self._preview_columns():
            self._set_heading(c)
        self._apply_sort()

    def _sort_item_key(self, iid: str):
        meta = self.metas.get(iid)
        if not meta:
            return (3, "")
        return sort_key_for_value(preview_value(self.sort_field, iid, meta, self.tokens))

    def _apply_sort(self):
        if not self.sort_field or not self.metas:
            return
        sorted_iids = sorted(
            self.tree.get_children(""),
            key=self._sort_item_key,
            reverse=self.sort_reverse,
        )
        for index, iid in enumerate(sorted_iids):
            self.tree.move(iid, "", index)
        self._reapply_stripes()

    def _add_preview_field(self):
        sel = self.preview_field_box.curselection()
        if not sel:
            messagebox.showinfo("提示", "请先在左侧选择一个预览字段")
            return
        key = list(FIELDS.keys())[sel[0]]
        if key in self.preview_fields:
            messagebox.showinfo("提示", "该字段已经在预览信息中")
            return
        self.preview_fields.append(key)
        self._refresh_preview_fields_display()
        self._configure_preview_columns()
        self._save_config()

    def _remove_preview_field(self):
        sel = self.preview_order_box.curselection()
        if not sel:
            return
        del self.preview_fields[sel[0]]
        if self.sort_field and self.sort_field not in self._preview_columns():
            self.sort_field = ""
        self._refresh_preview_fields_display()
        self._configure_preview_columns()
        self._save_config()

    def _reset_preview_fields(self):
        self.preview_fields = DEFAULT_PREVIEW_FIELDS.copy()
        if self.sort_field and self.sort_field not in self._preview_columns():
            self.sort_field = ""
        self._refresh_preview_fields_display()
        self._configure_preview_columns()
        self._save_config()

    # actions
    def _choose_folder(self):
        d = filedialog.askdirectory()
        if d:
            self.folder.set(d)
            self._scan()

    def _scan(self):
        d = self.folder.get()
        if not d or not Path(d).is_dir():
            messagebox.showerror("错误", "请先选择有效文件夹")
            return
        if self._scan_thread and self._scan_thread.is_alive():
            messagebox.showinfo("提示", "扫描仍在进行中，请稍候")
            return
        pdfs = sorted(Path(d).glob("*.pdf"))
        if not pdfs:
            messagebox.showinfo("提示", "该文件夹下未找到 PDF")
            return
        self.tree.delete(*self.tree.get_children())
        self.metas.clear()

        total = len(pdfs)

        def worker():
            for i, p in enumerate(pdfs, 1):
                # 提取在工作线程内执行,UI 更新通过 after 派发到主线程
                self.root.after(
                    0,
                    lambda i=i, name=p.name: self.status.config(
                        text=f"提取中 {i}/{total}: {name}"
                    ),
                )
                meta = extract_metadata(p)
                self.root.after(0, self._on_scan_one, str(p), meta)
            self.root.after(
                0, lambda: self.status.config(text=f"完成，共 {total} 个文件")
            )

        self._scan_thread = threading.Thread(target=worker, daemon=True)
        self._scan_thread.start()

    def _on_scan_one(self, path: str, meta: PaperMeta):
        """主线程回调:插入一行并更新排序、斑马纹。"""
        self.metas[path] = meta
        self.tree.insert(
            "", "end", iid=path,
            values=tree_values(path, meta, self.tokens, self.preview_fields),
        )
        self._apply_sort()
        self._reapply_stripes()

    def _reapply_stripes(self):
        for index, iid in enumerate(self.tree.get_children("")):
            self.tree.item(iid, tags=("even" if index % 2 == 0 else "odd",))

    def _refresh_preview(self):
        for path, meta in self.metas.items():
            try:
                self.tree.item(
                    path,
                    values=tree_values(path, meta, self.tokens, self.preview_fields),
                )
            except tk.TclError:
                pass
        self._apply_sort()
        self._reapply_stripes()

    def _export_excel(self):
        if not self.metas:
            messagebox.showinfo("提示", "请先扫描文件夹")
            return
        default_name = f"PDF文件预览_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        out = filedialog.asksaveasfilename(
            title="导出文件预览",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel 工作簿", "*.xlsx")],
        )
        if not out:
            return

        cols = self._preview_columns()
        headers = [PREVIEW_LABELS.get(c, c).replace(" ↑", "").replace(" ↓", "") for c in cols]
        rows = []
        for iid in self.tree.get_children(""):
            meta = self.metas.get(iid)
            if not meta:
                continue
            rows.append([
                preview_value(c, iid, meta, self.tokens)
                for c in cols
            ])
        try:
            write_xlsx(Path(out), headers, rows)
        except OSError as e:
            messagebox.showerror("导出失败", f"无法写入 Excel 文件:\n{e}")
            return
        messagebox.showinfo("导出完成", f"已导出 {len(rows)} 条记录:\n{out}")

    def _do_rename(self):
        if not self.metas:
            messagebox.showinfo("提示", "请先扫描文件夹")
            return
        if not self.tokens:
            messagebox.showerror("错误", "命名规则为空")
            return
        if not messagebox.askyesno("确认", f"对 {len(self.metas)} 个文件执行重命名?"):
            return

        ok, errs = 0, []
        new_metas = {}
        for path, meta in self.metas.items():
            p = Path(path)
            if not p.exists():
                continue
            new_path = p.with_name(build_filename(meta, self.tokens))
            if new_path == p:
                new_metas[str(p)] = meta
                continue
            # 防覆盖
            if new_path.exists():
                stem, suf = new_path.stem, new_path.suffix
                k = 1
                while new_path.exists():
                    new_path = p.with_name(f"{stem} ({k}){suf}")
                    k += 1
            try:
                p.rename(new_path)
                new_metas[str(new_path)] = meta
                ok += 1
            except OSError as e:
                errs.append(f"{p.name}: {e}")
                new_metas[str(p)] = meta

        self.metas = new_metas
        self._scan_after_rename()
        msg = f"完成,成功重命名 {ok} 个"
        if errs:
            msg += f"\n失败 {len(errs)}:\n" + "\n".join(errs[:8])
        messagebox.showinfo("结果", msg)

    def _scan_after_rename(self):
        self.tree.delete(*self.tree.get_children())
        for path, meta in self.metas.items():
            self.tree.insert(
                "", "end", iid=path,
                values=tree_values(path, meta, self.tokens, self.preview_fields),
            )
        self._apply_sort()
        self._reapply_stripes()


def main():
    if PdfReader is None:
        root = tk.Tk()
        root.withdraw()
        if messagebox.askyesno(
            "缺少依赖",
            "未检测到 pypdf 库, 是否现在自动安装?\n(将运行: pip install pypdf)",
        ):
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", "pypdf"])
                messagebox.showinfo("提示", "安装完成, 请重新启动本程序")
            except subprocess.CalledProcessError as e:
                messagebox.showerror("失败", f"安装失败: {e}\n请手动执行 pip install pypdf")
        return

    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
