#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
mendeley_cite.py  —  versi universal & cerdas
=============================================
Mengubah placeholder sitasi di Word (format apa pun: NR007, S12, REF-1, [1], [1-5], {A3}, dll.)
menjadi field Mendeley Cite asli (ADDIN CSL_CITATION), memakai metadata dari Excel
dan library lokal Mendeley Reference Manager.

Alur pakai:
    1. Taruh SATU file .docx dan SATU file .xlsx di folder ini.
    2. Jalankan Run.bat (atau python mendeley_cite.py).
    3. Hasil Word: "manuscript - Mendeley.docx"
    4. Hasil Laporan Excel: "laporan_sitasi.xlsx" (lengkap status Berhasil/Gagal dan alasannya).

Fitur unggulan:
    * Deteksi otomatis pola placeholder (huruf+angka, angka murni, tanda kurung (), [], {}, dll.)
    * Ekspansi rentang otomatis (misal [1-4] -> 1, 2, 3, 4 atau NR001-NR004 -> NR001, NR002, NR003, NR004)
    * Deteksi cerdas kolom ID, Judul, Penulis, Tahun, Jurnal, DOI di Excel
    * Pencocokan ke database lokal Mendeley Reference Manager (DOI -> Judul Persis -> Fuzzy)
    * Pengayaan metadata via Crossref API (volume, issue, page)
    * Pelaporan lengkap dalam Excel (.xlsx):
        - Sheet Ringkasan (metrik keberhasilan, info berkas)
        - Sheet Detail Sitasi Word (status Berhasil / Gagal Tidak Ada di Mendeley / Gagal Tidak Ada di Excel)
        - Sheet Perlu Tindakan (panduan penyelesaian masalah)
        - Sheet Referensi Excel Belum Dipakai
    * Dokumen asli Word tidak pernah ditimpa.
"""

from __future__ import annotations

import argparse
import collections
import copy
import csv
import difflib
import json
import os
import re
import shutil
import sqlite3
import sys
import tempfile
import time
import urllib.parse
import urllib.request
import zipfile
from datetime import datetime
from pathlib import Path

import lxml.etree as ET
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd

# ==========================================================================
# OOXML Constants & Helper
# ==========================================================================
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W_NS = "{%s}" % W
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

RPR_ORDER = [
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
    "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid",
    "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz",
    "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign",
    "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath",
]

CSL_SCHEMA = "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"
MENDELEY_URI = "http://www.mendeley.com/documents/?uuid={uuid}"

BASE_PARTS = ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml")
TEXT_PART_RE = re.compile(r"^word/(header|footer)\d*\.xml$")

HERE = Path(__file__).resolve().parent
CACHE_PATH = HERE / "crossref_cache.json"


def rpr_insert(rpr, child):
    tag = ET.QName(child).localname
    rank = RPR_ORDER.index(tag) if tag in RPR_ORDER else len(RPR_ORDER)
    for i, existing in enumerate(rpr):
        e = ET.QName(existing).localname
        if (RPR_ORDER.index(e) if e in RPR_ORDER else len(RPR_ORDER)) > rank:
            rpr.insert(i, child)
            return
    rpr.append(child)


# ==========================================================================
# Utilitas Umum & Normalisasi
# ==========================================================================
def norm_title(s) -> str:
    return re.sub(r"[^a-z0-9]", "", str(s or "").lower())


def norm_doi(s) -> str:
    d = str(s or "").strip().lower()
    for pre in ("https://doi.org/", "http://doi.org/", "https://dx.doi.org/",
                "http://dx.doi.org/", "https://www.doi.org/", "doi:"):
        d = d.removeprefix(pre)
    d = d.split()[0] if d.split() else ""
    return d.rstrip(".,;)")


def parse_year(v):
    m = re.search(r"(?:19|20)\d{2}", str(v or ""))
    return m.group(0) if m else None


def clean_id(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    try:
        f = float(s)
        if f == int(f):
            return str(int(f))
    except (ValueError, OverflowError):
        pass
    return s


def norm_hdr(s) -> str:
    return re.sub(r"[^a-z0-9]", "", str(s or "").lower())


def csl_name(family, given="") -> dict:
    return {"dropping-particle": "", "family": family or "", "given": given or "",
            "non-dropping-particle": "", "parse-names": False, "suffix": ""}


def parse_authors_full(text) -> list:
    """'Liang, Houhao (123); Jamaluddin, Azrin (456)' -> CSL names."""
    out = []
    for part in [p.strip() for p in str(text or "").split(";") if p.strip()]:
        part = re.sub(r"\s*\([^)]*\)\s*$", "", part).strip()
        if not part:
            continue
        if "," in part:
            fam, giv = part.split(",", 1)
        else:
            bits = part.split()
            fam, giv = (bits[0], " ".join(bits[1:])) if len(bits) > 1 else (part, "")
        out.append(csl_name(fam.strip(), giv.strip()))
    return out


def parse_authors_wos(text) -> list:
    """'Liang H.; Jamaluddin A.' -> CSL names."""
    out = []
    for part in [p.strip() for p in str(text or "").split(";") if p.strip()]:
        part = re.sub(r"\s*\([^)]*\)\s*$", "", part).strip()
        if not part:
            continue
        if "," in part:
            fam, giv = part.split(",", 1)
        else:
            bits = part.split()
            if len(bits) >= 2 and re.fullmatch(r"[A-Z][.\-]?[A-Z]?\.?", bits[-1]):
                fam, giv = " ".join(bits[:-1]), bits[-1]
            else:
                fam, giv = bits[0], " ".join(bits[1:])
        out.append(csl_name(fam.strip(), giv.strip()))
    return out


def authors_from_label(label) -> list:
    """'John et al. (2024)' -> [CSL name 'John']."""
    s = str(label or "").strip()
    if not s:
        return []
    s = re.split(r"\s+et\s+al|\s+&\s+|;|,|\(|\band\b|\bdan\b", s, maxsplit=1, flags=re.I)[0]
    s = s.strip().rstrip(".")
    return [csl_name(s, "")] if s else []


def author_year_label(item: dict) -> str:
    authors = item.get("author") or []
    year = ""
    try:
        year = str(item["issued"]["date-parts"][0][0])
    except Exception:
        pass
    fams = [a.get("family") or a.get("literal") or "" for a in authors]
    fams = [f for f in fams if f]
    if not fams:
        return f"Anon., {year}" if year else "Anon."
    if len(fams) == 1:
        who = fams[0]
    elif len(fams) == 2:
        who = f"{fams[0]} & {fams[1]}"
    else:
        who = f"{fams[0]} et al."
    return f"{who}, {year}" if year else who


# ==========================================================================
# Ekspansi Rentang (Range Expansion) misal [1-4] -> 1, 2, 3, 4
# ==========================================================================
RANGE_RE = re.compile(r"^([A-Za-z_-]*)(\d+)\s*[-–—]\s*([A-Za-z_-]*)(\d+)$")


def expand_range_token(token_str: str) -> list[str]:
    s = token_str.strip()
    m = RANGE_RE.match(s)
    if m:
        p1, n1, p2, n2 = m.groups()
        if not p2 or p1 == p2:
            try:
                start, end = int(n1), int(n2)
                if 0 < (end - start) <= 60:
                    width = len(n1)
                    return [f"{p1}{i:0{width}d}" for i in range(start, end + 1)]
            except ValueError:
                pass
    return [s] if s else []


def extract_bracket_tokens(inner_text: str) -> list[str]:
    """Mengurai isi kurung 'NR001; NR002-NR004' menjadi daftar kode individual."""
    raw_parts = re.split(r"[,;&]|\s+and\s+|\s+dan\s+", inner_text)
    tokens = []
    for part in raw_parts:
        part = part.strip()
        if not part:
            continue
        tokens.extend(expand_range_token(part))
    return tokens


# ==========================================================================
# Deteksi File Input di Folder
# ==========================================================================
def _docx_candidates(folder: Path, exclude: set) -> list[Path]:
    out = []
    for p in folder.glob("*.docx"):
        if p.name.startswith("~$") or p.name in exclude:
            continue
        if "mendeley" in p.stem.lower():
            continue
        out.append(p)
    return out


def _xlsx_candidates(folder: Path, exclude: set) -> list[Path]:
    out = []
    for p in folder.glob("*.xlsx"):
        if p.name.startswith("~$") or p.name in exclude:
            continue
        stem = p.stem.lower()
        if "laporan_sitasi" in stem or "citation_report" in stem:
            continue
        out.append(p)
    return out


def discover_inputs(folder: Path, docx_arg, workbook_arg, exclude=()):
    exclude = set(exclude)
    notes = []
    if docx_arg:
        docx = Path(docx_arg)
    else:
        cands = _docx_candidates(folder, exclude)
        if not cands:
            raise SystemExit(f"Tidak ada file .docx naskah di folder {folder}")
        cands.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        docx = cands[0]
        if len(cands) > 1:
            notes.append("Lebih dari satu .docx, otomatis memakai yang terbaru: " + docx.name)

    if workbook_arg:
        workbook = Path(workbook_arg)
    else:
        cands = _xlsx_candidates(folder, exclude)
        if not cands:
            raise SystemExit(f"Tidak ada file .xlsx referensi di folder {folder}")
        cands.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        workbook = cands[0]
        if len(cands) > 1:
            notes.append("Lebih dari satu .xlsx, otomatis memakai yang terbaru: " + workbook.name)

    return docx, workbook, notes


# ==========================================================================
# Ekstraksi Teks Word & Koleksi Token Potensial
# ==========================================================================
def story_parts(names) -> list[str]:
    parts = [p for p in BASE_PARTS if p in names]
    parts += [n for n in names if TEXT_PART_RE.match(n)]
    return parts


def read_doc_texts(docx: Path):
    texts, parts = [], []
    with zipfile.ZipFile(docx) as z:
        names = set(z.namelist())
        parts = story_parts(names)
        for part in parts:
            try:
                root = ET.fromstring(z.read(part), ET.XMLParser(huge_tree=True))
            except ET.XMLSyntaxError:
                continue
            for p in root.iter(W_NS + "p"):
                texts.append("".join(t.text or "" for t in p.iter(W_NS + "t")))
    return texts, parts


BRACKET_ALL_RE = re.compile(r"\[[^\[\]]{1,120}\]|\{[^{}]{1,120}\}|\([^()]{1,120}\)")
CODE_TOKEN_RE = re.compile(r"(?<![A-Za-z0-9])[A-Za-z]{1,12}[-_ ]?\d{1,6}(?![A-Za-z0-9])")


def collect_doc_tokens(texts) -> collections.Counter:
    """Kumpulkan semua token yang berpotensi menjadi kode sitasi."""
    toks = collections.Counter()
    for t in texts:
        for m in CODE_TOKEN_RE.finditer(t):
            toks[clean_id(m.group(0))] += 1
        for m in BRACKET_ALL_RE.finditer(t):
            inner = m.group(0)[1:-1]
            for tok in extract_bracket_tokens(inner):
                c = clean_id(tok)
                if c:
                    toks[c] += 1
    return toks


def diagnose_brackets(texts, limit=12) -> list[str]:
    c = collections.Counter()
    for t in texts:
        for m in BRACKET_ALL_RE.finditer(t):
            c[m.group(0)] += 1
    return [k for k, _ in c.most_common(limit)]


# ==========================================================================
# Deteksi Pola Placeholder
# ==========================================================================
class Pattern:
    __slots__ = ("group_re", "token_re", "bare", "label", "tok", "prefix")

    def __init__(self, group_re, token_re, bare, label, tok="", prefix=""):
        self.group_re = group_re
        self.token_re = token_re
        self.bare = bare
        self.label = label
        self.tok = tok
        self.prefix = prefix


DIGIT_MARK = "\x00"
LETTER_MARK = "\x01"


def _shape_of(val: str) -> str:
    s = re.sub(r"\d+", lambda m: DIGIT_MARK * len(m.group()), str(val))
    s = re.sub(r"[A-Za-z]+", lambda m: LETTER_MARK * len(m.group()), s)
    return s


def _shape_to_regex(shape: str) -> str:
    out = []
    for ch in shape:
        if ch == DIGIT_MARK:
            out.append(r"\d")
        elif ch == LETTER_MARK:
            out.append(r"[A-Za-z]")
        else:
            out.append(re.escape(ch))
    return "".join(out)


def _shape_label(shape: str) -> str:
    return shape.replace(DIGIT_MARK, "9").replace(LETTER_MARK, "A")


def _make_pattern(tok: str, bare: bool, label: str, prefix="") -> Pattern:
    sep = r"\s*[,;&]\s*|\s+and\s+|\s+dan\s+"
    # Range part: tok or tok-tok
    range_part = rf"(?:{tok}(?:\s*[-–—]\s*{tok})?)"
    group = re.compile(
        r"(?:"
        + r"\(" + range_part + r"(?:" + sep + range_part + r")*" + r"\)"
        + r"|" + r"\[" + range_part + r"(?:" + sep + range_part + r")*" + r"\]"
        + r"|" + r"\{" + range_part + r"(?:" + sep + range_part + r")*" + r"\}"
        + r")"
    )
    return Pattern(group, re.compile(tok), bare, label, tok, prefix)


def find_common_prefix(strings: list[str]) -> str:
    if not strings:
        return ""
    m = re.match(r"^([A-Za-z_-]+)", strings[0])
    if not m:
        return ""
    pre = m.group(1)
    for s in strings[1:]:
        if not s.startswith(pre):
            return ""
    return pre


def detect_placeholder(texts, ids, forced_regex=None, allow_bare=None):
    """Mendeteksi pola placeholder di Word yang paling sesuai dengan ID di Excel."""
    id_set = {clean_id(i) for i in ids if clean_id(i)}
    shapes = collections.Counter(_shape_of(i) for i in id_set)
    candidates = []

    if forced_regex:
        candidates.append((forced_regex, "manual regex", ""))

    common_pre = find_common_prefix(list(id_set))
    if common_pre and len(common_pre) >= 1:
        candidates.append((re.escape(common_pre) + r"[-_ ]?\d{1,6}", f"prefix '{common_pre}' + angka", common_pre))

    for shape, _n in shapes.most_common(3):
        if len(shape) <= 40:
            candidates.append((_shape_to_regex(shape), f"pola ID {_shape_label(shape)}", ""))

    candidates += [
        (r"[A-Za-z]{1,10}[-_ ]?\d{1,6}", "generik huruf+angka", ""),
        (r"\d{1,5}", "generik angka", ""),
    ]

    best = None
    for tok, label, pre in candidates:
        try:
            re.compile(tok)
        except re.error:
            continue
        for bare in ([allow_bare] if allow_bare is not None else [False, True]):
            if bare and re.fullmatch(r"\\d(\{1,5\})?|\\d\{1,5\}", tok):
                continue
            pat = _make_pattern(tok, bare, label, pre)
            found, in_ids, total = set(), set(), 0
            for t in texts:
                for m in pat.group_re.finditer(t):
                    inner = m.group(0)[1:-1]
                    for c in extract_bracket_tokens(inner):
                        total += 1
                        found.add(c)
                        if c in id_set:
                            in_ids.add(c)
                if bare:
                    covered = [(m.start(), m.end()) for m in pat.group_re.finditer(t)]
                    for m in pat.token_re.finditer(t):
                        if any(a <= m.start() < b for a, b in covered):
                            continue
                        total += 1
                        tok_val = m.group(0)
                        found.add(tok_val)
                        if tok_val in id_set:
                            in_ids.add(tok_val)
            if not in_ids:
                continue
            precision = len(in_ids) / len(found) if found else 0
            # Bobot: utamakan jumlah ID cocok dan presisi
            score = (len(in_ids), precision, total)
            if best is None or score > best[0]:
                best = (score, pat, sorted(in_ids))

    if best is None:
        return None, []
    _score, pat, matched = best
    if allow_bare is None:
        pat.bare = not re.fullmatch(r"\\d|\\d\{1,5\}", pat.tok or "")
    return pat, matched


# ==========================================================================
# Excel: Pemilihan Sheet, Header, dan Pemetaan Kolom
# ==========================================================================
ID_HEADER_KEYS = [
    "studyid", "study_id", "study", "studyno", "studynumber", "id", "idstudi",
    "id_studi", "kode", "code", "refid", "ref_id", "ref", "reference",
    "no", "nomor", "key", "citationkey", "citation_key", "nr", "label", "entry", "record"
]


def pick_table(xl: pd.ExcelFile, doc_tokens: collections.Counter):
    best = None
    fallback = None

    for sheet in xl.sheet_names:
        try:
            raw = xl.parse(sheet, header=None, dtype=str)
        except Exception:
            continue
        if raw.empty or raw.shape[1] == 0:
            continue
        max_h = min(12, len(raw))
        for h in range(max_h):
            body = raw.iloc[h + 1:]
            if body.empty:
                continue
            headers = [norm_hdr(x) for x in raw.iloc[h].tolist()]

            for ci in range(raw.shape[1]):
                vals = body.iloc[:, ci].dropna().astype(str).str.strip()
                vals = vals[(vals != "") & (vals.str.lower() != "nan")]
                if vals.empty:
                    continue

                cleaned_vals = {clean_id(v) for v in vals if clean_id(v)}
                uniq = cleaned_vals & set(doc_tokens.keys())
                hdr_name = headers[ci] if ci < len(headers) else ""
                is_id_named = any(k == hdr_name or hdr_name.startswith(k) for k in ID_HEADER_KEYS)

                uniqueness = len(cleaned_vals) / len(vals)
                if uniq:
                    prec = len(uniq) / len(cleaned_vals)
                    score = len(uniq) * 5 + prec * 25 + (10 if is_id_named else 0) + min(len(vals), 200) / 200
                    if best is None or score > best["score"]:
                        best = {"score": score, "sheet": sheet, "header": h, "col": ci,
                                "overlap": len(uniq), "raw": raw}
                elif is_id_named and uniqueness > 0.6:
                    score = 20 * uniqueness
                    if fallback is None or score > fallback["score"]:
                        fallback = {"score": score, "sheet": sheet, "header": h, "col": ci,
                                    "overlap": 0, "raw": raw}

    return best or fallback


def build_frame(xl, info):
    raw = info["raw"]
    h = info["header"]
    header = [str(x).strip() if str(x).strip() not in ("", "nan") else f"col{i}"
              for i, x in enumerate(raw.iloc[h].tolist())]
    seen, cols = {}, []
    for c in header:
        if c in seen:
            seen[c] += 1
            cols.append(f"{c}.{seen[c]}")
        else:
            seen[c] = 0
            cols.append(c)
    df = raw.iloc[h + 1:].copy()
    df.columns = cols
    df = df.dropna(how="all")
    return df


def _header_score(name, keys):
    n = norm_hdr(name)
    best = 0
    for k in keys:
        if n == k:
            return 3
        if n.startswith(k) or n.endswith(k):
            best = max(best, 2)
        elif k in n:
            best = max(best, 1)
    return best


def map_columns(df, id_col, doc_tokens):
    cols = [c for c in df.columns if c != id_col]
    used = {id_col}

    def pick(keys, min_score=1):
        cand = [(_header_score(c, keys), i, c) for i, c in enumerate(cols)]
        cand = [x for x in cand if x[0] >= min_score and x[2] not in used]
        if not cand:
            return None
        cand.sort(key=lambda x: (-x[0], x[1]))
        used.add(cand[0][2])
        return cand[0][2]

    mapping = {"id": id_col}
    mapping["doi"] = pick(["doi", "identifier", "digitalobjectidentifier", "linkdoi"])
    mapping["year"] = pick(["year", "tahun", "date", "publicationyear", "terbit", "pubyear"])
    mapping["journal"] = pick(["journal", "venue", "source", "container", "jurnal",
                                "publisher", "proceedings", "sourcetitle", "publication"])
    mapping["authors"] = pick(["author", "authors", "penulis", "creator", "authorfullnames"])
    mapping["title"] = pick(["title", "judul", "paper", "article", "documenttitle"])
    mapping["volume"] = pick(["volume", "vol"])
    mapping["issue"] = pick(["issue", "nomor", "no", "iss"])
    mapping["pages"] = pick(["page", "pages", "halaman", "pagination", "pp"])

    for c in df.columns:
        if c == id_col:
            continue
        n = norm_hdr(c)
        if "author" in n or "penulis" in n:
            mapping["authors_label"] = c
            break

    def col_values(c):
        return df[c].dropna().astype(str).str.strip()

    if mapping.get("doi") is None:
        for c in df.columns:
            if c in used:
                continue
            v = col_values(c)
            if v.empty:
                continue
            hit = v.str.match(r"^10\.\d{4,9}/\S+$").mean()
            if hit > 0.5:
                mapping["doi"] = c
                used.add(c)
                break

    if mapping.get("year") is None:
        for c in df.columns:
            if c in used:
                continue
            v = col_values(c)
            if v.empty:
                continue
            hit = v.str.fullmatch(r"(?:19|20)\d{2}").mean()
            if hit > 0.7:
                mapping["year"] = c
                used.add(c)
                break

    if mapping.get("title") is None:
        cand = []
        for c in df.columns:
            if c in used:
                continue
            v = col_values(c)
            if v.empty:
                continue
            avg = v.str.len().mean()
            if avg > 25:
                cand.append((avg, c))
        if cand:
            cand.sort(reverse=True)
            mapping["title"] = cand[0][1]
    return mapping


# ==========================================================================
# Crossref API & Cache
# ==========================================================================
def load_cache() -> dict:
    if CACHE_PATH.exists():
        try:
            return json.loads(CACHE_PATH.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_cache(cache):
    try:
        CACHE_PATH.write_text(json.dumps(cache, ensure_ascii=False, indent=1),
                              encoding="utf-8")
    except Exception:
        pass


def crossref_work(doi, cache, timeout=12):
    doi = norm_doi(doi)
    if not doi:
        return None
    if doi in cache:
        return cache[doi]
    try:
        req = urllib.request.Request(
            "https://api.crossref.org/works/" + urllib.parse.quote(doi),
            headers={"User-Agent": "mendeley-automate/2.5 (mailto:researcher@example.com)"},
        )
        with urllib.request.urlopen(req, timeout=timeout) as r:
            m = json.loads(r.read().decode("utf-8"))["message"]
    except Exception:
        cache[doi] = None
        return None

    authors = []
    for a in m.get("author") or []:
        fam = a.get("family") or a.get("name") or ""
        giv = a.get("given") or ""
        if fam:
            authors.append(csl_name(fam, giv))
    titles = m.get("title") or []
    conts = m.get("container-title") or []
    issued = (m.get("issued") or {}).get("date-parts") or [[None]]
    data = {
        "title": titles[0] if titles else "",
        "authors": authors,
        "journal": conts[0] if conts else "",
        "year": str(issued[0][0]) if issued and issued[0] and issued[0][0] else None,
        "volume": m.get("volume"),
        "issue": m.get("issue"),
        "page": m.get("page"),
    }
    cache[doi] = data
    time.sleep(0.04)
    return data


# ==========================================================================
# Library Mendeley Reference Manager (SQLite Lokal)
# ==========================================================================
def find_mrm_db():
    base = Path(os.environ.get("APPDATA", "")) / "Mendeley Reference Manager" / "mrm" / "databases"
    if not base.is_dir():
        return None
    dbs = [p for p in base.glob("*.db") if p.stat().st_size > 4096]
    return max(dbs, key=lambda p: p.stat().st_size) if dbs else None


def load_library(db_path):
    if db_path is None:
        return []
    tmp = Path(tempfile.gettempdir()) / "opencode" / "mrm_read.db"
    tmp.parent.mkdir(parents=True, exist_ok=True)
    for suf in ("", "-wal", "-shm"):
        src = Path(str(db_path) + suf)
        if src.exists():
            try:
                shutil.copy2(src, str(tmp) + suf)
            except OSError:
                pass
    try:
        con = sqlite3.connect(tmp)
        rows = con.execute("select c0,c1,c2,c3,c4,c8 from documents_fts_content").fetchall()
        con.close()
    except sqlite3.Error as exc:
        print(f"      ! gagal membaca database Mendeley: {exc}")
        return []
    out = []
    for uuid, title, authors, journal, year, doi in rows:
        if uuid:
            out.append({"uuid": uuid, "title": title or "", "authors": authors or "",
                        "journal": journal or "", "year": parse_year(year),
                        "doi": norm_doi(doi)})
    return out


class Library:
    def __init__(self, docs):
        self.docs = docs
        self.by_doi, self.by_title = {}, {}
        for d in docs:
            if d["doi"]:
                self.by_doi.setdefault(d["doi"], d)
            if d["title"]:
                self.by_title.setdefault(norm_title(d["title"]), d)

    def lookup(self, doi, title, year=None):
        doi = norm_doi(doi)
        if doi and doi in self.by_doi:
            return self.by_doi[doi], "DOI"
        nt = norm_title(title)
        if nt and nt in self.by_title:
            return self.by_title[nt], "Judul (Persis)"
        if nt:
            best, score = None, 0.0
            for key, d in self.by_title.items():
                if not key or abs(len(key) - len(nt)) > 40:
                    continue
                r = difflib.SequenceMatcher(None, nt, key).ratio()
                if r > score:
                    best, score = d, r
            if best is not None and score >= 0.90:
                return best, f"Judul (Fuzzy {int(score * 100)}%)"
        return None, ""


# ==========================================================================
# CSL & Payload Generator
# ==========================================================================
def build_csl_item(ref, item_id):
    item = {"author": ref.get("authors") or [],
            "container-title": ref.get("journal") or "",
            "id": item_id,
            "title": ref.get("title") or "",
            "type": ref.get("type") or "article-journal"}
    if ref.get("year"):
        item["issued"] = {"date-parts": [[str(ref["year"])]]}
    for k in ("volume", "issue", "page"):
        if ref.get(k):
            item[k] = str(ref[k])
    if ref.get("doi"):
        item["DOI"] = str(ref["doi"])
    return {k: v for k, v in item.items() if v not in ("", None, [])}


def build_field_json(items, display):
    payload = {
        "citationItems": [
            {"id": f"ITEM-{i + 1}",
             "itemData": dict(item, id=f"ITEM-{i + 1}"),
             "uris": [MENDELEY_URI.format(uuid=uuid)]}
            for i, (uuid, item) in enumerate(items)
        ],
        "mendeley": {"formattedCitation": display,
                     "plainTextFormattedCitation": display,
                     "previouslyFormattedCitation": display},
        "properties": {"noteIndex": 0},
        "schema": CSL_SCHEMA,
    }
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":"))


# ==========================================================================
# XML OpenXML Modification
# ==========================================================================
def nearest_par(node):
    n = node.getparent()
    while n is not None and n.tag != W_NS + "p":
        n = n.getparent()
    return n


def nearest_run(node):
    n = node.getparent()
    while n is not None and n.tag != W_NS + "r":
        n = n.getparent()
    return n


def own_text_nodes(p):
    return [t for t in p.iter(W_NS + "t") if nearest_par(t) is p]


def textmap(nodes):
    texts, spans, pos = [], [], 0
    for t in nodes:
        s = t.text or ""
        spans.append((t, pos, pos + len(s)))
        texts.append(s)
        pos += len(s)
    return "".join(texts), spans


def split_run_at(t, local):
    run = nearest_run(t)
    if run is None:
        return False
    content = [c for c in run if c.tag != W_NS + "rPr"]
    if len(content) != 1 or content[0] is not t:
        return False
    s = t.text or ""
    if not (0 < local < len(s)):
        return False
    head, tail = s[:local], s[local:]
    t.text = head
    if head and (head[0].isspace() or head[-1].isspace()):
        t.set(XML_SPACE, "preserve")
    new_run = copy.deepcopy(run)
    nt = new_run.find(W_NS + "t")
    nt.text = tail
    if tail and (tail[0].isspace() or tail[-1].isspace()):
        nt.set(XML_SPACE, "preserve")
    parent = run.getparent()
    parent.insert(parent.index(run) + 1, new_run)
    return True


def make_field_runs(rpr, field_code, display, lock=True):
    runs = []

    def run_with(child):
        r = ET.Element(W_NS + "r")
        if rpr is not None:
            r.append(copy.deepcopy(rpr))
        r.append(child)
        runs.append(r)

    begin = ET.Element(W_NS + "fldChar", {W_NS + "fldCharType": "begin"})
    if lock:
        begin.set(W_NS + "fldLock", "1")
    run_with(begin)

    instr = ET.Element(W_NS + "instrText", {XML_SPACE: "preserve"})
    instr.text = field_code
    run_with(instr)

    run_with(ET.Element(W_NS + "fldChar", {W_NS + "fldCharType": "separate"}))

    dr = ET.Element(W_NS + "r")
    drpr = copy.deepcopy(rpr) if rpr is not None else ET.Element(W_NS + "rPr")
    if drpr.find(W_NS + "noProof") is None:
        rpr_insert(drpr, ET.Element(W_NS + "noProof"))
    dr.append(drpr)
    t = ET.Element(W_NS + "t", {XML_SPACE: "preserve"})
    t.text = display
    dr.append(t)
    runs.append(dr)

    run_with(ET.Element(W_NS + "fldChar", {W_NS + "fldCharType": "end"}))
    return runs


def make_text_run(rpr, text):
    r = ET.Element(W_NS + "r")
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = ET.Element(W_NS + "t", {XML_SPACE: "preserve"})
    t.text = text
    r.append(t)
    return r


def apply_replacement(spans, start, end, field_runs):
    involved = [(t, s, e) for t, s, e in spans if max(start, s) < min(end, e)]
    if not involved:
        return False
    for t, s, e in involved:
        a, b = max(start, s), min(end, e)
        s_t = t.text or ""
        t.text = s_t[: a - s] + s_t[b - s:]
        if t.text and (t.text[0].isspace() or t.text[-1].isspace()):
            t.set(XML_SPACE, "preserve")
    anchor = None
    for t, s, e in spans:
        if e == start and e > s:
            anchor = nearest_run(t)
    if anchor is not None:
        parent = anchor.getparent()
        idx = parent.index(anchor) + 1
    else:
        first = nearest_run(involved[0][0])
        parent = first.getparent()
        idx = parent.index(first)
    for i, r in enumerate(field_runs):
        parent.insert(idx + i, r)
    return True


def patch_paragraph(p, resolver, style, counter, stats, pattern, known_ids, snippet_tracker):
    guard = 0
    while True:
        guard += 1
        if guard > 4000:
            return
        nodes = own_text_nodes(p)
        if not nodes:
            return
        text, spans = textmap(nodes)
        if not text:
            return

        groups = []
        for m in pattern.group_re.finditer(text):
            inner = m.group(0)[1:-1]
            codes = extract_bracket_tokens(inner)
            groups.append((m.start(), m.end(), codes, m.group(0)))

        covered = [(a, b) for a, b, _, _ in groups]
        matches = list(groups)
        if pattern.bare:
            for m in pattern.token_re.finditer(text):
                if any(a <= m.start() < b for a, b in covered):
                    continue
                code_val = clean_id(m.group(0))
                matches.append((m.start(), m.end(), [code_val], m.group(0)))

        if not matches:
            return
        matches.sort()

        # Rekam snippet kalimat untuk setiap kode yang ditemukan
        for start, end, codes, raw_str in matches:
            ctx_start = max(0, start - 50)
            ctx_end = min(len(text), end + 50)
            snip = text[ctx_start:ctx_end].replace("\r", " ").replace("\n", " ").strip()
            if ctx_start > 0:
                snip = "..." + snip
            if ctx_end < len(text):
                snip = snip + "..."
            for c in codes:
                tracker = snippet_tracker.setdefault(c, {"count": 0, "snippets": []})
                tracker["count"] += 1
                if snip not in tracker["snippets"] and len(tracker["snippets"]) < 3:
                    tracker["snippets"].append(snip)

        # Cari match yang setidaknya salah satu kodenya resolved
        pick = None
        for cand in matches:
            if any(resolver(c) is not None for c in cand[2]):
                pick = cand
                break

        if pick is None:
            # Tidak ada satupun yang resolved di grup ini
            for cand in matches:
                for c in cand[2]:
                    if c in known_ids:
                        stats["unresolved_mendeley"].add(c)
                    else:
                        stats["unresolved_excel"].add(c)
            stats["blocked_groups"] += len(matches)
            return

        start, end, codes, raw_str = pick

        spans = textmap(own_text_nodes(p))[1]
        node = next((t for t, s, e in spans if s <= start < e), None)
        if node is not None:
            s_ = next(s for t, s, e in spans if t is node)
            if start - s_ > 0 and split_run_at(node, start - s_):
                continue
        spans = textmap(own_text_nodes(p))[1]
        node2 = next((t for t, s, e in spans if s <= end - 1 < e), None)
        if node2 is not None:
            s2 = next(s for t, s, e in spans if t is node2)
            lo2 = end - s2
            if 0 < lo2 < len(node2.text or "") and split_run_at(node2, lo2):
                continue
        spans = textmap(own_text_nodes(p))[1]

        ok_codes = [c for c in codes if resolver(c) is not None]
        bad_mendeley = [c for c in codes if resolver(c) is None and c in known_ids]
        bad_excel = [c for c in codes if resolver(c) is None and c not in known_ids]

        items, labels = [], []
        for code in ok_codes:
            uuid, item, _m = resolver(code)
            items.append((uuid, item))
            if style == "numeric":
                labels.append((counter.setdefault(uuid, len(counter) + 1), str(counter[uuid])))
            else:
                labels.append((10 ** 9, author_year_label(item)))

        if style == "numeric":
            labels.sort(key=lambda x: x[0])
            display = "(" + ", ".join(lbl for _, lbl in labels) + ")"
        else:
            display = "(" + "; ".join(lbl for _, lbl in labels) + ")"

        field_code = "ADDIN CSL_CITATION " + build_field_json(items, display)
        anchor = next(t for t, s, e in spans if s <= start < e)
        rpr = anchor.getparent().find(W_NS + "rPr")
        runs = make_field_runs(rpr, field_code, display)

        # Sisa kode yang belum selesai tetap disisipkan sebagai teks agar tidak hilang
        unresolved_rest = bad_mendeley + bad_excel
        if unresolved_rest:
            runs.append(make_text_run(rpr, "(" + "; ".join(unresolved_rest) + ")"))
            for c in bad_mendeley:
                stats["unresolved_mendeley"].add(c)
            for c in bad_excel:
                stats["unresolved_excel"].add(c)
            stats["partial"] += 1

        if not apply_replacement(spans, start, end, runs):
            return
        stats["citations"] += 1
        stats["codes"].update(ok_codes)


def patch_xml_bytes(data, resolver, style, counter, stats, pattern, known_ids, snippet_tracker):
    root = ET.fromstring(data, ET.XMLParser(remove_blank_text=False, huge_tree=True))
    for p in list(root.iter(W_NS + "p")):
        patch_paragraph(p, resolver, style, counter, stats, pattern, known_ids, snippet_tracker)
    return ET.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def collect_leftovers(data, pattern):
    out = []
    try:
        root = ET.fromstring(data, ET.XMLParser(huge_tree=True))
    except ET.XMLSyntaxError:
        return out
    probe = pattern.token_re if pattern.bare else pattern.group_re
    for p in root.iter(W_NS + "p"):
        txt = "".join(t.text or "" for t in p.iter(W_NS + "t"))
        if probe.search(txt):
            out.append(re.sub(r"\s+", " ", txt).strip())
    return out


def rewrite_docx(src, dst, replaced):
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = replaced.get(item.filename)
            if data is None:
                data = zin.read(item.filename)
            zi = zipfile.ZipInfo(item.filename, date_time=item.date_time)
            zi.compress_type = item.compress_type
            zi.external_attr = item.external_attr
            zout.writestr(zi, data)


# ==========================================================================
# RIS Generator (Untuk referensi yang belum ada di Mendeley)
# ==========================================================================
def write_ris(path, refs):
    lines = []
    for r in refs:
        lines.append("TY  - JOUR")
        for a in r.get("authors") or []:
            lines.append(f"AU  - {a.get('family','')}, {a.get('given','')}".rstrip(", "))
        if r.get("title"):
            lines.append(f"TI  - {r['title']}")
        if r.get("journal"):
            lines.append(f"JO  - {r['journal']}")
        if r.get("year"):
            lines.append(f"PY  - {r['year']}")
        if r.get("volume"):
            lines.append(f"VL  - {r['volume']}")
        if r.get("issue"):
            lines.append(f"IS  - {r['issue']}")
        if r.get("page"):
            pg = str(r["page"]).replace("-", "\nEP  - ")
            lines.append(f"SP  - {pg}")
        if r.get("doi"):
            lines.append(f"DO  - {r['doi']}")
        lines.append("ER  - ")
        lines.append("")
    Path(path).write_text("\n".join(lines), encoding="utf-8")


# ==========================================================================
# Pembuat Laporan Excel (.xlsx) Lengkap & Elegan
# ==========================================================================
def generate_excel_report(path: Path, docx_path: Path, out_path: Path,
                          excel_path: Path, sheet_used: str, db_name: str,
                          summary_metrics: dict, all_word_citations: list[dict],
                          excel_unused: list[dict]):
    wb = openpyxl.Workbook()
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    c_navy = '1F4E78'
    c_blue_light = 'D9E1F2'
    c_green_bg = 'E2EFDA'
    c_green_fg = '375623'
    c_red_bg = 'FCE4D6'
    c_red_fg = 'C65911'
    c_yellow_bg = 'FFF2CC'
    c_yellow_fg = '833C0C'
    c_gray_bg = 'F2F2F2'
    c_gray_fg = '595959'

    font_title = Font(name='Calibri', size=16, bold=True, color=c_navy)
    font_section = Font(name='Calibri', size=12, bold=True, color=c_navy)
    font_hdr = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    fill_hdr = PatternFill(start_color=c_navy, end_color=c_navy, fill_type='solid')

    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')

    # -------------------------------------------------------------
    # SHEET 1: Ringkasan
    # -------------------------------------------------------------
    ws1 = wb.active
    ws1.title = "Ringkasan"
    ws1.views.sheetView[0].showGridLines = True

    ws1['A1'] = "LAPORAN OTOMATISASI SITASI MENDELEY"
    ws1['A1'].font = font_title
    ws1['A2'] = f"Waktu Eksekusi: {datetime.now().strftime('%d %B %Y, %H:%M:%S')}"
    ws1['A2'].font = Font(name='Calibri', size=10, italic=True, color='595959')

    ws1['A4'] = "Informasi Berkas"
    ws1['A4'].font = font_section

    info_rows = [
        ("Dokumen Word Asli", str(docx_path.name)),
        ("Dokumen Word Hasil", str(out_path.name)),
        ("Workbook Excel", f"{excel_path.name} (Sheet: '{sheet_used}')"),
        ("Database Mendeley", db_name),
    ]
    for r_idx, (k, v) in enumerate(info_rows, start=5):
        ws1.cell(row=r_idx, column=1, value=k).font = Font(bold=True)
        ws1.cell(row=r_idx, column=2, value=v)

    ws1['A10'] = "Ringkasan Eksekusi Sitasi"
    ws1['A10'].font = font_section

    headers_summary = ["Metrik", "Jumlah", "Persentase", "Status"]
    for c_idx, h in enumerate(headers_summary, start=1):
        cell = ws1.cell(row=11, column=c_idx, value=h)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.alignment = align_center

    tot_unique = summary_metrics.get("total_word_unique", 0)
    succ_cnt = summary_metrics.get("success_count", 0)
    fail_mend = summary_metrics.get("fail_mendeley_count", 0)
    fail_xl = summary_metrics.get("fail_excel_count", 0)
    pct_succ = (succ_cnt / tot_unique * 100) if tot_unique else 0

    metric_rows = [
        ("Total Sitasi Dikonversi di Dokumen", summary_metrics.get("citations_converted", 0), "-", "DIPROSES"),
        ("Total ID Unik Sitasi di Word", tot_unique, "100%", "TERDETEKSI"),
        ("Sitasi BERHASIL (Ada di Mendeley)", succ_cnt, f"{pct_succ:.1f}%", "BERHASIL"),
        ("Sitasi GAGAL (Tidak Ada di Mendeley)", fail_mend,
         f"{(fail_mend / tot_unique * 100) if tot_unique else 0:.1f}%", "PERIKSA MENDELEY"),
        ("Sitasi GAGAL (Tidak Terdaftar di Excel)", fail_xl,
         f"{(fail_xl / tot_unique * 100) if tot_unique else 0:.1f}%", "PERIKSA EXCEL"),
        ("Referensi di Excel yang Tidak Disitasi di Word", len(excel_unused), "-", "TIDAK DISITASI"),
    ]

    for r_idx, (m_label, m_val, m_pct, m_stat) in enumerate(metric_rows, start=12):
        c1 = ws1.cell(row=r_idx, column=1, value=m_label)
        c2 = ws1.cell(row=r_idx, column=2, value=m_val)
        c3 = ws1.cell(row=r_idx, column=3, value=m_pct)
        c4 = ws1.cell(row=r_idx, column=4, value=m_stat)

        c2.alignment = align_center
        c3.alignment = align_center
        c4.alignment = align_center

        for c in (c1, c2, c3, c4):
            c.border = thin_border

        if m_stat == "BERHASIL":
            c4.fill = PatternFill(start_color=c_green_bg, end_color=c_green_bg, fill_type='solid')
            c4.font = Font(bold=True, color=c_green_fg)
        elif "MENDELEY" in m_stat:
            c4.fill = PatternFill(start_color=c_red_bg, end_color=c_red_bg, fill_type='solid')
            c4.font = Font(bold=True, color=c_red_fg)
        elif "EXCEL" in m_stat:
            c4.fill = PatternFill(start_color=c_yellow_bg, end_color=c_yellow_bg, fill_type='solid')
            c4.font = Font(bold=True, color=c_yellow_fg)

    ws1.column_dimensions['A'].width = 45
    ws1.column_dimensions['B'].width = 20
    ws1.column_dimensions['C'].width = 15
    ws1.column_dimensions['D'].width = 22

    # -------------------------------------------------------------
    # SHEET 2: Detail Sitasi Word
    # -------------------------------------------------------------
    ws2 = wb.create_sheet(title="Detail Sitasi Word")
    ws2.views.sheetView[0].showGridLines = True
    ws2.freeze_panes = "A2"

    headers_detail = [
        "No", "ID Sitasi", "Status", "Alasan & Keterangan", "Kemunculan",
        "Metode Cocok Mendeley", "UUID Mendeley", "Judul Dokumen",
        "Penulis", "Tahun", "Jurnal", "DOI", "Cuplikan Paragraf di Word"
    ]
    for c_idx, h in enumerate(headers_detail, start=1):
        cell = ws2.cell(row=1, column=c_idx, value=h)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.alignment = align_center

    for r_idx, row in enumerate(all_word_citations, start=2):
        ws2.cell(row=r_idx, column=1, value=r_idx - 1).alignment = align_center
        c_id = ws2.cell(row=r_idx, column=2, value=row.get("id", ""))
        c_id.alignment = align_center
        c_id.font = Font(bold=True)

        c_st = ws2.cell(row=r_idx, column=3, value=row.get("status", ""))
        c_st.alignment = align_center

        status_str = row.get("status", "")
        if "BERHASIL" in status_str:
            c_st.fill = PatternFill(start_color=c_green_bg, end_color=c_green_bg, fill_type='solid')
            c_st.font = Font(bold=True, color=c_green_fg)
        elif "MENDELEY" in status_str:
            c_st.fill = PatternFill(start_color=c_red_bg, end_color=c_red_bg, fill_type='solid')
            c_st.font = Font(bold=True, color=c_red_fg)
        elif "EXCEL" in status_str:
            c_st.fill = PatternFill(start_color=c_yellow_bg, end_color=c_yellow_bg, fill_type='solid')
            c_st.font = Font(bold=True, color=c_yellow_fg)

        ws2.cell(row=r_idx, column=4, value=row.get("reason", ""))
        ws2.cell(row=r_idx, column=5, value=row.get("count", 1)).alignment = align_center
        ws2.cell(row=r_idx, column=6, value=row.get("match_method", "-")).alignment = align_center
        ws2.cell(row=r_idx, column=7, value=row.get("uuid", "-")).alignment = align_center
        ws2.cell(row=r_idx, column=8, value=row.get("title", ""))
        ws2.cell(row=r_idx, column=9, value=row.get("authors", ""))
        ws2.cell(row=r_idx, column=10, value=row.get("year", "")).alignment = align_center
        ws2.cell(row=r_idx, column=11, value=row.get("journal", ""))
        ws2.cell(row=r_idx, column=12, value=row.get("doi", ""))
        ws2.cell(row=r_idx, column=13, value=row.get("snippet", ""))

        for c_idx in range(1, 14):
            ws2.cell(row=r_idx, column=c_idx).border = thin_border

    # -------------------------------------------------------------
    # SHEET 3: Perlu Tindakan
    # -------------------------------------------------------------
    ws3 = wb.create_sheet(title="Perlu Tindakan")
    ws3.views.sheetView[0].showGridLines = True
    ws3.freeze_panes = "A2"

    headers_action = [
        "No", "ID Sitasi", "Kategori Masalah", "Alasan Kegagalan",
        "Langkah Penyelesaian / Rekomendasi", "Judul di Excel", "DOI", "Cuplikan di Word"
    ]
    for c_idx, h in enumerate(headers_action, start=1):
        cell = ws3.cell(row=1, column=c_idx, value=h)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.alignment = align_center

    failed_items = [r for r in all_word_citations if "GAGAL" in r.get("status", "")]
    if not failed_items:
        ws3.cell(row=2, column=2, value="Semua sitasi berhasil dikonversi! Tidak ada tindakan perbaikan yang diperlukan.")
        ws3.cell(row=2, column=2).font = Font(bold=True, color=c_green_fg)
    else:
        for r_idx, row in enumerate(failed_items, start=2):
            ws3.cell(row=r_idx, column=1, value=r_idx - 1).alignment = align_center
            c_id = ws3.cell(row=r_idx, column=2, value=row.get("id", ""))
            c_id.alignment = align_center
            c_id.font = Font(bold=True)

            status_str = row.get("status", "")
            c_cat = ws3.cell(row=r_idx, column=3, value=status_str)
            c_cat.alignment = align_center
            if "MENDELEY" in status_str:
                c_cat.fill = PatternFill(start_color=c_red_bg, end_color=c_red_bg, fill_type='solid')
                c_cat.font = Font(bold=True, color=c_red_fg)
                solusi = "Impor file referensi_belum_ada.ris atau file PDF artikel ke Mendeley Reference Manager, lalu jalankan Run.bat lagi."
            else:
                c_cat.fill = PatternFill(start_color=c_yellow_bg, end_color=c_yellow_bg, fill_type='solid')
                c_cat.font = Font(bold=True, color=c_yellow_fg)
                solusi = "Buka file Excel referensi, tambahkan baris untuk ID ini dan lengkapi Judul/DOI/Penulisnya."

            ws3.cell(row=r_idx, column=4, value=row.get("reason", ""))
            ws3.cell(row=r_idx, column=5, value=solusi)
            ws3.cell(row=r_idx, column=6, value=row.get("title", ""))
            ws3.cell(row=r_idx, column=7, value=row.get("doi", ""))
            ws3.cell(row=r_idx, column=8, value=row.get("snippet", ""))

            for c_idx in range(1, 9):
                ws3.cell(row=r_idx, column=c_idx).border = thin_border

    # -------------------------------------------------------------
    # SHEET 4: Excel Belum Digunakan
    # -------------------------------------------------------------
    ws4 = wb.create_sheet(title="Excel Belum Digunakan")
    ws4.views.sheetView[0].showGridLines = True
    ws4.freeze_panes = "A2"

    headers_unused = ["No", "ID di Excel", "Judul Artikel", "Penulis", "Tahun", "Jurnal", "DOI", "Keterangan"]
    for c_idx, h in enumerate(headers_unused, start=1):
        cell = ws4.cell(row=1, column=c_idx, value=h)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.alignment = align_center

    for r_idx, row in enumerate(excel_unused, start=2):
        ws4.cell(row=r_idx, column=1, value=r_idx - 1).alignment = align_center
        c_id = ws4.cell(row=r_idx, column=2, value=row.get("id", ""))
        c_id.alignment = align_center
        c_id.font = Font(bold=True)
        ws4.cell(row=r_idx, column=3, value=row.get("title", ""))
        ws4.cell(row=r_idx, column=4, value=row.get("authors", ""))
        ws4.cell(row=r_idx, column=5, value=row.get("year", "")).alignment = align_center
        ws4.cell(row=r_idx, column=6, value=row.get("journal", ""))
        ws4.cell(row=r_idx, column=7, value=row.get("doi", ""))
        ws4.cell(row=r_idx, column=8, value="Terdaftar di Excel tetapi tidak ditemukan sitasinya di naskah Word.")

        for c_idx in range(1, 9):
            ws4.cell(row=r_idx, column=c_idx).border = thin_border

    # Auto-fit column widths
    for sheet in (ws2, ws3, ws4):
        for col in sheet.columns:
            max_len = 0
            for cell in col:
                val = str(cell.value or '')
                if len(val) > max_len:
                    max_len = len(val)
            col_letter = get_column_letter(col[0].column)
            sheet.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 60)

    wb.save(path)


# ==========================================================================
# Alur Utama (Main Pipeline)
# ==========================================================================
def load_config(path: Path) -> dict:
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except Exception as exc:
            print(f"! config.json tidak valid: {exc}")
    return {}


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Sitasi placeholder Word -> field Mendeley Cite & Laporan Excel.")
    ap.add_argument("--config", default=str(HERE / "config.json"))
    ap.add_argument("--folder", help="folder berisi .docx dan .xlsx")
    ap.add_argument("--docx")
    ap.add_argument("--workbook")
    ap.add_argument("--out")
    ap.add_argument("--style", choices=["author-year", "numeric"])
    ap.add_argument("--id-regex", help="paksa pola ID, mis. 'NR\\d{3}' atau 'S\\d+'")
    ap.add_argument("--id-column", help="paksa nama kolom ID di Excel")
    ap.add_argument("--sheet", help="paksa nama sheet Excel")
    ap.add_argument("--no-enrich", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args(argv)

    cfg = load_config(Path(args.config))
    folder = Path(args.folder or cfg.get("folder") or HERE)
    style = args.style or cfg.get("style", "author-year")
    enrich = (not args.no_enrich) and cfg.get("enrich", True)

    print("=" * 72)
    print("  WORD  ->  MENDELEY CITATION AUTOMATION (UNIVERSAL EDITION)")
    print("=" * 72)
    print(f"  Folder Kerja : {folder}")

    docx, workbook, notes = discover_inputs(
        folder, args.docx or cfg.get("docx"), args.workbook or cfg.get("workbook"))
    for n in notes:
        print(f"  Catatan      : {n}")
    print(f"  Naskah Word  : {docx.name}")
    print(f"  Data Excel   : {workbook.name}")

    # ---------- 1. Teks & Token Word ----------
    print("\n[1/6] Membaca naskah dokumen Word ...")
    texts, parts = read_doc_texts(docx)
    print(f"      {len(texts)} paragraf, {len(parts)} bagian dokumen")
    doc_tokens = collect_doc_tokens(texts)
    print(f"      {sum(doc_tokens.values())} token kandidat sitasi terdeteksi")

    # ---------- 2. Excel ----------
    print("\n[2/6] Menganalisis file Excel & mencocokkan kolom ...")
    xl = pd.ExcelFile(workbook)
    info = pick_table(xl, doc_tokens)
    if info is None:
        raise SystemExit(
            "Tidak menemukan kolom ID di Excel yang cocok dengan isi dokumen Word.\n"
            "Coba jalankan dengan: --id-column \"NamaKolomID\" atau --id-regex \"POLA\"")
    df = build_frame(xl, info)
    id_col = df.columns[info["col"]]
    sheet_used = info['sheet']
    print(f"      Sheet terpilih: '{sheet_used}' (baris header: {info['header'] + 1})")
    print(f"      Kolom ID      : '{id_col}' ({info['overlap']} nilai langsung cocok dengan dokumen)")

    if args.id_column and args.id_column in df.columns:
        id_col = args.id_column
        print(f"      -> Kolom ID dipaksa: '{id_col}'")

    raw_ids = [clean_id(v) for v in df[id_col].tolist()]
    ids = [i for i in raw_ids if i]
    id_set = set(ids)
    print(f"      {len(id_set)} ID unik terdaftar di file Excel")

    # ---------- 3. Pola Placeholder ----------
    print("\n[3/6] Mendeteksi pola placeholder sitasi di Word ...")
    pattern, used_ids = detect_placeholder(
        texts, ids, forced_regex=args.id_regex,
        allow_bare=None if args.id_regex is None else True)
    if pattern is None:
        contoh = diagnose_brackets(texts)
        raise SystemExit(
            "Tidak menemukan pola placeholder yang cocok di dalam naskah Word.\n"
            f"Contoh tanda kurung di naskah: {contoh}\n"
            "Coba jalankan dengan: --id-regex \"POLA\"")
    print(f"      Pola Terpilih: {pattern.label}")
    print(f"      Format Regex : {pattern.group_re.pattern[:75]}")
    print(f"      {len(used_ids)} ID Excel terdeteksi disitasi di Word")

    # ---------- 4. Metadata Referensi ----------
    print("\n[4/6] Menyusun metadata referensi ...")
    colmap = map_columns(df, id_col, doc_tokens)
    print("      Kolom terpetakan: " + ", ".join(f"{k}='{v}'" for k, v in colmap.items() if v))

    cache = load_cache()
    refs = {}
    for _, row in df.iterrows():
        sid = clean_id(row.get(id_col))
        if not sid:
            continue

        def g(key):
            c = colmap.get(key)
            if not c:
                return ""
            v = row.get(c)
            if v is None or (isinstance(v, float) and pd.isna(v)):
                return ""
            s = str(v).strip()
            return "" if s.lower() == "nan" else s

        authors_list = (parse_authors_full(g("authors"))
                        or parse_authors_wos(g("authors"))
                        or authors_from_label(g("authors_label")))
        ref = {
            "id": sid,
            "title": g("title"),
            "doi": g("doi"),
            "journal": g("journal"),
            "year": parse_year(g("year")) or parse_year(g("authors_label")),
            "authors": authors_list,
            "volume": g("volume"),
            "issue": g("issue"),
            "page": g("pages"),
            "type": "article-journal",
        }
        refs[sid] = ref

    if enrich:
        n_doi = sum(1 for r in refs.values() if norm_doi(r["doi"]))
        done = 0
        for r in refs.values():
            if not norm_doi(r["doi"]):
                continue
            meta = crossref_work(r["doi"], cache)
            if meta:
                for k in ("volume", "issue", "page"):
                    if not r.get(k) and meta.get(k):
                        r[k] = meta[k]
                if not r["title"] and meta.get("title"):
                    r["title"] = meta["title"]
                if not r["authors"] and meta.get("authors"):
                    r["authors"] = meta["authors"]
                if not r["journal"] and meta.get("journal"):
                    r["journal"] = meta["journal"]
                if not r["year"] and meta.get("year"):
                    r["year"] = meta["year"]
            done += 1
            if done % 10 == 0 or done == n_doi:
                print(f"      Crossref {done}/{n_doi}", end="\r", flush=True)
        save_cache(cache)
        print(" " * 35)
    print(f"      {len(refs)} referensi siap dipadankan")

    # ---------- 5. Pencocokan Library Mendeley ----------
    print("\n[5/6] Mencocokkan dengan library Mendeley Reference Manager ...")
    db = find_mrm_db()
    docs = load_library(db)
    lib = Library(docs)
    db_display = db.name if db else "(database tidak ditemukan)"
    print(f"      Database SQLite: {db_display}")
    print(f"      Isi Library    : {len(docs)} dokumen | {len(lib.by_doi)} ber-DOI | {len(lib.by_title)} ber-Judul")

    resolved = {}
    for sid, ref in refs.items():
        d, method = lib.lookup(ref["doi"], ref["title"], ref["year"])
        if d is None:
            continue
        item = build_csl_item(ref, "ITEM-1")
        if not item.get("author") and d["authors"]:
            item["author"] = [csl_name(str(d["authors"]).split(";")[0].strip(), "")]
        if not item.get("container-title") and d["journal"]:
            item["container-title"] = d["journal"]
        resolved[sid] = (d["uuid"], item, method)

    print(f"      Berhasil dicocokkan ke Mendeley: {len(resolved)} dari {len(refs)} referensi")

    def resolver(code):
        return resolved.get(code)

    if args.dry_run:
        print("\n(--dry-run aktif: simulasi selesai, tidak menulis berkas)")
        return 0

    # ---------- 6. Penulisan Dokumen Word & Patching XML ----------
    print("\n[6/6] Memodifikasi XML Word & menyisipkan field Mendeley ...")
    stats = {
        "citations": 0, "partial": 0, "blocked_groups": 0,
        "unresolved_mendeley": set(),
        "unresolved_excel": set(),
        "codes": set()
    }
    counter = {}
    known_ids = set(refs.keys())
    snippet_tracker = {}
    replaced, final_parts = {}, {}

    with zipfile.ZipFile(docx) as z:
        for part in parts:
            data = z.read(part)
            new = patch_xml_bytes(data, resolver, style, counter, stats, pattern, known_ids, snippet_tracker)
            final_parts[part] = new
            if new != data:
                replaced[part] = new

    # Penamaan output Word: default "manuscript - Mendeley.docx"
    default_out_name = "manuscript - Mendeley.docx"
    chosen_out_name = cfg.get("output_name") or default_out_name
    out = Path(args.out) if args.out else docx.with_name(chosen_out_name)
    if out.resolve() == docx.resolve():
        out = docx.with_name(docx.stem + " - Mendeley.docx")
    rewrite_docx(docx, out, replaced)

    # ---------- 7. Penyusunan Data Laporan ----------
    all_word_ids = set(stats["codes"]) | set(stats["unresolved_mendeley"]) | set(stats["unresolved_excel"])
    all_word_citations = []

    for sid in sorted(all_word_ids):
        tracker = snippet_tracker.get(sid, {"count": 1, "snippets": [""]})
        count = tracker["count"]
        snippet = " | ".join(tracker["snippets"]) if tracker["snippets"] else ""

        if sid in stats["codes"]:
            uuid, item, method = resolved[sid]
            r = refs.get(sid, {})
            auth_str = ", ".join(a.get("family", "") for a in r.get("authors", []))
            all_word_citations.append({
                "id": sid,
                "status": "BERHASIL",
                "reason": f"Berhasil dikonversi menjadi field Mendeley Cite (Cocok via: {method})",
                "count": count,
                "match_method": method,
                "uuid": uuid,
                "title": r.get("title", ""),
                "authors": auth_str,
                "year": r.get("year", ""),
                "journal": r.get("journal", ""),
                "doi": r.get("doi", ""),
                "snippet": snippet,
            })
        elif sid in stats["unresolved_mendeley"]:
            r = refs.get(sid, {})
            auth_str = ", ".join(a.get("family", "") for a in r.get("authors", []))
            all_word_citations.append({
                "id": sid,
                "status": "GAGAL: TIDAK ADA DI MENDELEY",
                "reason": "Terdaftar di Excel dan disitasi di Word, namun metadata/DOI/Judul tidak ditemukan di library Mendeley",
                "count": count,
                "match_method": "Tidak Ditemukan",
                "uuid": "-",
                "title": r.get("title", ""),
                "authors": auth_str,
                "year": r.get("year", ""),
                "journal": r.get("journal", ""),
                "doi": r.get("doi", ""),
                "snippet": snippet,
            })
        else:
            all_word_citations.append({
                "id": sid,
                "status": "GAGAL: TIDAK ADA DI EXCEL",
                "reason": "Placeholder ditemukan di dokumen Word, tetapi kodenya tidak terdaftar pada tabel file Excel",
                "count": count,
                "match_method": "-",
                "uuid": "-",
                "title": "(Belum ada di Excel)",
                "authors": "-",
                "year": "-",
                "journal": "-",
                "doi": "-",
                "snippet": snippet,
            })

    # Referensi di Excel yang tidak disitasi di Word
    excel_unused = []
    for sid in sorted(set(refs.keys()) - all_word_ids):
        r = refs.get(sid, {})
        auth_str = ", ".join(a.get("family", "") for a in r.get("authors", []))
        excel_unused.append({
            "id": sid,
            "title": r.get("title", ""),
            "authors": auth_str,
            "year": r.get("year", ""),
            "journal": r.get("journal", ""),
            "doi": r.get("doi", ""),
        })

    summary_metrics = {
        "citations_converted": stats["citations"],
        "total_word_unique": len(all_word_ids),
        "success_count": len(stats["codes"]),
        "fail_mendeley_count": len(stats["unresolved_mendeley"]),
        "fail_excel_count": len(stats["unresolved_excel"]),
    }

    # Generate Laporan Excel
    excel_report_path = out.with_name("laporan_sitasi.xlsx")
    generate_excel_report(
        path=excel_report_path,
        docx_path=docx,
        out_path=out,
        excel_path=workbook,
        sheet_used=sheet_used,
        db_name=db_display,
        summary_metrics=summary_metrics,
        all_word_citations=all_word_citations,
        excel_unused=excel_unused
    )

    # Tampilkan Ringkasan di Terminal
    print("\n" + "=" * 72)
    print(f"  HASIL EKSEKUSI:")
    print(f"  * {stats['citations']} kelompok sitasi berhasil diubah menjadi field Mendeley")
    print(f"  * {len(stats['codes'])} ID unik sitasi sukses terhubung ke Mendeley")
    if stats["unresolved_mendeley"]:
        print(f"  * {len(stats['unresolved_mendeley'])} ID belum ada di Mendeley : {', '.join(sorted(stats['unresolved_mendeley']))}")
    if stats["unresolved_excel"]:
        print(f"  * {len(stats['unresolved_excel'])} ID belum ada di Excel    : {', '.join(sorted(stats['unresolved_excel']))}")
    print(f"\n  Berkas Hasil Word   : {out.name}")
    print(f"  Berkas Laporan Excel: {excel_report_path.name}")
    print("=" * 72)

    # File CSV Rekap (Kompatibilitas)
    rpt = out.with_name("mapping_report.csv")
    with open(rpt, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["ID", "status", "alasan", "uuid_mendeley", "judul"])
        for r in all_word_citations:
            w.writerow([r["id"], r["status"], r["reason"], r["uuid"], r["title"]])

    # File RIS untuk referensi yang belum ada di Mendeley
    missing_in_mendeley = sorted(stats["unresolved_mendeley"])
    if missing_in_mendeley:
        missing_refs = [refs[c] for c in missing_in_mendeley if c in refs]
        if missing_refs:
            ris = out.with_name("referensi_belum_ada.ris")
            write_ris(ris, missing_refs)
            print(f"  Berkas RIS          : {ris.name}  (drag & drop langsung ke Mendeley)")

    # File Teks Perlu Diperiksa
    leftovers = []
    for data in final_parts.values():
        leftovers += collect_leftovers(data, pattern)

    if missing_in_mendeley or stats["unresolved_excel"] or leftovers:
        chk = out.with_name("PERLU_DIPERIKSA.txt")
        with open(chk, "w", encoding="utf-8") as f:
            f.write("PANDUAN PERBAIKAN SITASI YANG BELUM LENGKAP\n" + "=" * 65 + "\n\n")
            if missing_in_mendeley:
                f.write(f"1. DOKUMEN BELUM ADA DI MENDELEY ({len(missing_in_mendeley)} ID):\n")
                f.write("   Buka aplikasi Mendeley Reference Manager, lalu drag & drop file\n")
                f.write("   'referensi_belum_ada.ris' ke dalamnya. Setelah itu jalankan Run.bat lagi.\n\n")
                for code in missing_in_mendeley:
                    r = refs.get(code, {})
                    f.write(f"   * {code}: {r.get('title', '')}\n")
                    f.write(f"     DOI: {r.get('doi') or '(kosong)'} | Jurnal: {r.get('journal') or ''}\n")
                f.write("\n")

            if stats["unresolved_excel"]:
                f.write(f"2. KODE BELUM ADA DI EXCEL ({len(stats['unresolved_excel'])} ID):\n")
                f.write("   Kode berikut ada di naskah Word tetapi tidak terdaftar di file Excel:\n")
                for code in sorted(stats["unresolved_excel"]):
                    snip = snippet_tracker.get(code, {}).get("snippets", [""])[0]
                    f.write(f"   * {code} -> Konteks: {snip}\n")
                f.write("\n")

            if leftovers:
                f.write("=" * 65 + "\n")
                f.write("3. CUPLIKAN PARAGRAF YANG MASIH MENYISAKAN PLACEHOLDER:\n\n")
                for i, txt in enumerate(leftovers, 1):
                    f.write(f"[{i}] {txt[:400]}\n\n")
        print(f"  Catatan Perbaikan   : {chk.name}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
