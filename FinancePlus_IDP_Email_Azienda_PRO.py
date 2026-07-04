# -*- coding: utf-8 -*-
"""
FinancePlus IDP Email Azienda PRO
Web app Streamlit/Python per archiviazione intelligente documenti, email e allegati.

Funzioni principali:
- OCR locale/cloud predisposto
- AI/rule-based document classification
- archivio automatico cliente
- database SQLite
- controllo duplicati SHA256
- coda da verificare
- comando CERCA AZIENDA per email e allegati
- VEDI TUTTO: anteprima, sintesi intelligente, selezione multipla
- SCARICA TUTTO / SCARICA SELEZIONATE: salva email PDF + allegati nella cartella del cliente
- struttura archivio email: Mittente / Azienda / documenti
- nomi file con data ricezione email in maiuscolo: _06_MAGGIO_2026

Avvio:
    streamlit run FinancePlus_IDP_Email_Azienda_PRO.py

Nota sicurezza:
    Non inserire password o chiavi API nel codice. Usa variabili d'ambiente o sidebar.
"""

from __future__ import annotations

import base64
import csv
import dataclasses
import email
import hashlib
import html
import imaplib
import io
import json
import os
import re
import shutil
import sqlite3
import tempfile
import textwrap
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime, date
from email.header import decode_header, make_header
from email.message import Message
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

try:
    import streamlit as st
except Exception:  # pragma: no cover
    st = None

try:
    import pandas as pd
except Exception:  # pragma: no cover
    pd = None

try:
    import fitz  # PyMuPDF
except Exception:  # pragma: no cover
    fitz = None

try:
    import pytesseract
except Exception:  # pragma: no cover
    pytesseract = None

try:
    from PIL import Image
except Exception:  # pragma: no cover
    Image = None

try:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
except Exception:  # pragma: no cover
    SimpleDocTemplate = None


APP_NAME = "FinancePlus IDP Email Azienda PRO"
APP_VERSION = "1.0.0"
DEFAULT_BASE_DIR = Path.home() / "FinancePlus_IDP_Archivio"
DB_NAME = "financeplus_idp_email.db"
EMAIL_ARCHIVE_DIR = "Archivio_Email_Aziende"
DOCUMENT_ARCHIVE_DIR = "Archivio_Documenti"
VERIFY_DIR = "Da_Verificare"
REPORT_DIR = "Report"
BACKUP_DIR = "Backup"

ITALIAN_MONTHS = {
    1: "GENNAIO",
    2: "FEBBRAIO",
    3: "MARZO",
    4: "APRILE",
    5: "MAGGIO",
    6: "GIUGNO",
    7: "LUGLIO",
    8: "AGOSTO",
    9: "SETTEMBRE",
    10: "OTTOBRE",
    11: "NOVEMBRE",
    12: "DICEMBRE",
}

DOCUMENT_CATEGORIES = {
    "Visura": ["visura", "camera di commercio", "registro imprese", "rea", "ateco", "capitale sociale"],
    "Bilancio": ["bilancio", "stato patrimoniale", "conto economico", "nota integrativa", "xbrl", "ebitda"],
    "Centrale_Rischi": ["centrale rischi", "banca d'italia", "accordato", "utilizzato", "sconfino", "garanzie"],
    "DURC": ["durc", "inps", "inail", "regolarita contributiva", "protocollo durc"],
    "Documento_Identita": ["carta di identita", "passaporto", "patente", "documento identita"],
    "Contratto": ["contratto", "decorrenza", "durata", "parti contraenti", "clausola"],
    "Fattura": ["fattura", "imponibile", "iva", "totale documento", "numero fattura"],
    "Estratto_Conto": ["estratto conto", "iban", "saldo iniziale", "saldo finale", "movimenti"],
    "Preventivo": ["preventivo", "offerta", "validita offerta", "proposta economica"],
    "Business_Plan": ["business plan", "piano economico", "previsionale", "investimento", "dscr"],
    "Report_Bancario": ["report bancario", "finanziamento", "garanzia", "mcc", "rating"],
    "Dichiarazione_Fiscale": ["dichiarazione", "redditi", "irap", "iva", "modello unico", "f24"],
}

SKIP_INLINE_IMAGE_MIMES = {
    "image/png",
    "image/jpeg",
    "image/gif",
    "image/webp",
}


@dataclass
class EmailAttachment:
    filename: str
    content_type: str
    size: int
    data: bytes
    content_id: str = ""
    is_inline: bool = False


@dataclass
class EmailItem:
    uid: str
    mailbox: str
    sender_name: str
    sender_email: str
    subject: str
    received_at: datetime
    body_text: str
    body_html: str
    attachments: List[EmailAttachment]
    company_score: float = 0.0
    company_reasons: str = ""
    smart_summary: str = ""


# -----------------------------------------------------------------------------
# General utilities
# -----------------------------------------------------------------------------


def decode_mime_header(value: Optional[str]) -> str:
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return value


def normalize_text(value: str) -> str:
    if value is None:
        return ""
    value = unicodedata.normalize("NFKD", str(value))
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = value.lower()
    value = re.sub(r"[^a-z0-9@._+\-/\s]", " ", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def safe_filename(value: str, fallback: str = "file", max_len: int = 120) -> str:
    value = decode_mime_header(value or "")
    value = unicodedata.normalize("NFKD", value)
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = re.sub(r"[\\/:*?\"<>|]", "_", value)
    value = re.sub(r"\s+", "_", value.strip())
    value = re.sub(r"_+", "_", value)
    value = value.strip("._ ")
    if not value:
        value = fallback
    if len(value) > max_len:
        stem, ext = os.path.splitext(value)
        value = stem[: max_len - len(ext) - 1].rstrip("_") + ext
    return value


def company_folder_name(company: str) -> str:
    cleaned = safe_filename(company, fallback="AZIENDA_NON_RICONOSCIUTA", max_len=100)
    return cleaned.upper()


def sender_folder_name(sender_email: str, sender_name: str = "") -> str:
    raw = sender_email or sender_name or "MITTENTE_SCONOSCIUTO"
    return safe_filename(raw, fallback="MITTENTE_SCONOSCIUTO", max_len=100).lower()


def italian_date_suffix(dt: datetime) -> str:
    return f"{dt.day:02d}_{ITALIAN_MONTHS.get(dt.month, str(dt.month))}_{dt.year}"


def append_date_suffix(filename: str, dt: datetime) -> str:
    filename = safe_filename(filename, fallback="documento")
    stem, ext = os.path.splitext(filename)
    suffix = italian_date_suffix(dt)
    stem = re.sub(r"_\d{2}_[A-Z]+_\d{4}$", "", stem, flags=re.IGNORECASE)
    return f"{stem}_{suffix}{ext or '.bin'}"


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    i = 2
    while True:
        candidate = parent / f"{stem}_{i}{suffix}"
        if not candidate.exists():
            return candidate
        i += 1


def sha256_bytes(data: bytes) -> str:
    h = hashlib.sha256()
    h.update(data)
    return h.hexdigest()


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def ensure_base_structure(base_dir: Path) -> Dict[str, Path]:
    paths = {
        "base": base_dir,
        "email_archive": base_dir / EMAIL_ARCHIVE_DIR,
        "document_archive": base_dir / DOCUMENT_ARCHIVE_DIR,
        "verify": base_dir / VERIFY_DIR,
        "report": base_dir / REPORT_DIR,
        "backup": base_dir / BACKUP_DIR,
    }
    for p in paths.values():
        p.mkdir(parents=True, exist_ok=True)
    return paths


def strip_html_to_text(value: str) -> str:
    if not value:
        return ""
    text = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", value)
    text = re.sub(r"(?i)<br\s*/?>", "\n", text)
    text = re.sub(r"(?i)</p>", "\n", text)
    text = re.sub(r"<[^>]+>", " ", text)
    text = html.unescape(text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n\s+", "\n", text)
    return text.strip()


def truncate_text(value: str, max_chars: int = 3500) -> str:
    value = value or ""
    if len(value) <= max_chars:
        return value
    return value[:max_chars].rstrip() + "..."


def split_keywords(company: str) -> List[str]:
    normalized = normalize_text(company)
    tokens = [t for t in re.split(r"\s+", normalized) if len(t) >= 3]
    legal_noise = {"srl", "spa", "sas", "snc", "soc", "coop", "ltd", "s", "r", "l", "s.p.a", "s.r.l"}
    tokens = [t for t in tokens if t not in legal_noise]
    variants = [normalized] + tokens
    return list(dict.fromkeys([v for v in variants if v]))


# -----------------------------------------------------------------------------
# SQLite database
# -----------------------------------------------------------------------------


def db_path(base_dir: Path) -> Path:
    return base_dir / DB_NAME


def connect_db(base_dir: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(str(db_path(base_dir)))
    conn.row_factory = sqlite3.Row
    return conn


def init_db(base_dir: Path) -> None:
    ensure_base_structure(base_dir)
    with connect_db(base_dir) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS clienti (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ragione_sociale TEXT,
                partita_iva TEXT,
                codice_fiscale TEXT,
                sede TEXT,
                amministratore TEXT,
                ateco TEXT,
                email TEXT,
                pec TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(ragione_sociale, partita_iva)
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS documenti (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cliente_id INTEGER,
                ragione_sociale TEXT,
                categoria TEXT,
                nome_originale TEXT,
                nome_archiviato TEXT,
                percorso TEXT,
                data_documento TEXT,
                importo TEXT,
                hash_sha256 TEXT UNIQUE,
                testo_estratto TEXT,
                stato TEXT,
                fonte TEXT,
                mittente TEXT,
                email_uid TEXT,
                data_importazione TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS email_archiviate (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                uid TEXT,
                mailbox TEXT,
                mittente_nome TEXT,
                mittente_email TEXT,
                azienda TEXT,
                oggetto TEXT,
                data_ricezione TEXT,
                percorso_pdf TEXT,
                sintesi TEXT,
                hash_contenuto TEXT UNIQUE,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS coda_verificare (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_path TEXT,
                testo_estratto TEXT,
                categoria_suggerita TEXT,
                cliente_suggerito TEXT,
                motivo_incertezza TEXT,
                stato TEXT DEFAULT 'Da verificare',
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS log_attivita (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                data TEXT DEFAULT CURRENT_TIMESTAMP,
                utente TEXT,
                operazione TEXT,
                documento TEXT,
                risultato TEXT,
                note TEXT
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS apprendimento (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                keyword TEXT,
                categoria_corretta TEXT,
                cliente_corretto TEXT,
                regola_appresa TEXT,
                fonte TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        conn.commit()


def log_activity(base_dir: Path, operation: str, document: str = "", result: str = "OK", notes: str = "", user: str = "utente") -> None:
    try:
        with connect_db(base_dir) as conn:
            conn.execute(
                "INSERT INTO log_attivita (utente, operazione, documento, risultato, note) VALUES (?, ?, ?, ?, ?)",
                (user, operation, document, result, notes),
            )
            conn.commit()
    except Exception:
        pass


def find_or_create_client(base_dir: Path, ragione_sociale: str, piva: str = "", cf: str = "", email_addr: str = "") -> int:
    ragione_sociale = ragione_sociale.strip() or "AZIENDA_NON_RICONOSCIUTA"
    with connect_db(base_dir) as conn:
        cur = conn.cursor()
        if piva:
            row = cur.execute("SELECT id FROM clienti WHERE partita_iva = ?", (piva,)).fetchone()
            if row:
                return int(row["id"])
        row = cur.execute("SELECT id FROM clienti WHERE lower(ragione_sociale) = lower(?)", (ragione_sociale,)).fetchone()
        if row:
            return int(row["id"])
        cur.execute(
            """
            INSERT INTO clienti (ragione_sociale, partita_iva, codice_fiscale, email)
            VALUES (?, ?, ?, ?)
            """,
            (ragione_sociale, piva, cf, email_addr),
        )
        conn.commit()
        return int(cur.lastrowid)


def insert_document_record(
    base_dir: Path,
    cliente_id: Optional[int],
    ragione_sociale: str,
    categoria: str,
    nome_originale: str,
    nome_archiviato: str,
    percorso: str,
    hash_value: str,
    testo: str = "",
    stato: str = "Archiviato",
    fonte: str = "documento",
    mittente: str = "",
    email_uid: str = "",
    data_documento: str = "",
    importo: str = "",
) -> bool:
    try:
        with connect_db(base_dir) as conn:
            conn.execute(
                """
                INSERT INTO documenti (
                    cliente_id, ragione_sociale, categoria, nome_originale, nome_archiviato,
                    percorso, data_documento, importo, hash_sha256, testo_estratto, stato,
                    fonte, mittente, email_uid
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    cliente_id,
                    ragione_sociale,
                    categoria,
                    nome_originale,
                    nome_archiviato,
                    percorso,
                    data_documento,
                    importo,
                    hash_value,
                    truncate_text(testo, 20000),
                    stato,
                    fonte,
                    mittente,
                    email_uid,
                ),
            )
            conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False


def hash_exists(base_dir: Path, hash_value: str) -> Optional[sqlite3.Row]:
    with connect_db(base_dir) as conn:
        return conn.execute("SELECT * FROM documenti WHERE hash_sha256 = ?", (hash_value,)).fetchone()


# -----------------------------------------------------------------------------
# Document reading, OCR, classification and IDP
# -----------------------------------------------------------------------------


def extract_text_from_pdf(path: Path) -> str:
    if fitz is None:
        return ""
    try:
        doc = fitz.open(str(path))
        parts = []
        for page in doc:
            parts.append(page.get_text("text"))
        return "\n".join(parts).strip()
    except Exception:
        return ""


def ocr_image_path(path: Path) -> str:
    if pytesseract is None or Image is None:
        return ""
    try:
        img = Image.open(path)
        return pytesseract.image_to_string(img, lang="ita+eng")
    except Exception:
        return ""


def extract_text_from_file(path: Path) -> str:
    ext = path.suffix.lower()
    if ext == ".txt":
        try:
            return path.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            return ""
    if ext == ".pdf":
        text = extract_text_from_pdf(path)
        # OCR fallback for scanned PDFs can be implemented page-by-page if needed.
        return text
    if ext in {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp"}:
        return ocr_image_path(path)
    return ""


def extract_identifiers(text: str) -> Dict[str, str]:
    out = {"piva": "", "cf": "", "date": "", "amount": ""}
    # Italian VAT number: 11 digits, often with IT prefix.
    piva_match = re.search(r"(?:p\.?\s*iva|partita\s+iva|vat)\s*[:\-]?\s*(?:IT)?\s*(\d{11})", text, re.I)
    if not piva_match:
        piva_match = re.search(r"\b(?:IT)?(\d{11})\b", text, re.I)
    if piva_match:
        out["piva"] = piva_match.group(1)
    cf_match = re.search(r"\b[A-Z]{6}\d{2}[A-Z]\d{2}[A-Z]\d{3}[A-Z]\b", text.upper())
    if cf_match:
        out["cf"] = cf_match.group(0)
    date_match = re.search(r"\b(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b", text)
    if date_match:
        out["date"] = date_match.group(1)
    amount_match = re.search(r"(?:euro|eur|totale|importo)\s*[:\-]?\s*([0-9\.]+,[0-9]{2})", text, re.I)
    if amount_match:
        out["amount"] = amount_match.group(1)
    return out


def classify_document(text: str, filename: str = "") -> Tuple[str, float, str]:
    corpus = normalize_text(f"{filename}\n{text}")
    scores: Dict[str, int] = {}
    reasons: Dict[str, List[str]] = defaultdict(list)
    for category, keywords in DOCUMENT_CATEGORIES.items():
        for kw in keywords:
            nkw = normalize_text(kw)
            if nkw and nkw in corpus:
                scores[category] = scores.get(category, 0) + 1
                reasons[category].append(kw)
    if not scores:
        return "Documento_Generico", 0.25, "Nessuna regola forte trovata"
    category, score = max(scores.items(), key=lambda x: x[1])
    confidence = min(0.95, 0.35 + 0.12 * score)
    return category, confidence, ", ".join(reasons[category][:5])


def infer_company_from_text(text: str, fallback: str = "") -> str:
    patterns = [
        r"ragione\s+sociale\s*[:\-]?\s*([^\n\r]{3,80})",
        r"denominazione\s*[:\-]?\s*([^\n\r]{3,80})",
        r"societa\s*[:\-]?\s*([^\n\r]{3,80})",
        r"cliente\s*[:\-]?\s*([^\n\r]{3,80})",
        r"spett\.?le\s+([^\n\r]{3,80})",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.I)
        if m:
            candidate = re.sub(r"\s+", " ", m.group(1)).strip(" -;:,.\t")
            if candidate:
                return candidate[:80]
    return fallback or "AZIENDA_NON_RICONOSCIUTA"


def archive_local_document(base_dir: Path, file_path: Path, forced_company: str = "") -> Tuple[bool, str]:
    ensure_base_structure(base_dir)
    init_db(base_dir)
    if not file_path.exists() or not file_path.is_file():
        return False, f"File non trovato: {file_path}"
    file_hash = sha256_file(file_path)
    existing = hash_exists(base_dir, file_hash)
    if existing:
        return False, f"Duplicato gia archiviato: {existing['percorso']}"
    text = extract_text_from_file(file_path)
    category, confidence, reason = classify_document(text, file_path.name)
    ids = extract_identifiers(text)
    company = forced_company or infer_company_from_text(text, fallback="AZIENDA_NON_RICONOSCIUTA")
    if company == "AZIENDA_NON_RICONOSCIUTA" or confidence < 0.45:
        target_dir = base_dir / VERIFY_DIR
        target_dir.mkdir(parents=True, exist_ok=True)
        target = unique_path(target_dir / safe_filename(file_path.name))
        shutil.copy2(file_path, target)
        with connect_db(base_dir) as conn:
            conn.execute(
                """
                INSERT INTO coda_verificare (file_path, testo_estratto, categoria_suggerita, cliente_suggerito, motivo_incertezza)
                VALUES (?, ?, ?, ?, ?)
                """,
                (str(target), truncate_text(text, 20000), category, company, f"Confidenza {confidence:.0%}. Motivo: {reason}"),
            )
            conn.commit()
        log_activity(base_dir, "Documento in coda da verificare", file_path.name, "DA_VERIFICARE", reason)
        return True, f"Inserito in coda da verificare: {target}"
    cliente_id = find_or_create_client(base_dir, company, ids.get("piva", ""), ids.get("cf", ""))
    target_dir = base_dir / DOCUMENT_ARCHIVE_DIR / company_folder_name(company)
    target_dir.mkdir(parents=True, exist_ok=True)
    dated_name = append_date_suffix(file_path.name, datetime.now())
    target = unique_path(target_dir / dated_name)
    shutil.copy2(file_path, target)
    ok = insert_document_record(
        base_dir,
        cliente_id,
        company,
        category,
        file_path.name,
        target.name,
        str(target),
        file_hash,
        text,
        fonte="documento_locale",
        data_documento=ids.get("date", ""),
        importo=ids.get("amount", ""),
    )
    if ok:
        log_activity(base_dir, "Documento archiviato", target.name, "OK", f"Categoria: {category}; confidenza: {confidence:.0%}")
        return True, f"Archiviato: {target}"
    return False, "Documento copiato ma record database duplicato"


# -----------------------------------------------------------------------------
# Email functions
# -----------------------------------------------------------------------------


def parse_sender(value: str) -> Tuple[str, str]:
    value = decode_mime_header(value)
    m = re.match(r"(?:(.*?)\s*)?<([^>]+)>", value)
    if m:
        return (m.group(1).strip(' "') or m.group(2).split("@")[0], m.group(2).strip().lower())
    if "@" in value:
        return (value.split("@")[0].strip(), value.strip().lower())
    return (value.strip(), "")


def parse_email_date(value: str) -> datetime:
    try:
        dt = parsedate_to_datetime(value)
        if dt.tzinfo:
            return dt.astimezone().replace(tzinfo=None)
        return dt
    except Exception:
        return datetime.now()


def get_payload_text(part: Message) -> str:
    try:
        payload = part.get_payload(decode=True)
        if payload is None:
            return ""
        charset = part.get_content_charset() or "utf-8"
        return payload.decode(charset, errors="replace")
    except Exception:
        try:
            return str(part.get_payload())
        except Exception:
            return ""


def parse_email_message(raw_bytes: bytes, uid: str, mailbox: str) -> EmailItem:
    msg = email.message_from_bytes(raw_bytes)
    subject = decode_mime_header(msg.get("Subject", ""))
    sender_name, sender_email = parse_sender(msg.get("From", ""))
    received_at = parse_email_date(msg.get("Date", ""))
    body_text_parts: List[str] = []
    body_html_parts: List[str] = []
    attachments: List[EmailAttachment] = []

    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type().lower()
            disp = (part.get("Content-Disposition") or "").lower()
            filename = decode_mime_header(part.get_filename() or "")
            content_id = part.get("Content-ID", "") or ""
            is_attachment = "attachment" in disp or bool(filename)
            is_inline = "inline" in disp or bool(content_id)
            if is_attachment:
                data = part.get_payload(decode=True) or b""
                if not filename:
                    ext = ctype.split("/")[-1] if "/" in ctype else "bin"
                    filename = f"allegato.{ext}"
                attachments.append(
                    EmailAttachment(
                        filename=safe_filename(filename, fallback="allegato"),
                        content_type=ctype,
                        size=len(data),
                        data=data,
                        content_id=content_id,
                        is_inline=is_inline,
                    )
                )
            elif ctype == "text/plain":
                body_text_parts.append(get_payload_text(part))
            elif ctype == "text/html":
                body_html_parts.append(get_payload_text(part))
    else:
        ctype = msg.get_content_type().lower()
        if ctype == "text/html":
            body_html_parts.append(get_payload_text(msg))
        else:
            body_text_parts.append(get_payload_text(msg))

    body_text = "\n".join(p.strip() for p in body_text_parts if p.strip())
    body_html = "\n".join(p.strip() for p in body_html_parts if p.strip())
    if not body_text and body_html:
        body_text = strip_html_to_text(body_html)

    return EmailItem(
        uid=uid,
        mailbox=mailbox,
        sender_name=sender_name,
        sender_email=sender_email,
        subject=subject,
        received_at=received_at,
        body_text=body_text,
        body_html=body_html,
        attachments=attachments,
    )


def imap_login(server: str, email_user: str, password: str, port: int = 993) -> imaplib.IMAP4_SSL:
    mail = imaplib.IMAP4_SSL(server, port)
    mail.login(email_user, password)
    return mail


def list_mailboxes(mail: imaplib.IMAP4_SSL) -> List[str]:
    status, data = mail.list()
    if status != "OK":
        return ["INBOX"]
    boxes = []
    for raw in data or []:
        line = raw.decode(errors="ignore")
        # Last quoted component is usually the mailbox name.
        matches = re.findall(r'"([^"]+)"', line)
        if matches:
            boxes.append(matches[-1])
        else:
            parts = line.split()
            if parts:
                boxes.append(parts[-1].strip('"'))
    preferred = ["INBOX"]
    for b in boxes:
        if b not in preferred:
            preferred.append(b)
    return preferred


def imap_search_uids(mail: imaplib.IMAP4_SSL, mailbox: str, company: str, max_results: int = 50) -> List[str]:
    try:
        status, _ = mail.select(mailbox, readonly=True)
        if status != "OK":
            return []
    except Exception:
        return []

    # Try a broad IMAP search first. Some servers support SUBJECT/BODY with UTF-8 poorly;
    # fallback to ALL is used when this fails.
    search_terms = split_keywords(company)
    uid_set: List[bytes] = []
    for term in search_terms[:4]:
        try:
            encoded = term.encode("utf-8")
            status, data = mail.uid("SEARCH", "CHARSET", "UTF-8", "OR", "SUBJECT", encoded, "BODY", encoded)
            if status == "OK" and data and data[0]:
                uid_set.extend(data[0].split())
        except Exception:
            continue
    if not uid_set:
        try:
            status, data = mail.uid("SEARCH", None, "ALL")
            if status == "OK" and data and data[0]:
                uid_set = data[0].split()[-max_results * 5 :]
        except Exception:
            uid_set = []
    # newest first by UID order
    decoded = [u.decode() if isinstance(u, bytes) else str(u) for u in uid_set]
    decoded = list(dict.fromkeys(decoded))
    return decoded[-max_results:][::-1]


def fetch_email_raw(mail: imaplib.IMAP4_SSL, uid: str) -> Optional[bytes]:
    try:
        status, data = mail.uid("FETCH", uid, "(RFC822)")
        if status != "OK":
            return None
        for item in data:
            if isinstance(item, tuple) and item[1]:
                return item[1]
    except Exception:
        return None
    return None


def score_email_for_company(item: EmailItem, company: str) -> Tuple[float, str]:
    keywords = split_keywords(company)
    subject_norm = normalize_text(item.subject)
    body_norm = normalize_text(item.body_text)
    attach_norm = normalize_text(" ".join(a.filename for a in item.attachments))
    score = 0.0
    reasons = []
    for kw in keywords:
        if not kw:
            continue
        if kw in subject_norm:
            score += 0.45
            reasons.append(f"oggetto contiene '{kw}'")
        if kw in body_norm:
            score += 0.35
            reasons.append(f"testo mail contiene '{kw}'")
        if kw in attach_norm:
            score += 0.20
            reasons.append(f"nome allegato contiene '{kw}'")
    if normalize_text(company) in subject_norm:
        score += 0.25
    if normalize_text(company) in body_norm:
        score += 0.15
    score = min(score, 1.0)
    if not reasons:
        reasons.append("nessuna corrispondenza forte nell'oggetto, corpo o allegati")
    return score, "; ".join(dict.fromkeys(reasons))


def attachment_text_preview(att: EmailAttachment, max_chars: int = 1200) -> str:
    ext = Path(att.filename).suffix.lower()
    if ext in {".txt", ".csv", ".xml", ".json"} or att.content_type.startswith("text/"):
        try:
            return att.data.decode("utf-8", errors="replace")[:max_chars]
        except Exception:
            return ""
    if ext == ".pdf" and fitz is not None:
        try:
            doc = fitz.open(stream=att.data, filetype="pdf")
            text = "\n".join(page.get_text("text") for page in doc[:3])
            return text[:max_chars]
        except Exception:
            return ""
    return ""


def summarize_text_rule_based(text: str, subject: str = "", max_sentences: int = 4) -> str:
    text = strip_html_to_text(text)
    text = re.sub(r"\s+", " ", text).strip()
    if not text:
        return "Nessun contenuto testuale leggibile."
    sentences = re.split(r"(?<=[.!?])\s+", text)
    sentences = [s.strip() for s in sentences if len(s.strip()) > 20]
    if not sentences:
        return truncate_text(text, 400)
    keywords = ["document", "bilancio", "visura", "fattura", "centrale rischi", "finanziamento", "preventivo", "contratto", "alleg", "azienda", "cliente", "banca"]
    ranked = []
    for s in sentences[:30]:
        ns = normalize_text(s)
        score = sum(1 for k in keywords if k in ns) + min(len(s) / 200, 2)
        ranked.append((score, s))
    best = [s for _, s in sorted(ranked, key=lambda x: x[0], reverse=True)[:max_sentences]]
    return " ".join(best)


def summarize_email(item: EmailItem, use_openai: bool = False, openai_api_key: str = "") -> str:
    attachment_notes = []
    for att in item.attachments[:5]:
        preview = attachment_text_preview(att, 800)
        if preview:
            attachment_notes.append(f"Allegato {att.filename}: {preview}")
    combined = f"Oggetto: {item.subject}\n\nCorpo mail:\n{item.body_text}\n\n" + "\n".join(attachment_notes)
    if use_openai and openai_api_key:
        try:
            from openai import OpenAI
            client = OpenAI(api_key=openai_api_key)
            resp = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "Sintetizza email e allegati per un consulente finanziario. Evidenzia azienda, documenti, scadenze, richieste, importi e criticita. Rispondi in italiano."},
                    {"role": "user", "content": truncate_text(combined, 12000)},
                ],
                temperature=0.1,
                max_tokens=350,
            )
            return resp.choices[0].message.content.strip()
        except Exception as exc:
            return f"Sintesi locale usata per errore IA: {exc}\n\n" + summarize_text_rule_based(combined, item.subject)
    return summarize_text_rule_based(combined, item.subject)


def search_company_emails(
    server: str,
    username: str,
    password: str,
    company: str,
    mailboxes: Sequence[str],
    max_results_per_box: int = 50,
    min_score: float = 0.25,
    use_openai: bool = False,
    openai_api_key: str = "",
) -> List[EmailItem]:
    mail = imap_login(server, username, password)
    try:
        results: List[EmailItem] = []
        for box in mailboxes:
            uids = imap_search_uids(mail, box, company, max_results=max_results_per_box)
            for uid in uids:
                raw = fetch_email_raw(mail, uid)
                if not raw:
                    continue
                item = parse_email_message(raw, uid, box)
                score, reasons = score_email_for_company(item, company)
                if score >= min_score:
                    item.company_score = score
                    item.company_reasons = reasons
                    item.smart_summary = summarize_email(item, use_openai, openai_api_key)
                    results.append(item)
        # Deduplicate by mailbox/uid and sort newest first
        seen = set()
        unique = []
        for item in results:
            key = (item.mailbox, item.uid)
            if key not in seen:
                seen.add(key)
                unique.append(item)
        unique.sort(key=lambda x: x.received_at, reverse=True)
        return unique
    finally:
        try:
            mail.logout()
        except Exception:
            pass


# -----------------------------------------------------------------------------
# PDF creation for email print
# -----------------------------------------------------------------------------


def pdf_available() -> bool:
    return SimpleDocTemplate is not None


def create_email_pdf(item: EmailItem, company: str, output_path: Path) -> None:
    if not pdf_available():
        # Fallback: save text with .pdf extension is not valid; use .txt beside expected name.
        txt_path = output_path.with_suffix(".txt")
        txt_path.write_text(
            f"MITTENTE: {item.sender_name} <{item.sender_email}>\nDATA: {item.received_at}\nOGGETTO: {item.subject}\nAZIENDA: {company}\n\n{item.body_text}\n\nALLEGATI:\n" + "\n".join(a.filename for a in item.attachments),
            encoding="utf-8",
        )
        return

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        rightMargin=1.6 * cm,
        leftMargin=1.6 * cm,
        topMargin=1.4 * cm,
        bottomMargin=1.4 * cm,
    )
    styles = getSampleStyleSheet()
    title = ParagraphStyle(
        "FPTitle",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=16,
        leading=20,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#17365D"),
        spaceAfter=12,
    )
    h = ParagraphStyle(
        "FPH",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=11,
        textColor=colors.HexColor("#17365D"),
        spaceBefore=10,
        spaceAfter=6,
    )
    body = ParagraphStyle(
        "FPBody",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=9,
        leading=12,
        alignment=TA_LEFT,
    )
    small = ParagraphStyle(
        "FPSmall",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=8,
        leading=10,
    )

    story: List[Any] = []
    story.append(Paragraph("FINANCEPLUS - STAMPA PDF EMAIL", title))
    meta = [
        ["Azienda", company],
        ["Mittente", f"{item.sender_name} <{item.sender_email}>"] ,
        ["Data ricezione", item.received_at.strftime("%d/%m/%Y %H:%M")],
        ["Oggetto", item.subject],
        ["Mailbox", item.mailbox],
    ]
    table = Table(meta, colWidths=[3.5 * cm, 13 * cm])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#EAF0F7")),
        ("TEXTCOLOR", (0, 0), (-1, -1), colors.HexColor("#111111")),
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTNAME", (1, 0), (1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#B7C7D9")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(table)
    story.append(Spacer(1, 0.25 * cm))
    story.append(Paragraph("Sintesi intelligente", h))
    story.append(Paragraph(html.escape(item.smart_summary or summarize_email(item)), body))
    story.append(Paragraph("Contenuto email", h))
    clean_body = item.body_text or strip_html_to_text(item.body_html) or "Nessun contenuto testuale."
    for chunk in textwrap.wrap(clean_body, width=1000, break_long_words=False, replace_whitespace=False):
        story.append(Paragraph(html.escape(chunk).replace("\n", "<br/>"), body))
        story.append(Spacer(1, 0.08 * cm))
    story.append(Paragraph("Lista allegati", h))
    if item.attachments:
        data = [["Nome file", "Tipo", "Dimensione"]]
        for att in item.attachments:
            data.append([att.filename, att.content_type, f"{att.size / 1024:.1f} KB"])
        tab = Table(data, colWidths=[9 * cm, 4 * cm, 3 * cm])
        tab.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#17365D")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#CCCCCC")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        story.append(tab)
    else:
        story.append(Paragraph("Nessun allegato presente.", small))
    doc.build(story)


def archive_email_item(
    base_dir: Path,
    item: EmailItem,
    company: str,
    skip_inline_images: bool = True,
    create_client: bool = True,
) -> Tuple[bool, str, List[str]]:
    ensure_base_structure(base_dir)
    init_db(base_dir)
    sender_dir = sender_folder_name(item.sender_email, item.sender_name)
    company_dir = company_folder_name(company)
    target_dir = base_dir / EMAIL_ARCHIVE_DIR / sender_dir / company_dir
    target_dir.mkdir(parents=True, exist_ok=True)
    saved: List[str] = []
    dt_suffix = italian_date_suffix(item.received_at)
    email_pdf_name = f"MAIL_{dt_suffix}.pdf"
    email_pdf_path = unique_path(target_dir / email_pdf_name)
    create_email_pdf(item, company, email_pdf_path)
    saved.append(str(email_pdf_path))

    cliente_id = find_or_create_client(base_dir, company, email_addr=item.sender_email) if create_client else None
    email_hash = sha256_bytes(f"{item.sender_email}|{item.received_at.isoformat()}|{item.subject}|{item.body_text}".encode("utf-8", errors="ignore"))
    try:
        with connect_db(base_dir) as conn:
            conn.execute(
                """
                INSERT OR IGNORE INTO email_archiviate (
                    uid, mailbox, mittente_nome, mittente_email, azienda, oggetto, data_ricezione,
                    percorso_pdf, sintesi, hash_contenuto
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    item.uid,
                    item.mailbox,
                    item.sender_name,
                    item.sender_email,
                    company,
                    item.subject,
                    item.received_at.isoformat(),
                    str(email_pdf_path),
                    item.smart_summary,
                    email_hash,
                ),
            )
            conn.commit()
    except Exception:
        pass

    # Save the email PDF as a document record too.
    try:
        pdf_hash = sha256_file(email_pdf_path)
        insert_document_record(
            base_dir,
            cliente_id,
            company,
            "Email_PDF",
            item.subject,
            email_pdf_path.name,
            str(email_pdf_path),
            pdf_hash,
            item.body_text,
            fonte="email",
            mittente=item.sender_email,
            email_uid=item.uid,
            data_documento=item.received_at.strftime("%d/%m/%Y"),
        )
    except Exception:
        pass

    for att in item.attachments:
        if skip_inline_images and att.is_inline and att.content_type in SKIP_INLINE_IMAGE_MIMES:
            continue
        if not att.data:
            continue
        att_hash = sha256_bytes(att.data)
        existing = hash_exists(base_dir, att_hash)
        if existing:
            log_activity(base_dir, "Allegato duplicato bloccato", att.filename, "DUPLICATO", str(existing["percorso"]))
            continue
        dated_name = append_date_suffix(att.filename, item.received_at)
        att_path = unique_path(target_dir / dated_name)
        att_path.write_bytes(att.data)
        saved.append(str(att_path))
        text_preview = attachment_text_preview(att, 6000)
        category, confidence, reason = classify_document(text_preview, att.filename)
        insert_document_record(
            base_dir,
            cliente_id,
            company,
            category,
            att.filename,
            att_path.name,
            str(att_path),
            att_hash,
            text_preview,
            fonte="allegato_email",
            mittente=item.sender_email,
            email_uid=item.uid,
            data_documento=item.received_at.strftime("%d/%m/%Y"),
        )
    log_activity(base_dir, "Email archiviata", item.subject, "OK", f"Azienda: {company}; file salvati: {len(saved)}")
    return True, str(target_dir), saved


# -----------------------------------------------------------------------------
# Streamlit UI
# -----------------------------------------------------------------------------


def render_header() -> None:
    st.set_page_config(page_title=APP_NAME, page_icon="📁", layout="wide")
    st.markdown(
        """
        <style>
        .main .block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
        .fp-title {font-size: 30px; font-weight: 800; color: #17365D; margin-bottom: 0.1rem;}
        .fp-sub {font-size: 14px; color: #666; margin-bottom: 1rem;}
        .fp-card {border: 1px solid #E5EAF0; border-radius: 14px; padding: 16px; background: #FFFFFF; box-shadow: 0 2px 8px rgba(0,0,0,0.04);}
        .fp-kpi {font-size: 26px; font-weight: 800; color: #17365D;}
        .fp-label {font-size: 13px; color: #666;}
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(f"<div class='fp-title'>{APP_NAME}</div>", unsafe_allow_html=True)
    st.markdown("<div class='fp-sub'>Archivio intelligente documenti, email e allegati - OCR, IA, IDP, ricerca azienda e download automatico.</div>", unsafe_allow_html=True)


def sidebar_settings() -> Dict[str, Any]:
    st.sidebar.header("⚙️ Configurazione")
    base_dir_str = st.sidebar.text_input("Cartella archivio principale", value=str(DEFAULT_BASE_DIR))
    base_dir = Path(base_dir_str).expanduser()
    init_db(base_dir)
    st.sidebar.caption("Le password non vengono salvate nel codice. Usa app password IMAP/Gmail quando necessario.")
    server = st.sidebar.text_input("Server IMAP", value=os.getenv("FINANCEPLUS_IMAP_SERVER", "imap.gmail.com"))
    username = st.sidebar.text_input("Email", value=os.getenv("FINANCEPLUS_IMAP_USER", ""))
    password = st.sidebar.text_input("Password/App password", value=os.getenv("FINANCEPLUS_IMAP_PASSWORD", ""), type="password")
    mailboxes_raw = st.sidebar.text_input("Mailbox da cercare", value="INBOX")
    max_results = st.sidebar.number_input("Max email per mailbox", min_value=5, max_value=500, value=80, step=5)
    min_score = st.sidebar.slider("Soglia riconoscimento azienda", min_value=0.10, max_value=0.95, value=0.25, step=0.05)
    skip_inline_images = st.sidebar.checkbox("Non salvare immagini inline/logo", value=True)
    st.sidebar.divider()
    use_openai = st.sidebar.checkbox("Usa sintesi IA OpenAI se disponibile", value=False)
    openai_api_key = st.sidebar.text_input("OPENAI_API_KEY", value=os.getenv("OPENAI_API_KEY", ""), type="password")
    return {
        "base_dir": base_dir,
        "server": server,
        "username": username,
        "password": password,
        "mailboxes": [m.strip() for m in mailboxes_raw.split(",") if m.strip()] or ["INBOX"],
        "max_results": int(max_results),
        "min_score": float(min_score),
        "skip_inline_images": skip_inline_images,
        "use_openai": use_openai,
        "openai_api_key": openai_api_key,
    }


def kpi(label: str, value: Any) -> None:
    st.markdown(f"<div class='fp-card'><div class='fp-kpi'>{value}</div><div class='fp-label'>{label}</div></div>", unsafe_allow_html=True)


def dashboard(settings: Dict[str, Any]) -> None:
    base_dir = settings["base_dir"]
    init_db(base_dir)
    with connect_db(base_dir) as conn:
        clienti = conn.execute("SELECT COUNT(*) AS c FROM clienti").fetchone()["c"]
        documenti = conn.execute("SELECT COUNT(*) AS c FROM documenti").fetchone()["c"]
        email_count = conn.execute("SELECT COUNT(*) AS c FROM email_archiviate").fetchone()["c"]
        verify = conn.execute("SELECT COUNT(*) AS c FROM coda_verificare WHERE stato <> 'Archiviato'").fetchone()["c"]
        logs = conn.execute("SELECT * FROM log_attivita ORDER BY id DESC LIMIT 10").fetchall()
    c1, c2, c3, c4 = st.columns(4)
    with c1: kpi("Clienti", clienti)
    with c2: kpi("Documenti", documenti)
    with c3: kpi("Email archiviate", email_count)
    with c4: kpi("Da verificare", verify)
    st.subheader("Ultime attività")
    if logs:
        df = pd.DataFrame([dict(r) for r in logs]) if pd is not None else [dict(r) for r in logs]
        st.dataframe(df, use_container_width=True)
    else:
        st.info("Nessuna attività registrata.")
    st.subheader("Struttura archivio")
    st.code(str(base_dir), language="text")


def tab_cerca_azienda(settings: Dict[str, Any]) -> None:
    st.subheader("🔎 CERCA AZIENDA")
    st.write("Scrivi il nome azienda. Il sistema cerca nell'oggetto, nel contenuto della mail e nei nomi degli allegati.")
    company = st.text_input("Nome azienda", placeholder="Esempio: BelGarden, PELCOM, STC, ETS GROUP")
    col_a, col_b = st.columns([1, 1])
    with col_a:
        btn_view = st.button("VEDI TUTTO", type="primary", use_container_width=True)
    with col_b:
        btn_download_all = st.button("SCARICA TUTTO", use_container_width=True)

    if (btn_view or btn_download_all) and not company.strip():
        st.error("Inserisci il nome dell'azienda.")
        return
    if (btn_view or btn_download_all) and (not settings["username"] or not settings["password"]):
        st.error("Configura email e password/app password IMAP nella sidebar.")
        return

    if btn_view or btn_download_all:
        with st.spinner("Ricerca email e allegati in corso..."):
            try:
                results = search_company_emails(
                    settings["server"],
                    settings["username"],
                    settings["password"],
                    company.strip(),
                    settings["mailboxes"],
                    max_results_per_box=settings["max_results"],
                    min_score=settings["min_score"],
                    use_openai=settings["use_openai"],
                    openai_api_key=settings["openai_api_key"],
                )
                st.session_state["company_search_results"] = results
                st.session_state["company_search_name"] = company.strip()
            except Exception as exc:
                st.error(f"Errore ricerca IMAP: {exc}")
                return

    results: List[EmailItem] = st.session_state.get("company_search_results", [])
    current_company = st.session_state.get("company_search_name", company.strip())

    if btn_download_all and results:
        with st.spinner("Scarico tutte le email e gli allegati nella cartella del cliente..."):
            all_saved = []
            for item in results:
                _, _, saved = archive_email_item(
                    settings["base_dir"],
                    item,
                    current_company,
                    skip_inline_images=settings["skip_inline_images"],
                )
                all_saved.extend(saved)
            st.success(f"Scaricati e archiviati {len(results)} email. File salvati: {len(all_saved)}")
            st.code("\n".join(all_saved[:50]) or "Nessun file salvato", language="text")
        return

    if results:
        st.success(f"Trovate {len(results)} email coerenti con '{current_company}'.")
        selected_keys = []
        for idx, item in enumerate(results):
            key = f"{item.mailbox}|{item.uid}"
            with st.expander(f"{idx + 1}. {item.received_at.strftime('%d/%m/%Y')} - {item.subject} - {item.sender_email} - score {item.company_score:.0%}", expanded=(idx == 0)):
                st.write(f"**Mittente:** {item.sender_name} <{item.sender_email}>")
                st.write(f"**Data ricezione:** {item.received_at.strftime('%d/%m/%Y %H:%M')}")
                st.write(f"**Motivo abbinamento:** {item.company_reasons}")
                st.write("**Sintesi intelligente:**")
                st.info(item.smart_summary or summarize_email(item))
                st.write("**Anteprima contenuto mail:**")
                st.text_area("", value=truncate_text(item.body_text, 2500), height=160, key=f"preview_{idx}")
                if item.attachments:
                    rows = [
                        {"Nome allegato": a.filename, "Tipo": a.content_type, "KB": round(a.size / 1024, 1), "Inline": a.is_inline}
                        for a in item.attachments
                    ]
                    st.write("**Allegati:**")
                    st.dataframe(pd.DataFrame(rows) if pd is not None else rows, use_container_width=True)
                else:
                    st.caption("Nessun allegato.")
                if st.checkbox("Seleziona questa email per SCARICA", key=f"select_{idx}"):
                    selected_keys.append(key)
        st.divider()
        if st.button("SCARICA EMAIL SELEZIONATE", type="primary"):
            selected_items = [x for x in results if f"{x.mailbox}|{x.uid}" in selected_keys]
            if not selected_items:
                st.warning("Nessuna email selezionata.")
            else:
                all_saved = []
                for item in selected_items:
                    _, _, saved = archive_email_item(
                        settings["base_dir"],
                        item,
                        current_company,
                        skip_inline_images=settings["skip_inline_images"],
                    )
                    all_saved.extend(saved)
                st.success(f"Archiviazione completata: {len(selected_items)} email, {len(all_saved)} file.")
                st.code("\n".join(all_saved), language="text")
    elif btn_view:
        st.warning("Nessuna email trovata con la soglia impostata. Abbassa la soglia o amplia le mailbox.")


def tab_importa_documenti(settings: Dict[str, Any]) -> None:
    st.subheader("📥 Importa e cataloga documenti")
    base_dir = settings["base_dir"]
    forced_company = st.text_input("Cliente/Azienda forzata opzionale", placeholder="Lascia vuoto per riconoscimento automatico")
    st.markdown("### Carica file singoli")
    uploads = st.file_uploader("PDF, immagini, TXT", accept_multiple_files=True, type=["pdf", "png", "jpg", "jpeg", "txt", "csv", "xml"])
    if uploads and st.button("Archivia file caricati"):
        temp_dir = Path(tempfile.mkdtemp(prefix="fp_upload_"))
        messages = []
        for up in uploads:
            path = temp_dir / safe_filename(up.name)
            path.write_bytes(up.getbuffer())
            ok, msg = archive_local_document(base_dir, path, forced_company=forced_company.strip())
            messages.append((ok, msg))
        for ok, msg in messages:
            st.success(msg) if ok else st.warning(msg)

    st.markdown("### Seleziona cartella locale con sottocartelle")
    folder = st.text_input("Percorso cartella da scansionare", placeholder=r"C:\Users\Danilo\Desktop\Documenti")
    recursive = st.checkbox("Leggi anche sottocartelle", value=True)
    if st.button("Scansiona e archivia cartella"):
        p = Path(folder).expanduser()
        if not p.exists() or not p.is_dir():
            st.error("Cartella non trovata.")
        else:
            exts = {".pdf", ".png", ".jpg", ".jpeg", ".txt", ".csv", ".xml"}
            files = [x for x in (p.rglob("*") if recursive else p.glob("*")) if x.is_file() and x.suffix.lower() in exts]
            progress = st.progress(0)
            results = []
            for i, f in enumerate(files, start=1):
                ok, msg = archive_local_document(base_dir, f, forced_company=forced_company.strip())
                results.append({"file": str(f), "ok": ok, "messaggio": msg})
                progress.progress(i / max(len(files), 1))
            st.success(f"Elaborati {len(files)} file.")
            st.dataframe(pd.DataFrame(results) if pd is not None else results, use_container_width=True)


def tab_ricerca_archivio(settings: Dict[str, Any]) -> None:
    st.subheader("🔍 Ricerca full-text archivio")
    q = st.text_input("Cerca per azienda, P.IVA, oggetto, banca, categoria, testo, allegato")
    if st.button("Cerca nell'archivio"):
        with connect_db(settings["base_dir"]) as conn:
            like = f"%{q}%"
            rows = conn.execute(
                """
                SELECT id, ragione_sociale, categoria, nome_archiviato, percorso, fonte, mittente, data_importazione
                FROM documenti
                WHERE ragione_sociale LIKE ? OR categoria LIKE ? OR nome_archiviato LIKE ? OR testo_estratto LIKE ? OR mittente LIKE ?
                ORDER BY id DESC LIMIT 300
                """,
                (like, like, like, like, like),
            ).fetchall()
        if rows:
            df = pd.DataFrame([dict(r) for r in rows]) if pd is not None else [dict(r) for r in rows]
            st.dataframe(df, use_container_width=True)
        else:
            st.info("Nessun risultato.")


def tab_coda(settings: Dict[str, Any]) -> None:
    st.subheader("🟠 Coda documenti da verificare")
    with connect_db(settings["base_dir"]) as conn:
        rows = conn.execute("SELECT * FROM coda_verificare ORDER BY id DESC LIMIT 200").fetchall()
    if not rows:
        st.success("Nessun documento da verificare.")
        return
    for row in rows:
        with st.expander(f"#{row['id']} - {Path(row['file_path']).name} - {row['stato']}"):
            st.write(f"**File:** {row['file_path']}")
            st.write(f"**Categoria suggerita:** {row['categoria_suggerita']}")
            st.write(f"**Cliente suggerito:** {row['cliente_suggerito']}")
            st.write(f"**Motivo:** {row['motivo_incertezza']}")
            st.text_area("Testo estratto", value=truncate_text(row["testo_estratto"], 3000), height=180, key=f"qtxt_{row['id']}")
            new_company = st.text_input("Cliente corretto", value=row["cliente_suggerito"] or "", key=f"qc_{row['id']}")
            new_category = st.text_input("Categoria corretta", value=row["categoria_suggerita"] or "Documento_Generico", key=f"qcat_{row['id']}")
            if st.button("Conferma e archivia", key=f"qbtn_{row['id']}"):
                src = Path(row["file_path"])
                if src.exists():
                    cliente_id = find_or_create_client(settings["base_dir"], new_company)
                    target_dir = settings["base_dir"] / DOCUMENT_ARCHIVE_DIR / company_folder_name(new_company)
                    target_dir.mkdir(parents=True, exist_ok=True)
                    target = unique_path(target_dir / safe_filename(src.name))
                    shutil.move(str(src), str(target))
                    file_hash = sha256_file(target)
                    insert_document_record(
                        settings["base_dir"], cliente_id, new_company, new_category, src.name, target.name,
                        str(target), file_hash, row["testo_estratto"], fonte="coda_verificare"
                    )
                    with connect_db(settings["base_dir"]) as conn:
                        conn.execute("UPDATE coda_verificare SET stato = 'Archiviato' WHERE id = ?", (row["id"],))
                        conn.execute(
                            "INSERT INTO apprendimento (keyword, categoria_corretta, cliente_corretto, regola_appresa, fonte) VALUES (?, ?, ?, ?, ?)",
                            (new_company, new_category, new_company, "Correzione coda verificare", str(target)),
                        )
                        conn.commit()
                    st.success("Documento archiviato e regola appresa.")
                else:
                    st.error("File non trovato.")


def tab_report_backup(settings: Dict[str, Any]) -> None:
    st.subheader("📊 Report e backup")
    base_dir = settings["base_dir"]
    if st.button("Genera report CSV documenti"):
        out = base_dir / REPORT_DIR / f"report_documenti_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        with connect_db(base_dir) as conn, open(out, "w", newline="", encoding="utf-8") as f:
            rows = conn.execute("SELECT * FROM documenti ORDER BY id DESC").fetchall()
            if rows:
                writer = csv.DictWriter(f, fieldnames=list(dict(rows[0]).keys()))
                writer.writeheader()
                for r in rows:
                    writer.writerow(dict(r))
        st.success(f"Report creato: {out}")
    if st.button("Backup database"):
        src = db_path(base_dir)
        dst = base_dir / BACKUP_DIR / f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{DB_NAME}"
        shutil.copy2(src, dst)
        st.success(f"Backup creato: {dst}")
    st.info("Per backup completo, copia tutta la cartella archivio principale indicata nella sidebar.")


def main() -> None:
    if st is None:
        raise RuntimeError("Streamlit non installato. Installa con: pip install streamlit")
    render_header()
    settings = sidebar_settings()
    tabs = st.tabs([
        "Dashboard",
        "CERCA AZIENDA",
        "Importa documenti",
        "Ricerca archivio",
        "Coda da verificare",
        "Report/Backup",
    ])
    with tabs[0]:
        dashboard(settings)
    with tabs[1]:
        tab_cerca_azienda(settings)
    with tabs[2]:
        tab_importa_documenti(settings)
    with tabs[3]:
        tab_ricerca_archivio(settings)
    with tabs[4]:
        tab_coda(settings)
    with tabs[5]:
        tab_report_backup(settings)


if __name__ == "__main__":
    main()
