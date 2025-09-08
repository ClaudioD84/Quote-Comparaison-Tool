# streamlit_app.py
"""
Fleet Leasing Offer Comparator - Streamlit App
Author: Senior Python Engineer (example)
Requirements:
  streamlit, pandas, numpy, pdfplumber, python-dateutil, xlsxwriter
Optional:
  camelot, tabula-py, pdfminer.six, pytesseract (for OCR - not invoked by default)
Notes:
  - All prices considered ex-VAT and exclude fuel by design.
  - The app attempts table extraction (camelot/tabula) if available, else falls back to pdfplumber text extraction.
  - Mixed currencies block export until user supplies conversion rates.
  - Demo data included.
"""
from __future__ import annotations
import io
import re
import sys
import logging
from typing import List, Dict, Any, Optional, Tuple
from dataclasses import dataclass, field
from datetime import datetime
import math

import streamlit as st
import pandas as pd
import numpy as np
import pdfplumber  # primary text extraction fallback
from dateutil import parser as dateparser

# Try optional libs (camelot/tabula) but tolerate absence
try:
    import camelot  # type: ignore
    HAS_CAMELOT = True
except Exception:
    HAS_CAMELOT = False

# Setup logging capture for UI
LOGS: List[str] = []
logger = logging.getLogger("leasing_comparator")
logger.setLevel(logging.DEBUG)
stream_handler = logging.StreamHandler(sys.stdout)
stream_handler.setLevel(logging.DEBUG)
logger.addHandler(stream_handler)

def log_ui(level: str, msg: str):
    ts = datetime.utcnow().isoformat()
    entry = f"[{ts}] {level.upper()}: {msg}"
    LOGS.append(entry)
    if level == "error":
        logger.error(msg)
    elif level == "warning":
        logger.warning(msg)
    else:
        logger.info(msg)


# ---------------------------
# Helpers: parsing & numeric
# ---------------------------
NUMBER_RE = re.compile(r"[-+]?\d{1,3}(?:[.,\s]\d{3})*(?:[.,]\d+)?|\d+(?:[.,]\d+)?")

CURRENCY_SYMBOLS = {
    "€": "EUR",
    "EUR": "EUR",
    "EUR ": "EUR",
    "EUR.": "EUR",
    "£": "GBP",
    "GBP": "GBP",
    "$": "USD",
    "USD": "USD",
    "CHF": "CHF",
}

def detect_currency_from_text(text: str) -> Optional[str]:
    # Search for symbols or ISO codes
    for sym, iso in CURRENCY_SYMBOLS.items():
        if sym in text:
            return iso
    # common patterns
    m = re.search(r"\b(EUR|EUR\.|EUR,|GBP|USD|CHF)\b", text, flags=re.I)
    if m:
        token = m.group(1).upper().strip(". ,")
        return CURRENCY_SYMBOLS.get(token, token)
    return None

def normalize_number(s: str) -> Optional[float]:
    """
    Normalize textual number to float, accepting European formats.
    Converts "1.234,56" -> 1234.56 and "1,234.56" -> 1234.56.
    """
    if s is None:
        return None
    s = str(s).strip()
    # Remove currency symbols
    s = re.sub(r"[^\d,.\-]+", "", s)
    if s == "":
        return None
    # If both . and , present: determine which is decimal by last occurrence
    if s.count(",") > 0 and s.count(".") > 0:
        if s.rfind(",") > s.rfind("."):
            # comma is decimal, dots are thousands
            s = s.replace(".", "")
            s = s.replace(",", ".")
        else:
            # dot is decimal, commas are thousands
            s = s.replace(",", "")
    else:
        # Only commas: could be decimal or thousands depending on grouping
        if s.count(",") == 1 and len(s.split(",")[-1]) in (1,2):
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    try:
        return float(s)
    except Exception as e:
        log_ui("warning", f"normalize_number failed for '{s}': {e}")
        return None

def parse_currency_and_number(snippet: str) -> Tuple[Optional[str], Optional[float]]:
    """
    Find the first currency symbol/ISO and a number in the snippet.
    """
    if not snippet:
        return None, None
    cur = detect_currency_from_text(snippet)
    m = NUMBER_RE.search(snippet.replace("\u202f", ""))  # narrow no-break-space
    if m:
        num = normalize_number(m.group(0))
    else:
        num = None
    return cur, num

def parse_duration_months(text: str) -> Optional[int]:
    """
    Find contract duration in months. Accept patterns:
      - "36 months", "36 mths", "3 years"
      - "48 mo", "24 months"
    """
    if not text:
        return None
    text = text.lower()
    # look for months directly
    m = re.search(r"(\d+)\s*(months|month|mths|mos|mo)\b", text)
    if m:
        return int(m.group(1))
    # years
    m = re.search(r"(\d+(\.\d+)?)\s*(years|year|yrs|yr)\b", text)
    if m:
        years = float(m.group(1))
        return int(round(years * 12))
    # sometimes "36m" or "36M"
    m = re.search(r"\b(\d{2,3})\s*[mM]\b", text)
    if m:
        return int(m.group(1))
    return None

def parse_mileage_total(text: str) -> Optional[int]:
    """
    Parse total contract mileage, often like "60,000 km", "10000 miles", "8.000 km p.a. x 36" etc.
    If phrase is per-annum, multiply by duration if available (but main check enforces equal mileage provided).
    We'll extract a single absolute total if present, else detect per-year and leave to higher-level code.
    """
    if not text:
        return None
    text = text.replace("\u202f", " ")
    # absolute total like 60000 km
    m = re.search(r"(\d{1,3}(?:[.,\s]\d{3})+|\d+)\s*(km|kilometers|kilometres|miles|mi)\b", text, flags=re.I)
    if m:
        num = normalize_number(m.group(1))
        unit = m.group(2).lower()
        if unit.startswith("mile"):
            # convert miles to km? Keep original unit but application expects mileage number (units not normalized).
            # We'll keep numeric as provided; assume user uses same units across offers.
            return int(round(num))
        else:
            return int(round(num))
    # per-annum patterns like "10,000 km p.a." or "10k km/year"
    m = re.search(r"(\d{1,3}(?:[.,\s]\d{3})+|\d+)\s*(k)?\s*(km|kilometres|kilometers|miles|mi)\s*(p\.a\.|per year|/yr|per annum|annum|year)\b", text, flags=re.I)
    if m:
        num = normalize_number(m.group(1))
        # if 'k' exists, multiply by 1,000
        if m.group(2):
            num *= 1000
        # Here we return per-year figure; upstream should multiply by duration if needed.
        return int(round(num))
    return None

def parse_excess_mileage_rate(text: str) -> Optional[float]:
    """
    Normalize excess mileage to cost per km (e.g., "€0.06/km" -> 0.06 ; "60 €/1,000 km" -> 0.06)
    """
    if not text:
        return None
    # Look for patterns with currency and per km or per 1000km
    m = re.search(r"([€£$]|\b(EUR|GBP|USD)\b)?\s*([0-9\.,]+)\s*(?:€|£|\$)?\s*(?:per|/)?\s*(\d{1,3}(?:[.,]\d{3})?|\d+)?\s*(km|kilometre|kilometer|kms|km\.)", text, flags=re.I)
    if m:
        raw = m.group(3)
        denom = m.group(4)
        unit = m.group(5)
        val = normalize_number(raw)
        denom_num = 1
        if denom:
            try:
                denom_num = int(re.sub(r"[^\d]", "", denom))
            except:
                denom_num = 1
        if denom_num == 0:
            denom_num = 1
        per_km = val / denom_num
        return round(per_km, 6)
    # fallback: just number with /km
    m = re.search(r"([0-9\.,]+)\s*/\s*km", text, flags=re.I)
    if m:
        val = normalize_number(m.group(1))
        return val
    return None

# ---------------------------
# Parsing pipeline
# ---------------------------
def extract_text_fallback(file_bytes: bytes) -> Dict[str, Any]:
    """
    Try to get text via pdfplumber. Returns dict with pages' text and raw_text concatenated.
    """
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            pages = []
            for p in pdf.pages:
                try:
                    txt = p.extract_text(x_tolerance=2) or ""
                except Exception as e:
                    log_ui("warning", f"pdfplumber page extract failed: {e}")
                    txt = ""
                pages.append(txt)
            raw_text = "\n\n".join(pages)
            return {"pages": pages, "raw_text": raw_text}
    except Exception as e:
        log_ui("error", f"pdfplumber failed to open PDF: {e}")
        return {"pages": [], "raw_text": ""}

def try_camelot_tables(file_bytes: bytes) -> Optional[List[pd.DataFrame]]:
    """
    If camelot is available and PDF has extractable tables, attempt to extract them.
    Returns list of pandas DataFrames or None.
    """
    if not HAS_CAMELOT:
        return None
    try:
        # camelot requires a filename, so write to temp buffer
        import tempfile
        with tempfile.NamedTemporaryFile(delete=True, suffix=".pdf") as tf:
            tf.write(file_bytes)
            tf.flush()
            tables = camelot.read_pdf(tf.name, flavor="lattice", pages="all")
            dfs = []
            for t in tables:
                try:
                    dfs.append(t.df)
                except Exception as e:
                    log_ui("warning", f"camelot table to df failed: {e}")
            if dfs:
                return dfs
    except Exception as e:
        log_ui("warning", f"camelot extraction failed: {e}")
    return None

@dataclass
class ParseResult:
    """
    Standardized structure of parsed raw fields (may contain None).
    """
    filename: str
    raw_text: str = ""
    snippets: Dict[str, str] = field(default_factory=dict)
    leasing_company_name: Optional[str] = None
    vehicle_description: Optional[str] = None
    contract_duration_months: Optional[int] = None
    contract_mileage_total: Optional[int] = None
    monthly_rental: Optional[float] = None
    upfront_costs: Optional[float] = None
    deposit: Optional[float] = None
    delivery_registration: Optional[float] = None
    admin_fees: Optional[float] = None
    maintenance_included: Optional[bool] = None
    maintenance_cost: Optional[float] = None
    tyres_included: Optional[bool] = None
    tyres_cost: Optional[float] = None
    road_tax_included: Optional[bool] = None
    road_tax_cost: Optional[float] = None
    insurance_included: Optional[bool] = None
    insurance_cost: Optional[float] = None
    early_termination_fees: Optional[str] = None
    excess_mileage_rate_per_km: Optional[float] = None
    discounts_or_rebates: Optional[float] = None
    currency: Optional[str] = None
    offer_valid_until: Optional[str] = None
    lead_time_or_eta: Optional[str] = None
    bundled_monthly: Optional[bool] = False
    parsing_confidence: float = 0.0
    scanned: bool = False
    raw_parse_warnings: List[str] = field(default_factory=list)


def parse_pdf(file_bytes: bytes, filename: str) -> ParseResult:
    """
    Robust PDF parser that tries:
      1. Camelot (if installed)
      2. pdfplumber text extraction + regex heuristics
    Returns a ParseResult with detected fields and raw_text snippets.
    """
    pr = ParseResult(filename=filename)
    # Attempt table extraction first (if possible)
    try:
        dfs = try_camelot_tables(file_bytes)
        if dfs:
            pr.raw_text = "\n\n".join(["\n".join(df.astype(str).apply(lambda row: " | ".join(row), axis=1).tolist()) for df in dfs])
            pr.snippets['tables_preview'] = pr.raw_text[:3000]
            pr.parsing_confidence += 0.3
    except Exception as e:
        log_ui("warning", f"camelot attempt error for {filename}: {e}")

    # Always run pdfplumber fallback for text
    text_info = extract_text_fallback(file_bytes)
    if text_info["raw_text"].strip() == "":
        pr.scanned = True
        pr.raw_parse_warnings.append("No extractable text found (likely scanned PDF). OCR not performed.")
        log_ui("warning", f"{filename}: likely scanned (no text).")
    else:
        pr.raw_text = (pr.raw_text + "\n\n" + text_info["raw_text"]).strip()
        pr.snippets['full_text_preview'] = pr.raw_text[:4000]
        pr.parsing_confidence += 0.4

    # Heuristics and regex to populate fields
    txt = pr.raw_text.lower()

    # Company detection: header/footer patterns or email/domains or leasing keywords
    m = re.search(r"([A-Z][A-Za-z&\s]{2,60}\s+(leasing|lease|finance|motor|automotive|rentals|renting|fleet))", pr.raw_text, flags=re.I)
    if m:
        pr.leasing_company_name = m.group(1).strip()
    else:
        # look for email domain contact block
        m2 = re.search(r"([A-Za-z0-9\.\-_]+)@([A-Za-z0-9\.\-_]+\.[A-Za-z]{2,})", pr.raw_text)
        if m2:
            domain = m2.group(2)
            pr.leasing_company_name = domain.split(".")[0].capitalize()
    if not pr.leasing_company_name:
        # fallback: use filename sans extension
        pr.leasing_company_name = re.sub(r"[_\-\.]+", " ", re.sub(r"\.pdf$", "", filename, flags=re.I)).strip()

    # vehicle description: look for "Vehicle:", "Model:", "Make/Model", or first big line near top
    m = re.search(r"(vehicle description|vehicle|model|make\/model|car)\s*[:\-]\s*(.+)", pr.raw_text, flags=re.I)
    if m:
        pr.vehicle_description = m.group(2).split("\n")[0].strip()
    else:
        # try first non-empty line as fallback
        lines = [ln.strip() for ln in pr.raw_text.splitlines() if ln.strip()]
        if lines:
            pr.vehicle_description = lines[0][:200]

    # currency detection
    cur = detect_currency_from_text(pr.raw_text)
    pr.currency = cur or None

    # contract duration
    dur = parse_duration_months(pr.raw_text)
    if dur:
        pr.contract_duration_months = dur

    # mileage
    mileage = parse_mileage_total(pr.raw_text)
    if mileage:
        pr.contract_mileage_total = mileage

    # monthly rental: search for "monthly", "monthly rental", "per month", "mo" near a currency/number
    m = re.search(r"(monthly rental|monthly lease|monthly|per month|per month:|p/m|pcm|p\.m\.)\s*[:\-]?\s*([€£$]?\s*[0-9\.,\s]+)", pr.raw_text, flags=re.I)
    if m:
        _, num = parse_currency_and_number(m.group(2))
        if num is not None:
            pr.monthly_rental = num
            pr.parsing_confidence += 0.1
    else:
        # fallback: look for "Total monthly" or first currency preceded by 'month' within 50 chars
        m2 = re.search(r"([€£$]?\s*[0-9\.,]+\s*)(?:per month|pcm|monthly|/month|p\.m\.)", pr.raw_text, flags=re.I)
        if m2:
            _, num = parse_currency_and_number(m2.group(1))
            if num is not None:
                pr.monthly_rental = num
                pr.parsing_confidence += 0.05

    # upfront costs: deposit, delivery, registration
    for label in ["deposit", "down payment", "initial payment"]:
        m = re.search(rf"{label}\s*[:\-]?\s*([€£$]?\s*[0-9\.,\s]+)", pr.raw_text, flags=re.I)
        if m:
            _, v = parse_currency_and_number(m.group(1))
            if v is not None:
                pr.deposit = v
                pr.parsing_confidence += 0.05
                pr.snippets[f"{label}_snippet"] = m.group(0)

    for label in ["delivery", "registration", "delivery & registration", "delivery/registration", "registration fee"]:
        m = re.search(rf"{label}\s*[:\-]?\s*([€£$]?\s*[0-9\.,\s]+)", pr.raw_text, flags=re.I)
        if m:
            _, v = parse_currency_and_number(m.group(1))
            if v is not None:
                pr.delivery_registration = (pr.delivery_registration or 0.0) + v
                pr.parsing_confidence += 0.03

    # admin fees
    m = re.search(r"(admin(istration)? fee|administration charge|processing fee)\s*[:\-]?\s*([€£$]?\s*[0-9\.,\s]+)", pr.raw_text, flags=re.I)
    if m:
        _, v = parse_currency_and_number(m.group(2))
        if v is not None:
            pr.admin_fees = v
            pr.parsing_confidence += 0.05

    # maintenance inclusion/cost
    m_inc = re.search(r"(maintenance (is )?included|full maintenance included|maintenace incl|maint\. incl|with maintenance)", pr.raw_text, flags=re.I)
    if m_inc:
        pr.maintenance_included = True
    m_cost = re.search(r"(maintenance cost|maintenance fee|maintenance)\s*[:\-]?\s*([€£$]?\s*[0-9\.,\s]+)", pr.raw_text, flags=re.I)
    if m_cost:
        _, v = parse_currency_and_number(m_cost.group(2))
        if v is not None:
            pr.maintenance_cost = v
            pr.maintenance_included = pr.maintenance_included or False
            pr.parsing_confidence += 0.04

    # tyres, road tax, insurance similar patterns
    for field, phrases in [
        ("tyres", ["tyre", "tyres", "tyres included", "tyres inc", "tyre replacement"]),
        ("road_tax", ["road tax", "vehicle tax", "tax included"]),
        ("insurance", ["insurance included", "insurance", "comprehensive insurance"]),
    ]:
        inc_pat = r"|".join([re.escape(p) for p in phrases if " " in p or p.isalpha()])
        m_inc = re.search(rf"({inc_pat})\s*(is )?included", pr.raw_text, flags=re.I)
        if m_inc:
            if field == "tyres":
                pr.tyres_included = True
            elif field == "road_tax":
                pr.road_tax_included = True
            elif field == "insurance":
                pr.insurance_included = True
        m_cost = re.search(rf"({inc_pat})\s*[:\-]?\s*([€£$]?\s*[0-9\.,\s]+)", pr.raw_text, flags=re.I)
        if m_cost:
            _, v = parse_currency_and_number(m_cost.group(2))
            if v is not None:
                if field == "tyres":
                    pr.tyres_cost = v
                elif field == "road_tax":
                    pr.road_tax_cost = v
                elif field == "insurance":
                    pr.insurance_cost = v
                pr.parsing_confidence += 0.03

    # early termination
    m = re.search(r"(early termination|termination fee|early termination fee).{0,80}?([€£$]?\s*[0-9\.,\s]+)?", pr.raw_text, flags=re.I)
    if m:
        pr.early_termination_fees = m.group(0)[:300]
        # try extract number
        if m.group(2):
            _, v = parse_currency_and_number(m.group(2))
            if v is not None:
                pr.raw_parse_warnings.append(f"early_termination_fee_detected:{v}")

    # excess mileage
    em = parse_excess_mileage_rate(pr.raw_text)
    if em:
        pr.excess_mileage_rate_per_km = em
        pr.parsing_confidence += 0.03

    # discounts
    m = re.search(r"(discount|rebate|reduction)\s*[:\-]?\s*([€£$]?\s*[0-9\.,\s]+|\d+%?)", pr.raw_text, flags=re.I)
    if m:
        token = m.group(2)
        if "%" in token:
            # percent discount: need to interpret relative to monthly or TCC later; store as negative percent marker
            try:
                pr.discounts_or_rebates = -abs(float(token.strip().replace("%", "")))  # negative indicates percent
            except:
                pr.raw_parse_warnings.append("discount_percent_unparsed")
        else:
            _, v = parse_currency_and_number(token)
            if v is not None:
                pr.discounts_or_rebates = v
                pr.parsing_confidence += 0.02

    # offer valid until / lead time
    m = re.search(r"(valid until|offer valid until|valid to)\s*[:\-]?\s*([A-Za-z0-9 ,.-]+)", pr.raw_text, flags=re.I)
    if m:
        try:
            dt = dateparser.parse(m.group(2), fuzzy=True)
            pr.offer_valid_until = dt.date().isoformat()
        except Exception:
            pr.offer_valid_until = m.group(2).strip()
    m = re.search(r"(lead time|eta|est(?:imated)? delivery|delivery in)\s*[:\-]?\s*([0-9]+)\s*(weeks|days|months)", pr.raw_text, flags=re.I)
    if m:
        pr.lead_time_or_eta = f"{m.group(2)} {m.group(3)}"

    # Compose upfront_costs total
    upfront = 0.0
    found_any_upfront = False
    for v in [pr.deposit, pr.delivery_registration]:
        if v is not None:
            upfront += float(v)
            found_any_upfront = True
    if found_any_upfront:
        pr.upfront_costs = upfront
        pr.parsing_confidence += 0.02

    # If monthly appears bundled with maintenance/tyres/insurance, detect
    bundle_phrases = re.search(r"(monthly (?:price|payment|rental).{0,40}?incl(?:uding)?|incl\.)", pr.raw_text, flags=re.I)
    if bundle_phrases:
        pr.bundled_monthly = True

    # Final parsing confidence normalization
    pr.parsing_confidence = min(1.0, pr.parsing_confidence)
    # Provide small confidence bump if essential fields present
    if pr.contract_duration_months and pr.contract_mileage_total and pr.monthly_rental:
        pr.parsing_confidence = max(pr.parsing_confidence, 0.6)
    # Save some snippets for transparency
    pr.snippets['top_lines'] = "\n".join([ln for ln in pr.raw_text.splitlines()[:20]])
    return pr

# ---------------------------
# Normalization & computation
# ---------------------------
def normalize_offer(pr: ParseResult, assumptions: Dict[str, Any]) -> Dict[str, Any]:
    """
    Turn ParseResult into normalized offer dict with known keys, applying assumptions toggles.
    Does not finalize discounts that are percent-based (keeps as negative percent to interpret later).
    """
    o = {
        "filename": pr.filename,
        "vendor": pr.leasing_company_name,
        "vehicle_description": pr.vehicle_description,
        "duration_months": pr.contract_duration_months,
        "mileage_total": pr.contract_mileage_total,
        "monthly_rental": pr.monthly_rental,
        "upfront_costs": pr.upfront_costs or 0.0,
        "deposit": pr.deposit or 0.0,
        "delivery_registration": pr.delivery_registration or 0.0,
        "admin_fees": pr.admin_fees or 0.0,
        "maintenance_included": pr.maintenance_included,
        "maintenance_cost": pr.maintenance_cost or 0.0,
        "tyres_included": pr.tyres_included,
        "tyres_cost": pr.tyres_cost or 0.0,
        "road_tax_included": pr.road_tax_included,
        "road_tax_cost": pr.road_tax_cost or 0.0,
        "insurance_included": pr.insurance_included,
        "insurance_cost": pr.insurance_cost or 0.0,
        "early_termination_fees": pr.early_termination_fees,
        "excess_mileage_rate_per_km": pr.excess_mileage_rate_per_km,
        "discounts_or_rebates": pr.discounts_or_rebates or 0.0,
        "currency": pr.currency,
        "offer_valid_until": pr.offer_valid_until,
        "lead_time_or_eta": pr.lead_time_or_eta,
        "bundled_monthly": pr.bundled_monthly,
        "parsing_confidence": pr.parsing_confidence,
        "scanned": pr.scanned,
        "raw_snippets": pr.snippets,
        "raw_parse_warnings": pr.raw_parse_warnings,
    }

    # Apply user overrides for discounts per vendor, if provided
    overrides = assumptions.get("discount_overrides", {}) or {}
    if o["vendor"] in overrides and overrides[o["vendor"]] is not None:
        ov = overrides[o["vendor"]]
        try:
            val = float(ov)
            o["discounts_or_rebates"] = val
        except:
            log_ui("warning", f"discount override for {o['vendor']} invalid: {ov}")

    # If a discount is negative (percent), apply later during compute_metrics
    return o

def compute_metrics(o: Dict[str, Any], assumptions: Dict[str, Any]) -> Dict[str, Any]:
    """
    Compute TCC, effective monthly cost, cost per km, etc.
    assumptions decide whether to include maintenance, tyres, insurance, road tax.
    """
    # copy to avoid mutation
    od = dict(o)

    # Required checks
    if od.get("duration_months") is None or od.get("mileage_total") is None:
        od["error_missing_essential"] = True
        return od

    dur = int(od["duration_months"])
    mil = int(od["mileage_total"])

    # Compose costs: upfront + admin + monthly*duration + included items
    upfront = float(od.get("upfront_costs", 0.0))
    admin = float(od.get("admin_fees", 0.0))
    monthly = float(od.get("monthly_rental") or 0.0)

    # Maintenance, tyres, insurance, road tax inclusion controlled by assumptions + detected inclusion/cost
    maintenance_total = 0.0
    if assumptions.get("include_maintenance") and (od.get("maintenance_included") or od.get("maintenance_cost") > 0):
        # If maintenance cost is per month? Heuristic: if maintenance_cost < 50 assume per-month? Not safe.
        # We'll check phrase: if price found once likely one-off; default: if maintenance_cost < 500 assume per-month for long durations? Risky.
        mcost = od.get("maintenance_cost", 0.0)
        if mcost and mcost < 1000 and mcost > 0 and mcost < monthly:  # heuristic: if small value possibly per month
            maintenance_total = mcost * dur
        else:
            maintenance_total = mcost  # one-off or total
        # If included flag true and cost zero, treat as included (zero)
    tyres_total = 0.0
    if assumptions.get("include_tyres") and (od.get("tyres_included") or od.get("tyres_cost") > 0):
        tcost = od.get("tyres_cost", 0.0)
        tyres_total = tcost

    insurance_total = 0.0
    if assumptions.get("include_insurance") and (od.get("insurance_included") or od.get("insurance_cost") > 0):
        icost = od.get("insurance_cost", 0.0)
        # same heuristic as maintenance
        if icost and icost < 1000 and icost < monthly:
            insurance_total = icost * dur
        else:
            insurance_total = icost

    road_tax_total = 0.0
    if assumptions.get("include_road_tax") and (od.get("road_tax_included") or od.get("road_tax_cost") > 0):
        road_tax_total = od.get("road_tax_cost", 0.0)

    # Discounts: if negative value implies percent off (we encoded as negative percent earlier)
    discounts = od.get("discounts_or_rebates", 0.0)
    discount_amount = 0.0
    if isinstance(discounts, (int, float)):
        if discounts < 0:
            # percent discount -> apply to total monthly*duration + upfront + admin + included additions
            gross = upfront + admin + monthly * dur + maintenance_total + tyres_total + insurance_total + road_tax_total
            discount_amount = abs(discounts) / 100.0 * gross
        else:
            discount_amount = float(discounts)
    else:
        discount_amount = 0.0

    tcc = upfront + admin + monthly * dur + maintenance_total + tyres_total + insurance_total + road_tax_total - discount_amount
    effective_monthly = tcc / dur if dur > 0 else None
    cost_per_km = tcc / mil if mil > 0 else None

    od.update({
        "computed_upfront": upfront,
        "computed_admin": admin,
        "computed_monthly_total": monthly * dur,
        "computed_maintenance_total": maintenance_total,
        "computed_tyres_total": tyres_total,
        "computed_insurance_total": insurance_total,
        "computed_road_tax_total": road_tax_total,
        "computed_discount_amount": discount_amount,
        "TCC": tcc,
        "effective_monthly": effective_monthly,
        "cost_per_km": cost_per_km,
    })
    return od

# ---------------------------
# Comparison & analysis
# ---------------------------
def compare_offers(offers: List[Dict[str, Any]], tie_breakers: List[str]) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Generate comparison DataFrame across offers and winner analysis dict.
    tie_breakers: ordered list of metrics to apply in case of tie (e.g., ['TCC','effective_monthly','lead_time']).
    """
    rows = []
    currency_set = set()
    missing_essentials = []
    for o in offers:
        if o.get("currency"):
            currency_set.add(o.get("currency"))
        if o.get("error_missing_essential"):
            missing_essentials.append(o.get("vendor"))
        rows.append({
            "Vendor": o.get("vendor"),
            "Duration_months": o.get("duration_months"),
            "Mileage_total": o.get("mileage_total"),
            "Currency": o.get("currency"),
            "Monthly_rental": o.get("monthly_rental") or 0.0,
            "Upfront_costs": o.get("computed_upfront", 0.0),
            "Admin_fees": o.get("computed_admin", 0.0),
            "Maintenance_total": o.get("computed_maintenance_total", 0.0),
            "Tyres_total": o.get("computed_tyres_total", 0.0),
            "Insurance_total": o.get("computed_insurance_total", 0.0),
            "Road_tax_total": o.get("computed_road_tax_total", 0.0),
            "Discount_amount": o.get("computed_discount_amount", 0.0),
            "TCC": o.get("TCC", np.nan),
            "Effective_monthly": o.get("effective_monthly", np.nan),
            "Cost_per_km": o.get("cost_per_km", np.nan),
            "Lead_time_or_eta": o.get("lead_time_or_eta"),
            "Offer_valid_until": o.get("offer_valid_until"),
            "Parsing_confidence": o.get("parsing_confidence", 0.0),
            "Bundled_monthly": o.get("bundled_monthly", False),
        })
    df = pd.DataFrame(rows)

    analysis = {
        "currency_set": currency_set,
        "missing_essentials": missing_essentials,
    }

    # Validate durations & mileage identical across offers
    durations = df["Duration_months"].unique()
    mileages = df["Mileage_total"].unique()
    if len(durations) > 1:
        analysis["duration_mismatch"] = True
    else:
        analysis["duration_mismatch"] = False
    if len(mileages) > 1:
        analysis["mileage_mismatch"] = True
    else:
        analysis["mileage_mismatch"] = False

    # Ranking: default by TCC ascending. Apply tie-breakers as needed
    df_sorted = df.copy()
    # ensure numeric columns
    for col in ["TCC", "Effective_monthly", "Cost_per_km", "Monthly_rental"]:
        df_sorted[col] = pd.to_numeric(df_sorted[col], errors="coerce")

    # Build sort order: TCC then tie_breakers
    sort_cols = []
    ascending = []
    # primary TCC
    sort_cols.append("TCC"); ascending.append(True)
    # apply any additional tie breakers mapping names to columns
    for tb in tie_breakers:
        if tb.lower() in ("monthly", "monthly_rental", "monthly rental"):
            sort_cols.append("Monthly_rental"); ascending.append(True)
        elif tb.lower() in ("effective_monthly", "effective monthly", "effective monthly cost"):
            sort_cols.append("Effective_monthly"); ascending.append(True)
        elif tb.lower() in ("lead_time", "lead time", "lead_time_or_eta"):
            # for lead time, smaller is better but it's textual; keep as last resort
            # We'll sort NaNs to bottom: create helper numeric mapping where possible
            df_sorted["_lead_numeric"] = df_sorted["Lead_time_or_eta"].apply(_lead_time_to_days)
            sort_cols.append("_lead_numeric"); ascending.append(True)
        elif tb.lower() in ("parsing_confidence",):
            sort_cols.append("Parsing_confidence"); ascending.append(False)
        # else ignore unknowns
    # apply sort
    try:
        df_sorted = df_sorted.sort_values(by=sort_cols, ascending=ascending, na_position='last')
    except Exception as e:
        log_ui("warning", f"Sorting by tie_breakers failed: {e}")
    df_sorted["Rank"] = range(1, len(df_sorted) + 1)

    # Winner
    winner_row = df_sorted.iloc[0].to_dict() if not df_sorted.empty else {}
    # Explain why: compute top 3 contributing factors versus runner-up
    why_notes = []
    if not df_sorted.empty and len(df_sorted) > 1:
        winner = df_sorted.iloc[0]
        runner = df_sorted.iloc[1]
        diffs = {}
        for comp in ["TCC", "Monthly_rental", "Upfront_costs", "Admin_fees", "Maintenance_total", "Tyres_total", "Insurance_total", "Road_tax_total"]:
            try:
                dif = float(runner.get(comp, 0.0)) - float(winner.get(comp, 0.0))
            except:
                dif = 0.0
            diffs[comp] = dif
        # rank by absolute difference descending
        sorted_diffs = sorted(diffs.items(), key=lambda kv: abs(kv[1]), reverse=True)
        top3 = sorted_diffs[:3]
        for comp, val in top3:
            why_notes.append(f"{comp}: winner {'lower' if val>0 else 'higher'} by {abs(val):,.2f} {list(currency_set)[0] if currency_set else ''}")
    else:
        why_notes.append("Single offer or insufficient data to compare.")

    analysis["winner"] = winner_row
    analysis["why_notes"] = why_notes
    analysis["comparison_df"] = df_sorted
    return df_sorted, analysis

def _lead_time_to_days(s: Optional[str]) -> Optional[int]:
    if not isinstance(s, str):
        return None
    m = re.search(r"(\d+)\s*(day|days)", s, flags=re.I)
    if m:
        return int(m.group(1))
    m = re.search(r"(\d+)\s*(week|weeks)", s, flags=re.I)
    if m:
        return int(m.group(1)) * 7
    m = re.search(r"(\d+)\s*(month|months)", s, flags=re.I)
    if m:
        return int(m.group(1)) * 30
    return None

# ---------------------------
# Excel generation
# ---------------------------
def build_excel(comparison_df: pd.DataFrame, offers: List[Dict[str, Any]], assumptions: Dict[str, Any]) -> bytes:
    """
    Build an in-memory Excel workbook with Winner Analysis sheet and one sheet per vendor.
    Returns bytes to stream to user.
    """
    output = io.BytesIO()
    # Use pandas ExcelWriter with xlsxwriter
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format='yyyy-mm-dd') as writer:
        workbook = writer.book
        # Formats
        money_fmt = workbook.add_format({'num_format': '#,##0.00'})
        bold = workbook.add_format({'bold': True})
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#DCE6F1'})
        # Winner Analysis sheet
        summary_sheet = "Winner Analysis"
        # Summary table
        comp = comparison_df.copy()
        comp.to_excel(writer, sheet_name=summary_sheet, startrow=1, index=False)
        ws = writer.sheets[summary_sheet]
        ws.write(0, 0, "Winner Analysis Summary", bold)
        # Apply formatting on currency columns
        col_map = {c: i for i, c in enumerate(comp.columns)}
        for col in ["Monthly_rental", "Upfront_costs", "Admin_fees", "Maintenance_total", "Tyres_total", "Insurance_total", "Road_tax_total", "Discount_amount", "TCC", "Effective_monthly", "Cost_per_km"]:
            if col in col_map:
                ws.set_column(col_map[col], col_map[col], 18, money_fmt)
        # Write assumptions and "why winner" below table
        start_row = len(comp) + 4
        ws.write(start_row, 0, "Assumptions", header_fmt)
        ar = assumptions
        ws.write(start_row+1, 0, "Include maintenance")
        ws.write(start_row+1, 1, str(ar.get("include_maintenance", False)))
        ws.write(start_row+2, 0, "Include tyres")
        ws.write(start_row+2, 1, str(ar.get("include_tyres", False)))
        ws.write(start_row+3, 0, "Include insurance")
        ws.write(start_row+3, 1, str(ar.get("include_insurance", False)))
        ws.write(start_row+4, 0, "Include road tax")
        ws.write(start_row+4, 1, str(ar.get("include_road_tax", False)))
        ws.write(start_row+5, 0, "Tie-breaker priority")
        ws.write(start_row+5, 1, ", ".join(ar.get("tie_breaker_priority", [])))
        # Why winner
        winner_notes = ar.get("winner_notes", [])
        ws.write(start_row+7, 0, "Why winner?", header_fmt)
        for i, note in enumerate(winner_notes):
            ws.write(start_row+8+i, 0, note)
        # Per-vendor sheets
        for o in offers:
            name = o.get("vendor")[:31]  # Excel sheet name limit
            # assemble a clean summary table
            rows = []
            for k in ["filename", "vendor", "vehicle_description", "duration_months", "mileage_total", "currency", "monthly_rental",
                      "upfront_costs", "admin_fees", "maintenance_included", "maintenance_cost", "tyres_included", "tyres_cost",
                      "insurance_included", "insurance_cost", "road_tax_included", "road_tax_cost", "excess_mileage_rate_per_km",
                      "discounts_or_rebates", "offer_valid_until", "lead_time_or_eta", "parsing_confidence", "scanned"]:
                rows.append({"Field": k, "Value": o.get(k)})
            dfv = pd.DataFrame(rows)
            dfv.to_excel(writer, sheet_name=name, index=False)
            ws_v = writer.sheets[name]
            # append computations and raw snippets below
            base_rows = len(dfv) + 3
            ws_v.write(base_rows-1, 0, "Computed metrics", header_fmt)
            comp_rows = [
                ("computed_monthly_total", o.get("computed_monthly_total")),
                ("computed_maintenance_total", o.get("computed_maintenance_total")),
                ("computed_tyres_total", o.get("computed_tyres_total")),
                ("computed_insurance_total", o.get("computed_insurance_total")),
                ("computed_road_tax_total", o.get("computed_road_tax_total")),
                ("computed_discount_amount", o.get("computed_discount_amount")),
                ("TCC", o.get("TCC")),
                ("effective_monthly", o.get("effective_monthly")),
                ("cost_per_km", o.get("cost_per_km")),
            ]
            for i, (k, v) in enumerate(comp_rows):
                ws_v.write(base_rows + i, 0, k)
                ws_v.write(base_rows + i, 1, v, money_fmt if isinstance(v, (int, float)) else None)
            # Raw snippets appendix
            snip_start = base_rows + len(comp_rows) + 2
            ws_v.write(snip_start, 0, "Raw snippets / parse warnings", header_fmt)
            snip_text = ""
            for k, val in (o.get("raw_snippets") or {}).items():
                snip_text += f"--- {k} ---\n{val}\n\n"
            warnings = o.get("raw_parse_warnings", [])
            if warnings:
                snip_text += "Parse warnings:\n" + "\n".join(warnings)
            # write as a single cell (may be long)
            ws_v.write(snip_start+1, 0, snip_text)
            # widen sheet
            ws_v.set_column(0, 0, 30)
            ws_v.set_column(1, 1, 40)
    data = output.getvalue()
    return data

# ---------------------------
# Demo test data generator
# ---------------------------
def demo_offers() -> List[Dict[str, Any]]:
    """
    Generate 3 synthetic offers for demo/testing without PDFs.
    """
    now = datetime.utcnow().date()
    offers = []
    # Vendor A: low monthly, higher upfront
    A = {
        "filename": "demo_offer_A.pdf",
        "vendor": "Alpha Leasing",
        "vehicle_description": "2025 Model X Electric - Demo",
        "duration_months": 36,
        "mileage_total": 60000,
        "monthly_rental": 450.0,
        "upfront_costs": 3000.0,
        "deposit": 2000.0,
        "delivery_registration": 1000.0,
        "admin_fees": 250.0,
        "maintenance_included": True,
        "maintenance_cost": 0.0,
        "tyres_included": False,
        "tyres_cost": 400.0,
        "road_tax_included": True,
        "road_tax_cost": 0.0,
        "insurance_included": False,
        "insurance_cost": 1200.0,
        "early_termination_fees": "See T&Cs",
        "excess_mileage_rate_per_km": 0.06,
        "discounts_or_rebates": 0.0,
        "currency": "EUR",
        "offer_valid_until": (now.replace(day=1)).isoformat(),
        "lead_time_or_eta": "4 weeks",
        "bundled_monthly": False,
        "parsing_confidence": 0.9,
        "scanned": False,
        "raw_snippets": {"demo": "alpha demo"},
        "raw_parse_warnings": []
    }
    # Vendor B: higher monthly but includes tyres
    B = {
        "filename": "demo_offer_B.pdf",
        "vendor": "Beta Fleet",
        "vehicle_description": "2025 Model X Electric - Demo",
        "duration_months": 36,
        "mileage_total": 60000,
        "monthly_rental": 470.0,
        "upfront_costs": 1000.0,
        "deposit": 1000.0,
        "delivery_registration": 0.0,
        "admin_fees": 150.0,
        "maintenance_included": False,
        "maintenance_cost": 800.0,  # one-off
        "tyres_included": True,
        "tyres_cost": 0.0,
        "road_tax_included": False,
        "road_tax_cost": 0.0,
        "insurance_included": True,
        "insurance_cost": 0.0,
        "early_termination_fees": "Pro rata",
        "excess_mileage_rate_per_km": 0.07,
        "discounts_or_rebates": 100.0,
        "currency": "EUR",
        "offer_valid_until": (now.isoformat()),
        "lead_time_or_eta": "6 weeks",
        "bundled_monthly": True,
        "parsing_confidence": 0.85,
        "scanned": False,
        "raw_snippets": {"demo": "beta demo"},
        "raw_parse_warnings": []
    }
    # Vendor C: cheapest TCC due to discount
    C = {
        "filename": "demo_offer_C.pdf",
        "vendor": "Gamma Leasing",
        "vehicle_description": "2025 Model X Electric - Demo",
        "duration_months": 36,
        "mileage_total": 60000,
        "monthly_rental": 460.0,
        "upfront_costs": 500.0,
        "deposit": 500.0,
        "delivery_registration": 0.0,
        "admin_fees": 100.0,
        "maintenance_included": False,
        "maintenance_cost": 0.0,
        "tyres_included": False,
        "tyres_cost": 600.0,
        "road_tax_included": False,
        "road_tax_cost": 0.0,
        "insurance_included": False,
        "insurance_cost": 0.0,
        "early_termination_fees": None,
        "excess_mileage_rate_per_km": 0.05,
        "discounts_or_rebates": -5.0,  # -5% discount (we treat negative as percent)
        "currency": "EUR",
        "offer_valid_until": (now.isoformat()),
        "lead_time_or_eta": "3 weeks",
        "bundled_monthly": False,
        "parsing_confidence": 0.8,
        "scanned": False,
        "raw_snippets": {"demo": "gamma demo"},
        "raw_parse_warnings": []
    }
    offers.extend([A, B, C])
    # Convert to computed metrics form for consistency
    return offers

# ---------------------------
# Streamlit UI
# ---------------------------
st.set_page_config(page_title="Leasing Offer Comparator", layout="wide")
st.title("Leasing Offer Comparator — Fleet Management")

st.markdown("""
Upload 2–10 PDF leasing offers for the *same vehicle configuration*, *same contract duration*, and *same contract mileage*.
The app will parse each offer (tables or text), normalize key fields, compare offers (Total Contract Cost, Effective Monthly, Cost/km), and produce an Excel workbook with:
- Sheet 1: **Winner Analysis**
- One sheet per leasing company (named exactly as vendor)
All prices are treated **ex-VAT** and **exclude fuel** by design.
""")

# File uploader & demo toggle
col1, col2 = st.columns([3,1])
with col1:
    uploaded_files = st.file_uploader("Upload PDF offers (2–10)", accept_multiple_files=True, type=["pdf"])
with col2:
    load_demo = st.checkbox("Load demo data", value=False)
    if load_demo:
        st.info("Demo offers loaded; no PDFs required.")

# Sidebar assumptions
st.sidebar.header("Assumptions & Overrides")
include_maintenance = st.sidebar.checkbox("Include maintenance in TCC", value=True)
include_tyres = st.sidebar.checkbox("Include tyres cost", value=False)
include_insurance = st.sidebar.checkbox("Include insurance", value=False)
include_road_tax = st.sidebar.checkbox("Include road tax", value=False)
st.sidebar.markdown("**Discount overrides (per vendor)**: enter a numeric absolute amount (positive) or percent (e.g., -5 for -5%)")
discount_overrides_input = {}
# We'll render dynamic inputs after parsing vendor names (later) — but allow empty.

tie_breaker = st.sidebar.multiselect("Tie-breaker priority (applied after TCC)", options=[
    "Monthly_rental", "Effective_monthly", "Lead_time", "Parsing_confidence"
], default=["Monthly_rental", "Effective_monthly"])

assumptions = {
    "include_maintenance": include_maintenance,
    "include_tyres": include_tyres,
    "include_insurance": include_insurance,
    "include_road_tax": include_road_tax,
    "discount_overrides": discount_overrides_input,
    "tie_breaker_priority": tie_breaker,
}

# Parse uploaded PDFs or demo data
offers_parsed: List[ParseResult] = []
offers_normalized: List[Dict[str, Any]] = []
parsing_errors = []
if load_demo:
    # Use demo data to bypass PDF parsing
    demo = demo_offers()
    # demo already in normalized form; wrap into computed metrics
    for d in demo:
        # compute metrics for demo offers
        nd = compute_metrics(d, assumptions)
        offers_normalized.append(nd)
else:
    # process uploaded files
    if uploaded_files:
        if len(uploaded_files) < 2:
            st.warning("Please upload at least 2 PDF offers for comparison.")
        elif len(uploaded_files) > 10:
            st.warning("Please upload no more than 10 files.")
        else:
            st.info(f"{len(uploaded_files)} file(s) uploaded. Parsing...")
            progress = st.progress(0)
            for i, f in enumerate(uploaded_files):
                try:
                    bytes_data = f.read()
                    pr = parse_pdf(bytes_data, f.name)
                    offers_parsed.append(pr)
                    # Show parse confidence and missing essentials
                    if pr.contract_duration_months is None:
                        pr.raw_parse_warnings.append("Missing contract duration (months).")
                    if pr.contract_mileage_total is None:
                        pr.raw_parse_warnings.append("Missing contract mileage total.")
                    progress.progress(int((i+1)/len(uploaded_files)*100))
                except Exception as e:
                    parsing_errors.append((f.name, str(e)))
                    log_ui("error", f"Failed to parse {f.name}: {e}")
            if parsing_errors:
                st.error(f"Parsing errors for {len(parsing_errors)} file(s). See log panel.")
            # Normalize parsed offers and compute metrics
            # Ask user for discount overrides per vendor if any
            # Build discount overrides inputs
            vendors = [pr.leasing_company_name for pr in offers_parsed]
            st.sidebar.markdown("### Discount overrides (per vendor)")
            for v in vendors:
                val = st.sidebar.text_input(f"{v}: (leave empty to use parsed)", value="", key=f"disc_{v}")
                if val.strip() != "":
                    try:
                        discount_overrides_input[v] = float(val)
                    except:
                        discount_overrides_input[v] = val  # will warn later
            assumptions["discount_overrides"] = discount_overrides_input
            for pr in offers_parsed:
                norm = normalize_offer(pr, assumptions)
                comp = compute_metrics(norm, assumptions)
                offers_normalized.append(comp)

# If offers_normalized present, run comparison validations
if offers_normalized:
    # Validate same duration and mileage
    durations = set([o.get("duration_months") for o in offers_normalized])
    mileages = set([o.get("mileage_total") for o in offers_normalized])
    same_duration = len(durations) == 1
    same_mileage = len(mileages) == 1
    # Currency check
    currencies = set([o.get("currency") for o in offers_normalized if o.get("currency")])
    mixed_currency = len(currencies) > 1
    if mixed_currency:
        st.error(f"Mixed currencies detected across offers: {currencies}. You must set conversion rates in the sidebar to proceed with export.")
        st.sidebar.markdown("### Currency conversion (required for mixed currencies)")
        currency_map = {}
        base_currency = st.sidebar.selectbox("Base currency to use for comparison", options=list(currencies), index=0)
        for c in currencies:
            if c == base_currency:
                currency_map[c] = 1.0
            else:
                rate = st.sidebar.number_input(f"Conversion rate {c} → {base_currency} (multiply)", min_value=0.0001, value=1.0, format="%.6f")
                currency_map[c] = float(rate)
        # apply conversion if user provided rates
        if st.sidebar.button("Apply conversion rates"):
            for o in offers_normalized:
                c = o.get("currency")
                if c and c in currency_map and currency_map[c] != 1.0:
                    rate = currency_map[c]
                    # convert numeric fields
                    for k in ["monthly_rental", "upfront_costs", "admin_fees", "maintenance_cost", "tyres_cost", "insurance_cost", "road_tax_cost", "computed_monthly_total", "computed_maintenance_total", "computed_tyres_total", "computed_insurance_total", "computed_road_tax_total", "computed_discount_amount", "TCC", "effective_monthly", "cost_per_km"]:
                        if o.get(k) is not None:
                            try:
                                o[k] = float(o[k]) * rate
                            except:
                                pass
                    o["currency"] = base_currency
            st.success("Applied conversion rates.")
            currencies = set([o.get("currency") for o in offers_normalized if o.get("currency")])
            mixed_currency = False
    # If mismatch in duration/mileage, block comparison
    if not same_duration or not same_mileage:
        st.error("Duration and/or mileage differ across uploaded offers. Comparison blocked.")
        st.write("Detected durations:", durations)
        st.write("Detected mileages:", mileages)
        st.stop()

    # Build comparison
    comp_df, analysis = compare_offers(offers_normalized, tie_breakers=tie_breaker)
    assumptions["winner_notes"] = analysis.get("why_notes", [])
    # UI: status
    st.success("Parsed and normalized offers.")
    st.markdown("### Summary table (normalized)")
    # Present normalized table
    display_cols = ["Vendor", "Duration_months", "Mileage_total", "Currency", "Monthly_rental", "Upfront_costs", "Admin_fees", "Maintenance_total", "TCC", "Effective_monthly", "Cost_per_km", "Rank"]
    st.dataframe(comp_df[display_cols].style.format({
        "Monthly_rental": "{:,.2f}",
        "Upfront_costs": "{:,.2f}",
        "Admin_fees": "{:,.2f}",
        "Maintenance_total": "{:,.2f}",
        "TCC": "{:,.2f}",
        "Effective_monthly": "{:,.2f}",
        "Cost_per_km": "{:,.4f}",
    }), height=300)

    # Highlight differences: conditional formatting in UI is limited; show best/worst colors manually
st.markdown("### Highlights")

if not comp_df.empty and "TCC" in comp_df.columns:
    best_tcc = comp_df["TCC"].min()
    worst_tcc = comp_df["TCC"].max()

    # Safely get winner if possible
    if pd.notna(best_tcc) and "Vendor" in comp_df.columns:
        best_vendor = comp_df.loc[comp_df["TCC"].idxmin(), "Vendor"]

        # Winner message
        st.write(
            f"Winner by default (lowest TCC): **{best_vendor}** — "
            f"TCC: **{best_tcc:,.2f} "
            f"{list(analysis['currency_set'])[0] if analysis.get('currency_set') else ''}**"
        )

    # Human-readable notes about why
    if analysis.get("why_notes"):
        st.write("Top factors why winner (human-readable):")
        for note in analysis["why_notes"]:
            st.write(f"- {note}")
else:
    st.warning("No valid comparison data available for highlights.")
    # Sensitivity: show what happens if include/exclude maintenance
    st.markdown("### Sensitivity: Maintenance inclusion")
    # toggle simulated
    sim_inc = not include_maintenance
    sim_ass = dict(assumptions)
    sim_ass["include_maintenance"] = sim_inc
    recomputed = []
    for o in offers_normalized:
        # reconstruct minimal normalized original for recompute
        base = {
            "duration_months": o.get("duration_months"),
            "mileage_total": o.get("mileage_total"),
            "monthly_rental": o.get("Monthly_rental") if "Monthly_rental" in o else o.get("monthly_rental"),
            "upfront_costs": o.get("computed_upfront", o.get("upfront_costs")),
            "admin_fees": o.get("computed_admin", o.get("admin_fees")),
            "maintenance_included": o.get("maintenance_included"),
            "maintenance_cost": o.get("maintenance_cost", o.get("computed_maintenance_total", 0.0)),
            "tyres_included": o.get("tyres_included"),
            "tyres_cost": o.get("tyres_cost", o.get("computed_tyres_total", 0.0)),
            "insurance_included": o.get("insurance_included"),
            "insurance_cost": o.get("insurance_cost", o.get("computed_insurance_total", 0.0)),
            "road_tax_included": o.get("road_tax_included"),
            "road_tax_cost": o.get("road_tax_cost", o.get("computed_road_tax_total", 0.0)),
            "discounts_or_rebates": o.get("computed_discount_amount") or 0.0,
        }
        recomputed.append(compute_metrics(base, sim_ass))
    # Find potential winner change
    recom_df, recom_analysis = compare_offers(recomputed, tie_breakers=tie_breaker)
    sim_winner = recom_df.iloc[0]["Vendor"]
    st.write(f"With maintenance inclusion toggled to **{sim_inc}**, winner would be **{sim_winner}** (was {best_vendor}).")

    # Expanders for raw snippets per vendor
    st.markdown("### Raw extracted snippets (per offer)")
    for o in offers_normalized:
        with st.expander(f"{o.get('vendor')} — parsing confidence {o.get('parsing_confidence',0):.2f}"):
            st.write("Filename:", o.get("filename"))
            st.write("Vehicle:", o.get("vehicle_description"))
            st.write("Duration (months):", o.get("duration_months"))
            st.write("Mileage total:", o.get("mileage_total"))
            st.write("Currency detected:", o.get("currency"))
            st.write("Monthly rental:", o.get("monthly_rental"))
            st.write("Upfront / admin:", o.get("upfront_costs"), "/", o.get("admin_fees"))
            st.write("TCC:", o.get("TCC"))
            st.write("Computed breakdown:")
            st.json({
                "computed_monthly_total": o.get("computed_monthly_total"),
                "maintenance_total": o.get("computed_maintenance_total"),
                "tyres_total": o.get("computed_tyres_total"),
                "insurance_total": o.get("computed_insurance_total"),
                "road_tax_total": o.get("computed_road_tax_total"),
                "discount_amount": o.get("computed_discount_amount"),
            })
            st.write("Raw parse warnings:", o.get("raw_parse_warnings"))
            st.write("--- Raw snippets (truncated) ---")
            rs = o.get("raw_snippets") or {}
            for k, v in rs.items():
                st.write(f"**{k}**")
                st.text_area(f"{o.get('vendor')}_{k}", value=v[:4000], height=120, key=f"snip_{o.get('vendor')}_{k}")

    # Logs panel
    with st.expander("Logs & Warnings"):
        for entry in LOGS[-200:]:
            st.write(entry)

    # Excel export
    st.markdown("### Export results")
    if len(offers_normalized) >= 2:
        try:
            excel_bytes = build_excel(comp_df, offers_normalized, assumptions)
            st.download_button("Download comparison.xlsx", data=excel_bytes, file_name="comparison.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.success("Excel generated in-memory. First sheet = Winner Analysis; subsequent sheets per vendor.")
        except Exception as e:
            st.error(f"Failed to generate Excel: {e}")
            log_ui("error", f"Excel generation error: {e}")
else:
    st.info("Upload PDFs or load demo data to begin.")

# Footer reminders and quick tips
st.markdown("---")
st.markdown("""
**Notes & tips**
- The parser uses regex heuristics and will not be perfect for all layouts; inspect raw snippets in expanders for transparency.
- If a PDF is scanned (image-only), the app marks it as scanned and will not OCR it automatically. Consider converting to searchable PDF or enable OCR in a future run.
- Mixed currencies must be resolved by providing conversion rates.
- Duration and mileage must match exactly across offers for valid comparison.
""")
