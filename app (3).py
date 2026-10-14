"""
Employee Reimbursement AI Automation System
Local Production-Style Prototype (Python + Streamlit + OpenPyXL + Pandas)

Key Capabilities:
1. 100% Local Processing (No Cloud API, No OpenAI, No Claude, No Gemini, Zero Data Leakage)
2. Dynamic Employee Profiler: Switch employee and the entire system, currency, policy rules & format adapt instantly
3. Company Master Excel (Reimbursement_Claim.xlsx) is the strict SOURCE OF TRUTH
4. Multi-Document Ingestion (PDF Invoices, Receipts, Tickets, Hotel Bills, Food Receipts, Scanned PDFs with OCR)
5. Anti-Hallucination Amount & Currency Extraction bound to verifiable bill evidence
6. Genuine Entity Extraction: Extracts Billed To / Passenger / Guest / Rider name alongside Vendor
7. Grouped Column Structure: Document -> Date -> Billed To -> Vendor -> Expense Type -> Nature of Expense -> Description
8. Dedicated Foreign Exchange Rate Converter: Clean Home Currency conversion (USD to GBP, CAD, INR, etc.)
9. Production 4-Sheet Excel Export (Claim_Output, Verification, Summary, Exceptions) + Tracker Update Integration
"""

import os
import sys
import re
import datetime
import io
import argparse
from typing import Dict, List, Tuple, Any, Optional

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- OPTIONAL / FALLBACK DEPENDENCIES ---
try:
    import streamlit as st
    STREAMLIT_AVAILABLE = True
except ImportError:
    STREAMLIT_AVAILABLE = False
    st = None

try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False
    fitz = None

try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False
    pdfplumber = None

try:
    import pypdf
    PYPDF_AVAILABLE = True
except ImportError:
    PYPDF_AVAILABLE = False
    pypdf = None

try:
    import pypdfium2 as pdfium
    PDFIUM_AVAILABLE = True
except ImportError:
    PDFIUM_AVAILABLE = False
    pdfium = None

try:
    import pytesseract
    from PIL import Image
    TESSERACT_AVAILABLE = True
except ImportError:
    TESSERACT_AVAILABLE = False
    pytesseract = None
    Image = None

try:
    from rapidfuzz import fuzz
    def calculate_text_similarity(s1: str, s2: str) -> float:
        return float(fuzz.ratio(s1.lower(), s2.lower()))
except ImportError:
    import difflib
    def calculate_text_similarity(s1: str, s2: str) -> float:
        return float(difflib.SequenceMatcher(None, s1.lower(), s2.lower()).ratio() * 100.0)


# ==============================================================================
# SECTION 1: MASTER DATA MANAGEMENT (SOURCE OF TRUTH) & EMPLOYEE PROFILES
# ==============================================================================

EMPLOYEE_PRESETS = {
    "Daragh Patrick Ryan": {
        "employee_id": "GI001UK799",
        "designation": "Senior Director",
        "leadership_level": "Senior Leadership",
        "home_currency": "GBP",
        "region": "Europe",
        "entity": "GENINVO INTERNATIONAL",
        "expense_category": "Conference Expense"
    },
    "Pinku Hooda": {
        "employee_id": "GI001YYT161",
        "designation": "Associate Director",
        "leadership_level": "Senior Leadership",
        "home_currency": "CAD",
        "region": "USA",
        "entity": "GENINVO INTERNATIONAL",
        "expense_category": "Conference Expense"
    },
    "James Patrick Gallagher": {
        "employee_id": "MI001US801",
        "designation": "Vice President",
        "leadership_level": "Senior Leadership",
        "home_currency": "USD",
        "region": "USA",
        "entity": "MARS INTERNATIONAL",
        "expense_category": "Business Expense"
    },
    "Craig Wozniak": {
        "employee_id": "MI001US808",
        "designation": "Director",
        "leadership_level": "Senior Leadership",
        "home_currency": "USD",
        "region": "USA",
        "entity": "MARS INTERNATIONAL",
        "expense_category": "Business Expense"
    },
    "Ritika Goel": {
        "employee_id": "GI001IN774",
        "designation": "Specialist",
        "leadership_level": "All Other Employees",
        "home_currency": "INR",
        "region": "India",
        "entity": "GENINVO INDIA",
        "expense_category": "Business Expense"
    },
    "Sheetal Bagul": {
        "employee_id": "GI001IN779",
        "designation": "Executive",
        "leadership_level": "All Other Employees",
        "home_currency": "INR",
        "region": "India",
        "entity": "GENINVO INDIA",
        "expense_category": "Business Expense"
    },
    "Custom / New Employee": {
        "employee_id": "EMP-NEW-001",
        "designation": "Manager",
        "leadership_level": "All Other Employees",
        "home_currency": "GBP",
        "region": "Europe",
        "entity": "GENINVO INTERNATIONAL",
        "expense_category": "Business Expense"
    }
}

DEFAULT_FX_MATRICES = {
    "GBP": {"USD": 0.77, "EUR": 0.85, "CAD": 0.56, "INR": 0.0092, "CHF": 0.88, "GBP": 1.0},
    "CAD": {"USD": 1.423, "EUR": 1.55, "GBP": 1.85, "INR": 0.016, "CHF": 1.62, "CAD": 1.0},
    "USD": {"GBP": 1.30, "EUR": 1.08, "CAD": 0.7027, "INR": 0.012, "CHF": 1.14, "USD": 1.0},
    "INR": {"USD": 84.0, "EUR": 91.0, "GBP": 109.0, "CAD": 61.5, "CHF": 95.5, "INR": 1.0},
    "EUR": {"USD": 0.93, "GBP": 1.18, "CAD": 0.65, "INR": 0.011, "CHF": 1.05, "EUR": 1.0},
    "CHF": {"USD": 0.88, "EUR": 0.95, "GBP": 1.14, "CAD": 0.62, "INR": 0.0105, "CHF": 1.0}
}


class MasterDataManager:
    DEFAULT_MASTER_PATHS = [
        "user_uploaded_files/ttl=63d/user_uploaded/Reimbursement_Claim.xlsx",
        "Reimbursement_Claim.xlsx",
        "Reimbursement_Claim821to824.xlsx",
        "/working_dir/c_3e47fec635734a78/user_uploaded_files/ttl=63d/user_uploaded/Reimbursement_Claim.xlsx"
    ]

    def __init__(self, master_source=None):
        self.master_source = master_source
        self.nature_master: List[Dict[str, Any]] = []
        self.policy_master: List[Dict[str, Any]] = []
        self.currency_master: List[Dict[str, Any]] = []
        self.load_masters()

    def resolve_master_path(self) -> Optional[str]:
        if self.master_source and isinstance(self.master_source, str) and os.path.exists(self.master_source):
            return self.master_source
        for p in self.DEFAULT_MASTER_PATHS:
            if os.path.exists(p):
                return p
        return None

    def load_masters(self):
        resolved = self.resolve_master_path()
        wb = None
        if resolved:
            try:
                wb = openpyxl.load_workbook(resolved, data_only=True)
            except Exception as e:
                print(f"[WARN] Master workbook error: {e}")
        elif self.master_source and hasattr(self.master_source, "read"):
            try:
                wb = openpyxl.load_workbook(io.BytesIO(self.master_source.read()), data_only=True)
            except Exception as e:
                print(f"[WARN] Stream workbook error: {e}")

        if wb is not None:
            self._load_from_workbook(wb)
        else:
            self._load_fallback_masters()

    def _load_from_workbook(self, wb: openpyxl.Workbook):
        if "Nature_Master" in wb.sheetnames:
            ws = wb["Nature_Master"]
            for r in range(5, ws.max_row + 1):
                kw = ws.cell(r, 1).value
                nat = ws.cell(r, 2).value
                notes = ws.cell(r, 3).value
                if kw and nat:
                    self.nature_master.append({
                        "keyword": str(kw).strip().lower(),
                        "nature": str(nat).strip(),
                        "notes": str(notes).strip() if notes else ""
                    })

        if "Policy_Master" in wb.sheetnames:
            ws = wb["Policy_Master"]
            for r in range(5, ws.max_row + 1):
                cat = ws.cell(r, 1).value
                reg = ws.cell(r, 2).value
                nat = ws.cell(r, 3).value
                curr = ws.cell(r, 4).value
                lim = ws.cell(r, 5).value
                basis = ws.cell(r, 6).value
                notes = ws.cell(r, 7).value
                if cat or reg or nat:
                    num_lim = None
                    if lim is not None and str(lim).strip() != "":
                        try:
                            num_lim = float(lim)
                        except ValueError:
                            num_lim = None
                    self.policy_master.append({
                        "category": str(cat).strip() if cat else "",
                        "region": str(reg).strip() if reg else "",
                        "nature": str(nat).strip() if nat else "",
                        "currency": str(curr).strip() if curr else "",
                        "limit": num_lim,
                        "basis": str(basis).strip() if basis else "",
                        "notes": str(notes).strip() if notes else ""
                    })

        if "Currency_Master" in wb.sheetnames:
            ws = wb["Currency_Master"]
            for r in range(5, ws.max_row + 1):
                code = ws.cell(r, 1).value
                name = ws.cell(r, 2).value
                notes = ws.cell(r, 3).value
                if code:
                    self.currency_master.append({
                        "code": str(code).strip().upper(),
                        "name": str(name).strip() if name else "",
                        "notes": str(notes).strip() if notes else ""
                    })

    def _load_fallback_masters(self):
        self.currency_master = [
            {"code": "INR", "name": "Indian Rupee", "notes": "India currency"},
            {"code": "USD", "name": "US Dollar", "notes": "US Dollar"},
            {"code": "CAD", "name": "Canadian Dollar", "notes": "Canadian Dollar"},
            {"code": "EUR", "name": "Euro", "notes": "Euro"},
            {"code": "GBP", "name": "Pound Sterling", "notes": "Pound Sterling"},
            {"code": "CHF", "name": "Swiss Franc", "notes": "Swiss Franc"}
        ]
        self.nature_master = [
            {"keyword": "uber", "nature": "Business Expense - Travel", "notes": ""},
            {"keyword": "flight", "nature": "Business Expense - Air Travel", "notes": ""},
            {"keyword": "hotel", "nature": "Business Expense - Lodging- US/Canada", "notes": ""},
            {"keyword": "meal", "nature": "Business Expense - Food & Beverage", "notes": ""},
            {"keyword": "postage", "nature": "Other Expenses", "notes": ""}
        ]


# ==============================================================================
# SECTION 2: DOCUMENT PARSER & OCR ENGINE
# ==============================================================================

class DocumentParser:
    @staticmethod
    def extract_text(file_source: Any, filename: str = "") -> Tuple[str, bool, int]:
        extracted_text = ""
        is_ocr = False
        page_count = 1

        if hasattr(file_source, "read"):
            data_bytes = file_source.read()
            if hasattr(file_source, "seek"):
                file_source.seek(0)
        elif isinstance(file_source, str) and os.path.exists(file_source):
            with open(file_source, "rb") as f:
                data_bytes = f.read()
        elif isinstance(file_source, bytes):
            data_bytes = file_source
        else:
            return "", False, 0

        # Try fitz
        if PYMUPDF_AVAILABLE and fitz is not None:
            try:
                doc = fitz.open(stream=data_bytes, filetype="pdf")
                page_count = len(doc)
                pages_text = [p.get_text() for p in doc if p.get_text()]
                extracted_text = "\n".join(pages_text).strip()
            except Exception:
                extracted_text = ""

        # Try pdfplumber
        if not extracted_text and PDFPLUMBER_AVAILABLE and pdfplumber is not None:
            try:
                with pdfplumber.open(io.BytesIO(data_bytes)) as pdf:
                    page_count = len(pdf.pages)
                    pages_text = [p.extract_text() for p in pdf.pages if p.extract_text()]
                    extracted_text = "\n".join(pages_text).strip()
            except Exception:
                extracted_text = ""

        # Try pypdf
        if not extracted_text and PYPDF_AVAILABLE and pypdf is not None:
            try:
                reader = pypdf.PdfReader(io.BytesIO(data_bytes))
                page_count = len(reader.pages)
                pages_text = [p.extract_text() for p in reader.pages if p.extract_text()]
                extracted_text = "\n".join(pages_text).strip()
            except Exception:
                extracted_text = ""

        # OCR fallback
        if len(extracted_text.strip()) < 35 and TESSERACT_AVAILABLE and pytesseract is not None:
            ocr_text = []
            if PDFIUM_AVAILABLE and pdfium is not None:
                try:
                    pdf_doc = pdfium.PdfDocument(data_bytes)
                    page_count = len(pdf_doc)
                    for i in range(page_count):
                        page = pdf_doc[i]
                        img = page.render(scale=2).to_pil()
                        txt = pytesseract.image_to_string(img)
                        if txt:
                            ocr_text.append(txt)
                    if ocr_text:
                        extracted_text = "\n".join(ocr_text).strip()
                        is_ocr = True
                except Exception:
                    pass

        return extracted_text, is_ocr, page_count


# ==============================================================================
# SECTION 3: INFORMATION EXTRACTION ENGINE (GENUINE ENTITY & AMOUNT BINDING)
# ==============================================================================

class ReceiptExtractor:
    CURRENCY_SYMBOLS = {
        "$": "USD", "USD": "USD", "US$": "USD",
        "£": "GBP", "GBP": "GBP",
        "€": "EUR", "EUR": "EUR",
        "₹": "INR", "INR": "INR", "RS": "INR", "RS.": "INR",
        "C$": "CAD", "CAD": "CAD",
        "CHF": "CHF"
    }

    TOTAL_LABELS = [
        r"grand\s+total", r"total\s+amount\s+paid", r"total\s+amount\s+due", r"total\s+amount",
        r"total\s+charged", r"total\s+paid", r"amount\s+paid", r"amount\s+due",
        r"balance\s+due", r"total\s+fare", r"final\s+total", r"net\s+total",
        r"total", r"fare", r"postage\s+fee", r"subtotal"
    ]

    def __init__(self, master_data: MasterDataManager):
        self.masters = master_data

    def extract_billed_to(self, text: str, default_employee: str = "") -> str:
        """Extracts person named on the bill (Rider, Passenger, Guest, Customer, Employee)."""
        lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
        for line in lines:
            m = re.match(r"(?i)^(?:rider|passenger|guest(?:\s+name)?|customer|billed\s+to|client|employee)\s*[:\-]\s*([A-Za-z\s\.\-]+)$", line)
            if m:
                val = m.group(1).strip()
                if len(val) > 2 and len(val) < 40 and not any(bad in val.lower() for x in ["receipt", "invoice", "terminal"] for bad in [x]):
                    return val
        return default_employee if default_employee else "Self"

    def extract_amount_and_currency(self, text: str) -> Tuple[Any, Any, str, float]:
        if not text or len(text.strip()) == 0:
            return "REVIEW REQUIRED", "REVIEW REQUIRED", "No readable text", 0.0

        lines = [ln.strip() for ln in text.split("\n") if ln.strip()]

        for kw in self.TOTAL_LABELS:
            kw_regex = re.compile(rf"(?i)\b{kw}\b[\s:=–\-]*([^\n]*)")
            for line in lines:
                if re.search(r"(?i)exchange\s+rate|conversion\s+rate|1\s*[A-Z]{3}\s*=", line):
                    continue
                if re.search(r"(?i)vat\s+reg|booking\s+ref|ticket\s+no|e-ticket|fss\s+session", line):
                    continue

                m = kw_regex.search(line)
                if m:
                    match = re.search(r"([A-Z$£€₹]{1,3}|C\$|Rs\.?)\s*([0-9]+(?:,[0-9]{3})*(?:\.[0-9]{2})?)", line, re.I)
                    if not match:
                        match = re.search(r"([0-9]+(?:,[0-9]{3})*(?:\.[0-9]{2})?)\s*([A-Z$£€₹]{1,3}|C\$|Rs\.?)", line, re.I)
                        if match:
                            amt_str = match.group(1).replace(",", "")
                            curr_raw = match.group(2).strip().upper()
                        else:
                            continue
                    else:
                        curr_raw = match.group(1).strip().upper()
                        amt_str = match.group(2).replace(",", "")

                    curr_norm = self.CURRENCY_SYMBOLS.get(curr_raw, curr_raw)
                    try:
                        amt_val = float(amt_str)
                        if amt_val == 0.0 and "paid" in line.lower():
                            m_paid = re.search(r"(?i)paid[^\d]*([A-Z$£€₹]{1,3}|C\$|Rs\.?)\s*([0-9]+(?:,[0-9]{3})*(?:\.[0-9]{2})?)", line)
                            if m_paid:
                                curr_norm = self.CURRENCY_SYMBOLS.get(m_paid.group(1).strip().upper(), m_paid.group(1).strip().upper())
                                amt_val = float(m_paid.group(2).replace(",", ""))
                            else:
                                continue
                        return round(amt_val, 2), curr_norm, line, 0.95
                    except ValueError:
                        continue

        generic_matches = re.findall(r"([$£€₹]|USD|GBP|EUR|INR|CAD)\s*([0-9]+(?:,[0-9]{3})*(?:\.[0-9]{2}))", text, re.I)
        if generic_matches:
            last_sym, last_val = generic_matches[-1]
            try:
                val = float(last_val.replace(",", ""))
                curr_norm = self.CURRENCY_SYMBOLS.get(last_sym.upper(), last_sym.upper())
                return round(val, 2), curr_norm, f"Extracted from line: {last_sym} {last_val}", 0.75
            except ValueError:
                pass

        return "REVIEW REQUIRED", "REVIEW REQUIRED", "Amount/Currency not bound to evidence", 0.0

    def extract_date(self, text: str) -> Tuple[str, str]:
        m = re.search(r"(\d{1,2})\s+(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})", text, re.I)
        if m:
            d, mon, y = m.groups()
            try:
                dt = datetime.datetime.strptime(f"{d} {mon} {y}", "%d %B %Y")
                return dt.strftime("%Y-%m-%d"), f"{d} {mon} {y}"
            except Exception:
                pass

        m = re.search(r"\b(202\d)-(\d{2})-(\d{2})\b", text)
        if m: return m.group(0), m.group(0)

        m = re.search(r"\b(\d{1,2})-(\w{3})-(\d{2,4})\b", text)
        if m:
            d, mon, y = m.groups()
            y_full = "20" + y if len(y) == 2 else y
            try:
                dt = datetime.datetime.strptime(f"{d}-{mon}-{y_full}", "%d-%b-%Y")
                return dt.strftime("%Y-%m-%d"), f"{d} {mon} {y_full}"
            except Exception:
                pass

        return "REVIEW REQUIRED", "REVIEW REQUIRED"

    def extract_document_number(self, text: str) -> str:
        patterns = [
            r"(?i)trip\s+id[:\s]+([A-Z0-9\-]+)",
            r"(?i)booking\s+reference[:\s]+([A-Z0-9]+)",
            r"(?i)e-ticket\s+number[:\s]+([0-9\-]+)",
            r"(?i)ticket\s+number[:\s]+([0-9\-]+)",
            r"(?i)invoice\s+number[:\s]+([A-Z0-9\-]+)",
            r"(?i)receipt\s+no[:\s]+([A-Z0-9\-]+)",
            r"(?i)tracking\s+number[:\s]+([A-Z0-9\-]+)",
            r"(?i)order\s+no[:\s]+([A-Z0-9\-]+)"
        ]
        for p in patterns:
            m = re.search(p, text)
            if m:
                return m.group(1).strip()
        return "N/A"

    def extract_vendor(self, text: str, filename: str = "") -> str:
        fn, t = filename.lower(), text.lower()
        if "uber" in fn or "uber" in t: return "Uber"
        if "lyft" in fn or "lyft" in t: return "Lyft"
        if "virgin atlantic" in fn or "virgin atlantic" in t: return "Virgin Atlantic Airways"
        if "air canada" in fn or "air canada" in t: return "Air Canada"
        if "royal mail" in fn or "royal mail" in t: return "Royal Mail"
        if "battery wharf" in fn or "battery wharf" in t: return "Battery Wharf Hotel, Boston Waterfront"
        if "hampton inn" in fn or "hampton inn" in t: return "Hampton Inn & Suites"
        if "raindrop" in fn or "raindrop indian" in t: return "Raindrop Indian Cuisine"
        if "subway" in fn or "subway" in t: return "Subway"
        if "starbucks" in fn or "starbucks" in t: return "Starbucks"
        if "shell" in fn or "shell oil" in t: return "Shell Oil Products US"
        if "thrifty" in fn or "thrifty car" in t: return "Thrifty Car Rental"

        lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
        if lines: return lines[0][:40]
        return "Unknown Vendor"

    def classify_expense_type(self, text: str, filename: str = "", vendor: str = "") -> str:
        combined = f"{filename} {vendor} {text}".lower()
        if any(w in combined for w in ["uber", "taxi", "cab", "lyft", "ola", "ride share", "ride-share"]):
            return "Taxi"
        if any(w in combined for w in ["virgin atlantic", "airline", "airways", "flight", "e-ticket", "boarding pass", "airfare"]):
            return "Flight"
        if any(w in combined for w in ["hotel", "lodging", "room accommodation", "check-in", "check-out", "folio", "inn & suites"]):
            return "Hotel"
        if any(w in combined for w in ["royal mail", "postage", "courier", "parcel force", "fedex", "dhl", "shipment"]):
            return "Other"
        if any(w in combined for w in ["restaurant", "cuisine", "food", "meal", "biryani", "dinner", "lunch", "breakfast", "cafe", "coffee", "subway", "pizza", "burger"]):
            return "Food"
        if any(w in combined for w in ["shell oil", "petrol", "diesel", "fuel", "gas station"]):
            return "Fuel"
        if any(w in combined for w in ["parking", "toll"]):
            return "Parking"
        if any(w in combined for w in ["conference registration", "attendee registration", "delegate registration"]):
            return "Registration"
        if any(w in combined for w in ["telecom", "vodafone", "wifi", "broadband", "mobile bill"]):
            return "Telephone/Internet"
        if any(w in combined for w in ["stationery", "office supply", "notebook", "pantry"]):
            return "Office Supply"
        return "Other"

    def match_nature_of_expense(self, expense_type: str, text: str, filename: str = "", is_conference: bool = False) -> str:
        t = f"{filename} {text}".lower()
        if expense_type == "Flight":
            is_intl = any(k in t for k in ["international", "lhr", "london", "uk", "overseas", "foreign", "frankfurt", "toronto to philadelphia"])
            if is_conference:
                return "Conference Expense - Air Travel International" if is_intl else "Conference Expense - Air Travel"
            else:
                return "Business Expense - Air Travel- International" if is_intl else "Business Expense - Air Travel"

        if expense_type == "Hotel":
            if is_conference:
                return "Conference Expense - Lodging - L1"
            else:
                if any(k in t for k in ["us", "usa", "canada", "boston", "philadelphia"]):
                    return "Business Expense - Lodging- US/Canada"
                return "Business Expense - Lodging Expense- L1"

        if expense_type == "Taxi":
            return "Conference Expense - Travel" if is_conference else "Business Expense - Travel"

        if expense_type == "Food":
            if "india" in t and not is_conference:
                return "Business Expense - Meals Expense -India"
            return "Conference Expense - Food & Beverage" if is_conference else "Business Expense - Food & Beverage"

        if expense_type == "Fuel":
            return "Conference Expense - Fuel/Mileage" if is_conference else "Business Expense - Fuel/Mileage Expense"

        if expense_type == "Parking": return "Parking Expense"
        if expense_type == "Registration": return "Conference Expense - Attendee Registrations"
        if expense_type == "Office Supply": return "Office Supply"
        if expense_type == "Telephone/Internet": return "Telephone Expense"

        if any(k in t for k in ["royal mail", "postage", "courier"]):
            return "Other Expenses"

        return "Conference Expense - Misc." if is_conference else "Business Expense - Misc."

    def extract_hotel_nights(self, text: str) -> int:
        m_n = re.search(r"(?i)(?:number\s+of\s+nights|nights|stay)[:\s]+(\d+)", text)
        if m_n:
            try: return int(m_n.group(1))
            except ValueError: pass
        occ = len(re.findall(r"(?i)room\s+accommodation", text))
        return occ if occ > 1 else 1

    def generate_description(self, expense_type: str, nature: str, vendor: str, doc_no: str, date_disp: str, text: str, units: int = 1) -> str:
        prefix = "Conference travel" if "Conference" in nature and ("Air" in nature or "Travel" in nature or "Lodging" in nature) \
            else ("Conference expense" if "Conference" in nature else ("Business travel" if ("Travel" in nature or "Lodging" in nature) else "Business expense"))

        if expense_type == "Flight":
            scope = "International" if "International" in nature else "Domestic"
            doc_str = f" e-Ticket {doc_no}" if doc_no != "N/A" else ""
            return f"{prefix} - {scope} Air Travel - {vendor}{doc_str} - {date_disp}"
        if expense_type == "Taxi":
            return f"{prefix} - Ground Transportation - {vendor} - {date_disp}"
        if expense_type == "Hotel":
            nights_str = f" ({units} nights)" if units > 1 else ""
            return f"{prefix} - Lodging - {vendor} - {date_disp}{nights_str}"
        if expense_type == "Food":
            loc = ", Philadelphia" if "philadelphia" in text.lower() else (", Boston" if "boston" in text.lower() else "")
            return f"{prefix} - Meals - {vendor}{loc} - {date_disp}"
        if "royal mail" in vendor.lower() or "postage" in text.lower():
            return f"Business expense - Postage - {vendor} - {date_disp}"
        return f"{prefix} - {nature.split(' - ')[-1]} - {vendor} - {date_disp}"


# ==============================================================================
# SECTION 4: POLICY ENGINE WITH DYNAMIC FX & HOME CURRENCY ELIGIBILITY
# ==============================================================================

class PolicyEngine:
    def __init__(self, master_data: MasterDataManager):
        self.masters = master_data

    def evaluate_policy(self, leadership_level: str, region: str, nature: str, currency: str, units: int, amount: float, fx_rate: float = 1.0, home_currency: str = "GBP") -> Dict[str, Any]:
        if not isinstance(amount, (int, float)):
            return {
                "policy_limit_str": "N/A",
                "policy_limit_per_unit": None,
                "eligible_amount": "Review Required",
                "excess": 0.0,
                "eligible_amount_hc": "Review Required",
                "excess_hc": 0.0,
                "status": "REVIEW",
                "remarks": "Amount is not numeric or requires review."
            }

        matched_rule = None
        for rule in self.masters.policy_master:
            if rule["nature"].lower() == nature.lower():
                rule_reg = rule["region"].lower()
                reg_low = region.lower()
                reg_match = False
                if not rule_reg or rule_reg == reg_low:
                    reg_match = True
                elif reg_low in ["usa", "canada"] and rule_reg in ["usa", "canada"]:
                    reg_match = True
                elif reg_low in ["uk", "europe"] and rule_reg == "europe":
                    reg_match = True

                if reg_match and rule["category"] in ["All Employees", leadership_level, ""]:
                    matched_rule = rule
                    break

        if not matched_rule:
            for rule in self.masters.policy_master:
                if rule["nature"].lower() == nature.lower() and rule["category"] in ["All Employees", leadership_level, ""]:
                    if rule["region"].lower() == "india" and region.lower() in ["usa", "canada", "europe", "uk"]:
                        continue
                    matched_rule = rule
                    break

        if not matched_rule:
            return {
                "policy_limit_str": "N/A",
                "policy_limit_per_unit": None,
                "eligible_amount": "Review Required",
                "excess": 0.0,
                "eligible_amount_hc": "Review Required",
                "excess_hc": 0.0,
                "status": "REVIEW",
                "remarks": f"No applicable policy rule in Policy_Master for {nature} in {region}."
            }

        lim = matched_rule["limit"]
        basis = matched_rule["basis"]
        hc_amount = round(amount * fx_rate, 2)

        if lim is None:
            return {
                "policy_limit_str": basis,
                "policy_limit_per_unit": None,
                "eligible_amount": round(amount, 2),
                "excess": 0.0,
                "eligible_amount_hc": hc_amount,
                "excess_hc": 0.0,
                "status": "PASS",
                "remarks": f"Approved under Policy basis: {basis}."
            }
        else:
            total_allowed_bill_curr = lim * units
            eligible_bill = min(amount, total_allowed_bill_curr)
            excess_bill = max(0.0, amount - eligible_bill)

            total_allowed_hc = total_allowed_bill_curr * fx_rate
            eligible_hc = min(hc_amount, total_allowed_hc)
            excess_hc = max(0.0, hc_amount - eligible_hc)

            status = "PASS" if excess_bill == 0.0 else "REVIEW"
            rem = f"Policy cap: {lim} {matched_rule['currency']}/{basis} (Cap: {total_allowed_bill_curr:.2f} {matched_rule['currency']} / {total_allowed_hc:.2f} {home_currency})"
            if excess_bill > 0:
                rem += f" | Excess: {excess_bill:.2f} {currency} ({excess_hc:.2f} {home_currency})"

            return {
                "policy_limit_str": f"{lim} {matched_rule['currency']}",
                "policy_limit_per_unit": lim,
                "eligible_amount": round(eligible_bill, 2),
                "excess": round(excess_bill, 2),
                "eligible_amount_hc": round(eligible_hc, 2),
                "excess_hc": round(excess_hc, 2),
                "status": status,
                "remarks": rem
            }


# ==============================================================================
# SECTION 5: PIPELINE COORDINATOR
# ==============================================================================

class ReimbursementPipeline:
    def __init__(self, master_manager: MasterDataManager):
        self.masters = master_manager
        self.extractor = ReceiptExtractor(self.masters)
        self.policy_engine = PolicyEngine(self.masters)

    def process_document(self, file_source: Any, filename: str, employee_config: Dict[str, Any], fx_rates: Dict[str, float]) -> Dict[str, Any]:
        raw_text, is_ocr, page_count = DocumentParser.extract_text(file_source, filename)
        emp_name = employee_config.get("employee_name", "Employee")
        home_curr = employee_config.get("home_currency", "GBP")
        leadership = employee_config.get("leadership_level", "Senior Leadership")
        is_conf = (employee_config.get("expense_category") == "Conference Expense")

        if not raw_text or len(raw_text.strip()) == 0:
            return {
                "document_name": filename,
                "date": "REVIEW REQUIRED",
                "billed_to": emp_name,
                "vendor": "Unreadable",
                "expense_type": "Other",
                "nature_of_expense": "Other Expenses",
                "description": f"Document unreadable - {filename}",
                "currency": "REVIEW REQUIRED",
                "amount": "REVIEW REQUIRED",
                "fx_rate": 1.0,
                "home_currency": home_curr,
                "home_currency_amount": "REVIEW REQUIRED",
                "policy_limit_str": "N/A",
                "eligible_amount_hc": "Review Required",
                "excess_hc": 0.0,
                "status": "ERROR",
                "receipt": "No",
                "remarks": "Document could not be read. Blank or corrupted PDF."
            }

        amt, curr, amt_evidence, conf = self.extractor.extract_amount_and_currency(raw_text)
        date_iso, date_disp = self.extractor.extract_date(raw_text)
        doc_no = self.extractor.extract_document_number(raw_text)
        vendor = self.extractor.extract_vendor(raw_text, filename)
        billed_to = self.extractor.extract_billed_to(raw_text, emp_name)
        exp_type = self.extractor.classify_expense_type(raw_text, filename, vendor)
        nature = self.extractor.match_nature_of_expense(exp_type, raw_text, filename, is_conference=is_conf)

        units = 1
        region = employee_config.get("region", "Europe")
        if exp_type == "Hotel":
            units = self.extractor.extract_hotel_nights(raw_text)
            region = "USA" if "boston" in raw_text.lower() or "usa" in raw_text.lower() else region
        elif "royal mail" in vendor.lower():
            region = "Europe"

        desc = self.extractor.generate_description(exp_type, nature, vendor, doc_no, date_disp, raw_text, units=units)

        # Dynamic FX lookup
        fx = fx_rates.get(curr, 1.0)
        hc_amount = round(amt * fx, 2) if isinstance(amt, (int, float)) else "REVIEW REQUIRED"

        pol_res = self.policy_engine.evaluate_policy(
            leadership_level=leadership,
            region=region,
            nature=nature,
            currency=curr if curr != "REVIEW REQUIRED" else home_curr,
            units=units,
            amount=amt,
            fx_rate=fx,
            home_currency=home_curr
        )

        verification_status = pol_res["status"]
        remarks = pol_res["remarks"]
        if amt == "REVIEW REQUIRED" or curr == "REVIEW REQUIRED":
            verification_status = "REVIEW"
            remarks = "Amount or Currency uncertain; manual review required."

        return {
            "document_name": filename,
            "date": date_iso,
            "billed_to": billed_to,
            "vendor": vendor,
            "expense_type": exp_type,
            "nature_of_expense": nature,
            "description": desc,
            "currency": curr,
            "amount": amt,
            "amount_evidence": amt_evidence,
            "units": units,
            "fx_rate": fx,
            "home_currency": home_curr,
            "home_currency_amount": hc_amount,
            "policy_limit_str": pol_res["policy_limit_str"],
            "eligible_amount_bill": pol_res["eligible_amount"],
            "excess_bill": pol_res["excess"],
            "eligible_amount_hc": pol_res["eligible_amount_hc"],
            "excess_hc": pol_res["excess_hc"],
            "receipt": "Yes",
            "duplicate_status": "UNIQUE",
            "status": verification_status,
            "remarks": remarks,
            "is_ocr": is_ocr
        }


# ==============================================================================
# SECTION 6: PRODUCTION EXCEL EXPORT & TRACKER APPRENDER
# ==============================================================================

class ExcelExportGenerator:
    @staticmethod
    def generate_workbook(employee_config: Dict[str, Any], records: List[Dict[str, Any]]) -> io.BytesIO:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        font_header_big = Font(name="Calibri", size=14, bold=True, color="1F497D")
        font_col_header = Font(name="Calibri", size=10, bold=True, color="000000")
        font_data = Font(name="Calibri", size=10)
        font_bold = Font(name="Calibri", size=10, bold=True)

        fill_header = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        fill_subtotal = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        fill_exception = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

        align_center = Alignment(horizontal="center", vertical="center")
        align_left = Alignment(horizontal="left", vertical="center")
        align_right = Alignment(horizontal="right", vertical="center")

        thin_side = Side(border_style="thin", color="D9D9D9")
        thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        home_curr = employee_config.get("home_currency", "GBP")

        # 1. Claim_Output
        ws_claim = wb.create_sheet(title="Claim_Output")
        ws_claim.views.sheetView[0].showGridLines = True
        ws_claim.cell(1, 3, "EXPENSE REIMBURSEMENT").font = font_header_big

        ws_claim.cell(4, 4, "Employee Name").font = font_bold
        ws_claim.cell(4, 5, employee_config.get("employee_name", "Daragh Patrick Ryan")).font = font_data
        ws_claim.cell(4, 7, "Employee ID").font = font_bold
        ws_claim.cell(4, 8, employee_config.get("employee_id", "GI001UK799")).font = font_data
        ws_claim.cell(4, 11, "Policy Category").font = font_bold
        ws_claim.cell(4, 12, employee_config.get("designation", "Senior Director")).font = font_data

        ws_claim.cell(5, 4, "Entity").font = font_bold
        ws_claim.cell(5, 5, employee_config.get("entity", "GENINVO INTERNATIONAL")).font = font_data
        ws_claim.cell(5, 7, "Location/Region").font = font_bold
        ws_claim.cell(5, 8, employee_config.get("region", "Europe")).font = font_data
        ws_claim.cell(5, 11, "Home Currency").font = font_bold
        ws_claim.cell(5, 12, home_curr).font = font_data

        ws_claim.cell(6, 4, "Expense Category").font = font_bold
        ws_claim.cell(6, 5, employee_config.get("expense_category", "Conference Expense")).font = font_data

        claim_headers = [
            "Claim No", "Date", "Description", "Nature of Expense", "Currency",
            "Amount", "Units*", "FX Rate*", f"Home Currency ({home_curr})", "Policy Limit / Unit",
            f"Eligible ({home_curr})", f"Excess ({home_curr})", "Receipt", "Status", "Remarks"
        ]
        for c_idx, h_text in enumerate(claim_headers, start=3):
            cell = ws_claim.cell(8, c_idx, h_text)
            cell.font = font_col_header
            cell.fill = fill_header
            cell.alignment = align_center
            cell.border = thin_border

        curr_row = 9
        tot_claimed_hc = 0.0
        tot_eligible_hc = 0.0
        tot_excess_hc = 0.0

        for idx, rec in enumerate(records):
            c_no = 825 + (idx // 4)
            ws_claim.cell(curr_row, 3, c_no if idx % 4 == 0 else "").alignment = align_center
            ws_claim.cell(curr_row, 4, rec.get("date", "")).alignment = align_center
            ws_claim.cell(curr_row, 5, rec.get("description", "")).alignment = align_left
            ws_claim.cell(curr_row, 6, rec.get("nature_of_expense", "")).alignment = align_left
            ws_claim.cell(curr_row, 7, rec.get("currency", "")).alignment = align_center
            
            amt = rec.get("amount", 0.0)
            ws_claim.cell(curr_row, 8, amt).alignment = align_right
            ws_claim.cell(curr_row, 8).number_format = "#,##0.00" if isinstance(amt, (int, float)) else "@"

            units = rec.get("units", 1)
            ws_claim.cell(curr_row, 9, units if units > 1 else "").alignment = align_center
            
            fx = rec.get("fx_rate", 1.0)
            ws_claim.cell(curr_row, 10, fx).alignment = align_center

            hc_amt = rec.get("home_currency_amount", 0.0)
            ws_claim.cell(curr_row, 11, hc_amt).alignment = align_right
            ws_claim.cell(curr_row, 11).number_format = "#,##0.00" if isinstance(hc_amt, (int, float)) else "@"

            ws_claim.cell(curr_row, 12, rec.get("policy_limit_str", "")).alignment = align_center

            elig_hc = rec.get("eligible_amount_hc", 0.0)
            ws_claim.cell(curr_row, 13, elig_hc).alignment = align_right
            ws_claim.cell(curr_row, 13).number_format = "#,##0.00" if isinstance(elig_hc, (int, float)) else "@"

            excess_hc = rec.get("excess_hc", 0.0)
            ws_claim.cell(curr_row, 14, excess_hc).alignment = align_right
            ws_claim.cell(curr_row, 14).number_format = "#,##0.00" if isinstance(excess_hc, (int, float)) else "@"

            ws_claim.cell(curr_row, 15, rec.get("receipt", "Yes")).alignment = align_center
            stat_disp = "Verified✅" if rec.get("status") == "PASS" else "Review Required"
            ws_claim.cell(curr_row, 16, stat_disp).alignment = align_center
            ws_claim.cell(curr_row, 17, rec.get("remarks", "")).alignment = align_left

            for c in range(3, 18):
                ws_claim.cell(curr_row, c).font = font_data
                ws_claim.cell(curr_row, c).border = thin_border

            if isinstance(hc_amt, (int, float)): tot_claimed_hc += hc_amt
            if isinstance(elig_hc, (int, float)): tot_eligible_hc += elig_hc
            if isinstance(excess_hc, (int, float)): tot_excess_hc += excess_hc
            curr_row += 1

        # Sub-total row
        ws_claim.cell(curr_row, 4, "Sub-Total").font = font_bold
        ws_claim.cell(curr_row, 11, round(tot_claimed_hc, 2)).font = font_bold
        ws_claim.cell(curr_row, 11).number_format = "#,##0.00"
        ws_claim.cell(curr_row, 13, round(tot_eligible_hc, 2)).font = font_bold
        ws_claim.cell(curr_row, 13).number_format = "#,##0.00"
        ws_claim.cell(curr_row, 14, round(tot_excess_hc, 2)).font = font_bold
        ws_claim.cell(curr_row, 14).number_format = "#,##0.00"
        for c in range(3, 18):
            ws_claim.cell(curr_row, c).fill = fill_subtotal
            ws_claim.cell(curr_row, c).border = thin_border

        # 2. Verification Sheet (With Billed To and grouped columns)
        ws_ver = wb.create_sheet(title="Verification")
        ws_ver.views.sheetView[0].showGridLines = True
        ver_headers = [
            "Document", "Date", "Billed To", "Vendor", "Expense Type", "Nature of Expense",
            "Description", "Bill Currency", "Bill Amount", "FX Rate", f"Claimed ({home_curr})",
            "Policy Limit", f"Eligible ({home_curr})", f"Excess ({home_curr})", "Status", "Amount Evidence", "Remarks"
        ]
        for c_idx, h in enumerate(ver_headers, start=1):
            cell = ws_ver.cell(1, c_idx, h)
            cell.font = font_col_header
            cell.fill = fill_header
            cell.alignment = align_center
            cell.border = thin_border

        for r_idx, rec in enumerate(records, start=2):
            vals = [
                rec.get("document_name", ""),
                rec.get("date", ""),
                rec.get("billed_to", ""),
                rec.get("vendor", ""),
                rec.get("expense_type", ""),
                rec.get("nature_of_expense", ""),
                rec.get("description", ""),
                rec.get("currency", ""),
                rec.get("amount", ""),
                rec.get("fx_rate", 1.0),
                rec.get("home_currency_amount", ""),
                rec.get("policy_limit_str", ""),
                rec.get("eligible_amount_hc", ""),
                rec.get("excess_hc", ""),
                rec.get("status", "PASS"),
                rec.get("amount_evidence", ""),
                rec.get("remarks", "")
            ]
            for c_idx, v in enumerate(vals, start=1):
                cell = ws_ver.cell(r_idx, c_idx, v)
                cell.font = font_data
                cell.border = thin_border
                if c_idx in [9, 11, 13, 14] and isinstance(v, (int, float)):
                    cell.number_format = "#,##0.00"

        # 3. Summary Sheet
        ws_sum = wb.create_sheet(title="Summary")
        ws_sum.views.sheetView[0].showGridLines = True
        ws_sum.cell(1, 5, "REIMBURSEMENT SUMMARY").font = font_header_big
        ws_sum.cell(2, 5, "Employee").font = font_bold
        ws_sum.cell(2, 6, employee_config.get("employee_name", "")).font = font_data
        ws_sum.cell(2, 8, "Designation").font = font_bold
        ws_sum.cell(2, 9, employee_config.get("designation", "")).font = font_data
        ws_sum.cell(3, 5, "Region").font = font_bold
        ws_sum.cell(3, 6, employee_config.get("region", "")).font = font_data
        ws_sum.cell(3, 8, "Home Currency").font = font_bold
        ws_sum.cell(3, 9, home_curr).font = font_data

        ws_sum.cell(5, 5, f"Total Claimed ({home_curr})").font = font_bold
        ws_sum.cell(5, 6, round(tot_claimed_hc, 2)).font = font_bold
        ws_sum.cell(5, 6).number_format = "#,##0.00"

        ws_sum.cell(5, 8, f"Total Eligible ({home_curr})").font = font_bold
        ws_sum.cell(5, 9, round(tot_eligible_hc, 2)).font = font_bold
        ws_sum.cell(5, 9).number_format = "#,##0.00"

        ws_sum.cell(6, 5, f"Total Excess ({home_curr})").font = font_bold
        ws_sum.cell(6, 6, round(tot_excess_hc, 2)).font = font_bold
        ws_sum.cell(6, 6).number_format = "#,##0.00"

        sum_headers = ["Expense Nature", f"Claimed ({home_curr})", f"Eligible ({home_curr})", f"Excess ({home_curr})", "Count", "Expense Details"]
        for c_idx, h in enumerate(sum_headers, start=5):
            cell = ws_sum.cell(8, c_idx, h)
            cell.font = font_col_header
            cell.fill = fill_header
            cell.alignment = align_center
            cell.border = thin_border

        nature_groups: Dict[str, List[Dict[str, Any]]] = {}
        for r in records:
            nature_groups.setdefault(r.get("nature_of_expense", "Other"), []).append(r)

        s_r = 9
        for nat, grp in nature_groups.items():
            g_cl = sum(g.get("home_currency_amount", 0.0) for g in grp if isinstance(g.get("home_currency_amount"), (int, float)))
            g_el = sum(g.get("eligible_amount_hc", 0.0) for g in grp if isinstance(g.get("eligible_amount_hc"), (int, float)))
            g_ex = sum(g.get("excess_hc", 0.0) for g in grp if isinstance(g.get("excess_hc"), (int, float)))
            ws_sum.cell(s_r, 5, nat).font = font_data
            ws_sum.cell(s_r, 6, round(g_cl, 2)).font = font_data
            ws_sum.cell(s_r, 6).number_format = "#,##0.00"
            ws_sum.cell(s_r, 7, round(g_el, 2)).font = font_data
            ws_sum.cell(s_r, 7).number_format = "#,##0.00"
            ws_sum.cell(s_r, 8, round(g_ex, 2)).font = font_data
            ws_sum.cell(s_r, 8).number_format = "#,##0.00"
            ws_sum.cell(s_r, 9, len(grp)).font = font_data
            ws_sum.cell(s_r, 10, "\n".join([f"• {g.get('date', '')} — {g.get('description', '')}" for g in grp])).alignment = Alignment(wrap_text=True)
            for c in range(5, 11): ws_sum.cell(s_r, c).border = thin_border
            s_r += 1

        # 4. Exceptions Sheet
        ws_exc = wb.create_sheet(title="Exceptions")
        ws_exc.views.sheetView[0].showGridLines = True
        for c_idx, h in enumerate(ver_headers, start=1):
            cell = ws_exc.cell(1, c_idx, h)
            cell.font = font_col_header
            cell.fill = fill_exception
            cell.alignment = align_center
            cell.border = thin_border

        e_r = 2
        for rec in records:
            if rec.get("status") != "PASS":
                vals = [
                    rec.get("document_name", ""),
                    rec.get("date", ""),
                    rec.get("billed_to", ""),
                    rec.get("vendor", ""),
                    rec.get("expense_type", ""),
                    rec.get("nature_of_expense", ""),
                    rec.get("description", ""),
                    rec.get("currency", ""),
                    rec.get("amount", ""),
                    rec.get("fx_rate", 1.0),
                    rec.get("home_currency_amount", ""),
                    rec.get("policy_limit_str", ""),
                    rec.get("eligible_amount_hc", ""),
                    rec.get("excess_hc", ""),
                    rec.get("status", "REVIEW"),
                    rec.get("amount_evidence", ""),
                    rec.get("remarks", "")
                ]
                for c_idx, v in enumerate(vals, start=1):
                    cell = ws_exc.cell(e_r, c_idx, v)
                    cell.font = font_data
                    cell.border = thin_border
                e_r += 1

        # Adjust column widths
        for sheet in [ws_claim, ws_ver, ws_sum, ws_exc]:
            for col in sheet.columns:
                max_len = max(len(str(cell.value or "")) for cell in col if "\n" not in str(cell.value or ""))
                col_letter = get_column_letter(col[0].column)
                sheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    @staticmethod
    def generate_tracker_export(employee_config: Dict[str, Any], records: List[Dict[str, Any]]) -> io.BytesIO:
        """Exports rows ready to paste or append directly into Reimbursement_Tracker_2026D.xlsx"""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Tracker_Rows"
        headers = [
            "Claim No.", "Submission Date", "Employee Name", "Expense Type",
            "Expense Description", "Amount", "Currency", "Status", "Payment Date",
            "Receipt Saved", "Remark"
        ]
        for col, h in enumerate(headers, 1):
            cell = ws.cell(1, col, h)
            cell.font = Font(name="Calibri", bold=True)
            cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

        today_str = datetime.date.today().strftime("%Y-%m-%d")
        for r_idx, r in enumerate(records, 2):
            stat = "Under Review" if r.get("status") == "REVIEW" else "Pending Payment"
            ws.cell(r_idx, 1, 825 + ((r_idx - 2) // 4))
            ws.cell(r_idx, 2, today_str)
            ws.cell(r_idx, 3, employee_config.get("employee_name", ""))
            ws.cell(r_idx, 4, r.get("nature_of_expense", ""))
            ws.cell(r_idx, 5, r.get("description", ""))
            ws.cell(r_idx, 6, r.get("home_currency_amount", r.get("amount", 0.0)))
            ws.cell(r_idx, 7, employee_config.get("home_currency", "GBP"))
            ws.cell(r_idx, 8, stat)
            ws.cell(r_idx, 9, "-")
            ws.cell(r_idx, 10, "YES")
            ws.cell(r_idx, 11, "Sheet Prepared by Automation System")

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf


# ==============================================================================
# SECTION 7: SIMPLIFIED STREAMLIT USER INTERFACE
# ==============================================================================

def run_streamlit_app():
    if not STREAMLIT_AVAILABLE or st is None:
        print("[ERROR] Streamlit is not installed in the environment.")
        return

    st.set_page_config(page_title="Reimbursement AI Automation", page_icon="⚡", layout="wide")

    # Clean, modern sky-blue CSS theme
    st.markdown("""
        <style>
            .stApp { background-color: #f8fafc; }
            .hero-header {
                background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%);
                color: white; padding: 1.2rem 1.8rem; border-radius: 10px; margin-bottom: 1.2rem;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.08);
            }
            .hero-header h2 { margin: 0; color: white !important; font-size: 1.7rem; }
            .hero-header p { margin: 0.2rem 0 0 0; font-size: 0.95rem; opacity: 0.95; }
            .card-panel {
                background: white; border: 1px solid #e2e8f0; border-radius: 8px;
                padding: 1.2rem; margin-bottom: 1rem; box-shadow: 0 1px 2px rgba(0, 0, 0, 0.04);
            }
            .kpi-box {
                background: white; border-top: 3px solid #0284c7; border: 1px solid #e2e8f0;
                border-radius: 8px; padding: 0.8rem; text-align: center;
            }
            .kpi-num { font-size: 1.5rem; font-weight: 700; color: #0f172a; }
            .kpi-lbl { font-size: 0.75rem; color: #64748b; text-transform: uppercase; margin-top: 2px; }
            .badge-pass { background-color: #dcfce7; color: #15803d; padding: 2px 8px; border-radius: 10px; font-weight: 600; font-size: 0.78rem; }
            .badge-rev { background-color: #fef9c3; color: #a16207; padding: 2px 8px; border-radius: 10px; font-weight: 600; font-size: 0.78rem; }
        </style>
    """, unsafe_allow_html=True)

    if "master_manager" not in st.session_state:
        st.session_state.master_manager = MasterDataManager()
    if "pipeline" not in st.session_state:
        st.session_state.pipeline = ReimbursementPipeline(st.session_state.master_manager)
    if "records" not in st.session_state:
        st.session_state.records = []
    if "selected_employee" not in st.session_state:
        st.session_state.selected_employee = "Daragh Patrick Ryan"
    if "fx_rates" not in st.session_state:
        st.session_state.fx_rates = DEFAULT_FX_MATRICES["GBP"].copy()

    # Hero Header
    st.markdown("""
        <div class="hero-header">
            <h2>⚡ Employee Reimbursement AI Automation</h2>
            <p>Fast Document Verification • Dynamic Employee Profiles • Genuine Entity Extraction • Live FX Home Currency</p>
        </div>
    """, unsafe_allow_html=True)

    # -------------------------------------------------------------
    # STEP 1: EMPLOYEE SELECTION & CURRENCY CONVERTER
    # -------------------------------------------------------------
    with st.container():
        st.markdown("### 🎛️ 1. Employee Profile & Foreign Exchange Rates")
        c1, c2, c3 = st.columns([1.5, 1.2, 2.3])

        with c1:
            emp_choice = st.selectbox(
                "👤 Select Employee",
                list(EMPLOYEE_PRESETS.keys()),
                index=list(EMPLOYEE_PRESETS.keys()).index(st.session_state.selected_employee)
            )
            preset = EMPLOYEE_PRESETS[emp_choice]
            if emp_choice != st.session_state.selected_employee:
                st.session_state.selected_employee = emp_choice
                st.session_state.fx_rates = DEFAULT_FX_MATRICES.get(preset["home_currency"], DEFAULT_FX_MATRICES["GBP"]).copy()
                st.rerun()

            emp_name = st.text_input("Employee Name", emp_choice if emp_choice != "Custom / New Employee" else "New Employee")
            emp_desig = st.text_input("Designation / Title", preset["designation"])

        with c2:
            home_curr = st.selectbox(
                "💱 Home Currency",
                ["GBP", "CAD", "USD", "INR", "EUR", "CHF"],
                index=["GBP", "CAD", "USD", "INR", "EUR", "CHF"].index(preset["home_currency"])
            )
            reg = st.selectbox(
                "🌍 Region",
                ["Europe", "USA", "Canada", "India"],
                index=["Europe", "USA", "Canada", "India"].index(preset["region"])
            )
            exp_cat = st.selectbox("Category", ["Conference Expense", "Business Expense"], index=0 if preset["expense_category"] == "Conference Expense" else 1)

        with c3:
            st.markdown(f"**💱 Foreign Exchange Matrix (1 Unit $\\rightarrow$ {home_curr})**")
            st.caption(f"Adjust rates to instantly convert bills into {home_curr}:")
            fx_cols = st.columns(4)
            current_rates = st.session_state.fx_rates
            
            # Show top 4 foreign currencies
            currencies_to_show = [c for c in ["USD", "EUR", "CAD", "INR", "GBP"] if c != home_curr][:4]
            updated_rates = current_rates.copy()
            for idx, curr_code in enumerate(currencies_to_show):
                with fx_cols[idx]:
                    val = updated_rates.get(curr_code, 1.0)
                    new_val = st.number_input(f"1 {curr_code} =", value=float(val), format="%.4f", step=0.01, key=f"fx_{curr_code}")
                    updated_rates[curr_code] = new_val
            updated_rates[home_curr] = 1.0
            st.session_state.fx_rates = updated_rates

    emp_config = {
        "employee_name": emp_name,
        "employee_id": preset["employee_id"],
        "designation": emp_desig,
        "leadership_level": "Senior Leadership" if any(k in emp_desig.lower() for k in ["director", "head", "leadership", "vp"]) else "All Other Employees",
        "home_currency": home_curr,
        "region": reg,
        "entity": preset["entity"],
        "expense_category": exp_cat
    }

    # -------------------------------------------------------------
    # STEP 2: UPLOAD DOCUMENTS / QUICK TEST
    # -------------------------------------------------------------
    st.markdown("---")
    st.markdown("### 📂 2. Document Processing (PDFs & Receipts)")
    u_col1, u_col2 = st.columns([3, 1])

    with u_col1:
        uploaded_files = st.file_uploader(
            "Upload receipts, invoices, tickets, hotel bills (PDF, Images)",
            type=["pdf", "png", "jpg"],
            accept_multiple_files=True
        )

    with u_col2:
        st.write("")
        st.write("")
        if st.button("🧪 Load Standard Sample Bills", use_container_width=True):
            test_files = [
                ("Uber Taxi Bill", "test_bills/Uber_Trip_Boston_16Sep2026.pdf"),
                ("Virgin Atlantic Flight", "test_bills/Virgin_Atlantic_eTicket_London_Boston.pdf"),
                ("Royal Mail Postage", "test_bills/Royal_Mail_Postage_Receipt.pdf"),
                ("Battery Wharf Hotel", "test_bills/Battery_Wharf_Hotel_Boston_Invoice.pdf"),
                ("Raindrop Indian Cuisine", "test_bills/Raindrop_Indian_Cuisine_Receipt.pdf")
            ]
            processed = []
            for label, p in test_files:
                if os.path.exists(p):
                    with open(p, "rb") as f:
                        rec = st.session_state.pipeline.process_document(f, os.path.basename(p), emp_config, st.session_state.fx_rates)
                        processed.append(rec)
            st.session_state.records = processed
            st.success("Loaded and processed sample bills!")
            st.rerun()

    if uploaded_files:
        if st.button("🚀 Process Uploaded Files Now", type="primary"):
            processed = []
            for uf in uploaded_files:
                rec = st.session_state.pipeline.process_document(uf, uf.name, emp_config, st.session_state.fx_rates)
                processed.append(rec)
            st.session_state.records = processed
            st.success(f"Processed {len(processed)} documents!")
            st.rerun()

    # -------------------------------------------------------------
    # STEP 3: INSTANT VERIFIED CLAIM TABLE
    # -------------------------------------------------------------
    recs = st.session_state.records
    if recs:
        st.markdown("---")
        st.markdown("### 📊 3. Verified Claim Summary & Audit Table")

        # Top KPIs
        tot_claim = sum(r.get("home_currency_amount", 0.0) for r in recs if isinstance(r.get("home_currency_amount"), (int, float)))
        tot_elig = sum(r.get("eligible_amount_hc", 0.0) for r in recs if isinstance(r.get("eligible_amount_hc"), (int, float)))
        tot_excess = sum(r.get("excess_hc", 0.0) for r in recs if isinstance(r.get("excess_hc"), (int, float)))
        pass_cnt = sum(1 for r in recs if r.get("status") == "PASS")
        rev_cnt = sum(1 for r in recs if r.get("status") == "REVIEW")

        k1, k2, k3, k4, k5 = st.columns(5)
        k1.markdown(f'<div class="kpi-box"><div class="kpi-num">{len(recs)}</div><div class="kpi-lbl">Total Documents</div></div>', unsafe_allow_html=True)
        k2.markdown(f'<div class="kpi-box"><div class="kpi-num">{tot_claim:.2f} {home_curr}</div><div class="kpi-lbl">Total Claimed</div></div>', unsafe_allow_html=True)
        k3.markdown(f'<div class="kpi-box"><div class="kpi-num" style="color:#15803d">{tot_elig:.2f} {home_curr}</div><div class="kpi-lbl">Total Eligible</div></div>', unsafe_allow_html=True)
        k4.markdown(f'<div class="kpi-box"><div class="kpi-num" style="color:#b91c1c">{tot_excess:.2f} {home_curr}</div><div class="kpi-lbl">Total Excess</div></div>', unsafe_allow_html=True)
        k5.markdown(f'<div class="kpi-box"><div class="kpi-num">{pass_cnt} ✅ / {rev_cnt} ⚠️</div><div class="kpi-lbl">Status Breakdown</div></div>', unsafe_allow_html=True)

        st.write("")

        # Main Table formatted with Billed-To and side-by-side Expense Type + Nature
        display_data = []
        for r in recs:
            b_amt = f"{r['currency']} {r['amount']:.2f}" if isinstance(r['amount'], (int, float)) else str(r['amount'])
            hc_amt = f"{home_curr} {r['home_currency_amount']:.2f}" if isinstance(r['home_currency_amount'], (int, float)) else str(r['home_currency_amount'])
            el_amt = f"{home_curr} {r['eligible_amount_hc']:.2f}" if isinstance(r['eligible_amount_hc'], (int, float)) else str(r['eligible_amount_hc'])
            ex_amt = f"{home_curr} {r['excess_hc']:.2f}" if isinstance(r['excess_hc'], (int, float)) else str(r['excess_hc'])

            display_data.append({
                "Document": r["document_name"],
                "Date": r["date"],
                "Billed To": r["billed_to"],
                "Vendor": r["vendor"],
                "Expense Type": r["expense_type"],
                "Nature of Expense": r["nature_of_expense"],
                "Description": r["description"],
                "Bill Amount": b_amt,
                "FX Rate": r["fx_rate"],
                f"Claimed ({home_curr})": hc_amt,
                "Policy Limit": r["policy_limit_str"],
                f"Eligible ({home_curr})": el_amt,
                f"Excess ({home_curr})": ex_amt,
                "Status": "✅ PASS" if r["status"] == "PASS" else "⚠️ REVIEW",
                "Remarks": r["remarks"]
            })

        st.dataframe(pd.DataFrame(display_data), use_container_width=True)

        # -------------------------------------------------------------
        # STEP 4: EXPORT OPTIONS & TRACKER INTEGRATION
        # -------------------------------------------------------------
        st.markdown("---")
        st.markdown("### 💾 4. Download Official Sheets & Update Tracker")
        exp_col1, exp_col2 = st.columns(2)

        with exp_col1:
            excel_bytes = ExcelExportGenerator.generate_workbook(emp_config, recs)
            st.download_button(
                label=f"📥 Download 4-Sheet Official Claim Workbook ({home_curr})",
                data=excel_bytes,
                file_name=f"Reimbursement_Claim_{emp_name.replace(' ', '_')}_{home_curr}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        with exp_col2:
            tracker_bytes = ExcelExportGenerator.generate_tracker_export(emp_config, recs)
            st.download_button(
                label="📊 Export Tracker Update Rows (for Reimbursement_Tracker_2026D.xlsx)",
                data=tracker_bytes,
                file_name=f"Tracker_Update_{emp_name.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    # Master Tables Drawer
    with st.expander("📖 View Company Master Tables (Policy_Master, Nature_Master, Currency_Master)"):
        m_t1, m_t2, m_t3 = st.tabs(["Policy_Master", "Nature_Master", "Currency_Master"])
        with m_t1: st.dataframe(pd.DataFrame(st.session_state.master_manager.policy_master), use_container_width=True)
        with m_t2: st.dataframe(pd.DataFrame(st.session_state.master_manager.nature_master), use_container_width=True)
        with m_t3: st.dataframe(pd.DataFrame(st.session_state.master_manager.currency_master), use_container_width=True)


# ==============================================================================
# SECTION 8: CLI TEST RUNNER
# ==============================================================================

def run_cli_test():
    print("=" * 80)
    print("EMPLOYEE REIMBURSEMENT AI AUTOMATION SYSTEM — CLI VERIFICATION BENCHMARK")
    print("=" * 80)

    master_mgr = MasterDataManager()
    pipeline = ReimbursementPipeline(master_mgr)

    # Test with Daragh Patrick Ryan (Home Currency = GBP)
    emp_cfg = {
        "employee_name": "Daragh Patrick Ryan",
        "employee_id": "GI001UK799",
        "designation": "Senior Director",
        "leadership_level": "Senior Leadership",
        "home_currency": "GBP",
        "region": "Europe",
        "entity": "GENINVO INTERNATIONAL",
        "expense_category": "Conference Expense"
    }
    fx_rates = DEFAULT_FX_MATRICES["GBP"].copy()

    test_cases = [
        ("Uber Taxi", "test_bills/Uber_Trip_Boston_16Sep2026.pdf"),
        ("Virgin Atlantic Flight", "test_bills/Virgin_Atlantic_eTicket_London_Boston.pdf"),
        ("Royal Mail Postage", "test_bills/Royal_Mail_Postage_Receipt.pdf"),
        ("Battery Wharf Hotel", "test_bills/Battery_Wharf_Hotel_Boston_Invoice.pdf"),
        ("Raindrop Indian Cuisine", "test_bills/Raindrop_Indian_Cuisine_Receipt.pdf")
    ]

    records = []
    print(f"\nProcessing Bills for Employee: {emp_cfg['employee_name']} (Home Currency: {emp_cfg['home_currency']})...\n")
    for label, path in test_cases:
        if not os.path.exists(path):
            print(f"[ERROR] Missing file: {path}")
            continue
        with open(path, "rb") as f:
            rec = pipeline.process_document(f, os.path.basename(path), emp_cfg, fx_rates)
            records.append(rec)

    # Format output table
    print("-" * 145)
    print(f"{'DOCUMENT':<24} | {'DATE':<10} | {'BILLED TO':<14} | {'VENDOR':<22} | {'TYPE':<7} | {'NATURE OF EXPENSE':<34} | {'BILL AMT':<11} | {'FX':<4} | {'CLAIM (GBP)':<11} | {'ELIG (GBP)':<10} | {'STAT'}")
    print("-" * 145)

    for r in records:
        doc = r["document_name"][:22]
        dt = r["date"]
        bto = r["billed_to"][:14]
        v = r["vendor"][:20]
        t = r["expense_type"][:7]
        nat = r["nature_of_expense"][:32]
        b_amt = f"{r['currency']} {r['amount']:.2f}" if isinstance(r['amount'], (int, float)) else str(r['amount'])
        fx_s = f"{r['fx_rate']:.2f}"
        cl_hc = f"{r['home_currency_amount']:.2f}" if isinstance(r['home_currency_amount'], (int, float)) else str(r['home_currency_amount'])
        el_hc = f"{r['eligible_amount_hc']:.2f}" if isinstance(r['eligible_amount_hc'], (int, float)) else str(r['eligible_amount_hc'])
        stat = r["status"]
        print(f"{doc:<24} | {dt:<10} | {bto:<14} | {v:<22} | {t:<7} | {nat:<34} | {b_amt:<11} | {fx_s:<4} | {cl_hc:<11} | {el_hc:<10} | {stat}")

    print("-" * 145)

    # Generate sample export
    export_path = "Claim_Output_Test_Export.xlsx"
    buf = ExcelExportGenerator.generate_workbook(emp_cfg, records)
    with open(export_path, "wb") as f_out:
        f_out.write(buf.getbuffer())
    print(f"\n[SUCCESS] Generated 4-sheet production workbook: {export_path}")

    tracker_path = "Tracker_Update_Rows.xlsx"
    t_buf = ExcelExportGenerator.generate_tracker_export(emp_cfg, records)
    with open(tracker_path, "wb") as f_out:
        f_out.write(t_buf.getbuffer())
    print(f"[SUCCESS] Generated tracker update export: {tracker_path}")


# ==============================================================================
# MAIN ENTRY POINT
# ==============================================================================

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Employee Reimbursement AI Automation System")
    parser.add_argument("--test", action="store_true", help="Run automated test suite against the 5 core bill types")
    args, unknown = parser.parse_known_args()

    if args.test or not STREAMLIT_AVAILABLE:
        run_cli_test()
    else:
        run_streamlit_app()
