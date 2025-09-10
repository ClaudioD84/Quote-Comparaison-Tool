"""
AI-Powered Fleet Leasing Offer Comparator - Streamlit App
This version uses a Large Language Model (LLM) to intelligently parse PDF content,
populates a structured Excel template, and automatically names the output file.

Author: Fleet Management Tool
Version: 2.1 (with fixes for data writing & winner detection)

Requirements:
  streamlit, pandas, numpy, pdfplumber, python-dateutil, xlsxwriter, openpyxl
"""

import io
import re
import sys
import logging
import json
import traceback
from typing import List, Dict, Any, Optional
from dataclasses import dataclass, field, asdict

import streamlit as st
import pandas as pd
import numpy as np
import pdfplumber
from dateutil import parser as dateparser
import openpyxl
import difflib

# --- Constants ---
FIELD_TO_ROW_MAP = {
    'quote_number': 'Quote number',
    'driver_name': 'Driver name',
    'manufacturer': 'Manufacturer',
    'model': 'Model',
    'version': 'Version',
    'fuel_type': 'Fuel type',
    'term_months': 'Term (months)',
    'mileage_km_per_year': 'Mileage per year (in km)',
    'vehicle_list_price': 'Vehicle list price (excl. VAT, excl. options)',
    'options_price': 'Options (excl. taxes)',
    'delivery_fee': 'Delivery fee',
    'total_net_investment': 'Total net investment',
    'taxation_value': 'Taxation value',
    'monthly_financial_rate': 'Monthly financial rate (depreciation + interest)',
    'maintenance_repairs_tires': 'Maintenance, repairs and tires',
    'insurance': 'Insurance',
    'administration_fee': 'Administration fee',
    'cost_per_month': 'Leasing payment',
}

# --- Logging Setup ---
@st.cache_resource
def setup_logging():
    logger = logging.getLogger("leasing_comparator")
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        handler = logging.StreamHandler(sys.stdout)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    return logger

logger = setup_logging()

# --- Data Structures ---
@dataclass
class ParsedOffer:
    filename: str
    vendor: Optional[str] = None
    company_name: Optional[str] = None
    driver_name: Optional[str] = None
    quote_number: Optional[str] = None

    vehicle_description: Optional[str] = None
    manufacturer: Optional[str] = None
    model: Optional[str] = None
    version: Optional[str] = None
    fuel_type: Optional[str] = None

    term_months: Optional[int] = None
    mileage_km_per_year: Optional[int] = None

    vehicle_list_price: Optional[float] = None
    options_price: Optional[float] = None
    delivery_fee: Optional[float] = None
    total_net_investment: Optional[float] = None
    taxation_value: Optional[float] = None
    monthly_financial_rate: Optional[float] = None
    maintenance_repairs_tires: Optional[float] = None
    insurance: Optional[float] = None
    administration_fee: Optional[float] = None
    cost_per_month: Optional[float] = None
    total_contract_cost: Optional[float] = None

    parsing_confidence: float = 0.0
    warnings: List[str] = field(default_factory=list)
    raw_text: str = ""

# --- PDF Parsing ---
def extract_text_from_pdf(file_content: bytes) -> str:
    text = ""
    with pdfplumber.open(io.BytesIO(file_content)) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += f"--- PAGE {page.page_number} ---\n\n{page_text}\n\n"
    logger.info(f"Extracted {len(text)} characters from PDF.")
    return text

# --- Mock LLM ---
def mock_llm_api_call(pdf_text: str, filename: str) -> Dict[str, Any]:
    logger.info(f"Mocking LLM call for: {filename}")
    if "ARVAL" in pdf_text.upper():
        return {
            "vendor": "Arval",
            "company_name": "Grundfos",
            "driver_name": "Mikkel Mikkelsen",
            "quote_number": "2508.120.036",
            "vehicle_description": "Opel Grandland EL 210 73kWh F GS Sky 5d",
            "manufacturer": "Opel",
            "model": "Grandland EL 210",
            "version": "73kWh F GS Sky 5d",
            "fuel_type": "EV",
            "term_months": 48,
            "mileage_km_per_year": 35000,
            "cost_per_month": 5784.39,
            "total_contract_cost": 277650.92,
            "vehicle_list_price": 260408.00,
            "options_price": 6000.00,
            "delivery_fee": 3820.00,
            "total_net_investment": 284928.00,
            "taxation_value": 350000,
            "monthly_financial_rate": 4310.27,
            "maintenance_repairs_tires": 782.72,
            "insurance": 318.80,
            "administration_fee": 65.00
        }
    elif "AYVENS" in pdf_text.upper():
        return {
            "vendor": "Ayvens",
            "company_name": "Grundfos",
            "driver_name": "Mikkel Mikkelsen",
            "quote_number": "3052514/001",
            "vehicle_description": "OPEL GRANDLAND EL 210 73kWh F GS Sky",
            "manufacturer": "Opel",
            "model": "GRANDLAND EL 210",
            "version": "73kWh F GS Sky",
            "fuel_type": "EV",
            "term_months": 48,
            "mileage_km_per_year": 35000,
            "cost_per_month": 5871.39,
            "total_contract_cost": 281826.72,
            "vehicle_list_price": 274408.00,
            "options_price": 16000.00,
            "delivery_fee": 3820.00,
            "total_net_investment": 291528.00,
            "taxation_value": 347490,
            "monthly_financial_rate": 4075.97,
            "maintenance_repairs_tires": 800.00,
            "insurance": 350.00,
            "administration_fee": 70.00
        }
    else:
        return {
            "vendor": "Unknown", "company_name": "Unknown", "driver_name": "Unknown",
            "cost_per_month": 0, "term_months": 0, "mileage_km_per_year": 0
        }

def call_llm_for_parsing(pdf_text: str, filename: str) -> Dict[str, Any]:
    try:
        return mock_llm_api_call(pdf_text, filename)
    except Exception as e:
        logger.error(f"Error calling LLM API: {e}")
        return {}

# --- File Processing ---
def process_uploaded_files(uploaded_files: List[Any]) -> List[ParsedOffer]:
    parsed_offers = []
    for uploaded_file in uploaded_files:
        try:
            logger.info(f"Processing file: {uploaded_file.name}")
            file_content = uploaded_file.getvalue()
            raw_text = extract_text_from_pdf(file_content)
            llm_result = call_llm_for_parsing(raw_text, uploaded_file.name)

            if not llm_result:
                parsed_offers.append(ParsedOffer(filename=uploaded_file.name, warnings=["LLM parsing failed."]))
                continue

            offer = ParsedOffer(filename=uploaded_file.name, raw_text=raw_text, **llm_result)
            if not all([offer.vendor, offer.cost_per_month, offer.term_months]):
                offer.warnings.append("Core fields (vendor, cost, term) are missing.")
            else:
                offer.parsing_confidence = 0.95
                if not offer.total_contract_cost and offer.cost_per_month and offer.term_months:
                    offer.total_contract_cost = offer.cost_per_month * offer.term_months
            parsed_offers.append(offer)

        except Exception as e:
            logger.error(f"Failed to process {uploaded_file.name}: {e}")
            traceback.print_exc()
            parsed_offers.append(ParsedOffer(filename=uploaded_file.name, warnings=[f"Error: {e}"]))
    return parsed_offers

# --- Report Generation ---
def generate_structured_report(offers: List[ParsedOffer], template_file: Any) -> io.BytesIO:
    """
    Populates the Excel template with parsed offers.
    - Each offer gets its own column with the vendor name as header.
    - Fields are written into the correct rows based on FIELD_TO_ROW_MAP.
    - Winning vendor is written into template or a new 'Summary' sheet.
    """
    try:
        workbook = openpyxl.load_workbook(io.BytesIO(template_file.getvalue()))
        sheet = workbook.active
        logger.info(f"Loaded template workbook. Active sheet: '{sheet.title}'")

        # --- 1. Identify the header row (row containing 'Quote number' etc in column B) ---
        header_row = None
        for r in range(1, sheet.max_row + 1):
            val = sheet.cell(row=r, column=2).value
            if val and str(val).strip().lower() in [s.lower() for s in FIELD_TO_ROW_MAP.values()]:
                header_row = r - 1  # assume headers are the row above the field names
                break
        if not header_row:
            header_row = 2  # fallback

        # --- 2. Ensure vendor columns exist or add them ---
        vendor_to_col_map = {}
        existing_headers = {str(sheet.cell(row=header_row, column=c).value).strip(): c
                            for c in range(3, sheet.max_column + 1)
                            if sheet.cell(row=header_row, column=c).value}

        for offer in offers:
            vendor_name = offer.vendor or offer.filename
            if vendor_name in existing_headers:
                vendor_to_col_map[vendor_name] = existing_headers[vendor_name]
            else:
                # Add new column at the end
                new_col = sheet.max_column + 1
                sheet.cell(row=header_row, column=new_col, value=vendor_name)
                vendor_to_col_map[vendor_name] = new_col
                logger.info(f"Added new column {new_col} for vendor '{vendor_name}'")

        # --- 3. Map descriptions (column B) to rows ---
        description_to_row_map = {}
        for row_idx in range(1, sheet.max_row + 1):
            cell_value = sheet.cell(row=row_idx, column=2).value
            if cell_value and isinstance(cell_value, str):
                description_to_row_map[cell_value.strip()] = row_idx

        # --- 4. Write each offer into its vendor column ---
        for offer in offers:
            col_to_write = vendor_to_col_map.get(offer.vendor or offer.filename)
            if not col_to_write:
                continue
            for field_name, row_description in FIELD_TO_ROW_MAP.items():
                row_to_write = description_to_row_map.get(row_description)
                value = getattr(offer, field_name, None)
                if row_to_write and value is not None:
                    sheet.cell(row=row_to_write, column=col_to_write, value=value)

        # --- 5. Add winning vendor ---
        valid_offers = [o for o in offers if o.cost_per_month and o.cost_per_month > 0]
        if valid_offers:
            winner_offer = min(valid_offers, key=lambda o: float(o.cost_per_month))
            # Try to write in template
            winner_row = None
            for desc, r in description_to_row_map.items():
                if "winning" in desc.lower() or "winner" in desc.lower():
                    winner_row = r
                    break
            if winner_row:
                sheet.cell(row=winner_row, column=3, value=winner_offer.vendor)
            else:
                summary = workbook.create_sheet("Summary") if "Summary" not in workbook.sheetnames else workbook["Summary"]
                summary.append(["Winning leasing company", winner_offer.vendor])

        # --- Save and return buffer ---
        buffer = io.BytesIO()
        workbook.save(buffer)
        buffer.seek(0)
        return buffer

    except Exception as e:
        logger.error(f"Excel generation failed: {e}")
        traceback.print_exc()
        return io.BytesIO()


# --- UI ---
def display_parsing_results(offers: List[ParsedOffer]):
    st.header("📊 Parsing Results")
    col1, col2 = st.columns(2)
    with col1:
        avg_confidence = np.mean([o.parsing_confidence for o in offers]) if offers else 0
        st.metric("Average Confidence", f"{avg_confidence:.1%}")
    with col2:
        st.metric("Total Warnings", sum(len(o.warnings) for o in offers))

    with st.expander("📋 Detailed Parsing Results"):
        for offer in offers:
            st.write(f"**{offer.vendor or offer.filename}**")
            st.json(asdict(offer), expanded=False)

def main():
    st.set_page_config(page_title="Fleet Leasing Comparator", layout="wide")
    st.title("🤖 AI-Powered Fleet Leasing Comparator")
    st.write("Upload PDF offers and an Excel template, then generate a structured TCO report.")

    with st.sidebar:
        st.header("1. Upload Offer PDFs")
        uploaded_files = st.file_uploader("Upload PDF files", type=["pdf"], accept_multiple_files=True)
        st.header("2. Upload TCO Template")
        template_file = st.file_uploader("Upload Excel template", type=["xlsx"])
        process_button = st.button("🚀 Process Files", type="primary", use_container_width=True,
                                   disabled=not (uploaded_files and template_file))

    if process_button:
        with st.spinner("Analyzing documents..."):
            parsed_offers = process_uploaded_files(uploaded_files)
        if parsed_offers:
            st.success("✅ Processing complete!")
            display_parsing_results(parsed_offers)

            report_buffer = generate_structured_report(parsed_offers, template_file)
            first_offer = next((o for o in parsed_offers if o.company_name and o.driver_name), None)
            output_filename = (f"{first_offer.company_name}_{first_offer.driver_name}.xlsx"
                               if first_offer else "Leasing_Report.xlsx")

            with st.sidebar:
                st.header("✅ Report Ready")
                st.download_button(
                    label="📥 Download Report",
                    data=report_buffer.getvalue(),
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.error("No offers parsed successfully.")
    else:
        st.info("Upload offers and template, then click Process.")

if __name__ == "__main__":
    main()
