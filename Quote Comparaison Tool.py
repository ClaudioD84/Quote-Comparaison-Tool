"""
AI-Powered Fleet Leasing Offer Comparator - Streamlit App
This version uses a Large Language Model (LLM) to intelligently parse PDF content,
populates a structured Excel template, and automatically names the output file.

Author: Fleet Management Tool
Version: 2.0

Requirements:
  streamlit, pandas, numpy, pdfplumber, python-dateutil, xlsxwriter, openpyxl

Key Changes in this version:
  - Added 'openpyxl' to requirements for reading and modifying the Excel template
    while preserving its structure and formatting.
  - The script now requires an Excel template file to be uploaded.
  - A new function `generate_structured_report` populates the template with parsed data.
  - The output filename is automatically generated based on the extracted
    company and driver name, e.g., 'Grundfos_Mikkel_Mikkelsen.xlsx'.
  - The `ParsedOffer` data structure and the mock LLM call have been expanded
    to extract more detailed information required by the template.
"""

import io
import re
import sys
import logging
import tempfile
import json
import traceback
from typing import List, Dict, Any, Optional, Tuple, Union
from dataclasses import dataclass, field, asdict
from datetime import datetime, date
import requests
import difflib
from collections import defaultdict

import streamlit as st
import pandas as pd
import numpy as np
import pdfplumber
from dateutil import parser as dateparser
import openpyxl  # Added for template manipulation

# --- Constants ---
# This map links the attributes of the ParsedOffer dataclass to the exact text
# found in column B of the Excel template. This allows the script to find the
# correct row to write each piece of data to.
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
    """Sets up a Streamlit-friendly logger."""
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
    """
    Standardized structure for parsed leasing offer data.
    Expanded to include more fields for the structured template and file naming.
    """
    # Core attributes for identification and file naming
    filename: str
    vendor: Optional[str] = None
    company_name: Optional[str] = None
    driver_name: Optional[str] = None
    quote_number: Optional[str] = None

    # Vehicle details
    vehicle_description: Optional[str] = None
    manufacturer: Optional[str] = None
    model: Optional[str] = None
    version: Optional[str] = None
    fuel_type: Optional[str] = None

    # Contract terms
    term_months: Optional[int] = None
    mileage_km_per_year: Optional[int] = None

    # Cost breakdown
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

    # Parsing metadata
    parsing_confidence: float = 0.0
    warnings: List[str] = field(default_factory=list)
    raw_text: str = ""

# --- Core Logic: PDF Parsing & Data Extraction ---

def extract_text_from_pdf(file_content: bytes) -> str:
    """Extracts text from all pages of a PDF file."""
    text = ""
    with pdfplumber.open(io.BytesIO(file_content)) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += f"--- PAGE {page.page_number} ---\n\n{page_text}\n\n"
    logger.info(f"Extracted {len(text)} characters from PDF.")
    return text

def mock_llm_api_call(pdf_text: str, filename: str) -> Dict[str, Any]:
    """
    Mocks a call to an LLM API for data extraction.
    This function simulates finding data in the PDF text and returns a JSON object.
    It has been updated to return the new, more detailed fields.
    """
    logger.info(f"Mocking LLM call for: {filename}")
    
    # Simulate extraction based on vendor-specific keywords
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
            "monthly_financial_rate": 4310.27, # depreciation + interest
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
        # Return a default/empty structure if vendor is not recognized
        return {
            "vendor": "Unknown", "company_name": "Unknown", "driver_name": "Unknown",
            "cost_per_month": 0, "term_months": 0, "mileage_km_per_year": 0
        }

def call_llm_for_parsing(pdf_text: str, filename: str) -> Dict[str, Any]:
    """
    Wrapper for the LLM API call.
    In a real application, this would contain the actual API request logic.
    """
    try:
        # Replace this with your actual API call, e.g., to Gemini
        response_data = mock_llm_api_call(pdf_text, filename)
        return response_data
    except Exception as e:
        logger.error(f"Error calling LLM API: {e}")
        return {}

def process_uploaded_files(uploaded_files: List[Any]) -> List[ParsedOffer]:
    """Processes a list of uploaded PDF files."""
    parsed_offers = []
    for uploaded_file in uploaded_files:
        try:
            logger.info(f"Processing file: {uploaded_file.name}")
            file_content = uploaded_file.getvalue()
            raw_text = extract_text_from_pdf(file_content)
            
            llm_result = call_llm_for_parsing(raw_text, uploaded_file.name)
            
            if not llm_result:
                offer = ParsedOffer(filename=uploaded_file.name, warnings=["LLM parsing failed."])
                parsed_offers.append(offer)
                continue

            # Create a ParsedOffer object from the LLM result
            offer = ParsedOffer(
                filename=uploaded_file.name,
                raw_text=raw_text,
                **llm_result
            )
            
            # Basic validation
            if not all([offer.vendor, offer.cost_per_month, offer.term_months]):
                offer.warnings.append("Core fields (vendor, cost, term) are missing.")
            else:
                offer.parsing_confidence = 0.95 # Assume high confidence for mock
                if not offer.total_contract_cost:
                    offer.total_contract_cost = offer.cost_per_month * offer.term_months

            parsed_offers.append(offer)

        except Exception as e:
            logger.error(f"Failed to process {uploaded_file.name}: {e}")
            traceback.print_exc()
            parsed_offers.append(ParsedOffer(filename=uploaded_file.name, warnings=[f"An error occurred: {e}"]))
            
    return parsed_offers

# --- Core Logic: Excel Report Generation ---

def generate_structured_report(offers: List[ParsedOffer], template_file: Any) -> io.BytesIO:
    """
    Populates a structured Excel template with data from parsed offers.

    Args:
        offers: A list of ParsedOffer objects.
        template_file: The uploaded Excel template file object.

    Returns:
        A BytesIO buffer containing the populated Excel file.
    """
    try:
        # Load the workbook and select the active sheet
        template_file.seek(0)
        workbook = openpyxl.load_workbook(template_file)
        sheet = workbook.active
        logger.info(f"Loaded template workbook. Active sheet: '{sheet.title}'")

        # 1. Map vendor names to their respective column index in the template
        vendor_to_col_map = {}
        # Vendors are typically listed in row 2, starting from column C (index 3)
        for col_idx in range(3, sheet.max_column + 1):
            cell_value = sheet.cell(row=2, column=col_idx).value
            if cell_value:
                # Find which offer vendor best matches this column header
                for offer in offers:
                    # Use fuzzy matching to handle minor name differences (e.g., Ayvens vs AYVENS)
                    matches = difflib.get_close_matches(offer.vendor.lower(), [str(cell_value).lower()], n=1, cutoff=0.8)
                    if matches:
                        vendor_to_col_map[offer.vendor] = col_idx
                        logger.info(f"Mapped vendor '{offer.vendor}' to template column {col_idx}")
                        break # Move to the next column once a match is found
        
        # 2. Map field descriptions from the template (Column B) to their row index
        description_to_row_map = {}
        for row_idx in range(1, sheet.max_row + 1):
            cell_value = sheet.cell(row=row_idx, column=2).value # Column B
            if cell_value and isinstance(cell_value, str):
                description_to_row_map[cell_value.strip()] = row_idx

        # 3. Populate the sheet with data from each offer
        for offer in offers:
            col_to_write = vendor_to_col_map.get(offer.vendor)
            if not col_to_write:
                logger.warning(f"Could not find a matching column for vendor '{offer.vendor}' in the template.")
                continue

            # Use the predefined map to write each field to the correct row
            for field_name, row_description in FIELD_TO_ROW_MAP.items():
                row_to_write = description_to_row_map.get(row_description)
                value = getattr(offer, field_name, None)

                if row_to_write and value is not None:
                    try:
                        sheet.cell(row=row_to_write, column=col_to_write, value=value)
                        logger.info(f"Writing '{value}' to cell({row_to_write}, {col_to_write}) for {offer.vendor}")
                    except Exception as cell_error:
                        logger.error(f"Could not write value '{value}' to cell. Error: {cell_error}")

        # Save the populated workbook to a memory buffer
        buffer = io.BytesIO()
        workbook.save(buffer)
        buffer.seek(0)
        return buffer

    except Exception as e:
        logger.error(f"An error occurred during Excel report generation: {e}")
        traceback.print_exc()
        # Return an empty buffer in case of error
        return io.BytesIO()

# --- Streamlit UI ---

def display_parsing_results(offers: List[ParsedOffer]):
    """Display parsing results summary"""
    st.header("📊 Parsing Results")
    col1, col2 = st.columns(2)
    with col1:
        avg_confidence = np.mean([o.parsing_confidence for o in offers]) if offers else 0
        st.metric("Average Confidence", f"{avg_confidence:.1%}")
    with col2:
        warning_count = sum(len(o.warnings) for o in offers)
        st.metric("Total Warnings", warning_count)
    
    with st.expander("📋 Detailed Parsing Results"):
        for offer in offers:
            st.write(f"**{offer.vendor or offer.filename}**")
            st.json(asdict(offer), expanded=False)

def main():
    """Main function to run the Streamlit app."""
    st.set_page_config(page_title="Fleet Leasing Comparator", layout="wide")
    st.title("🤖 AI-Powered Fleet Leasing Comparator")
    st.write("This tool extracts data from PDF leasing offers, compares them, and generates a structured TCO report.")

    # --- Sidebar for File Uploads and Actions ---
    with st.sidebar:
        st.header("1. Upload Offer PDFs")
        uploaded_files = st.file_uploader(
            "Upload one or more PDF files from leasing companies",
            type=["pdf"],
            accept_multiple_files=True
        )

        st.header("2. Upload TCO Template")
        template_file = st.file_uploader(
            "Upload the Excel TCO template file",
            type=["xlsx"]
        )

        process_button = st.button("🚀 Process Files", type="primary", use_container_width=True, disabled=not (uploaded_files and template_file))

    # --- Main Content Area ---
    if process_button:
        with st.spinner("Analyzing documents and generating report..."):
            parsed_offers = process_uploaded_files(uploaded_files)

        if parsed_offers:
            st.success("✅ Processing complete!")
            display_parsing_results(parsed_offers)

            # --- Generate Report and Prepare Download ---
            report_buffer = generate_structured_report(parsed_offers, template_file)
            
            # Determine filename from the first valid offer
            first_offer = next((o for o in parsed_offers if o.company_name and o.driver_name), None)
            if first_offer:
                company = first_offer.company_name.replace(' ', '_')
                driver = first_offer.driver_name.replace(' ', '_')
                output_filename = f"{company}_{driver}.xlsx"
            else:
                output_filename = "Leasing_Report.xlsx"

            with st.sidebar:
                st.header("✅ Report Ready")
                st.download_button(
                    label="📥 Download Report",
                    data=report_buffer,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.error("Could not parse any of the uploaded files. Please check the files and try again.")
    else:
        st.info("Please upload your PDF offers and the Excel template, then click 'Process Files'.")

if __name__ == "__main__":
    main()
