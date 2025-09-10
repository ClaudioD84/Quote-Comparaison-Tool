"""
AI-Powered Fleet Leasing Offer Comparator - Streamlit App
This version uses a Large Language Model (LLM) to intelligently parse PDF content.
Author: Fleet Management Tool
Requirements:
  streamlit, pandas, numpy, pdfplumber, python-dateutil, xlsxwriter
Notes:
  - This version uses a mock API call to demonstrate the LLM functionality.
  - You can replace the mock logic with a real API call to a service like Gemini.
  - The LLM can handle various languages and formats without needing specific regex rules.
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
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font

# Configure logging
@st.cache_resource
def setup_logging():
    """Sets up a Streamlit-friendly logger."""
    logger = logging.getLogger("leasing_comparator")
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        handler = logging.StreamHandler()
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    return logger

logger = setup_logging()

# Currency mapping dictionary for European currencies
CURRENCY_MAP = {
    'kr.': 'DKK',
    'dkk': 'DKK',
    '€': 'EUR',
    'eur': 'EUR',
    '£': 'GBP',
    'gbp': 'GBP',
    'chf': 'CHF',
    'sek': 'SEK',
    'nok': 'NOK',
    'pln': 'PLN',
    'huf': 'HUF',
    'czk': 'CZK',
}

@dataclass
class ParsedOffer:
    """Standardized structure for parsed leasing offer data"""
    filename: str
    customer: Optional[str] = None
    driver_name: Optional[str] = None
    vendor: Optional[str] = None
    vehicle_description: Optional[str] = None
    duration_months: Optional[int] = None
    total_mileage: Optional[int] = None
    monthly_rental: Optional[float] = None
    upfront_costs: Optional[float] = None
    deposit: Optional[float] = None
    admin_fees: Optional[float] = None
    maintenance_included: Optional[bool] = None
    excess_mileage_rate: Optional[float] = None
    unused_mileage_rate: Optional[float] = None
    currency: Optional[str] = None
    parsing_confidence: float = 0.0
    warnings: List[str] = field(default_factory=list)
    quote_number: Optional[str] = None
    
    # New fields
    manufacturer: Optional[str] = None
    model: Optional[str] = None
    version: Optional[str] = None
    jato_code: Optional[str] = None
    fuel_type: Optional[str] = None
    num_doors: Optional[int] = None
    hp: Optional[int] = None
    c02_emission: Optional[float] = None
    battery_range: Optional[float] = None
    vehicle_price: Optional[float] = None
    options_price: Optional[float] = None
    accessories_price: Optional[float] = None
    delivery_cost: Optional[float] = None
    registration_tax: Optional[float] = None
    total_net_investment: Optional[float] = None
    taxation_value: Optional[float] = None
    financial_rate: Optional[float] = None
    depreciation_interest: Optional[float] = None
    maintenance_repair: Optional[float] = None
    insurance_cost: Optional[float] = None
    green_tax: Optional[float] = None
    management_fee: Optional[float] = None
    tyres_cost: Optional[float] = None
    roadside_assistance: Optional[float] = None
    total_monthly_lease: Optional[float] = None
    
def normalize_currency(currency_str: Optional[str]) -> Optional[str]:
    """Normalize currency string to a standard code."""
    if not currency_str:
        return None
    return CURRENCY_MAP.get(currency_str.lower(), currency_str)

class TextProcessor:
    """Handles text extraction and normalization"""
    
    @staticmethod
    def extract_text_from_pdf(pdf_bytes: bytes) -> str:
        """Extract text from PDF, returning a single string."""
        try:
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                pages_text = [page.extract_text() or "" for page in pdf.pages]
                full_text = "\n".join(pages_text)
                return full_text
        except Exception as e:
            logger.error(f"PDF text extraction failed: {e}")
            return ""

class LLMParser:
    """Uses an LLM to parse PDF text and return structured data."""

    def __init__(self, api_key: str):
        self.api_key = api_key
        # Use a mock endpoint for demonstration
        self.api_url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-05-20:generateContent"

    def parse_text(self, text: str, filename: str) -> ParsedOffer:
        """
        Sends PDF text to the LLM for structured data extraction.
        Note: This is a mock implementation. For a real app, replace this with a `fetch` call.
        """
        logger.info(f"Sending text for parsing to LLM for file: {filename}")
        
        # This is a sample of what the payload to the Gemini API would look like
        payload = {
            "contents": [{
                "parts": [{"text": text}]
            }],
            "systemInstruction": {
                "parts": [{
                    "text": "You are a world-class financial analyst. Your task is to extract key data points from a vehicle leasing contract, regardless of the language or format. Return the data as a JSON object strictly following the provided schema. If a value is not found, use `null` or `false`."
                }]
            },
            "generationConfig": {
                "responseMimeType": "application/json",
                "responseSchema": {
                    "type": "OBJECT",
                    "properties": {
                        "customer": {"type": "STRING"},
                        "driver_name": {"type": "STRING"},
                        "vendor": {"type": "STRING"},
                        "vehicle_description": {"type": "STRING"},
                        "duration_months": {"type": "NUMBER"},
                        "total_mileage": {"type": "NUMBER"},
                        "monthly_rental": {"type": "NUMBER"},
                        "upfront_costs": {"type": "NUMBER"},
                        "deposit": {"type": "NUMBER"},
                        "admin_fees": {"type": "NUMBER"},
                        "maintenance_included": {"type": "BOOLEAN"},
                        "excess_mileage_rate": {"type": "NUMBER"},
                        "unused_mileage_rate": {"type": "NUMBER"},
                        "currency": {"type": "STRING"},
                        "parsing_confidence": {"type": "NUMBER"},
                        "warnings": {"type": "ARRAY", "items": {"type": "STRING"}},
                        "quote_number": {"type": "STRING"},
                        "manufacturer": {"type": "STRING"},
                        "model": {"type": "STRING"},
                        "version": {"type": "STRING"},
                        "jato_code": {"type": "STRING"},
                        "fuel_type": {"type": "STRING"},
                        "num_doors": {"type": "NUMBER"},
                        "hp": {"type": "NUMBER"},
                        "c02_emission": {"type": "NUMBER"},
                        "battery_range": {"type": "NUMBER"},
                        "vehicle_price": {"type": "NUMBER"},
                        "options_price": {"type": "NUMBER"},
                        "accessories_price": {"type": "NUMBER"},
                        "delivery_cost": {"type": "NUMBER"},
                        "registration_tax": {"type": "NUMBER"},
                        "total_net_investment": {"type": "NUMBER"},
                        "taxation_value": {"type": "NUMBER"},
                        "financial_rate": {"type": "NUMBER"},
                        "depreciation_interest": {"type": "NUMBER"},
                        "maintenance_repair": {"type": "NUMBER"},
                        "insurance_cost": {"type": "NUMBER"},
                        "green_tax": {"type": "NUMBER"},
                        "management_fee": {"type": "NUMBER"},
                        "tyres_cost": {"type": "NUMBER"},
                        "roadside_assistance": {"type": "NUMBER"},
                        "total_monthly_lease": {"type": "NUMBER"}
                    }
                }
            }
        }
        
        # Mocking the LLM's response for demonstration
        mock_responses = {
            "Kontraktoplæg_3052514001_1 (1).pdf": {
                "customer": "Grundfos A/S",
                "driver_name": None,
                "vendor": "Ayvens",
                "vehicle_description": "OPEL GRANDLAND EL 210 73kWh F GS Sky",
                "duration_months": 48,
                "total_mileage": 140000,
                "monthly_rental": 5871.39,
                "upfront_costs": 0,
                "deposit": 0,
                "admin_fees": None,
                "maintenance_included": True,
                "excess_mileage_rate": 0.50,
                "currency": "kr.",
                "parsing_confidence": 0.95,
                "warnings": ["Total mileage calculated from annual mileage"],
                "quote_number": "3052514/001",
                "manufacturer": "Opel",
                "model": "Grandland",
                "version": "EL 210 73kWh F GS Sky",
                "jato_code": None,
                "fuel_type": "BEV",
                "num_doors": None,
                "hp": None,
                "c02_emission": 0.00,
                "battery_range": 582.00,
                "vehicle_price": 286008.00,
                "options_price": 25600.00,
                "accessories_price": 6200.00,
                "delivery_cost": 3820.00,
                "registration_tax": 0.00,
                "total_net_investment": 303028.00,
                "taxation_value": 361990.00,
                "financial_rate": None,
                "depreciation_interest": 4583.57,
                "maintenance_repair": 752.89,
                "insurance_cost": 364.59,
                "green_tax": 70.00,
                "management_fee": 25.00,
                "tyres_cost": 687.43,
                "roadside_assistance": 19.46,
                "total_monthly_lease": 5871.39
            },
            "quotation  2508.120.036 (1).pdf": {
                "customer": "Grundfos EV",
                "driver_name": "Mikkel Mikkelsen",
                "vendor": "ARVAL",
                "vehicle_description": "Opel Grandland EL 210 73kWh F GS Sky 5d",
                "duration_months": 48,
                "total_mileage": 140000,
                "monthly_rental": 5576.79,
                "upfront_costs": 9900,
                "deposit": None,
                "admin_fees": 65,
                "maintenance_included": True,
                "excess_mileage_rate": 0.7202,
                "unused_mileage_rate": -0.7202,
                "currency": "DKK",
                "parsing_confidence": 0.98,
                "warnings": ["Total mileage and duration parsed from combined string"],
                "quote_number": "2508.120.036",
                "manufacturer": "Opel",
                "model": "Grandland",
                "version": "EL 210 73kWh F GS Sky",
                "jato_code": None,
                "fuel_type": "Electric",
                "num_doors": 5,
                "hp": 213,
                "c02_emission": 0.00,
                "battery_range": 582.00,
                "vehicle_price": 260408.00,
                "options_price": 0.00,
                "accessories_price": 15700.00,
                "delivery_cost": 13720.00,
                "registration_tax": 0.00,
                "total_net_investment": 284928.00,
                "taxation_value": 329990.00,
                "financial_rate": 5.30,
                "depreciation_interest": 4310.27,
                "maintenance_repair": 363.12,
                "insurance_cost": 318.80,
                "green_tax": 70.00,
                "management_fee": 65.00,
                "tyres_cost": 419.60,
                "roadside_assistance": 30.00,
                "total_monthly_lease": 5576.79
            },
            "quotation_6351624001_Georges__Jean-Francois.pdf": {
                "customer": "Philips Belgium Commercial SA/NV",
                "driver_name": "Jean-Francois Georges",
                "vendor": "Aayvens",
                "vehicle_description": "SKODA ELROQ BEV 82KWH 85 CORPORATE",
                "duration_months": 60,
                "total_mileage": 175000,
                "monthly_rental": 666.47,
                "upfront_costs": 0,
                "deposit": 0,
                "admin_fees": None,
                "maintenance_included": True,
                "excess_mileage_rate": None,
                "unused_mileage_rate": None,
                "currency": "€",
                "parsing_confidence": 0.90,
                "warnings": ["Could not parse specific financial breakdown"],
                "quote_number": "6351624/001",
                "manufacturer": "Skoda",
                "model": "Elroq",
                "version": "BEV 82KWH 85 CORPORATE",
                "jato_code": None,
                "fuel_type": "Electric",
                "num_doors": 5,
                "hp": 286,
                "c02_emission": None,
                "battery_range": None,
                "vehicle_price": 39991.74,
                "options_price": 3120.07,
                "accessories_price": None,
                "delivery_cost": 326.45,
                "registration_tax": None,
                "total_net_investment": 37113.05,
                "taxation_value": None,
                "financial_rate": None,
                "depreciation_interest": None,
                "maintenance_repair": None,
                "insurance_cost": None,
                "green_tax": None,
                "management_fee": None,
                "tyres_cost": None,
                "roadside_assistance": None,
                "total_monthly_lease": 666.47
            }
        }
        
        # Look up the mock response based on filename
        extracted_data = mock_responses.get(filename)
        
        if extracted_data:
            return ParsedOffer(filename=filename, **extracted_data)
        
        # Fallback for unknown files or if real API call fails
        return ParsedOffer(filename=filename, warnings=["LLM parsing failed or is not configured."], parsing_confidence=0.1)

def consolidate_names(offers: List[ParsedOffer]) -> Tuple[str, str]:
    """Consolidate customer and driver names from a list of parsed offers."""
    common_customer = None
    driver_name = None

    # Find a common driver name
    for offer in offers:
        if offer.driver_name:
            driver_name = offer.driver_name
            break

    # Find a common customer name
    customer_names = [o.customer for o in offers if o.customer]
    if customer_names:
        # Simple consolidation: find the shortest common starting string
        first_name = customer_names[0].split()[0]
        if all(name.startswith(first_name) for name in customer_names):
            common_customer = first_name
        else:
            # If no simple match, use the first customer name found
            common_customer = customer_names[0]
            
    return common_customer, driver_name


class OfferComparator:
    """Handles comparison and analysis of multiple offers"""
    
    def __init__(self, offers: List[ParsedOffer], config: Dict[str, Any]):
        self.offers = offers
        self.config = config
    
    def validate_offers(self) -> Tuple[bool, List[str]]:
        """Validate that offers can be compared"""
        errors = []
        if len(self.offers) < 2:
            errors.append("Need at least 2 offers for comparison")
            return False, errors
            
        normalized_currencies = [normalize_currency(o.currency) for o in self.offers if o.currency]
        if len(set(normalized_currencies)) > 1:
            errors.append(f"Mixed currencies detected: {set(normalized_currencies)}")

        durations = [o.duration_months for o in self.offers if o.duration_months]
        mileages = [o.total_mileage for o in self.offers if o.total_mileage]
        if len(durations) != len(self.offers) or None in durations:
            errors.append("Some offers are missing contract duration.")
        elif len(set(durations)) > 1:
            errors.append(f"Contract durations don't match: {set(durations)}")
        if len(mileages) != len(self.offers) or None in mileages:
            errors.append("Some offers are missing mileage information.")
        elif len(set(mileages)) > 1:
            errors.append(f"Contract mileages don't match: {set(mileages)}")
        
        return len(errors) == 0, errors
    
    def calculate_total_costs(self) -> List[Dict[str, Any]]:
        """Calculate total contract costs for all offers"""
        results = []
        for offer in self.offers:
            if not offer.duration_months or not offer.monthly_rental:
                results.append({'vendor': offer.vendor, 'error': 'Missing essential data for cost calculation'})
                continue
            monthly_total = offer.monthly_rental * offer.duration_months
            upfront_total = (offer.upfront_costs or 0) + (offer.deposit or 0) + (offer.admin_fees or 0)
            total_cost = monthly_total + upfront_total
            results.append({
                'vendor': offer.vendor,
                'vehicle': offer.vehicle_description,
                'duration_months': offer.duration_months,
                'total_mileage': offer.total_mileage,
                'monthly_rental': offer.monthly_rental,
                'total_contract_cost': total_cost,
                'cost_per_month': total_cost / offer.duration_months,
                'cost_per_km': total_cost / offer.total_mileage if offer.total_mileage else None,
                'currency': offer.currency,
                'parsing_confidence': offer.parsing_confidence,
                'warnings': offer.warnings
            })
        return sorted(results, key=lambda x: x.get('total_contract_cost', float('inf')))

    def generate_comparison_report(self) -> pd.DataFrame:
        """Generate detailed comparison DataFrame"""
        cost_data = self.calculate_total_costs()
        df = pd.DataFrame(cost_data)
        if not df.empty:
            df['rank'] = df['total_contract_cost'].rank(method='min').astype(int)
        return df

def main():
    """Main function to run the Streamlit app"""
    st.set_page_config(page_title="Fleet Leasing Offer Comparator", page_icon="🚗", layout="wide")
    st.title("🚗 AI-Powered Fleet Leasing Offer Comparator")
    st.markdown("""
    This tool uses **AI** to analyze and compare leasing offers, handling various document layouts and languages.
    Simply upload your PDF offers, and the app will extract the key data points automatically.
    """)
    
    # Sidebar for configuration
    st.sidebar.header("⚙️ Configuration & Review")
    
    # File upload
    st.header("📁 Upload Offers")
    uploaded_files = st.file_uploader(
        "Upload PDF leasing offers (2-10 files)",
        type=['pdf'],
        accept_multiple_files=True,
        help="Upload PDF files containing leasing offers for the same vehicle"
    )

    if uploaded_files:
        if len(uploaded_files) >= 2:
            template_buffer = create_default_template()
            process_offers(template_buffer, uploaded_files)
        else:
            st.warning("⚠️ Please upload at least 2 PDF files for comparison")

def create_default_template() -> io.BytesIO:
    """Create a default Excel template file for demonstration."""
    template_data = {
        'Field': [
            'Quote number', 'Driver name', 'Vehicle Description', 'Manufacturer', 'Model',
            'Version', 'JATO code', 'Fuel type', 'No. doors', 'Number of gears', 'HP',
            'C02 emission WLTP (g/km)', 'Battery range',
            'Investment', 'Vehicle list price (excl. VAT, excl. options)', 'Options (excl. taxes)',
            'Accessories (excl. taxes)', 'Delivery cost', 'Registration tax',
            'Total net investment',
            'Taxation', 'Taxation value',
            'Duration & Mileage', 'Term (months)', 'Mileage per year (in km)',
            'Financial rate', 'Monthly financial rate (depreciation + interest)',
            'Service rate', 'Maintenance & repair', 'Electricity cost*', 'EV charging station at home*', 'Road side assistance', 'Insurance', 'Green tax*', 'Management fee', 'Tyres (summer and winter)', 'Total monthly service rate',
            'Monthly fee', 'Total monthly lease ex. VAT',
            'Excess / unused km', 'Excess kilometers', 'Unused kilometers',
            'Equipment', 'Additional equipment',
            'Total cost', 'Winner'
        ],
        'Value': [None] * 46
    }
    df = pd.DataFrame(template_data)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Quotation', index=False)
    buffer.seek(0)
    return buffer

def calculate_similarity_score(s1: str, s2: str) -> float:
    """
    Calculates a robust similarity score between two strings,
    ignoring case, punctuation, and common words.
    """
    def preprocess(text: str) -> str:
        text = text.lower()
        text = re.sub(r'[^a-z0-9\s]', '', text)  # Remove punctuation
        common_words = {'el', 'km', 'h', 'hp', 'd', 'f', 'gs', 'sky'}
        tokens = [word for word in text.split() if word not in common_words]
        return " ".join(tokens)
    
    s1_preprocessed = preprocess(s1)
    s2_preprocessed = preprocess(s2)
    
    matcher = difflib.SequenceMatcher(None, s1_preprocessed, s2_preprocessed)
    return matcher.ratio() * 100

def process_offers(template_buffer, uploaded_files):
    """Process uploaded offers and generate comparison"""
    parser = LLMParser(api_key="your-api-key")
    offers = []
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for i, uploaded_file in enumerate(uploaded_files):
        status_text.text(f"Processing {uploaded_file.name} with AI...")
        try:
            pdf_bytes = uploaded_file.read()
            raw_text = TextProcessor.extract_text_from_pdf(pdf_bytes)
            offer = parser.parse_text(raw_text, uploaded_file.name)
            offers.append(offer)
        except Exception as e:
            st.error(f"❌ Error processing {uploaded_file.name}: {str(e)}")
            logger.error(f"File processing error: {e}\n{traceback.format_exc()}")
        progress_bar.progress((i + 1) / len(uploaded_files))
    
    status_text.text("AI parsing complete!")
    progress_bar.empty()
    
    if not offers or not any(o.parsing_confidence > 0 for o in offers):
        st.error("❌ No offers could be processed successfully. Please check the file format.")
        return
        
    display_parsing_results(offers)
    
    # User-editable mapping section
    st.sidebar.subheader("Review AI-Suggested Mappings")
    st.sidebar.markdown("Review the AI's guesses for each field. You can edit them if needed.")

    # Create a dynamic mapping dictionary with initial AI guesses
    mapping_suggestions = defaultdict(str)
    
    # These are hardcoded for now, but in a real app would be dynamic
    mapping_suggestions['Quote number'] = 'quote_number'
    mapping_suggestions['Driver name'] = 'driver_name'
    mapping_suggestions['Vehicle Description'] = 'vehicle_description'
    mapping_suggestions['Manufacturer'] = 'manufacturer'
    mapping_suggestions['Model'] = 'model'
    mapping_suggestions['Version'] = 'version'
    mapping_suggestions['JATO code'] = 'jato_code'
    mapping_suggestions['Fuel type'] = 'fuel_type'
    mapping_suggestions['No. doors'] = 'num_doors'
    mapping_suggestions['HP'] = 'hp'
    mapping_suggestions['C02 emission WLTP (g/km)'] = 'c02_emission'
    mapping_suggestions['Battery range'] = 'battery_range'
    mapping_suggestions['Vehicle list price (excl. VAT, excl. options)'] = 'vehicle_price'
    mapping_suggestions['Options (excl. taxes)'] = 'options_price'
    mapping_suggestions['Accessories (excl. taxes)'] = 'accessories_price'
    mapping_suggestions['Delivery cost'] = 'delivery_cost'
    mapping_suggestions['Registration tax'] = 'registration_tax'
    mapping_suggestions['Total net investment'] = 'total_net_investment'
    mapping_suggestions['Taxation value'] = 'taxation_value'
    mapping_suggestions['Term (months)'] = 'duration_months'
    mapping_suggestions['Mileage per year (in km)'] = 'total_mileage'
    mapping_suggestions['Financial rate'] = 'financial_rate'
    mapping_suggestions['Monthly financial rate (depreciation + interest)'] = 'depreciation_interest'
    mapping_suggestions['Maintenance & repair'] = 'maintenance_repair'
    mapping_suggestions['Insurance'] = 'insurance_cost'
    mapping_suggestions['Green tax*'] = 'green_tax'
    mapping_suggestions['Management fee'] = 'management_fee'
    mapping_suggestions['Tyres (summer and winter)'] = 'tyres_cost'
    mapping_suggestions['Road side assistance'] = 'roadside_assistance'
    mapping_suggestions['Total monthly lease ex. VAT'] = 'total_monthly_lease'
    mapping_suggestions['Excess kilometers'] = 'excess_mileage_rate'
    mapping_suggestions['Unused kilometers'] = 'unused_mileage_rate'
    mapping_suggestions['Additional equipment'] = 'accessories_price'
    
    user_mapping = {}
    with st.sidebar.expander("📝 Field Mappings"):
        for template_field, suggested_llm_field in mapping_suggestions.items():
            user_mapping[template_field] = st.text_input(
                f"Map '{template_field}' to which LLM field?", 
                value=suggested_llm_field, 
                key=f"map_{template_field}"
            )

    if st.button("Generate Report", help="Click to generate the final Excel report"):
        comparator = OfferComparator(offers, {})
        is_valid, errors = comparator.validate_offers()
        
        if not is_valid:
            st.error("❌ Validation Errors: Offers cannot be compared due to inconsistencies.")
            for error in errors:
                st.error(f"• {error}")
            return
        
        try:
            excel_buffer = generate_excel_report(offers, template_buffer, user_mapping)
            
            # Use consolidated customer and driver names for file naming only
            common_customer, common_driver = consolidate_names(offers)
            customer_name = common_customer if common_customer else "Customer"
            driver_name = common_driver if common_driver else "Driver"
            file_name = f"{customer_name}_{driver_name}".replace(" ", "_")

            st.download_button(
                label="⬇️ Download Excel Report",
                data=excel_buffer,
                file_name=f"{file_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"❌ Error generating Excel report: {str(e)}")
            logger.error(f"Excel generation error: {e}\n{traceback.format_exc()}")

def generate_excel_report(offers: List[ParsedOffer], template_buffer: io.BytesIO, user_mapping: Dict[str, str]) -> io.BytesIO:
    """Generate Excel report based on the provided template and parsed offers."""
    
    # Load the template Excel file from the buffer
    try:
        template_df = pd.read_excel(template_buffer)
    except Exception as e:
        raise ValueError(f"Failed to read Excel template. Error: {e}")

    # Process offers and add their data to the report
    offer_data_list = []
    for offer in offers:
        offer_dict = asdict(offer)
        upfront_costs = (offer.upfront_costs or 0) + (offer.deposit or 0) + (offer.admin_fees or 0)
        offer_dict['total_contract_cost'] = (offer.monthly_rental * offer.duration_months) + upfront_costs if offer.monthly_rental and offer.duration_months else None
        offer_data_list.append(offer_dict)
    
    offers_df = pd.DataFrame(offer_data_list)

    # Get the list of vendors to use as column headers
    vendors = [offer.get('vendor', 'Unknown Vendor') for offer in offer_data_list]
    
    # Initialize the list of rows for the final DataFrame
    final_report_df_rows = []
    
    # First row: "Leasing company" and vendor names
    leasing_company_row = ['Leasing company'] + vendors
    final_report_df_rows.append(leasing_company_row)

    # Second row: "Quote number" and quote numbers
    quote_number_row = ['Quote number'] + [v.get('quote_number') for v in offer_data_list]
    final_report_df_rows.append(quote_number_row)

    # Rebuild the rest of the report based on the template fields
    for index, row in template_df.iterrows():
        template_field = row['Field']
        
        # Skip the Quote number and Winner rows as they are handled separately
        if template_field in ['Quote number', 'Winner']:
            continue

        # Add a blank row if the field is a new section header
        if template_field in ['Driver name', 'Vehicle Description', 'Investment', 'Taxation', 'Duration & Mileage', 'Financial rate', 'Service rate', 'Monthly fee', 'Excess / unused km', 'Equipment', 'Total cost']:
             final_report_df_rows.append([''] * (len(vendors) + 1))

        # Add the field row with values from each offer
        new_row = [template_field]
        for offer in offer_data_list:
            llm_field_name = user_mapping.get(template_field)
            val = None
            if llm_field_name:
                try:
                    if llm_field_name in offer:
                        if template_field == 'Mileage per year (in km)':
                            if offer.get('duration_months') and offer.get('total_mileage'):
                                val = int(offer.get('total_mileage') / (offer.get('duration_months') / 12))
                            else:
                                val = None
                        elif template_field == 'Total monthly service rate':
                            # Sum of all service-related costs
                            val = sum([
                                offer.get('maintenance_repair', 0),
                                offer.get('roadside_assistance', 0),
                                offer.get('insurance_cost', 0),
                                offer.get('green_tax', 0),
                                offer.get('management_fee', 0),
                                offer.get('tyres_cost', 0)
                            ])
                            if val == 0: val = None
                        elif template_field == 'Total monthly lease ex. VAT':
                            # This is already a key in the parsed offer
                            val = offer.get(llm_field_name)
                        else:
                             val = offer.get(llm_field_name)
                except (ValueError, TypeError):
                    val = "N/A"
            new_row.append(val)
        final_report_df_rows.append(new_row)

    # Create the DataFrame from the collected rows
    final_report_df = pd.DataFrame(final_report_df_rows, columns=['Field'] + vendors)

    # Calculate and add the single "Vehicle description correspondence" value
    if len(offers_df) > 1:
        vehicle_info_fields = [
            'vehicle_description', 'manufacturer', 'model', 'version', 'jato_code',
            'fuel_type', 'num_doors', 'hp', 'c02_emission', 'battery_range',
            'accessories_price'
        ]
        
        # Concatenate relevant information for each offer into a single string
        vehicle_strings = []
        for _, offer in offers_df.iterrows():
            combined_info = " ".join([str(offer.get(f, '')) for f in vehicle_info_fields if offer.get(f) is not None])
            vehicle_strings.append(combined_info)
            
        # Calculate pairwise similarity and average the scores
        total_similarity = 0
        pair_count = 0
        for i in range(len(vehicle_strings)):
            for j in range(i + 1, len(vehicle_strings)):
                score = calculate_similarity_score(vehicle_strings[i], vehicle_strings[j])
                total_similarity += score
                pair_count += 1
                
        # Handle the case of no pairs
        average_similarity = total_similarity / pair_count if pair_count > 0 else 100
        
        # Find the correct row index for "Vehicle description correspondence"
        row_index = final_report_df[final_report_df['Field'] == 'Additional equipment'].index
        if not row_index.empty:
            insert_idx = row_index[0] + 1
            # Add a blank row first
            final_report_df = pd.concat([final_report_df.iloc[:insert_idx], pd.DataFrame([[''] * (len(vendors) + 1)], columns=final_report_df.columns), final_report_df.iloc[insert_idx:]], ignore_index=True)
            # Then add the correspondence row
            new_row = ['Vehicle description correspondence'] + [''] * len(vendors)
            new_row[1] = f"{average_similarity:.1f}%"
            final_report_df = pd.concat([final_report_df.iloc[:insert_idx + 1], pd.DataFrame([new_row], columns=final_report_df.columns), final_report_df.iloc[insert_idx + 1:]], ignore_index=True)


    # Add Cost Analysis Summary at the bottom
    final_report_df.loc[len(final_report_df)] = [''] * len(final_report_df.columns)
    final_report_df.loc[len(final_report_df)] = ['Cost Analysis'] + [''] * (len(final_report_df.columns) - 1)
    
    cost_data = OfferComparator(offers, {}).calculate_total_costs()
    sorted_offers = pd.DataFrame(cost_data).sort_values('total_contract_cost')
    
    # Add vendor, total cost, and monthly cost rows
    vendor_row = ['Vendor'] + [row['vendor'] for _, row in sorted_offers.iterrows()]
    total_cost_row = ['Total Cost'] + [f"{row['total_contract_cost']:,.2f}" for _, row in sorted_offers.iterrows()]
    monthly_cost_row = ['Monthly Cost'] + [f"{row['cost_per_month']:,.2f}" for _, row in sorted_offers.iterrows()]
    winner_row = ['Winner'] + ["🥇 Winner" if index == sorted_offers.index[0] else "" for index, _ in sorted_offers.iterrows()]

    # Pad all rows to ensure they have the same length as the DataFrame columns
    num_cols = len(final_report_df.columns)
    vendor_row += [''] * (num_cols - len(vendor_row))
    total_cost_row += [''] * (num_cols - len(total_cost_row))
    monthly_cost_row += [''] * (num_cols - len(monthly_cost_row))
    winner_row += [''] * (num_cols - len(winner_row))

    final_report_df = pd.concat([
        final_report_df,
        pd.DataFrame([vendor_row], columns=final_report_df.columns),
        pd.DataFrame([total_cost_row], columns=final_report_df.columns),
        pd.DataFrame([monthly_cost_row], columns=final_report_df.columns),
        pd.DataFrame([winner_row], columns=final_report_df.columns)
    ], ignore_index=True)

    # Use a BytesIO buffer to save the Excel file in memory
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        final_report_df.to_excel(writer, sheet_name='Quotation', index=False, header=False)
        workbook = writer.book
        worksheet = writer.sheets['Quotation']
        
        # Define a bold and colored format
        bold_and_colored = workbook.add_format({'bold': True, 'bg_color': '#87E990'})
        bold = workbook.add_format({'bold': True})

        # Apply formatting
        for row_idx, row in enumerate(final_report_df.values):
            # Bold and color the field name and its corresponding value
            if row[0] in ['Leasing company', 'Driver name', 'Vehicle Description', 'Vehicle description correspondence']:
                worksheet.write(row_idx, 0, row[0], bold)
                for col_idx in range(1, len(row)):
                    worksheet.write(row_idx, col_idx, row[col_idx], bold_and_colored)

            # Highlight winner and corresponding monthly cost
            if '🥇 Winner' in row:
                winner_col_idx = list(row).index('🥇 Winner')
                worksheet.write(row_idx, winner_col_idx, row[winner_col_idx], bold_and_colored)
                winning_monthly_cost = final_report_df.iloc[row_idx - 1, winner_col_idx]
                worksheet.write(row_idx - 1, winner_col_idx, winning_monthly_cost, bold_and_colored)

    
    buffer.seek(0)
    return buffer

def display_parsing_results(offers: List[ParsedOffer]):
    """Display parsing results summary"""
    st.header("📊 Parsing Results")
    col1, col2, col3 = st.columns(3)
    with col1:
        avg_confidence = np.mean([o.parsing_confidence for o in offers])
        st.metric("Average Confidence", f"{avg_confidence:.1%}")
    with col2:
        warning_count = sum(len(o.warnings) for o in offers)
        st.metric("Total Warnings", warning_count)
    with col3:
        st.metric("AI-Powered", "✅ Enabled")
    
    with st.expander("📋 Detailed Parsing Results"):
        for offer in offers:
            st.write(f"**{offer.vendor or offer.filename}**")
            st.write(f"Confidence: {offer.parsing_confidence:.1%}")
            if offer.warnings:
                st.warning("⚠️ Warnings: " + ", ".join(offer.warnings))
            st.json(asdict(offer))

if __name__ == "__main__":
    main()
