"""
AI-Powered Fleet Leasing Offer Comparator - Streamlit App
This version uses a Large Language Model (LLM) to intelligently parse PDF content.
Author: Fleet Management Tool
Requirements:
  streamlit, pandas, numpy, pdfplumber, python-dateutil, xlsxwriter, google-generativeai
Notes:
  - This version uses a real API call to the Google Gemini API.
  - You must provide a valid API key to use the parsing functionality.
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
import xlsxwriter
import google.generativeai as genai

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
    'kr': 'DKK',
    'dkk': 'DKK',
    '€': 'EUR',
    'eur': 'EUR',
    'euro': 'EUR',
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
    vendor: Optional[str] = None
    vehicle_description: Optional[str] = None
    # Separate fields for actual and maximum contract terms
    max_duration_months: Optional[int] = None
    max_total_mileage: Optional[int] = None
    offer_duration_months: Optional[int] = None
    offer_total_mileage: Optional[int] = None
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

    # New fields to support the extended functionality
    quote_number: Optional[str] = None
    manufacturer: Optional[str] = None
    model: Optional[str] = None
    version: Optional[str] = None
    internal_colour: Optional[str] = None
    external_colour: Optional[str] = None
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
    # The LLM should be flexible enough to recognize "Arval Assistance", etc.
    roadside_assistance: Optional[float] = None
    tyres_cost: Optional[float] = None
    total_monthly_lease: Optional[float] = None
    driver_name: Optional[str] = None
    customer: Optional[str] = None
    # New fields for itemized lists
    options_list: List[Dict[str, Union[str, float]]] = field(default_factory=list)
    accessories_list: List[Dict[str, Union[str, float]]] = field(default_factory=list)

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
    """Uses the Gemini LLM to parse PDF text and return structured data."""

    def __init__(self, api_key: str):
        """Initializes the Gemini client with the provided API key."""
        if not api_key:
            raise ValueError("An API key for the Gemini API is required.")
        self.api_key = api_key
        genai.configure(api_key=self.api_key)
        logger.info("Gemini client configured successfully.")

    def parse_text(self, text: str, filename: str) -> ParsedOffer:
        """
        Sends PDF text to the Gemini API for structured data extraction.
        """
        logger.info(f"Sending text for parsing to Gemini 2.5 Pro for file: {filename}")

        # The instructions are now embedded directly in the prompt
        prompt_text = f"""
        You are a world-class financial analyst specializing in fleet leasing. Your task is to extract key data points from a vehicle leasing contract, regardless of the language or format.

        IMPORTANT:
        1. Distinguish between the **maximum allowed** contract terms and the **actual terms of the offer**.
        - `max_duration_months` and `max_total_mileage` refer to the maximum possible contract length and total mileage allowed (e.g., "Max contract: 60 months / 300,000 km").
        - `offer_duration_months` and `offer_total_mileage` refer to the specific terms of this offer (e.g., "Current offer: 36 months / 175,000 km").

        2. All extracted price and cost amounts, including `monthly_rental` and `total_monthly_lease`, must be **excluding VAT (Value-Added Tax)**. Look for cues like "excl. VAT", "HT", "net price", etc.

        3. Differentiate between the driver and the customer. `driver_name` is the employee using the car. `customer` is the company renting the car, usually in the header.

        4. For `roadside_assistance`, include amounts for phrases like "Arval assistance" or "Ayvens assistance".

        5. Calculate `offer_total_mileage` if annual mileage is given. Example: "35,000 km per year / 48 months" -> total mileage is 35000 * 48 / 12 = 140000.

        6. `total_net_investment` is sometimes called "Taxable value" or "Total tax of the car".

        7. Treat "BEV", "Battery Electric Vehicle", "Electric", and "electricity" as the same `fuel_type`.

        8. Extract the `internal_colour` and `external_colour` of the vehicle.

        Return the data as a JSON object strictly following the provided schema. If a value is not found, use `null`. Do not invent values.
        
        <DOCUMENT_TO_PARSE>
        {text}
        </DOCUMENT_TO_PARSE>
        """

        # This defines the JSON structure we want the LLM to return
        json_schema = {
            "type": "OBJECT",
            "properties": {
                "vendor": {"type": "STRING"},
                "vehicle_description": {"type": "STRING"},
                "max_duration_months": {"type": "NUMBER"},
                "max_total_mileage": {"type": "NUMBER"},
                "offer_duration_months": {"type": "NUMBER"},
                "offer_total_mileage": {"type": "NUMBER"},
                "monthly_rental": {"type": "NUMBER"},
                "upfront_costs": {"type": "NUMBER"},
                "deposit": {"type": "NUMBER"},
                "admin_fees": {"type": "NUMBER"},
                "maintenance_included": {"type": "BOOLEAN"},
                "excess_mileage_rate": {"type": "NUMBER"},
                "unused_mileage_rate": {"type": "NUMBER"},
                "currency": {"type": "STRING"},
                "parsing_confidence": {"type": "NUMBER", "description": "A value between 0.0 and 1.0 indicating how confident you are in the extracted data."},
                "warnings": {"type": "ARRAY", "items": {"type": "STRING"}},
                "quote_number": {"type": "STRING"},
                "manufacturer": {"type": "STRING"},
                "model": {"type": "STRING"},
                "version": {"type": "STRING"},
                "internal_colour": {"type": "STRING"},
                "external_colour": {"type": "STRING"},
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
                "total_monthly_lease": {"type": "NUMBER"},
                "driver_name": {"type": "STRING"},
                "customer": {"type": "STRING"},
                "options_list": {"type": "ARRAY", "items": {"type": "OBJECT", "properties": {"name": {"type": "STRING"}, "price": {"type": "NUMBER"}}}},
                "accessories_list": {"type": "ARRAY", "items": {"type": "OBJECT", "properties": {"name": {"type": "STRING"}, "price": {"type": "NUMBER"}}}}
            }
        }

        # Configure the generation settings to force JSON output
        generation_config = genai.types.GenerationConfig(
            response_mime_type="application/json",
            response_schema=json_schema
        )
        
        model = genai.GenerativeModel(
            model_name='gemini-2.5-pro',
            generation_config=generation_config
        )

        try:
            response = model.generate_content(prompt_text)
            extracted_data = json.loads(response.text)
            extracted_data['options_list'] = extracted_data.get('options_list') or []
            extracted_data['accessories_list'] = extracted_data.get('accessories_list') or []

            return ParsedOffer(filename=filename, **extracted_data)

        except Exception as e:
            logger.error(f"Error during Gemini 2.5 Pro API call for {filename}: {str(e)}")
            logger.error(traceback.format_exc())
            return ParsedOffer(
                filename=filename,
                warnings=[f"LLM parsing failed due to an API error: {str(e)}"],
                parsing_confidence=0.1
            )

class OfferComparator:
    """Handles comparison and analysis of multiple offers"""

    def __init__(self, offers: List[ParsedOffer]):
        self.offers = offers

    def validate_offers(self) -> Tuple[bool, List[str]]:
        """Validate that offers can be compared"""
        errors = []
        if len(self.offers) < 2:
            errors.append("Need at least 2 offers for comparison")
            return False, errors

        normalized_currencies = [normalize_currency(o.currency) for o in self.offers if o.currency]
        if len(set(normalized_currencies)) > 1:
            errors.append(f"Mixed currencies detected: {set(normalized_currencies)}")

        durations = [o.offer_duration_months for o in self.offers if o.offer_duration_months]
        mileages = [o.offer_total_mileage for o in self.offers if o.offer_total_mileage]

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
            if not offer.offer_duration_months or not offer.monthly_rental:
                results.append({'vendor': offer.vendor, 'error': 'Missing essential data for cost calculation'})
                continue
            monthly_total = offer.monthly_rental * offer.offer_duration_months
            upfront_total = (offer.upfront_costs or 0) + (offer.deposit or 0) + (offer.admin_fees or 0)
            total_cost = monthly_total + upfront_total
            results.append({
                'vendor': offer.vendor,
                'vehicle': offer.vehicle_description,
                'duration_months': offer.offer_duration_months,
                'total_mileage': offer.offer_total_mileage,
                'monthly_rental': offer.monthly_rental,
                'total_contract_cost': total_cost,
                'cost_per_month': total_cost / offer.offer_duration_months,
                'cost_per_km': total_cost / offer.offer_total_mileage if offer.offer_total_mileage else None,
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

    st.sidebar.header("⚙️ Configuration & Review")

    api_key = st.sidebar.text_input(
        "Enter your Google AI API Key",
        type="password",
        help="Get your API key from Google AI Studio. For deployed apps, use st.secrets."
    )

    st.header("📁 Upload Offers")

    reference_file = st.file_uploader(
        "Upload the Reference Offer (1 file)",
        type=['pdf'],
        accept_multiple_files=False,
        help="Upload the PDF file that will be used as the benchmark for comparison"
    )

    other_files = st.file_uploader(
        "Upload Other Offers (1-9 files)",
        type=['pdf'],
        accept_multiple_files=True,
        help="Upload the other PDF files you want to compare against the reference offer"
    )

    if reference_file or other_files:
        uploaded_files = [reference_file] + other_files
        current_file_names = [f.name for f in uploaded_files if f is not None]

        if 'offers' not in st.session_state or st.session_state.get('uploaded_files') != current_file_names:
            if not api_key:
                st.error("❌ Please enter your Google AI API Key in the sidebar to proceed.")
                st.stop()

            st.session_state.offers = process_offers_internal(api_key, uploaded_files)
            st.session_state.uploaded_files = current_file_names

        if st.session_state.offers:
            display_parsing_results(st.session_state.offers)

    st.sidebar.subheader("Review AI-Suggested Mappings")
    st.sidebar.markdown("Review the AI's guesses for each field. You can edit them if needed.")

    mapping_suggestions = defaultdict(str)
    mapping_suggestions['Quote number'] = 'quote_number'
    mapping_suggestions['Driver name'] = 'driver_name'
    mapping_suggestions['Vehicle Description'] = 'vehicle_description'
    mapping_suggestions['Manufacturer'] = 'manufacturer'
    mapping_suggestions['Model'] = 'model'
    mapping_suggestions['Version'] = 'version'
    mapping_suggestions['Internal colour'] = 'internal_colour'
    mapping_suggestions['External colour'] = 'external_colour'
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
    mapping_suggestions['Term (months)'] = 'offer_duration_months'
    mapping_suggestions['Mileage per year (in km)'] = 'offer_total_mileage'
    mapping_suggestions['Monthly financial rate (depreciation + interest)'] = 'depreciation_interest'
    mapping_suggestions['Maintenance & repair'] = 'maintenance_repair'
    mapping_suggestions['Insurance'] = 'insurance_cost'
    mapping_suggestions['Green tax*'] = 'green_tax'
    mapping_suggestions['Management fee'] = 'management_fee'
    mapping_suggestions['Tyres (summer and winter)'] = 'tyres_cost'
    mapping_suggestions['Road side assistance'] = 'roadside_assistance'
    mapping_suggestions['Total monthly service rate'] = 'total_monthly_service_rate'
    mapping_suggestions['Total monthly lease ex. VAT'] = 'total_monthly_lease'
    mapping_suggestions['Excess kilometers'] = 'excess_mileage_rate'
    mapping_suggestions['Unused kilometers'] = 'unused_mileage_rate'

    user_mapping = {}
    with st.sidebar.expander("📝 Field Mappings"):
        for template_field, suggested_llm_field in mapping_suggestions.items():
            user_mapping[template_field] = st.text_input(
                f"Map '{template_field}' to which LLM field?",
                value=suggested_llm_field,
                key=f"map_{template_field}"
            )

    if st.button("Generate Report", help="Click to generate the final Excel report"):
        offers = st.session_state.get('offers')
        if not offers:
            st.error("❌ No offers found. Please upload files first.")
            return

        comparator = OfferComparator(offers)
        is_valid, errors = comparator.validate_offers()

        if not is_valid:
            st.error("❌ Validation Errors: Offers cannot be compared due to inconsistencies.")
            for error in errors:
                st.error(f"• {error}")
            return

        try:
            template_buffer = create_default_template()
            excel_buffer = generate_excel_report(offers, template_buffer, user_mapping)
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

def process_offers_internal(api_key: str, uploaded_files) -> List[ParsedOffer]:
    """Helper function to process offers and return the list of parsed objects."""
    try:
        parser = LLMParser(api_key=api_key)
    except ValueError as e:
        st.error(f"❌ Initialization Error: {e}")
        return []

    offers = []
    progress_bar = st.progress(0)
    status_text = st.empty()

    for i, uploaded_file in enumerate(uploaded_files):
        if uploaded_file is None:
            continue
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
        st.error("❌ No offers could be processed successfully. Please check the file format or API key.")
        return []
    return offers

def create_default_template() -> io.BytesIO:
    """Create a default Excel template file for demonstration."""
    # Define the fields in a list to control the order
    fields = [
        'Quote number', 'Driver name', 'Vehicle Description', 'Manufacturer', 'Model',
        'Version', 'Internal colour', 'External colour', 'Fuel type',
        'No. doors', 'Number of gears', 'HP', 'C02 emission WLTP (g/km)', 'Battery range',
        'Equipment', 'Additional equipment', 'Additional equipment price',
        'Investment', 'Vehicle list price (excl. VAT, excl. options)', 'Options (excl. taxes)',
        'Accessories (excl. taxes)', 'Delivery cost', 'Registration tax',
        'Total net investment',
        'Taxation', 'Taxation value',
        'Duration & Mileage', 'Term (months)', 'Mileage per year (in km)',
        'Financial rate', 'Monthly financial rate (depreciation + interest)',
        'Service rate', 'Maintenance & repair', 'Electricity cost*', 'EV charging station at home*', 
        'Road side assistance', 'Insurance', 'Green tax*', 'Management fee', 'Tyres (summer and winter)', 
        'Total monthly service rate',
        'Monthly fee', 'Total monthly lease ex. VAT',
        'Excess / unused km', 'Excess kilometers', 'Unused kilometers',
        'Winner'
    ]
    
    # Build the dictionary dynamically to prevent length mismatches
    template_data = {
        'Field': fields,
        'Value': [None] * len(fields)
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
    ignoring case, punctuation, and common irrelevant words.
    """
    def preprocess(text: str) -> str:
        text = text.lower()
        text = re.sub(r'[^a-z0-9\s]', '', text)
        common_words = {'el', 'km', 'h', 'hp', 'd', 'f', 'gs', 'sky', 'hk', 'auto', 'farve', 'color'}
        tokens = [word for word in text.split() if word not in common_words]
        return " ".join(tokens)

    s1_preprocessed = preprocess(s1)
    s2_preprocessed = preprocess(s2)

    matcher = difflib.SequenceMatcher(None, s1_preprocessed, s2_preprocessed)
    return matcher.ratio() * 100

def get_offer_diff(offer1: ParsedOffer, offer2: ParsedOffer) -> str:
    """Compares two ParsedOffer objects and returns a string summarizing the differences."""
    diff_summary = []
    SIMILARITY_THRESHOLD = 90.0
    ELECTRIC_SYNONYMS = {'bev', 'electric', 'battery electric vehicle', 'electricity'}

    # === Compare standard fields ===
    fields_to_compare = [
        'vehicle_description', 'manufacturer', 'model', 'version',
        'internal_colour', 'external_colour', # Added colours
        'offer_duration_months', 'offer_total_mileage',
        'currency', 'taxation_value', 'green_tax',
        'fuel_type',
    ]

    for field in fields_to_compare:
        val1 = getattr(offer1, field)
        val2 = getattr(offer2, field)
        
        # Normalize strings for comparison
        val1_str = str(val1 or '').strip().lower()
        val2_str = str(val2 or '').strip().lower()

        # Check for a perfect case-insensitive match
        if val1_str == val2_str:
            continue
        
        # Special handling for fuzzy matches and substrings
        if field in ['vehicle_description', 'version']:
            # Check for substring match, e.g., 'EV3' in 'EV3 77kWh'
            if val1_str in val2_str or val2_str in val1_str:
                continue
            # Fallback to the similarity score for more complex matches
            score = calculate_similarity_score(val1_str, val2_str)
            if score >= SIMILARITY_THRESHOLD:
                continue
        elif field == 'currency':
            if normalize_currency(val1_str) == normalize_currency(val2_str):
                continue
        elif field == 'fuel_type':
            if val1_str in ELECTRIC_SYNONYMS and val2_str in ELECTRIC_SYNONYMS:
                continue

        # Convert back to original case for the diff message
        val1_display = str(val1) if val1 is not None else "MISSING"
        val2_display = str(val2) if val2 is not None else "MISSING"

        if val1 is None and val2 is not None:
            diff_summary.append(f"• {field.replace('_', ' ').title()}: MISSING vs {val2_display}")
        elif val1 is not None and val2 is None:
            diff_summary.append(f"• {field.replace('_', ' ').title()}: {val1_display} vs MISSING")
        else:
            diff_summary.append(f"• {field.replace('_', ' ').title()}: {val1_display} vs {val2_display}")

    # === Compare equipment lists ===
    equip1 = {item['name'].strip() for item in offer1.options_list + offer1.accessories_list}
    equip2 = {item['name'].strip() for item in offer2.options_list + offer2.accessories_list}
    
    added_equip = equip2 - equip1
    removed_equip = equip1 - equip2

    if added_equip:
        diff_summary.append(f"• Equipment Added: {', '.join(sorted(list(added_equip)))}")
    if removed_equip:
        diff_summary.append(f"• Equipment Removed: {', '.join(sorted(list(removed_equip)))}")

    return "\n".join(diff_summary) if diff_summary else "No significant differences found."


def consolidate_names(offers: List[ParsedOffer]) -> Tuple[str, str]:
    """Consolidate customer and driver names from a list of parsed offers."""
    common_customer = None
    driver_name = None

    for offer in offers:
        if offer.driver_name:
            driver_name = offer.driver_name
            break

    customer_names = [o.customer for o in offers if o.customer]
    if customer_names:
        first_name = customer_names[0].split()[0]
        if all(name.startswith(first_name) for name in customer_names):
            common_customer = first_name
        else:
            common_customer = customer_names[0]

    return common_customer, driver_name

def generate_excel_report(offers: List[ParsedOffer], template_buffer: io.BytesIO, user_mapping: Dict[str, str]) -> io.BytesIO:
    """Generate Excel report based on the provided template and parsed offers."""
    try:
        template_df = pd.read_excel(template_buffer)
    except Exception as e:
        raise ValueError(f"Failed to read Excel template. Error: {e}")

    offer_data_list = []
    for offer in offers:
        offer_dict = asdict(offer)
        upfront_costs = (offer.upfront_costs or 0) + (offer.deposit or 0) + (offer.admin_fees or 0)
        offer_dict['total_contract_cost'] = (offer.monthly_rental * offer.offer_duration_months) + upfront_costs if offer.monthly_rental and offer.offer_duration_months else None
        offer_data_list.append(offer_dict)

    reference_offer = offers[0]
    other_offers = offers[1:]
    vendors = [offer.get('vendor', 'Unknown Vendor') for offer in offer_data_list]
    final_report_df_rows = []
    final_report_df_rows.append(['Leasing company'] + vendors)

    # Define fields that are titles and should have empty rows
    TITLE_ONLY_FIELDS = [
        'Investment', 'Taxation', 'Duration & Mileage', 'Financial rate', 
        'Service rate', 'Monthly fee', 'Excess / unused km', 'Equipment'
    ]

    for index, row in template_df.iterrows():
        template_field = row['Field']

        if template_field in ['Leasing company', 'Winner']:
            continue
        
        # Handle title-only fields
        if template_field in TITLE_ONLY_FIELDS:
            final_report_df_rows.append([''] * (len(vendors) + 1)) # Blank separator row
            final_report_df_rows.append([template_field] + [''] * len(vendors))
            continue

        # Add a blank row before specific sections
        if template_field in ['Driver name', 'Vehicle Description']:
             final_report_df_rows.append([''] * (len(vendors) + 1))

        # Handle aggregated equipment fields
        if template_field == 'Additional equipment':
            new_row = [template_field]
            for offer in offers:
                all_names = [item['name'] for item in offer.options_list + offer.accessories_list]
                new_row.append(", ".join(all_names) if all_names else None)
            final_report_df_rows.append(new_row)
            continue

        if template_field == 'Additional equipment price':
            new_row = [template_field]
            for offer in offers:
                all_prices = [item.get('price', 0) or 0 for item in offer.options_list + offer.accessories_list]
                total_price = sum(all_prices)
                if total_price == 0:
                    total_price = (offer.options_price or 0) + (offer.accessories_price or 0)
                new_row.append(total_price if total_price > 0 else None)
            final_report_df_rows.append(new_row)
            continue

        # Default behavior for regular data fields
        new_row = [template_field]
        for offer in offer_data_list:
            llm_field_name = user_mapping.get(template_field)
            val = None
            if llm_field_name:
                try:
                    if llm_field_name in offer:
                        if template_field == 'Mileage per year (in km)':
                            if offer.get('offer_duration_months') and offer.get('offer_total_mileage'):
                                val = int(offer.get('offer_total_mileage') / (offer.get('offer_duration_months') / 12))
                            else:
                                val = None
                        elif template_field == 'Total monthly service rate':
                            val = sum([
                                offer.get('maintenance_repair', 0) or 0, offer.get('roadside_assistance', 0) or 0,
                                offer.get('insurance_cost', 0) or 0, offer.get('green_tax', 0) or 0,
                                offer.get('management_fee', 0) or 0, offer.get('tyres_cost', 0) or 0
                            ])
                            if val == 0: val = None
                        else:
                             val = offer.get(llm_field_name)
                except (ValueError, TypeError):
                    val = "N/A"
            new_row.append(val if val is not None and val != '' else "MISSING")
        final_report_df_rows.append(new_row)

    final_report_df = pd.DataFrame(final_report_df_rows, columns=['Field'] + vendors)

    if len(offers) > 1:
        row_index = final_report_df[final_report_df['Field'] == 'Total net investment'].index
        if not row_index.empty:
            insert_idx = row_index[0] + 1
            rows_to_insert = [
                [''] * (len(vendors) + 1),
                ['Vehicle description correspondence', '100.0%'] + [f"{calculate_similarity_score(reference_offer.vehicle_description, o.vehicle_description):.1f}%" for o in other_offers],
                [''] * (len(vendors) + 1),
                ['Gap analysis', 'N/A'] + [get_offer_diff(reference_offer, o) for o in other_offers]
            ]
            insert_df = pd.DataFrame(rows_to_insert, columns=final_report_df.columns)
            final_report_df = pd.concat([final_report_df.iloc[:insert_idx], insert_df, final_report_df.iloc[insert_idx:]]).reset_index(drop=True)

    # Cost Analysis section
    final_report_df.loc[len(final_report_df)] = [''] * len(final_report_df.columns)
    final_report_df.loc[len(final_report_df)] = ['Cost Analysis (excl. VAT)'] + [''] * (len(final_report_df.columns) - 1)

    cost_data = OfferComparator(offers).calculate_total_costs()
    original_vendor_order = [offer.get('vendor', 'Unknown Vendor') for offer in offer_data_list]
    cost_df = pd.DataFrame(cost_data).set_index('vendor')
    sorted_cost_df = cost_df.loc[original_vendor_order].reset_index()
    min_cost = sorted_cost_df['total_contract_cost'].min()

    total_cost_row = ['Total Cost (excl. VAT)'] + [row['total_contract_cost'] for _, row in sorted_cost_df.iterrows()]
    monthly_cost_row = ['Monthly Cost (excl. VAT)'] + [row['cost_per_month'] for _, row in sorted_cost_df.iterrows()]
    winner_row = ['Winner'] + ["🥇 Winner" if row['total_contract_cost'] == min_cost else "" for _, row in sorted_cost_df.iterrows()]
    summary_df = pd.DataFrame([total_cost_row, monthly_cost_row, winner_row], columns=final_report_df.columns)
    final_report_df = pd.concat([final_report_df, summary_df], ignore_index=True)
    
    # --- EXCEL GENERATION AND FORMATTING ---
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        final_report_df.to_excel(writer, sheet_name='Quotation', index=False, header=False)
        workbook = writer.book
        worksheet = writer.sheets['Quotation']

        # Define formats
        bold_format = workbook.add_format({'bold': True})
        winner_format = workbook.add_format({'bold': True, 'bg_color': '#C6EFCE', 'font_color': '#006100'})
        wrap_format = workbook.add_format({'text_wrap': True, 'valign': 'top'})
        green_highlight_match = workbook.add_format({'bg_color': '#C6EFCE'})
        red_highlight_mismatch = workbook.add_format({'bg_color': '#FFC7CE'})
        orange_highlight_variation = workbook.add_format({'bg_color': '#FFEB9C'})

        # Define fields for spec comparison formatting
        spec_fields_to_format = [
            'Vehicle Description', 'Manufacturer', 'Model', 'Version', 'Internal colour', 
            'External colour', 'Fuel type', 'No. doors', 'Number of gears', 'HP', 
            'C02 emission WLTP (g/km)', 'Battery range', 'Additional equipment', 
            'Additional equipment price'
        ]

        # Determine "Taxation value" formatting
        tax_row_values = final_report_df[final_report_df['Field'] == 'Taxation value'].iloc[0, 1:].tolist()
        numeric_tax_values = [float(v) for v in tax_row_values if isinstance(v, (int, float))]
        tax_format_to_apply = None
        if len(numeric_tax_values) > 1:
            diff = max(numeric_tax_values) - min(numeric_tax_values)
            if diff == 0:
                tax_format_to_apply = green_highlight_match
            elif diff < 1:
                tax_format_to_apply = orange_highlight_variation
            else:
                tax_format_to_apply = red_highlight_mismatch
        
        # Find winner column
        winner_col_idx = -1
        for i, val in enumerate(winner_row):
            if val == "🥇 Winner":
                winner_col_idx = i
                break

        # Apply formatting row by row
        for r_idx, row in enumerate(final_report_df.values):
            field_name = str(row[0])
            if field_name in TITLE_ONLY_FIELDS + ['Leasing company', 'Driver name', 'Vehicle Description', 'Cost Analysis (excl. VAT)', 'Gap analysis', 'Vehicle description correspondence']:
                worksheet.write(r_idx, 0, field_name, bold_format)
            
            if field_name in ['Gap analysis', 'Additional equipment']:
                for c_idx in range(1, len(row)):
                    worksheet.write(r_idx, c_idx, row[c_idx], wrap_format)

            if field_name == 'Taxation value' and tax_format_to_apply:
                for c_idx in range(1, len(row)):
                    worksheet.write(r_idx, c_idx, row[c_idx], tax_format_to_apply)
            
            # Apply spec comparison formatting
            if field_name in spec_fields_to_format and len(row) > 2:
                ref_val_str = str(row[1] or '').strip().lower()
                worksheet.write(r_idx, 1, row[1], green_highlight_match) # Color reference cell green
                
                for c_idx in range(2, len(row)):
                    current_val = row[c_idx]
                    current_val_str = str(current_val or '').strip().lower()
                    
                    # Check for a perfect match (case-insensitive)
                    if current_val_str == ref_val_str:
                        worksheet.write(r_idx, c_idx, current_val, green_highlight_match)
                    
                    # Check for a partial match (substring)
                    elif ref_val_str in current_val_str or current_val_str in ref_val_str:
                        worksheet.write(r_idx, c_idx, current_val, orange_highlight_variation)
                    
                    # If neither is true, it's a mismatch
                    else:
                        worksheet.write(r_idx, c_idx, current_val, red_highlight_mismatch)

            if winner_col_idx != -1:
                if field_name in ['Total Cost (excl. VAT)', 'Monthly Cost (excl. VAT)', 'Winner']:
                    worksheet.write(r_idx, winner_col_idx, row[winner_col_idx], winner_format)
                    worksheet.write(0, winner_col_idx, final_report_df.iloc[0, winner_col_idx], winner_format)

        worksheet.set_column(0, 0, 40)
        for i in range(1, len(final_report_df.columns)):
            worksheet.set_column(i, i, 25)

    buffer.seek(0)
    return buffer

def display_parsing_results(offers: List[ParsedOffer]):
    """Display parsing results summary"""
    st.header("📊 Parsing Results")
    col1, col2, col3 = st.columns(3)
    with col1:
        avg_confidence = np.mean([o.parsing_confidence for o in offers if o.parsing_confidence is not None])
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
                st.warning("⚠️ Warnings:")
                for w in offer.warnings:
                    st.write(f"- {w}")
            
            st.write("---")
            st.write("**Extracted Data**")
            st.json(asdict(offer))

if __name__ == '__main__':
    main()
