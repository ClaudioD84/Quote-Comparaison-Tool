"""
AI-Powered Fleet Leasing Offer Comparator - Streamlit App
This version uses a Large Large Model (LLM) to intelligently parse PDF content.
Author: Fleet Management Tool
Version: 2.0 (Refactored for maintainability, performance, and UX)
Requirements:
  streamlit, pandas, numpy, pdfplumber, python-dateutil, xlsxwriter, openpyxl, google-generativeai
Notes:
  - This version uses a real API call to the Google Gemini API.
  - You must provide a valid API key to use the parsing functionality.
  - For deployed apps, store the key in Streamlit secrets as "GOOGLE_API_KEY".
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

    # Use st.secrets for professional deployment, with a fallback to manual input
    try:
        api_key = st.secrets["GOOGLE_API_KEY"]
        st.sidebar.success("✅ API Key loaded from secrets.", icon=" L")
    except (FileNotFoundError, KeyError):
        st.sidebar.warning("API Key not found in secrets. Please enter it manually.")
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
        uploaded_files = [f for f in [reference_file] + other_files if f is not None]
        current_file_names = [f.name for f in uploaded_files]

        if 'offers' not in st.session_state or st.session_state.get('uploaded_files') != current_file_names:
            if not api_key:
                st.error("❌ Please enter your Google AI API Key in the sidebar to proceed.")
                st.stop()

            try:
                parser = LLMParser(api_key=api_key)
                st.session_state.offers = process_offers_internal(_parser=parser, uploaded_files=uploaded_files)
                st.session_state.uploaded_files = current_file_names
            except ValueError as e:
                st.error(f"❌ Initialization Error: {e}")
                st.stop()


        if st.session_state.get('offers'):
            offers = st.session_state.offers
            tab1, tab2, tab3 = st.tabs(["📊 Parsing Results", "🔍 Gap & Spec Analysis", "💰 Cost Comparison"])

            with tab1:
                display_parsing_results(offers)

            with tab2:
                display_gap_analysis(offers)

            with tab3:
                display_cost_comparison(offers)


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

    if st.sidebar.button("Generate & Download Report", help="Click to generate the final Excel report", type="primary"):
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
            with st.spinner("Generating Excel report..."):
                template_buffer = create_default_template()
                excel_buffer = generate_excel_report(offers, template_buffer, user_mapping)
                common_customer, common_driver = consolidate_names(offers)
                customer_name = common_customer if common_customer else "Customer"
                driver_name = common_driver if common_driver else "Driver"
                file_name = f"{customer_name}_{driver_name}".replace(" ", "_")

            st.sidebar.download_button(
                label="⬇️ Download Excel Report",
                data=excel_buffer,
                file_name=f"{file_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.sidebar.success("✅ Report generated successfully!")
        except Exception as e:
            st.error(f"❌ Error generating Excel report: {str(e)}")
            logger.error(f"Excel generation error: {e}\n{traceback.format_exc()}")

@st.cache_data
def process_offers_internal(_parser: LLMParser, uploaded_files: List[st.runtime.uploaded_file_manager.UploadedFile]) -> List[ParsedOffer]:
    """Helper function to process offers and return the list of parsed objects. Cached for performance."""
    offers = []
    progress_bar = st.progress(0, "Initializing AI processing...")
    
    for i, uploaded_file in enumerate(uploaded_files):
        if uploaded_file is None:
            continue
        progress_text = f"Processing {uploaded_file.name} with AI... ({i+1}/{len(uploaded_files)})"
        progress_bar.progress((i + 1) / len(uploaded_files), text=progress_text)
        try:
            pdf_bytes = uploaded_file.getvalue() # Use getvalue() for cached function
            raw_text = TextProcessor.extract_text_from_pdf(pdf_bytes)
            offer = _parser.parse_text(raw_text, uploaded_file.name)
            offers.append(offer)
        except Exception as e:
            st.error(f"❌ Error processing {uploaded_file.name}: {str(e)}")
            logger.error(f"File processing error: {e}\n{traceback.format_exc()}")
            
    progress_bar.empty()

    if not offers or not any(o.parsing_confidence > 0 for o in offers):
        st.error("❌ No offers could be processed successfully. Please check the file format or API key.")
        return []
    return offers

def create_default_template() -> io.BytesIO:
    """Create a default Excel template file for demonstration."""
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
        'Service rate', 'Maintenance & repair', 'Road side assistance', 'Insurance', 'Management fee', 'Tyres (summer and winter)', 
        'Total monthly service rate',
        'Monthly fee', 'Total monthly lease ex. VAT',
        'Excess / unused km', 'Excess kilometers', 'Unused kilometers',
        'Winner'
    ]
    template_data = {'Field': fields, 'Value': [None] * len(fields)}
    df = pd.DataFrame(template_data)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Quotation', index=False)
    buffer.seek(0)
    return buffer

def calculate_similarity_score(s1: str, s2: str) -> float:
    """Calculates a robust similarity score between two strings."""
    def preprocess(text: str) -> str:
        text = text.lower()
        text = re.sub(r'[^a-z0-9\s]', '', text)
        common_words = {'el', 'km', 'h', 'hp', 'd', 'f', 'gs', 'sky', 'hk', 'auto', 'farve', 'color'}
        tokens = [word for word in text.split() if word not in common_words]
        return " ".join(tokens)

    s1_preprocessed = preprocess(str(s1 or ''))
    s2_preprocessed = preprocess(str(s2 or ''))
    matcher = difflib.SequenceMatcher(None, s1_preprocessed, s2_preprocessed)
    return matcher.ratio() * 100

def get_offer_diff(offer1: ParsedOffer, offer2: ParsedOffer) -> str:
    """Compares two ParsedOffer objects and returns a string summarizing the differences."""
    diff_summary = []
    SIMILARITY_THRESHOLD = 90.0
    ELECTRIC_SYNONYMS = {'bev', 'electric', 'battery electric vehicle', 'electricity'}

    fields_to_compare = [
        'vehicle_description', 'manufacturer', 'model', 'version',
        'internal_colour', 'external_colour', 'offer_duration_months', 
        'offer_total_mileage', 'currency', 'taxation_value', 'green_tax', 'fuel_type',
    ]
    for field in fields_to_compare:
        val1, val2 = getattr(offer1, field), getattr(offer2, field)
        val1_str, val2_str = str(val1 or '').strip().lower(), str(val2 or '').strip().lower()
        if val1_str == val2_str: continue
        
        if field in ['vehicle_description', 'version'] and (val1_str in val2_str or val2_str in val1_str or calculate_similarity_score(val1_str, val2_str) >= SIMILARITY_THRESHOLD): continue
        if field == 'currency' and normalize_currency(val1_str) == normalize_currency(val2_str): continue
        if field == 'fuel_type' and val1_str in ELECTRIC_SYNONYMS and val2_str in ELECTRIC_SYNONYMS: continue
        if field == 'green_tax': continue

        val1_display = str(val1) if val1 is not None else "MISSING"
        val2_display = str(val2) if val2 is not None else "MISSING"
        diff_summary.append(f"• {field.replace('_', ' ').title()}: {val1_display} vs {val2_display}")

    equip1 = {item['name'].strip() for item in offer1.options_list + offer1.accessories_list}
    equip2 = {item['name'].strip() for item in offer2.options_list + offer2.accessories_list}
    if added := equip2 - equip1: diff_summary.append(f"• Equipment Added: {', '.join(sorted(list(added)))}")
    if removed := equip1 - equip2: diff_summary.append(f"• Equipment Removed: {', '.join(sorted(list(removed)))}")

    return "\n".join(diff_summary) if diff_summary else "No significant differences found."

def consolidate_names(offers: List[ParsedOffer]) -> Tuple[str, str]:
    """Consolidate customer and driver names from a list of parsed offers."""
    driver_name = next((o.driver_name for o in offers if o.driver_name), None)
    customer_names = [o.customer for o in offers if o.customer]
    common_customer = None
    if customer_names:
        first_name = customer_names[0].split()[0]
        common_customer = first_name if all(name.startswith(first_name) for name in customer_names) else customer_names[0]
    return common_customer, driver_name

def _safe_float_convert(val: Any) -> Optional[float]:
    """Converts a value to a float, handling common European number formats."""
    if isinstance(val, (int, float)): return float(val)
    if isinstance(val, str):
        try: return float(val.replace('.', '').replace(',', '.'))
        except (ValueError, TypeError): return None
    return None

def _prepare_main_data(offers: List[ParsedOffer], template_df: pd.DataFrame, user_mapping: Dict[str, str]) -> pd.DataFrame:
    """Prepares the main data section of the report."""
    offer_data_list = [asdict(offer) for offer in offers]
    vendors = [offer.get('vendor', 'Unknown Vendor') for offer in offer_data_list]
    
    final_rows = [['Leasing company'] + vendors]
    
    TITLE_ONLY_FIELDS = ['Investment', 'Taxation', 'Duration & Mileage', 'Financial rate', 'Service rate', 'Monthly fee', 'Excess / unused km', 'Equipment']
    FIELDS_TO_REMOVE = ['Electricity cost*', 'EV charging station at home*', 'Green tax*']
    SERVICE_RATE_FIELDS = ['maintenance_repair', 'roadside_assistance', 'insurance_cost', 'management_fee', 'tyres_cost']
    ZERO_MEANS_MISSING = ['maintenance_repair', 'roadside_assistance', 'management_fee', 'tyres_cost']

    for _, row in template_df.iterrows():
        template_field = row['Field']
        if template_field in ['Leasing company', 'Winner'] or template_field in FIELDS_TO_REMOVE: continue
        if template_field in TITLE_ONLY_FIELDS:
            final_rows.extend([[''] * (len(vendors) + 1), [template_field] + [''] * len(vendors)])
            continue
        if template_field in ['Driver name', 'Vehicle Description']: final_rows.append([''] * (len(vendors) + 1))

        if template_field == 'Additional equipment':
            row_data = [", ".join(item['name'] for item in offer.options_list + offer.accessories_list) or None for offer in offers]
        elif template_field == 'Additional equipment price':
            row_data = []
            for offer in offers:
                total_price = sum(_safe_float_convert(item.get('price', 0)) or 0 for item in offer.options_list + offer.accessories_list)
                if total_price == 0: total_price = (_safe_float_convert(offer.options_price) or 0) + (_safe_float_convert(offer.accessories_price) or 0)
                row_data.append(total_price if total_price > 0 else None)
        elif template_field == 'Total monthly service rate':
            row_data = [sum(_safe_float_convert(offer.get(f, 0)) or 0 for f in SERVICE_RATE_FIELDS) or "MISSING" for offer in offer_data_list]
        else:
            row_data = []
            llm_field_name = user_mapping.get(template_field)
            for offer in offer_data_list:
                val = None
                if llm_field_name:
                    if template_field == 'Mileage per year (in km)':
                        val = int(offer['offer_total_mileage'] / (offer['offer_duration_months'] / 12)) if offer.get('offer_duration_months') and offer.get('offer_total_mileage') else None
                    else:
                        val = offer.get(llm_field_name)
                        if llm_field_name in ZERO_MEANS_MISSING and val == 0: val = "MISSING"
                row_data.append(val if val is not None and val != '' else "MISSING")
        
        final_rows.append([template_field] + row_data)
        
    return pd.DataFrame(final_rows, columns=['Field'] + vendors)

def _calculate_gap_analysis_rows(reference_offer: ParsedOffer, other_offers: List[ParsedOffer], num_vendors: int) -> List[List[Any]]:
    """Calculates the gap analysis section."""
    if not other_offers: return []
    return [
        [''] * (num_vendors + 1),
        ['Vehicle description correspondence', '100.0%'] + [f"{calculate_similarity_score(reference_offer.vehicle_description, o.vehicle_description):.1f}%" for o in other_offers],
        [''] * (num_vendors + 1),
        ['Gap analysis', 'N/A'] + [get_offer_diff(reference_offer, o) for o in other_offers]
    ]

def _calculate_cost_analysis_df(offers: List[ParsedOffer], original_vendor_order: List[str]) -> pd.DataFrame:
    """Calculates the cost analysis section."""
    cost_data = OfferComparator(offers).calculate_total_costs()
    cost_df = pd.DataFrame(cost_data).set_index('vendor').loc[original_vendor_order].reset_index()
    min_cost = cost_df['total_contract_cost'].min()
    
    rows = [
        [''] * (len(offers) + 1),
        ['Cost Analysis (excl. VAT)'] + [''] * len(offers),
        ['Total Cost (excl. VAT)'] + cost_df['total_contract_cost'].tolist(),
        ['Monthly Cost (excl. VAT)'] + cost_df['cost_per_month'].tolist(),
        ['Winner'] + ["🥇 Winner" if cost == min_cost else "" for cost in cost_df['total_contract_cost']]
    ]
    return pd.DataFrame(rows)

def _apply_excel_formatting(writer: pd.ExcelWriter, df: pd.DataFrame):
    """Applies all xlsxwriter formatting to the final report."""
    workbook = writer.book
    worksheet = writer.sheets['Quotation']
    
    # Define formats
    formats = {
        'bold': workbook.add_format({'bold': True}),
        'winner': workbook.add_format({'bold': True, 'bg_color': '#C6EFCE', 'font_color': '#006100'}),
        'wrap': workbook.add_format({'text_wrap': True, 'valign': 'top'}),
        'green': workbook.add_format({'bg_color': '#C6EFCE'}),
        'red': workbook.add_format({'bg_color': '#FFC7CE'}),
        'orange': workbook.add_format({'bg_color': '#FFEB9C'})
    }
    
    # Find winner column
    winner_row = df[df[0] == 'Winner'].values.flatten().tolist()
    winner_col_idx = winner_row.index("🥇 Winner") if "🥇 Winner" in winner_row else -1

    # Apply formatting
    for r_idx, row in enumerate(df.values):
        field_name = str(row[0])
        if field_name in ['Investment', 'Taxation', 'Duration & Mileage', 'Financial rate', 'Service rate', 'Monthly fee', 'Excess / unused km', 'Equipment', 'Leasing company', 'Driver name', 'Vehicle Description', 'Cost Analysis (excl. VAT)', 'Gap analysis', 'Vehicle description correspondence']:
            worksheet.write(r_idx, 0, field_name, formats['bold'])
        
        if field_name in ['Gap analysis', 'Additional equipment']:
            for c_idx in range(1, len(row)): worksheet.write(r_idx, c_idx, row[c_idx], formats['wrap'])
        
        # Spec comparison formatting
        spec_fields = ['Vehicle Description', 'Manufacturer', 'Model', 'Version', 'Internal colour', 'External colour', 'Fuel type', 'No. doors', 'Number of gears', 'HP', 'C02 emission WLTP (g/km)', 'Battery range', 'Additional equipment', 'Additional equipment price']
        if field_name in spec_fields and len(row) > 2:
            ref_val_str = str(row[1] or '').strip().lower()
            worksheet.write(r_idx, 1, row[1], formats['green'])
            for c_idx in range(2, len(row)):
                curr_val_str = str(row[c_idx] or '').strip().lower()
                fmt = formats['red']
                if curr_val_str == ref_val_str: fmt = formats['green']
                elif ref_val_str in curr_val_str or curr_val_str in ref_val_str: fmt = formats['orange']
                worksheet.write(r_idx, c_idx, row[c_idx], fmt)

        # Taxation value formatting
        if field_name == 'Taxation value':
            numeric_vals = [float(v) for v in row[1:] if isinstance(v, (int, float))]
            if len(numeric_vals) > 1:
                diff = max(numeric_vals) - min(numeric_vals)
                fmt = formats['red'] if diff >= 1 else (formats['orange'] if diff > 0 else formats['green'])
                for c_idx in range(1, len(row)): worksheet.write(r_idx, c_idx, row[c_idx], fmt)
        
        # Winner column formatting
        if winner_col_idx != -1 and field_name in ['Total Cost (excl. VAT)', 'Monthly Cost (excl. VAT)', 'Winner']:
            worksheet.write(r_idx, winner_col_idx, row[winner_col_idx], formats['winner'])
            worksheet.write(0, winner_col_idx, df.iloc[0, winner_col_idx], formats['winner'])

    # Set column widths
    worksheet.set_column(0, 0, 40)
    for i in range(1, len(df.columns)): worksheet.set_column(i, i, 25)

def generate_excel_report(offers: List[ParsedOffer], template_buffer: io.BytesIO, user_mapping: Dict[str, str]) -> io.BytesIO:
    """Orchestrates the generation of the final Excel report."""
    template_df = pd.read_excel(template_buffer)
    
    # 1. Prepare main data
    final_report_df = _prepare_main_data(offers, template_df, user_mapping)
    
    # 2. Calculate and insert gap analysis
    if len(offers) > 1:
        row_index = final_report_df[final_report_df['Field'] == 'Total net investment'].index
        if not row_index.empty:
            insert_idx = row_index[0] + 1
            gap_rows = _calculate_gap_analysis_rows(offers[0], offers[1:], len(offers))
            insert_df = pd.DataFrame(gap_rows, columns=final_report_df.columns)
            final_report_df = pd.concat([final_report_df.iloc[:insert_idx], insert_df, final_report_df.iloc[insert_idx:]]).reset_index(drop=True)
            
    # 3. Calculate and append cost analysis
    original_vendor_order = final_report_df.columns[1:].tolist()
    cost_df = _calculate_cost_analysis_df(offers, original_vendor_order)
    cost_df.columns = final_report_df.columns
    final_report_df = pd.concat([final_report_df, cost_df], ignore_index=True)

    # 4. Generate Excel file and apply formatting
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        final_report_df.to_excel(writer, sheet_name='Quotation', index=False, header=False)
        _apply_excel_formatting(writer, final_report_df)
        
    buffer.seek(0)
    return buffer

def display_parsing_results(offers: List[ParsedOffer]):
    """Display parsing results summary in the Streamlit app."""
    st.header("📊 AI Parsing Results")
    col1, col2, col3 = st.columns(3)
    avg_confidence = np.mean([o.parsing_confidence for o in offers if o.parsing_confidence is not None])
    warning_count = sum(len(o.warnings) for o in offers)
    col1.metric("Average Confidence", f"{avg_confidence:.1%}")
    col2.metric("Total Warnings", warning_count)
    col3.metric("AI Model", "Gemini 2.5 Pro")

    with st.expander("📋 View Detailed Extracted Data (JSON)"):
        for offer in offers:
            st.subheader(f"📄 {offer.vendor or offer.filename}")
            st.json(asdict(offer))

def display_gap_analysis(offers: List[ParsedOffer]):
    """Display gap and specification analysis in the Streamlit app."""
    st.header("🔍 Gap & Specification Analysis")
    if len(offers) < 2:
        st.info("Upload at least two offers to perform a gap analysis.")
        return

    reference_offer = offers[0]
    other_offers = offers[1:]
    
    st.subheader("Vehicle Description Similarity")
    st.markdown(f"Comparing all offers against the reference: **{reference_offer.vehicle_description}**")
    
    cols = st.columns(len(other_offers))
    for i, offer in enumerate(other_offers):
        score = calculate_similarity_score(reference_offer.vehicle_description, offer.vehicle_description)
        with cols[i]:
            st.metric(label=f"vs. {offer.vendor}", value=f"{score:.1f}%")
            
    st.subheader("Key Differences Detected")
    for offer in other_offers:
        st.markdown(f"---")
        st.markdown(f"#### Gaps between `{reference_offer.vendor}` and `{offer.vendor}`")
        diff_text = get_offer_diff(reference_offer, offer)
        if diff_text == "No significant differences found.":
            st.success(f"✅ {diff_text}")
        else:
            st.text(diff_text)
            
def display_cost_comparison(offers: List[ParsedOffer]):
    """Display cost comparison in the Streamlit app."""
    st.header("💰 Cost Comparison")
    comparator = OfferComparator(offers)
    is_valid, errors = comparator.validate_offers()

    if not is_valid:
        st.warning("Offers may not be directly comparable due to inconsistencies.")
        for error in errors:
            st.error(f"• {error}")
        return

    report_df = comparator.generate_comparison_report()
    if report_df.empty:
        st.error("Could not generate a cost comparison report.")
        return
        
    # Reorder columns for better readability
    display_cols = ['rank', 'vendor', 'total_contract_cost', 'cost_per_month', 'cost_per_km', 'vehicle', 'duration_months', 'total_mileage', 'currency']
    report_df = report_df[[col for col in display_cols if col in report_df.columns]]
    
    st.dataframe(report_df.style.format({
        'total_contract_cost': '{:,.2f}',
        'cost_per_month': '{:,.2f}',
        'cost_per_km': '{:,.4f}',
        'parsing_confidence': '{:.1%}'
    }), use_container_width=True)


if __name__ == '__main__':
    main()
