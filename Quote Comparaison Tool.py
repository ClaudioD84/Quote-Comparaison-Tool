# -*- coding: utf-8 -*-
"""
AI-Powered Fleet Leasing Offer Comparator - Streamlit App (Version 2.1 - Free Tier Optimized)
This version incorporates best practices for performance, maintainability, and user experience.
It is optimized to use the Gemini 1.0 Pro model for its generous free tier.

Author: Fleet Management Tool (Refactored by Gemini)
Date: 2025-09-16
Requirements:
  streamlit, pandas, pdfplumber, python-dateutil, openpyxl, xlsxwriter, google-generativeai, difflib
"""

import io
import re
import json
import logging
import traceback
from typing import List, Dict, Any, Optional, Tuple, Union
from dataclasses import dataclass, field, asdict
from collections import defaultdict
import difflib

import streamlit as st
import pandas as pd
import numpy as np
import pdfplumber
import google.generativeai as genai


# --- CONFIGURATION & CONSTANTS ---

@st.cache_resource
def setup_logging():
    """Sets up a Streamlit-friendly logger that persists across reruns."""
    logger = logging.getLogger("leasing_comparator")
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        handler = logging.StreamHandler()
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    return logger

logger = setup_logging()

CURRENCY_MAP = {
    'kr.': 'DKK', 'kr': 'DKK', 'dkk': 'DKK',
    '€': 'EUR', 'eur': 'EUR', 'euro': 'EUR',
    '£': 'GBP', 'gbp': 'GBP',
    'chf': 'CHF', 'sek': 'SEK', 'nok': 'NOK',
    'pln': 'PLN', 'huf': 'HUF', 'czk': 'CZK',
}

# --- DATA STRUCTURES ---

@dataclass
class ParsedOffer:
    """Standardized structure for parsed leasing offer data."""
    filename: str
    vendor: Optional[str] = None
    vehicle_description: Optional[str] = None
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
    roadside_assistance: Optional[float] = None
    tyres_cost: Optional[float] = None
    total_monthly_lease: Optional[float] = None
    driver_name: Optional[str] = None
    customer: Optional[str] = None
    options_list: List[Dict[str, Union[str, float]]] = field(default_factory=list)
    accessories_list: List[Dict[str, Union[str, float]]] = field(default_factory=list)

# --- CORE LOGIC CLASSES ---

class TextProcessor:
    """Handles text extraction from PDF files."""
    @staticmethod
    def extract_text_from_pdf(pdf_bytes: bytes) -> str:
        """Extracts all text from a PDF file's bytes."""
        try:
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                return "\n".join(page.extract_text() or "" for page in pdf.pages)
        except Exception as e:
            logger.error(f"PDF text extraction failed: {e}")
            return ""

class LLMParser:
    """Uses the Gemini LLM to parse text into a structured ParsedOffer object."""
    def __init__(self, api_key: str):
        if not api_key:
            raise ValueError("An API key for the Gemini API is required.")
        genai.configure(api_key=api_key)
        logger.info("Gemini client configured successfully.")
        
        # Define the JSON schema once during initialization
        self.json_schema = {
            "type": "OBJECT",
            "properties": {
                "vendor": {"type": "STRING"}, "vehicle_description": {"type": "STRING"},
                "max_duration_months": {"type": "NUMBER"}, "max_total_mileage": {"type": "NUMBER"},
                "offer_duration_months": {"type": "NUMBER"}, "offer_total_mileage": {"type": "NUMBER"},
                "monthly_rental": {"type": "NUMBER"}, "upfront_costs": {"type": "NUMBER"},
                "deposit": {"type": "NUMBER"}, "admin_fees": {"type": "NUMBER"},
                "maintenance_included": {"type": "BOOLEAN"}, "excess_mileage_rate": {"type": "NUMBER"},
                "unused_mileage_rate": {"type": "NUMBER"}, "currency": {"type": "STRING"},
                "parsing_confidence": {"type": "NUMBER", "description": "Confidence (0.0 to 1.0) in the extracted data."},
                "warnings": {"type": "ARRAY", "items": {"type": "STRING"}},
                "quote_number": {"type": "STRING"}, "manufacturer": {"type": "STRING"},
                "model": {"type": "STRING"}, "version": {"type": "STRING"},
                "internal_colour": {"type": "STRING"}, "external_colour": {"type": "STRING"},
                "fuel_type": {"type": "STRING"}, "num_doors": {"type": "NUMBER"},
                "hp": {"type": "NUMBER"}, "c02_emission": {"type": "NUMBER"},
                "battery_range": {"type": "NUMBER"}, "vehicle_price": {"type": "NUMBER"},
                "options_price": {"type": "NUMBER"}, "accessories_price": {"type": "NUMBER"},
                "delivery_cost": {"type": "NUMBER"}, "registration_tax": {"type": "NUMBER"},
                "total_net_investment": {"type": "NUMBER"}, "taxation_value": {"type": "NUMBER"},
                "financial_rate": {"type": "NUMBER"}, "depreciation_interest": {"type": "NUMBER"},
                "maintenance_repair": {"type": "NUMBER"}, "insurance_cost": {"type": "NUMBER"},
                "green_tax": {"type": "NUMBER"}, "management_fee": {"type": "NUMBER"},
                "tyres_cost": {"type": "NUMBER"}, "roadside_assistance": {"type": "NUMBER"},
                "total_monthly_lease": {"type": "NUMBER"}, "driver_name": {"type": "STRING"},
                "customer": {"type": "STRING"},
                "options_list": {"type": "ARRAY", "items": {"type": "OBJECT", "properties": {"name": {"type": "STRING"}, "price": {"type": "NUMBER"}}}},
                "accessories_list": {"type": "ARRAY", "items": {"type": "OBJECT", "properties": {"name": {"type": "STRING"}, "price": {"type": "NUMBER"}}}}
            }
        }
        
        generation_config = genai.types.GenerationConfig(
            response_mime_type="application/json",
            response_schema=self.json_schema
        )
        self.model = genai.GenerativeModel(
            model_name='gemini-1.0-pro', # UPDATED to use model with more generous free tier
            generation_config=generation_config
        )

    def parse_text(self, text: str, filename: str) -> ParsedOffer:
        """Sends PDF text to the Gemini API for structured data extraction."""
        logger.info(f"Sending text to Gemini for parsing file: {filename}")
        
        prompt_text = f"""
        You are a world-class financial analyst specializing in fleet leasing. Your task is to extract key data points from a vehicle leasing contract, regardless of the language or format.

        KEY INSTRUCTIONS:
        1.  **Differentiate Contract Terms**:
            - `max_duration_months` & `max_total_mileage`: The maximum possible contract terms (e.g., "Max contract: 60 months / 300,000 km").
            - `offer_duration_months` & `offer_total_mileage`: The terms for this specific offer (e.g., "Current offer: 36 months / 175,000 km").

        2.  **Prices Exclude VAT**: All extracted prices and costs (`monthly_rental`, `total_monthly_lease`, etc.) must be EXCLUDING VAT. Look for terms like "excl. VAT", "HT", "net price".

        3.  **Driver vs. Customer**: `driver_name` is the employee. `customer` is the company renting the car.

        4.  **Calculate Total Mileage**: If only annual mileage is given, calculate the total. Example: "35,000 km per year / 48 months" -> offer_total_mileage is 35000 * (48 / 12) = 140000.

        5.  **Synonyms**: Treat "BEV", "Electric", and "electricity" as the same `fuel_type`. "Total net investment" can also be "Taxable value". Include amounts for "Arval assistance" or "Ayvens assistance" under `roadside_assistance`.

        Return the data as a JSON object strictly following the provided schema. If a value is not found, use `null`. Do not invent values.
        
        <DOCUMENT_TO_PARSE>
        {text}
        </DOCUMENT_TO_PARSE>
        """

        try:
            response = self.model.generate_content(prompt_text)
            extracted_data = json.loads(response.text)
            # Ensure list fields exist even if empty
            extracted_data['options_list'] = extracted_data.get('options_list') or []
            extracted_data['accessories_list'] = extracted_data.get('accessories_list') or []
            return ParsedOffer(filename=filename, **extracted_data)
        except Exception as e:
            logger.error(f"Gemini API call failed for {filename}: {str(e)}\n{traceback.format_exc()}")
            return ParsedOffer(filename=filename, warnings=[f"LLM parsing failed: {str(e)}"], parsing_confidence=0.1)

class OfferComparator:
    """Handles comparison, validation, and analysis of multiple leasing offers."""
    def __init__(self, offers: List[ParsedOffer]):
        self.offers = offers

    def validate_offers(self) -> Tuple[bool, List[str]]:
        """Validates that offers have consistent terms for a fair comparison."""
        errors = []
        if len(self.offers) < 2:
            return False, ["At least 2 offers are needed for comparison."]

        currencies = {o.currency for o in self.offers if o.currency}
        if len(currencies) > 1:
            errors.append(f"Mixed currencies found: {', '.join(currencies)}")

        durations = {o.offer_duration_months for o in self.offers if o.offer_duration_months}
        if len(durations) > 1:
            errors.append(f"Contract durations do not match: {', '.join(map(str, durations))}")

        mileages = {o.offer_total_mileage for o in self.offers if o.offer_total_mileage}
        if len(mileages) > 1:
            errors.append(f"Contract mileages do not match: {', '.join(map(str, mileages))}")

        return not errors, errors

    def calculate_total_costs(self) -> List[Dict[str, Any]]:
        """Calculates total contract cost and per-unit costs for each offer."""
        results = []
        for offer in self.offers:
            cost_info = {'vendor': offer.vendor or offer.filename, 'error': None}
            if not offer.offer_duration_months or not offer.monthly_rental:
                cost_info['error'] = 'Missing duration or monthly rental for cost calculation.'
                results.append(cost_info)
                continue
            
            monthly_total = offer.monthly_rental * offer.offer_duration_months
            upfront_total = (offer.upfront_costs or 0) + (offer.deposit or 0) + (offer.admin_fees or 0)
            total_cost = monthly_total + upfront_total
            
            results.append({
                'vendor': offer.vendor or offer.filename,
                'vehicle': offer.vehicle_description,
                'total_contract_cost': total_cost,
                'cost_per_month': total_cost / offer.offer_duration_months,
                'cost_per_km': total_cost / offer.offer_total_mileage if offer.offer_total_mileage else None,
                'currency': offer.currency
            })
        return sorted(results, key=lambda x: x.get('total_contract_cost', float('inf')))

# --- UTILITY FUNCTIONS ---

def _safe_float_convert(val: Any) -> Optional[float]:
    """Safely converts a value to a float, handling European number formats."""
    if isinstance(val, (int, float)): return float(val)
    if isinstance(val, str):
        try:
            return float(val.replace('.', '').replace(',', '.'))
        except (ValueError, TypeError):
            return None
    return None

def calculate_similarity_score(s1: Optional[str], s2: Optional[str]) -> float:
    """Calculates a similarity score between two strings, ignoring common irrelevant words."""
    if not s1 or not s2: return 0.0
    
    def preprocess(text: str) -> str:
        text = text.lower()
        text = re.sub(r'[^a-z0-9\s]', '', text)
        common_words = {'el', 'km', 'h', 'hp', 'd', 'f', 'auto', 'color', 'farve'}
        tokens = [word for word in text.split() if word not in common_words]
        return " ".join(tokens)

    return difflib.SequenceMatcher(None, preprocess(s1), preprocess(s2)).ratio() * 100

def get_offer_diff(offer1: ParsedOffer, offer2: ParsedOffer) -> str:
    """Compares two offers and returns a string summarizing the key differences."""
    diffs = []
    fields_to_compare = [
        ('Vehicle Description', 'vehicle_description'), ('Manufacturer', 'manufacturer'),
        ('Model', 'model'), ('Version', 'version'), ('Fuel Type', 'fuel_type'),
        ('External Colour', 'external_colour'), ('Taxation Value', 'taxation_value')
    ]
    for name, field in fields_to_compare:
        val1, val2 = getattr(offer1, field), getattr(offer2, field)
        if str(val1 or '').lower() != str(val2 or '').lower():
            diffs.append(f"• **{name}**: {val1 or 'N/A'} vs **{val2 or 'N/A'}**")

    equip1 = {item['name'].strip().lower() for item in offer1.options_list + offer1.accessories_list}
    equip2 = {item['name'].strip().lower() for item in offer2.options_list + offer2.accessories_list}
    
    if added := equip2 - equip1: diffs.append(f"• **Equipment Added**: {', '.join(sorted(list(added)))}")
    if removed := equip1 - equip2: diffs.append(f"• **Equipment Removed**: {', '.join(sorted(list(removed)))}")

    return "\n".join(diffs) if diffs else "✅ No significant differences found."

def consolidate_names(offers: List[ParsedOffer]) -> Tuple[str, str]:
    """Finds the most likely customer and driver name from all offers."""
    driver_name = next((o.driver_name for o in offers if o.driver_name), "Driver")
    customer_names = [o.customer for o in offers if o.customer]
    customer_name = customer_names[0] if customer_names else "Customer"
    return customer_name, driver_name

# --- DATA PROCESSING (CACHED) ---

@st.cache_data(show_spinner="🧠 Processing PDFs with AI... This may take a moment.")
def process_offers_internal(_parser: LLMParser, uploaded_files: list) -> List[ParsedOffer]:
    """
    Processes uploaded PDF files using the LLMParser.
    This function is cached to avoid reprocessing and re-running expensive API calls.
    """
    offers = []
    for uploaded_file in uploaded_files:
        if uploaded_file is None: continue
        uploaded_file.seek(0)
        pdf_bytes = uploaded_file.read()
        raw_text = TextProcessor.extract_text_from_pdf(pdf_bytes)
        if raw_text:
            offer = _parser.parse_text(raw_text, uploaded_file.name)
            offers.append(offer)
        else:
            st.warning(f"Could not extract text from {uploaded_file.name}. The file might be an image-based PDF.")
    return offers

# --- EXCEL REPORT GENERATION (MODULARIZED) ---

def generate_excel_report(offers: List[ParsedOffer], user_mapping: Dict[str, str]) -> io.BytesIO:
    """Orchestrates the creation of the final Excel report."""
    report_df = _prepare_report_dataframe(offers, user_mapping)
    if len(offers) > 1:
        report_df = _insert_gap_analysis(report_df, offers)
    report_df = _add_cost_analysis(report_df, offers)
    return _format_excel_workbook(report_df)

def _prepare_report_dataframe(offers: List[ParsedOffer], user_mapping: Dict[str, str]) -> pd.DataFrame:
    """Prepares the main data section of the report."""
    vendors = [o.vendor or f"Offer {i+1}" for i, o in enumerate(offers)]
    report_rows = [['Leasing company'] + vendors]
    
    for template_field, llm_field in user_mapping.items():
        if not llm_field:  # Skip if user cleared the mapping
            continue
        
        # Handle special aggregated or calculated fields
        if llm_field == 'section_title':
            report_rows.append([''] * (len(vendors) + 1))
            report_rows.append([template_field] + [''] * len(vendors))
            continue
        
        if llm_field == 'additional_equipment':
            row_data = [template_field] + [", ".join([item['name'] for item in o.options_list + o.accessories_list]) or "MISSING" for o in offers]
            report_rows.append(row_data)
            continue
            
        if llm_field == 'total_monthly_service_rate':
            SERVICE_FIELDS = ['maintenance_repair', 'roadside_assistance', 'insurance_cost', 'management_fee', 'tyres_cost']
            row_data = [template_field]
            for o in offers:
                total_sum = sum(_safe_float_convert(getattr(o, f, 0)) or 0 for f in SERVICE_FIELDS)
                row_data.append(total_sum if total_sum > 0 else "MISSING")
            report_rows.append(row_data)
            continue
        
        # Default behavior for standard fields
        new_row = [template_field]
        ZERO_MEANS_MISSING = ['maintenance_repair', 'roadside_assistance', 'management_fee', 'tyres_cost']
        for o in offers:
            val = "MISSING"
            raw_val = getattr(o, llm_field, None)
            if raw_val is not None and raw_val != '':
                val = "MISSING" if llm_field in ZERO_MEANS_MISSING and raw_val == 0 else raw_val
            new_row.append(val)
        report_rows.append(new_row)
        
    return pd.DataFrame(report_rows, columns=['Field'] + vendors)

def _insert_gap_analysis(report_df: pd.DataFrame, offers: List[ParsedOffer]) -> pd.DataFrame:
    """Inserts vehicle similarity and gap analysis into the report."""
    ref_offer, other_offers = offers[0], offers[1:]
    
    # Find insert index safely
    idx_series = report_df[report_df['Field'] == 'Total net investment'].index
    if idx_series.empty:
        return report_df # Return original df if anchor not found
    insert_idx = idx_series[0] + 1
    
    rows_to_insert = [
        [''] * len(report_df.columns),
        ['Vehicle description correspondence', '100.0%'] + [f"{calculate_similarity_score(ref_offer.vehicle_description, o.vehicle_description):.1f}%" for o in other_offers],
        [''] * len(report_df.columns),
        ['Gap analysis', 'N/A'] + [get_offer_diff(ref_offer, o) for o in other_offers]
    ]
    
    insert_df = pd.DataFrame(rows_to_insert, columns=report_df.columns)
    return pd.concat([report_df.iloc[:insert_idx], insert_df, report_df.iloc[insert_idx:]]).reset_index(drop=True)

def _add_cost_analysis(report_df: pd.DataFrame, offers: List[ParsedOffer]) -> pd.DataFrame:
    """Appends the final cost analysis summary."""
    cost_data = OfferComparator(offers).calculate_total_costs()
    cost_df = pd.DataFrame(cost_data)
    
    # Only proceed if cost calculation was successful
    if 'total_contract_cost' in cost_df.columns and not cost_df['total_contract_cost'].isnull().all():
        min_cost = cost_df['total_contract_cost'].min()
        total_cost_row = ['Total Cost (excl. VAT)'] + cost_df['total_contract_cost'].tolist()
        monthly_cost_row = ['Monthly Cost (excl. VAT)'] + cost_df['cost_per_month'].tolist()
        winner_row = ['Winner'] + ["🥇 Winner" if cost == min_cost else "" for cost in cost_df['total_contract_cost']]
    else:
        # Handle case where cost calculation failed for all offers
        num_offers = len(offers)
        error_msg = "Calculation Failed"
        total_cost_row = ['Total Cost (excl. VAT)'] + [error_msg] * num_offers
        monthly_cost_row = ['Monthly Cost (excl. VAT)'] + [error_msg] * num_offers
        winner_row = ['Winner'] + [""] * num_offers

    summary_rows = [
        [''] * len(report_df.columns),
        ['Cost Analysis (excl. VAT)'] + [''] * (len(report_df.columns) - 1),
        total_cost_row,
        monthly_cost_row,
        winner_row
    ]
    summary_df = pd.DataFrame(summary_rows, columns=report_df.columns)
    return pd.concat([report_df, summary_df], ignore_index=True)

def _format_excel_workbook(report_df: pd.DataFrame) -> io.BytesIO:
    """Writes a DataFrame to a fully formatted Excel file in memory."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        report_df.to_excel(writer, sheet_name='Quotation Comparison', index=False, header=False)
        workbook = writer.book
        worksheet = writer.sheets['Quotation Comparison']

        # Formats
        bold_format = workbook.add_format({'bold': True})
        winner_format = workbook.add_format({'bold': True, 'bg_color': '#C6EFCE', 'font_color': '#006100'})
        wrap_format = workbook.add_format({'text_wrap': True, 'valign': 'top'})
        
        # Apply formatting
        gap_row_idx_series = report_df[report_df['Field'] == 'Gap analysis'].index
        if not gap_row_idx_series.empty:
            worksheet.set_row(gap_row_idx_series[0], 150, wrap_format)
            
        winner_row_series = report_df[report_df['Field'] == 'Winner'].values
        if winner_row_series.size > 0:
            winner_row = winner_row_series.flatten().tolist()
            if "🥇 Winner" in winner_row:
                winner_col_idx = winner_row.index("🥇 Winner")
                for r_idx, field in enumerate(report_df['Field']):
                    if field in ['Total Cost (excl. VAT)', 'Monthly Cost (excl. VAT)', 'Winner', 'Leasing company']:
                        worksheet.write(r_idx, winner_col_idx, report_df.iloc[r_idx, winner_col_idx], winner_format)

        worksheet.set_column(0, 0, 40)
        worksheet.set_column(1, len(report_df.columns) - 1, 25)

    buffer.seek(0)
    return buffer

# --- STREAMLIT UI ---

def display_results_in_app(offers: List[ParsedOffer]):
    """Displays parsing summary and comparison results in the Streamlit UI using tabs."""
    st.header("📊 AI Analysis Results", divider='rainbow')

    tab1, tab2, tab3 = st.tabs(["**Parsing Summary**", "**↔️ Gap Analysis**", "**💰 Cost Comparison**"])

    with tab1:
        st.subheader("Extraction Overview")
        cols = st.columns(len(offers))
        for i, offer in enumerate(offers):
            with cols[i]:
                st.metric(
                    label=f"**{offer.vendor or f'Offer {i+1}'}**",
                    value=f"{offer.parsing_confidence:.1%}",
                    help="AI confidence in the accuracy of the extracted data."
                )
        st.info("The details below show all data extracted by the AI from each document.")
        for offer in offers:
            with st.expander(f"📄 View Extracted Data for **{offer.filename}**"):
                # Display a warning if the API call failed for this offer
                if offer.warnings and "LLM parsing failed" in offer.warnings[0]:
                    st.error(f"AI parsing failed for this document. Error: {offer.warnings[0]}")
                st.json(asdict(offer))

    if len(offers) > 1:
        ref_offer, other_offers = offers[0], offers[1:]
        
        with tab2:
            st.subheader(f"Comparison against Reference: **{ref_offer.vendor or ref_offer.filename}**")
            for i, other_offer in enumerate(other_offers):
                st.markdown(f"---")
                st.markdown(f"#### Analyzing **{other_offer.vendor or other_offer.filename}**")
                similarity = calculate_similarity_score(ref_offer.vehicle_description, other_offer.vehicle_description)
                st.progress(int(similarity), text=f"Vehicle Description Similarity: **{similarity:.1f}%**")
                
                diff = get_offer_diff(ref_offer, other_offer)
                st.markdown(diff)

        with tab3:
            st.subheader("Financial Comparison")
            comparator = OfferComparator(offers)
            is_valid, errors = comparator.validate_offers()
            if not is_valid:
                for error in errors:
                    st.warning(f"⚠️ {error}")
                st.error("Financial comparison may be inaccurate due to inconsistencies.")

            cost_df = pd.DataFrame(comparator.calculate_total_costs())

            # Check to prevent crash if parsing fails
            if 'total_contract_cost' in cost_df.columns:
                cost_df['rank'] = cost_df['total_contract_cost'].rank(method='min').astype(int)
                st.dataframe(cost_df.style.format({
                    'total_contract_cost': '{:,.2f}',
                    'cost_per_month': '{:,.2f}',
                    'cost_per_km': '{:,.2f}',
                }).highlight_min(subset=['total_contract_cost'], color='#D4EDDA'), use_container_width=True)
            else:
                st.error("🔴 Cost comparison could not be performed because the AI failed to extract financial data from the documents, likely due to an API error.")
                st.dataframe(cost_df)

def main():
    """Main function to run the Streamlit application."""
    st.set_page_config(page_title="Fleet Leasing Offer Comparator", page_icon="🚗", layout="wide")
    st.title("🚗 AI-Powered Fleet Leasing Offer Comparator")
    st.markdown("Upload a **reference offer** and one or more **other offers** to automatically extract, compare, and analyze the key terms using AI.")

    with st.sidebar:
        st.header("⚙️ Configuration")
        api_key = st.text_input("Enter your Google AI API Key", type="password", help="For deployed apps, use st.secrets for security.")
        if not api_key:
            st.info("Get your API key from [Google AI Studio](https://aistudio.google.com/app/apikey).")

        mapping_suggestions = {
            "Driver & Vehicle": "section_title", "Quote number": "quote_number", "Driver name": "driver_name",
            "Vehicle Description": "vehicle_description", "Manufacturer": "manufacturer", "Model": "model", 
            "Version": "version", "Internal colour": "internal_colour", "External colour": "external_colour", 
            "Fuel type": "fuel_type", "No. doors": "num_doors", "HP": "hp", 
            "C02 emission WLTP (g/km)": "c02_emission", "Battery range": "battery_range",
            "Equipment": "section_title", "Additional equipment": "additional_equipment",
            "Investment": "section_title", "Vehicle list price (excl. VAT, excl. options)": "vehicle_price", 
            "Options (excl. taxes)": "options_price", "Accessories (excl. taxes)": "accessories_price", 
            "Delivery cost": "delivery_cost", "Registration tax": "registration_tax", 
            "Total net investment": "total_net_investment",
            "Taxation": "section_title", "Taxation value": "taxation_value",
            "Duration & Mileage": "section_title", "Term (months)": "offer_duration_months", 
            "Mileage per year (in km)": "offer_total_mileage",
            "Financial rate": "section_title", "Monthly financial rate (depreciation + interest)": "depreciation_interest",
            "Service rate": "section_title", "Maintenance & repair": "maintenance_repair", "Insurance": "insurance_cost", 
            "Green tax*": "green_tax", "Management fee": "management_fee", "Tyres (summer and winter)": "tyres_cost",
            "Road side assistance": "roadside_assistance", "Total monthly service rate": "total_monthly_service_rate",
            "Monthly fee": "section_title", "Total monthly lease ex. VAT": "total_monthly_lease",
            "Excess / unused km": "section_title", "Excess kilometers": "excess_mileage_rate", 
            "Unused kilometers": "unused_mileage_rate"
        }
        user_mapping = mapping_suggestions.copy()

    st.header("📁 Upload Your Offers", divider='rainbow')
    col1, col2 = st.columns(2)
    with col1:
        reference_file = st.file_uploader("1. Upload the **Reference** Offer", type='pdf')
    with col2:
        other_files = st.file_uploader("2. Upload **Other** Offers (up to 9)", type='pdf', accept_multiple_files=True)

    if reference_file and other_files:
        uploaded_files = [reference_file] + other_files
        
        if not api_key:
            st.error("❌ Please enter your Google AI API Key in the sidebar to proceed.")
            st.stop()

        try:
            parser = LLMParser(api_key=api_key)
            offers = process_offers_internal(parser, uploaded_files)
            
            if offers:
                st.session_state.offers = offers
                display_results_in_app(offers)
                
                st.header("📄 Generate Final Report", divider='rainbow')
                st.info("After reviewing the analysis above, click the button below to generate a formatted Excel spreadsheet with all the details.")
                
                if st.button("Generate Excel Report", type="primary"):
                    with st.spinner("Creating your Excel file..."):
                        excel_buffer = generate_excel_report(offers, user_mapping)
                        customer_name, driver_name = consolidate_names(offers)
                        file_name = f"Leasing_Comparison_{customer_name}_{driver_name}.xlsx".replace(" ", "_")
                        
                        st.download_button(
                            label="⬇️ Download Excel Report",
                            data=excel_buffer,
                            file_name=file_name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
            else:
                st.warning("AI processing did not return any valid offer data. Please check your files and API key.")

        except Exception as e:
            st.error(f"An unexpected error occurred: {e}")
            logger.error(f"Top-level error in main loop: {traceback.format_exc()}")

    elif reference_file:
        st.info("Please upload at least one other offer to begin the comparison.")

if __name__ == '__main__':
    main()
