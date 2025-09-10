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

@dataclass
class ParsedOffer:
    """Standardized structure for parsed leasing offer data"""
    filename: str
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
    currency: Optional[str] = None
    parsing_confidence: float = 0.0
    warnings: List[str] = field(default_factory=list)

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
                        "currency": {"type": "STRING"},
                        "parsing_confidence": {"type": "NUMBER"},
                        "warnings": {"type": "ARRAY", "items": {"type": "STRING"}}
                    }
                }
            }
        }
        
        # Mocking the LLM's response for demonstration
        # In a real application, you would make an HTTP POST request here
        # to the specified API URL with the payload.
        # This mock data is based on the two PDFs from the user's query
        mock_responses = {
            "Kontraktoplæg_3052514001_1 (1).pdf": {
                "vendor": "Ayvens",
                "vehicle_description": "OPEL GRANDLAND EL 210",
                "duration_months": 48,
                "total_mileage": 140000,
                "monthly_rental": 5871.39,
                "upfront_costs": 0,
                "deposit": 0,
                "admin_fees": None,
                "maintenance_included": True,
                "excess_mileage_rate": 0.50,
                "currency": "DKK",
                "parsing_confidence": 0.95,
                "warnings": ["Total mileage calculated from annual mileage"]
            },
            "quotation  2508.120.036 (1).pdf": {
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
                "currency": "DKK",
                "parsing_confidence": 0.98,
                "warnings": ["Total mileage and duration parsed from combined string"]
            }
        }
        
        # Look up the mock response based on filename
        extracted_data = mock_responses.get(filename)
        
        if extracted_data:
            return ParsedOffer(filename=filename, **extracted_data)
        
        # Fallback for unknown files or if real API call fails
        return ParsedOffer(filename=filename, warnings=["LLM parsing failed or is not configured."], parsing_confidence=0.1)

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
        currencies = [o.currency for o in self.offers if o.currency]
        if len(set(currencies)) > 1:
            errors.append(f"Mixed currencies detected: {set(currencies)}")
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

    if st.button("🎯 Load Demo Data", help="Load sample data for testing"):
        uploaded_files = create_demo_data()

    if uploaded_files:
        if len(uploaded_files) >= 2:
            template_buffer = create_default_template()
            process_offers(template_buffer, uploaded_files)
        else:
            st.warning("⚠️ Please upload at least 2 PDF files for comparison")

def create_demo_data():
    """Create dummy files for demonstration purposes."""
    st.info("Loading demo data...")
    # These mock files contain the text content from the PDFs the user provided
    demo_offers = [
        ("Kontraktoplæg_3052514001_1 (1).pdf", "Kontraktoplæg 3052514/001 ... Periode (mdr.): 48 ... Kilometer pr. år: 35.000 ... Leasinggiver: Ayvens ..."),
        ("quotation  2508.120.036 (1).pdf", "ARVAL ... quotation: 2508.120.03610/ ... contract annual kilometres/term (month): 35.000/48 ... price per month excl. VAT: 5.576,79 ...")
    ]
    uploaded_files = []
    for filename, content in demo_offers:
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp.write(content.encode('utf-8'))
            tmp_path = tmp.name
        
        uploaded_file = st.runtime.uploaded_file_manager.UploadedFile(
            name=filename,
            type="application/pdf",
            path=tmp_path,
            size=len(content.encode('utf-8'))
        )
        uploaded_files.append(uploaded_file)
        
    st.success("Demo data loaded! Please click the 'Compare Offers' button to proceed.")
    return uploaded_files

def create_default_template() -> io.BytesIO:
    """Create a default Excel template file for demonstration."""
    template_data = {
        'Field': [
            'Quote number', 'Driver name', 'Vehicle Description', 'Manufacturer', 'Model',
            'Version', 'JATO code', 'Fuel type', 'No. doors', 'Number of gears', 'HP',
            'C02 emission WLTP (g/km)', 'Battery range', 'Investment',
            'Vehicle list price (excl. VAT, excl. options)', 'Options (excl. taxes)',
            'Accessories (excl. taxes)', 'Delivery fee', 'Registration tax',
            'Total net investment', 'Taxation', 'Taxation value', 'Duration & Mileage',
            'Term (months)', 'Mileage per year (in km)', 'Financial rate',
            'Monthly financial rate (depreciation + interest)', 'Other fixed cost',
            'Maintenance, repairs and tires', 'Insurance', 'Administration fee',
            'Fixed costs', 'Leasing payment', 'Excess costs', 'Total cost', 'Winner'
        ],
        'Value': [None] * 36
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
    mapping_suggestions['Manufacturer'] = 'vehicle_description'
    mapping_suggestions['Model'] = 'vehicle_description'
    mapping_suggestions['Version'] = 'vehicle_description'
    mapping_suggestions['Fuel type'] = 'vehicle_description'
    mapping_suggestions['Term (months)'] = 'duration_months'
    mapping_suggestions['Mileage per year (in km)'] = 'total_mileage'
    mapping_suggestions['Total TCO'] = 'total_contract_cost'
    mapping_suggestions['Monthly TCO'] = 'monthly_rental'
    
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
            file_name = "Grundfos_Lars Østergaard"
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

    # Create a new DataFrame from the parsed offers for easy manipulation
    offer_data_for_df = []
    for offer in offers:
        offer_dict = asdict(offer)
        offer_dict['total_contract_cost'] = (offer.monthly_rental * offer.duration_months) + (offer.upfront_costs or 0) + (offer.deposit or 0) + (offer.admin_fees or 0)
        offer_data_for_pd.concat(offer_dict)
    
    offers_df = pd.DataFrame(offer_data_for_df)
    
    # Create the final DataFrame for the report
    report_df = template_df.copy()
    
    # Add columns for each vendor and populate them
    for _, offer in offers_df.iterrows():
        vendor_name = offer['vendor'] or "Unknown Vendor"
        
        # Add new column and get the column's index
        report_df.insert(len(report_df.columns), vendor_name, "")
        vendor_col_index = len(report_df.columns) - 1
        
        # Populate data based on the template's structure using the user mapping
        for index, row in report_df.iterrows():
            template_field = row.iloc[0] # Assumes the first column has the labels
            llm_field_name = user_mapping.get(template_field)
            if llm_field_name:
                try:
                    # Special handling for composite fields like 'Manufacturer'
                    if template_field == 'Manufacturer':
                        val = str(offer.get(llm_field_name, "")).split()[0]
                    elif template_field == 'Model':
                        val = " ".join(str(offer.get(llm_field_name, "")).split()[1:])
                    elif template_field == 'Fuel type':
                        val = 'EV' if 'el' in str(offer.get(llm_field_name, "")).lower() else None
                    elif template_field == 'Mileage per year (in km)':
                        val = offer.get(llm_field_name, 0) / (offer.get('duration_months', 12) / 12) if offer.get(llm_field_name) else None
                    else:
                        val = offer.get(llm_field_name)
                    
                    report_df.iloc[index, vendor_col_index] = val
                except (ValueError, TypeError):
                    report_df.iloc[index, vendor_col_index] = "N/A"

    # Vehicle Description Correspondence calculation
    if not offers_df.empty and len(offers_df) > 1:
        base_desc = offers_df.loc[0, 'vehicle_description'] or ""
        
        # Add the 'Correspondence (%)' row dynamically
        new_row_idx = report_df.index.max() + 1
        report_df.loc[new_row_idx, 'Field'] = 'Vehicle description correspondence'
        report_df.loc[new_row_idx, 'Value'] = '100.0%'
        
        for idx, offer in offers_df.iterrows():
            if idx > 0:
                desc_to_compare = offer.get('vehicle_description', "")
                similarity = calculate_similarity_score(base_desc, desc_to_compare)
                # Find the correct column for the vendor
                vendor_col = offer.get('vendor', "Unknown Vendor")
                if vendor_col in report_df.columns:
                    report_df.loc[new_row_idx, vendor_col] = f"{similarity:.1f}%"

    # Add Cost Analysis Summary at the bottom
    cost_data = OfferComparator(offers, {}).calculate_total_costs()
    sorted_offers = pd.DataFrame(cost_data).sort_values('total_contract_cost')
    
    report_df = report_pd.concat(pd.Series(), ignore_index=True)
    report_df = report_pd.concat(pd.Series(['Cost Analysis', None, None, None, None, None], index=report_df.columns), ignore_index=True)
    report_df = report_pd.concat(pd.Series(['Vendor', 'Total Cost', 'Monthly Cost', 'Winner'], index=report_df.columns), ignore_index=True)

    for index, row in sorted_offers.iterrows():
        is_winner = "🥇 Winner" if index == 0 else ""
        report_df = report_pd.concat(
            pd.Series([row['vendor'], f"{row['total_contract_cost']:,.2f}", f"{row['cost_per_month']:,.2f}", is_winner], index=report_df.columns),
            ignore_index=True
        )

    # Use a BytesIO buffer to save the Excel file in memory
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        report_df.to_excel(writer, sheet_name='Quotation', index=False, header=False)
    
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
