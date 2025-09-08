"""
Fleet Leasing Offer Comparator - Streamlit App (Improved Version)
Author: Fleet Management Tool
Requirements:
  streamlit, pandas, numpy, pdfplumber, python-dateutil, xlsxwriter
Optional:
  camelot-py[cv], tabula-py, pdfminer.six, pytesseract
Notes:
  - All prices considered ex-VAT and exclude fuel by design
  - Enhanced error handling and validation
  - Improved parsing accuracy with fallback mechanisms
  - Better UI/UX with progress indicators and clear feedback
"""

import io
import re
import sys
import logging
import tempfile
from typing import List, Dict, Any, Optional, Tuple, Union
from dataclasses import dataclass, field
from datetime import datetime, date
import traceback

import streamlit as st
import pandas as pd
import numpy as np
import pdfplumber
from dateutil import parser as dateparser

# Optional imports with graceful fallbacks
try:
    import camelot
    HAS_CAMELOT = True
except ImportError:
    HAS_CAMELOT = False
    st.warning("camelot-py not installed. Table extraction will use fallback methods only.")

try:
    import tabula
    HAS_TABULA = True
except ImportError:
    HAS_TABULA = False

# Configure logging
@st.cache_resource
def setup_logging():
    logger = logging.getLogger("leasing_comparator")
    logger.setLevel(logging.INFO)
    
    if not logger.handlers:
        handler = logging.StreamHandler()
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    
    return logger

logger = setup_logging()

# Global constants
CURRENCY_SYMBOLS = {
    "€": "EUR", "EUR": "EUR", "£": "GBP", "GBP": "GBP",
    "$": "USD", "USD": "USD", "CHF": "CHF"
}

NUMBER_PATTERN = re.compile(r"[-+]?\d{1,3}(?:[.,\s]\d{3})*(?:[.,]\d+)?|\d+(?:[.,]\d+)?")

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
    delivery_fees: Optional[float] = None
    admin_fees: Optional[float] = None
    maintenance_included: Optional[bool] = None
    maintenance_cost: Optional[float] = None
    tyres_included: Optional[bool] = None
    tyres_cost: Optional[float] = None
    insurance_included: Optional[bool] = None
    insurance_cost: Optional[float] = None
    road_tax_included: Optional[bool] = None
    road_tax_cost: Optional[float] = None
    excess_mileage_rate: Optional[float] = None
    discount_amount: Optional[float] = None
    currency: Optional[str] = None
    offer_valid_until: Optional[str] = None
    delivery_time: Optional[str] = None
    raw_text: str = ""
    parsing_confidence: float = 0.0
    warnings: List[str] = field(default_factory=list)
    is_scanned: bool = False

class TextProcessor:
    """Handles text extraction and normalization"""
    
    @staticmethod
    def extract_text_from_pdf(pdf_bytes: bytes) -> Tuple[str, bool]:
        """Extract text from PDF, return (text, is_scanned)"""
        try:
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                pages_text = []
                for page in pdf.pages:
                    try:
                        text = page.extract_text() or ""
                        pages_text.append(text)
                    except Exception as e:
                        logger.warning(f"Failed to extract text from page: {e}")
                        pages_text.append("")
                
                full_text = "\n".join(pages_text)
                is_scanned = len(full_text.strip()) < 50  # Heuristic for scanned PDFs
                
                return full_text, is_scanned
                
        except Exception as e:
            logger.error(f"PDF text extraction failed: {e}")
            return "", True
    
    @staticmethod
    def normalize_number(text: str) -> Optional[float]:
        """Convert various number formats to float"""
        if not text:
            return None
            
        # Clean the text
        clean_text = re.sub(r'[^\d,.\-+]', '', str(text).strip())
        if not clean_text:
            return None
        
        try:
            # Handle European format (1.234,56) vs American (1,234.56)
            if ',' in clean_text and '.' in clean_text:
                # Determine which is decimal separator by position
                last_comma = clean_text.rfind(',')
                last_dot = clean_text.rfind('.')
                
                if last_comma > last_dot:
                    # European format: comma is decimal
                    clean_text = clean_text.replace('.', '').replace(',', '.')
                else:
                    # American format: dot is decimal
                    clean_text = clean_text.replace(',', '')
            elif ',' in clean_text:
                # Only comma: check if it's decimal or thousands separator
                parts = clean_text.split(',')
                if len(parts) == 2 and len(parts[1]) <= 2:
                    # Likely decimal separator
                    clean_text = clean_text.replace(',', '.')
                else:
                    # Thousands separator
                    clean_text = clean_text.replace(',', '')
            
            return float(clean_text)
            
        except (ValueError, AttributeError):
            logger.warning(f"Failed to normalize number: {text}")
            return None
    
    @staticmethod
    def detect_currency(text: str) -> Optional[str]:
        """Detect currency from text"""
        if not text:
            return None
            
        text_upper = text.upper()
        for symbol, code in CURRENCY_SYMBOLS.items():
            if symbol in text or code in text_upper:
                return code
        return None

class OfferParser:
    """Main parser for PDF leasing offers"""
    
    def __init__(self):
        self.text_processor = TextProcessor()
    
    def parse_pdf(self, pdf_bytes: bytes, filename: str) -> ParsedOffer:
        """Parse a PDF leasing offer"""
        offer = ParsedOffer(filename=filename)
        
        try:
            # Extract text
            offer.raw_text, offer.is_scanned = self.text_processor.extract_text_from_pdf(pdf_bytes)
            
            if offer.is_scanned:
                offer.warnings.append("Document appears to be scanned - OCR may be needed for better accuracy")
                offer.parsing_confidence = 0.1
                return offer
            
            # Parse individual fields
            self._parse_vendor(offer)
            self._parse_vehicle_description(offer)
            self._parse_financial_details(offer)
            self._parse_contract_terms(offer)
            self._parse_additional_costs(offer)
            self._parse_dates_and_delivery(offer)
            
            # Calculate overall confidence
            offer.parsing_confidence = self._calculate_confidence(offer)
            
        except Exception as e:
            logger.error(f"Error parsing {filename}: {e}")
            offer.warnings.append(f"Parsing error: {str(e)}")
            offer.parsing_confidence = 0.0
        
        return offer
    
    def _parse_vendor(self, offer: ParsedOffer):
        """Extract vendor/leasing company name"""
        text = offer.raw_text
        
        # Look for common patterns
        patterns = [
            r'([A-Z][A-Za-z\s&]+(?:leasing|lease|finance|motor|automotive|rentals|fleet))',
            r'from[\s:]+([A-Z][A-Za-z\s&]+)',
            r'([A-Z][A-Za-z\s&]+)\s+(?:offers|presents|quotes)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                offer.vendor = match.group(1).strip()[:50]  # Limit length
                break
        
        # Fallback: use filename
        if not offer.vendor:
            offer.vendor = re.sub(r'[._-]', ' ', offer.filename.replace('.pdf', '')).strip()
    
    def _parse_vehicle_description(self, offer: ParsedOffer):
        """Extract vehicle description"""
        text = offer.raw_text
        
        patterns = [
            r'(?:vehicle|model|car)[\s:]+([^\n\r]{5,100})',
            r'(?:make/model)[\s:]+([^\n\r]{5,100})',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                offer.vehicle_description = match.group(1).strip()[:200]
                break
    
    def _parse_financial_details(self, offer: ParsedOffer):
        """Parse financial information"""
        text = offer.raw_text
        
        # Currency detection
        offer.currency = self.text_processor.detect_currency(text)
        
        # Monthly rental
        monthly_patterns = [
            r'monthly[\s\w]*?rental[\s:]*([€£$]?\s*[\d,.\s]+)',
            r'per\s+month[\s:]*([€£$]?\s*[\d,.\s]+)',
            r'monthly[\s:]*([€£$]?\s*[\d,.\s]+)',
        ]
        
        for pattern in monthly_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                offer.monthly_rental = self.text_processor.normalize_number(match.group(1))
                if offer.monthly_rental:
                    break
        
        # Deposit
        deposit_patterns = [
            r'deposit[\s:]*([€£$]?\s*[\d,.\s]+)',
            r'down\s+payment[\s:]*([€£$]?\s*[\d,.\s]+)',
        ]
        
        for pattern in deposit_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                offer.deposit = self.text_processor.normalize_number(match.group(1))
                if offer.deposit:
                    break
        
        # Admin fees
        admin_pattern = r'admin(?:istration)?\s+fee[\s:]*([€£$]?\s*[\d,.\s]+)'
        match = re.search(admin_pattern, text, re.IGNORECASE)
        if match:
            offer.admin_fees = self.text_processor.normalize_number(match.group(1))
    
    def _parse_contract_terms(self, offer: ParsedOffer):
        """Parse contract duration and mileage"""
        text = offer.raw_text
        
        # Contract duration
        duration_patterns = [
            r'(\d+)\s*(?:months?|mths?)',
            r'(\d+(?:\.\d+)?)\s*(?:years?|yrs?)(?:\s*[=*]\s*(\d+)\s*months?)?',
        ]
        
        for pattern in duration_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                if 'year' in pattern:
                    years = float(match.group(1))
                    offer.duration_months = int(years * 12)
                else:
                    offer.duration_months = int(match.group(1))
                break
        
        # Total mileage
        mileage_patterns = [
            r'(\d{1,3}(?:[,.\s]\d{3})*)\s*(?:km|miles?)',
            r'(\d+(?:,\d{3})*)\s*(?:km|miles?)\s*(?:total|per\s+contract)',
        ]
        
        for pattern in mileage_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                offer.total_mileage = int(self.text_processor.normalize_number(match.group(1)) or 0)
                if offer.total_mileage:
                    break
    
    def _parse_additional_costs(self, offer: ParsedOffer):
        """Parse maintenance, insurance, etc."""
        text = offer.raw_text.lower()
        
        # Maintenance
        if re.search(r'maintenance\s+included', text):
            offer.maintenance_included = True
        
        maintenance_match = re.search(r'maintenance[\s:]*([€£$]?\s*[\d,.\s]+)', text)
        if maintenance_match:
            offer.maintenance_cost = self.text_processor.normalize_number(maintenance_match.group(1))
        
        # Similar patterns for tyres, insurance, road tax
        for service in ['tyres', 'insurance', 'road tax']:
            if re.search(rf'{service}\s+included', text):
                setattr(offer, f'{service.replace(" ", "_")}_included', True)
            
            cost_match = re.search(rf'{service}[\s:]*([€£$]?\s*[\d,.\s]+)', text)
            if cost_match:
                setattr(offer, f'{service.replace(" ", "_")}_cost', 
                       self.text_processor.normalize_number(cost_match.group(1)))
    
    def _parse_dates_and_delivery(self, offer: ParsedOffer):
        """Parse validity dates and delivery times"""
        text = offer.raw_text
        
        # Offer validity
        validity_match = re.search(r'valid\s+until[\s:]*([^\n\r]{5,50})', text, re.IGNORECASE)
        if validity_match:
            try:
                parsed_date = dateparser.parse(validity_match.group(1))
                if parsed_date:
                    offer.offer_valid_until = parsed_date.date().isoformat()
            except:
                offer.offer_valid_until = validity_match.group(1).strip()
        
        # Delivery time
        delivery_match = re.search(r'(?:delivery|lead\s+time)[\s:]*(\d+\s*(?:weeks?|months?|days?))', text, re.IGNORECASE)
        if delivery_match:
            offer.delivery_time = delivery_match.group(1)
    
    def _calculate_confidence(self, offer: ParsedOffer) -> float:
        """Calculate parsing confidence score"""
        score = 0.0
        
        # Essential fields
        if offer.monthly_rental:
            score += 0.3
        if offer.duration_months:
            score += 0.2
        if offer.total_mileage:
            score += 0.2
        if offer.vendor:
            score += 0.1
        if offer.currency:
            score += 0.1
        
        # Additional fields
        if offer.deposit is not None:
            score += 0.05
        if offer.admin_fees is not None:
            score += 0.05
        
        return min(1.0, score)

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
        
        # Check for essential data
        durations = [o.duration_months for o in self.offers if o.duration_months]
        mileages = [o.total_mileage for o in self.offers if o.total_mileage]
        
        if len(durations) != len(self.offers):
            errors.append("Some offers missing contract duration")
        elif len(set(durations)) > 1:
            errors.append(f"Contract durations don't match: {set(durations)}")
        
        if len(mileages) != len(self.offers):
            errors.append("Some offers missing mileage information")
        elif len(set(mileages)) > 1:
            errors.append(f"Contract mileages don't match: {set(mileages)}")
        
        # Currency check
        currencies = [o.currency for o in self.offers if o.currency]
        if len(set(currencies)) > 1:
            errors.append(f"Mixed currencies detected: {set(currencies)}")
        
        return len(errors) == 0, errors
    
    def calculate_total_costs(self) -> List[Dict[str, Any]]:
        """Calculate total contract costs for all offers"""
        results = []
        
        for offer in self.offers:
            if not offer.duration_months or not offer.monthly_rental:
                results.append({
                    'vendor': offer.vendor,
                    'error': 'Missing essential data for cost calculation'
                })
                continue
            
            # Base costs
            monthly_total = offer.monthly_rental * offer.duration_months
            upfront_total = (offer.deposit or 0) + (offer.delivery_fees or 0) + (offer.admin_fees or 0)
            
            # Optional costs based on configuration
            additional_costs = 0
            if self.config.get('include_maintenance') and offer.maintenance_cost:
                additional_costs += offer.maintenance_cost
            if self.config.get('include_tyres') and offer.tyres_cost:
                additional_costs += offer.tyres_cost
            if self.config.get('include_insurance') and offer.insurance_cost:
                additional_costs += offer.insurance_cost
            if self.config.get('include_road_tax') and offer.road_tax_cost:
                additional_costs += offer.road_tax_cost
            
            # Discounts
            discount = offer.discount_amount or 0
            
            # Total contract cost
            total_cost = monthly_total + upfront_total + additional_costs - discount
            
            results.append({
                'vendor': offer.vendor,
                'vehicle': offer.vehicle_description,
                'duration_months': offer.duration_months,
                'total_mileage': offer.total_mileage,
                'monthly_rental': offer.monthly_rental,
                'monthly_total': monthly_total,
                'upfront_total': upfront_total,
                'additional_costs': additional_costs,
                'discount': discount,
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
            df['rank'] = range(1, len(df) + 1)
        
        return df

# Streamlit UI
def main():
    st.set_page_config(
        page_title="Fleet Leasing Offer Comparator",
        page_icon="🚗",
        layout="wide"
    )
    
    st.title("🚗 Fleet Leasing Offer Comparator")
    
    st.markdown("""
    Compare multiple leasing offers for the same vehicle configuration. 
    Upload PDF offers with identical contract duration and mileage for accurate comparison.
    
    **Key Features:**
    - Automatic PDF parsing and data extraction
    - Standardized cost comparison (ex-VAT, excluding fuel)
    - Excel export with detailed breakdown
    - Configurable cost inclusions (maintenance, tyres, etc.)
    """)
    
    # Sidebar configuration
    st.sidebar.header("⚙️ Configuration")
    
    config = {
        'include_maintenance': st.sidebar.checkbox("Include Maintenance Costs", value=True),
        'include_tyres': st.sidebar.checkbox("Include Tyre Costs", value=False),
        'include_insurance': st.sidebar.checkbox("Include Insurance", value=False),
        'include_road_tax': st.sidebar.checkbox("Include Road Tax", value=False),
    }
    
    # File upload
    st.header("📁 Upload Offers")
    uploaded_files = st.file_uploader(
        "Upload PDF leasing offers (2-10 files)",
        type=['pdf'],
        accept_multiple_files=True,
        help="Upload PDF files containing leasing offers for the same vehicle"
    )
    
    # Demo data option
    if st.checkbox("🎯 Load Demo Data", help="Load sample data for testing"):
        uploaded_files = create_demo_data()
    
    if uploaded_files and len(uploaded_files) >= 2:
        process_offers(uploaded_files, config)
    elif uploaded_files:
        st.warning("⚠️ Please upload at least 2 PDF files for comparison")

def create_demo_data():
    """Create demo data for testing"""
    # This would create sample ParsedOffer objects
    # Implementation would depend on your specific demo needs
    return []

def process_offers(uploaded_files, config):
    """Process uploaded offers and generate comparison"""
    parser = OfferParser()
    offers = []
    
    # Progress tracking
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    # Parse each file
    for i, uploaded_file in enumerate(uploaded_files):
        status_text.text(f"Processing {uploaded_file.name}...")
        
        try:
            pdf_bytes = uploaded_file.read()
            offer = parser.parse_pdf(pdf_bytes, uploaded_file.name)
            offers.append(offer)
            
        except Exception as e:
            st.error(f"❌ Error processing {uploaded_file.name}: {str(e)}")
            logger.error(f"File processing error: {e}\n{traceback.format_exc()}")
        
        progress_bar.progress((i + 1) / len(uploaded_files))
    
    status_text.text("Processing complete!")
    progress_bar.empty()
    
    if not offers:
        st.error("❌ No offers could be processed successfully")
        return
    
    # Display parsing results
    display_parsing_results(offers)
    
    # Validate and compare
    comparator = OfferComparator(offers, config)
    is_valid, errors = comparator.validate_offers()
    
    if not is_valid:
        st.error("❌ Validation Errors:")
        for error in errors:
            st.error(f"• {error}")
        return
    
    # Generate comparison
    comparison_df = comparator.generate_comparison_report()
    display_comparison_results(comparison_df, config)
    
    # Export functionality
    provide_export_options(comparison_df, offers, config)

def display_parsing_results(offers: List[ParsedOffer]):
    """Display parsing results summary"""
    st.header("📊 Parsing Results")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        avg_confidence = np.mean([o.parsing_confidence for o in offers])
        st.metric("Average Confidence", f"{avg_confidence:.1%}")
    
    with col2:
        scanned_count = sum(1 for o in offers if o.is_scanned)
        st.metric("Scanned PDFs", scanned_count)
    
    with col3:
        warning_count = sum(len(o.warnings) for o in offers)
        st.metric("Total Warnings", warning_count)
    
    # Detailed results
    with st.expander("📋 Detailed Parsing Results"):
        for offer in offers:
            st.write(f"**{offer.vendor or offer.filename}**")
            st.write(f"Confidence: {offer.parsing_confidence:.1%}")
            if offer.warnings:
                st.write("⚠️ Warnings:", ", ".join(offer.warnings))
            st.write("---")

def display_comparison_results(df: pd.DataFrame, config: Dict[str, Any]):
    """Display comparison results"""
    st.header("🏆 Comparison Results")
    
    if df.empty:
        st.error("No valid offers to compare")
        return
    
    # Winner announcement
    winner = df.iloc[0]
    st.success(f"🥇 **Winner: {winner['vendor']}** - Total Cost: {winner['total_contract_cost']:,.2f} {winner.get('currency', '')}")
    
    # Comparison table
    st.subheader("📈 Detailed Comparison")
    
    display_columns = [
        'rank', 'vendor', 'monthly_rental', 'upfront_total', 
        'additional_costs', 'total_contract_cost', 'cost_per_km', 'parsing_confidence'
    ]
    
    # Format the dataframe for display
    display_df = df[display_columns].copy()
    
    st.dataframe(
        display_df.style.format({
            'monthly_rental': '{:,.2f}',
            'upfront_total': '{:,.2f}',
            'additional_costs': '{:,.2f}',
            'total_contract_cost': '{:,.2f}',
            'cost_per_km': '{:.4f}',
            'parsing_confidence': '{:.1%}'
        }).highlight_min(['total_contract_cost'], color='lightgreen'),
        use_container_width=True
    )
    
    # Cost breakdown chart
    st.subheader("💰 Cost Breakdown")
    
    chart_data = df[['vendor', 'monthly_total', 'upfront_total', 'additional_costs']].set_index('vendor')
    st.bar_chart(chart_data)

def provide_export_options(df: pd.DataFrame, offers: List[ParsedOffer], config: Dict[str, Any]):
    """Provide export options for the results"""
    st.header("📤 Export Results")
    
    if st.button("🔽 Generate Excel Report"):
        try:
            excel_buffer = generate_excel_report(df, offers, config)
            
            st.download_button(
                label="📊 Download Excel Report",
                data=excel_buffer.getvalue(),
                file_name=f"leasing_comparison_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.success("✅ Excel report generated successfully!")
            
        except Exception as e:
            st.error(f"❌ Error generating Excel report: {str(e)}")
            logger.error(f"Excel generation error: {e}")

def generate_excel_report(df: pd.DataFrame, offers: List[ParsedOffer], config: Dict[str, Any]) -> io.BytesIO:
    """Generate Excel report with multiple sheets"""
    buffer = io.BytesIO()
    
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        # Winner Analysis sheet
        df.to_excel(writer, sheet_name='Winner Analysis', index=False)
        
        # Individual offer sheets
        for offer in offers:
            sheet_name = (offer.vendor or offer.filename)[:31]  # Excel sheet name limit
            
            offer_data = {
                'Field': [
                    'Vendor', 'Vehicle', 'Duration (months)', 'Total Mileage',
                    'Monthly Rental', 'Deposit', 'Admin Fees', 'Maintenance Cost',
                    'Currency', 'Parsing Confidence', 'Warnings'
                ],
                'Value': [
                    offer.vendor, offer.vehicle_description, offer.duration_months,
                    offer.total_mileage, offer.monthly_rental, offer.deposit,
                    offer.admin_fees, offer.maintenance_cost, offer.currency,
                    f"{offer.parsing_confidence:.1%}", "; ".join(offer.warnings)
                ]
            }
            
            offer_df = pd.DataFrame(offer_data)
            offer_df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    buffer.seek(0)
    return buffer

if __name__ == "__main__":
    main()
