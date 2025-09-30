import streamlit as st
import json
import logging
from typing import Dict, List, IO

# --- Import Core Logic from Modules ---
from modules.llm_core import LLMParser  # CORRECTED: Was LLMManager
from modules.models import ParsedOffer
from modules.offer_comparator import OfferComparator, get_offer_diff, calculate_similarity_score
from modules.pdf_parser import extract_text_from_pdf
from modules.report_generator import generate_excel_report

# --- Configuration & Setup ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)
RECIPES_FILE = 'config/recipes.json'

# --- Helper Functions ---
@st.cache_data
def load_recipes() -> Dict:
    """Loads the recipes from the JSON file. Cached for performance."""
    try:
        with open(RECIPES_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        st.error(f"Fatal Error: `{RECIPES_FILE}` not found. Please ensure it is in the 'config' directory.")
        return {}
    except json.JSONDecodeError:
        st.error(f"Fatal Error: Could not parse `{RECIPES_FILE}`. Please check for syntax errors.")
        return {}

@st.cache_data
def process_offers_internal(_parser: LLMParser, uploaded_files: List[IO[bytes]]) -> List[ParsedOffer]:
    """
    Internal function to process uploaded files with caching.
    Extracts text and sends it to the LLM parser.
    """
    offers = []
    progress_bar = st.progress(0, "Initializing AI processing...")
    total_files = len(uploaded_files)

    for i, uploaded_file in enumerate(uploaded_files):
        filename = uploaded_file.name
        progress_text = f"Processing {filename} with AI... ({i+1}/{total_files})"
        progress_bar.progress((i + 1) / total_files, text=progress_text)
        
        try:
            raw_text = extract_text_from_pdf(uploaded_file)
            if raw_text and raw_text.strip():
                offer = _parser.parse_text(raw_text, filename)
                offers.append(offer)
            else:
                st.warning(f"⚠️ Could not extract any text from '{filename}'. It might be an image-based PDF or corrupted.")
        except Exception as e:
            st.error(f"❌ Error processing {filename}: {str(e)}")
            logger.error(f"File processing error for {filename}: {e}", exc_info=True)

    progress_bar.empty()
    return offers

# --- Main Streamlit Application UI ---
def main():
    st.set_page_config(page_title="Leasing Quote Comparator", layout="wide")
    st.title("🤖 AI-Powered Leasing Quote Comparison Tool")
    st.markdown("Upload a reference offer and one or more competitor offers to get a detailed, side-by-side comparison.")

    recipes = load_recipes()
    if not recipes:
        st.stop()

    # --- Sidebar for Configuration ---
    with st.sidebar:
        st.header("⚙️ Configuration")
        try:
            # Recommended: Use Streamlit secrets for API keys
            api_key = st.secrets["GOOGLE_API_KEY"]
            st.success("✅ API Key loaded from secrets.")
        except (FileNotFoundError, KeyError):
            st.warning("API Key not found in secrets.")
            api_key = st.text_input("Enter your Google AI API Key", type="password")

        st.markdown("---")
        st.header("Recipe Selection")
        st.info("Select a recipe to guide the AI on how to interpret the offers.")

        countries = list(recipes.keys())
        selected_country = st.selectbox("Select Country", options=countries)
        
        selected_vendor = None
        if selected_country:
            vendors = list(recipes[selected_country].keys())
            selected_vendor = st.selectbox("Select Vendor", options=vendors)

    # --- Main Page for File Upload and Processing ---
    st.header("📁 Upload Offer Documents")
    ref_file = st.file_uploader("1. Upload Reference Offer (PDF)", type=['pdf'])
    comp_files = st.file_uploader("2. Upload Competitor Offers (PDF)", type=['pdf'], accept_multiple_files=True)

    if st.button("🚀 Compare Offers", type="primary"):
        # --- Initial Checks ---
        if not api_key:
            st.error("🚨 Please enter your Google AI API key in the sidebar.")
        elif not ref_file:
            st.warning("⚠️ Please upload a reference offer.")
        elif not comp_files:
            st.warning("⚠️ Please upload at least one competitor offer.")
        else:
            with st.spinner("AI is analyzing the documents... This may take a moment."):
                try:
                    parser = LLMParser(api_key=api_key) # CORRECTED: Was LLMManager
                    all_files = [ref_file] + comp_files
                    
                    # Process all offers and store them in the session state
                    st.session_state.offers = process_offers_internal(parser, all_files)
                    
                    if not st.session_state.offers:
                         st.error("Processing complete, but no data was successfully extracted.")
                    
                except ValueError as e:
                    st.error(f"Initialization Error: {e}")
                except Exception as e:
                    st.error(f"An unexpected error occurred during processing: {e}")
                    logger.error("Offer processing failed", exc_info=True)
    
    # --- Display Results if Offers are Processed ---
    if 'offers' in st.session_state and st.session_state.offers:
        st.success("🎉 Analysis complete!")
        offers = st.session_state.offers
        reference_offer = offers[0]
        competitor_offers = offers[1:]

        comparator = OfferComparator(offers)
        is_valid, validation_errors = comparator.validate_offers()

        tab1, tab2, tab3 = st.tabs(["📊 Cost Comparison", "🔍 Specification Gap Analysis", "📄 Raw Extracted Data"])

        with tab1:
            st.header("Cost and Financial Comparison")
            if not is_valid:
                st.warning("Offers may not be directly comparable due to inconsistencies.")
                for error in validation_errors:
                    st.error(f"• {error}")
            
            report_df = comparator.generate_comparison_report()
            if not report_df.empty:
                st.dataframe(report_df.style.format({
                    'total_contract_cost': '{:,.2f}',
                    'cost_per_month': '{:,.2f}',
                    'cost_per_km': '{:,.4f}'
                }), use_container_width=True)
                
                # Download button for the full side-by-side Excel report
                excel_bytes = generate_excel_report(offers)
                st.download_button(
                    label="⬇️ Download Full Comparison Report (Excel)",
                    data=excel_bytes,
                    file_name="Leasing_Comparison_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("Could not generate a cost comparison report.")

        with tab2:
            st.header("Vehicle Specification Gap Analysis")
            st.markdown(f"Comparing all offers against the reference: **{reference_offer.vehicle_description or 'N/A'}** from **{reference_offer.vendor or 'N/A'}**")
            
            for offer in competitor_offers:
                st.markdown("---")
                col1, col2 = st.columns([1, 2])
                with col1:
                    score = calculate_similarity_score(reference_offer.vehicle_description, offer.vehicle_description)
                    st.metric(label=f"Similarity vs. {offer.vendor or offer.filename}", value=f"{score:.1f}%")
                with col2:
                    st.markdown(f"**Key Differences vs. Reference:**")
                    diff_text = get_offer_diff(reference_offer, offer)
                    if "No significant differences" in diff_text:
                        st.success(f"✅ {diff_text}")
                    else:
                        st.text(diff_text)

        with tab3:
            st.header("Raw Extracted Data (JSON)")
            st.info("This is the structured data extracted by the AI from each document.")
            for offer in offers:
                with st.expander(f"📄 {offer.filename} ({offer.vendor or 'Unknown Vendor'})"):
                    st.json(offer.to_dict())

if __name__ == "__main__":
    main()

