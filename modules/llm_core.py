import json
import logging
import traceback
import google.generativeai as genai
from typing import Dict

# Import the data model to structure the output
from .models import ParsedOffer

# Get a logger for this module
logger = logging.getLogger(__name__)

class LLMParser:
    """
    A class to handle all interactions with the Google Gemini Large Language Model (LLM).
    It is responsible for sending prompts, defining the expected output structure,
    and parsing the AI's response.
    """

    def __init__(self, api_key: str):
        """
        Initializes the Gemini client with the provided API key.

        Args:
            api_key: The Google AI API key for authentication.
        
        Raises:
            ValueError: If the API key is not provided.
        """
        if not api_key:
            raise ValueError("A Google AI API key is required to use the LLM parser.")
        self.api_key = api_key
        genai.configure(api_key=self.api_key)
        logger.info("Google Gemini client configured successfully.")

    def parse_text(self, text: str, filename: str) -> ParsedOffer:
        """
        Sends the extracted text from an offer to the Gemini API and parses the
        structured JSON response into a ParsedOffer object.

        Args:
            text: The raw text extracted from the PDF or spreadsheet.
            filename: The name of the file being processed, for context.

        Returns:
            A ParsedOffer object populated with the data extracted by the AI.
            If parsing fails, it returns an object with low confidence and a warning.
        """
        logger.info(f"Sending text from '{filename}' to Gemini for parsing.")

        # This detailed prompt instructs the AI on exactly how to behave and what to extract.
        prompt_text = f"""
        You are a world-class financial analyst specializing in fleet leasing. Your task is to extract key data points from a vehicle leasing contract, regardless of the language or format.

        **CRITICAL INSTRUCTIONS:**
        1.  **Distinguish Terms:** Differentiate between the maximum allowed contract terms and the actual terms of this specific offer.
        2.  **VAT Exclusion:** All extracted prices and costs (`monthly_rental`, `total_monthly_lease`, etc.) MUST be excluding VAT (Value-Added Tax).
        3.  **Identify Parties:** `driver_name` is the employee; `customer` is the company.
        4.  **Calculate Mileage:** If annual mileage is given (e.g., "35,000 km per year / 48 months"), you MUST calculate the `offer_total_mileage` (35000 * (48 / 12) = 140000).
        5.  **Fuel Type:** Treat "BEV", "Electric", and "electricity" as the same fuel type.

        Return the data as a JSON object strictly following the provided schema. If a value is not found or is not applicable, use `null`. Do not invent values.
        
        <DOCUMENT_TO_PARSE>
        {text}
        </DOCUMENT_TO_PARSE>
        """

        # This schema perfectly mirrors the ParsedOffer dataclass in models.py.
        # It forces the LLM to return a predictable, structured JSON object.
        json_schema = {
            "type": "OBJECT",
            "properties": {
                "filename": {"type": "STRING"},
                "vendor": {"type": "STRING"},
                "vehicle_description": {"type": "STRING"},
                "max_duration_months": {"type": "NUMBER"},
                "max_total_mileage": {"type": "NUMBER"},
                "offer_duration_months": {"type": "NUMBER"},
                "offer_total_mileage": {"type": "NUMBER"},
                "monthly_rental": {"type": "NUMBER"},
                "total_monthly_lease": {"type": "NUMBER"},
                "currency": {"type": "STRING"},
                "upfront_costs": {"type": "NUMBER"},
                "deposit": {"type": "NUMBER"},
                "admin_fees": {"type": "NUMBER"},
                "excess_mileage_rate": {"type": "NUMBER"},
                "unused_mileage_rate": {"type": "NUMBER"},
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
                "roadside_assistance": {"type": "NUMBER"},
                "tyres_cost": {"type": "NUMBER"},
                "maintenance_included": {"type": "BOOLEAN"},
                "driver_name": {"type": "STRING"},
                "customer": {"type": "STRING"},
                "options_list": {"type": "ARRAY", "items": {"type": "OBJECT", "properties": {"name": {"type": "STRING"}, "price": {"type": "NUMBER"}}}},
                "accessories_list": {"type": "ARRAY", "items": {"type": "OBJECT", "properties": {"name": {"type": "STRING"}, "price": {"type": "NUMBER"}}}},
                "parsing_confidence": {"type": "NUMBER", "description": "A value between 0.0 and 1.0 indicating your confidence in the extracted data."},
                "warnings": {"type": "ARRAY", "items": {"type": "STRING"}},
            }
        }

        # Configure the model to use the JSON schema for its response
        generation_config = genai.types.GenerationConfig(
            response_mime_type="application/json",
            response_schema=json_schema
        )
        
        model = genai.GenerativeModel(
            model_name='gemini-1.5-pro-latest',
            generation_config=generation_config
        )

        try:
            # Make the API call to the Gemini model
            response = model.generate_content(prompt_text)
            extracted_data = json.loads(response.text)
            
            # Ensure list fields exist to prevent errors later
            extracted_data['options_list'] = extracted_data.get('options_list') or []
            extracted_data['accessories_list'] = extracted_data.get('accessories_list') or []
            extracted_data['filename'] = filename # Ensure filename is set

            # Create and return the structured ParsedOffer object
            return ParsedOffer(**extracted_data)

        except Exception as e:
            logger.error(f"Error during Gemini API call for {filename}: {str(e)}")
            logger.error(traceback.format_exc())
            # Return a default object with an error message
            return ParsedOffer(
                filename=filename,
                warnings=[f"LLM parsing failed: {str(e)}"],
                parsing_confidence=0.1
            )


