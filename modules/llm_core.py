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
    It executes the detailed instructions provided by a recipe prompt.
    """

    def __init__(self, api_key: str):
        """Initializes the Gemini client with the provided API key."""
        if not api_key:
            raise ValueError("A Google AI API key is required to use the LLM parser.")
        self.api_key = api_key
        genai.configure(api_key=self.api_key)
        logger.info("Google Gemini client configured successfully.")

    def parse_text(self, text: str, filename: str, prompt_template: str) -> ParsedOffer:
        """
        Sends the extracted text and a specific recipe prompt to the Gemini API
        and parses the structured JSON response into a ParsedOffer object.
        """
        logger.info(f"Sending text from '{filename}' to Gemini for parsing using a customer-specific recipe.")
        
        # The final prompt is a combination of the recipe and the document text
        prompt_text = f"{prompt_template}\n\n<DOCUMENT_TO_PARSE>\n{text}\n</DOCUMENT_TO_PARSE>"

        # This JSON schema enforces a consistent output structure from the AI
        json_schema = {
            "type": "OBJECT",
            "properties": {
                "filename": {"type": "STRING"}, "vendor": {"type": "STRING"}, "vehicle_description": {"type": "STRING"},
                "max_duration_months": {"type": "NUMBER"}, "max_total_mileage": {"type": "NUMBER"}, "offer_duration_months": {"type": "NUMBER"},
                "offer_total_mileage": {"type": "NUMBER"}, "monthly_rental": {"type": "NUMBER"}, "total_monthly_lease": {"type": "NUMBER"},
                "currency": {"type": "STRING"}, "upfront_costs": {"type": "NUMBER"}, "deposit": {"type": "NUMBER"},
                "admin_fees": {"type": "NUMBER"}, "excess_mileage_rate": {"type": "NUMBER"}, "unused_mileage_rate": {"type": "NUMBER"},
                "quote_number": {"type": "STRING"}, "manufacturer": {"type": "STRING"}, "model": {"type": "STRING"},
                "version": {"type": "STRING"}, "internal_colour": {"type": "STRING"}, "external_colour": {"type": "STRING"},
                "fuel_type": {"type": "STRING"}, "num_doors": {"type": "NUMBER"}, "hp": {"type": "NUMBER"}, "c02_emission": {"type": "NUMBER"},
                "battery_range": {"type": "NUMBER"}, "vehicle_price": {"type": "NUMBER"}, "options_price": {"type": "NUMBER"},
                "accessories_price": {"type": "NUMBER"}, "delivery_cost": {"type": "NUMBER"}, "registration_tax": {"type": "NUMBER"},
                "total_net_investment": {"type": "NUMBER"}, "taxation_value": {"type": "NUMBER"}, "financial_rate": {"type": "NUMBER"},
                "depreciation_interest": {"type": "NUMBER"}, "maintenance_repair": {"type": "NUMBER"}, "insurance_cost": {"type": "NUMBER"},
                "green_tax": {"type": "NUMBER"}, "management_fee": {"type": "NUMBER"}, "roadside_assistance": {"type": "NUMBER"},
                "tyres_cost": {"type": "NUMBER"}, "maintenance_included": {"type": "BOOLEAN"}, "driver_name": {"type": "STRING"},
                "customer": {"type": "STRING"},
                "options_list": {"type": "ARRAY", "items": {"type": "OBJECT", "properties": {"name": {"type": "STRING"}, "price": {"type": "NUMBER"}}}},
                "accessories_list": {"type": "ARRAY", "items": {"type": "OBJECT", "properties": {"name": {"type": "STRING"}, "price": {"type": "NUMBER"}}}},
                "parsing_confidence": {"type": "NUMBER"}, "warnings": {"type": "ARRAY", "items": {"type": "STRING"}},
            }
        }

        generation_config = genai.types.GenerationConfig(response_mime_type="application/json", response_schema=json_schema)
        model = genai.GenerativeModel('gemini-1.5-pro-latest', generation_config=generation_config)

        try:
            response = model.generate_content(prompt_text)
            extracted_data = json.loads(response.text)
            
            # Ensure essential fields are set correctly
            extracted_data.update({
                'options_list': extracted_data.get('options_list') or [],
                'accessories_list': extracted_data.get('accessories_list') or [],
                'filename': filename
            })
            return ParsedOffer(**extracted_data)
        except Exception as e:
            logger.error(f"Error during Gemini API call for {filename}: {str(e)}")
            logger.error(traceback.format_exc())
            return ParsedOffer(filename=filename, warnings=[f"LLM parsing failed: {str(e)}"], parsing_confidence=0.1)

