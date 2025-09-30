import json
import logging
import traceback
import openai  # UPDATED: Switched from google.generativeai to openai
from typing import Dict

# Import the data model to structure the output
from modules.models import ParsedOffer

# Get a logger for this module
logger = logging.getLogger(__name__)

class LLMParser:
    """
    A class to handle all interactions with the OpenAI Large Language Model (LLM).
    It executes the detailed instructions provided by a recipe prompt.
    """

    def __init__(self, api_key: str):
        """Initializes the OpenAI client with the provided API key."""
        if not api_key:
            raise ValueError("An OpenAI API key is required to use the LLM parser.")
        # UPDATED: Configure the OpenAI client
        self.client = openai.OpenAI(api_key=api_key)
        logger.info("OpenAI client configured successfully.")

    def parse_text(self, text: str, filename: str, prompt_template: str) -> ParsedOffer:
        """
        Sends the extracted text and a specific recipe prompt to the OpenAI API
        and parses the structured JSON response into a ParsedOffer object.
        """
        logger.info(f"Sending text from '{filename}' to OpenAI for parsing using a customer-specific recipe.")
        
        prompt_text = f"{prompt_template}\n\n<DOCUMENT_TO_PARSE>\n{text}\n</DOCUMENT_TO_PARSE>"

        # This JSON schema enforces a consistent output structure from the AI.
        # It is compatible with OpenAI's JSON mode.
        json_schema = {
            "type": "object",
            "properties": {
                "filename": {"type": "string"}, "vendor": {"type": "string"}, "vehicle_description": {"type": "string"},
                "max_duration_months": {"type": "number"}, "max_total_mileage": {"type": "number"}, "offer_duration_months": {"type": "number"},
                "offer_total_mileage": {"type": "number"}, "monthly_rental": {"type": "number"}, "total_monthly_lease": {"type": "number"},
                "currency": {"type": "string"}, "upfront_costs": {"type": "number"}, "deposit": {"type": "number"},
                "admin_fees": {"type": "number"}, "excess_mileage_rate": {"type": "number"}, "unused_mileage_rate": {"type": "number"},
                "quote_number": {"type": "string"}, "manufacturer": {"type": "string"}, "model": {"type": "string"},
                "version": {"type": "string"}, "internal_colour": {"type": "string"}, "external_colour": {"type": "string"},
                "fuel_type": {"type": "string"}, "num_doors": {"type": "number"}, "hp": {"type": "number"}, "c02_emission": {"type": "number"},
                "battery_range": {"type": "number"}, "vehicle_price": {"type": "number"}, "options_price": {"type": "number"},
                "accessories_price": {"type": "number"}, "delivery_cost": {"type": "number"}, "registration_tax": {"type": "number"},
                "total_net_investment": {"type": "number"}, "taxation_value": {"type": "number"}, "financial_rate": {"type": "number"},
                "depreciation_interest": {"type": "number"}, "maintenance_repair": {"type": "number"}, "insurance_cost": {"type": "number"},
                "green_tax": {"type": "number"}, "management_fee": {"type": "number"}, "roadside_assistance": {"type": "number"},
                "tyres_cost": {"type": "number"}, "maintenance_included": {"type": "boolean"}, "driver_name": {"type": "string"},
                "customer": {"type": "string"},
                "options_list": {"type": "array", "items": {"type": "object", "properties": {"name": {"type": "string"}, "price": {"type": "number"}}}},
                "accessories_list": {"type": "array", "items": {"type": "object", "properties": {"name": {"type": "string"}, "price": {"type": "number"}}}},
                "parsing_confidence": {"type": "number"}, "warnings": {"type": "array", "items": {"type": "string"}},
            }
        }
        
        try:
            # UPDATED: API call changed to the OpenAI client format
            response = self.client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": "You are an expert financial analyst. Your task is to extract data from leasing offers and return it ONLY as a valid JSON object that conforms to the provided schema. Do not include any other text or explanations in your response."},
                    {"role": "user", "content": prompt_text}
                ],
                response_format={"type": "json_object", "schema": json_schema}
            )
            extracted_data = json.loads(response.choices[0].message.content)
            
            extracted_data.update({
                'options_list': extracted_data.get('options_list') or [],
                'accessories_list': extracted_data.get('accessories_list') or [],
                'filename': filename
            })
            return ParsedOffer(**extracted_data)

        except Exception as e:
            logger.error(f"Error during OpenAI API call for {filename}: {str(e)}")
            logger.error(traceback.format_exc())
            return ParsedOffer(filename=filename, warnings=[f"LLM parsing failed: {str(e)}"], parsing_confidence=0.1)

