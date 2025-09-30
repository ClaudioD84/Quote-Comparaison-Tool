import json
import logging
import traceback
import openai
from typing import Dict

# Import the data model and the asdict function to help with structuring
from modules.models import ParsedOffer, to_dict_helper

# Get a logger for this module
logger = logging.getLogger(__name__)

class LLMParser:
    """
    A class to handle all interactions with the OpenAI Large Language Model (LLM).
    """

    def __init__(self, api_key: str):
        """
        Initializes the OpenAI client with the provided API key.
        """
        if not api_key:
            raise ValueError("An OpenAI API key is required.")
        self.client = openai.OpenAI(api_key=api_key)
        logger.info("OpenAI client configured successfully.")

    def parse_text(self, text: str, filename: str, prompt_template: str) -> ParsedOffer:
        """
        Sends the extracted text to the OpenAI API and parses the structured JSON response.
        """
        logger.info(f"Sending text from '{filename}' to OpenAI for parsing using a customer-specific recipe.")
        
        # This is the standard structure we expect the AI to return.
        # We will describe it in the prompt.
        json_schema_description = json.dumps(to_dict_helper(ParsedOffer(filename="")), indent=2)

        # The final prompt combines the user's recipe with instructions for the AI
        final_prompt = f"""
        {prompt_template}

        **Final Output Requirement:**
        Your final response MUST be a single, valid JSON object that strictly follows the structure shown below. Do not include any text, explanations, or markdown formatting before or after the JSON object.

        **JSON Structure to Follow:**
        {json_schema_description}

        <DOCUMENT_TO_PARSE>
        {text}
        </DOCUMENT_TO_PARSE>
        """

        try:
            # Make the API call to the OpenAI model
            response = self.client.chat.completions.create(
                model="gpt-4o",
                messages=[{"role": "user", "content": final_prompt}],
                # CORRECTED: The 'schema' parameter is not supported by OpenAI's API.
                # JSON mode is enabled, and the structure is enforced via the prompt.
                response_format={"type": "json_object"},
                temperature=0.0,
            )
            
            content = response.choices[0].message.content
            extracted_data = json.loads(content)
            
            # Ensure list fields exist to prevent errors later
            extracted_data['options_list'] = extracted_data.get('options_list') or []
            extracted_data['accessories_list'] = extracted_data.get('accessories_list') or []
            extracted_data['filename'] = filename

            return ParsedOffer(**extracted_data)

        except Exception as e:
            logger.error(f"Error during OpenAI API call for {filename}: {e}")
            logger.error(traceback.format_exc())
            return ParsedOffer(
                filename=filename,
                warnings=[f"LLM parsing failed: {e}"],
                parsing_confidence=0.1
            )

