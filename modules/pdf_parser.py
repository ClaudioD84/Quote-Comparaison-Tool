import io
import pdfplumber
from typing import IO

def extract_text_from_pdf(pdf_bytes: bytes) -> str:
    """
    Extracts all text from the pages of a PDF file.

    Args:
        pdf_bytes: The raw bytes of the PDF file.

    Returns:
        A single string containing all the extracted text from the PDF.
        Returns an empty string if the PDF cannot be read or contains no text.
    """
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            # Extract text from each page and handle pages with no text (None)
            pages_text = [page.extract_text() or "" for page in pdf.pages]
            full_text = "\n".join(pages_text)
            return full_text
    except Exception as e:
        # If any error occurs during PDF processing, return an empty string.
        # The main app will handle the warning to the user.
        print(f"Error extracting text from PDF: {e}")
        return ""

