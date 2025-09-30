import io
import logging
import pdfplumber
import pytesseract
from PIL import Image
from typing import IO

logger = logging.getLogger(__name__)

def extract_text_from_pdf(pdf_bytes_io: IO[bytes]) -> str:
    """
    Extracts text from an uploaded PDF file using a hybrid approach.
    It first tries direct text extraction and falls back to OCR if needed.
    """
    full_text = ""
    pdf_bytes_io.seek(0)
    
    try:
        with pdfplumber.open(pdf_bytes_io) as pdf:
            # First, try a simple, fast extraction on the whole document
            for page in pdf.pages:
                page_text = page.extract_text(x_tolerance=2)
                if page_text:
                    full_text += page_text + "\n\n"
            
            # If simple extraction fails or gets very little text, try OCR
            if len(full_text.strip()) < 100:
                logger.warning("Initial text extraction yielded little result. Falling back to OCR.")
                full_text = "" # Reset text to rebuild with OCR results
                pdf_bytes_io.seek(0) # Reset buffer for page iteration
                with pdfplumber.open(pdf_bytes_io) as ocr_pdf:
                    for i, page in enumerate(ocr_pdf.pages):
                        try:
                            # Convert page to a high-resolution image
                            im = page.to_image(resolution=300)
                            # Use pytesseract to perform OCR on the image
                            # 'deu' for German, 'eng' for English.
                            page_text_ocr = pytesseract.image_to_string(im.original, lang='deu+eng')
                            if page_text_ocr:
                                full_text += page_text_ocr + "\n\n"
                                logger.info(f"Successfully extracted text from page {i+1} using OCR.")
                        except Exception as ocr_error:
                            logger.error(f"OCR failed for page {i+1}: {ocr_error}")
                            
        if not full_text.strip():
            logger.error("Both standard extraction and OCR failed to get text from the PDF.")
            
        return full_text
        
    except Exception as e:
        logger.error(f"Fatal error reading PDF: {e}", exc_info=True)
        return ""

