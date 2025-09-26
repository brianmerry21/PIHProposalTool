#!/usr/bin/env python3
import os
import sys
import logging
import tempfile
from pypdf import PdfReader, PdfWriter
import shutil

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def append_marketing_pdf(main_pdf_path, marketing_pdf_path='static/assets/modula_vlm_marketing.pdf'):
    """
    Append a marketing PDF to the end of the main PDF
    
    Args:
        main_pdf_path: Path to the main PDF file
        marketing_pdf_path: Path to the marketing PDF to append
    
    Returns:
        Path to the combined PDF file
    """
    logger.info(f"Appending marketing PDF {marketing_pdf_path} to {main_pdf_path}")
    try:
        if not os.path.exists(marketing_pdf_path):
            logger.error(f"Marketing PDF not found at {marketing_pdf_path}")
            return main_pdf_path
            
        if not os.path.exists(main_pdf_path):
            logger.error(f"Main PDF not found at {main_pdf_path}")
            return main_pdf_path
            
        # Create a PDF writer object
        writer = PdfWriter()
        
        # Add pages from main PDF
        reader = PdfReader(main_pdf_path)
        for page_num in range(len(reader.pages)):
            writer.add_page(reader.pages[page_num])
        
        # Add pages from marketing PDF
        marketing_reader = PdfReader(marketing_pdf_path)
        for page_num in range(len(marketing_reader.pages)):
            writer.add_page(marketing_reader.pages[page_num])
        
        # Create temporary file for the output
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
            combined_pdf_path = tmp_file.name
            
        # Write the combined PDF to the temporary file
        with open(combined_pdf_path, 'wb') as output_pdf:
            writer.write(output_pdf)
        
        logger.info(f"Successfully created combined PDF at {combined_pdf_path}")
        
        # Copy the combined PDF to the original location
        shutil.copyfile(combined_pdf_path, main_pdf_path)
        os.unlink(combined_pdf_path)  # Remove temporary file
        
        logger.info(f"Marketing PDF successfully appended to {main_pdf_path}")
        return main_pdf_path
    except Exception as e:
        logger.error(f"Error appending marketing PDF: {str(e)}")
        return main_pdf_path

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python append_marketing_to_pdf.py <path_to_input_pdf> [path_to_marketing_pdf]")
        sys.exit(1)
    
    input_pdf = sys.argv[1]
    
    marketing_pdf = 'static/assets/modula_vlm_marketing.pdf'
    if len(sys.argv) >= 3:
        marketing_pdf = sys.argv[2]
    
    append_marketing_pdf(input_pdf, marketing_pdf)
    print(f"Successfully appended marketing PDF to {input_pdf}")