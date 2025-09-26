"""
This script fixes the duplicate header issue on pages 11, 12, and 13
"""
import os
import logging
import re

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def fix_duplicate_headers():
    """Fix the duplicate header issue on pages 11, 12, and 13"""
    pdf_processor_path = "utils/pdf_processor.py"
    if not os.path.exists(pdf_processor_path):
        logger.error(f"Could not find {pdf_processor_path}")
        return False
    
    logger.info(f"Found pdf_processor.py at {pdf_processor_path}")
    
    # Read the file
    with open(pdf_processor_path, 'r') as f:
        content = f.read()
    
    # Check the add_page_header function
    header_function_pattern = r"def add_page_header\(doc, customer_info\):[\s\S]*?return header_paragraph"
    header_function_match = re.search(header_function_pattern, content)
    
    if header_function_match:
        header_function = header_function_match.group(0)
        
        # Check if we need to add code to prevent duplicate headers
        if "section.different_first_page = True" not in header_function:
            # Add code to prevent duplicate headers
            updated_header_function = header_function.replace(
                "def add_page_header(doc, customer_info):",
                """def add_page_header(doc, customer_info):
    # Remove any existing headers
    for section in doc.sections:
        # Make sure we're not inheriting headers from previous sections
        section.header.is_linked_to_previous = False
        section.different_first_page = True
        
        # Clear any existing content in the header
        section.header.paragraphs.clear()"""
            )
            
            # Replace the function with our updated version
            content = content.replace(header_function, updated_header_function)
            logger.info("Updated add_page_header function to prevent duplicate headers")
    
    # Also find where we set up the actual header in the process_pdf_to_word function
    pricing_sections_pattern = r"# ======== PAGE \d+: PRICING TABLE ========[\s\S]*?add_page_header\(doc, customer_info\)"
    
    if re.search(pricing_sections_pattern, content):
        # Ensure headers are cleared before adding new ones
        modified_content = re.sub(
            r"(# ======== PAGE \d+: PRICING TABLE ========[\s\S]*?)add_page_header\(doc, customer_info\)",
            r"\1# Clear any existing headers\n        for section in doc.sections:\n            section.header.is_linked_to_previous = False\n            section.header.paragraphs.clear()\n        \n        add_page_header(doc, customer_info)",
            content
        )
        
        if modified_content != content:
            content = modified_content
            logger.info("Added code to clear headers before adding new ones")
    
    # Look for any code adding header/footer after page 10
    header_setup_pattern = r"add_page_header\(doc, customer_info\)"
    page_count = len(re.findall(header_setup_pattern, content))
    
    if page_count > 3:  # If headers are being added more than 3 times, it's likely causing duplicates
        logger.info(f"Found {page_count} header additions - reducing to prevent duplicates")
        
        # Modify the document creation to handle header inheritance better
        doc_creation_pattern = r"# Create the Word document[\s\S]*?doc = Document\(\)"
        
        if re.search(doc_creation_pattern, content):
            updated_doc_creation = """    # Create the Word document
    try:
        doc = Document()
        
        # Configure document to properly handle headers
        for section in doc.sections:
            section.different_first_page = True"""
            
            content = re.sub(doc_creation_pattern, updated_doc_creation, content)
            logger.info("Updated document creation to handle headers better")
    
    # Write the changes back to the file
    with open(pdf_processor_path, 'w') as f:
        f.write(content)
    
    logger.info("Successfully applied fixes for duplicate headers")
    return True

if __name__ == "__main__":
    fix_duplicate_headers()