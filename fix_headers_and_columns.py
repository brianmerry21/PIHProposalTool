"""
This script fixes several issues with the document generation:
1. Removes "Options Not Included" from showing as a line item
2. Adjusts column widths for the options table
3. Fixes image sizing on page 7
4. Removes the recommended items note text
5. Fixes duplicate header issue on pages 11, 12, and 13
"""
import os
import logging
import re

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def fix_header_and_columns():
    """Fix header, column widths, and image sizing issues"""
    # Find the pdf_processor file
    pdf_processor_path = "utils/pdf_processor.py"
    if not os.path.exists(pdf_processor_path):
        logger.error(f"Could not find {pdf_processor_path}")
        return False
    
    logger.info(f"Found pdf_processor.py at {pdf_processor_path}")
    
    # Read the file
    with open(pdf_processor_path, 'r') as f:
        content = f.read()
    
    # 1. Fix the column widths in the options table
    column_widths_pattern = r"width_price_ea = Inches\(0\.85\).*?width_option = Inches\(1\.0\)"
    column_widths_replacement = "width_price_ea = Inches(0.95)  # 15% for price each\n            width_qty = Inches(0.65)       # 10% for quantity\n            width_option = Inches(0.9)     # 13% for option column"
    
    content = re.sub(column_widths_pattern, column_widths_replacement, content, flags=re.DOTALL)
    logger.info("Updated column widths")
    
    # 2. Remove the "Note: Items marked in bold..." text
    recommended_note_pattern = r"# Add note about recommended items.*?note_para\.add_run\('Note: Items marked in bold.*?'\)\.italic = True"
    
    if recommended_note_pattern in content:
        content = re.sub(recommended_note_pattern, "# Note about recommended items removed", content, flags=re.DOTALL)
        logger.info("Removed 'Note: Items marked in bold...' text")
    
    # 3. Fix the image extraction for page 7 (PDF page 5)
    # Make sure we're only cropping the top margin, not sides or bottom
    image_extraction_pattern = r"Cropping header, sides, and bottom: Removing top 250px, 100px from sides, and 100px from bottom"
    
    if image_extraction_pattern in content:
        # Find the extract_page_as_image method in image_processor.py
        image_processor_path = "utils/image_processor.py"
        if os.path.exists(image_processor_path):
            with open(image_processor_path, 'r') as f:
                image_processor_content = f.read()
            
            # Modify the extract_page_as_image_special function to only crop the top
            special_extract_pattern = r"def extract_page_as_image_special\(.*?\):.*?width, height = img\.size.*?cropped_img = img\.crop\(\(.*?\)\)"
            special_extract_replacement = """def extract_page_as_image_special(pdf_path, page_num, output_path, dpi=200):
    '''
    Extract a page from a PDF as an image with minimal cropping.
    Only removes the header, preserves sides and bottom.
    Specifically for page 7 of the document.
    
    Args:
        pdf_path: Path to the PDF file
        page_num: Page number to extract (1-indexed)
        output_path: Path to save the extracted image
        dpi: DPI for the extracted image (default: 200)
        
    Returns:
        Path to the extracted image or None on error
    '''
    try:
        # Convert PDF page to image
        from pdf2image import convert_from_path
        
        # Extract just the requested page
        images = convert_from_path(
            pdf_path, 
            dpi=dpi, 
            first_page=page_num, 
            last_page=page_num
        )
        
        if not images:
            logger.error(f"Failed to extract page {page_num} from {pdf_path}")
            return None
            
        # Get the first (and only) image
        img = images[0]
        
        # For page 7, only crop the top header (top 170px) - preserve sides and bottom completely
        width, height = img.size
        cropped_img = img.crop((0, 170, width, height))"""
            
            # Update the extract_page_as_image_special function
            image_processor_content = re.sub(special_extract_pattern, special_extract_replacement, image_processor_content, flags=re.DOTALL)
            
            # Write the changes to image_processor.py
            with open(image_processor_path, 'w') as f:
                f.write(image_processor_content)
            
            logger.info("Updated extract_page_as_image_special function to only crop the top header")
    
    # 4. Fix the duplicate header issue on pages 11, 12, and 13
    # Look for the code that sets up the document header and footer
    
    # First, find where we add headers to section
    header_setup_pattern = r"def add_page_header\(doc, customer_info\):.*?return header_paragraph"
    header_setup_match = re.search(header_setup_pattern, content, re.DOTALL)
    
    if header_setup_match:
        header_setup_code = header_setup_match.group(0)
        
        # Check if we need to modify the add_page_header function
        if "doc.sections[-1].different_first_page = True" not in header_setup_code:
            # Modify the add_page_header function to ensure no duplicate headers
            modified_header_code = header_setup_code.replace(
                "def add_page_header(doc, customer_info):",
                """def add_page_header(doc, customer_info):
    # Make sure we're not using any previous header
    for section in doc.sections:
        section.header.is_linked_to_previous = False
        # Different first page to avoid duplicating header on first page
        section.different_first_page = True"""
            )
            
            # Replace the old header setup code with the modified version
            content = content.replace(header_setup_code, modified_header_code)
            logger.info("Updated header setup to prevent duplicate headers")
    
    # 5. Check if we need to filter out "Options Not Included" line items
    # Look for where we set up the optional items for the document
    if "# Filter optional items to exclude section headers and 'Options Not Included'" not in content:
        optional_items_pattern = r"(\s+)optional_items_data = extract_optional_items\(extracted_pdf_text\)"
        
        if re.search(optional_items_pattern, content):
            indent = re.search(optional_items_pattern, content).group(1)
            filtering_code = f"{indent}# Filter optional items to exclude section headers and 'Options Not Included'\n"
            filtering_code += f"{indent}optional_items_data = [item for item in optional_items_data if 'Options Not Included' not in item.get('description', '')]"
            
            content = re.sub(
                optional_items_pattern,
                f"\\1optional_items_data = extract_optional_items(extracted_pdf_text)\n{filtering_code}",
                content
            )
            logger.info("Added filtering to remove 'Options Not Included' from optional items")
    
    # Write all the changes back to the file
    with open(pdf_processor_path, 'w') as f:
        f.write(content)
    
    logger.info("Successfully applied all fixes")
    return True

if __name__ == "__main__":
    fix_header_and_columns()