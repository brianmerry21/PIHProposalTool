"""
This script modifies the image sizes in the pdf_processor.py file to make them larger,
extending into the margins by 0.3 inches on each side.
"""

import os
import re
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def fix_image_sizes():
    """Increase the size of all images in the document by 0.3 inches on each side"""
    pdf_processor_path = "utils/pdf_processor.py"
    
    if not os.path.exists(pdf_processor_path):
        logger.error(f"Could not find {pdf_processor_path}")
        return False
        
    logger.info(f"Reading {pdf_processor_path}...")
    
    with open(pdf_processor_path, 'r') as f:
        content = f.read()
    
    # Pattern to find available_width calculation and image addition
    pattern = r"([ \t]+)available_width = section\.page_width - section\.left_margin - section\.right_margin\n\1doc\.add_picture\(([^,]+), width=available_width\)"
    
    # Replacement with extended width
    replacement = r"\1available_width = section.page_width - section.left_margin - section.right_margin\n\1# Increase width by 0.6 inches (0.3 inches on each side) to extend into margins\n\1extended_width = available_width + Inches(0.6)\n\1doc.add_picture(\2, width=extended_width)"
    
    # Apply the replacement
    modified_content = re.sub(pattern, replacement, content)
    
    # Also increase the cutsheet image size
    cutsheet_pattern = r"([ \t]+)doc\.add_picture\(page_nine_image_path, width=Inches\(6\.5\)\)"
    cutsheet_replacement = r"\1doc.add_picture(page_nine_image_path, width=Inches(7.1))  # Increased from 6.5 to 7.1 inches"
    
    # Apply the cutsheet replacement
    modified_content = re.sub(cutsheet_pattern, cutsheet_replacement, modified_content)
    
    # Check if we made any changes
    if content == modified_content:
        logger.warning("No changes were made to the file")
        return False
    
    # Write the modified content back to the file
    with open(pdf_processor_path, 'w') as f:
        f.write(modified_content)
    
    logger.info("Successfully increased image sizes")
    return True

if __name__ == "__main__":
    fix_image_sizes()