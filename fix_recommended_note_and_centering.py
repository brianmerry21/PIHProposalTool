"""
This script adds back the recommended items note and centers specific headings
in the Optional Items section.
"""
import os
import re
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def fix_recommended_note_and_centering():
    """
    Add back the note about recommended items and center the headings
    in the Optional Items section.
    """
    pdf_processor_path = "utils/pdf_processor.py"
    
    if not os.path.exists(pdf_processor_path):
        logger.error(f"Could not find {pdf_processor_path}")
        return False
        
    logger.info(f"Reading {pdf_processor_path}...")
    
    with open(pdf_processor_path, 'r') as f:
        content = f.read()
        
    # Store original content for comparison later
    original_content = content
    
    # 1. First, center the "Optional Items & Accessories Pricing" heading
    # Find the line that creates this heading
    optional_title_pattern = r"([ \t]+)doc\.add_heading\('Optional Items & Accessories Pricing', level=1\)"
    optional_title_replacement = r"\1heading = doc.add_heading('Optional Items & Accessories Pricing', level=1)\n\1heading.alignment = WD_ALIGN_PARAGRAPH.CENTER"
    
    # Apply the centering replacement
    content = re.sub(optional_title_pattern, optional_title_replacement, content)
    
    # 2. Add the recommended items note below the options table
    # This was likely removed in a previous fix, so we need to add it back
    # Find the line after the options table that adds lead time info
    lead_time_pattern = r"([ \t]+# Note about recommended items removed\n\n[ \t]+# Add lead time information\n[ \t]+lead_time_para = doc\.add_paragraph\(\)\n[ \t]+)lead_time_para\.add_run\(\\"(\nLead Time Currently 10-12 weeks)\\"\)\.bold = True"
    
    # Replacement with recommendation note and centered lead time
    lead_time_replacement = r"\1# Add note about recommended items\n\1recommended_note_para = doc.add_paragraph()\n\1recommended_note_para.add_run(\"Note: Items marked in bold on \\\"Options Not Included\\\" are highly recommended by PIH for this project.\")\n\1\n\1# Add lead time information\n\1lead_time_para = doc.add_paragraph()\n\1lead_time_para.alignment = WD_ALIGN_PARAGRAPH.CENTER\n\1lead_time_para.add_run(\"\2\").bold = True"
    
    # Apply the note and centering replacement
    content = re.sub(lead_time_pattern, lead_time_replacement, content)
    
    # Check if we made any changes
    if content == original_content:
        logger.warning("No changes were made to the file")
        return False
    
    # Write the modified content back to the file
    with open(pdf_processor_path, 'w') as f:
        f.write(content)
    
    logger.info("Successfully added recommended items note and centered headings")
    return True

if __name__ == "__main__":
    fix_recommended_note_and_centering()