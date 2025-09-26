"""
This script specifically fixes:
1. Notes section font size (change from 11pt to 10pt)
2. Remove remaining spacing between sections on page 14
"""
import os
import re
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def fix_notes_and_spacing():
    """
    Fix Notes font size and remove all remaining spacing between sections
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
    
    # 1. Fix Note 1 text font size
    note1_pattern = r"(        # Add Note 1\n        note1_para = doc\.add_paragraph\(\).*?\n        note1_para\.paragraph_format\.space_before = Pt\(0\)\n        note1_para\.paragraph_format\.space_after = Pt\(0\).*?note1_para\.add_run\(\"Note 1:\"\)\.bold = True\n        note1_para\.add_run\(\"\\tAll Permits are \\\"By Customer\\\".*?\"\))"
    note1_replacement = r"\1\n        # Ensure 10pt font size for all note text\n        for run in note1_para.runs:\n            run.font.size = Pt(10)"
    content = re.sub(note1_pattern, note1_replacement, content, flags=re.DOTALL)
    
    # 2. Fix Note 2 text font size
    note2_pattern = r"(        # Add Note 2\n        note2_para = doc\.add_paragraph\(\).*?\n        note2_para\.paragraph_format\.space_before = Pt\(0\)\n        note2_para\.paragraph_format\.space_after = Pt\(0\).*?note2_para\.add_run\(\"Note 2:\"\)\.bold = True\n        note2_para\.add_run\(\"\\tSeismic calculations for the equipment.*?\"\))"
    note2_replacement = r"\1\n        # Ensure 10pt font size for all note text\n        for run in note2_para.runs:\n            run.font.size = Pt(10)"
    content = re.sub(note2_pattern, note2_replacement, content, flags=re.DOTALL)
    
    # 3. Remove spacing between "Lead Time:" and "Payment Schedule:"
    lead_time_pattern = r"(        # Add Lead Time section - more compact\n        lead_time_heading = doc\.add_paragraph\(\).*?\n.*?\"Lead Time\"\)\.bold = True\n        lead_time_heading\.add_run\(\": 10-12 Weeks\"\).*?\n.*?run\.font\.size = Pt\(10\)\n\n)        # Add Payment Schedule section\n"
    lead_time_replacement = r"\1        # Add Payment Schedule section directly after Lead Time\n"
    content = re.sub(lead_time_pattern, lead_time_replacement, content, flags=re.DOTALL)
    
    # 4. Fix the lead time section to remove extra space before and after
    lead_time_fix_pattern = r"        # Add Lead Time section - more compact\n        lead_time_heading = doc\.add_paragraph\(\)"
    lead_time_fix_replacement = r"        # Add Lead Time section - more compact with minimal spacing\n        lead_time_heading = doc.add_paragraph()\n        lead_time_heading.paragraph_format.space_before = Pt(0)\n        lead_time_heading.paragraph_format.space_after = Pt(0)"
    content = content.replace(lead_time_fix_pattern, lead_time_fix_replacement)
    
    # 5. Remove excessive paragraph spacing between payment schedule and notes
    payment_spacing_pattern = r"(        for item in payment_schedule:.*?p\.paragraph_format\.space_after = Pt\(0\).*?run\.font\.size = Pt\(10\)\n\n)        # Add Note 1"
    payment_spacing_replacement = r"\1        # Add Note 1"
    content = re.sub(payment_spacing_pattern, payment_spacing_replacement, content, flags=re.DOTALL)
    
    # 6. Remove excessive paragraph spacing between Note 1 and Note 2
    note1_spacing_pattern = r"(        # Add Note 1.*?note1_para\.add_run\(\"\\tAll Permits are \\\"By Customer\\\".*?\"\).*?run\.font\.size = Pt\(10\)\n\n)        # Add Note 2"
    note1_spacing_replacement = r"\1        # Add Note 2"
    content = re.sub(note1_spacing_pattern, note1_spacing_replacement, content, flags=re.DOTALL)
    
    # Check if we made any changes
    if content == original_content:
        logger.warning("No changes were made to the file")
        return False
    
    # Write the modified content back to the file
    with open(pdf_processor_path, 'w') as f:
        f.write(content)
    
    logger.info("Successfully fixed Notes font size and spacing issues")
    return True

if __name__ == "__main__":
    fix_notes_and_spacing()