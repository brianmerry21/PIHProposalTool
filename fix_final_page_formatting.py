"""
This script applies additional formatting fixes to page 14:
1. Remove spacing between section headers and bullet points
2. Reduce bottom margin to 0.5 inches
"""
import os
import re
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def fix_final_page_formatting():
    """
    Apply additional formatting fixes to page 14
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
    
    # 1. Remove spacing between section headers and bullet points
    # Find the section that begins the page 14 content
    page_14_section_pattern = r"(# Set global style for page 14 - compact formatting.*?)# Add \"Proposal includes:\" heading"
    page_14_section = re.search(page_14_section_pattern, content, re.DOTALL)
    
    if page_14_section:
        # Add the bottom margin setting
        margin_setting = """        # Set smaller bottom margin for page 14
        for section in doc.sections:
            if section == doc.sections[-1]:  # Last section (page 14)
                section.bottom_margin = Inches(0.5)
        """
        
        # Insert margin setting after the global style for page 14
        modified_section = page_14_section.group(1) + margin_setting + "\n        # Add \"Proposal includes:\" heading"
        content = content.replace(page_14_section.group(0), modified_section)
    
    # 2. Remove the empty paragraph after "Proposal includes:" heading
    # This empty paragraph adds unwanted space
    empty_para_pattern = r"(        # Add \"Proposal includes:\" heading - black, underlined, smaller.*?underline = True.*?# Black color\n)        doc\.add_paragraph\(\"\"\)  # Keep first one but empty"
    
    empty_para_replacement = r"\1"
    content = re.sub(empty_para_pattern, empty_para_replacement, content, flags=re.DOTALL)
    
    # 3. Remove the empty paragraph after heading section for Proposal does not include
    empty_para2_pattern = r"(        # Add \"Proposal does not include:\" heading - black, underlined, smaller.*?underline = True.*?# Black color\n)        doc\.add_paragraph\(\)"
    
    empty_para2_replacement = r"\1"
    content = re.sub(empty_para2_pattern, empty_para2_replacement, content, flags=re.DOTALL)
    
    # 4. Remove the empty paragraph after the included items bulleted list
    empty_para3_pattern = r"(        for item in included_items:.*?p\.add_run\(item\).*?run\.font\.size = Pt\(10\)\n\n)        doc\.add_paragraph\(\)"
    
    empty_para3_replacement = r"\1"
    content = re.sub(empty_para3_pattern, empty_para3_replacement, content, flags=re.DOTALL)
    
    # 5. Remove the empty paragraph after excluded items bulleted list
    empty_para4_pattern = r"(        for item in excluded_items:.*?p\.add_run\(item\).*?run\.font\.size = Pt\(10\)\n\n)        doc\.add_paragraph\(\)"
    
    empty_para4_replacement = r"\1"
    content = re.sub(empty_para4_pattern, empty_para4_replacement, content, flags=re.DOTALL)
    
    # Check if we made any changes
    if content == original_content:
        logger.warning("No changes were made to the file")
        return False
    
    # Write the modified content back to the file
    with open(pdf_processor_path, 'w') as f:
        f.write(content)
    
    logger.info("Successfully applied additional formatting fixes to page 14")
    return True

if __name__ == "__main__":
    fix_final_page_formatting()