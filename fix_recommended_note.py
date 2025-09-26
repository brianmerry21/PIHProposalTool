"""
This script removes the recommended note from the document
"""
import os
import logging
import re

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def remove_recommended_note():
    """Remove the recommended items note from the document"""
    pdf_processor_path = "utils/pdf_processor.py"
    if not os.path.exists(pdf_processor_path):
        logger.error(f"Could not find {pdf_processor_path}")
        return False
    
    logger.info(f"Found pdf_processor.py at {pdf_processor_path}")
    
    # Read the file
    with open(pdf_processor_path, 'r') as f:
        content = f.read()
    
    # Find the section that adds the note about recommended items
    note_pattern = r"# Add note about recommended items[\s\S]*?note_para\.add_run\('Note: Items marked in bold.*?\)\.italic = True"
    
    # Replace with a blank line
    if re.search(note_pattern, content):
        new_content = re.sub(note_pattern, "# Note about recommended items removed", content)
        
        # Write the changes back to the file
        with open(pdf_processor_path, 'w') as f:
            f.write(new_content)
        
        logger.info("Successfully removed the recommended items note")
        return True
    else:
        logger.warning("Could not find the recommended items note section")
        return False

if __name__ == "__main__":
    remove_recommended_note()