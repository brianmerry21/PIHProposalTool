"""
This script adds a specialized image extraction function for page 7 (PDF page 5)
that only crops the header but preserves the sides and bottom of the image.
"""
import os
import logging
import re

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def fix_page7_image():
    """Add a specialized image extraction function for page 7"""
    pdf_processor_path = "utils/pdf_processor.py"
    utils_image_processor_path = "utils/image_processor.py"
    
    if os.path.exists(utils_image_processor_path):
        # We'll add our specialized function to the image_processor.py file
        logger.info(f"Found image_processor.py at {utils_image_processor_path}")
        
        # Read the file
        with open(utils_image_processor_path, 'r') as f:
            image_processor_content = f.read()
        
        # Check if we already have a specialized function
        if "extract_page_as_image_special" in image_processor_content:
            logger.info("Specialized image function already exists")
            return True
        
        # Add a specialized image extraction function
        special_func = """
def extract_page_as_image_special(pdf_path, page_num, output_path, dpi=200):
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
        
        # For page 7, only crop the header (top 250px) - preserve sides and bottom
        width, height = img.size
        cropped_img = img.crop((0, 250, width, height))
        
        # Save the cropped image
        cropped_img.save(output_path)
        logger.info(f"Saved page {page_num} image with special handling: {output_path}")
        
        return output_path
    except Exception as e:
        logger.error(f"Error extracting page as image with special handling: {e}")
        return None
"""
        
        # Add the function to the end of the file
        image_processor_content += special_func
        
        # Write the updated content back
        with open(utils_image_processor_path, 'w') as f:
            f.write(image_processor_content)
            
        logger.info("Added specialized image extraction function to image_processor.py")
    
    # Now update pdf_processor.py to use the specialized function for page 5 (which is page 7 in the document)
    if not os.path.exists(pdf_processor_path):
        logger.error(f"Could not find {pdf_processor_path}")
        return False
    
    logger.info(f"Found pdf_processor.py at {pdf_processor_path}")
    
    # Read the file
    with open(pdf_processor_path, 'r') as f:
        pdf_processor_content = f.read()
    
    # Update the imports to include our specialized function
    if "from utils.image_processor import extract_pdf_region_as_image" in pdf_processor_content:
        old_import = "from utils.image_processor import extract_pdf_region_as_image"
        if ", extract_page_as_image_special" in old_import:
            logger.info("Import for specialized function already exists")
        else:
            new_import = "from utils.image_processor import extract_pdf_region_as_image, extract_page_as_image_special"
            pdf_processor_content = pdf_processor_content.replace(old_import, new_import)
            logger.info("Updated imports in pdf_processor.py")
    elif "from utils.image_processor import" in pdf_processor_content:
        # Find the existing import line
        import_pattern = r"from utils\.image_processor import .*"
        import_match = re.search(import_pattern, pdf_processor_content)
        if import_match:
            old_import = import_match.group(0)
            # Append our function to the import
            if "extract_page_as_image_special" not in old_import:
                new_import = old_import + ", extract_page_as_image_special"
                pdf_processor_content = pdf_processor_content.replace(old_import, new_import)
                logger.info("Updated imports in pdf_processor.py")
    
    # Find the code that extracts pages and modify it
    extract_pattern = r"# Extract specific pages from the PDF as images[\s\S]*?for page_num in range\(\d+, \d+\):"
    extract_match = re.search(extract_pattern, pdf_processor_content)
    
    if extract_match:
        # Find the page extraction loop
        loop_start = extract_match.end()
        loop_code = pdf_processor_content[loop_start:loop_start + 600]
        
        # Look for img_path assignment
        img_path_pattern = r"img_path = .*?extract_page_as_image\("
        img_path_match = re.search(img_path_pattern, loop_code)
        
        if img_path_match:
            # Find the entire img_path assignment
            img_path_start = loop_code.find(img_path_match.group(0))
            img_path_end = loop_code.find(")", img_path_start) + 1
            img_path_code = loop_code[img_path_start:img_path_end]
            
            # Create a conditional version
            conditional_code = """            # Use special handling for page 5 (which is page 7 in the document)
            if page_num == 5:
                img_path = extract_page_as_image_special(
                    pdf_path, 
                    page_num, 
                    temp.name,
                    dpi=200
                )
            else:
                img_path = extract_page_as_image(
                    pdf_path, 
                    page_num, 
                    temp.name, 
                    dpi=200
                )"""
            
            # Replace the img_path assignment with our conditional version
            modified_loop_code = loop_code.replace(img_path_code, conditional_code)
            
            # Replace the loop code in the full content
            pdf_processor_content = pdf_processor_content.replace(loop_code, modified_loop_code)
            logger.info("Updated page extraction loop to use specialized function for page 5")
    
    # Write the updated content back
    with open(pdf_processor_path, 'w') as f:
        f.write(pdf_processor_content)
    
    logger.info("Successfully updated pdf_processor.py to fix page 7 image")
    return True

if __name__ == "__main__":
    fix_page7_image()