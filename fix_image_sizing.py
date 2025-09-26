"""
This script modifies the image extraction to only crop the top header
for page 7, preserving the sides and bottom.
"""
import os
import logging
import re

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def fix_image_sizing():
    """Fix image sizing for page 7 to prevent it from pushing to next page"""
    image_processor_path = "utils/image_processor.py"
    if not os.path.exists(image_processor_path):
        logger.error(f"Could not find {image_processor_path}")
        return False
    
    logger.info(f"Found image_processor.py at {image_processor_path}")
    
    # Read the file
    with open(image_processor_path, 'r') as f:
        content = f.read()
    
    # Create/modify a special extraction function for page 5 (page 7 in document)
    if "extract_page_as_image_special" not in content:
        # Add the special function
        special_function = """
def extract_page_as_image_special(pdf_path, page_num, output_path, dpi=200):
    '''
    Extract a page from a PDF as an image with minimal cropping.
    Only removes the header, preserves sides and bottom completely.
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
        
        # Only crop the top header (top 170px) - preserve sides and bottom completely
        width, height = img.size
        cropped_img = img.crop((0, 170, width, height))
        
        # Save the cropped image
        cropped_img.save(output_path)
        logger.info(f"Saved page {page_num} image with special handling: {output_path}")
        
        return output_path
    except Exception as e:
        logger.error(f"Error extracting page as image with special handling: {e}")
        return None
"""
        
        # Add the function to the end of the file
        content += special_function
        logger.info("Added special image extraction function")
    else:
        # Modify the existing function to ensure it only crops the top
        special_function_pattern = r"def extract_page_as_image_special\(.*?\):.*?cropped_img = img\.crop\(\(.*?\)\)"
        
        if re.search(special_function_pattern, content, re.DOTALL):
            # Update the crop parameters to only crop the top
            modified_crop = """        # Only crop the top header (top 170px) - preserve sides and bottom completely
        width, height = img.size
        cropped_img = img.crop((0, 170, width, height))"""
            
            # Find the existing crop code
            crop_pattern = r".*width, height = img\.size.*?cropped_img = img\.crop\(\(.*?\)\)"
            
            # Replace with our modified crop
            content = re.sub(crop_pattern, modified_crop, content, re.DOTALL)
            logger.info("Updated crop parameters in special image extraction function")
    
    # Write the changes back to the file
    with open(image_processor_path, 'w') as f:
        f.write(content)
    
    # Now update pdf_processor.py to use our special function for page 5
    pdf_processor_path = "utils/pdf_processor.py"
    if os.path.exists(pdf_processor_path):
        with open(pdf_processor_path, 'r') as f:
            pdf_content = f.read()
        
        # Ensure we import the special function
        import_pattern = r"from utils\.image_processor import (.*)"
        import_match = re.search(import_pattern, pdf_content)
        
        if import_match:
            imports = import_match.group(1)
            if "extract_page_as_image_special" not in imports:
                # Add our function to the imports
                new_imports = imports + ", extract_page_as_image_special"
                pdf_content = pdf_content.replace(imports, new_imports)
                logger.info("Added extract_page_as_image_special to imports")
        
        # Find where page 5 is extracted and use our special function
        page5_pattern = r"# Extract the full page 5 of the PDF[\s\S]*?extracted_image = extract_pdf_region_as_image\(pdf_path, page_num=5,.*?\)"
        
        if re.search(page5_pattern, pdf_content):
            # Replace with our special function
            special_extraction = """# Extract the full page 5 of the PDF with special handling to preserve sides and bottom
            try:
                # Create a temporary file to save the image
                with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as temp_img:
                    temp_path = temp_img.name

                # Use special extraction for page 5 to only crop the header
                extracted_image = extract_page_as_image_special(
                    pdf_path, 
                    5,  # Page number
                    temp_path,
                    dpi=200  # Lower DPI to prevent timeouts
                )"""
            
            pdf_content = re.sub(page5_pattern, special_extraction, pdf_content)
            logger.info("Updated page 5 extraction to use special function")
        
        # Write the changes back to the file
        with open(pdf_processor_path, 'w') as f:
            f.write(pdf_content)
    
    logger.info("Successfully updated image extraction for page 7")
    return True

if __name__ == "__main__":
    fix_image_sizing()