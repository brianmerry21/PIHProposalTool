"""
This script addresses three remaining issues:
1. Remove "Options Not Included" from appearing as a line item in the optional items section
2. Add Price Ea. column to options not included section and use "Option" instead of dash marks
3. Create a separate function for handling the image on page 7 without cropping sides/bottom
"""
import os
import sys
import logging
from pathlib import Path

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def fix_options_header_in_pdf_processor():
    """
    Fix the PDF processor to not include the "Options Not Included" header as a line item
    and format the options section correctly.
    """
    try:
        # Look for pdf_processor.py in both current directory and utils/
        processor_path = None
        possible_paths = ["utils/pdf_processor.py", "pdf_processor.py"]
        
        for path in possible_paths:
            if os.path.exists(path):
                processor_path = path
                break
        
        if not processor_path:
            logger.error("Could not find pdf_processor.py")
            return False
        
        logger.info(f"Found pdf_processor.py at {processor_path}")
        
        # Read the file
        with open(processor_path, 'r') as f:
            content = f.read()
        
        # 1. Fix the function that adds the pricing table to exclude section headers marked as optional
        # Find where we're creating the main pricing table
        if "# Add standard line items to the table" in content:
            logger.info("Fixing pricing table generation to exclude optional section headers")
            
            # Original pattern - likely includes all items including section headers marked as optional
            old_pattern = """    # Add standard line items to the table
    for i, item in enumerate(line_items):
        # Skip optional items, they'll go in a separate table
        if item.is_optional:
            continue"""
            
            # New pattern - specifically exclude section headers that are marked as optional
            new_pattern = """    # Add standard line items to the table
    for i, item in enumerate(line_items):
        # Skip optional items and section headers marked as optional,
        # they'll go in a separate table or be excluded
        if item.is_optional:
            continue
        
        # Skip section headers that are marked as optional
        if item.is_section_header and item.is_optional:
            continue"""
            
            # Replace the pattern
            if old_pattern in content:
                content = content.replace(old_pattern, new_pattern)
                logger.info("Updated standard line items logic")
            else:
                logger.warning("Could not find standard line items pattern to update")
        
        # 2. Fix the optional items table to use "Option" instead of dash marks and include Price Ea. column
        if "# Add optional items to a separate table" in content:
            logger.info("Fixing optional items table format")
            
            # Find the optional items table creation section
            # We'll search for a pattern where we're adding the headers or setting up the table
            old_optional_table_pattern = """    # Add optional items to a separate table
    if optional_items:
        doc.add_paragraph("").add_run("Options Not Included").bold = True
        
        # Create a table for optional items
        optional_table = doc.add_table(rows=1, cols=3)
        optional_table.style = 'Table Grid'
        
        # Set column widths
        set_column_width(optional_table.columns[0], 4.5)  # Item description
        set_column_width(optional_table.columns[1], 0.75)  # Quantity
        set_column_width(optional_table.columns[2], 1.25)  # Price
        
        # Add header row
        header_cells = optional_table.rows[0].cells
        header_cells[0].text = "Item"
        header_cells[1].text = "Qty."
        header_cells[2].text = "Prices"
        
        # Make header row bold
        for cell in header_cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.bold = True"""
            
            new_optional_table_pattern = """    # Add optional items to a separate table
    if optional_items:
        doc.add_paragraph("").add_run("Options Not Included").bold = True
        
        # Create a table for optional items
        optional_table = doc.add_table(rows=1, cols=4)  # Add a column for Price Ea.
        optional_table.style = 'Table Grid'
        
        # Set column widths
        set_column_width(optional_table.columns[0], 4.0)  # Item description
        set_column_width(optional_table.columns[1], 0.85)  # Price Ea.
        set_column_width(optional_table.columns[2], 0.65)  # Quantity
        set_column_width(optional_table.columns[3], 1.0)   # Price
        
        # Add header row
        header_cells = optional_table.rows[0].cells
        header_cells[0].text = "Item"
        header_cells[1].text = "Price Ea."
        header_cells[2].text = "Qty."
        header_cells[3].text = "Prices"
        
        # Make header row bold
        for cell in header_cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.bold = True"""
            
            if old_optional_table_pattern in content:
                content = content.replace(old_optional_table_pattern, new_optional_table_pattern)
                logger.info("Updated optional table pattern")
            else:
                logger.warning("Could not find optional table pattern to update")
            
            # Fix how we're adding rows to the optional items table
            old_optional_item_add = """        # Add data rows for optional items
        for item in optional_items:
            # Skip section headers
            if item.is_section_header:
                continue
                
            row_cells = optional_table.add_row().cells
            row_cells[0].text = item.description
            
            # Add quantity
            if item.quantity and item.quantity > 0:
                row_cells[1].text = f"{item.quantity:.0f}" if item.quantity == int(item.quantity) else f"{item.quantity:.1f}"
            else:
                row_cells[1].text = "-"
            
            # Add price - handle TBD items
            if item.is_tbd_price:
                row_cells[2].text = "TBD"
            else:
                # Format the price with commas for thousands
                price = item.price_total or 0.0
                row_cells[2].text = f"${price:,.2f}" if price else "-"
"""
            
            new_optional_item_add = """        # Add data rows for optional items
        for item in optional_items:
            # Skip section headers
            if item.is_section_header:
                continue
                
            row_cells = optional_table.add_row().cells
            row_cells[0].text = item.description
            
            # Add Price Ea. column
            if item.is_tbd_price:
                row_cells[1].text = "TBD"
            else:
                # Format the price with commas for thousands
                price_each = item.price_each or 0.0
                row_cells[1].text = f"${price_each:,.2f}" if price_each else "Option"
            
            # Add quantity
            if item.quantity and item.quantity > 0:
                row_cells[2].text = f"{item.quantity:.0f}" if item.quantity == int(item.quantity) else f"{item.quantity:.1f}"
            else:
                row_cells[2].text = "1"  # Default quantity
            
            # Add price - handle TBD items
            if item.is_tbd_price:
                row_cells[3].text = "TBD"
            else:
                # Format the price with commas for thousands
                price = item.price_total or 0.0
                row_cells[3].text = f"${price:,.2f}" if price else "Option"
"""
            
            if old_optional_item_add in content:
                content = content.replace(old_optional_item_add, new_optional_item_add)
                logger.info("Updated optional item addition")
            else:
                logger.warning("Could not find optional item addition pattern to update")
        
        # 3. Create a separate function for handling the page 7 image without cropping
        if "def extract_page_as_image" in content:
            logger.info("Adding special handling for page 7 image")
            
            # Define the new function for special page handling
            special_page_function = """
def extract_page_as_image_special(pdf_path, page_num, output_path, dpi=200):
    \"\"\"
    Extract a page from a PDF as an image with minimal cropping, only removing the header.
    Specifically designed for page 7 where we don't want to crop the sides or bottom.
    
    Args:
        pdf_path: Path to the PDF file
        page_num: Page number to extract (1-indexed)
        output_path: Path to save the extracted image
        dpi: DPI for the extracted image
    
    Returns:
        Path to the extracted image
    \"\"\"
    try:
        # Convert to image using pdf2image
        from pdf2image import convert_from_path
        
        # Extract the specified page
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
        
        # Only crop off the header (top 250px)
        # Do not crop sides or bottom
        width, height = img.size
        cropped_img = img.crop((0, 250, width, height))
        
        # Save the cropped image
        cropped_img.save(output_path)
        logger.info(f"Saved special handling image to {output_path}")
        
        return output_path
    except Exception as e:
        logger.error(f"Error extracting page {page_num} as image: {e}")
        return None"""
            
            # Find the right spot to add the new function
            # Let's try to add it after the existing extract_page_as_image function
            if "def extract_page_as_image(" in content:
                # Find the end of the function
                extract_func_start = content.find("def extract_page_as_image(")
                # Find the next function after extract_page_as_image
                next_func_start = content.find("def ", extract_func_start + 10)
                
                if next_func_start > 0:
                    # Insert our new function before the next function
                    content = content[:next_func_start] + special_page_function + "\n\n" + content[next_func_start:]
                    logger.info("Added special page handling function")
                else:
                    # If there's no next function, add to the end of the file
                    content += "\n\n" + special_page_function
                    logger.info("Added special page handling function at the end of the file")
            else:
                logger.warning("Could not find extract_page_as_image function")
        
            # Now modify the part where we extract pages from the PDF to use our special function for page 7
            if "# Extract specific pages from the PDF as images" in content:
                old_page_extract = """    # Extract specific pages from the PDF as images
    for page_num in range(2, 6):  # Pages 2-5 (1-indexed)
        # Extract the page as an image
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as temp:
            img_path = extract_page_as_image(
                pdf_path, 
                page_num, 
                temp.name,
                dpi=200  # Reduced from 400 to 200 to prevent worker timeouts
            )
            
            if img_path:
                # Add the extracted image to the document
                doc.add_picture(img_path, width=Inches(6.5))
                logger.info(f"Added full page from PDF page {page_num} to Word document")
                
                # Add a page break after each image except the last one
                if page_num < 5:
                    doc.add_page_break()"""
                
                new_page_extract = """    # Extract specific pages from the PDF as images
    for page_num in range(2, 6):  # Pages 2-5 (1-indexed)
        # Extract the page as an image
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as temp:
            # Use special handling for page 5 (which becomes page 7 in the document)
            # This is specifically for the sizing chart
            if page_num == 5:  # Page 5 in PDF becomes page 7 in document
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
                    dpi=200  # Reduced from 400 to 200 to prevent worker timeouts
                )
            
            if img_path:
                # Add the extracted image to the document
                doc.add_picture(img_path, width=Inches(6.5))
                logger.info(f"Added full page from PDF page {page_num} to Word document")
                
                # Add a page break after each image except the last one
                if page_num < 5:
                    doc.add_page_break()"""
                
                if old_page_extract in content:
                    content = content.replace(old_page_extract, new_page_extract)
                    logger.info("Updated page extraction logic for special handling of page 7")
                else:
                    logger.warning("Could not find page extraction logic to update")
        
        # Write the updated content back to the file
        with open(processor_path, 'w') as f:
            f.write(content)
        
        logger.info("Successfully updated pdf_processor.py")
        return True
    
    except Exception as e:
        logger.error(f"Error fixing options header: {e}")
        return False

if __name__ == "__main__":
    fix_options_header_in_pdf_processor()