"""
This script fixes the remaining formatting issues:
1. Remove the "Options Not Included" line from appearing in the options table
2. Add Price Ea. column to optional items table and use "Option" instead of dash marks
3. Fix image sizing for page 7 (page 5 in PDF) to avoid excessive cropping
"""
import os
import sys
import logging
import re

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def modify_pdf_processor():
    """Find and modify the pdf_processor.py file to fix the remaining formatting issues"""
    try:
        pdf_processor_path = "utils/pdf_processor.py"
        if not os.path.exists(pdf_processor_path):
            logger.error(f"Could not find {pdf_processor_path}")
            return False
        
        # Read the file
        with open(pdf_processor_path, 'r') as f:
            content = f.read()
        
        # Look for the extract_page_as_image function to add our specialized version
        image_func_match = re.search(r'def extract_page_as_image\([^)]*\):', content)
        if not image_func_match:
            logger.error("Could not find extract_page_as_image function")
            return False
        
        # Find the end of the function to add our new specialized function
        func_start = image_func_match.start()
        func_end = content.find("def ", func_start + 10)
        if func_end == -1:
            # If no more functions are found, look for other markers
            func_end = content.find("\n\n", func_start + 100)  # Skip ahead a bit to find the end
        
        # Create our specialized image handling function
        special_image_func = """
def extract_page_as_image_special(pdf_path, page_num, output_path, dpi=200):
    '''
    Extract a page from a PDF as an image with minimal cropping (only removing the header).
    Specifically designed for page 7 where we don't want to crop sides or bottom.
    
    Args:
        pdf_path: Path to the PDF file
        page_num: Page number to extract (1-indexed)
        output_path: Path to save the extracted image
        dpi: DPI for the extracted image
    
    Returns:
        Path to the extracted image
    '''
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
        
        # Only crop off the header (top 250px) - no side or bottom cropping
        width, height = img.size
        cropped_img = img.crop((0, 250, width, height))
        
        # Save the cropped image
        cropped_img.save(output_path)
        logger.info(f"Saved special handling image for page {page_num} to {output_path}")
        
        return output_path
    except Exception as e:
        logger.error(f"Error extracting page {page_num} as image with special handling: {e}")
        return None
"""
        
        # Insert the new function after the existing one
        modified_content = content[:func_end] + special_image_func + content[func_end:]
        
        # Now find the code that extracts pages as images and modify it to use our special function for page 5
        page_extract_pattern = r'# Extract specific pages from the PDF as images[\s\S]*?for page_num in range\(2, 6\):'
        page_extract_match = re.search(page_extract_pattern, modified_content)
        
        if not page_extract_match:
            logger.error("Could not find page extraction code")
            return False
        
        # Find the code inside the loop
        loop_start = page_extract_match.end()
        inside_loop = modified_content[loop_start:loop_start + 1000]  # Get a chunk to work with
        
        # Find the line where extract_page_as_image is called
        extract_line_pattern = r'img_path = extract_page_as_image\('
        extract_line_match = re.search(extract_line_pattern, inside_loop)
        
        if not extract_line_match:
            logger.error("Could not find extract_page_as_image call in loop")
            return False
        
        # Replace the extract code with a conditional based on page number
        original_extract_code = inside_loop[extract_line_match.start():extract_line_match.start() + 200]
        # Find where the code block for the extract call ends
        extract_end = original_extract_code.find(")")
        if extract_end == -1:
            logger.error("Could not find end of extract_page_as_image call")
            return False
        
        extract_end += 1  # Include the closing parenthesis
        
        # Create the modified extraction code
        modified_extract_code = """
            # Use special handling for page 5 (which becomes page 7 in the document)
            if page_num == 5:  # Use special handling for sizing chart image
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
"""
        
        # Replace the original extraction code with our conditional version
        modified_content = modified_content.replace(
            original_extract_code[:extract_end], 
            modified_extract_code
        )
        
        # Now find the code that creates the optional items table
        # Search for where we create the optional items table and add "Options Not Included" section
        options_table_pattern = r'# Add optional items to a separate table[\s\S]*?if optional_items:'
        options_table_match = re.search(options_table_pattern, modified_content)
        
        if not options_table_match:
            logger.error("Could not find optional items table code")
            return False
        
        options_start = options_table_match.end()
        options_code = modified_content[options_start:options_start + 1000]
        
        # Find the part where we create the table and set its columns
        table_create_pattern = r'optional_table = doc\.add_table\(rows=1, cols=(\d)\)'
        table_create_match = re.search(table_create_pattern, options_code)
        
        if not table_create_match:
            logger.error("Could not find optional table creation code")
            return False
        
        # Change the table to have 4 columns instead of 3
        original_cols = table_create_match.group(1)
        modified_options_code = options_code.replace(
            f"optional_table = doc.add_table(rows=1, cols={original_cols})",
            "optional_table = doc.add_table(rows=1, cols=4)"  # Update to 4 columns
        )
        
        # Update the column widths
        column_widths_pattern = r'# Set column widths[\s\S]*?set_column_width\(optional_table\.columns\[\d\], [^)]*\)'
        column_widths_match = re.search(column_widths_pattern, modified_options_code)
        
        if not column_widths_match:
            logger.error("Could not find column width setting code")
            return False
        
        column_widths_code = modified_options_code[column_widths_match.start():column_widths_match.end() + 500]
        
        # Find all the set_column_width calls
        width_calls = re.findall(r'set_column_width\(optional_table\.columns\[\d\], [^)]*\)', column_widths_code)
        
        if len(width_calls) < int(original_cols):
            logger.error(f"Expected at least {original_cols} column width settings, found {len(width_calls)}")
            return False
        
        # Replace the column width settings with our 4-column version
        new_column_widths = """        # Set column widths
        set_column_width(optional_table.columns[0], 4.0)  # Item description  
        set_column_width(optional_table.columns[1], 0.85)  # Price Ea.
        set_column_width(optional_table.columns[2], 0.65)  # Quantity
        set_column_width(optional_table.columns[3], 1.0)   # Prices"""
        
        modified_options_code = re.sub(
            r'# Set column widths[\s\S]*?' + re.escape(width_calls[-1]),
            new_column_widths,
            modified_options_code
        )
        
        # Update the header cells
        header_cells_pattern = r'header_cells\[\d\]\.text = "[^"]*"'
        header_cells = re.findall(header_cells_pattern, modified_options_code)
        
        if len(header_cells) < int(original_cols):
            logger.error(f"Expected at least {original_cols} header cell settings, found {len(header_cells)}")
            return False
        
        # Replace the header cell settings with our 4-column version
        new_header_cells = """        header_cells[0].text = "Item"
        header_cells[1].text = "Price Ea."
        header_cells[2].text = "Qty."
        header_cells[3].text = "Prices\""""
        
        modified_options_code = re.sub(
            r'header_cells\[\d\]\.text = "[^"]*"[\s\S]*?' + re.escape(header_cells[-1]),
            new_header_cells,
            modified_options_code
        )
        
        # Now update the part where we add data rows to the optional items table
        data_rows_pattern = r'# Add data rows for optional items[\s\S]*?for item in optional_items:'
        data_rows_match = re.search(data_rows_pattern, modified_options_code)
        
        if not data_rows_match:
            logger.error("Could not find optional items data rows code")
            return False
        
        data_rows_start = data_rows_match.end()
        data_rows_code = modified_options_code[data_rows_start:data_rows_start + 500]
        
        # Find the part where we check for section headers and skip them
        skip_section_pattern = r'# Skip section headers[\s\S]*?if item\.is_section_header:[\s\S]*?continue'
        skip_section_match = re.search(skip_section_pattern, data_rows_code)
        
        if not skip_section_match:
            logger.error("Could not find skip section header code")
            return False
        
        # First add the modified options code to the content
        modified_content = modified_content.replace(options_code, modified_options_code)
        
        # Now find and replace the part where we add the data rows
        row_cells_pattern = r'row_cells = optional_table\.add_row\(\)\.cells[\s\S]*?row_cells\[\d\]\.text = "[^"]*"'
        row_cells_code = re.search(row_cells_pattern, modified_content)
        
        if not row_cells_code:
            logger.error("Could not find row cells code")
            return False
        
        # Find the entire row data setting block
        row_data_start = row_cells_code.start()
        row_data_code = modified_content[row_data_start:row_data_start + 1000]
        
        # Find all the row_cells assignments
        row_cells_assigns = re.findall(r'row_cells\[\d\]\.text = .*?(?=\n)', row_data_code)
        
        if len(row_cells_assigns) < int(original_cols):
            logger.error(f"Expected at least {original_cols} row cell assignments, found {len(row_cells_assigns)}")
            return False
        
        # Create our modified row data code
        new_row_data = """            row_cells = optional_table.add_row().cells
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
        
        # Replace the row data code
        modified_content = modified_content.replace(
            row_data_code[:row_cells_code.end() + 200],  # Include some extra to make sure we get it all
            new_row_data
        )
        
        # Write the modified content back to the file
        with open(pdf_processor_path, 'w') as f:
            f.write(modified_content)
        
        logger.info("Successfully modified pdf_processor.py")
        
        # Now let's find and modify the app.py to filter out section headers from the optional items
        app_path = "app.py"
        if not os.path.exists(app_path):
            logger.error(f"Could not find {app_path}")
            return False
        
        # Read the app.py file
        with open(app_path, 'r') as f:
            app_content = f.read()
        
        # Find where we filter items for the review page
        filter_pattern = r'# Calculate total price \(excluding optional items and section headers\)'
        filter_match = re.search(filter_pattern, app_content)
        
        if not filter_match:
            logger.error("Could not find filtering code in app.py")
            return False
        
        filter_start = filter_match.end()
        filter_code = app_content[filter_start:filter_start + 500]
        
        # Find the non_optional_items filter
        non_optional_pattern = r'non_optional_items = \[.*?\]'
        non_optional_match = re.search(non_optional_pattern, filter_code)
        
        if not non_optional_match:
            logger.error("Could not find non_optional_items filter")
            return False
        
        # Update the filter to also exclude section headers that are marked as optional
        modified_app_content = app_content.replace(
            "non_optional_items = [item for item in line_items if not item.is_included and not item.is_section_header and not item.is_optional]",
            "non_optional_items = [item for item in line_items if not item.is_included and not item.is_section_header and not item.is_optional]"
        )
        
        # Also find where we're extracting optional items for the document
        optional_filter_pattern = r'optional_items = \[.*?is_optional.*?\]'
        optional_filter_match = re.search(optional_filter_pattern, app_content)
        
        if optional_filter_match:
            # Update this filter to exclude section headers from optional items
            modified_app_content = modified_app_content.replace(
                "optional_items = [item for item in line_items if item.is_optional]",
                "optional_items = [item for item in line_items if item.is_optional and not item.is_section_header]"
            )
        
        # Write the modified app content back
        with open(app_path, 'w') as f:
            f.write(modified_app_content)
        
        logger.info("Successfully modified app.py")
        
        return True
    
    except Exception as e:
        logger.error(f"Error modifying files: {e}")
        return False

if __name__ == "__main__":
    modify_pdf_processor()