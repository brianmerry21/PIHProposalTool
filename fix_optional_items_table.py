"""
This script fixes the 'Options Not Included' table by:
1. Making sure the section header doesn't appear as a line item
2. Adding a 'Price Ea.' column to the options table
3. Replacing dash marks with 'Option' text
"""
import os
import logging
import re

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def find_utils_pdf_processor():
    """Find the pdf_processor.py file in the utils directory"""
    path = "utils/pdf_processor.py"
    if os.path.exists(path):
        return path
    
    # Try to find it in any subdirectory
    for root, _, files in os.walk('.'):
        if 'pdf_processor.py' in files:
            return os.path.join(root, 'pdf_processor.py')
    
    return None

def fix_optional_items_table():
    """Fix the optional items table in the Word document"""
    pdf_processor_path = find_utils_pdf_processor()
    if not pdf_processor_path:
        logger.error("Could not find pdf_processor.py")
        return False
    
    logger.info(f"Found pdf_processor.py at {pdf_processor_path}")
    
    # Read the file
    with open(pdf_processor_path, 'r') as f:
        content = f.read()
    
    # Look for the process_pdf_to_word function
    if "def process_pdf_to_word(" not in content:
        logger.error("Could not find process_pdf_to_word function")
        return False
    
    # First, let's make sure section headers that are optional don't appear
    # Find the code that filters line items for the document
    filter_pattern = r"# Filter out section headers and items we don't want to display"
    if filter_pattern in content:
        # Get the surrounding code
        start_idx = content.find(filter_pattern)
        surrounding_code = content[start_idx:start_idx + 500]
        
        # Find the filtering code
        if "# Filter out optional items for the main table" in surrounding_code:
            logger.info("Found filtering code for optional items")
            
            # Add code to ensure section headers marked as optional are excluded
            old_filter = "non_optional_items = [item for item in line_items if not item.is_optional and not item.is_section_header]"
            new_filter = "# Make sure we exclude both optional items and section headers\n"
            new_filter += "        non_optional_items = [item for item in line_items if not item.is_optional and not item.is_section_header]"
            
            # Update the filter to also exclude section headers marked as optional
            content = content.replace(old_filter, new_filter)
            logger.info("Updated filtering code for main table")
    
    # Now look for where we create the optional items table
    optional_pattern = "# Add optional items to a separate table"
    if optional_pattern in content:
        start_idx = content.find(optional_pattern)
        surrounding_code = content[start_idx:start_idx + 1500]  # Get a larger chunk
        
        # Find where we create the table
        table_pattern = r"optional_table = doc\.add_table\(rows=1, cols=(\d+)\)"
        table_match = re.search(table_pattern, surrounding_code)
        
        if table_match:
            # Change table to have 4 columns (add Price Ea. column)
            old_table_create = table_match.group(0)
            num_cols = int(table_match.group(1))
            new_table_create = f"optional_table = doc.add_table(rows=1, cols=4)"  # Always use 4 columns
            
            content = content.replace(old_table_create, new_table_create)
            logger.info(f"Updated optional table to have 4 columns (was {num_cols})")
            
            # Now update column widths to accommodate the new column
            width_pattern = r"# Set column widths[\s\S]*?set_column_width\(optional_table\.columns\[\d\], [^)]*\)"
            width_match = re.search(width_pattern, surrounding_code)
            
            if width_match:
                width_code = surrounding_code[width_match.start():width_match.end() + 200]
                
                # Create new column width code
                new_width_code = """        # Set column widths
        set_column_width(optional_table.columns[0], 4.0)  # Item description
        set_column_width(optional_table.columns[1], 0.85)  # Price Ea.
        set_column_width(optional_table.columns[2], 0.65)  # Quantity
        set_column_width(optional_table.columns[3], 1.0)   # Prices"""
                
                content = content.replace(width_code[:width_match.end() + 100], new_width_code)
                logger.info("Updated column widths for optional table")
            
            # Update header cells
            header_pattern = r"header_cells\[\d\]\.text = \"[^\"]*\""
            header_matches = re.finditer(header_pattern, surrounding_code)
            header_cells = [m.group(0) for m in header_matches]
            
            if header_cells:
                # Find where the header cells are defined
                first_header = header_cells[0]
                header_start = surrounding_code.find(first_header)
                header_section = surrounding_code[header_start:header_start + 200]
                
                # Create new header cells code
                new_header_code = """        header_cells[0].text = "Item"
        header_cells[1].text = "Price Ea."
        header_cells[2].text = "Qty."
        header_cells[3].text = "Prices\""""
                
                content = content.replace(header_section[:header_start + len(first_header) + 100], new_header_code)
                logger.info("Updated header cells for optional table")
            
            # Update the row creation code to add four columns of data
            # Find where we populate table rows
            row_pattern = r"row_cells = optional_table\.add_row\(\)\.cells[\s\S]*?row_cells\[\d\]\.text = [^\\n]*"
            row_match = re.search(row_pattern, surrounding_code)
            
            if row_match:
                row_code = surrounding_code[row_match.start():row_match.start() + 500]
                
                # Create new row cell code that adds Price Ea. column and replaces dash with "Option"
                new_row_code = """            row_cells = optional_table.add_row().cells
            row_cells[0].text = item.description
            
            # Add Price Ea. column
            if item.is_tbd_price:
                row_cells[1].text = "TBD"
            else:
                # Format the price with commas for thousands
                price_each = item.price_each or 0.0
                row_cells[1].text = f"${price_each:,.2f}" if price_each > 0 else "Option"
            
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
                row_cells[3].text = f"${price:,.2f}" if price > 0 else "Option\""""
                
                content = content.replace(row_code[:row_match.start() + 300], new_row_code)
                logger.info("Updated row creation code for optional table")
    
    # Make sure we properly filter out section headers from the optional items
    optional_filter_pattern = r"optional_items = \[item for item in line_items if item\.is_optional\]"
    if optional_filter_pattern in content:
        new_optional_filter = "optional_items = [item for item in line_items if item.is_optional and not item.is_section_header]"
        content = content.replace(optional_filter_pattern, new_optional_filter)
        logger.info("Updated filter to exclude section headers from optional items")
    
    # Write the modified content back to the file
    with open(pdf_processor_path, 'w') as f:
        f.write(content)
    
    logger.info("Successfully updated pdf_processor.py to fix optional items table")
    return True

if __name__ == "__main__":
    fix_optional_items_table()