"""
This script fixes the headers in the Options Not Included table:
1. Renames the headers to: Item, Price Ea., Qty., Option
2. Makes sure the word "Option" appears in all cells in the last column
"""
import os
import logging
import re

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def fix_options_headers():
    """Fix the headers in the Options Not Included table"""
    
    # Find the pdf_processor file
    pdf_processor_path = "utils/pdf_processor.py"
    if not os.path.exists(pdf_processor_path):
        # Try to find it in other locations
        for root, _, files in os.walk("."):
            if "pdf_processor.py" in files:
                pdf_processor_path = os.path.join(root, "pdf_processor.py")
                break
        
        if not os.path.exists(pdf_processor_path):
            logger.error("Could not find pdf_processor.py")
            return False
    
    logger.info(f"Found pdf_processor.py at {pdf_processor_path}")
    
    # Read the file
    with open(pdf_processor_path, 'r') as f:
        content = f.read()
    
    # Fix the header cells in the Options Not Included table
    # Find the part where header cells are defined for the optional items table
    header_pattern = r"(\s+)header_cells\[0\]\.text = \".*?\"\s+header_cells\[1\]\.text = \".*?\"\s+header_cells\[2\]\.text = \".*?\"\s+header_cells\[3\]\.text = \".*?\""
    header_match = re.search(header_pattern, content)
    
    if not header_match:
        logger.error("Could not find header cells code")
        return False
    
    # Create the new header code with correct column names
    indent = header_match.group(1)
    new_header_code = f"{indent}header_cells[0].text = \"Item\"\n"
    new_header_code += f"{indent}header_cells[1].text = \"Price Ea.\"\n"
    new_header_code += f"{indent}header_cells[2].text = \"Qty.\"\n"
    new_header_code += f"{indent}header_cells[3].text = \"Option\""
    
    # Replace the header cells code
    content = content.replace(header_match.group(0), new_header_code)
    logger.info("Updated header cells for optional items table")
    
    # Now find and completely replace the code for populating the optional items rows
    # We need a comprehensive fix for the row data in the options table
    options_row_pattern = r"# Add data rows for optional items.*?for item in optional_items:(.*?)(?=\n\s*# Add section headings|\n\s*# Add fixed header|\n\s*doc\.add_paragraph)"
    options_row_match = re.search(options_row_pattern, content, re.DOTALL)
    
    if not options_row_match:
        logger.error("Could not find optional items row code")
        return False
    
    # Create a completely new row code block that ensures the last column shows Option
    # Indent level will match what was there before
    row_code_block = options_row_match.group(1)
    
    # Extract the indent level from the first line of code in the row block
    indent_match = re.search(r"^(\s+)", row_code_block.lstrip("\n"))
    if not indent_match:
        indent = "            "  # Default indent if we can't extract it
    else:
        indent = indent_match.group(1)
    
    new_row_code = f"\n{indent}# Skip section headers\n"
    new_row_code += f"{indent}if item.is_section_header:\n"
    new_row_code += f"{indent}    continue\n\n"
    new_row_code += f"{indent}row_cells = optional_table.add_row().cells\n"
    new_row_code += f"{indent}row_cells[0].text = item.description\n\n"
    new_row_code += f"{indent}# Add Price Ea. column\n"
    new_row_code += f"{indent}if item.is_tbd_price:\n"
    new_row_code += f"{indent}    row_cells[1].text = \"TBD\"\n"
    new_row_code += f"{indent}else:\n"
    new_row_code += f"{indent}    # Format the price with commas for thousands\n"
    new_row_code += f"{indent}    price_each = item.price_each or 0.0\n"
    new_row_code += f"{indent}    row_cells[1].text = f\"${{price_each:,.2f}}\" if price_each > 0 else \"TBD\"\n\n"
    new_row_code += f"{indent}# Add quantity\n"
    new_row_code += f"{indent}if item.quantity and item.quantity > 0:\n"
    new_row_code += f"{indent}    row_cells[2].text = f\"{{item.quantity:.0f}}\" if item.quantity == int(item.quantity) else f\"{{item.quantity:.1f}}\"\n"
    new_row_code += f"{indent}else:\n"
    new_row_code += f"{indent}    row_cells[2].text = \"1\"  # Default quantity\n\n"
    new_row_code += f"{indent}# Last column always shows \"Option\"\n"
    new_row_code += f"{indent}row_cells[3].text = \"Option\"\n"
    
    # Replace the old row code with our new version
    new_options_section = content.replace(options_row_match.group(1), new_row_code)
    
    if new_options_section != content:
        content = new_options_section
        logger.info("Updated optional items row code to ensure last column always shows 'Option'")
    
    # Write the changes back to the file
    with open(pdf_processor_path, 'w') as f:
        f.write(content)
    
    logger.info("Successfully updated options headers in pdf_processor.py")
    return True

if __name__ == "__main__":
    fix_options_headers()