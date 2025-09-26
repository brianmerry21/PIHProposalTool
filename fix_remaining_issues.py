"""
This script fixes two remaining issues:
1. "Options Not Included" section header appearing in main pricing table
2. Total price in the document incorrectly including optional items

The fix:
1. Mark section headers for optional items with is_optional=True
2. Update the total price calculation to exclude optional items
"""

import logging
import os

def fix_optional_section_header():
    """
    Apply necessary changes to prevent "Options Not Included" from appearing in main table
    and fix total price calculation to exclude optional items
    """
    app_file = 'app.py'
    temp_file = 'app.py.temp'
    
    try:
        # Process line by line
        with open(app_file, 'r') as original, open(temp_file, 'w') as temp:
            processing_first_section = False
            processing_second_section = False
            
            for line in original:
                # Look for the section header creation lines
                if '# 7. Optional Items section' in line:
                    temp.write(line)
                    processing_first_section = True
                    continue
                
                # Mark the section header as optional
                if processing_first_section and 'is_section_header=True' in line:
                    # Add is_optional=True before the closing parenthesis
                    modified_line = line.rstrip()
                    if modified_line.endswith(','):
                        temp.write(modified_line + '\n')
                    else:
                        # Add comma if needed
                        if modified_line.endswith(')'):
                            modified_line = modified_line[:-1] + ',\n'
                            temp.write(modified_line)
                            temp.write('            is_optional=True  # Mark as optional to filter from main table\n')
                            temp.write('        )\n')
                            processing_first_section = False
                            continue
                        else:
                            temp.write(modified_line + ',\n')
                elif processing_first_section and line.strip() == ')':
                    temp.write('            is_optional=True  # Mark as optional to filter from main table\n')
                    temp.write(line)
                    processing_first_section = False
                    continue
                
                # Look for the total price calculation in process_extraction
                if 'Calculate total price' in line:
                    temp.write(line)
                    processing_second_section = True
                    continue
                
                # Update the total price calculation
                if processing_second_section and 'total_price = sum' in line:
                    # Replace with the corrected version
                    temp.write('        total_price = sum(item.price_total for item in line_items \n')
                    temp.write('                     if not item.is_included \n')
                    temp.write('                     and not item.is_section_header \n')
                    temp.write('                     and not item.is_optional)\n')
                    processing_second_section = False
                    continue
                
                # Write unmodified line
                temp.write(line)
        
        # Replace the original file
        os.replace(temp_file, app_file)
        print("Fixed optional section headers and total price calculation in app.py")
        return True
        
    except Exception as e:
        logging.error(f"Error fixing optional section headers: {str(e)}")
        # Clean up temp file if it exists
        if os.path.exists(temp_file):
            os.remove(temp_file)
        return False

def fix_pdf_processor():
    """
    Fix the document generation to properly filter optional items from
    the main pricing table in utils/pdf_processor.py
    """
    file_path = 'utils/pdf_processor.py'
    temp_file = 'utils/pdf_processor.py.temp'
    
    try:
        with open(file_path, 'r') as original, open(temp_file, 'w') as temp:
            for line in original:
                if "# Skip items that are marked as not to be included" in line:
                    temp.write(line)
                    # Add the corrected filtering code
                    temp.write("            # Skip items that are marked as not to be included or are optional (separate section)\n")
                    temp.write("            if hasattr(item, 'get') and (\n")
                    temp.write("                ('user_selected' in item and not item.get('user_selected', True)) or\n")
                    temp.write("                item.get('is_optional', False)  # Skip optional items in main pricing table\n")
                    temp.write("            ):\n")
                    temp.write("                continue\n")
                    
                    # Skip the next few lines (the original if condition)
                    for _ in range(4):  # Skip 4 lines
                        next(original, '')
                    continue
                
                # Write unmodified line
                temp.write(line)
        
        # Replace the original file
        os.replace(temp_file, file_path)
        print("Fixed document generation in utils/pdf_processor.py")
        return True
        
    except Exception as e:
        logging.error(f"Error fixing pdf_processor.py: {str(e)}")
        # Clean up temp file if it exists
        if os.path.exists(temp_file):
            os.remove(temp_file)
        return False

if __name__ == "__main__":
    success1 = fix_optional_section_header()
    success2 = fix_pdf_processor()
    
    if success1 and success2:
        print("All fixes applied successfully.")
    else:
        print("Some fixes could not be applied.")