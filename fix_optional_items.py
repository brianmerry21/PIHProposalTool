import os
import logging

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def fix_extract_preview():
    """
    Apply the necessary fixes to app.py to prevent duplicate optional items
    and fix the formatting in the Options Not Included section.
    """
    try:
        # Path to app.py
        app_path = 'app.py'
        
        # Read the file
        with open(app_path, 'r') as f:
            content = f.read()
        
        # Add imports at the top if needed
        if 'from utils.pdf_processor import extract_optional_items' not in content:
            content = content.replace(
                'from utils.pdf_processor import extract_line_items, extract_total_value',
                'from utils.pdf_processor import extract_line_items, extract_total_value, extract_optional_items'
            )
        
        # Implement the fix to filter out optional items from main items
        if 'optional_items_data = extract_optional_items(full_text)' not in content:
            # Find the line where we extract line items
            extract_line_items_pos = content.find('items_data = extract_line_items(full_text)')
            if extract_line_items_pos > 0:
                # Get the indentation level
                indentation = ''
                for i in range(extract_line_items_pos - 1, 0, -1):
                    if content[i] == '\n':
                        break
                    indentation = content[i] + indentation
                
                # Create the code to insert
                optional_items_code = f"""
{indentation}# Get optional items first to make sure we don't include them in regular items
{indentation}optional_items_data = extract_optional_items(full_text)

{indentation}# Create a list of optional item descriptions to filter them out of regular items
{indentation}optional_descriptions = [item.get('description', '').lower() for item in optional_items_data]

{indentation}# Extract all line items"""
                
                # Insert code before the line items extraction
                content = content[:extract_line_items_pos] + optional_items_code + content[extract_line_items_pos:]
                
                # Add filter after items_data = extract_line_items(full_text)
                line_items_end_pos = content.find('\n', extract_line_items_pos)
                filter_code = f"""

{indentation}# Filter out any items that are in the optional items list to avoid duplication
{indentation}filtered_items_data = []
{indentation}for item in items_data:
{indentation}    item_description = item.get('description', '').lower()
{indentation}    # Skip if this item is in the optional items list
{indentation}    if any(opt_desc in item_description or item_description in opt_desc for opt_desc in optional_descriptions):
{indentation}        logger.info(f"Filtering out optional item from main items: {{item_description}}")
{indentation}        continue
{indentation}    filtered_items_data.append(item)"""
                
                # Add filter code
                content = content[:line_items_end_pos] + filter_code + content[line_items_end_pos:]
                
                # Replace items_data with filtered_items_data throughout the code
                loop_through_items_pos = content.find('# Loop through all line items and add them to the database')
                if loop_through_items_pos > 0:
                    loop_code = content[loop_through_items_pos:]
                    # Replace with filtered items
                    loop_code = loop_code.replace('for item_data in items_data:', 'for item_data in filtered_items_data:')
                    content = content[:loop_through_items_pos] + loop_code
        
        # Make sure is_recommended is set to False in the code
        content = content.replace("is_recommended=item_data.get('is_recommended', False)", "is_recommended=False")
        
        # Write the file back
        with open(app_path, 'w') as f:
            f.write(content)
            
        logger.info("Successfully applied fixes to app.py")
        
    except Exception as e:
        logger.error(f"Error applying fixes: {str(e)}")
        raise

if __name__ == "__main__":
    fix_extract_preview()
    print("Fixes applied successfully.")