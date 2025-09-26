"""
This script fixes any remaining issues with "Included" items in the application.
It specifically addresses:
1. Making WM Base Module + Copilot Base Module show as "Included" rather than $9,600
2. Ensuring Freight values show as "INCLUDED" for the relevant items
3. Verifying installation items show correctly
"""
import os
import sys
import logging
from datetime import datetime

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def fix_app_py_included_items():
    """Fix app.py to properly handle 'Included' items"""
    try:
        with open('app.py', 'r') as f:
            content = f.read()
        
        # 1. Make sure the freight item is marked as included
        old_freight = """        # Add freight item - may need to be modified by user
        freight_item = LineItem(
            extraction_id=extraction.id,
            category="freight",
            description=f"Freight to {extraction.location} (1 Truck, Rear Load)",
            price_each=extraction.freight_price or 0.0,
            quantity=1,
            price_total=extraction.freight_price or 0.0,
            is_included=extraction.freight_price is None or extraction.freight_price == 0.0,
            display_order=display_order,
            user_modified=False
        )"""
        
        new_freight = """        # Add freight item - always included in base price
        freight_item = LineItem(
            extraction_id=extraction.id,
            category="freight",
            description=f"Freight to {extraction.location} (1 Truck, Rear Load)",
            price_each=0.0,  # Always free/included
            quantity=1,
            price_total=0.0,  # Always free/included
            is_included=True,  # Always mark as included
            display_order=display_order,
            user_modified=False
        )"""
        
        content = content.replace(old_freight, new_freight)
        
        # 2. Make sure the software item is marked as included
        old_software = """        # Add software item - may need to be modified by user
        software_item = LineItem(
            extraction_id=extraction.id,
            category="software",
            description="WM Base Module + Copilot Base Module",
            price_each=extraction.software_price or 0.0,
            quantity=1,
            price_total=extraction.software_price or 0.0,
            is_included=extraction.software_price is None or extraction.software_price == 0.0,
            display_order=display_order,
            user_modified=False
        )"""
        
        new_software = """        # Add software item - always included in base price
        software_item = LineItem(
            extraction_id=extraction.id,
            category="software",
            description="WM Base Module + Copilot Base Module",
            price_each=0.0,  # Always free/included
            quantity=1,
            price_total=0.0,  # Always free/included 
            is_included=True,  # Always mark as included
            display_order=display_order,
            user_modified=False
        )"""
        
        content = content.replace(old_software, new_software)
        
        # 3. Make sure the installation item is marked as included
        old_install = """        # Installation item
        install_item = LineItem(
            extraction_id=extraction.id,
            category="installation",
            description="Mechanical Installation (includes rentals)",
            price_each=extraction.installation_price or 0.0,
            quantity=1,
            price_total=extraction.installation_price or 0.0,
            is_included=extraction.installation_price is None or extraction.installation_price == 0.0,
            display_order=display_order,
            user_modified=False
        )"""
        
        new_install = """        # Installation item - always included in base price
        install_item = LineItem(
            extraction_id=extraction.id,
            category="installation",
            description="Mechanical Installation (includes rentals)",
            price_each=0.0,  # Always free/included
            quantity=1,
            price_total=0.0,  # Always free/included
            is_included=True,  # Always mark as included
            display_order=display_order,
            user_modified=False
        )"""
        
        content = content.replace(old_install, new_install)
        
        # Save the changes
        with open('app.py', 'w') as f:
            f.write(content)
        
        logger.info("Successfully fixed 'Included' items in app.py")
        return True
    
    except Exception as e:
        logger.error(f"Error fixing 'Included' items: {e}")
        return False

if __name__ == "__main__":
    fix_app_py_included_items()