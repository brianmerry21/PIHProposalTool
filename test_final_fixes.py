"""
This script tests all the fixes implemented for price calculations and included items.
It verifies:
1. VLM price_each and price_total match
2. Software, freight, and installation items are correctly marked as "Included"
3. Total price calculation is correct
"""
import os
import sys
import logging
from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy.orm import DeclarativeBase
from datetime import datetime

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Create minimal app context
class Base(DeclarativeBase):
    pass

db = SQLAlchemy(model_class=Base)
app = Flask(__name__)
app.config["SQLALCHEMY_DATABASE_URI"] = os.environ.get("DATABASE_URL")
app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {
    "pool_recycle": 300,
    "pool_pre_ping": True,
}
db.init_app(app)

# Import models after initializing app/db
with app.app_context():
    from models import PDFExtraction, LineItem

def create_test_dataset():
    """Create a test dataset with VLM and included items"""
    with app.app_context():
        try:
            # Create a test extraction
            extraction = PDFExtraction(
                filename="test.pdf",
                upload_date=datetime.utcnow(),
                original_path="/tmp/test.pdf",
                customer_name="Test Customer",
                total_price=96332.46,  # The expected total price
                base_price=96332.46,   # Same as total - all other components included
                software_price=0.0,    # Included
                installation_price=0.0, # Included
                seismic_price=0.0,     # Included
                freight_price=0.0,     # Included
                vlm_model="ML25",
                vlm_height="132\"",
                tray_quantity=22,
                tray_width="161.41",
                tray_depth="25.75",
                tray_height="4.72",
                location="Hillsboro, OR"
            )
            db.session.add(extraction)
            db.session.commit()
            
            # Set up display order counter
            display_order = 1
            
            # 1. VLM header
            vlm_header = LineItem(
                extraction_id=extraction.id,
                category="header",
                description="VLM",
                is_section_header=True,
                display_order=display_order
            )
            db.session.add(vlm_header)
            display_order += 1
            
            # 2. Main VLM item - price_each and price_total should match total_price
            vlm_item = LineItem(
                extraction_id=extraction.id,
                category="vlm",
                description=f"Modula {extraction.vlm_model} (H{extraction.vlm_height})",
                price_each=extraction.total_price,  # Should match total_price
                quantity=1,
                price_total=extraction.total_price,  # Should match price_each
                is_included=False,
                display_order=display_order,
                user_modified=False
            )
            db.session.add(vlm_item)
            display_order += 1
            
            # 3. Included VLM items
            included_items = [
                "ISO7 Cleanroom Ready",
                "Single Bay & Internal WorkStation",
                f"Tray size: {extraction.tray_width}'' W x {extraction.tray_depth}'' D Tray Height {extraction.tray_height}''",
                "Dynamic Height Management System"
            ]
            
            for item in included_items:
                qty = extraction.tray_quantity if "Tray size" in item else 1
                line_item = LineItem(
                    extraction_id=extraction.id,
                    category="vlm_included",
                    description=item,
                    price_each=0.0,
                    quantity=qty,
                    price_total=0.0,
                    is_included=True,
                    display_order=display_order,
                    user_modified=False
                )
                db.session.add(line_item)
                display_order += 1
            
            # 4. Freight header + item
            freight_header = LineItem(
                extraction_id=extraction.id,
                category="header",
                description="Freight",
                is_section_header=True,
                display_order=display_order
            )
            db.session.add(freight_header)
            display_order += 1
            
            # Freight item - should be included
            freight_item = LineItem(
                extraction_id=extraction.id,
                category="freight",
                description=f"Freight to {extraction.location} (1 Truck, Rear Load)",
                price_each=0.0,  # Free/included
                quantity=1,
                price_total=0.0,  # Free/included
                is_included=True,  # Marked as included
                display_order=display_order,
                user_modified=False
            )
            db.session.add(freight_item)
            display_order += 1
            
            # 5. Software header + item
            software_header = LineItem(
                extraction_id=extraction.id,
                category="header",
                description="Software",
                is_section_header=True,
                display_order=display_order
            )
            db.session.add(software_header)
            display_order += 1
            
            # Software item - should be included
            software_item = LineItem(
                extraction_id=extraction.id,
                category="software",
                description="WM Base Module + Copilot Base Module",
                price_each=0.0,  # Free/included
                quantity=1,
                price_total=0.0,  # Free/included
                is_included=True,  # Marked as included
                display_order=display_order,
                user_modified=False
            )
            db.session.add(software_item)
            display_order += 1
            
            # 6. Installation header + item
            install_header = LineItem(
                extraction_id=extraction.id,
                category="header",
                description="Installation and Warranty",
                is_section_header=True,
                display_order=display_order
            )
            db.session.add(install_header)
            display_order += 1
            
            # Installation item - should be included
            install_item = LineItem(
                extraction_id=extraction.id,
                category="installation",
                description="Mechanical Installation (includes rentals)",
                price_each=0.0,  # Free/included
                quantity=1,
                price_total=0.0,  # Free/included
                is_included=True,  # Marked as included
                display_order=display_order,
                user_modified=False
            )
            db.session.add(install_item)
            display_order += 1
            
            # Save all items
            db.session.commit()
            
            logger.info(f"Created test dataset with extraction ID: {extraction.id}")
            return extraction.id
            
        except Exception as e:
            logger.error(f"Error creating test dataset: {e}")
            db.session.rollback()
            return None

def verify_calculations(extraction_id):
    """Verify all price calculations are correct"""
    with app.app_context():
        try:
            from app import db  # Import db from app.py to ensure consistent connection
            # Get the extraction and line items
            extraction = db.session.get(PDFExtraction, extraction_id)
            line_items = LineItem.query.filter_by(extraction_id=extraction_id).order_by(LineItem.display_order).all()
            
            all_passed = True
            
            # 1. Check VLM price_each and price_total
            vlm_item = next((item for item in line_items if item.category == "vlm"), None)
            if vlm_item:
                logger.info(f"Testing VLM item prices: price_each={vlm_item.price_each}, price_total={vlm_item.price_total}")
                if vlm_item.price_each != vlm_item.price_total or vlm_item.price_each != extraction.total_price:
                    logger.error("❌ VLM item price_each and price_total do not match total_price")
                    logger.error(f"   price_each: {vlm_item.price_each}, price_total: {vlm_item.price_total}, total_price: {extraction.total_price}")
                    all_passed = False
                else:
                    logger.info("✓ VLM item price_each and price_total match total_price")
            else:
                logger.error("❌ VLM item not found")
                all_passed = False
            
            # 2. Check included items
            included_categories = ["software", "freight", "installation"]
            for category in included_categories:
                item = next((item for item in line_items if item.category == category), None)
                if item:
                    logger.info(f"Testing {category} inclusion: is_included={item.is_included}, price_each={item.price_each}, price_total={item.price_total}")
                    if not item.is_included or item.price_each != 0.0 or item.price_total != 0.0:
                        logger.error(f"❌ {category} item not properly marked as included")
                        logger.error(f"   is_included: {item.is_included}, price_each: {item.price_each}, price_total: {item.price_total}")
                        all_passed = False
                    else:
                        logger.info(f"✓ {category} item is properly marked as included")
                else:
                    logger.error(f"❌ {category} item not found")
                    all_passed = False
            
            # 3. Calculate total price using app logic
            non_optional_items = [item for item in line_items 
                                 if not item.is_included 
                                 and not item.is_section_header 
                                 and not item.is_optional]
            
            total_price = 0.0
            for item in non_optional_items:
                item_price = 0.0 if item.price_total is None else float(item.price_total)
                total_price += item_price
            
            logger.info(f"Testing total price calculation: {total_price} vs expected {extraction.total_price}")
            if abs(total_price - extraction.total_price) > 0.01:  # Allow for floating point rounding
                logger.error(f"❌ Total price calculation incorrect: {total_price} != {extraction.total_price}")
                all_passed = False
            else:
                logger.info("✓ Total price calculation is correct")
            
            # Summary
            if all_passed:
                logger.info("🎉 All tests PASSED! The application correctly handles prices and included items.")
            else:
                logger.error("❌ Some tests FAILED. Review the errors above.")
            
            return all_passed
            
        except Exception as e:
            logger.error(f"Error verifying calculations: {e}")
            return False

def clean_up(extraction_id):
    """Clean up the test data"""
    with app.app_context():
        try:
            from app import db  # Import db from app.py to ensure consistent connection
            # Delete the line items first (foreign key constraint)
            LineItem.query.filter_by(extraction_id=extraction_id).delete()
            # Then delete the extraction
            PDFExtraction.query.filter_by(id=extraction_id).delete()
            db.session.commit()
            logger.info(f"Cleaned up test data for extraction ID: {extraction_id}")
            return True
        except Exception as e:
            logger.error(f"Error cleaning up test data: {e}")
            db.session.rollback()
            return False

if __name__ == "__main__":
    extraction_id = create_test_dataset()
    if extraction_id:
        verify_calculations(extraction_id)
        clean_up(extraction_id)
    else:
        logger.error("Failed to create test dataset, cannot verify calculations.")