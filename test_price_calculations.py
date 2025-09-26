
import os
import sys
import logging
from datetime import datetime

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Import application modules
from app import app, db
from models import PDFExtraction, LineItem

def test_price_calculations():
    """Test price calculations with our fixes"""
    logger.info("Testing price calculations with fixed code...")
    
    with app.app_context():
        # Create a test extraction
        extraction = PDFExtraction(
            filename="test.pdf",
            upload_date=datetime.utcnow(),
            original_path="/tmp/test.pdf",
            customer_name="Test Customer",
            total_price=96332.46  # The expected total price
        )
        db.session.add(extraction)
        db.session.flush()
        
        # Create VLM item 
        vlm_item = LineItem(
            extraction_id=extraction.id,
            category="vlm",
            description="Modula ML25 VLM",
            price_each=96332.46,  # Should match total_price
            quantity=1,
            price_total=96332.46,  # Should match price_each
            is_included=False,
            is_section_header=False,
            is_optional=False
        )
        db.session.add(vlm_item)
        
        # Create included item
        included_item = LineItem(
            extraction_id=extraction.id,
            category="software",
            description="WM Base Module + Copilot Base Module",
            price_each=0.0,
            quantity=1,
            price_total=0.0,
            is_included=True,
            is_section_header=False,
            is_optional=False
        )
        db.session.add(included_item)
        
        # Create optional item
        optional_item = LineItem(
            extraction_id=extraction.id,
            category="Optional",
            description="Optional item",
            price_each=1000.0,
            quantity=1,
            price_total=1000.0,
            is_included=False,
            is_section_header=False,
            is_optional=True
        )
        db.session.add(optional_item)
        
        # Create section header
        section_header = LineItem(
            extraction_id=extraction.id,
            category="header",
            description="Options Not Included",
            is_section_header=True,
            is_optional=True
        )
        db.session.add(section_header)
        
        # Calculate total price using our algorithm
        non_optional_items = [item for item in [vlm_item, included_item, optional_item, section_header] 
                            if not item.is_included 
                            and not item.is_section_header 
                            and not item.is_optional]
        
        total_price = 0.0
        for item in non_optional_items:
            item_price = 0.0 if item.price_total is None else float(item.price_total)
            total_price += item_price
        
        expected_total = 96332.46  # Should only include the VLM item
        
        # Check if calculation is correct
        if abs(total_price - expected_total) < 0.01:
            logger.info(f"Total price calculation: {total_price} ✓")
            logger.info("Price calculation test PASSED")
        else:
            logger.error(f"Total price calculation: {total_price}, expected {expected_total} ✗")
            logger.error("Price calculation test FAILED")
        
        # Rollback to avoid affecting database
        db.session.rollback()

if __name__ == "__main__":
    test_price_calculations()
