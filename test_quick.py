"""
Simple test to check if the VLM item shows correct price values and included items are marked properly.
This test directly imports from app.py to avoid SQLAlchemy context issues.
"""
import logging
from app import app, db
from models import PDFExtraction, LineItem

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def test_prices():
    """Test if prices are set correctly in the application"""
    with app.app_context():
        # Reset database first (optional)
        # db.drop_all()
        # db.create_all()
        
        # Check if any extractions exist
        extractions = PDFExtraction.query.all()
        if extractions:
            logger.info(f"Found {len(extractions)} existing extractions")
            
            # Check the most recent extraction
            extraction = extractions[-1]
            logger.info(f"Checking extraction ID: {extraction.id}")
            logger.info(f"Total price: {extraction.total_price}")
            
            # Get line items
            line_items = LineItem.query.filter_by(extraction_id=extraction.id).order_by(LineItem.display_order).all()
            
            # Check VLM item
            vlm_item = next((item for item in line_items if item.category == "vlm"), None)
            if vlm_item:
                logger.info(f"VLM Item: {vlm_item.description}")
                logger.info(f"  Price Ea: {vlm_item.price_each}")
                logger.info(f"  Quantity: {vlm_item.quantity}")
                logger.info(f"  Price Total: {vlm_item.price_total}")
                
                # Check if price_each and price_total match
                if vlm_item.price_each == vlm_item.price_total and vlm_item.price_total == extraction.total_price:
                    logger.info("✓ VLM item prices are consistent")
                else:
                    logger.error("❌ VLM item prices are inconsistent")
            else:
                logger.error("No VLM item found")
            
            # Check if software, freight, and installation items are marked as included
            for category in ["software", "freight", "installation"]:
                item = next((item for item in line_items if item.category == category), None)
                if item:
                    logger.info(f"{category.capitalize()} Item: {item.description}")
                    logger.info(f"  is_included: {item.is_included}")
                    logger.info(f"  Price Ea: {item.price_each}")
                    logger.info(f"  Price Total: {item.price_total}")
                    
                    if item.is_included and item.price_each == 0.0 and item.price_total == 0.0:
                        logger.info(f"✓ {category.capitalize()} item is correctly marked as included")
                    else:
                        logger.error(f"❌ {category.capitalize()} item is not correctly marked as included")
                else:
                    logger.error(f"No {category} item found")
            
            # Check total price calculation
            non_optional_items = [item for item in line_items 
                                if not item.is_included 
                                and not item.is_section_header 
                                and not item.is_optional]
            
            total_price = 0.0
            for item in non_optional_items:
                item_price = 0.0 if item.price_total is None else float(item.price_total)
                total_price += item_price
            
            logger.info(f"Calculated total price: {total_price}")
            logger.info(f"Expected total price: {extraction.total_price}")
            
            if abs(total_price - extraction.total_price) < 0.01:  # Allow for floating point rounding
                logger.info("✓ Total price calculation is correct")
            else:
                logger.error(f"❌ Total price calculation is incorrect: {total_price} != {extraction.total_price}")
                
        else:
            logger.warning("No extractions found in database. Upload a PDF first to test.")

if __name__ == "__main__":
    test_prices()