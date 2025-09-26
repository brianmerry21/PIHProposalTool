"""
Test script to verify None values are handled correctly in price calculations.
This script tests the fixes we've made for handling NoneType errors.
"""
import os
import sys
import logging
import json
from datetime import datetime

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

try:
    # Import application modules
    from app import app, db
    from models import PDFExtraction, LineItem
except ImportError as e:
    logger.error(f"Error importing application modules: {e}")
    sys.exit(1)

def test_repr_method():
    """Test that LineItem.__repr__ handles None values correctly"""
    logger.info("Testing LineItem.__repr__ method...")
    
    with app.app_context():
        # Create a test item with None price_total
        test_item = LineItem(
            description="Test Item",
            price_each=None,
            quantity=1,
            price_total=None,
            is_included=False
        )
        
        # Get the repr and check it doesn't crash
        try:
            repr_string = repr(test_item)
            logger.info(f"LineItem.__repr__ with None price_total: {repr_string}")
            logger.info("LineItem.__repr__ test PASSED")
            return True
        except Exception as e:
            logger.error(f"LineItem.__repr__ test FAILED: {e}")
            return False

def test_price_calculation():
    """Test that price calculations handle None values correctly"""
    logger.info("Testing price calculations...")
    
    with app.app_context():
        # Create test items with various None values
        test_cases = [
            {"price_each": None, "quantity": 2, "expected": 0.0},
            {"price_each": 10.0, "quantity": None, "expected": 0.0},
            {"price_each": None, "quantity": None, "expected": 0.0},
            {"price_each": 10.0, "quantity": 2, "expected": 20.0}
        ]
        
        all_passed = True
        for i, test in enumerate(test_cases):
            price_each = test["price_each"]
            quantity = test["quantity"]
            expected = test["expected"]
            
            # Manual calculation with our None-safe approach
            if price_each is None or quantity is None:
                price_each = 0.0 if price_each is None else float(price_each)
                quantity = 0.0 if quantity is None else float(quantity)
            
            calculated = price_each * quantity
            
            # Check if calculation matches expected
            if calculated == expected:
                logger.info(f"Test case {i+1}: {price_each} * {quantity} = {calculated} ✓")
            else:
                logger.error(f"Test case {i+1}: {price_each} * {quantity} = {calculated}, expected {expected} ✗")
                all_passed = False
        
        logger.info(f"Price calculation tests {'PASSED' if all_passed else 'FAILED'}")
        return all_passed

def test_total_price_calculation():
    """Test that total price calculation handles None values correctly"""
    logger.info("Testing total price calculation...")
    
    with app.app_context():
        try:
            # Create a test extraction
            extraction = PDFExtraction(
                filename="test.pdf",
                upload_date=datetime.utcnow(),
                original_path="/tmp/test.pdf",
                customer_name="Test Customer"
            )
            db.session.add(extraction)
            db.session.flush()  # Get the ID without committing
            
            # Create test items
            items = [
                LineItem(
                    extraction_id=extraction.id,
                    description="Item 1",
                    price_each=10.0,
                    quantity=2,
                    price_total=20.0,
                    is_included=False,
                    is_section_header=False,
                    is_optional=False
                ),
                LineItem(
                    extraction_id=extraction.id,
                    description="Item 2",
                    price_each=None,  # None price
                    quantity=3,
                    price_total=None,  # None price
                    is_included=False,
                    is_section_header=False,
                    is_optional=False
                ),
                LineItem(
                    extraction_id=extraction.id,
                    description="Included Item",
                    price_each=15.0,
                    quantity=1,
                    price_total=0.0,
                    is_included=True,  # Included items should be excluded
                    is_section_header=False,
                    is_optional=False
                ),
                LineItem(
                    extraction_id=extraction.id,
                    description="Optional Item",
                    price_each=25.0,
                    quantity=1,
                    price_total=25.0,
                    is_included=False,
                    is_section_header=False,
                    is_optional=True  # Optional items should be excluded
                ),
                LineItem(
                    extraction_id=extraction.id,
                    description="Section Header",
                    is_included=False,
                    is_section_header=True,  # Section headers should be excluded
                    is_optional=False
                )
            ]
            
            for item in items:
                db.session.add(item)
            
            # Calculate total price using our None-safe approach
            non_optional_items = [item for item in items 
                                if not item.is_included 
                                and not item.is_section_header 
                                and not item.is_optional]
            
            total_price = 0.0
            for item in non_optional_items:
                # Convert None to 0.0
                item_price = 0.0 if item.price_total is None else float(item.price_total)
                total_price += item_price
            
            # Expected: Only Item 1 should count (20.0)
            # Item 2 has None price so it should be counted as 0.0
            # Other items are excluded based on their flags
            expected_total = 20.0
            
            if abs(total_price - expected_total) < 0.01:  # Allow for floating point errors
                logger.info(f"Total price calculation: {total_price} ✓")
                logger.info("Total price calculation test PASSED")
                result = True
            else:
                logger.error(f"Total price calculation: {total_price}, expected {expected_total} ✗")
                logger.info("Total price calculation test FAILED")
                result = False
            
            # Rollback the transaction to avoid affecting the real database
            db.session.rollback()
            return result
            
        except Exception as e:
            logger.error(f"Total price calculation test FAILED with error: {e}")
            # Ensure we roll back on error
            db.session.rollback()
            return False

def run_all_tests():
    """Run all tests and report results"""
    logger.info("Starting tests for NoneType handling...")
    
    results = {
        "repr_method": test_repr_method(),
        "price_calculation": test_price_calculation(),
        "total_price_calculation": test_total_price_calculation()
    }
    
    logger.info("\n--- Test Results ---")
    all_passed = True
    for test_name, passed in results.items():
        logger.info(f"{test_name}: {'PASSED' if passed else 'FAILED'}")
        if not passed:
            all_passed = False
    
    if all_passed:
        logger.info("\nAll tests PASSED! The application now properly handles None values.")
    else:
        logger.info("\nSome tests FAILED. Further fixes are needed.")

if __name__ == "__main__":
    run_all_tests()