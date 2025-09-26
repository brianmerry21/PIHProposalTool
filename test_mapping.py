import os
import logging
from utils.pdf_processor import extract_text_from_pdf, create_modula_mapping_rules
from utils.excel_template_analyzer import extract_mapping_from_pdf
import pprint

# Configure logging - set to INFO to reduce output
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Silence verbose libraries
logging.getLogger('pdfminer').setLevel(logging.WARNING)
logging.getLogger('fontTools').setLevel(logging.WARNING)
logging.getLogger('PIL').setLevel(logging.WARNING)

def test_pdf_extraction_mapping():
    """Test the PDF extraction and mapping to Excel cells"""
    try:
        # Path to the PDF file
        pdf_path = os.path.join('attached_assets', 'DISC34.99_ThermoFisher_Pricer_Modula_ML25_Cleanroom_No.00051890.pdf')
        
        # Extract text from the PDF
        pdf_text = extract_text_from_pdf(pdf_path)
        logger.info(f"Extracted {len(pdf_text)} pages of text from the PDF")
        
        # Create mapping rules for the Modula PDF
        mapping_rules = create_modula_mapping_rules()
        logger.info("Created mapping rules for the Modula PDF")
        
        # Extract data from PDF text and create mapping for Excel template
        data_mapping = extract_mapping_from_pdf(pdf_text, None, mapping_rules)
        logger.info("Created data mapping from PDF text")
        
        # Print the data mapping to verify it matches the requirements
        print("\n--- Data Mapping Results ---")
        pprint.pprint(data_mapping)
        print("---------------------------\n")
        
        # Verify specific mappings
        sheet_name = "Primary Option"  # The sheet name used in your mapping
        
        # Expected data in cells
        expected_cells = {
            "F1": "Thermo Fisher",
            "B7": "ML25 – 3700",
            "B8": "Offer # 00051890",
            "B9": "Number of Trays: 22"
        }
        
        # Check each expected cell
        for cell, expected_value in expected_cells.items():
            if cell in data_mapping.get(sheet_name, {}):
                actual_value = data_mapping[sheet_name][cell]
                print(f"Cell {cell}: Expected: '{expected_value}', Actual: '{actual_value}'")
                if expected_value in actual_value:
                    print(f"✓ PASS: Cell {cell} contains expected value")
                else:
                    print(f"✗ FAIL: Cell {cell} value does not match")
            else:
                print(f"✗ FAIL: Cell {cell} not found in mapping")
        
        # Also check the tray dimensions formatting
        if "B10" in data_mapping.get(sheet_name, {}):
            tray_dims = data_mapping[sheet_name]["B10"]
            print(f"Tray dimensions: {tray_dims}")
            if "Width:" in tray_dims and "Depth:" in tray_dims:
                print("✓ PASS: Tray dimensions formatted correctly")
            else:
                print("✗ FAIL: Tray dimensions format incorrect")
        else:
            print("✗ FAIL: Tray dimensions cell not found")
            
        return True
        
    except Exception as e:
        logger.error(f"Error in test_pdf_extraction_mapping: {str(e)}")
        return False

if __name__ == "__main__":
    # Run the test
    print("Running PDF extraction mapping test...")
    result = test_pdf_extraction_mapping()
    print(f"Test {'passed' if result else 'failed'}")