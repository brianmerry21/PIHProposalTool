"""
This script adds a new page 14 to the generated Word document 
with proposal details, what's included/excluded, payment schedule, 
and acceptance signature lines.
"""
import os
import re
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def add_proposal_details_page():
    """
    Add a new page 14 to the Word document with proposal details, including:
    - What's included in the proposal
    - What's not included in the proposal
    - Lead time
    - Payment schedule
    - Notes about permits and seismic calculations
    - Signature block
    """
    pdf_processor_path = "utils/pdf_processor.py"
    
    if not os.path.exists(pdf_processor_path):
        logger.error(f"Could not find {pdf_processor_path}")
        return False
        
    logger.info(f"Reading {pdf_processor_path}...")
    
    with open(pdf_processor_path, 'r') as f:
        content = f.read()
        
    # Store original content for comparison later
    original_content = content
    
    # Find the appropriate location to add the new page
    # Look for the code that adds contact information (near the end of the document)
    contact_info_pattern = r"(        # Add contact information\n        contact_para = doc\.add_paragraph\(\)\n        contact_para\.add_run\(\"\\nFor questions or to place an order, please contact:\"\)\.italic = True\n        contact_para\.add_run\(f\"\\n{contact_name}\\nPacific Integrated Handling\\nCell: {contact_phone}\\n{contact_email}\"\))"
    
    # The new code to add the proposal details page
    proposal_details_code = r"""\1
        
        # Add a page break to start the Proposal Details page
        doc.add_page_break()
        
        # Add page 14 header with page number
        section = doc.sections[-1]
        header = section.header
        header_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        header_para.style = 'Header'
        header_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        tab_stops = header_para.paragraph_format.tab_stops
        tab_stops.add_tab_stop(Inches(1.0), WD_TAB_ALIGNMENT.LEFT)
        header_para.text = ""
        header_para.add_run(f"{customer_name}\t").bold = True
        current_date = datetime.now().strftime('%B %d, %Y')
        header_para.add_run(f"Page 14\\n{current_date}")
        
        # Add "Proposal includes:" heading
        doc.add_heading('Proposal includes:', level=1)
        doc.add_paragraph()
        
        # Create a bulleted list of included items
        included_items = [
            "All mechanical hardware as outlined in this proposal",
            f"Freight to {{{{CUSTOMER_LOCATION}}}} (2 Calistoga Trucks, Side Load)",
            "Unloading the equipment upon arrival",
            "Moving equipment from dock area to installation area",
            "Mechanical installation using our own forklift and scissor lifts",
            "Electrical installation and controls within the equipment",
            "Configuring controls and checking the equipment for proper operation",
            "Operator and maintenance training",
            "Manuals and Documentation",
            "PIH Exclusive: 3 Scheduled Maintenance Visits over the first 2 years (Modula requires 1)",
            "2 Year Warranty Parts and Labor",
            "Seismic Anchoring Calculations and Certifications (Price to be determined)"
        ]
        
        # Try to extract location from customer_info or extraction
        freight_location = "Tualatin, OR"
        try:
            if hasattr(customer_info, 'get') and customer_info.get('location'):
                freight_location = customer_info.get('location')
            
            # Replace the placeholder with actual location if found
            included_items[1] = included_items[1].replace("{{CUSTOMER_LOCATION}}", freight_location)
        except:
            # Keep the placeholder if location can't be determined
            pass
            
        for item in included_items:
            p = doc.add_paragraph(style='List Bullet')
            p.add_run(item)
            
        doc.add_paragraph()
        
        # Add "Proposal does not include:" heading
        doc.add_heading('Proposal does not include:', level=1)
        doc.add_paragraph()
        
        # Create a bulleted list of excluded items
        excluded_items = [
            "Electrical Hookup to VLM provided by Western Precision",
            "Fire Suppression",
            "Installation Permits",
            "Providing a clear path for moving equipment to the installation area",
            "Providing a clear area for installation and the erected equipment",
            "Computer Hardware and data drops",
            "Local and State Taxes on applicable items",
            "Seismic reports may require modifications to the Western Precision facility which will be the customer's responsibility"
        ]
        
        for item in excluded_items:
            p = doc.add_paragraph(style='List Bullet')
            p.add_run(item)
            
        doc.add_paragraph()
        
        # Add Lead Time section
        lead_time_heading = doc.add_paragraph()
        lead_time_heading.add_run("Lead Time").bold = True
        lead_time_heading.add_run(": 10-12 Weeks")
        
        doc.add_paragraph()
        
        # Add Payment Schedule section
        payment_heading = doc.add_paragraph()
        payment_heading.add_run("Payment Schedule:").bold = True
        
        payment_schedule = [
            "35% due upon purchase order, Net 10 days",
            "55% due upon shipping, Net 30 days",
            "10% due upon acceptance, Net 30 days"
        ]
        
        for item in payment_schedule:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run(item)
            
        doc.add_paragraph()
        
        # Add notes
        note1_para = doc.add_paragraph()
        note1_para.add_run("Note 1.").bold = True
        note1_para.add_run("\\tAll Permits are "By Customer". Please note that processing time can take 3 to 5 weeks or more.")
        
        note2_para = doc.add_paragraph()
        note2_para.add_run("Note 2.").bold = True
        note2_para.add_run("\\tSeismic calculations for the equipment are required to obtain installation permits. Seismic calculations will be prepared by an Architect selected by Pacific Integrated Handling. The purpose of the calculations is to show that the floor can support the installed storage equipment. The Architect will prepare final project calculations and drawings based on customer supplied floor data. The data needed includes floor hardness and thickness, rebar size and spacing, and the soil compaction or floor load rating and construction if not on ground floor. If the floor proves to be inadequate to support the equipment, the cost of changes to the floor and/or equipment is the responsibility of the customer.")
        
        doc.add_paragraph()
        
        # Add closing and signature block
        closing_para = doc.add_paragraph()
        closing_para.add_run("Best regards,")
        
        # Add some space before signature block
        for _ in range(3):
            doc.add_paragraph()
            
        # Create signature block with a 2x2 table for alignment
        signature_table = doc.add_table(rows=2, cols=2)
        signature_table.autofit = False
        signature_table.style = 'Table Grid'
        signature_table.style = 'Normal Table'  # Remove borders
        
        # Set column widths
        signature_table.columns[0].width = Inches(3.0)
        signature_table.columns[1].width = Inches(3.5)
        
        # Add content to signature table
        signature_table.cell(0, 0).text = "Josh Jancola, Sales Representative"
        
        signature_table.cell(0, 1).paragraphs[0].add_run("Accepted by: ").bold = True
        signature_table.cell(0, 1).paragraphs[0].add_run("_________________________")
        
        signature_table.cell(1, 0).text = ""
        
        signature_table.cell(1, 1).paragraphs[0].add_run("Signature: ").bold = True
        signature_table.cell(1, 1).paragraphs[0].add_run("_________________________")
        
        # Add date line
        date_para = doc.add_paragraph()
        date_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        date_para.add_run("Date: ").bold = True
        date_para.add_run("_____________________________")"""
    
    # Add the proposal details page code to the file
    modified_content = re.sub(contact_info_pattern, proposal_details_code, content)
    
    # Check if we made any changes
    if modified_content == original_content:
        logger.warning("No changes were made to the file")
        return False
    
    # Write the modified content back to the file
    with open(pdf_processor_path, 'w') as f:
        f.write(modified_content)
    
    logger.info("Successfully added proposal details page")
    return True

if __name__ == "__main__":
    add_proposal_details_page()