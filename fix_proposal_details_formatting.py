"""
This script fixes the formatting issues with the proposal details page:
1. Remove duplicate header
2. Fix heading style (black, underlined, smaller font)
3. Fix payment schedule alignment and spacing
4. Reduce overall spacing to fit content on one page
5. Fix Lead Time spacing
"""
import os
import re
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def fix_proposal_details_formatting():
    """
    Fix formatting issues with the proposal details page to match the
    desired layout from the reference image.
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
    
    # 1. Fix the header issue - remove the duplicate header
    # Replace the header code to use the same style as other pages
    
    header_pattern = r"""        # Add page 14 header with page number
        section = doc\.sections\[-1\]
        header = section\.header
        header_para = header\.paragraphs\[0\] if header\.paragraphs else header\.add_paragraph\(\)
        header_para\.style = 'Header'
        header_para\.alignment = WD_ALIGN_PARAGRAPH\.RIGHT
        tab_stops = header_para\.paragraph_format\.tab_stops
        tab_stops\.add_tab_stop\(Inches\(1\.0\), WD_TAB_ALIGNMENT\.LEFT\)
        header_para\.text = ""
        # Get customer name from the context or use a default
        if hasattr\(customer_info, 'get'\) and customer_info\.get\('customer_name'\):
            cust_name = customer_info\.get\('customer_name'\)
        elif hasattr\(customer_info, 'customer_name'\) and customer_info\.customer_name:
            cust_name = customer_info\.customer_name
        else:
            cust_name = "Customer"
            
        header_para\.add_run\(f"\{cust_name\}\\t"\)\.bold = True
        current_date = datetime\.now\(\)\.strftime\('%B %d, %Y'\)
        header_para\.add_run\(f"Page 14\\n\{current_date\}"\)"""
    
    header_replacement = """        # Setup header and footer like other pages
        # Use the existing setup_document_headers_and_footers function
        setup_document_headers_and_footers(
            doc, doc.sections[-1],
            customer_info.get('customer_name', customer_info.customer_name if hasattr(customer_info, 'customer_name') else "Customer"),
            customer_info.get('proposal_number', ''),
            datetime.now().strftime('%m/%d/%Y'),
            14  # Page number
        )"""
    
    content = re.sub(header_pattern, header_replacement, content)
    
    # 2. Fix the heading style - change from blue to black, reduce font size, add underline
    # Update the "Proposal includes:" heading
    heading1_pattern = r"""        # Add "Proposal includes:" heading
        doc\.add_heading\('Proposal includes:', level=1\)"""
    
    heading1_replacement = """        # Add "Proposal includes:" heading - black, underlined, smaller
        proposal_includes_para = doc.add_paragraph()
        proposal_includes_run = proposal_includes_para.add_run('Proposal includes:')
        proposal_includes_run.font.size = Pt(11)  # Smaller font size
        proposal_includes_run.font.bold = True
        proposal_includes_run.font.underline = True  # Add underline
        proposal_includes_run.font.color.rgb = RGBColor(0, 0, 0)  # Black color"""
    
    content = re.sub(heading1_pattern, heading1_replacement, content)
    
    # Update the "Proposal does not include:" heading
    heading2_pattern = r"""        # Add "Proposal does not include:" heading
        doc\.add_heading\('Proposal does not include:', level=1\)"""
    
    heading2_replacement = """        # Add "Proposal does not include:" heading - black, underlined, smaller
        proposal_excludes_para = doc.add_paragraph()
        proposal_excludes_run = proposal_excludes_para.add_run('Proposal does not include:')
        proposal_excludes_run.font.size = Pt(11)  # Smaller font size
        proposal_excludes_run.font.bold = True
        proposal_excludes_run.font.underline = True  # Add underline
        proposal_excludes_run.font.color.rgb = RGBColor(0, 0, 0)  # Black color"""
    
    content = re.sub(heading2_pattern, heading2_replacement, content)
    
    # 3 & 4. Fix payment schedule and general spacing
    # Remove the empty paragraph after headings to reduce spacing
    content = content.replace(r'doc.add_paragraph()', 'doc.add_paragraph("")', 1)  # Keep first one but empty
    content = content.replace(r'doc.add_paragraph()', '', 2)  # Remove the second and third
    
    # Fix payment schedule centering and make left-aligned
    payment_pattern = r"""        for item in payment_schedule:
            p = doc\.add_paragraph\(\)
            p\.alignment = WD_ALIGN_PARAGRAPH\.CENTER
            p\.add_run\(item\)"""
    
    payment_replacement = """        for item in payment_schedule:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.left_indent = Inches(2.0)  # Indent for better alignment
            p.add_run(item)"""
    
    content = re.sub(payment_pattern, payment_replacement, content)
    
    # 5. Fix Lead Time spacing - make it more compact
    lead_time_pattern = r"""        # Add Lead Time section
        lead_time_heading = doc\.add_paragraph\(\)
        lead_time_heading\.add_run\("Lead Time"\)\.bold = True
        lead_time_heading\.add_run\(": 10-12 Weeks"\)
        
        doc\.add_paragraph\(\)"""
    
    lead_time_replacement = """        # Add Lead Time section - more compact
        lead_time_heading = doc.add_paragraph()
        lead_time_heading.add_run("Lead Time").bold = True
        lead_time_heading.add_run(": 10-12 Weeks")"""
    
    content = re.sub(lead_time_pattern, lead_time_replacement, content)
    
    # Fix payment heading spacing
    payment_heading_pattern = r"""        # Add Payment Schedule section
        payment_heading = doc\.add_paragraph\(\)
        payment_heading\.add_run\("Payment Schedule:"\)\.bold = True"""
    
    payment_heading_replacement = """        # Add Payment Schedule section - less spacing
        payment_heading = doc.add_paragraph()
        payment_heading.paragraph_format.space_before = Pt(6)  # Reduce space before
        payment_heading.add_run("Payment Schedule:").bold = True"""
    
    content = re.sub(payment_heading_pattern, payment_heading_replacement, content)
    
    # Add imports for the new code if needed
    rgb_import = "from docx.shared import RGBColor, Pt"
    if rgb_import not in content:
        content = content.replace("from docx.shared import Inches, Pt", "from docx.shared import Inches, Pt, RGBColor")
    
    # Check if we made any changes
    if content == original_content:
        logger.warning("No changes were made to the file")
        return False
    
    # Write the modified content back to the file
    with open(pdf_processor_path, 'w') as f:
        f.write(content)
    
    logger.info("Successfully fixed proposal details formatting")
    return True

if __name__ == "__main__":
    fix_proposal_details_formatting()