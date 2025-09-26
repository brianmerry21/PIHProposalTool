"""
This script optimizes the spacing on page 14 (proposal details page) to ensure
all content fits on a single page:
1. Reduce font size to 10pt for all text
2. Set single line spacing
3. Remove all space before/after paragraphs
"""
import os
import re
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def fix_page_14_spacing():
    """
    Modify the formatting of page 14 to make it more compact
    and ensure all content fits on a single page.
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
    
    # Find the section that begins the proposal details page
    details_page_pattern = r"        # Add a page break to start the Proposal Details page\n        doc\.add_page_break\(\)"
    
    # Add code to set global page 14 style right after the page break
    page_14_style_code = """        # Add a page break to start the Proposal Details page
        doc.add_page_break()
        
        # Set global style for page 14 - compact formatting
        style = doc.styles['Normal']
        style.font.size = Pt(10)  # Smaller font size for all text
        style.paragraph_format.space_before = Pt(0)  # No space before paragraphs
        style.paragraph_format.space_after = Pt(0)   # No space after paragraphs
        style.paragraph_format.line_spacing = 1.0    # Single line spacing"""
    
    # Replace the page break code with our enhanced version
    content = content.replace(details_page_pattern, page_14_style_code)
    
    # Explicitly set font size to 10pt for all runs on page 14
    # Find all paragraphs and add font size settings
    
    # 1. Proposal includes heading
    includes_pattern = r"(        proposal_includes_run = proposal_includes_para\.add_run\('Proposal includes:'\))\n        proposal_includes_run\.font\.size = Pt\(11\)"
    includes_replacement = r"\1\n        proposal_includes_run.font.size = Pt(10)"
    content = re.sub(includes_pattern, includes_replacement, content)
    
    # 2. Proposal does not include heading
    excludes_pattern = r"(        proposal_excludes_run = proposal_excludes_para\.add_run\('Proposal does not include:'\))\n        proposal_excludes_run\.font\.size = Pt\(11\)"
    excludes_replacement = r"\1\n        proposal_excludes_run.font.size = Pt(10)"
    content = re.sub(excludes_pattern, excludes_replacement, content)
    
    # 3. Add font size setting for bulleted list items
    bullet_pattern = r"(            p = doc\.add_paragraph\(style='List Bullet'\)\n            p\.add_run\(item\))"
    bullet_replacement = r"\1\n            for run in p.runs:\n                run.font.size = Pt(10)"
    content = re.sub(bullet_pattern, bullet_replacement, content)
    
    # 4. Set paragraph format for all paragraphs after the bullet lists
    # Find the Lead Time section
    lead_time_pattern = r"(        # Add Lead Time section - more compact\n        lead_time_heading = doc\.add_paragraph\(\))"
    lead_time_replacement = r"\1\n        lead_time_heading.paragraph_format.space_before = Pt(0)\n        lead_time_heading.paragraph_format.space_after = Pt(0)"
    content = re.sub(lead_time_pattern, lead_time_replacement, content)
    
    # 5. Lead time text font size
    lead_time_run_pattern = r"(        lead_time_heading\.add_run\(\"Lead Time\"\)\.bold = True\n        lead_time_heading\.add_run\(\": 10-12 Weeks\"\))"
    lead_time_run_replacement = r"\1\n        for run in lead_time_heading.runs:\n            run.font.size = Pt(10)"
    content = re.sub(lead_time_run_pattern, lead_time_run_replacement, content)
    
    # 6. Payment schedule heading
    payment_heading_pattern = r"(        payment_heading = doc\.add_paragraph\(\)\n        payment_heading\.paragraph_format\.space_before = Pt\(6\))"
    payment_heading_replacement = r"\1\n        payment_heading.paragraph_format.space_before = Pt(0)\n        payment_heading.paragraph_format.space_after = Pt(0)"
    content = re.sub(payment_heading_pattern, payment_heading_replacement, content)
    
    # 7. Payment schedule heading text
    payment_heading_run_pattern = r"(        payment_heading\.add_run\(\"Payment Schedule:\"\)\.bold = True)"
    payment_heading_run_replacement = r"\1\n        for run in payment_heading.runs:\n            run.font.size = Pt(10)"
    content = re.sub(payment_heading_run_pattern, payment_heading_run_replacement, content)
    
    # 8. Payment schedule items
    payment_items_pattern = r"(            p = doc\.add_paragraph\(\)\n            p\.alignment = WD_ALIGN_PARAGRAPH\.LEFT\n            p\.paragraph_format\.left_indent = Inches\(2\.0\)  # Indent for better alignment\n            p\.add_run\(item\))"
    payment_items_replacement = r"\1\n            p.paragraph_format.space_before = Pt(0)\n            p.paragraph_format.space_after = Pt(0)\n            for run in p.runs:\n                run.font.size = Pt(10)"
    content = re.sub(payment_items_pattern, payment_items_replacement, content)
    
    # 9. Note paragraphs
    note_pattern = r"(        note1_para = doc\.add_paragraph\(\))"
    note_replacement = r"\1\n        note1_para.paragraph_format.space_before = Pt(0)\n        note1_para.paragraph_format.space_after = Pt(0)"
    content = re.sub(note_pattern, note_replacement, content)
    
    note2_pattern = r"(        note2_para = doc\.add_paragraph\(\))"
    note2_replacement = r"\1\n        note2_para.paragraph_format.space_before = Pt(0)\n        note2_para.paragraph_format.space_after = Pt(0)"
    content = re.sub(note2_pattern, note2_replacement, content)
    
    # 10. Note text
    note1_run_pattern = r"(        note1_para\.add_run\(\"Note 1:\"\)\.bold = True\n        note1_para\.add_run\(\"\\tAll Permits are \\\"By Customer\\\".*?\"\))"
    note1_run_replacement = r"\1\n        for run in note1_para.runs:\n            run.font.size = Pt(10)"
    content = re.sub(note1_run_pattern, note1_run_replacement, content, flags=re.DOTALL)
    
    note2_run_pattern = r"(        note2_para\.add_run\(\"Note 2:\"\)\.bold = True\n        note2_para\.add_run\(\"\\tSeismic calculations for the equipment.*?\"\))"
    note2_run_replacement = r"\1\n        for run in note2_para.runs:\n            run.font.size = Pt(10)"
    content = re.sub(note2_run_pattern, note2_run_replacement, content, flags=re.DOTALL)
    
    # 11. Best regards and signature section
    closing_pattern = r"(        # Add closing and signature block\n        closing_para = doc\.add_paragraph\(\))"
    closing_replacement = r"\1\n        closing_para.paragraph_format.space_before = Pt(0)\n        closing_para.paragraph_format.space_after = Pt(0)"
    content = re.sub(closing_pattern, closing_replacement, content)
    
    closing_run_pattern = r"(        closing_para\.add_run\(\"Best regards,\"\))"
    closing_run_replacement = r"\1\n        for run in closing_para.runs:\n            run.font.size = Pt(10)"
    content = re.sub(closing_run_pattern, closing_run_replacement, content)
    
    # 12. Reduce space before signature block
    signature_pattern = r"        # Add some space before signature block\n        for _ in range\(3\):\n            doc\.add_paragraph\(\)"
    signature_replacement = r"        # Add minimal space before signature block\n        doc.add_paragraph()"
    content = re.sub(signature_pattern, signature_replacement, content)
    
    # 13. Signature table content
    signature_text_pattern = r"(        # Add content to signature table\n        signature_table\.cell\(0, 0\)\.text = \"Josh Jancola, Sales Representative\")"
    signature_text_replacement = r"\1\n        for para in signature_table.cell(0, 0).paragraphs:\n            for run in para.runs:\n                run.font.size = Pt(10)"
    content = re.sub(signature_text_pattern, signature_text_replacement, content)
    
    # 14. Signature block accepted by/signature
    for i in range(2):
        for j in range(2):
            if i == 0 and j == 0:
                continue  # Skip the first cell which we already handled
            
            cell_pattern = f"(        signature_table\.cell\({i}, {j}\)\.paragraphs\[0\]\.add_run\(.*?\))(\.bold = True)?"
            cell_replacement = r"\1\2\n        for run in signature_table.cell({i}, {j}).paragraphs[0].runs:\n            run.font.size = Pt(10)".format(i=i, j=j)
            content = re.sub(cell_pattern, cell_replacement, content)
    
    # 15. Date line
    date_pattern = r"(        # Add date line\n        date_para = doc\.add_paragraph\(\))"
    date_replacement = r"\1\n        date_para.paragraph_format.space_before = Pt(0)"
    content = re.sub(date_pattern, date_replacement, content)
    
    date_run_pattern = r"(        date_para\.add_run\(\"Date: \"\)\.bold = True\n        date_para\.add_run\(\"_____________________________\"\))"
    date_run_replacement = r"\1\n        for run in date_para.runs:\n            run.font.size = Pt(10)"
    content = re.sub(date_run_pattern, date_run_replacement, content)
    
    # Check if we made any changes
    if content == original_content:
        logger.warning("No changes were made to the file")
        return False
    
    # Write the modified content back to the file
    with open(pdf_processor_path, 'w') as f:
        f.write(content)
    
    logger.info("Successfully modified page 14 spacing")
    return True

if __name__ == "__main__":
    fix_page_14_spacing()