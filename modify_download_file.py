"""
This script modifies the download_file function in app.py to automatically
append the marketing PDF to downloaded PDFs.
"""
import os
import re
import logging

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def modify_download_file():
    """Modify the download_file function in app.py to append marketing PDF to downloaded PDFs"""
    app_py_path = "app.py"
    
    if not os.path.exists(app_py_path):
        logger.error(f"Could not find {app_py_path}")
        return False
    
    logger.info(f"Reading {app_py_path}...")
    
    with open(app_py_path, 'r') as f:
        content = f.read()
    
    # Add imports for the PDF append utility if they don't exist
    if "from utils.pdf_append import append_marketing_pdf" not in content:
        # Find import section
        imports_end = content.find("\n\n", content.find("import"))
        if imports_end != -1:
            new_import = "\nfrom utils.pdf_append import append_marketing_pdf"
            content = content[:imports_end] + new_import + content[imports_end:]
            logger.info("Added import for pdf_append utility")
        else:
            logger.error("Could not find import section")
            return False
    
    # Look for the download_file function
    download_func_pattern = r"@app\.route\('/download/.*?'\)\ndef download_file\(file_type, filename_base\):"
    download_match = re.search(download_func_pattern, content)
    
    if not download_match:
        logger.error("Could not find download_file function")
        return False
    
    # Find where file_type == 'word' in the download_file function
    word_pattern = r"elif file_type == 'word':"
    word_match = re.search(word_pattern, content)
    
    if not word_match:
        logger.error("Could not find 'word' file type handling in download_file function")
        return False
    
    # Find the line that returns send_file in the word section
    send_file_pattern = r"return send_file\(file_path, as_attachment=True,.*?\)"
    send_file_matches = re.finditer(send_file_pattern, content[word_match.end():])
    
    # Get the first match after the word_match
    send_file_match = None
    for match in send_file_matches:
        send_file_match = match
        break
    
    if not send_file_match:
        logger.error("Could not find send_file in 'word' section of download_file function")
        return False
    
    # Calculate the actual position in the full content
    send_file_pos = word_match.end() + send_file_match.start()
    
    # Check if our PDF conversion code is already present
    pdf_check = "# Convert DOCX to PDF and append marketing PDF"
    if pdf_check in content[word_match.end():send_file_pos]:
        logger.info("PDF conversion and append code already exists")
        return True
    
    # Find the file_path assignment line
    file_path_pattern = r"file_path = os\.path\.join\(UPLOAD_FOLDER, f\"\{filename_base\}\.docx\"\)"
    file_path_match = re.search(file_path_pattern, content[word_match.end():send_file_pos])
    
    if not file_path_match:
        logger.error("Could not find file_path assignment in 'word' section")
        return False
    
    # Calculate the actual position in the full content
    file_path_pos = word_match.end() + file_path_match.end()
    
    # Find indentation by looking at the line with file_path
    lines = content[:file_path_pos].split('\n')
    indentation = ''
    for char in lines[-1]:
        if char == ' ' or char == '\t':
            indentation += char
        else:
            break
    
    # Create our PDF conversion and append code with the correct indentation
    pdf_code = f"\n{indentation}# If PDF is requested instead, convert DOCX to PDF and append marketing PDF\n"
    pdf_code += f"{indentation}if request.args.get('format') == 'pdf':\n"
    pdf_code += f"{indentation}    # Create a PDF filename based on the original DOCX\n"
    pdf_code += f"{indentation}    pdf_path = os.path.join(UPLOAD_FOLDER, f\"{{filename_base}}_download.pdf\")\n"
    pdf_code += f"{indentation}    \n"
    pdf_code += f"{indentation}    # Convert DOCX to PDF using LibreOffice\n"
    pdf_code += f"{indentation}    import subprocess\n"
    pdf_code += f"{indentation}    logger.info(f\"Converting {{file_path}} to PDF...\")\n"
    pdf_code += f"{indentation}    cmd = [\n"
    pdf_code += f"{indentation}        \"libreoffice\", \n"
    pdf_code += f"{indentation}        \"--headless\", \n"
    pdf_code += f"{indentation}        \"--convert-to\", \n"
    pdf_code += f"{indentation}        \"pdf\", \n"
    pdf_code += f"{indentation}        \"--outdir\", \n"
    pdf_code += f"{indentation}        UPLOAD_FOLDER,\n"
    pdf_code += f"{indentation}        file_path\n"
    pdf_code += f"{indentation}    ]\n"
    pdf_code += f"{indentation}    process = subprocess.run(cmd, capture_output=True, text=True)\n"
    pdf_code += f"{indentation}    \n"
    pdf_code += f"{indentation}    # Check if conversion was successful\n"
    pdf_code += f"{indentation}    if process.returncode != 0:\n"
    pdf_code += f"{indentation}        logger.error(f\"Error converting DOCX to PDF: {{process.stderr}}\")\n"
    pdf_code += f"{indentation}        flash(\"Error converting document to PDF\", 'danger')\n"
    pdf_code += f"{indentation}        return send_file(file_path, as_attachment=True, download_name=f\"report_{{filename_base}}.docx\")\n"
    pdf_code += f"{indentation}    \n"
    pdf_code += f"{indentation}    # LibreOffice creates the PDF with the same base name\n"
    pdf_code += f"{indentation}    libreoffice_pdf = os.path.join(UPLOAD_FOLDER, f\"{{filename_base}}.pdf\")\n"
    pdf_code += f"{indentation}    if os.path.exists(libreoffice_pdf):\n"
    pdf_code += f"{indentation}        # Move to our download filename\n"
    pdf_code += f"{indentation}        os.rename(libreoffice_pdf, pdf_path)\n"
    pdf_code += f"{indentation}        \n"
    pdf_code += f"{indentation}        # Append marketing PDF\n"
    pdf_code += f"{indentation}        pdf_path = append_marketing_pdf(pdf_path)\n"
    pdf_code += f"{indentation}        \n"
    pdf_code += f"{indentation}        # Send the PDF file\n"
    pdf_code += f"{indentation}        return send_file(pdf_path, as_attachment=True, download_name=f\"report_{{filename_base}}.pdf\")\n"
    pdf_code += f"{indentation}    else:\n"
    pdf_code += f"{indentation}        logger.error(f\"PDF not created at expected path: {{libreoffice_pdf}}\")\n"
    pdf_code += f"{indentation}        flash(\"Error creating PDF\", 'danger')\n"
    pdf_code += f"{indentation}        return send_file(file_path, as_attachment=True, download_name=f\"report_{{filename_base}}.docx\")\n"
    
    # Insert the code
    content = content[:file_path_pos] + pdf_code + content[file_path_pos:]
    
    # Write the updated content back to app.py
    with open(app_py_path, 'w') as f:
        f.write(content)
    
    logger.info("Successfully modified download_file function to support PDF format with marketing materials")
    return True

def add_pdf_button_to_template():
    """Add PDF download button to the result template"""
    template_path = "templates/result.html"
    
    if not os.path.exists(template_path):
        logger.error(f"Could not find {template_path}")
        return False
    
    logger.info(f"Reading {template_path}...")
    
    with open(template_path, 'r') as f:
        content = f.read()
    
    # Check if our PDF button is already present
    if "Download PDF with Marketing Materials" in content:
        logger.info("PDF download button already exists")
        return True
    
    # Find the Word download button
    word_button_pattern = r'<a href="{{ file_data\.word_download_url }}".*?class="btn btn-.*?".*?>.*?Word.*?</a>'
    word_button_match = re.search(word_button_pattern, content)
    
    if not word_button_match:
        logger.error("Could not find Word download button in template")
        return False
    
    # Get the full button HTML
    word_button_html = word_button_match.group(0)
    
    # Create our PDF button based on the Word button
    pdf_button_html = word_button_html.replace(
        '{{ file_data.word_download_url }}', 
        '{{ file_data.word_download_url }}?format=pdf'
    ).replace(
        'Word', 
        'PDF with Marketing'
    )
    
    # Insert the new button after the Word button
    new_buttons = word_button_html + '\n            ' + pdf_button_html
    content = content.replace(word_button_html, new_buttons)
    
    # Write the updated content back to the template
    with open(template_path, 'w') as f:
        f.write(content)
    
    logger.info("Successfully added PDF download button to the result template")
    return True

if __name__ == "__main__":
    success = modify_download_file()
    if success:
        add_pdf_button_to_template()
        print("Successfully modified app.py to support PDF downloads with marketing materials")
    else:
        print("Failed to modify app.py")