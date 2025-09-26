"""
This script modifies the preview_file function in app.py to automatically
append the marketing PDF to generated preview PDFs.
"""
import os
import re
import logging

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def modify_preview_file():
    """Modify the preview_file function in app.py to append marketing PDF"""
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
    
    # Look for the preview_file function
    preview_func_pattern = r"@app\.route\('/preview/.*?'\)\ndef preview_file\(file_type, filename_base\):"
    preview_match = re.search(preview_func_pattern, content)
    
    if not preview_match:
        logger.error("Could not find preview_file function")
        return False
    
    # Find the rename operation in the preview_file function
    rename_pattern = r"os\.rename\(libreoffice_pdf, pdf_path\)"
    rename_match = re.search(rename_pattern, content)
    
    if not rename_match:
        logger.error("Could not find PDF rename operation in preview_file function")
        return False
    
    # Check if our append code is already present
    append_check = "# Append marketing PDF to the generated PDF"
    if append_check in content:
        logger.info("Marketing PDF append code already exists")
        return True
    
    # Insert our code after the rename operation
    insert_pos = rename_match.end()
    
    # Find indentation by looking at the line with the rename operation
    lines = content[:insert_pos].split('\n')
    indentation = ''
    for char in lines[-1]:
        if char == ' ' or char == '\t':
            indentation += char
        else:
            break
    
    # Create our append code with the correct indentation
    append_code = f"\n{indentation}# Append marketing PDF to the generated PDF\n"
    append_code += f"{indentation}pdf_path = append_marketing_pdf(pdf_path)\n"
    
    # Insert the code
    content = content[:insert_pos] + append_code + content[insert_pos:]
    
    # Write the updated content back to app.py
    with open(app_py_path, 'w') as f:
        f.write(content)
    
    logger.info("Successfully modified preview_file function to append marketing PDF")
    return True

if __name__ == "__main__":
    success = modify_preview_file()
    if success:
        print("Successfully modified app.py to append marketing PDF to preview files")
    else:
        print("Failed to modify app.py")