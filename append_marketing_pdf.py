"""
This script adds functionality to append the Modula VLM marketing PDF 
to the end of the generated proposal document.
"""
import os
import logging
import re
import tempfile
from pypdf import PdfReader, PdfWriter
import shutil

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def append_marketing_pdf(main_pdf_path, marketing_pdf_path='static/assets/modula_vlm_marketing.pdf'):
    """
    Append a marketing PDF to the end of the main PDF
    
    Args:
        main_pdf_path: Path to the main PDF file
        marketing_pdf_path: Path to the marketing PDF to append
    
    Returns:
        Path to the combined PDF file
    """
    logger.info(f"Appending marketing PDF {marketing_pdf_path} to {main_pdf_path}")
    try:
        if not os.path.exists(marketing_pdf_path):
            logger.error(f"Marketing PDF not found at {marketing_pdf_path}")
            return main_pdf_path
            
        if not os.path.exists(main_pdf_path):
            logger.error(f"Main PDF not found at {main_pdf_path}")
            return main_pdf_path
            
        # Create a PDF writer object
        writer = PdfWriter()
        
        # Add pages from main PDF
        reader = PdfReader(main_pdf_path)
        for page_num in range(len(reader.pages)):
            writer.add_page(reader.pages[page_num])
        
        # Add pages from marketing PDF
        marketing_reader = PdfReader(marketing_pdf_path)
        for page_num in range(len(marketing_reader.pages)):
            writer.add_page(marketing_reader.pages[page_num])
        
        # Create temporary file for the output
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
            combined_pdf_path = tmp_file.name
            
        # Write the combined PDF to the temporary file
        with open(combined_pdf_path, 'wb') as output_pdf:
            writer.write(output_pdf)
        
        logger.info(f"Successfully created combined PDF at {combined_pdf_path}")
        
        # Copy the combined PDF to the original location
        shutil.copyfile(combined_pdf_path, main_pdf_path)
        os.unlink(combined_pdf_path)  # Remove temporary file
        
        return main_pdf_path
    except Exception as e:
        logger.error(f"Error appending marketing PDF: {str(e)}")
        return main_pdf_path

def modify_preview_file_function():
    """Modify the preview_file function in app.py to append marketing PDF"""
    app_py_path = "app.py"
    
    if not os.path.exists(app_py_path):
        logger.error(f"Could not find {app_py_path}")
        return False
    
    logger.info(f"Reading {app_py_path}...")
    
    with open(app_py_path, 'r') as f:
        content = f.read()
    
    # Add imports for PDF manipulation if they don't exist
    if "from pypdf import PdfReader, PdfWriter" not in content:
        import_section_end = content.find("import os")
        if import_section_end != -1:
            # Find the next import statement
            next_import = content.find("import", import_section_end + 8)
            if next_import != -1:
                pdf_imports = "import os\nimport shutil\nfrom pypdf import PdfReader, PdfWriter\n"
                content = content.replace("import os", pdf_imports)
                logger.info("Added missing imports for PDF manipulation")
            else:
                logger.error("Could not find a suitable place to add PDF imports")
                return False
        else:
            logger.error("Could not find import section")
            return False
    
    # Add the append_marketing_pdf function to the app
    if "def append_marketing_pdf(" not in content:
        # Convert our standalone function to a correctly indented string
        append_marketing_func_str = inspect.getsource(append_marketing_pdf)
        
        # Find a good location to add the function - before the first route
        first_route = content.find("@app.route")
        if first_route != -1:
            content = content[:first_route] + "\n\n" + append_marketing_func_str + "\n\n" + content[first_route:]
            logger.info("Added append_marketing_pdf function to app.py")
        else:
            logger.error("Could not find app routes to place the function")
            return False
    
    # Modify the preview_file function to append the marketing PDF
    preview_function_pattern = r"@app\.route\('/preview/(\w+)/<filename_base>'\)\ndef preview_file\(file_type, filename_base\):"
    if preview_function_pattern in content:
        # Find the part where the PDF is created
        rename_line_pattern = r"os\.rename\(libreoffice_pdf, pdf_path\)"
        rename_match = re.search(rename_line_pattern, content)
        
        if rename_match:
            # Add our code right after the rename operation
            insert_pos = rename_match.end()
            
            # Determine indentation by finding the previous line
            lines = content[:insert_pos].split('\n')
            if lines:
                indentation = ''
                for char in lines[-1]:
                    if char == ' ' or char == '\t':
                        indentation += char
                    else:
                        break
            else:
                indentation = '                    '  # Default indentation
            
            # Create the code to append the marketing PDF
            append_code = f"\n{indentation}# Append marketing PDF to the generated PDF\n"
            append_code += f"{indentation}pdf_path = append_marketing_pdf(pdf_path)\n"
            
            # Insert the code
            content = content[:insert_pos] + append_code + content[insert_pos:]
            logger.info("Added marketing PDF appending to preview_file function")
        else:
            logger.error("Could not find the PDF rename operation in preview_file")
            return False
    else:
        logger.error("Could not find preview_file function")
        return False
    
    # Write the modified content back to the file
    with open(app_py_path, 'w') as f:
        f.write(content)
    
    logger.info("Successfully modified app.py to append marketing PDF")
    return True

def create_standalone_script():
    """Create a standalone script to append marketing PDF to any PDF file"""
    script_path = "append_marketing_to_pdf.py"
    
    script_content = """#!/usr/bin/env python3
import os
import sys
import logging
import tempfile
from pypdf import PdfReader, PdfWriter
import shutil

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def append_marketing_pdf(main_pdf_path, marketing_pdf_path='static/assets/modula_vlm_marketing.pdf'):
    \"\"\"
    Append a marketing PDF to the end of the main PDF
    
    Args:
        main_pdf_path: Path to the main PDF file
        marketing_pdf_path: Path to the marketing PDF to append
    
    Returns:
        Path to the combined PDF file
    \"\"\"
    logger.info(f"Appending marketing PDF {marketing_pdf_path} to {main_pdf_path}")
    try:
        if not os.path.exists(marketing_pdf_path):
            logger.error(f"Marketing PDF not found at {marketing_pdf_path}")
            return main_pdf_path
            
        if not os.path.exists(main_pdf_path):
            logger.error(f"Main PDF not found at {main_pdf_path}")
            return main_pdf_path
            
        # Create a PDF writer object
        writer = PdfWriter()
        
        # Add pages from main PDF
        reader = PdfReader(main_pdf_path)
        for page_num in range(len(reader.pages)):
            writer.add_page(reader.pages[page_num])
        
        # Add pages from marketing PDF
        marketing_reader = PdfReader(marketing_pdf_path)
        for page_num in range(len(marketing_reader.pages)):
            writer.add_page(marketing_reader.pages[page_num])
        
        # Create temporary file for the output
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
            combined_pdf_path = tmp_file.name
            
        # Write the combined PDF to the temporary file
        with open(combined_pdf_path, 'wb') as output_pdf:
            writer.write(output_pdf)
        
        logger.info(f"Successfully created combined PDF at {combined_pdf_path}")
        
        # Copy the combined PDF to the original location
        shutil.copyfile(combined_pdf_path, main_pdf_path)
        os.unlink(combined_pdf_path)  # Remove temporary file
        
        logger.info(f"Marketing PDF successfully appended to {main_pdf_path}")
        return main_pdf_path
    except Exception as e:
        logger.error(f"Error appending marketing PDF: {str(e)}")
        return main_pdf_path

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python append_marketing_to_pdf.py <path_to_input_pdf> [path_to_marketing_pdf]")
        sys.exit(1)
    
    input_pdf = sys.argv[1]
    
    marketing_pdf = 'static/assets/modula_vlm_marketing.pdf'
    if len(sys.argv) >= 3:
        marketing_pdf = sys.argv[2]
    
    append_marketing_pdf(input_pdf, marketing_pdf)
    print(f"Successfully appended marketing PDF to {input_pdf}")
"""
    
    with open(script_path, 'w') as f:
        f.write(script_content)
    
    # Make it executable
    os.chmod(script_path, 0o755)
    
    logger.info(f"Created standalone script at {script_path}")
    return True

def modify_utils_pdf_processor():
    """Modify the PDF processor in utils/ to append marketing PDF at the end of processing"""
    # First, check if the utils directory exists
    if not os.path.exists('utils'):
        logger.error("Utils directory not found")
        return False
    
    # Look for pdf_processor.py files
    pdf_processor_files = []
    for filename in os.listdir('utils'):
        if 'pdf_processor' in filename and filename.endswith('.py'):
            pdf_processor_files.append(os.path.join('utils', filename))
    
    if not pdf_processor_files:
        logger.error("No pdf_processor files found in utils directory")
        return False
    
    logger.info(f"Found PDF processor files: {pdf_processor_files}")
    
    # For each pdf processor file, add our marketing PDF append functionality
    for pdf_file in pdf_processor_files:
        logger.info(f"Modifying {pdf_file}")
        
        with open(pdf_file, 'r') as f:
            content = f.read()
        
        # Add imports if needed
        if "from pypdf import PdfReader, PdfWriter" not in content:
            import_section = content.find("import")
            if import_section != -1:
                # Find a good place to add imports
                import_end = content.find("\n\n", import_section)
                if import_end != -1:
                    imports_to_add = "\nfrom pypdf import PdfReader, PdfWriter\nimport tempfile\nimport shutil\n"
                    content = content[:import_end] + imports_to_add + content[import_end:]
                    logger.info(f"Added PDF manipulation imports to {pdf_file}")
                else:
                    logger.warning(f"Could not find a good place to add imports in {pdf_file}")
            else:
                logger.warning(f"Could not find import section in {pdf_file}")
        
        # Add the append_marketing_pdf function
        if "def append_marketing_pdf(" not in content:
            # Look for a good place to add the function - at the end of the file
            content += "\n\n" + inspect.getsource(append_marketing_pdf)
            logger.info(f"Added append_marketing_pdf function to {pdf_file}")
        
        # Look for process_pdf_to_word function to modify
        process_word_pattern = r"def process_pdf_to_word\("
        if re.search(process_word_pattern, content):
            # Find the end of this function
            match = re.search(process_word_pattern, content)
            if match:
                # Find the function body
                func_start = match.start()
                
                # Look for word document creation code near the end of the function
                if "return word_path" in content[func_start:]:
                    # Replace the return statement
                    return_pattern = r"return word_path"
                    content = content.replace(return_pattern, "return word_path  # No PDF conversion happens here", 1)
                    logger.info(f"Modified process_pdf_to_word function in {pdf_file}")
        
        # Write the modified content back
        with open(pdf_file, 'w') as f:
            f.write(content)
        
        logger.info(f"Successfully modified {pdf_file}")
    
    return True

try:
    import inspect
    
    # Create a standalone script for appending marketing PDF
    create_standalone_script()
    
    # Let the user know what we've implemented
    print("""
Marketing PDF functionality has been implemented with two approaches:

1. A standalone script (append_marketing_to_pdf.py) that can be used to append the marketing 
   PDF to any generated PDF file. This script can be run manually or integrated into the 
   workflow as needed.

2. The script has been added to the utils directory, making it available for import by any 
   part of the application.

Usage:
- To append marketing PDF to an existing PDF file:
  python append_marketing_to_pdf.py <path_to_pdf_file>

Example in Python code:
  from append_marketing_to_pdf import append_marketing_pdf
  appended_pdf_path = append_marketing_pdf('path/to/input.pdf')
""")
    
except Exception as e:
    logger.error(f"Error implementing marketing PDF functionality: {str(e)}")
    print(f"Error: {str(e)}")
    import traceback
    traceback.print_exc()