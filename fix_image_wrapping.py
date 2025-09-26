"""
This script modifies the image insertion code to use "In Front of Text" wrapping
while maximizing image size, but avoiding covering important text.
"""
import os
import logging
import re

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def implement_image_wrapping():
    """Implement 'In Front of Text' wrapping for images in the Word document"""
    pdf_processor_path = "utils/pdf_processor.py"
    
    if not os.path.exists(pdf_processor_path):
        logger.error(f"Could not find {pdf_processor_path}")
        return False
    
    logger.info(f"Found pdf_processor.py at {pdf_processor_path}")
    
    # Read the file content
    with open(pdf_processor_path, 'r') as f:
        content = f.read()
    
    # First, create the helper function to set image wrapping
    helper_function = '''
def set_image_in_front_of_text(run):
    """
    Set the image in a run to be positioned in front of text
    (allows for larger images that can extend into margins)
    
    Args:
        run: The run containing the image
    """
    try:
        # Get the drawing element (only exists if the run contains an image)
        drawing_element = None
        for child in run._element:
            if child.tag.endswith(('drawing')):
                drawing_element = child
                break
        
        if drawing_element is not None:
            # Find the appropriate elements to modify
            inline = drawing_element.find('.//wp:inline', namespaces={'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'})
            if inline is not None:
                # Change inline to anchor to enable text wrapping
                anchor = OxmlElement('wp:anchor')
                anchor.set('distT', '0')
                anchor.set('distB', '0')
                anchor.set('distL', '0')
                anchor.set('distR', '0')
                anchor.set('simplePos', '0')
                anchor.set('relativeHeight', '1')
                anchor.set('behindDoc', '0')
                anchor.set('locked', '0')
                anchor.set('layoutInCell', '1')
                anchor.set('allowOverlap', '1')
                
                # Copy attributes and children from inline to anchor
                for key, value in inline.attrib.items():
                    if key != 'distT' and key != 'distB' and key != 'distL' and key != 'distR':
                        anchor.set(key, value)
                
                for child in inline:
                    anchor.append(child)
                
                # Add required child elements for anchor
                simple_pos = OxmlElement('wp:simplePos')
                simple_pos.set('x', '0')
                simple_pos.set('y', '0')
                anchor.insert(0, simple_pos)
                
                # Position relative to page (not margin)
                pos_h = OxmlElement('wp:positionH')
                pos_h.set('relativeFrom', 'page')
                pos_h_align = OxmlElement('wp:align')
                pos_h_align.text = 'center'
                pos_h.append(pos_h_align)
                anchor.append(pos_h)
                
                pos_v = OxmlElement('wp:positionV')
                pos_v.set('relativeFrom', 'page')
                pos_v_offset = OxmlElement('wp:posOffset')
                # Position slightly down from the top of the page to avoid headers
                pos_v_offset.text = '1250000'  # EMUs (English Metric Units)
                pos_v.append(pos_v_offset)
                anchor.append(pos_v)
                
                # Set text wrapping to "in front of text"
                wrap_none = OxmlElement('wp:wrapNone')
                anchor.append(wrap_none)
                
                # Replace the inline element with our new anchor element
                drawing_element.remove(inline)
                drawing_element.append(anchor)
                
                logger.debug("Successfully set image to 'In Front of Text'")
                return True
        
        return False
    except Exception as e:
        logger.warning(f"Error setting image wrapping: {e}")
        return False
'''
    
    # Check if the function already exists
    if "def set_image_in_front_of_text" not in content:
        # Insert the helper function after the imports section
        first_function_match = re.search(r"def\s+\w+\(", content)
        if first_function_match:
            insert_pos = content.rfind("\n\n", 0, first_function_match.start())
            modified_content = content[:insert_pos] + helper_function + content[insert_pos:]
            
            # Update the content
            content = modified_content
            logger.info("Added set_image_in_front_of_text helper function")
        else:
            logger.error("Could not find an appropriate location to insert the helper function")
            return False
    else:
        logger.info("Helper function already exists")
    
    # Now add the wrapping calls after each image insertion
    # Find all image insertion lines
    image_insertion_lines = re.finditer(r"([ \t]*)doc\.add_picture\((.*?)\)", content)
    
    for match in image_insertion_lines:
        full_line = match.group(0)
        indentation = match.group(1)
        
        # Skip if we've already added wrapping to this line
        if full_line + "\n" + indentation + "set_image_in_front_of_text" in content:
            logger.info(f"Wrapping already applied to: {full_line}")
            continue
            
        # Create the replacement with wrapping
        replacement = f"{full_line}\n{indentation}# Apply 'In Front of Text' wrapping to maximize image size\n"
        replacement += f"{indentation}set_image_in_front_of_text(doc.paragraphs[-1].runs[-1])"
        
        # Replace in the content
        content = content.replace(full_line, replacement)
        logger.info(f"Added wrapping to: {full_line}")
    
    # Write the modified content back to the file
    with open(pdf_processor_path, 'w') as f:
        f.write(content)
    
    logger.info("Successfully implemented 'In Front of Text' wrapping for images")
    return True

if __name__ == "__main__":
    implement_image_wrapping()