import logging
import tempfile
# import numpy as np  # Temporarily disabled due to system dependency issues
import os
from PIL import Image
from pdf2image import convert_from_path

logger = logging.getLogger(__name__)

def process_logo_image(image_path, output_size=(400, 200), output_path=None):
    """
    Process and resize a logo image for inclusion in documents

    Args:
        image_path: Path to the source image
        output_size: Tuple of (width, height) for the output image
        output_path: Optional path to save the processed image

    Returns:
        Path to the processed image or the image object if output_path is None
    """
    try:
        # Open the image
        img = Image.open(image_path)

        # Resize while preserving aspect ratio
        img.thumbnail(output_size, Image.LANCZOS)

        # If output path specified, save the image
        if output_path:
            # Make sure directory exists
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            # Save with transparency if it has an alpha channel
            img.save(output_path)
            logger.info(f"Processed logo saved to {output_path}")
            return output_path
        else:
            # Return the image object
            return img

    except Exception as e:
        logger.error(f"Error processing logo image: {str(e)}")
        return None

def detect_orange_bar(img, threshold_y=200):
    """
    Detects the position of the orange bar in a Modula PDF page image

    Args:
        img: PIL Image object
        threshold_y: Maximum y-coordinate to search (to avoid false positives further down)

    Returns:
        y-coordinate of the bottom of the orange bar, or None if not found
    """
    try:
        # Simplified approach without numpy - just use basic image processing
        width, height = img.size

        # Convert to RGB if needed
        if img.mode != 'RGB':
            img = img.convert('RGB')

        # Search for orange-colored header section by scanning pixels
        # Define the orange color range to detect (in RGB)
        orange_lower = (200, 100, 0)   # Lower bound for orange
        orange_upper = (255, 170, 80)  # Upper bound for orange

        # Only search in the top portion of the image
        search_height = min(threshold_y, height)

        # Search for orange bar from top to bottom
        for y in range(20, search_height):  # Start at 20px to skip very top
            orange_pixels = 0
            
            # Count orange pixels in this row
            for x in range(width):
                pixel = img.getpixel((x, y))
                r, g, b = pixel[:3]  # Handle both RGB and RGBA
                
                # Check if pixel is in orange range
                if (orange_lower[0] <= r <= 255 and
                    orange_lower[1] <= g <= orange_upper[1] and
                    orange_lower[2] <= b <= orange_upper[2]):
                    orange_pixels += 1

            # If we find a significant number of orange pixels (at least 20% of width),
            # this is likely our orange bar
            if orange_pixels > width * 0.2:
                # Look for the bottom of the orange bar by finding where orange ends
                for bottom_y in range(y, min(y + 50, height)):
                    orange_count = 0
                    for x in range(width):
                        pixel = img.getpixel((x, bottom_y))
                        r, g, b = pixel[:3]
                        if (orange_lower[0] <= r <= 255 and
                            orange_lower[1] <= g <= orange_upper[1] and
                            orange_lower[2] <= b <= orange_upper[2]):
                            orange_count += 1
                    
                    if orange_count < width * 0.1:
                        # We've found the bottom of the orange bar
                        logger.info(f"Detected orange bar at y-position: {bottom_y}")
                        return bottom_y

                # If we didn't find the bottom edge, just return a bit below current position
                logger.info(f"Detected partial orange bar at y-position: {y + 20}")
                return y + 20

        # If we get here, no orange bar was detected - use a reasonable default
        logger.info("No orange bar detected, using default crop position of 120px")
        return 120  # Default crop position

    except Exception as e:
        logger.error(f"Error detecting orange bar: {str(e)}")
        # Return a reasonable default in case of error
        return 120

def auto_crop_modula_header(img, fixed_crop_height=None, crop_sides=True, crop_bottom=True):
    """
    Automatically crop the header (logo and page number) from a Modula PDF page image
    and optionally crop left, right, and bottom sides to remove excess white space

    Args:
        img: PIL Image object of a page from Modula PDF
        fixed_crop_height: Optional fixed height in pixels to crop from top of image
        crop_sides: Whether to crop left and right sides to remove excess white space (0.5 inches)
        crop_bottom: Whether to crop the bottom of the image (0.75 inches)

    Returns:
        Cropped PIL Image with header removed and sides/bottom trimmed
    """
    # Get original dimensions
    width, height = img.size

    # Don't crop if original image is very short (likely not a technical page)
    if height <= 500:
        logger.info(f"Image too short ({height}px), skipping auto-crop")
        return img

    try:
        # If a fixed crop height is specified, use that
        # Otherwise calculate it based on DPI (200 dpi × 1 inch = 200 pixels)
        if fixed_crop_height:
            top = fixed_crop_height
        else:
            # Default to 1.25 inches at 200 DPI
            top = 250

        # Set up crop parameters
        # Crop sides if requested (0.5 inches from each side = 100px at 200DPI)
        # Crop bottom if requested (0.5 inches = 100px at 200DPI)
        if crop_sides and crop_bottom:
            # At 200 DPI, 0.5 inch = 100 pixels
            side_crop = 100
            bottom_crop = 100  # Reduced from 150 (0.75") to 100 (0.5")
            left, right = side_crop, width - side_crop
            bottom = height - bottom_crop
            logger.info(f"Cropping header, sides, and bottom: Removing top {top}px, {side_crop}px from sides, and {bottom_crop}px from bottom")
        elif crop_sides:
            side_crop = 100
            left, right = side_crop, width - side_crop
            bottom = height
            logger.info(f"Cropping header and sides: Removing top {top}px and {side_crop}px from each side")
        elif crop_bottom:
            bottom_crop = 100  # Changed from 150 (0.75") to 100 (0.5")
            left, right = 0, width
            bottom = height - bottom_crop
            logger.info(f"Cropping header and bottom: Removing top {top}px and {bottom_crop}px from bottom")
        else:
            left, right, bottom = 0, width, height
            logger.info(f"Cropping header only: Removing top {top}px from image")
            
        cropped_img = img.crop((left, top, right, bottom))
        return cropped_img
    except Exception as crop_e:
        logger.error(f"Error during auto-cropping: {str(crop_e)}")
        return img  # Fallback to original if cropping fails

def extract_pdf_region_as_image(pdf_path, page_num=1, bbox=None, output_path=None, auto_crop_header=True):
    """
    Extract a specific region from a PDF page and save it as an image

    Args:
        pdf_path: Path to the PDF file
        page_num: Page number to extract from (1-based index)
        bbox: Bounding box as PIL-style (left, top, right, bottom). If None, extracts the entire page.
        output_path: Optional path to save the extracted image
        auto_crop_header: Whether to automatically crop the top header with Modula logo and page number

    Returns:
        Path to the saved image or the image object if output_path is None
    """
    try:
        if bbox:
            logger.info(f"Extracting region {bbox} from page {page_num} of {pdf_path}")
        else:
            logger.info(f"Extracting full page {page_num} of {pdf_path}")

        # Create a temporary directory for image conversion
        with tempfile.TemporaryDirectory() as temp_dir:
            # Convert PDF page to image at medium DPI to improve performance
            images = convert_from_path(pdf_path, dpi=200, first_page=page_num, last_page=page_num)

            if not images:
                logger.error(f"Failed to convert page {page_num} to image")
                return None

            # Get the first (and only) image
            img = images[0]

            # Auto-crop the header if requested (and no specific bbox is provided)
            if auto_crop_header and not bbox and page_num >= 2:  # Only crop from page 2 onward
                # Apply top, side, and bottom cropping
                result_img = auto_crop_modula_header(img, crop_sides=True, crop_bottom=True)
            # Extract the region if specified, otherwise use the whole image
            elif bbox:
                result_img = img.crop(bbox)
            else:
                result_img = img

            # If output path specified, save the image
            if output_path:
                # Use a different approach to write the file to avoid PIL issues
                with open(output_path, 'wb') as f:
                    result_img.save(f, format='PNG')
                logger.info(f"Saved extracted image to {output_path}")
                return output_path
            else:
                # Return the image object
                return result_img

    except Exception as e:
        logger.error(f"Error extracting PDF as image: {str(e)}")
        return None
def extract_page_as_image_special(pdf_path, page_num, output_path, dpi=200):
    '''
    Extract a page from a PDF as an image with minimal cropping.
    Only removes the header, preserves sides and bottom.
    Specifically for page 7 of the document.
    
    Args:
        pdf_path: Path to the PDF file
        page_num: Page number to extract (1-indexed)
        output_path: Path to save the extracted image
        dpi: DPI for the extracted image (default: 200)
        
    Returns:
        Path to the extracted image or None on error
    '''
    try:
        # Convert PDF page to image
        from pdf2image import convert_from_path
        
        # Extract just the requested page
        images = convert_from_path(
            pdf_path, 
            dpi=dpi, 
            first_page=page_num, 
            last_page=page_num
        )
        
        if not images:
            logger.error(f"Failed to extract page {page_num} from {pdf_path}")
            return None
            
        # Get the first (and only) image
        img = images[0]
        
        # For page 7, only crop the header (top 250px) - preserve sides and bottom
        width, height = img.size
        cropped_img = img.crop((0, 250, width, height))
        
        # Save the cropped image
        cropped_img.save(output_path)
        logger.info(f"Saved page {page_num} image with special handling: {output_path}")
        
        return output_path
    except Exception as e:
        logger.error(f"Error extracting page as image with special handling: {e}")
        return None
