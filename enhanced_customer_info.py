def extract_customer_info_from_pdf(pdf_text):
    """
    Extract customer specific information from PDF
    
    Args:
        pdf_text: Full text content from PDF as a string
        
    Returns:
        Dictionary with customer info (name, location, date, proposal_number, contact_info)
    """
    import re
    import logging
    from datetime import datetime
    logger = logging.getLogger(__name__)
    
    customer_info = {
        "name": "",  # Default   #Thermo Fisher
        "location": "Hillsboro, OR",  # Default
        "date": datetime.now().strftime('%B %d, %Y'),
        "proposal_number": "00000000",
        "model": "Modula ML25 Vertical Lift Module",
        "contact_name": "Josh Jancola",
        "contact_email": "joshjancola@pacificintegrated.com",
        "contact_phone": "253.500.4193",
        "contact_office": "888.550.5888"
    }
    
    # Try to extract customer name from the header
    # customer_pattern = r"(Thermo\s+Fisher|.*?(?:Scientific|Inc\.?|Corp\.?|Corporation|Company|LLC))"
    # customer_match = re.search(customer_pattern, pdf_text[:500])
    # if customer_match:
    #     customer_info["name"] = customer_match.group(1).strip()
    import re

    # Assuming pdf_text_pages is a list of strings, each string is a page's text
    first_page_text = pdf_text_pages[0]

    # Remove repeated header/footer if you know what it looks like
    # For example, remove "Thermo Fisher" if it's always in the header
    first_page_text = first_page_text.replace("Dart Aerospace", "")

    # Now search for the actual customer name with common suffixes
    customer_match = re.search(
        r"\b([A-Z][A-Za-z&,\.\s]*(?:Inc\.?|Corp\.?|Corporation|Aerospace|Company|LLC|Scientific)?)\b",
        first_page_text
    )

    if customer_match:
        client_name = customer_match.group(1).strip()
    else:
        client_name = "Unknown Customer"
    # Try to extract proposal number (based on screenshot format)
    proposal_pattern = r"Proposal\s+[#]?([0-9]+)"
    proposal_match = re.search(proposal_pattern, pdf_text[:1000])
    if proposal_match:
        customer_info["proposal_number"] = proposal_match.group(1)
    
    # Try to extract location - look for cleanroom location
    location_pattern = r"(?:cleanroom at the|facility at|facility in)\s+(.*?)(?:facility|\.)"
    location_match = re.search(location_pattern, pdf_text[:2000])
    if location_match:
        location = location_match.group(1).strip()
        if location:
            customer_info["location"] = location
    
    # Backup location pattern
    if "location" not in customer_info or not customer_info["location"]:
        location_backup = r"(?:Hillsboro|Seattle|Portland|Tacoma|Bellevue|Vancouver),\s+(?:OR|WA|CA|ID)"
        location_match = re.search(location_backup, pdf_text[:2000])
        if location_match:
            customer_info["location"] = location_match.group(0)
    
    # Try to extract date from the header
    date_pattern = r"((?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4})"
    date_match = re.search(date_pattern, pdf_text[:1000])
    if date_match:
        customer_info["date"] = date_match.group(1)
    
    # Try to extract model from the header or content
    model_pattern = r"Modula\s+(?:VLM\s+)?(?:ML|MC|MX|MS)\d+(?:-\d+)?"
    model_match = re.search(model_pattern, pdf_text[:2000])
    if model_match:
        customer_info["model"] = model_match.group(0)
    
    # Try to extract tray information
    tray_pattern = r"contain\s+(\d+)\s+slotted\s+trays"
    tray_match = re.search(tray_pattern, pdf_text)
    if tray_match:
        customer_info["num_trays"] = tray_match.group(1)
    
    # Try to extract tray dimensions
    dimension_pattern = r"sized\s+as\s+([\d\.]+)(?:\"|\s+inches|\s*\')\s+wide\s+and\s+([\d\.]+)(?:\"|\s+inches|\s*\')\s+deep"
    dimension_match = re.search(dimension_pattern, pdf_text)
    if dimension_match:
        width = dimension_match.group(1)
        depth = dimension_match.group(2)
        customer_info["tray_dimensions"] = f"Width: {width}'' x Depth: {depth}''"
    
    # Try to extract contact info (based on screenshot format)
    contact_name_pattern = r"([A-Z][a-z]+\s+[A-Z][a-z]+)[\s\n]*Pacific\s+Integrated\s+Handling"
    contact_name_match = re.search(contact_name_pattern, pdf_text)
    if contact_name_match:
        customer_info["contact_name"] = contact_name_match.group(1)
    
    # Try to extract phone
    phone_pattern = r"Cell:\s+([\d\.\-]+)"
    phone_match = re.search(phone_pattern, pdf_text)
    if phone_match:
        customer_info["contact_phone"] = phone_match.group(1)
    
    # Try to extract email
    email_pattern = r"\b([A-Za-z0-9._%+-]+@pacificintegrated\.com)\b"
    email_match = re.search(email_pattern, pdf_text)
    if email_match:
        customer_info["contact_email"] = email_match.group(1)
    
    # Look for recipient name (person being addressed)
    recipient_pattern = r"([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+),\s*[\n\r]"
    recipient_match = re.search(recipient_pattern, pdf_text[:2000])
    if recipient_match:
        customer_info["recipient_name"] = recipient_match.group(1)
    
    # Extract storage capacity information
    capacity_pattern = r"storage\s+capacity\s+of\s+([\d\.]+)\s+sq\s+ft\.\s+and\s+([\d\.]+)\s+cubic\s+ft\."
    capacity_match = re.search(capacity_pattern, pdf_text)
    if capacity_match:
        sq_ft = capacity_match.group(1)
        cubic_ft = capacity_match.group(2)
        customer_info["storage_capacity"] = f"{sq_ft} sq ft. / {cubic_ft} cubic ft."
    
    # Extract footprint
    footprint_pattern = r"footprint\s+of\s+([\d\.]+)\s+sq\.\s+ft\."
    footprint_match = re.search(footprint_pattern, pdf_text)
    if footprint_match:
        customer_info["unit_footprint"] = f"{footprint_match.group(1)} sq. ft."
    
    # Extract height
    height_pattern = r"height\s+of\s+([\d\.]+)\s+inches\s+\(([\d\'\.\"]+)\)"
    height_match = re.search(height_pattern, pdf_text)
    if height_match:
        inches = height_match.group(1)
        feet = height_match.group(2)
        customer_info["unit_height"] = f"{inches} inches ({feet})"
    
    return customer_info