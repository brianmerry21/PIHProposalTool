#!/usr/bin/env python3
"""
Script to fix the save buttons in the review.html template by removing the
onclick event that is preventing them from working properly.
"""

import re

def fix_save_buttons():
    """
    Remove the onclick attribute from save buttons in the review.html template
    """
    template_path = 'templates/review.html'
    
    with open(template_path, 'r') as file:
        content = file.read()
    
    # Replace the onclick attribute using regex for more precision
    fixed_content = re.sub(
        r'class="btn btn-sm btn-outline-success save-item-btn d-none rounded-circle me-1"([^>]*?)onclick="event\.preventDefault\(\); return false;"',
        r'class="btn btn-sm btn-outline-success save-item-btn d-none rounded-circle me-1"\1',
        content
    )
    
    # Write the modified content back to the file
    with open(template_path, 'w') as file:
        file.write(fixed_content)
    
    print("Save buttons fixed in review.html")

if __name__ == "__main__":
    fix_save_buttons()
