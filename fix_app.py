"""
Fix syntax errors in app.py
"""
import re

def fix_app_py():
    with open('app.py', 'r') as f:
        content = f.read()
    
    # Fix the duplicate is_section_header and missing comma issues
    pattern = r'is_section_header=True,\s+is_section_header=True,\s+display_order=display_order\s+is_optional=True'
    replacement = 'is_section_header=True,\n            is_optional=True,  # Mark as optional to filter from main table\n            display_order=display_order'
    
    content = re.sub(pattern, replacement, content)
    
    # Write the fixed content back
    with open('app.py', 'w') as f:
        f.write(content)
    
    print("Fixed app.py")

if __name__ == "__main__":
    fix_app_py()