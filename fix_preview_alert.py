#!/usr/bin/env python3
"""
Script to fix the alert message not appearing when downloading from the preview page 
by ensuring the bootstrap modal script is properly initialized.
"""

def fix_preview_alert():
    """
    Update the JavaScript in preview.html to ensure the alert modal is properly initialized
    """
    preview_path = 'templates/preview.html'
    
    with open(preview_path, 'r') as file:
        content = file.read()
    
    # Make sure bootstrap.Modal can be initialized correctly
    if 'new bootstrap.Modal' in content:
        # Ensure that the bootstrap JS is properly loaded
        # Check if there's any error in the modal initialization
        lines = content.split('\n')
        for i, line in enumerate(lines):
            if 'new bootstrap.Modal' in line:
                # Verify we have access to bootstrap
                debug_code = """
            // Make sure bootstrap is available
            try {
                console.log('Bootstrap check:', typeof bootstrap);
                const modal = new bootstrap.Modal(document.getElementById('docAlertModal'), {
                    backdrop: 'static'  // Prevents closing when clicking outside
                });
                modal.show();
            } catch (error) {
                console.error('Bootstrap Modal Error:', error);
                alert('Please personalize the Executive Summary on page 2 before sending to the customer. Mike and John are watching 👀');
            }
                """
                # Replace the existing modal initialization with our debugged version
                lines[i] = line.replace('const modal = new bootstrap.Modal', '// Original: const modal = new bootstrap.Modal')
                lines.insert(i, debug_code)
                break
        
        # Write back the modified content
        with open(preview_path, 'w') as file:
            file.write('\n'.join(lines))
        
        print("Preview alert fixed in preview.html")
    else:
        print("Could not find bootstrap.Modal initialization in preview.html")

if __name__ == "__main__":
    fix_preview_alert()
