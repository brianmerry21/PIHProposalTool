import os
import logging

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def remove_recommended_functionality():
    """
    Remove the 'is_recommended' functionality from the application
    """
    try:
        # 1. Remove/disable the update-recommended route in app.py
        app_path = 'app.py'
        with open(app_path, 'r') as f:
            app_content = f.read()
            
        # Comment out or modify the update-recommended route
        if '@app.route(\'/update-recommended/<int:item_id>\', methods=[\'POST\'])' in app_content:
            route_start = app_content.find('@app.route(\'/update-recommended/<int:item_id>\', methods=[\'POST\'])')
            route_end = app_content.find('return jsonify', route_start)
            route_end = app_content.find('}\n    )', route_end) + 7  # Find the end of the route function
            
            # Get the indentation
            indent = ''
            for i in range(route_start - 1, 0, -1):
                if app_content[i] == '\n':
                    break
                indent = app_content[i] + indent
                
            # Replace the route with a simplified version that does nothing
            simplified_route = f"""@app.route('/update-recommended/<int:item_id>', methods=['POST'])
def update_recommended(item_id):
    \"\"\"This route is deprecated and no longer used\"\"\"
    return jsonify({{'success': True, 'item_id': item_id, 'is_recommended': False}})"""
            
            new_app_content = app_content[:route_start] + simplified_route + app_content[route_end:]
            
            # Write the modified file
            with open(app_path, 'w') as f:
                f.write(new_app_content)
                
            logger.info("Successfully removed update-recommended route functionality from app.py")
            
        # 2. Remove the Javascript function from review.html
        review_path = 'templates/review.html'
        
        # Fix the "Option" column for optional items
        with open(review_path, 'r') as f:
            review_content = f.read()
            
        # Find the function that updates recommended status and remove it
        update_rec_func_start = review_content.find('function updateRecommendedStatus(')
        if update_rec_func_start > 0:
            update_rec_func_end = review_content.find('};', update_rec_func_start)
            update_rec_func_end = review_content.find('\n', update_rec_func_end) + 1
            
            # Remove the function
            new_review_content = review_content[:update_rec_func_start] + review_content[update_rec_func_end:]
            
            # Write the modified file
            with open(review_path, 'w') as f:
                f.write(new_review_content)
                
            logger.info("Successfully removed updateRecommendedStatus function from review.html")
            
        # 3. Make sure "Option" is displayed in the Word document for optional items
        pdf_processor_path = 'utils/pdf_processor.py'
        with open(pdf_processor_path, 'r') as f:
            pdf_content = f.read()
        
        if 'cells[2].text = f"${price_each:,.2f}" if price_each > 0 else "Option"' in pdf_content:
            # Ensure all optional items show "Option" in the prices column
            new_pdf_content = pdf_content.replace(
                'cells[2].text = f"${price_each:,.2f}" if price_each > 0 else "Option"',
                'cells[2].text = "Option"'
            )
            
            with open(pdf_processor_path, 'w') as f:
                f.write(new_pdf_content)
                
            logger.info("Successfully updated Word document generation for optional items")
            
        return True
        
    except Exception as e:
        logger.error(f"Error while removing recommended functionality: {str(e)}")
        return False

if __name__ == "__main__":
    success = remove_recommended_functionality()
    if success:
        print("Successfully removed 'recommended' functionality from the application.")
    else:
        print("Error occurred while removing 'recommended' functionality.")