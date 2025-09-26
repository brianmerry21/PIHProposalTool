"""
This script fixes "must be real number, not NoneType" errors throughout the application
by properly handling None values in templates and calculations.
"""
import os
import re

def fix_price_formatting_in_templates():
    """
    Fix price formatting in templates to handle None values correctly
    """
    # Get the review.html content
    review_html_path = os.path.join('templates', 'review.html')
    with open(review_html_path, 'r') as f:
        content = f.read()
    
    # Replace all instances of price_each without None check
    content = re.sub(
        r'"%.2f"\|format\(item\.price_each\)', 
        '"%.2f"|format(item.price_each or 0.0)', 
        content
    )
    
    # Replace all instances of price_total without None check
    content = re.sub(
        r'"%.2f"\|format\(item\.price_total\)', 
        '"%.2f"|format(item.price_total or 0.0)', 
        content
    )
    
    # Replace total_price formatting without None check
    content = re.sub(
        r'"%.2f"\|format\(total_price\)', 
        '"%.2f"|format(total_price or 0.0)', 
        content
    )
    
    # Write the modified content back
    with open(review_html_path, 'w') as f:
        f.write(content)
    
    print(f"Updated {review_html_path} to properly handle None values in price formatting.")
    
    # Also check other templates like result.html and preview.html
    for template_name in ['result.html', 'preview.html']:
        template_path = os.path.join('templates', template_name)
        if os.path.exists(template_path):
            with open(template_path, 'r') as f:
                content = f.read()
            
            # Replace price formatting
            content = re.sub(
                r'"%.2f"\|format\((.*?)\)', 
                r'"%.2f"|format(\1 or 0.0)', 
                content
            )
            
            with open(template_path, 'w') as f:
                f.write(content)
            
            print(f"Updated {template_path} to properly handle None values in price formatting.")

def fix_repr_method_in_models():
    """
    Update the LineItem __repr__ method to handle None values
    """
    models_path = 'models.py'
    with open(models_path, 'r') as f:
        content = f.read()
    
    # Check if we already fixed the __repr__ method
    if 'price_display =' in content:
        print("LineItem __repr__ method already fixed.")
        return
    
    # Replace the __repr__ method
    old_repr = r"""    def __repr__\(self\):
        return f"<LineItem {self.description} - {'Included' if self.is_included else self.price_total}>"
"""
    
    new_repr = """    def __repr__(self):
        # Handle the case where price_total might be None
        if self.is_included:
            price_display = 'Included'
        else:
            price_display = f"{self.price_total or 0.0}"
        return f"<LineItem {self.description} - {price_display}>"
"""
    
    content = re.sub(old_repr, new_repr, content)
    
    with open(models_path, 'w') as f:
        f.write(content)
    
    print(f"Updated {models_path} to fix the LineItem __repr__ method.")

def fix_handling_in_app_calculation():
    """
    Ensure all price calculations in app.py properly handle None values
    """
    app_path = 'app.py'
    with open(app_path, 'r') as f:
        content = f.read()
    
    # Check if we need to fix price calculation in update_item
    if 'price_each = 0.0 if item.price_each is None else float(item.price_each)' in content:
        print("Price calculations in app.py already handle None values.")
        return
    
    # Update price calculations in update_item function
    old_calc = r"""        # Calculate the total price based on quantity and unit price
        # Only if the item isn't marked as "included"
        if not item.is_included:
            item.price_total = item.price_each * item.quantity
        else:
            item.price_total = 0.0"""
    
    new_calc = """        # Calculate the total price based on quantity and unit price
        # Only if the item isn't marked as "included"
        if not item.is_included:
            # Ensure price_each and quantity are not None
            price_each = 0.0 if item.price_each is None else float(item.price_each)
            quantity = 0.0 if item.quantity is None else float(item.quantity)
            item.price_total = price_each * quantity
        else:
            item.price_total = 0.0"""
    
    content = re.sub(old_calc, new_calc, content)
    
    with open(app_path, 'w') as f:
        f.write(content)
    
    print(f"Updated {app_path} to fix price calculations.")

if __name__ == "__main__":
    print("Fixing None value handling throughout the application...")
    fix_price_formatting_in_templates()
    fix_repr_method_in_models()
    fix_handling_in_app_calculation()
    print("All fixes applied successfully!")