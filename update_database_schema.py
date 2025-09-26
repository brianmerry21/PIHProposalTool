#!/usr/bin/env python3
"""
Database schema update script to add excel_path field to PDFExtraction model
"""

import os
import sys
from sqlalchemy import text

# Add the current directory to the Python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import app, db
from models import PDFExtraction

def update_database_schema():
    """Add excel_path column to pdf_extraction table if it doesn't exist"""
    with app.app_context():
        try:
            # Check if the column already exists
            result = db.session.execute(text("PRAGMA table_info(pdf_extraction)"))
            columns = [row[1] for row in result.fetchall()]
            
            if 'excel_path' not in columns:
                print("Adding excel_path column to pdf_extraction table...")
                db.session.execute(text("ALTER TABLE pdf_extraction ADD COLUMN excel_path VARCHAR(512)"))
                db.session.commit()
                print("✓ excel_path column added successfully!")
            else:
                print("✓ excel_path column already exists!")
                
        except Exception as e:
            print(f"Error updating database schema: {e}")
            db.session.rollback()
            return False
            
    return True

if __name__ == "__main__":
    print("Updating database schema...")
    if update_database_schema():
        print("Database schema updated successfully!")
    else:
        print("Failed to update database schema!")
        sys.exit(1)