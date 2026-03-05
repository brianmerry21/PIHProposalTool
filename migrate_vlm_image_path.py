"""
Database migration script to add vlm_image_path column to pdf_extraction table
Run this script to update your existing database schema
"""

import sqlite3
import os

# Path to the database
db_path = os.path.join(os.path.dirname(__file__), 'pih_data.db')

def migrate_database():
    """Add vlm_image_path column to pdf_extraction table if it doesn't exist"""
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # Check if column already exists
        cursor.execute("PRAGMA table_info(pdf_extraction)")
        columns = [column[1] for column in cursor.fetchall()]
        
        if 'vlm_image_path' not in columns:
            print("Adding vlm_image_path column to pdf_extraction table...")
            cursor.execute("ALTER TABLE pdf_extraction ADD COLUMN vlm_image_path VARCHAR(512)")
            conn.commit()
            print("✓ Successfully added vlm_image_path column")
        else:
            print("✓ vlm_image_path column already exists")
        
        conn.close()
        print("\nDatabase migration completed successfully!")
        
    except Exception as e:
        print(f"Error during migration: {str(e)}")
        raise

if __name__ == "__main__":
    print("Starting database migration...")
    migrate_database()
