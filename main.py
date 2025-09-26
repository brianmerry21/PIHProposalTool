import logging
import os
from app import app, db
from sqlalchemy import create_engine, text

# Configure logging
logging.basicConfig(level=logging.DEBUG)

# Ensure DATABASE_URL is set 
if not os.environ.get('DATABASE_URL'):
    os.environ['DATABASE_URL'] = os.environ.get('DATABASE_URL', 'sqlite:///pih_data.db')

def update_database_schema():
    """
    Add any missing columns to the database tables
    """
    database_url = os.environ.get("DATABASE_URL")
    try:
        # Create engine and connect to database
        engine = create_engine(database_url)
        with engine.connect() as conn:
            # Check if logo_path column exists
            check_sql = text(
                "SELECT column_name FROM information_schema.columns "
                "WHERE table_name = 'pdf_extraction' AND column_name = 'logo_path'"
            )
            result = conn.execute(check_sql)
            rows = result.fetchall()
            
            if not rows:
                # Add the logo_path column if it doesn't exist
                alter_sql = text(
                    "ALTER TABLE pdf_extraction ADD COLUMN logo_path VARCHAR(512)"
                )
                conn.execute(alter_sql)
                conn.commit()
                logging.info("Added logo_path column to pdf_extraction table")
    except Exception as e:
        logging.error(f"Error updating database schema: {str(e)}")
        # Continue with application startup even if schema update fails
        pass

# Create database tables
with app.app_context():
    db.create_all()
    logging.info("Database tables created successfully")
    # Update schema with any missing columns
    update_database_schema()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=True)
