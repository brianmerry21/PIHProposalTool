import logging
import os
from app import app, db
from sqlalchemy import create_engine, text

# ------------------------------------------------------------
# Configure logging
# ------------------------------------------------------------
logging.basicConfig(level=logging.DEBUG, format="%(asctime)s [%(levelname)s] %(message)s")

# ------------------------------------------------------------
#  Ensure DATABASE_URL is set (default: SQLite)
# ------------------------------------------------------------
if not os.environ.get('DATABASE_URL'):
    os.environ['DATABASE_URL'] = 'sqlite:///pih_data.db'

# ------------------------------------------------------------
#  Function to check and update database schema dynamically
# ------------------------------------------------------------
def update_database_schema():
    """
    Add any missing columns to the database tables for both SQLite and PostgreSQL/MySQL.
    """

    database_url = os.environ.get("DATABASE_URL")
    logging.info(f"Using database: {database_url}")

    # Define all expected columns for the pdf_extraction table
    expected_columns = {
        'logo_path': 'TEXT',
        'vlm_image_path': 'TEXT',
        'contact_name': 'TEXT',
        'contact_email': 'TEXT',
        'contact_phone': 'TEXT',
        'contact_office': 'TEXT'
    }

    try:
        # Create engine and connect
        engine = create_engine(database_url)
        with engine.connect() as conn:
            # ------------------------------------------------------------
            # 🔹 Case 1: SQLite database
            # ------------------------------------------------------------
            if database_url.startswith("sqlite"):
                result = conn.execute(text("PRAGMA table_info(pdf_extraction);"))
                columns = [row[1] for row in result]

                for col_name, col_type in expected_columns.items():
                    if col_name not in columns:
                        conn.execute(text(f"ALTER TABLE pdf_extraction ADD COLUMN {col_name} {col_type};"))
                        logging.info(f" Added missing column '{col_name}' ({col_type}) to pdf_extraction (SQLite)")
                    else:
                        logging.info(f"Column '{col_name}' already exists (SQLite)")

            # ------------------------------------------------------------
            # 🔹 Case 2: PostgreSQL / MySQL database
            # ------------------------------------------------------------
            else:
                for col_name, col_type in expected_columns.items():
                    check_sql = text(f"""
                        SELECT column_name 
                        FROM information_schema.columns
                        WHERE table_name = 'pdf_extraction' AND column_name = '{col_name}'
                    """)
                    result = conn.execute(check_sql)
                    rows = result.fetchall()

                    if not rows:
                        conn.execute(text(f"ALTER TABLE pdf_extraction ADD COLUMN {col_name} {col_type};"))
                        conn.commit()
                        logging.info(f" Added missing column '{col_name}' ({col_type}) to pdf_extraction (Postgres/MySQL)")
                    else:
                        logging.info(f" Column '{col_name}' already exists (Postgres/MySQL)")

    except Exception as e:
        logging.error(f"Error updating database schema: {str(e)}")
        pass  # continue with app startup


# ------------------------------------------------------------
# ✅ Initialize database and update schema
# ------------------------------------------------------------
with app.app_context():
    db.create_all()
    logging.info(" Database tables created successfully")
    update_database_schema()

# ------------------------------------------------------------
# ✅ Run the Flask application
# ------------------------------------------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=True)