"""
This script resets the database and ensures the default prices are correctly set.
Use this if you want to start fresh with the correct pricing structure.
"""
import os
import sys
import logging
from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy.orm import DeclarativeBase
from datetime import datetime

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Create minimal app context
class Base(DeclarativeBase):
    pass

db = SQLAlchemy(model_class=Base)
app = Flask(__name__)
app.config["SQLALCHEMY_DATABASE_URI"] = os.environ.get("DATABASE_URL")
app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {
    "pool_recycle": 300,
    "pool_pre_ping": True,
}
db.init_app(app)

# Import models after initializing app/db
with app.app_context():
    from models import PDFExtraction, LineItem

def reset_database():
    """Reset the database and create tables"""
    try:
        with app.app_context():
            logger.info("Dropping all tables...")
            db.drop_all()
            
            logger.info("Creating all tables...")
            db.create_all()
            
            logger.info("Database tables created successfully")
        return True
    except Exception as e:
        logger.error(f"Error resetting database: {str(e)}")
        return False

if __name__ == "__main__":
    reset_database()