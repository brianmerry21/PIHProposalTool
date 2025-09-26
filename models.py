from datetime import datetime
from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()

class PDFExtraction(db.Model):
    """Store information about an uploaded PDF and its extracted data"""
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(255), nullable=False)
    upload_date = db.Column(db.DateTime, default=datetime.utcnow)
    original_path = db.Column(db.String(512))
    
    # Basic extraction data
    customer_name = db.Column(db.String(255))
    proposal_number = db.Column(db.String(50))
    proposal_date = db.Column(db.String(50))
    location = db.Column(db.String(255))
    logo_path = db.Column(db.String(512))  # Path to uploaded customer logo
    # excel_path = db.Column(db.String(512))  # Path to generated/reuploaded Excel file - temporarily commented out
    
    # VLM specifications
    vlm_model = db.Column(db.String(50))
    vlm_height = db.Column(db.String(50))
    
    # Extracted values
    base_price = db.Column(db.Float)
    software_price = db.Column(db.Float)
    installation_price = db.Column(db.Float)
    seismic_price = db.Column(db.Float)
    freight_price = db.Column(db.Float)
    total_price = db.Column(db.Float)
    
    # Tray information
    tray_quantity = db.Column(db.Integer)
    tray_width = db.Column(db.String(20))
    tray_depth = db.Column(db.String(20))
    tray_height = db.Column(db.String(20))
    
    # Related items
    line_items = db.relationship("LineItem", back_populates="extraction", cascade="all, delete-orphan")
    
    def __repr__(self):
        return f"<PDFExtraction {self.filename} - {self.customer_name}>"

class LineItem(db.Model):
    """Store individual line items from the extraction"""
    id = db.Column(db.Integer, primary_key=True)
    extraction_id = db.Column(db.Integer, db.ForeignKey('pdf_extraction.id'))
    extraction = db.relationship("PDFExtraction", back_populates="line_items")
    
    # Item details
    category = db.Column(db.String(50))  # e.g., "VLM", "Software", "Installation", etc.
    description = db.Column(db.String(255))
    price_each = db.Column(db.Float)
    quantity = db.Column(db.Float)
    price_total = db.Column(db.Float)
    is_included = db.Column(db.Boolean, default=False)  # For items marked as "Included"
    margin_percent = db.Column(db.Float, default=0.0)   # Stores the margin percentage
    
    # For user interface
    display_order = db.Column(db.Integer)
    is_section_header = db.Column(db.Boolean, default=False)  # For section headers like "Software" or "Installation and Warranty"
    
    # For optional items not included in base price
    is_optional = db.Column(db.Boolean, default=False)  # For items listed in the "Options Not Included" section
    is_recommended = db.Column(db.Boolean, default=False)  # For optional items that are recommended by PIH
    is_tbd_price = db.Column(db.Boolean, default=False)  # For optional items with TBD pricing
    
    # User modifications
    user_modified = db.Column(db.Boolean, default=False)
    original_price_each = db.Column(db.Float)
    original_quantity = db.Column(db.Float)
    original_price_total = db.Column(db.Float)
    original_is_included = db.Column(db.Boolean)
    
    def __repr__(self):
        # Handle the case where price_total might be None
        if self.is_included:
            price_display = 'Included'
        else:
            price_display = f"{self.price_total or 0.0}"
        return f"<LineItem {self.description} - {price_display}>"