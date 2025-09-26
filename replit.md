# PIH Quote Generator

## Overview
This is a Flask-based web application designed to process Modula VLM (Vertical Lift Module) PDF quotes and generate professional Word documents and Excel spreadsheets. The application extracts pricing information, line items, and customer details from uploaded PDFs and creates formatted proposals with integrated marketing materials.

## System Architecture
The application follows a traditional Flask MVC architecture with the following layers:
- **Frontend**: HTML templates with Bootstrap for responsive UI
- **Backend**: Flask web framework with SQLAlchemy ORM
- **Database**: SQLite (development) with PostgreSQL support for production
- **File Processing**: PDF extraction using pdfplumber and document generation using python-docx
- **Image Processing**: PDF2Image and PIL for image extraction and manipulation

## Key Components

### Core Application (`app.py`)
- Flask application with route handlers for file upload, processing, review, and download
- Database operations for storing extracted data and line items
- PDF to Word/Excel conversion orchestration
- File serving and preview functionality

### Database Models (`models.py`)
- **PDFExtraction**: Stores metadata about uploaded PDFs and extracted pricing information
- **LineItem**: Stores individual line items with categories, pricing, and display options
- Relationships between extractions and their associated line items

### PDF Processing (`utils/pdf_processor.py`)
- Text extraction from PDF files using pdfplumber
- Line item parsing and categorization
- Word document generation with custom formatting
- Image extraction from PDF pages for document embedding
- Marketing PDF integration

### Image Processing (`utils/image_processor.py`)
- Logo processing and resizing
- PDF page extraction as images
- Image optimization for document embedding

### PDF Utilities (`utils/pdf_append.py`)
- Marketing PDF appending functionality
- PDF manipulation using pypdf

## Data Flow
1. **Upload**: User uploads Modula VLM PDF quote
2. **Extraction**: System extracts text and parses line items, pricing, and customer information
3. **Processing**: Data is structured and stored in database with proper categorization
4. **Review**: User can review and modify extracted data through web interface
5. **Generation**: System generates Word document with formatted tables, images, and marketing materials
6. **Download**: User can download generated documents or preview them in browser

## External Dependencies
- **PDF Processing**: pdfplumber, pypdf, pdf2image
- **Document Generation**: python-docx, openpyxl
- **Image Processing**: Pillow, opencv-python
- **Web Framework**: Flask, Flask-SQLAlchemy
- **Database**: SQLAlchemy with PostgreSQL/SQLite support
- **System Dependencies**: poppler-utils, libreoffice for PDF conversion

## Deployment Strategy
- **Development**: SQLite database with local file storage
- **Production**: PostgreSQL database with Gunicorn WSGI server
- **Replit Configuration**: Uses autoscale deployment target
- **File Handling**: Temporary file processing with cleanup
- **Static Assets**: Marketing PDFs and images stored in static/assets

## Recent Changes
- June 24, 2025: Added VLM image library with visual selection interface for cover page images
- June 24, 2025: Created 6 VLM model images (ML15-ML100) with blur/selection effects
- June 24, 2025: Implemented CSS styling for image gallery with hover and selection states
- June 24, 2025: Added salesperson dropdown with automated contact information population
- June 24, 2025: Implemented contact information for 7 salespeople (Josh, Noah, Tyler, Ivan, Matthew, Mike, John)
- June 24, 2025: Fixed WD_BREAK import error in PDF processor
- June 24, 2025: Updated download functionality to provide .docx files instead of PDFs
- June 24, 2025: Fixed marketing content integration with proper page breaks
- June 24, 2025: Resolved image processing issues without numpy dependencies
- June 24, 2025: Marketing PDF now properly converts to PNG images in Word documents

## Changelog
- June 19, 2025: Initial setup

## User Preferences
Preferred communication style: Simple, everyday language.