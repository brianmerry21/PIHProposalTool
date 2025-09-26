import os
import sys
import logging
import argparse
from utils.excel_template_analyzer import analyze_template_and_save

# Set up logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

def main():
    parser = argparse.ArgumentParser(description='Analyze Excel template structure')
    parser.add_argument('template_path', help='Path to the Excel template file')
    parser.add_argument('--output', '-o', help='Path to save analysis output (default: template_analysis.txt)')
    
    args = parser.parse_args()
    
    template_path = args.template_path
    output_path = args.output or 'template_analysis.txt'
    
    try:
        logger.info(f"Analyzing Excel template: {template_path}")
        analyze_template_and_save(template_path, output_path)
        logger.info(f"Analysis saved to: {output_path}")
    except Exception as e:
        logger.error(f"Error: {str(e)}")
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main())