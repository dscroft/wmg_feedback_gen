#!/usr/bin/env python3
"""
Example script demonstrating how to use the wmg_feedback_gen library.

This script shows how to generate feedback documents using the new library structure.
"""

import sys
import os

# Add the src directory to the Python path for development
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

import wmg_feedback_gen

if __name__ == "__main__":
    """Main function demonstrating library usage."""
    
    # Configuration
    xlsx_filename = "demo_marks.xlsx"
    template_filename = "demo_feedback_template.docx"
    worksheet = "marks"  # Adjust based on your Excel file
    output_filename = "feedback/feedback_{{STUDENTID}}.docx"

    wmg_feedback_gen.generate(
        xlsx_filename=xlsx_filename,
        template_filename=template_filename,
        worksheet=worksheet,
        output_filename=output_filename)

    
  