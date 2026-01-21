#!/usr/bin/env python3
"""
Example script demonstrating how to use the wmg_feedback_gen library.

This script shows how to generate feedback documents using the new library structure.

This example uses the default validators and post-processing steps as suitable for 
the default WMG feedback templates. Namely, it assumes the presence of a STUDENTID column
in the provided Excel worksheet which contains a 7 digit student identifier.
It also highlights the appropriate cells in the feedback table based on the last word in the 
category column, i.e. "GOOD" will cause the "Good pass" cell to be highlighted.
"""

import sys
import os

# Add the src directory to the Python path for development
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

import wmg_feedback_gen
from pathlib import Path

if __name__ == "__main__":
    """Main function demonstrating library usage."""
    
    # Configuration
    xlsx_filename = Path(__file__).parent / "demo_marks.xlsx"
    template_filename = Path(__file__).parent / "demo_feedback_template.docx"
    worksheet = "marks"  # Adjust based on your Excel file
    output_filename = Path(__file__).parent / "output" / "feedback_{{STUDENTID}}.docx"

    # Generate feedback using default validators and post-processing
    print( "Generate feedback documents with default validators and post-processing..." )
    wmg_feedback_gen.generate(
        xlsx_filename=xlsx_filename,
        template_filename=template_filename,
        worksheet=worksheet,
        output_filename=output_filename)

    # Example without the default validators and post-processing
    print( "Generate feedback documents without validators and post-processing..." )
    wmg_feedback_gen.generate(
        xlsx_filename=xlsx_filename,
        template_filename=template_filename,
        worksheet=worksheet,
        output_filename=Path(__file__).parent / "output" / "alternative_{{STUDENTID}}.docx",
        validators=None,
        post_processing=None)
  