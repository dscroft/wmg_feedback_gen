#!/usr/bin/env python3
"""
Example script demonstrating how to use the wmg_feedback_gen library.

This script shows how to customise the generating of feedback documents 
to handle different feedback forms
"""

import os
import sys
from docx import Document

# Add the src directory to the Python path for development
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

import wmg_feedback_gen
from wmg_feedback_gen.document_generator import highlight_cell

def custom_highlight(row_data, filename):
    """This alternative template is organised column-wise as opposed to the 
        standard row-wise WMG template and the grade bounds are different.
        This function transposes the table and then highlights the required cells.
    """

    docx = Document(filename)
    for table in docx.tables:
        # Transpose the table: rows become columns and columns become rows
        transposed_cells = list(zip(*[row.cells for row in table.rows]))
        for col in transposed_cells:
            if len(col) != 9:  # not the correct table
                continue

    def row_index(mark):
        if mark >= 80: return 1
        elif mark >= 70: return 2
        elif mark >= 60: return 3
        elif mark >= 50: return 4
        elif mark >= 40: return 5
        elif mark >= 30: return 6
        else: return 7

    highlight_cell(transposed_cells[1][row_index(row_data['LO2'])])
    highlight_cell(transposed_cells[2][row_index(row_data['LO3'])])
    highlight_cell(transposed_cells[3][row_index(row_data['LO4'])])
    highlight_cell(transposed_cells[4][row_index(row_data['LO5'])])

    docx.save(filename)


if __name__ == "__main__":
    """Main function demonstrating library usage."""
    
    # Configuration
    xlsx_filename = "demo_marks.xlsx"
    template_filename = "demo_alternative_template.docx"
    worksheet = "marks"  # Adjust based on your Excel file

    # Output filename can use Jinja2 templating, in this case we are using lastname 
    # but converted to lowercase
    output_filename = "feedback/feedback_{{NAME | lower}}.docx"

    # The default validators check for a valid student ID in the STUDENTID column
    # But we can replace or extend these if needed. For example here we add one
    # to check that they have a valid mark
    validators = {
        'total': lambda x: str(x).isdigit() and 0 < int(x) and int(x) <= 100
    }

    # Columns that need to be present in the Excel file
    # For the most part these will be automatically detected based on the 
    # template file, validators and output filename. But if you are using 
    # a custom post-processing function you may need to adapt them as is the case
    # here.
    expected_vars = ['LO2', 'LO3', 'LO4', 'LO5']

    wmg_feedback_gen.generate(
        xlsx_filename=xlsx_filename,
        template_filename=template_filename,
        worksheet=worksheet,
        output_filename=output_filename,
        validators=validators,
        post_processing=custom_highlight,
        expected_vars=expected_vars
    )

