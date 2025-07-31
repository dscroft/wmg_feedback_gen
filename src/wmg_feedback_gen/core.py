"""Core functionality for feedback generation.

This module contains the main functions for processing Excel data
and generating feedback documents.
"""

import re
import jinja2
import openpyxl
from docxtpl import DocxTemplate
import logging

default_validators = {
    'STUDENTID': lambda x: bool(re.match(r"[0-9]{7}", str(x).strip()))
}

def mark_category( mark ):
    mark = float(mark)
    if mark >= 80: return "OUTSTANDING"
    if mark >= 70: return "DISTINCTION"
    if mark >= 60: return "GOOD"
    if mark >= 50: return "PASS"
    if mark >= 40: return "MARGINAL"
    return "FAIL"

def find_columns(sheet, expected):
    """
    Find the columns in the given worksheet that match the expected variable names.

    Args:
        sheet: The openpyxl worksheet object to search.
        expected: A list of expected variable names to find in the sheet.

    Returns:
        A dictionary mapping each expected variable name to its column index (0-based).

    Details:
        Iterates through the rows of the worksheet and searches for the expected variable names.
        Stops searching once all expected variables are found.
        Cells are searched in row->column order, the first occurrence of each variable name
        is recorded, and the search stops once all expected variables are found.

    Warning:
        This function assumes that the expected variable names are unique within the worksheet.
        Or that the first occurrence of each variable name is the one we want.
        If a variable name appears multiple times, only the first occurrence will be recorded.
    """
    columns = {i: None for i in expected}

    for row in sheet.iter_rows(values_only=True):
        # if all columns are found, we can stop
        if None not in columns.values():
            break

        # check all cells in row to find matches for missing variables
        for idx, cell in enumerate(row):
            c = str(cell).strip()
            if c in columns and columns[c] is None:
                columns[c] = idx

    logging.debug(f"Found columns: {columns}")

    return columns


def process_to_dicts(sheet, expected, validators=default_validators):
    """
    Process worksheet rows and yield validated row data as dictionaries.

    Args:
        sheet: The openpyxl worksheet object to process.
        expected: List of expected column names.
        validators: Dictionary of validation functions for each column.

    Yields:
        Dictionary containing row data for each valid row.
    """
    columns = find_columns(sheet, expected)

    for row in sheet.iter_rows(values_only=True):
        row_data = {var: row[idx] for var, idx in columns.items() if idx is not None}
        if validate_row_data(row_data, validators):
            yield row_data
        else:
            logging.warning(f"Row data did not pass validation: {row_data}")


def extract_row_data(row, columns):
    """Extract row data based on column mapping."""
    return {var: row[idx] for var, idx in columns.items() if idx is not None}


def validate_row_data(row_data, validators=default_validators):
    """
    Validate row data using provided validators.

    Args:
        row_data: Dictionary containing row data.
        validators: Dictionary of validation functions.

    Returns:
        True if all validations pass, False otherwise.

    Raises:
        ValueError: If validator is missing or not callable.
    """
    for var, func in validators.items():
        if var not in row_data:
            raise ValueError(f"Validator '{var}' not found in columns.")

        if not callable(func):
            raise ValueError(f"Validator for '{var}' is not callable.")
        
        if func(row_data[var]) != True:
            return False
        
    return True


def gen_filename(template, row_data):
    """
    Generate filename using Jinja2 template and row data.

    Args:
        template: Jinja2 template string for filename.
        row_data: Dictionary containing template variables.

    Returns:
        Generated filename string.
    """
    template = jinja2.Template(template)
    return template.render(**row_data)
