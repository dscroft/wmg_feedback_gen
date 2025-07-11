# WMG Feedback Generator

A Python library for generating student feedback documents from Excel data and Word templates.

## Features

- Generate personalized feedback documents from Excel spreadsheets
- Support for Word template-based document generation
- Automatic marking categorization (Outstanding, Distinction, Good, Pass, Marginal, Fail)
- Flexible column mapping and data validation
- Batch processing of multiple students

## Installation

### From Source

```bash
git clone https://github.com/davidcroft/wmg-feedback-gen.git
cd wmg-feedback-gen
pip install -e .
```

### Dependencies

```bash
pip install -r requirements.txt
```

## Quick Start

### Using the DocumentGenerator Class

```python
from wmg_feedback_gen import DocumentGenerator

# Create a document generator
generator = DocumentGenerator(
    xlsx_filename="student_marks.xlsx",
    template_filename="feedback_template.docx",
    worksheet="marks",
    output_path="generated_feedback"
)

# Generate feedback documents
generated_files = generator.generate_feedback_documents(
    marker="Dr David Croft",
    component="PMA"
)

print(f"Generated {len(generated_files)} feedback documents")
```

### Using Core Functions

```python
from wmg_feedback_gen import find_columns, process_to_dicts, category
import openpyxl

# Open Excel file
workbook = openpyxl.load_workbook("marks.xlsx", data_only=True)
sheet = workbook["marks"]

# Find columns automatically
expected_vars = ["STUDENTID", "FEEDBACK", "LO2", "LO3", "LO4", "LO5"]
columns = find_columns(sheet, expected_vars)

# Process rows
for row_data in process_to_dicts(sheet, expected_vars):
    print(f"Student {row_data['STUDENTID']}: {category(row_data['LO2'])}")
```

## Project Structure

```
wmg_feedback_gen/
├── src/
│   └── wmg_feedback_gen/
│       ├── __init__.py
│       ├── core.py                 # Core processing functions
│       └── document_generator.py   # Document generation
├── tests/
│   ├── test_core.py
│   └── test_document_generator.py
├── examples/
│   ├── generate_feedback.py       # Example usage
│   ├── legacy_pma_feedback.py     # Legacy compatibility
│   ├── demo_marks.xlsx            # Sample data
│   └── demo_feedback_template.docx # Sample template
├── pyproject.toml                 # Modern packaging
├── requirements.txt               # Runtime dependencies
└── README.md
```

## Excel File Format

Your Excel file should contain columns with the following data:
- Student ID (7-digit number)
- Feedback text
- Learning outcome marks (LO2, LO3, LO4, LO5)
- Any additional data needed for your template

## Word Template Format

Your Word template should contain placeholder text that will be replaced:
- `{{STUDENTID}}` - Student ID
- `{{COMPONENT}}` - Component name
- `{{MARKER}}` - Marker name
- `{{LO2}}`, `{{LO3}}`, etc. - Learning outcome categories
- `{{FEEDBACK}}` - Feedback text

## Development

### Install Development Dependencies

```bash
pip install -r requirements-dev.txt
```

### Run Tests

```bash
pytest
```

### Code Formatting

```bash
black src/ tests/
```

### Type Checking

```bash
mypy src/
```

## License

MIT License - see LICENSE file for details.
