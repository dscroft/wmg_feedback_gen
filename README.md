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

See the `examples/` directory for usage examples.

## Excel File Format

For the standard default settings your Excel file should contain columns with the following data:
- STUDENTID (7-digit number)
- Any additional data needed for your template

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
