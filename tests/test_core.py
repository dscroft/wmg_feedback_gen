"""Tests for core functionality."""

import pytest
from wmg_feedback_gen.core import (
    category,
    find_columns,
    validate_row_data,
    gen_filename,
    default_validators
)


class TestCategory:
    """Test the category function."""

    def test_outstanding(self):
        assert category(85) == "OUTSTANDING"
        assert category(80) == "OUTSTANDING"

    def test_distinction(self):
        assert category(75) == "DISTINCTION"
        assert category(70) == "DISTINCTION"

    def test_good(self):
        assert category(65) == "GOOD"
        assert category(60) == "GOOD"

    def test_pass(self):
        assert category(55) == "PASS"
        assert category(50) == "PASS"

    def test_marginal(self):
        assert category(45) == "MARGINAL"
        assert category(40) == "MARGINAL"

    def test_fail(self):
        assert category(35) == "FAIL"
        assert category(0) == "FAIL"

    def test_string_input(self):
        assert category("75") == "DISTINCTION"


class TestValidateRowData:
    """Test the validate_row_data function."""

    def test_valid_student_id(self):
        row_data = {"STUDENTID": "1234567"}
        assert validate_row_data(row_data) == True

    def test_invalid_student_id(self):
        row_data = {"STUDENTID": "123"}
        assert validate_row_data(row_data) == False

    def test_missing_validator(self):
        row_data = {"OTHER": "value"}
        with pytest.raises(ValueError, match="Validator 'STUDENTID' not found"):
            validate_row_data(row_data)


class TestGenFilename:
    """Test the gen_filename function."""

    def test_simple_template(self):
        template = "feedback_{{STUDENTID}}.docx"
        row_data = {"STUDENTID": "1234567"}
        result = gen_filename(template, row_data)
        assert result == "feedback_1234567.docx"

    def test_complex_template(self):
        template = "{{COMPONENT}}_feedback_{{STUDENTID}}_{{MARKER}}.docx"
        row_data = {
            "STUDENTID": "1234567",
            "COMPONENT": "PMA",
            "MARKER": "DrCroft"
        }
        result = gen_filename(template, row_data)
        assert result == "PMA_feedback_1234567_DrCroft.docx"
