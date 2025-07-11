"""Tests for document generator functionality."""

import pytest
from wmg_feedback_gen.document_generator import category


class TestDocumentGenerator:
    """Test the DocumentGenerator class and related functions."""

    def test_category_function(self):
        """Test the category function with various inputs."""
        assert category(85) == "OUTSTANDING"
        assert category(75) == "DISTINCTION"
        assert category(65) == "GOOD"
        assert category(55) == "PASS"
        assert category(45) == "MARGINAL"
        assert category(35) == "FAIL"
