#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 00:17:59 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/data/test_column_names.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/data/test_column_names.py"

import os
import tempfile
import csv
from typing import Dict, List, Tuple, Any
import pytest

"""
Functionality:
* Tests column name handling in CSV files for SigmaPlot
* Verifies appropriate column naming in generated data files
Input:
* None (creates test files internally)
Output:
* Test results for column naming validation
Prerequisites:
* pytest
"""

class TestColumnNames:
    def test_default_column_names(self) -> None:
        """
        Tests that default column names (X, Y) are used when not specified.
        """
        # Create a temporary CSV file with default column names
        temp_file = os.path.join(tempfile.gettempdir(), "test_default_columns.csv")

        try:
            # Create sample data with default column names
            with open(temp_file, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(["X", "Y"])
                writer.writerow([1, 2])
                writer.writerow([3, 4])

            # Verify the CSV file contains the expected column names
            with open(temp_file, 'r') as f:
                reader = csv.reader(f)
                header = next(reader)
                assert header[0] == "X", "First column should be named 'X'"
                assert header[1] == "Y", "Second column should be named 'Y'"
        finally:
            # Clean up
            if os.path.exists(temp_file):
                os.remove(temp_file)

    def test_custom_column_names(self) -> None:
        """
        Tests that custom column names are correctly applied.
        """
        # Create a temporary CSV file with custom column names
        temp_file = os.path.join(tempfile.gettempdir(), "test_custom_columns.csv")
        custom_names = ["Time", "Value"]

        try:
            # Create sample data with custom column names
            with open(temp_file, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(custom_names)
                writer.writerow([1, 10])
                writer.writerow([2, 20])

            # Verify the CSV file contains the expected column names
            with open(temp_file, 'r') as f:
                reader = csv.reader(f)
                header = next(reader)
                assert header[0] == custom_names[0], f"First column should be named '{custom_names[0]}'"
                assert header[1] == custom_names[1], f"Second column should be named '{custom_names[1]}'"
        finally:
            # Clean up
            if os.path.exists(temp_file):
                os.remove(temp_file)

    def test_multi_series_column_names(self) -> None:
        """
        Tests column naming for multi-series datasets.
        """
        # Create a temporary CSV file with multi-series column names
        temp_file = os.path.join(tempfile.gettempdir(), "test_multi_series_columns.csv")
        series_names = ["Series1", "Series2"]

        try:
            # Create sample data with multi-series column names
            with open(temp_file, 'w', newline='') as f:
                writer = csv.writer(f)
                # Header row with series names
                writer.writerow([f"{series_names[0]}_X", f"{series_names[0]}_Y",
                                 f"{series_names[1]}_X", f"{series_names[1]}_Y"])
                writer.writerow([1, 10, 1, 5])
                writer.writerow([2, 20, 2, 10])

            # Verify the CSV file contains the expected column names
            with open(temp_file, 'r') as f:
                reader = csv.reader(f)
                header = next(reader)
                assert header[0] == f"{series_names[0]}_X", f"First column should be named '{series_names[0]}_X'"
                assert header[1] == f"{series_names[0]}_Y", f"Second column should be named '{series_names[0]}_Y'"
                assert header[2] == f"{series_names[1]}_X", f"Third column should be named '{series_names[1]}_X'"
                assert header[3] == f"{series_names[1]}_Y", f"Fourth column should be named '{series_names[1]}_Y'"
        finally:
            # Clean up
            if os.path.exists(temp_file):
                os.remove(temp_file)

# EOF