#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 00:18:10 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/data/test_data_generation.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/data/test_data_generation.py"

import os
import csv
import math
import tempfile
from typing import Dict, List, Tuple, Optional, Any
import pytest

"""
Functionality:
* Tests data generation functions for SigmaPlot graphing
* Verifies correct creation of different data types (sine, cosine, etc.)
* Ensures generated data files have proper format and content
Input:
* None (creates test data internally)
Output:
* Test results for data generation validation
Prerequisites:
* pytest
"""

class TestDataGeneration:
    @pytest.fixture
    def sample_data_functions(self) -> Dict[str, Any]:
        """
        Import sample data functions for testing.

        Returns
        -------
        Dict[str, Any]
            Dictionary of data generation functions
        """
        try:
            from pysigmacro.data.generators import (
                create_sample_data,
                create_scatter_data,
                create_time_series_data,
                create_categorical_data,
                create_3d_data,
                prepare_multi_series_data
            )
            return {
                "create_sample_data": create_sample_data,
                "create_scatter_data": create_scatter_data,
                "create_time_series_data": create_time_series_data,
                "create_categorical_data": create_categorical_data,
                "create_3d_data": create_3d_data,
                "prepare_multi_series_data": prepare_multi_series_data
            }
        except ImportError:
            pytest.skip("Pysigmacro data generators not available")
            return {}

    def test_create_sample_data_sine(self, sample_data_functions: Dict[str, Any]) -> None:
        """
        Tests generation of sine wave sample data.

        Parameters
        ----------
        sample_data_functions : Dict[str, Any]
            Dictionary of data generation functions
        """
        if not sample_data_functions:
            pytest.skip("Sample data functions not available")

        create_sample_data = sample_data_functions["create_sample_data"]

        # Create sample sine data
        temp_file = os.path.join(tempfile.gettempdir(), "test_sine_data.csv")
        try:
            csv_path, x_values, y_values = create_sample_data(
                data_type="sine",
                num_points=10,
                output_path=temp_file
            )

            # Verify the file was created
            assert os.path.exists(csv_path), "CSV file was not created"

            # Verify data content
            with open(csv_path, 'r') as f:
                reader = csv.reader(f)
                header = next(reader)
                assert header[0] == "X", "First column should be labeled 'X'"
                assert header[1] == "Y", "Second column should be labeled 'Y'"

                # Verify at least one row of data
                row = next(reader)
                assert len(row) == 2, "Data row should have two values"

                # Verify returned data matches
                assert len(x_values) == 10, "Should contain 10 x values"
                assert len(y_values) == 10, "Should contain 10 y values"

                # Verify sine wave characteristics (values between -1 and 1)
                assert all(-1.01 <= float(y) <= 1.01 for y in y_values), "Sine values should be between -1 and 1"
        finally:
            # Clean up
            if os.path.exists(temp_file):
                os.remove(temp_file)

    def test_create_scatter_data(self, sample_data_functions: Dict[str, Any]) -> None:
        """
        Tests generation of scatter plot data.

        Parameters
        ----------
        sample_data_functions : Dict[str, Any]
            Dictionary of data generation functions
        """
        if not sample_data_functions:
            pytest.skip("Sample data functions not available")

        create_scatter_data = sample_data_functions["create_scatter_data"]

        # Create sample scatter data
        temp_file = os.path.join(tempfile.gettempdir(), "test_scatter_data.csv")
        try:
            num_points = 20
            correlation = 0.8

            csv_path, x_values, y_values = create_scatter_data(
                num_points=num_points,
                correlation=correlation,
                output_path=temp_file
            )

            # Verify the file was created
            assert os.path.exists(csv_path), "CSV file was not created"

            # Verify data content
            with open(csv_path, 'r') as f:
                reader = csv.reader(f)
                header = next(reader)
                assert header[0] == "X", "First column should be labeled 'X'"
                assert header[1] == "Y", "Second column should be labeled 'Y'"

                # Count rows
                data_rows = list(reader)
                assert len(data_rows) == num_points, f"Should contain {num_points} data rows"

            # Verify returned data matches requested size
            assert len(x_values) == num_points, f"Should contain {num_points} x values"
            assert len(y_values) == num_points, f"Should contain {num_points} y values"

            # Calculate observed correlation (basic)
            if len(x_values) >= 2:
                x_mean = sum(x_values) / len(x_values)
                y_mean = sum(y_values) / len(y_values)

                numerator = sum((x - x_mean) * (y - y_mean) for x, y in zip(x_values, y_values))
                denominator_x = sum((x - x_mean) ** 2 for x in x_values)
                denominator_y = sum((y - y_mean) ** 2 for y in y_values)

                if denominator_x > 0 and denominator_y > 0:
                    observed_correlation = numerator / (math.sqrt(denominator_x) * math.sqrt(denominator_y))

                    # Correlation should be approximately the requested value
                    # Allow some deviation due to random generation
                    assert abs(observed_correlation - correlation) < 0.3, \
                        f"Observed correlation ({observed_correlation}) differs too much from requested ({correlation})"
        finally:
            # Clean up
            if os.path.exists(temp_file):
                os.remove(temp_file)

    def test_categorical_data(self, sample_data_functions: Dict[str, Any]) -> None:
        """
        Tests generation of categorical data.

        Parameters
        ----------
        sample_data_functions : Dict[str, Any]
            Dictionary of data generation functions
        """
        if not sample_data_functions:
            pytest.skip("Sample data functions not available")

        create_categorical_data = sample_data_functions["create_categorical_data"]

        # Create sample categorical data
        temp_file = os.path.join(tempfile.gettempdir(), "test_categorical_data.csv")
        try:
            test_categories = ["Cat1", "Cat2", "Cat3"]
            test_values = [5.0, 10.0, 15.0]

            csv_path, categories, values = create_categorical_data(
                categories=test_categories,
                values=test_values,
                output_path=temp_file
            )

            # Verify the file was created
            assert os.path.exists(csv_path), "CSV file was not created"

            # Verify data content
            with open(csv_path, 'r') as f:
                reader = csv.reader(f)
                header = next(reader)
                assert header[0] == "Category", "First column should be labeled 'Category'"
                assert header[1] == "Value", "Second column should be labeled 'Value'"

                # Verify categories and values
                rows = list(reader)
                assert len(rows) == len(test_categories), f"Should contain {len(test_categories)} data rows"

                for i, row in enumerate(rows):
                    assert row[0] == test_categories[i], f"Category at row {i} should be {test_categories[i]}"
                    assert float(row[1]) == test_values[i], f"Value at row {i} should be {test_values[i]}"

            # Verify returned data matches input
            assert categories == test_categories, "Returned categories should match input"
            assert values == test_values, "Returned values should match input"
        finally:
            # Clean up
            if os.path.exists(temp_file):
                os.remove(temp_file)

    def test_3d_data(self, sample_data_functions: Dict[str, Any]) -> None:
        """
        Tests generation of 3D data.

        Parameters
        ----------
        sample_data_functions : Dict[str, Any]
            Dictionary of data generation functions
        """
        if not sample_data_functions:
            pytest.skip("Sample data functions not available")

        create_3d_data = sample_data_functions["create_3d_data"]

        # Create sample 3D data
        temp_file = os.path.join(tempfile.gettempdir(), "test_3d_data.csv")
        try:
            nx, ny = 5, 5

            csv_path, x_grid, y_grid, z_values = create_3d_data(
                nx=nx,
                ny=ny,
                function_type="peaks",
                output_path=temp_file
            )

            # Verify the file was created
            assert os.path.exists(csv_path), "CSV file was not created"

            # Verify data content
            with open(csv_path, 'r') as f:
                reader = csv.reader(f)
                header = next(reader)
                assert header[0] == "X", "First column should be labeled 'X'"
                assert header[1] == "Y", "Second column should be labeled 'Y'"
                assert header[2] == "Z", "Third column should be labeled 'Z'"

                # Verify data size (nx * ny points)
                rows = list(reader)
                assert len(rows) == nx * ny, f"Should contain {nx * ny} data rows"

            # Verify returned data lengths
            assert len(x_grid) == nx * ny, f"x_grid should have {nx * ny} values"
            assert len(y_grid) == nx * ny, f"y_grid should have {nx * ny} values"
            assert len(z_values) == nx * ny, f"z_values should have {nx * ny} values"
        finally:
            # Clean up
            if os.path.exists(temp_file):
                os.remove(temp_file)

    def test_multi_series_data(self, sample_data_functions: Dict[str, Any]) -> None:
        """
        Tests preparation of multi-series data.

        Parameters
        ----------
        sample_data_functions : Dict[str, Any]
            Dictionary of data generation functions
        """
        if not sample_data_functions:
            pytest.skip("Sample data functions not available")

        prepare_multi_series_data = sample_data_functions["prepare_multi_series_data"]

        # Create sample multi-series data
        temp_file = os.path.join(tempfile.gettempdir(), "test_multi_series_data.csv")
        try:
            # Create data series
            data_series = {
                "Series1": ([1, 2, 3], [10, 20, 30]),
                "Series2": ([1, 2, 3], [5, 15, 25])
            }

            csv_path = prepare_multi_series_data(
                data_series=data_series,
                output_path=temp_file
            )

            # Verify the file was created
            assert os.path.exists(csv_path), "CSV file was not created"

            # Verify data content
            with open(csv_path, 'r') as f:
                reader = csv.reader(f)
                header = next(reader)
                assert len(header) == 4, "Should have 4 columns for 2 series"
                assert header[0] == "Series1_X", "First column should be labeled 'Series1_X'"
                assert header[1] == "Series1_Y", "Second column should be labeled 'Series1_Y'"
                assert header[2] == "Series2_X", "Third column should be labeled 'Series2_X'"
                assert header[3] == "Series2_Y", "Fourth column should be labeled 'Series2_Y'"

                # Verify data rows
                rows = list(reader)
                assert len(rows) == 3, "Should contain 3 data rows"

                # Verify first row values
                assert float(rows[0][0]) == 1, "Series1_X first value should be 1"
                assert float(rows[0][1]) == 10, "Series1_Y first value should be 10"
                assert float(rows[0][2]) == 1, "Series2_X first value should be 1"
                assert float(rows[0][3]) == 5, "Series2_Y first value should be 5"
        finally:
            # Clean up
            if os.path.exists(temp_file):
                os.remove(temp_file)

# EOF