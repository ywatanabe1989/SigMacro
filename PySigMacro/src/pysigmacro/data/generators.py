#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-08 23:18:31 (ywatanabe)"
# File: /home/ywatanabe/proj/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/data/generators.py

THIS_FILE = "/home/ywatanabe/proj/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/data/generators.py"

"""
Data generation utilities for SigmaPlot graphs.
"""

import os
import csv
import math
import tempfile
import random
from datetime import datetime
from typing import List, Dict, Tuple, Union, Optional, Any, Callable

try:
    import numpy as np

    NUMPY_AVAILABLE = True
except ImportError:
    NUMPY_AVAILABLE = False


def create_sample_data(
    data_type: str = "sine",
    num_points: int = 100,
    output_path: Optional[str] = None,
) -> Tuple[str, List, List]:
    """
    Create sample data for testing SigmaPlot graphing functionality.

    Args:
        data_type (str): Type of data to generate ('sine', 'cosine', 'linear', 'exponential')
        num_points (int): Number of data points to generate
        output_path (str, optional): Path to save CSV file. If None, creates temp file.

    Returns:
        Tuple[str, List, List]: (csv_path, x_values, y_values)
    """
    # Generate x values
    x_values = [i * (10 / num_points) for i in range(num_points)]

    # Generate y values based on data_type
    if data_type.lower() == "sine":
        y_values = [math.sin(x) for x in x_values]
        title = "Sine Wave"
    elif data_type.lower() == "cosine":
        y_values = [math.cos(x) for x in x_values]
        title = "Cosine Wave"
    elif data_type.lower() == "linear":
        y_values = [2 * x + 1 for x in x_values]
        title = "Linear Function (2x + 1)"
    elif data_type.lower() == "exponential":
        y_values = [math.exp(x / 5) for x in x_values]
        title = "Exponential Function"
    elif data_type.lower() == "gaussian":
        y_values = [math.exp(-((x - 5) ** 2) / 2) for x in x_values]
        title = "Gaussian Function"
    else:
        # Default to sine
        y_values = [math.sin(x) for x in x_values]
        title = "Sine Wave"

    # Determine output path
    if output_path is None:
        temp_dir = tempfile.gettempdir()
        output_path = os.path.join(temp_dir, f"sigmaplot_{data_type}_data.csv")

    # Write to CSV file
    with open(output_path, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["X", "Y"])
        for i in range(len(x_values)):
            writer.writerow([x_values[i], y_values[i]])

    print(f"Sample {title} data created at: {output_path}")
    return output_path, x_values, y_values


def prepare_multi_series_data(
    data_series: Dict[str, Tuple[List, List]],
    output_path: Optional[str] = None,
) -> str:
    """
    Prepare multi-series data for graphs with multiple data series.

    Args:
        data_series (Dict[str, Tuple[List, List]]): Dictionary of series_name: (x_values, y_values)
        output_path (str, optional): Path to save CSV file. If None, creates temp file.

    Returns:
        str: Path to the CSV file with prepared data
    """
    # Determine output path
    if output_path is None:
        temp_dir = tempfile.gettempdir()
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(
            temp_dir, f"sigmaplot_multi_series_{timestamp}.csv"
        )

    # Get all series names
    series_names = list(data_series.keys())

    # Open CSV file for writing
    with open(output_path, "w", newline="") as f:
        writer = csv.writer(f)

        # Write header with series names
        header = []
        for name in series_names:
            header.extend([f"{name}_X", f"{name}_Y"])
        writer.writerow(header)

        # Find the max number of points across all series
        max_points = max([len(data_series[name][0]) for name in series_names])

        # Write data rows
        for i in range(max_points):
            row = []
            for name in series_names:
                x_values, y_values = data_series[name]
                if i < len(x_values):
                    row.extend([x_values[i], y_values[i]])
                else:
                    row.extend(["", ""])
            writer.writerow(row)

    print(f"Multi-series data file created at: {output_path}")
    return output_path


def create_scatter_data(
    num_points: int = 50,
    correlation: float = 0.7,
    output_path: Optional[str] = None,
) -> Tuple[str, List, List]:
    """
    Create scatter plot data with specified correlation.

    Args:
        num_points (int): Number of data points to generate
        correlation (float): Desired correlation between x and y (-1 to 1)
        output_path (str, optional): Path to save CSV file. If None, creates temp file.

    Returns:
        Tuple[str, List, List]: (csv_path, x_values, y_values)
    """
    # Generate correlated random data
    if NUMPY_AVAILABLE:
        # Set random seed for reproducibility
        np.random.seed(42)

        # Generate x values
        x_values = np.random.normal(size=num_points)

        # Generate y values with desired correlation
        y_values = correlation * x_values + np.sqrt(
            1 - correlation**2
        ) * np.random.normal(size=num_points)

        # Convert to Python lists
        x_values = x_values.tolist()
        y_values = y_values.tolist()
    else:
        # Fallback if NumPy is not available
        random.seed(42)

        # Generate random x values
        x_values = [random.normalvariate(0, 1) for _ in range(num_points)]

        # Simple approach to generate correlated y values
        y_values = [
            correlation * x
            + (1 - abs(correlation)) * random.normalvariate(0, 1)
            for x in x_values
        ]

    # Determine output path
    if output_path is None:
        temp_dir = tempfile.gettempdir()
        output_path = os.path.join(temp_dir, f"sigmaplot_scatter_data.csv")

    # Write to CSV file
    with open(output_path, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["X", "Y"])
        for i in range(len(x_values)):
            writer.writerow([x_values[i], y_values[i]])

    print(
        f"Scatter plot data (correlation={correlation}) created at: {output_path}"
    )
    return output_path, x_values, y_values


def create_time_series_data(
    start_date: str = "2023-01-01",
    days: int = 30,
    trend: float = 0.1,
    seasonality: float = 0.5,
    noise: float = 0.2,
    output_path: Optional[str] = None,
) -> Tuple[str, List, List]:
    """
    Create time series data with trend and seasonality.

    Args:
        start_date (str): Starting date in 'YYYY-MM-DD' format
        days (int): Number of days to generate
        trend (float): Trend coefficient (slope)
        seasonality (float): Seasonality amplitude
        noise (float): Random noise amplitude
        output_path (str, optional): Path to save CSV file. If None, creates temp file.

    Returns:
        Tuple[str, List, List]: (csv_path, dates, values)
    """
    from datetime import datetime, timedelta

    # Parse start date
    start = datetime.strptime(start_date, "%Y-%m-%d")

    # Generate dates
    dates = [
        (start + timedelta(days=i)).strftime("%Y-%m-%d") for i in range(days)
    ]

    # Generate values with trend, seasonality, and noise
    values = []
    for i in range(days):
        # Trend component
        trend_component = trend * i

        # Seasonality component (weekly pattern)
        seasonality_component = seasonality * math.sin(2 * math.pi * i / 7)

        # Random noise
        noise_component = noise * (random.random() * 2 - 1)

        # Combine components
        value = trend_component + seasonality_component + noise_component
        values.append(value)

    # Determine output path
    if output_path is None:
        temp_dir = tempfile.gettempdir()
        output_path = os.path.join(temp_dir, f"sigmaplot_time_series_data.csv")

    # Write to CSV file
    with open(output_path, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Date", "Value"])
        for i in range(len(dates)):
            writer.writerow([dates[i], values[i]])

    print(f"Time series data created at: {output_path}")
    return output_path, dates, values


def create_categorical_data(
    categories: Optional[List[str]] = None,
    values: Optional[List[float]] = None,
    output_path: Optional[str] = None,
) -> Tuple[str, List, List]:
    """
    Create categorical data for bar charts, pie charts, etc.

    Args:
        categories (List[str], optional): List of category names
        values (List[float], optional): List of values for each category
        output_path (str, optional): Path to save CSV file. If None, creates temp file.

    Returns:
        Tuple[str, List, List]: (csv_path, categories, values)
    """
    # Default categories if not provided
    if categories is None:
        categories = [
            "Category A",
            "Category B",
            "Category C",
            "Category D",
            "Category E",
        ]

    # Default values if not provided
    if values is None:
        values = [random.uniform(1, 10) for _ in range(len(categories))]

    # Ensure lists have the same length
    if len(categories) != len(values):
        raise ValueError("Categories and values must have the same length")

    # Determine output path
    if output_path is None:
        temp_dir = tempfile.gettempdir()
        output_path = os.path.join(temp_dir, f"sigmaplot_categorical_data.csv")

    # Write to CSV file
    with open(output_path, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Category", "Value"])
        for i in range(len(categories)):
            writer.writerow([categories[i], values[i]])

    print(f"Categorical data created at: {output_path}")
    return output_path, categories, values


def create_3d_data(
    nx: int = 20,
    ny: int = 20,
    function_type: str = "peaks",
    output_path: Optional[str] = None,
) -> Tuple[str, List, List, List]:
    """
    Create 3D data for surface and contour plots.

    Args:
        nx (int): Number of points in x dimension
        ny (int): Number of points in y dimension
        function_type (str): Type of function ('peaks', 'sincos', 'volcano')
        output_path (str, optional): Path to save CSV file. If None, creates temp file.

    Returns:
        Tuple[str, List, List, List]: (csv_path, x_grid, y_grid, z_values)
    """
    # Check if NumPy is available for more efficient grid generation
    if not NUMPY_AVAILABLE:
        print(
            "Warning: NumPy not available, using slower implementation for 3D data"
        )

        # Generate x and y grids
        x_values = [i / (nx - 1) * 4 - 2 for i in range(nx)]
        y_values = [i / (ny - 1) * 4 - 2 for i in range(ny)]

        # Generate all combinations for the grid
        x_grid = []
        y_grid = []
        for y in y_values:
            for x in x_values:
                x_grid.append(x)
                y_grid.append(y)

        # Generate z values based on function type
        z_values = []
        for i in range(len(x_grid)):
            x, y = x_grid[i], y_grid[i]

            if function_type == "peaks":
                z = (
                    3 * (1 - x) ** 2 * math.exp(-(x**2) - (y + 1) ** 2)
                    - 10 * (x / 5 - x**3 - y**5) * math.exp(-(x**2) - y**2)
                    - 1 / 3 * math.exp(-((x + 1) ** 2) - y**2)
                )
            elif function_type == "sincos":
                z = math.sin(3 * x) * math.cos(3 * y)
            elif function_type == "volcano":
                r = math.sqrt(x**2 + y**2)
                z = 2 * math.exp(-(r**2)) if r > 0 else 2
            else:
                # Default to peaks function
                z = (
                    3 * (1 - x) ** 2 * math.exp(-(x**2) - (y + 1) ** 2)
                    - 10 * (x / 5 - x**3 - y**5) * math.exp(-(x**2) - y**2)
                    - 1 / 3 * math.exp(-((x + 1) ** 2) - y**2)
                )

            z_values.append(z)
    else:
        # NumPy implementation (more efficient)
        x = np.linspace(-2, 2, nx)
        y = np.linspace(-2, 2, ny)

        # Create coordinate grid
        X, Y = np.meshgrid(x, y)

        # Calculate Z values based on function type
        if function_type == "peaks":
            Z = (
                3 * (1 - X) ** 2 * np.exp(-(X**2) - (Y + 1) ** 2)
                - 10 * (X / 5 - X**3 - Y**5) * np.exp(-(X**2) - Y**2)
                - 1 / 3 * np.exp(-((X + 1) ** 2) - Y**2)
            )
        elif function_type == "sincos":
            Z = np.sin(3 * X) * np.cos(3 * Y)
        elif function_type == "volcano":
            R = np.sqrt(X**2 + Y**2)
            Z = 2 * np.exp(-(R**2))
        else:
            # Default to peaks function
            Z = (
                3 * (1 - X) ** 2 * np.exp(-(X**2) - (Y + 1) ** 2)
                - 10 * (X / 5 - X**3 - Y**5) * np.exp(-(X**2) - Y**2)
                - 1 / 3 * np.exp(-((X + 1) ** 2) - Y**2)
            )

        # Reshape for output
        x_grid = X.flatten().tolist()
        y_grid = Y.flatten().tolist()
        z_values = Z.flatten().tolist()

    # Determine output path
    if output_path is None:
        temp_dir = tempfile.gettempdir()
        output_path = os.path.join(
            temp_dir, f"sigmaplot_3d_data_{function_type}.csv"
        )

    # Write to CSV file
    with open(output_path, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["X", "Y", "Z"])
        for i in range(len(x_grid)):
            writer.writerow([x_grid[i], y_grid[i], z_values[i]])

    print(f"3D data ({function_type}) created at: {output_path}")
    return output_path, x_grid, y_grid, z_values

# EOF