#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-08 23:07:16 (ywatanabe)"
# File: /home/ywatanabe/proj/SigmaPlot-v12.0-Pysigmacro/scripts/sigmaplot-py/src/sigmaplot/data/importers.py

THIS_FILE = "/home/ywatanabe/proj/SigmaPlot-v12.0-Pysigmacro/scripts/sigmaplot-py/src/sigmaplot/data/importers.py"

"""
Data import utilities for SigmaPlot graphs.
"""

import os
import csv
import json
import tempfile
from typing import List, Dict, Tuple, Union, Optional, Any

def import_csv_to_dict(csv_path: str,
                      has_header: bool = True,
                      delimiter: str = ',') -> Dict[str, List]:
    """
    Import CSV file into a dictionary of column data.

    Args:
        csv_path (str): Path to CSV file
        has_header (bool): Whether CSV has a header row
        delimiter (str): CSV delimiter character

    Returns:
        Dict[str, List]: Dictionary with column names as keys and data as lists
    """
    data = {}

    with open(csv_path, 'r', newline='') as f:
        reader = csv.reader(f, delimiter=delimiter)

        # Handle header row
        if has_header:
            headers = next(reader)
            for header in headers:
                data[header] = []
        else:
            # Create default column names
            row = next(reader)
            headers = [f"Column{i+1}" for i in range(len(row))]
            for i, header in enumerate(headers):
                data[header] = [row[i]]

        # Process data rows
        for row in reader:
            for i, value in enumerate(row):
                if i < len(headers):
                    try:
                        # Try to convert to numeric
                        numeric_value = float(value)
                        # If it's an integer, store as int
                        if numeric_value.is_integer():
                            numeric_value = int(numeric_value)
                        data[headers[i]].append(numeric_value)
                    except (ValueError, TypeError):
                        # If not numeric, store as string
                        data[headers[i]].append(value)

    return data

def import_excel_to_dict(excel_path: str,
                        sheet_name: Optional[str] = None,
                        has_header: bool = True) -> Dict[str, List]:
    """
    Import Excel file into a dictionary of column data.

    Args:
        excel_path (str): Path to Excel file
        sheet_name (str, optional): Sheet name to import (default: first sheet)
        has_header (bool): Whether Excel has a header row

    Returns:
        Dict[str, List]: Dictionary with column names as keys and data as lists
    """
    try:
        import pandas as pd

        # Read Excel file
        if sheet_name:
            df = pd.read_excel(excel_path, sheet_name=sheet_name)
        else:
            df = pd.read_excel(excel_path)

        # Convert to dictionary
        if has_header:
            data = {col: df[col].tolist() for col in df.columns}
        else:
            # Use default column names
            df.columns = [f"Column{i+1}" for i in range(len(df.columns))]
            data = {col: df[col].tolist() for col in df.columns}

        return data

    except ImportError:
        print("Warning: pandas not installed. Cannot import Excel files.")
        return {}

def export_dict_to_csv(data: Dict[str, List],
                      output_path: Optional[str] = None) -> str:
    """
    Export dictionary data to CSV file.

    Args:
        data (Dict[str, List]): Dictionary with column names as keys and data as lists
        output_path (str, optional): Path to save CSV file. If None, creates temp file.

    Returns:
        str: Path to the created CSV file
    """
    # Determine output path
    if output_path is None:
        temp_dir = tempfile.gettempdir()
        output_path = os.path.join(temp_dir, "sigmaplot_exported_data.csv")

    # Get column names
    columns = list(data.keys())

    # Find the max length of all columns
    max_length = max([len(data[col]) for col in columns])

    # Write to CSV file
    with open(output_path, 'w', newline='') as f:
        writer = csv.writer(f)

        # Write header
        writer.writerow(columns)

        # Write data rows
        for i in range(max_length):
            row = []
            for col in columns:
                if i < len(data[col]):
                    row.append(data[col][i])
                else:
                    row.append('')
            writer.writerow(row)

    print(f"Data exported to: {output_path}")
    return output_path

def convert_dataframe_to_csv(dataframe, output_path: Optional[str] = None) -> str:
    """
    Convert pandas DataFrame to CSV for SigmaPlot.

    Args:
        dataframe: pandas DataFrame
        output_path (str, optional): Path to save CSV file. If None, creates temp file.

    Returns:
        str: Path to the created CSV file
    """
    try:
        # Determine output path
        if output_path is None:
            temp_dir = tempfile.gettempdir()
            output_path = os.path.join(temp_dir, "sigmaplot_dataframe_data.csv")

        # Export DataFrame to CSV
        dataframe.to_csv(output_path, index=False)

        print(f"DataFrame exported to: {output_path}")
        return output_path

    except Exception as e:
        print(f"Error exporting DataFrame: {e}")
        return None

# EOF