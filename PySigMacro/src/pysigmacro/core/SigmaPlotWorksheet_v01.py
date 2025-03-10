#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 08:44:29 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/core/SigmaPlotWorksheet.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/core/SigmaPlotWorksheet.py"

"""
Worksheet manipulation utilities for SigmaPlot.
"""

import os
from typing import List, Dict, Tuple, Optional

from pysigmacro.core.connection import connect

class SigmaPlotWorksheet:
    """Class to handle SigmaPlot worksheet operations"""

    @staticmethod
    def new(visible=True):
        """
        Create a new worksheet.

        Args:
            visible (bool): Make SigmaPlot visible

        Returns:
            SigmaPlot application COM object or None if failed
        """
        app = connect(visible=visible)
        if not app:
            print("Failed to connect to SigmaPlot")
            return None

        # Create a new worksheet
        app.NewWorksheet()
        return app

    @staticmethod
    def import_csv(csv_path, visible=True):
        """
        Import CSV data into a new worksheet.

        Args:
            csv_path (str): Path to CSV file
            visible (bool): Make SigmaPlot visible

        Returns:
            SigmaPlot application COM object or None if failed
        """
        app = connect(visible=visible)
        if not app:
            print("Failed to connect to SigmaPlot")
            return None

        # Create a new worksheet
        app.NewWorksheet()

        # Import CSV data
        try:
            app.CurrentWorksheet.ImportData(csv_path, 1, 1, ",")
            return app
        except Exception as e:
            print(f"Error importing CSV: {e}")
            return None

    @staticmethod
    def set_data(app, x_values: List, y_values: List):
        """
        Set x and y values directly in worksheet.

        Args:
            app: SigmaPlot application COM object
            x_values (List): X-axis values
            y_values (List): Y-axis values

        Returns:
            SigmaPlot application COM object or None if failed
        """
        if not app:
            print("No application connection provided")
            return None

        try:
            # Create a new worksheet if needed
            if not hasattr(app, 'CurrentWorksheet'):
                app.NewWorksheet()

            # Set X values (column 1)
            for i, x in enumerate(x_values):
                # Try different methods to set cell values
                try:
                    app.CurrentWorksheet.Cells(i+1, 1).Value = x
                except:
                    try:
                        app.CurrentWorksheet.SetCell(i+1, 1, x)
                    except:
                        print(f"Warning: Could not set cell value at {i+1}, 1")

            # Set Y values (column 2)
            for i, y in enumerate(y_values):
                try:
                    app.CurrentWorksheet.Cells(i+1, 2).Value = y
                except:
                    try:
                        app.CurrentWorksheet.SetCell(i+1, 2, y)
                    except:
                        print(f"Warning: Could not set cell value at {i+1}, 2")

            return app

        except Exception as e:
            print(f"Error setting worksheet data: {e}")
            return None

    @staticmethod
    def set_multi_column_data(app, data_columns: Dict[str, List]):
        """
        Set multiple columns of data in a worksheet.

        Args:
            app: SigmaPlot application COM object
            data_columns (Dict[str, List]): Dictionary mapping column names to data lists

        Returns:
            SigmaPlot application COM object or None if failed
        """
        if not app:
            print("No application connection provided")
            return None

        try:
            # Create a new worksheet if needed
            if not hasattr(app, 'CurrentWorksheet'):
                app.NewWorksheet()

            # Set column headers
            col_names = list(data_columns.keys())
            for col_idx, name in enumerate(col_names):
                try:
                    app.CurrentWorksheet.Cells(0, col_idx+1).Value = name
                except:
                    try:
                        app.CurrentWorksheet.SetCell(0, col_idx+1, name)
                    except:
                        print(f"Warning: Could not set header for column {col_idx+1}")

            # Find the maximum length of data columns
            max_rows = max([len(data_columns[col]) for col in col_names])

            # Set data for each column
            for col_idx, name in enumerate(col_names):
                column_data = data_columns[name]
                for row_idx, value in enumerate(column_data):
                    try:
                        app.CurrentWorksheet.Cells(row_idx+1, col_idx+1).Value = value
                    except:
                        try:
                            app.CurrentWorksheet.SetCell(row_idx+1, col_idx+1, value)
                        except:
                            print(f"Warning: Could not set value at {row_idx+1}, {col_idx+1}")

            return app

        except Exception as e:
            print(f"Error setting multi-column data: {e}")
            return None

# EOF