#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 11:12:53 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/core/SigmaPlotWorksheet.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/core/SigmaPlotWorksheet.py"
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 10:16:30 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/core/SigmaPlotWorksheet.py
THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/core/SigmaPlotWorksheet.py"
"""
Worksheet manipulation utilities for SigmaPlot.
"""
import os
import time
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
        app.Execute("NewWorksheet")
        time.sleep(1)
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
        app.Execute("NewWorksheet")
        time.sleep(1)

        # Import CSV data
        try:
            # Try direct import if available
            if hasattr(app, 'Execute'):
                app.Execute(f"ImportASCII(\"{csv_path}\")")
                time.sleep(1)
                return app
            else:
                print("Execute method not available for importing CSV")
                return None
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
            if not hasattr(app, 'ActiveDocument'):
                app.Execute("NewWorksheet")
                time.sleep(1)

            # Get active document
            active_doc = app.ActiveDocument

            # Set column headers
            app.Execute("WorksheetCell(1, 1) = \"X\"")
            app.Execute("WorksheetCell(1, 2) = \"Y\"")

            # Set data values
            for i, (x, y) in enumerate(zip(x_values, y_values)):
                # Use Execute to set cell values
                app.Execute(f"WorksheetCell({i+2}, 1) = {float(x)}")
                app.Execute(f"WorksheetCell({i+2}, 2) = {float(y)}")

                # Process in smaller batches to avoid command buffer issues
                if i % 20 == 0:
                    time.sleep(0.1)

            return app
        except Exception as e:
            print(f"Error setting worksheet data: {e}")
            # Try alternative method using DataTable
            try:
                print("Trying DataTable method...")
                worksheet = app.ActiveDocument.ActiveSheet
                if hasattr(worksheet, 'DataTable'):
                    data_table = worksheet.DataTable

                    # Set headers
                    data_table.putData(1, 1, ["X"])
                    data_table.putData(2, 1, ["Y"])

                    # Set data values
                    for i, (x, y) in enumerate(zip(x_values, y_values)):
                        data_table.putData(1, i+2, [float(x)])
                        data_table.putData(2, i+2, [float(y)])
                    return app
                else:
                    print("DataTable property not available")
                    return None
            except Exception as e2:
                print(f"Error with alternative method: {e2}")
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
            if not hasattr(app, 'ActiveDocument'):
                app.Execute("NewWorksheet")
                time.sleep(1)

            # Get active document
            active_doc = app.ActiveDocument

            # Set column headers
            col_names = list(data_columns.keys())
            for col_idx, name in enumerate(col_names):
                app.Execute(f"WorksheetCell(1, {col_idx+1}) = \"{name}\"")

            # Find the maximum length of data columns
            max_rows = max([len(data_columns[col]) for col in col_names])

            # Set data for each column
            for col_idx, name in enumerate(col_names):
                column_data = data_columns[name]
                for row_idx, value in enumerate(column_data):
                    app.Execute(f"WorksheetCell({row_idx+2}, {col_idx+1}) = {float(value)}")

                    # Process in smaller batches to avoid command buffer issues
                    if row_idx % 20 == 0:
                        time.sleep(0.1)

            return app
        except Exception as e:
            print(f"Error setting multi-column data: {e}")
            # Try alternative method using DataTable
            try:
                print("Trying DataTable method...")
                worksheet = app.ActiveDocument.ActiveSheet
                if hasattr(worksheet, 'DataTable'):
                    data_table = worksheet.DataTable

                    # Set headers
                    col_names = list(data_columns.keys())
                    for col_idx, name in enumerate(col_names):
                        data_table.putData(col_idx+1, 1, [name])

                    # Set data for each column
                    for col_idx, name in enumerate(col_names):
                        column_data = data_columns[name]
                        for row_idx, value in enumerate(column_data):
                            data_table.putData(col_idx+1, row_idx+2, [float(value)])
                    return app
                else:
                    print("DataTable property not available")
                    return None
            except Exception as e2:
                print(f"Error with alternative method: {e2}")
                return None

# EOF