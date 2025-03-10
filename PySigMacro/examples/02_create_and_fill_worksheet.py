#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 10:30:54 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/examples/02_create_and_fill_worksheet.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/examples/02_create_and_fill_worksheet.py"

import argparse
from pysigmacro.utils.inspect_sigmaplot import inspect_sigmaplot
from pysigmacro.core.connection import connect
import random
import time

def display_results(results):
    """
    Display inspection results in a formatted way.

    Args:
        results: Dictionary with inspection results
    """
    print("===== SigmaPlot COM Inspection Results =====")

    # APPLICATION STRUCTURE
    print("APPLICATION STRUCTURE:")
    print("--------------------------------------------------")
    if 'application' in results and 'type' in results['application']:
        print(f"Type: {results['application']['type']}")
    else:
        print("Type: Unknown")

    # APPLICATION PROPERTIES
    if 'app_properties' in results:
        print("\nAPPLICATION PROPERTIES:")
        print("--------------------------------------------------")
        for prop, value in results['app_properties'].items():
            print(f"{prop}: {value}")

    # NOTEBOOKS
    if 'notebooks' in results and isinstance(results['notebooks'], dict):
        print("\nNOTEBOOKS:")
        print("--------------------------------------------------")
        print(f"Count: {results['notebooks'].get('count', 0)}")

        items = results['notebooks'].get('items', {})
        if items:
            print("Notebook items:")
            for key, notebook in items.items():
                if isinstance(notebook, dict) and 'name' in notebook:
                    print(f"- {notebook['name']}")
                else:
                    print(f"- {key}")

    # ACTIVE DOCUMENT
    if 'active_document' in results:
        print("\nACTIVE DOCUMENT:")
        print("--------------------------------------------------")
        if isinstance(results['active_document'], dict):
            active_doc = results['active_document']
            print(f"Name: {active_doc.get('name', 'Unknown')}")
            print(f"Type: {active_doc.get('type', 'Unknown')}")

            if 'collections' in active_doc:
                print("Collections:")
                for coll_name in active_doc['collections'].keys():
                    print(f"- {coll_name}")
        else:
            print(f"Information: {results['active_document']}")

    # NOTEBOOK OPERATIONS
    if 'notebook_operations' in results:
        print("\nNOTEBOOK OPERATIONS:")
        print("--------------------------------------------------")
        ops = results['notebook_operations']

        if 'activate' in ops:
            if isinstance(ops['activate'], dict) and 'success' in ops['activate']:
                success = "Success" if ops['activate']['success'] else "Failed"
                print(f"activate: {success}")
            else:
                print(f"activate: {ops['activate']}")

        if 'available_methods' in ops:
            print(f"available_methods: {ops['available_methods']}")

    print("===== End of Inspection Results =====")

def create_and_fill_worksheet(sigmaplot_app=None, data=None, worksheet_name="Data", notebook_index=None):
    """
    Create a new worksheet in SigmaPlot and populate it with data.

    Args:
        sigmaplot_app: SigmaPlot application object. If None, will connect to existing instance.
        data: Dictionary or pandas DataFrame with data to fill. If None, creates empty worksheet.
        worksheet_name: Name for the new worksheet
        notebook_index: Index of notebook to use (1-based). If None, uses active notebook.

    Returns:
        Tuple of (worksheet, success_status)
    """
    import time
    from pysigmacro.core.connection import connect
    import logging
    import pythoncom
    import win32com.client

    # Set up logging
    logger = logging.getLogger(__name__)

    # Connect to SigmaPlot if needed
    if sigmaplot_app is None:
        sigmaplot_app = connect(visible=True, launch_if_not_found=True)
        time.sleep(1)

    # Activate the specified notebook if requested
    if notebook_index is not None:
        try:
            notebooks = sigmaplot_app.Notebooks
            if notebooks.Count >= notebook_index:
                # Use direct function call which seems to work
                notebook = notebooks(notebook_index)
                notebook.Activate()
                time.sleep(0.5)
        except Exception as e:
            logger.error(f"Error activating notebook {notebook_index}: {e}")
            return None, False

    # Get active document
    try:
        active_doc = sigmaplot_app.ActiveDocument
        if not active_doc:
            logger.error("No active document available")
            return None, False

        print(f"Active document: {active_doc.Name}")
    except Exception as e:
        logger.error(f"Error accessing active document: {e}")
        return None, False

    # Try method using SigmaPlot's Execute command (most reliable)
    worksheet = None
    try:
        print("Trying Execute NewWorksheet command...")
        sigmaplot_app.Execute("NewWorksheet")
        time.sleep(2)

        # Try to get the newly created worksheet from active sheet
        try:
            if hasattr(active_doc, 'ActiveSheet'):
                print("Getting ActiveSheet...")
                worksheet = active_doc.ActiveSheet
            elif hasattr(active_doc, 'ActiveWorksheet'):
                print("Getting ActiveWorksheet...")
                worksheet = active_doc.ActiveWorksheet
        except Exception as e:
            print(f"Error getting active sheet: {e}")
    except Exception as e:
        print(f"Error executing NewWorksheet command: {e}")

    # Alternate method using SendKeys to simulate keyboard shortcut
    if worksheet is None:
        try:
            import win32com.client
            print("Trying SendKeys method...")
            shell = win32com.client.Dispatch("WScript.Shell")
            sigmaplot_app.Activate()
            time.sleep(0.5)

            # Send Ctrl+W which is the keyboard shortcut for New Worksheet in SigmaPlot
            shell.SendKeys("^w")
            time.sleep(2)

            # Try to get active sheet again
            try:
                if hasattr(active_doc, 'ActiveSheet'):
                    worksheet = active_doc.ActiveSheet
                elif hasattr(active_doc, 'ActiveWorksheet'):
                    worksheet = active_doc.ActiveWorksheet
            except Exception as e:
                print(f"Error getting active sheet after SendKeys: {e}")
        except Exception as e:
            print(f"Error using SendKeys method: {e}")

    # SigmaPlot 12 alternative - use Sheets collection
    if worksheet is None and hasattr(active_doc, 'Sheets'):
        try:
            print("Trying Sheets.Add method...")
            sheets = active_doc.Sheets
            if hasattr(sheets, 'Add'):
                # Try with specific worksheet type
                sheet_types = {"Worksheet": 0, "Graph": 1, "Report": 2}
                worksheet = sheets.Add(sheet_types["Worksheet"])
                time.sleep(1)
            elif hasattr(sheets, 'AddWorksheet'):
                worksheet = sheets.AddWorksheet()
                time.sleep(1)
        except Exception as e:
            print(f"Error using Sheets.Add method: {e}")

    # Try explicit menu command as another alternative
    if worksheet is None:
        try:
            print("Trying menu command...")
            # Menu for "New Worksheet" in SigmaPlot
            menu_command = "File.New.Worksheet"
            sigmaplot_app.Execute(menu_command)
            time.sleep(2)

            # Try to get active sheet again
            if hasattr(active_doc, 'ActiveSheet'):
                worksheet = active_doc.ActiveSheet
            elif hasattr(active_doc, 'ActiveWorksheet'):
                worksheet = active_doc.ActiveWorksheet
        except Exception as e:
            print(f"Error executing menu command: {e}")

    # Check if we have a worksheet
    if worksheet is None:
        print("All worksheet creation methods failed. Let's see what's available:")

        # Debug: Print available methods and properties of active document
        try:
            print("\nActive Document Properties:")
            for attr_name in dir(active_doc):
                try:
                    attr = getattr(active_doc, attr_name)
                    print(f"- {attr_name}: {type(attr)}")
                except:
                    print(f"- {attr_name}: [error accessing]")
        except:
            print("Could not list active document properties")

        logger.error("Failed to create worksheet using any method")
        return None, False

    # Try to set worksheet name
    try:
        print(f"Setting worksheet name to: {worksheet_name}")
        if hasattr(worksheet, 'Name'):
            worksheet.Name = worksheet_name
    except Exception as e:
        print(f"Could not set worksheet name: {e}")

    # Fill with data if provided
    if data is not None:
        try:
            print("Filling worksheet with data...")
            # Convert pandas DataFrame to dict if needed
            if hasattr(data, 'to_dict'):
                data = data.to_dict('list')

            # Try multiple methods to access cells
            cell_access_methods = [
                lambda r, c: worksheet.Cells(r, c),
                lambda r, c: worksheet.Cell(r, c),
                lambda r, c: getattr(worksheet, f"Cell({r},{c})"),
                lambda r, c: worksheet.Cells.Item(r, c)
            ]

            # Determine data dimensions
            columns = list(data.keys())
            num_columns = len(columns)
            if num_columns == 0:
                return worksheet, True

            # Assuming all columns have same length
            first_col = data[columns[0]]
            num_rows = len(first_col) if isinstance(first_col, (list, tuple)) else 1

            # Print data dimensions
            print(f"Data has {num_columns} columns and up to {num_rows} rows")

            # Fill headers (column names)
            cell_method = None
            for method in cell_access_methods:
                try:
                    # Try first cell as test
                    cell = method(1, 1)
                    cell.Value = "Test"
                    # If we got here, the method works
                    cell_method = method
                    print(f"Found working cell access method")
                    break
                except Exception as e:
                    continue

            if cell_method is None:
                print("Could not find working cell access method")
                return worksheet, False

            # Fill headers with working method
            for col_idx, col_name in enumerate(columns):
                try:
                    # Convert to 1-based indexing
                    cell = cell_method(1, col_idx + 1)
                    cell.Value = col_name
                except Exception as e:
                    print(f"Error setting header for column {col_idx}: {e}")

            # Fill data values
            for col_idx, col_name in enumerate(columns):
                col_data = data[col_name]

                # Handle scalar vs list data
                if not isinstance(col_data, (list, tuple)):
                    col_data = [col_data]

                for row_idx, value in enumerate(col_data):
                    try:
                        # +2 because row 1 is headers, and convert to 1-based indexing
                        cell = cell_method(row_idx + 2, col_idx + 1)
                        cell.Value = value
                    except Exception as e:
                        print(f"Error setting cell ({row_idx+2}, {col_idx+1}): {e}")

            print("Data filling completed")
        except Exception as e:
            print(f"Error populating worksheet with data: {e}")
            return worksheet, False

    print("Worksheet creation and population succeeded!")
    return worksheet, True

def test_create_worksheet():
    """Test function to create and fill a worksheet with sample data."""
    print("Starting worksheet creation test...")

    # Sample data
    data = {
        "X": [i for i in range(10)],
        "Y1": [random.random() * 10 for _ in range(10)],
        "Y2": [random.random() * 5 + 5 for _ in range(10)]
    }

    # Connect to SigmaPlot
    print("Connecting to SigmaPlot...")
    sigmaplot = connect(visible=True, launch_if_not_found=True)

    # Create and fill worksheet
    print("Creating and filling worksheet...")
    worksheet, success = create_and_fill_worksheet(
        sigmaplot_app=sigmaplot,
        data=data,
        worksheet_name="Sample Data"
    )

    if success:
        print("Successfully created and filled worksheet!")
    else:
        print("Failed to create or fill worksheet")

    return worksheet

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="SigmaPlot worksheet creation and inspection tool")
    parser.add_argument('--inspect', action='store_true', help='Run inspection of SigmaPlot COM interface')
    parser.add_argument('--create', action='store_true', help='Create and fill a sample worksheet')

    args = parser.parse_args()

    if args.inspect:
        # Run inspection
        results = inspect_sigmaplot()
        display_results(results)

    if args.create:
        # Create worksheet
        worksheet = test_create_worksheet()

    # If no arguments, default to inspection
    if not (args.inspect or args.create):
        results = inspect_sigmaplot()
        display_results(results)

# EOF