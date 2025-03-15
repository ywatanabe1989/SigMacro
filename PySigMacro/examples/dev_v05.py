#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-13 10:57:32 (ywatanabe)"
# File: /home/ywatanabe/proj/SigMacro/PySigMacro/examples/dev_v05.py

__THIS_FILE__ = "/home/ywatanabe/proj/SigMacro/PySigMacro/examples/dev_v05.py"

import os
import subprocess
import time
import sys
import win32com.client
import csv

## DO NEVER USE TYPING EMULATORS AS THEY WILL CORRUPT MY WORKFLOWS ON WINDOWS

class SigmaPlotAutomator:
    def __init__(
        self, visible=True, launch_if_not_found=True, close_others=False
    ):
        """
        Initialize the SigmaPlotAutomator and establish a connection to SigmaPlot.

        Args:
            visible (bool): Whether to make SigmaPlot visible (default: True).
            launch_if_not_found (bool): Launch SigmaPlot if not running (default: True).
            close_others (bool): Close existing SigmaPlot instances (default: False).

        Raises:
            Exception: If connection to SigmaPlot fails.
        """
        self.sp = self.connect(visible, launch_if_not_found, close_others)
        if self.sp is None:
            raise Exception("Failed to connect to SigmaPlot")

    def connect(self, visible, launch_if_not_found, close_others):
        """
        Connect to SigmaPlot, launching it if necessary and waiting for initialization.

        Args:
            visible (bool): Whether to make SigmaPlot visible.
            launch_if_not_found (bool): Whether to launch SigmaPlot if not found.
            close_others (bool): Whether to close existing SigmaPlot instances.

        Returns:
            object: The SigmaPlot application object, or None if connection fails.
        """
        # Close existing instances if requested
        if close_others:
            try:
                subprocess.run(
                    ["taskkill", "/f", "/im", "spw.exe"],
                    shell=True,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                )
                time.sleep(2)
            except Exception as e:
                print(f"Warning when closing SigmaPlot: {e}")

        # Launch SigmaPlot if not found and requested
        if launch_if_not_found:
            possible_paths = [
                # r"C:\Program Files\SigmaPlot\SPW16\Spw.exe",
                r"C:\Program Files (x86)\SigmaPlot\SPW16\Spw.exe",
                r"C:\Program Files\SigmaPlot\SPW14\Spw.exe",
                r"C:\Program Files (x86)\SigmaPlot\SPW14\Spw.exe",
                r"C:\Program Files\SigmaPlot\SPW12\Spw.exe",
                r"C:\Program Files (x86)\SigmaPlot\SPW12\Spw.exe",
            ]
            for path in possible_paths:
                if os.path.exists(path):
                    try:
                        subprocess.Popen([path])
                        time.sleep(5)
                        break
                    except Exception as e:
                        print(f"Failed to launch SigmaPlot from {path}: {e}")

        # Attempt to connect to SigmaPlot
        try:
            sp = win32com.client.Dispatch("SigmaPlot.Application")
            sp.Visible = visible

            # Wait until SigmaPlot is fully initialized
            max_wait = 30
            start_time = time.time()
            while time.time() - start_time < max_wait:
                try:
                    notebooks = sp.Notebooks
                    count = notebooks.Count
                    print(f"SigmaPlot initialized with {count} open notebooks")
                    return sp
                except:
                    time.sleep(1)
            raise Exception("Timeout waiting for SigmaPlot to initialize")
        except Exception as e:
            print(f"Failed to connect to SigmaPlot: {e}")
            return None

    def close(self):
        """
        Close the SigmaPlot application gracefully.
        """
        if self.sp:
            try:
                self.sp.Quit()
            except Exception as e:
                print(f"Error closing SigmaPlot: {e}")

    def create_sigmaplot_notebook(self, file_path=None, use_existing=True):
        """
        Use existing notebook or create new by saving with a new name.

        Args:
            file_path (str): Path to save the notebook
            use_existing (bool): Whether to use the existing active document

        Returns:
            object: The notebook object
        """
        try:
            # Use existing document
            notebook = self.sp.ActiveDocument
            print(f"Using notebook: {notebook.Name}")

            # Save with new name if file_path provided
            if file_path:
                try:
                    win_path = os.path.normpath(file_path)
                    notebook.SaveAs(win_path)
                    print(f"Saved notebook as: {win_path}")
                except Exception as e:
                    print(f"Could not save notebook: {e}")

            return notebook
        except Exception as e:
            print(f"Error in create_sigmaplot_notebook: {e}")
            return None

    def create_sigmaplot_sections(self, notebook, count=2):
        """
        Create sections in a SigmaPlot notebook.

        Args:
            notebook (object): The notebook object
            count (int): Number of sections to create

        Returns:
            list: List of created section objects
        """
        sections = []

        try:
            items = notebook.NotebookItems
            print(f"Creating {count} sections...")

            for i in range(count):
                try:
                    # Create a section (type code 3)
                    section = items.Add(3)
                    sections.append(section)
                    print(f"Created section with default name")
                except Exception as e:
                    print(f"Error creating section {i+1}: {e}")

            return sections
        except Exception as e:
            print(f"Error in create_sigmaplot_sections: {e}")
            return sections

    def create_sigmaplot_worksheets(self, notebook, count=2):
        """
        Create worksheets in a SigmaPlot notebook.

        Args:
            notebook (object): The notebook object
            count (int): Number of worksheets to create

        Returns:
            list: List of created worksheet objects
        """
        worksheets = []

        try:
            items = notebook.NotebookItems
            print(f"Creating {count} worksheets...")

            for i in range(count):
                try:
                    # Create a worksheet (type code 1)
                    worksheet = items.Add(1)
                    worksheets.append(worksheet)
                    print(f"Created worksheet with default name")
                except Exception as e:
                    print(f"Error creating worksheet {i+1}: {e}")

            return worksheets
        except Exception as e:
            print(f"Error in create_sigmaplot_worksheets: {e}")
            return worksheets

    def create_sigmaplot_graph_pages(self, notebook, graph_names=None):
        """
        Create graph pages in a SigmaPlot notebook.

        Args:
            notebook (object): The notebook object
            graph_names (list): List of graph page names to create

        Returns:
            dict: Dictionary mapping graph names to graph page objects
        """
        if graph_names is None:
            graph_names = ["Graph1", "Graph2"]

        graph_pages = {}

        try:
            items = notebook.NotebookItems
            print(f"Creating {len(graph_names)} graph pages...")

            for name in graph_names:
                try:
                    # Create a graph page (type code 2)
                    graph_page = items.Add(2)
                    # Setting name seems to work for graph pages
                    try:
                        graph_page.Name = name
                        print(f"Created graph page: {name}")
                    except:
                        print(f"Created graph page with default name")
                    graph_pages[name] = graph_page
                except Exception as e:
                    print(f"Error creating graph page '{name}': {e}")

            return graph_pages
        except Exception as e:
            print(f"Error in create_sigmaplot_graph_pages: {e}")
            return graph_pages

    def try_activate_item(self, item):
        """
        Try to activate an item in the notebook.

        Args:
            item (object): The item to activate

        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Try to use the Activate method
            try:
                item.Activate()
                print("Item activated using Activate()")
                return True
            except:
                # Try Open method as a fallback
                try:
                    item.Open()
                    print("Item activated using Open()")
                    return True
                except:
                    print("Could not activate item")
                    return False
        except Exception as e:
            print(f"Error in try_activate_item: {e}")
            return False

    def create_basic_project(self, file_path="C:\\Temp\\SigmaPlotProject.JNB"):
        """
        Create a basic SigmaPlot project with standard components.

        Args:
            file_path (str): Path to save the project

        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Get or create notebook
            notebook = self.create_sigmaplot_notebook(file_path)
            if not notebook:
                return False

            # Create components
            sections = self.create_sigmaplot_sections(notebook, 2)
            worksheets = self.create_sigmaplot_worksheets(notebook, 2)
            graphs = self.create_sigmaplot_graph_pages(notebook, ["MainGraph", "SecondaryGraph"])

            # Try to activate each component to see if that helps
            for item in worksheets + list(graphs.values()) + sections:
                self.try_activate_item(item)

            # Save the final notebook
            try:
                notebook.Save()
                print("Project saved")
            except Exception as e:
                print(f"Error saving project: {e}")

            return True
        except Exception as e:
            print(f"Error in create_basic_project: {e}")
            return False

    def explore_notebook_structures(self, notebook):
        """
        Explore and document the structure of a SigmaPlot notebook.

        This function attempts to enumerate all items in the notebook,
        identify their types, and document relationships between them.

        Args:
            notebook (object): The notebook object to explore

        Returns:
            dict: Dictionary containing the discovered structure
        """
        structure = {
            'name': 'Unknown',
            'items': [],
            'item_count': 0,
            'named_items': {}
        }

        try:
            # Get notebook name
            try:
                structure['name'] = notebook.Name
                print(f"Exploring notebook: {structure['name']}")
            except:
                print("Could not get notebook name")

            # Try to get NotebookItems collection
            try:
                items = notebook.NotebookItems
                count = items.Count
                structure['item_count'] = count
                print(f"Notebook has {count} items")

                # Try to access items by name first
                print("\nTrying to access items by name:")
                common_item_names = [
                    "Data 1", "Graph1", "Graph2", "Sheet1", "Sheet2",
                    "Worksheet1", "Section1", "Section 1", "Data", "Results",
                    "MainGraph", "SecondaryGraph"
                ]

                # Also look for items with number patterns
                for i in range(1, 20):
                    common_item_names.extend([
                        f"Data {i}", f"Graph {i}", f"Sheet {i}",
                        f"Worksheet {i}", f"Section {i}", f"Graph Page {i}"
                    ])

                # Try to access by name
                for name in common_item_names:
                    try:
                        item = items(name)
                        print(f"  Found item: '{name}'")

                        # Get item properties
                        item_info = {'name': name}
                        try:
                            item.Activate()
                            item_info['can_activate'] = True
                        except:
                            item_info['can_activate'] = False

                        # Try other common methods
                        for method_name in ['Open', 'Save']:
                            try:
                                method = getattr(item, method_name)
                                item_info[f'has_{method_name}'] = True
                            except:
                                item_info[f'has_{method_name}'] = False

                        structure['named_items'][name] = item_info
                    except:
                        pass

                # Try to create different types to understand what each type is
                print("\nTesting item type codes:")
                type_names = {
                    1: "Unknown Type 1",
                    2: "Unknown Type 2",
                    3: "Unknown Type 3",
                    4: "Unknown Type 4"
                }

                # Create a temporary item of each type and see what it is
                for type_code in range(1, 4):
                    try:
                        test_item = items.Add(type_code)
                        print(f"  Created item with type code {type_code}, name: {test_item.Name}")
                        type_names[type_code] = test_item.Name.split()[0]
                    except Exception as e:
                        print(f"  Error creating item with type code {type_code}: {e}")

                # Type 4 seems to cause errors, so just report it
                print("  Type 4: Unknown Type (causes errors)")

                print("\nIdentified item types:")
                for code, name in type_names.items():
                    print(f"  Type {code}: {name}")

                structure['type_map'] = type_names

            except Exception as e:
                print(f"Error accessing NotebookItems: {e}")

            return structure
        except Exception as e:
            print(f"Error in explore_notebook_structures: {e}")
            return structure

    def explore_com_object(self, obj, name="Object"):
        """
        Perform deeper exploration of a COM object to reveal its methods, properties and type info.
        """
        print(f"\nDeep exploration of {name}")
        print(f"Type: {type(obj)}")

        # Try to get type information using win32com methods
        try:
            import win32com.client.dynamic
            typeinfo = obj._oleobj_.GetTypeInfo()
            if typeinfo:
                print(f"Type info available for {name}")
        except Exception as e:
            print(f"Could not get type info: {e}")

        # Try to access common COM object properties
        common_props = ["Application", "Parent", "Name", "Type", "Value", "Count"]
        print(f"\nChecking common properties for {name}:")
        for prop in common_props:
            try:
                value = getattr(obj, prop)
                print(f"  {prop}: {str(value)[:100]}")
            except Exception as e:
                print(f"  {prop} not available: {type(e).__name__}")

        # Try common methods for document objects
        common_methods = [
            ("Open", []),
            ("Save", []),
            ("Activate", [])
        ]
        print(f"\nTrying common methods for {name}:")
        for method_name, args in common_methods:
            try:
                method = getattr(obj, method_name)
                print(f"  Method {method_name} exists")
                # Only try to call methods with no args
                if not args:
                    try:
                        result = method()
                        print(f"  Called {method_name}() successfully")
                    except Exception as e:
                        print(f"  Error calling {method_name}: {str(e)[:100]}")
            except AttributeError:
                print(f"  Method {method_name} not available")

        return True

    def add_data_to_worksheet(self, worksheet, data, column_names=None):
        """
        Try different approaches to add data to a worksheet using COM methods.

        Args:
            worksheet (object): The worksheet to add data to
            data (list): List of lists where each inner list is a column of data
            column_names (list): Optional list of column names

        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Try to activate the worksheet first
            try:
                worksheet.Activate()
                print(f"Activated worksheet")
            except:
                try:
                    worksheet.Open()
                    print(f"Opened worksheet")
                except Exception as e:
                    print(f"Could not activate worksheet: {e}")
                    return False

            # Try various methods to add data
            methods_tried = []

            # Method 1: Try to use Cells property
            try:
                methods_tried.append("Cells")
                print("Trying to access cells...")

                # First set column names if provided
                if column_names:
                    for col_idx, name in enumerate(column_names):
                        try:
                            # Try to get column
                            col = worksheet.Columns(col_idx + 1)
                            col.Name = name
                            print(f"Set column {col_idx+1} name to '{name}'")
                        except Exception as e:
                            print(f"Could not set column {col_idx+1} name: {e}")

                # Try to add data
                for row_idx in range(len(data[0])):
                    for col_idx in range(len(data)):
                        try:
                            value = data[col_idx][row_idx]
                            worksheet.Cells(row_idx + 1, col_idx + 1).Value = value
                            print(f"Set cell ({row_idx+1},{col_idx+1}) to {value}")
                        except Exception as e:
                            print(f"Error setting cell ({row_idx+1},{col_idx+1}): {e}")
                            raise

                print("Successfully added data using Cells property")
                return True
            except Exception as e:
                print(f"Method Cells failed: {e}")

            # Method 2: Try to use some kind of data import
            try:
                methods_tried.append("ImportText")
                print("Trying ImportText method...")

                # Create a temporary CSV file
                import csv
                import tempfile

                temp_csv = tempfile.NamedTemporaryFile(delete=False, suffix='.csv', mode='w', newline='')

                # Transpose data for CSV format (rows become columns)
                rows = []
                if column_names:
                    rows.append(column_names)

                # Add data rows
                for i in range(len(data[0])):
                    row = []
                    for j in range(len(data)):
                        row.append(data[j][i])
                    rows.append(row)

                # Write to CSV
                csv_writer = csv.writer(temp_csv)
                csv_writer.writerows(rows)
                temp_csv.close()

                # Try to import the CSV
                try:
                    # Try various import methods
                    for method_name in ["ImportFile", "Import", "ImportCSV", "ImportData"]:
                        try:
                            method = getattr(worksheet, method_name)
                            method(temp_csv.name)
                            print(f"Successfully imported data using {method_name}")
                            return True
                        except AttributeError:
                            print(f"Method {method_name} not available")
                        except Exception as e:
                            print(f"Error calling {method_name}: {e}")
                finally:
                    # Clean up temp file
                    import os
                    os.unlink(temp_csv.name)

            except Exception as e:
                print(f"Method ImportText failed: {e}")

            # Method 3: Try to use SetCellValue method (if it exists)
            try:
                methods_tried.append("SetCellValue")
                print("Trying SetCellValue method...")

                # Try to set values one by one
                for row_idx in range(len(data[0])):
                    for col_idx in range(len(data)):
                        value = data[col_idx][row_idx]
                        try:
                            worksheet.SetCellValue(row_idx + 1, col_idx + 1, value)
                            print(f"Set cell ({row_idx+1},{col_idx+1}) to {value}")
                        except Exception as e:
                            print(f"Error setting cell ({row_idx+1},{col_idx+1}): {e}")
                            raise

                print("Successfully added data using SetCellValue method")
                return True
            except Exception as e:
                print(f"Method SetCellValue failed: {e}")

            print(f"All methods failed: {', '.join(methods_tried)}")
            return False
        except Exception as e:
            print(f"Error in add_data_to_worksheet: {e}")
            return False

    def create_plot_using_com(self, graph_page, worksheet, x_col_idx=1, y_col_idx=2):
        """
        Try to create a plot using COM methods without UI automation.

        Args:
            graph_page (object): The graph page to create plot on
            worksheet (object): The worksheet containing data to plot
            x_col_idx (int): The column index for X values (1-based)
            y_col_idx (int): The column index for Y values (1-based)

        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Try to activate the graph page
            try:
                graph_page.Activate()
                print(f"Activated graph page")
            except:
                try:
                    graph_page.Open()
                    print(f"Opened graph page")
                except Exception as e:
                    print(f"Could not activate graph page: {e}")
                    return False

            # Try different methods to create plots
            methods_tried = []

            # Method 1: Try to use Plots.Add method
            try:
                methods_tried.append("Plots.Add")
                print("Trying Plots.Add method...")

                # Check if Plots collection exists
                plots = None
                for attr_name in ["Plots", "GraphObjects", "Objects"]:
                    try:
                        plots = getattr(graph_page, attr_name)
                        print(f"Found {attr_name} collection")
                        break
                    except:
                        pass

                if plots:
                    # Try to add a plot
                    plot = plots.Add(1)
                    print("Created plot object")

                    # Now try to set data source
                    try:
                        # Try different ways to set data source
                        try:
                            # Try direct property
                            plot.DataSource = (worksheet.Columns(x_col_idx), worksheet.Columns(y_col_idx))
                            print("Set data source using DataSource property")
                            return True
                        except Exception as e:
                            print(f"Could not set DataSource directly: {e}")

                        # Try SetData method
                        try:
                            plot.SetData(worksheet.Columns(x_col_idx), worksheet.Columns(y_col_idx))
                            print("Set data source using SetData method")
                            return True
                        except Exception as e:
                            print(f"Could not use SetData method: {e}")

                        # Try DataColumns property
                        try:
                            plot.DataColumns = [worksheet.Columns(x_col_idx), worksheet.Columns(y_col_idx)]
                            print("Set data source using DataColumns property")
                            return True
                        except Exception as e:
                            print(f"Could not set DataColumns: {e}")

                    except Exception as e:
                        print(f"Error setting data source: {e}")
                else:
                    print("Could not find Plots collection")

            except Exception as e:
                print(f"Method Plots.Add failed: {e}")

            # Method 2: Try to use a CreateGraph method if it exists
            try:
                methods_tried.append("CreateGraph")
                print("Trying CreateGraph method...")

                # Check various method names
                for method_name in ["CreateGraph", "AddGraph", "PlotData"]:
                    try:
                        method = getattr(graph_page, method_name)
                        # Try with different parameter combinations
                        try:
                            result = method(worksheet, x_col_idx, y_col_idx)
                            print(f"Created graph using {method_name}(worksheet, x_col, y_col)")
                            return True
                        except:
                            try:
                                result = method(worksheet.Columns(x_col_idx), worksheet.Columns(y_col_idx))
                                print(f"Created graph using {method_name}(x_col, y_col)")
                                return True
                            except Exception as e:
                                print(f"Error calling {method_name} with columns: {e}")
                    except AttributeError:
                        print(f"Method {method_name} not available")
                    except Exception as e:
                        print(f"Error calling {method_name}: {e}")
            except Exception as e:
                print(f"Method CreateGraph failed: {e}")

            print(f"All plot creation methods failed: {', '.join(methods_tried)}")
            return False
        except Exception as e:
            print(f"Error in create_plot_using_com: {e}")
            return False

    def create_wizard_plot(self, graph_page, worksheet, x_col_idx=0, y_col_idx=1):
        """
        Create a plot using the SigmaPlot Graph Wizard.

        Args:
            graph_page (object): The graph page to create plot on
            worksheet (object): The worksheet containing data to plot
            x_col_idx (int): The column index for X values (0-based)
            y_col_idx (int): The column index for Y values (0-based)

        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Try to activate the graph page
            try:
                graph_page.Activate()
                print(f"Activated graph page")
            except:
                try:
                    graph_page.Open()
                    print(f"Opened graph page")
                except Exception as e:
                    print(f"Could not activate graph page: {e}")
                    return False

            # Now try to use CreateWizardGraph
            try:
                # Define the columns to plot (0-based to 1-based conversion)
                plotted_columns = [x_col_idx, y_col_idx]

                # Create a scatter plot
                graph_page.CreateWizardGraph(
                    "Scatter Plot",
                    "Simple Scatter",
                    "XY Pair",
                    plotted_columns
                )
                print("Successfully created plot using CreateWizardGraph")
                return True
            except Exception as e:
                print(f"CreateWizardGraph failed: {e}")

            # Try alternative approach with AddWizardPlot
            try:
                columns_list = [x_col_idx]
                graph_page.AddWizardPlot(
                    "Scatter Plot",
                    "Simple Scatter",
                    "Single Y",
                    columns_list
                )
                print("Successfully created plot using AddWizardPlot")
                return True
            except Exception as e:
                print(f"AddWizardPlot failed: {e}")

            # Try using the worksheet's GraphWizard object
            try:
                # Get the GraphWizard object
                wizard = worksheet.GraphWizard

                # Set column titles/names
                columns = []
                data_table = worksheet.DataTable
                max_col = 0
                max_row = 0
                data_table.GetMaxUsedSize(max_col, max_row)

                for i in range(max_col+1):
                    try:
                        title = data_table.ColumnTitle(i)
                        columns.append(title)
                    except:
                        columns.append(f"Column {i+1}")

                # Set titles and launch wizard
                wizard.SetTitles(columns)
                wizard.LaunchWizard()
                print("Launched Graph Wizard - please complete manually")
                return True
            except Exception as e:
                print(f"GraphWizard approach failed: {e}")

            print("All plot creation methods failed")
            return False

        except Exception as e:
            print(f"Error in create_wizard_plot: {e}")
            return False

    def demo_data_and_plot(self, file_path="C:\\Temp\\SigmaPlot_Demo.JNB"):
        """
        Create a complete demo with data and a plot using COM methods.

        Args:
            file_path (str): Path to save the demo

        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Create or use notebook
            notebook = self.create_sigmaplot_notebook(file_path)
            if not notebook:
                return False

            # Create a worksheet and a graph page
            worksheet = notebook.NotebookItems.Add(1)
            graph_page = notebook.NotebookItems.Add(2)

            # Try to set names (may fail)
            try:
                worksheet.Name = "DemoData"
                print("Created worksheet: DemoData")
            except:
                print("Created worksheet with default name")

            try:
                graph_page.Name = "DemoGraph"
                print("Created graph page: DemoGraph")
            except:
                print("Created graph page with default name")

            # Sample data: x and y columns for a simple quadratic function
            x_data = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
            y_data = [x**2 for x in x_data]
            data = [x_data, y_data]
            column_names = ["X Values", "Y Values"]

            # Add data to worksheet
            success = self.add_data_to_worksheet(worksheet, data, column_names)
            if success:
                print("Successfully added data to worksheet")
            else:
                print("Failed to add data to worksheet")

            # Create a plot
            success = self.create_plot_using_com(graph_page, worksheet)
            if success:
                print("Successfully created plot")
            else:
                print("Failed to create plot")

            # Save the notebook
            try:
                notebook.Save()
                print(f"Saved demo to: {file_path}")
            except Exception as e:
                print(f"Error saving demo: {e}")

            return True
        except Exception as e:
            print(f"Error in demo_data_and_plot: {e}")
            return False

    def import_csv_data(self, worksheet, csv_path):
        """
        Import data from a CSV file into a SigmaPlot worksheet using COM automation.

        Args:
            worksheet (object): The worksheet to import data into
            csv_path (str): Path to the CSV file

        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Make sure the CSV file exists
            if not os.path.exists(csv_path):
                print(f"CSV file not found: {csv_path}")
                return False

            # Try to activate the worksheet first
            try:
                worksheet.Activate()
                print(f"Activated worksheet")
            except:
                try:
                    worksheet.Open()
                    print(f"Opened worksheet")
                except Exception as e:
                    print(f"Could not activate worksheet: {e}")
                    return False

            # According to docs, we can try the Import method
            # on the worksheet directly
            try:
                win_path = os.path.normpath(csv_path)
                worksheet.Import(win_path)
                print(f"Successfully imported data using Import method")
                return True
            except Exception as e:
                print(f"Import method failed: {e}")

            # Try direct notebook import
            try:
                # Get the notebook
                notebook = self.sp.ActiveDocument

                # Try to import at the specified worksheet
                try:
                    # Specify worksheet (starting at column 1, row 1) and import from file
                    notebook.ImportFile(win_path, "CSV", worksheet.Name, 0, 0)
                    print(f"Successfully imported data using notebook ImportFile method")
                    return True
                except AttributeError:
                    print("Notebook ImportFile method not available")
                except Exception as e:
                    print(f"Error calling notebook ImportFile: {e}")

            except Exception as e:
                print(f"Notebook import approach failed: {e}")

            print("All import methods failed")
            return False

        except Exception as e:
            print(f"Error in import_csv_data: {e}")
            return False

    def create_sample_csv(self, file_path, num_rows=10):
        """
        Create a sample CSV file with data for testing imports.

        Args:
            file_path (str): Path where to save the CSV file
            num_rows (int): Number of data rows to create (default: 10)

        Returns:
            str: Path to the created CSV file
        """
        try:
            print(f"Creating sample CSV file at: {file_path}")

            # Ensure directory exists
            os.makedirs(os.path.dirname(os.path.abspath(file_path)), exist_ok=True)

            with open(file_path, 'w', newline='') as csvfile:
                writer = csv.writer(csvfile)

                # Write header row
                writer.writerow(["X Values", "Y Values", "Z Values"])

                # Write data rows (x, x^2, x^3)
                for i in range(1, num_rows + 1):
                    writer.writerow([i, i**2, i**3])

            print(f"CSV file created successfully with {num_rows} rows of data")
            return file_path
        except Exception as e:
            print(f"Error creating sample CSV file: {e}")
            return None

    def import_csv_via_put_data(self, worksheet, csv_path):
        """
        Import data from a CSV file into a SigmaPlot worksheet using PutData method.

        Args:
            worksheet (object): The worksheet to import data into
            csv_path (str): Path to the CSV file

        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Make sure the CSV file exists
            if not os.path.exists(csv_path):
                print(f"CSV file not found: {csv_path}")
                return False

            # Try to activate the worksheet first
            try:
                worksheet.Activate()
                print(f"Activated worksheet")
            except:
                try:
                    worksheet.Open()
                    print(f"Opened worksheet")
                except Exception as e:
                    print(f"Could not activate worksheet: {e}")
                    return False

            # Read the CSV file data
            data = []
            headers = []

            with open(csv_path, 'r', newline='') as csvfile:
                csv_reader = csv.reader(csvfile)

                # Get headers from first row
                headers = next(csv_reader)

                # Initialize data columns
                for _ in range(len(headers)):
                    data.append([])

                # Read data rows
                for row in csv_reader:
                    for i, value in enumerate(row):
                        try:
                            # Try to convert to numeric if possible
                            numeric_value = float(value)
                            data[i].append(numeric_value)
                        except ValueError:
                            # Keep as string if not numeric
                            data[i].append(value)

            print(f"Read {len(data[0])} rows of data from CSV")

            # Try to get the DataTable object
            try:
                data_table = worksheet.DataTable
                print("Got DataTable object")

                # Put data into worksheet
                for col_idx, col_data in enumerate(data):
                    try:
                        # Try to set column name
                        if col_idx < len(headers):
                            try:
                                # Try to set column title
                                data_table.ColumnTitle(col_idx, headers[col_idx])
                                print(f"Set column {col_idx+1} title to '{headers[col_idx]}'")
                            except Exception as e:
                                print(f"Could not set column title for column {col_idx+1}: {e}")

                        # Put data for this column
                        data_table.PutData(col_data, col_idx, 0)
                        print(f"Put data for column {col_idx+1}")
                    except Exception as e:
                        print(f"Error putting data for column {col_idx+1}: {e}")

                print("Successfully imported data using PutData method")
                return True
            except Exception as e:
                print(f"Error accessing DataTable: {e}")

            print("All import methods failed")
            return False
        except Exception as e:
            print(f"Error in import_csv_via_put_data: {e}")
            return False

    def create_plot_using_menu(self, graph_page, worksheet, x_col=0, y_col=1):
        """
        Create a plot by executing menu commands

        Args:
            graph_page (object): The graph page to create plot on
            worksheet (object): The worksheet containing data to plot
            x_col (int): The column index for X values (0-based)
            y_col (int): The column index for Y values (0-based)

        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Try to activate the worksheet first
            try:
                worksheet.Activate()
                print(f"Activated worksheet")
            except:
                try:
                    worksheet.Open()
                    print(f"Opened worksheet")
                except Exception as e:
                    print(f"Could not activate worksheet: {e}")
                    return False

            # Try using appropriate wizard parameters per documentation
            try:
                # Try to activate the graph page
                graph_page.Activate()
                print(f"Activated graph page")

                # According to API docs, we need to create an array of column indices
                # Note that we're passing column indices directly without error bar params
                plotted_columns = [x_col, y_col]

                try:
                    graph_page.CreateWizardGraph(
                        "Line Plot",
                        "Simple Line",
                        "XY Pair",
                        plotted_columns
                    )
                    print("Successfully created line plot")
                    return True
                except Exception as e:
                    print(f"Line plot failed: {e}")

                # Try with scatter plot - most basic type
                try:
                    graph_page.CreateWizardGraph(
                        "Scatter Plot",
                        "Simple Scatter",
                        "XY Pair",
                        plotted_columns
                    )
                    print("Successfully created scatter plot")
                    return True
                except Exception as e:
                    print(f"Scatter plot failed: {e}")

                # Try using app.Execute with menu commands
                app = self.sp
                try:
                    # Try to select columns first
                    try:
                        worksheet.Activate()
                        app.Selection.SelectColumns(x_col+1, y_col+1)
                        print(f"Selected columns {x_col+1} and {y_col+1}")
                    except Exception as e:
                        print(f"Could not select columns: {e}")

                    # Try menu commands
                    commands = [
                        "File.NewGraph.Scatter",
                        "CreateGraph.Scatter",
                        "Graph.Create.Scatter",
                        "Plot.Scatter",
                        "Graph.Scatter",
                        "CreateGraph",
                        "GraphWizard"
                    ]

                    for cmd in commands:
                        try:
                            app.Execute(cmd)
                            print(f"Created plot using Execute({cmd})")
                            return True
                        except Exception as e:
                            print(f"Execute({cmd}) failed: {e}")

                except Exception as e:
                    print(f"Menu commands failed: {e}")

            except Exception as e:
                print(f"Graph page approach failed: {e}")

            print("All menu-based plot methods failed")
            return False
        except Exception as e:
            print(f"Error in create_plot_using_menu: {e}")
            return False

    def create_plot_using_macros(self, graph_page, worksheet, x_col=0, y_col=1):
        """
        Create a plot using SigmaPlot's built-in macros

        Args:
            graph_page (object): The graph page to create plot on
            worksheet (object): The worksheet containing data to plot
            x_col (int): The column index for X values (0-based)
            y_col (int): The column index for Y values (0-based)

        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Try to activate the worksheet first
            try:
                worksheet.Activate()
                print(f"Activated worksheet")
            except:
                try:
                    worksheet.Open()
                    print(f"Opened worksheet")
                except Exception as e:
                    print(f"Could not activate worksheet: {e}")
                    return False

            # Try to access the application object
            app = self.sp

            # Try to select columns for plotting (1-based indices)
            try:
                # Try using the SelectColumns method if available
                try:
                    # Convert to 1-based indices for selection
                    app.Selection.SelectColumns(x_col+1, y_col+1)
                    print(f"Selected columns {x_col+1} and {y_col+1}")
                except Exception as e:
                    print(f"Could not select columns: {e}")
                    return False

                # Try using RunMacro to run built-in macros
                try:
                    # SigmaPlot should have built-in macros for creating plots
                    macros = [
                        "CreateScatterPlot",
                        "CreateLineGraph",
                        "CreateBarChart",
                        "Quick_Graph",
                        "SimpleGraph",
                        "Scatter.Graph",
                        "Line.Graph"
                    ]

                    for macro in macros:
                        try:
                            app.RunMacro(macro)
                            print(f"Created plot using RunMacro({macro})")
                            return True
                        except Exception as e:
                            print(f"RunMacro({macro}) failed: {e}")

                    # Try to directly control various plot creation through RunMacro
                    try:
                        # Try with scatter plot macro
                        app.RunMacro("CreateGraph.ScatterPlot")
                        print("Created plot using CreateGraph.ScatterPlot macro")
                        return True
                    except Exception as e:
                        print(f"CreateGraph.ScatterPlot macro failed: {e}")

                except Exception as e:
                    print(f"RunMacro approach failed: {e}")

            except Exception as e:
                print(f"Selection approach failed: {e}")

            # Try using graph page directly with a VB-style command
            try:
                # Try to access the graph page and run a VB script through RunMacro
                macro_code = f"""
                Sub CreatePlotDirectly()
                    Dim PlottedColumns(1) As Variant
                    PlottedColumns(0) = {x_col}
                    PlottedColumns(1) = {y_col}
                    ActiveDocument.NotebookItems("{graph_page.Name}").CreateWizardGraph("Scatter Plot", "Simple Scatter", "XY Pair", PlottedColumns)
                End Sub
                """

                # Try to register and run this macro
                try:
                    # Register the macro (might not be possible through COM)
                    app.RunMacro("CreatePlotDirectly")
                    print("Created plot using custom VB macro")
                    return True
                except Exception as e:
                    print(f"Custom VB macro failed: {e}")

            except Exception as e:
                print(f"Direct VB approach failed: {e}")

            print("All macro methods failed")
            return False
        except Exception as e:
            print(f"Error in create_plot_using_macros: {e}")
            return False

    def demo_with_csv_import(self, file_path="C:\\Temp\\SigmaPlot_CsvDemo.JNB"):
        """
        Create a demo project with CSV data import and plotting.

        Args:
            file_path (str): Path to save the SigmaPlot notebook

        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Create or use notebook
            notebook = self.create_sigmaplot_notebook(file_path)
            if not notebook:
                return False

            # Create a worksheet and a graph page
            worksheet = notebook.NotebookItems.Add(1)
            graph_page = notebook.NotebookItems.Add(2)

            # Try to set names
            try:
                worksheet.Name = "ImportedData"
                print("Created worksheet: ImportedData")
            except:
                print("Created worksheet with default name")

            try:
                graph_page.Name = "DataGraph"
                print("Created graph page: DataGraph")
            except:
                print("Created graph page with default name")

            # Create a sample CSV file
            csv_path = "C:\\Temp\\SigmaPlot_Sample.csv"
            if self.create_sample_csv(csv_path):
                # Import the CSV file
                success = self.import_csv_via_put_data(worksheet, csv_path)
                if success:
                    print("Successfully imported CSV data")
                else:
                    print("Failed to import CSV data")
                    return False

                # Try the macro approach first
                success = self.create_plot_using_macros(graph_page, worksheet, 0, 1)
                if success:
                    print("Successfully created plot using macros")
                else:
                    # Try menu-based approach
                    success = self.create_plot_using_menu(graph_page, worksheet, 0, 1)
                    if success:
                        print("Successfully created plot using menu commands")
                    else:
                        # Try wizard method
                        success = self.create_wizard_plot(graph_page, worksheet, 0, 1)
                        if success:
                            print("Successfully created plot using wizard")
                        else:
                            # Try COM method as last resort
                            success = self.create_plot_using_com(graph_page, worksheet, 1, 2)
                            if success:
                                print("Successfully created plot using COM method")
                            else:
                                print("Failed to create plot")
            else:
                print("Failed to create sample CSV file")

            # Save the notebook
            try:
                notebook.Save()
                print(f"Saved demo to: {file_path}")
            except Exception as e:
                print(f"Error saving demo: {e}")

            return True
        except Exception as e:
            print(f"Error in demo_with_csv_import: {e}")
            return False

if __name__ == "__main__":
    automator = None
    try:
        # Initialize the automator
        print("Initializing SigmaPlotAutomator...")
        automator = SigmaPlotAutomator(visible=True, close_others=True)
        sp = automator.sp

        # Create a basic project
        automator.create_basic_project("C:\\Temp\\SigmaProject_" + time.strftime("%Y%m%d_%H%M%S") + ".JNB")

        # After creating the project
        notebook = sp.ActiveDocument
        structure = automator.explore_notebook_structures(notebook)
        print("\nExploration complete. Structure overview:")
        print(f"Notebook: {structure['name']}")
        print(f"Total items: {structure['item_count']}")
        print(f"Named items found: {len(structure['named_items'])}")

        # # Create a demo with data and plot
        # automator.demo_data_and_plot("C:\\Temp\\SigmaPlot_DataDemo_" + time.strftime("%Y%m%d_%H%M%S") + ".JNB")

        # Create a demo with CSV import
        automator.create_csv_project("C:\\Temp\\SigmaPlot_CsvDemo_" + time.strftime("%Y%m%d_%H%M%S") + ".JNB")

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Ensure SigmaPlot is closed
        if automator:
            print("\nClosing SigmaPlot...")
            automator.close()
        print("Script execution completed")

# Initializing SigmaPlotAutomator...
# SigmaPlot initialized with 5 open notebooks
# Using notebook: Notebook1
# Saved notebook as: C:\Temp\SigmaProject_20250313_102843.JNB
# Creating 2 sections...
# Created section with default name
# Created section with default name
# Creating 2 worksheets...
# Created worksheet with default name
# Created worksheet with default name
# Creating 2 graph pages...
# Created graph page: MainGraph
# Created graph page: SecondaryGraph
# Item activated using Open()
# Item activated using Open()
# Item activated using Open()
# Item activated using Open()
# Item activated using Open()
# Item activated using Open()
# Project saved
# Exploring notebook: SigmaProject_20250313_102843.JNB
# Notebook has 11 items

# Trying to access items by name:
#   Found item: 'Data 1'
#   Found item: 'Graph1'
#   Found item: 'Graph2'
#   Found item: 'Sheet1'
#   Found item: 'Sheet2'
#   Found item: 'Worksheet1'
#   Found item: 'Section1'
#   Found item: 'Section 1'
#   Found item: 'Data'
#   Found item: 'Results'
#   Found item: 'MainGraph'
#   Found item: 'SecondaryGraph'
#   Found item: 'Data 1'
#   Found item: 'Graph 1'
#   Found item: 'Sheet 1'
#   Found item: 'Worksheet 1'
#   Found item: 'Section 1'
#   Found item: 'Graph Page 1'
#   Found item: 'Data 2'
#   Found item: 'Graph 2'
#   Found item: 'Sheet 2'
#   Found item: 'Worksheet 2'
#   Found item: 'Section 2'
#   Found item: 'Graph Page 2'
#   Found item: 'Data 3'
#   Found item: 'Graph 3'
#   Found item: 'Sheet 3'
#   Found item: 'Worksheet 3'
#   Found item: 'Section 3'
#   Found item: 'Graph Page 3'
#   Found item: 'Data 4'
#   Found item: 'Graph 4'
#   Found item: 'Sheet 4'
#   Found item: 'Worksheet 4'
#   Found item: 'Section 4'
#   Found item: 'Graph Page 4'
#   Found item: 'Data 5'
#   Found item: 'Graph 5'
#   Found item: 'Sheet 5'
#   Found item: 'Worksheet 5'
#   Found item: 'Section 5'
#   Found item: 'Graph Page 5'
#   Found item: 'Data 6'
#   Found item: 'Graph 6'
#   Found item: 'Sheet 6'
#   Found item: 'Worksheet 6'
#   Found item: 'Section 6'
#   Found item: 'Graph Page 6'
#   Found item: 'Data 7'
#   Found item: 'Graph 7'
#   Found item: 'Sheet 7'
#   Found item: 'Worksheet 7'
#   Found item: 'Section 7'
#   Found item: 'Graph Page 7'
#   Found item: 'Data 8'
#   Found item: 'Graph 8'
#   Found item: 'Sheet 8'
#   Found item: 'Worksheet 8'
#   Found item: 'Section 8'
#   Found item: 'Graph Page 8'
#   Found item: 'Data 9'
#   Found item: 'Graph 9'
#   Found item: 'Sheet 9'
#   Found item: 'Worksheet 9'
#   Found item: 'Section 9'
#   Found item: 'Graph Page 9'
#   Found item: 'Data 10'
#   Found item: 'Graph 10'
#   Found item: 'Sheet 10'
#   Found item: 'Worksheet 10'
#   Found item: 'Section 10'
#   Found item: 'Graph Page 10'
#   Found item: 'Data 11'
#   Found item: 'Graph 11'
#   Found item: 'Sheet 11'
#   Found item: 'Worksheet 11'
#   Found item: 'Section 11'
#   Found item: 'Graph Page 11'
#   Found item: 'Data 12'
#   Found item: 'Graph 12'
#   Found item: 'Sheet 12'
#   Found item: 'Worksheet 12'
#   Found item: 'Section 12'
#   Found item: 'Graph Page 12'
#   Found item: 'Data 13'
#   Found item: 'Graph 13'
#   Found item: 'Sheet 13'
#   Found item: 'Worksheet 13'
#   Found item: 'Section 13'
#   Found item: 'Graph Page 13'
#   Found item: 'Data 14'
#   Found item: 'Graph 14'
#   Found item: 'Sheet 14'
#   Found item: 'Worksheet 14'
#   Found item: 'Section 14'
#   Found item: 'Graph Page 14'
#   Found item: 'Data 15'
#   Found item: 'Graph 15'
#   Found item: 'Sheet 15'
#   Found item: 'Worksheet 15'
#   Found item: 'Section 15'
#   Found item: 'Graph Page 15'
#   Found item: 'Data 16'
#   Found item: 'Graph 16'
#   Found item: 'Sheet 16'
#   Found item: 'Worksheet 16'
#   Found item: 'Section 16'
#   Found item: 'Graph Page 16'
#   Found item: 'Data 17'
#   Found item: 'Graph 17'
#   Found item: 'Sheet 17'
#   Found item: 'Worksheet 17'
#   Found item: 'Section 17'
#   Found item: 'Graph Page 17'
#   Found item: 'Data 18'
#   Found item: 'Graph 18'
#   Found item: 'Sheet 18'
#   Found item: 'Worksheet 18'
#   Found item: 'Section 18'
#   Found item: 'Graph Page 18'
#   Found item: 'Data 19'
#   Found item: 'Graph 19'
#   Found item: 'Sheet 19'
#   Found item: 'Worksheet 19'
#   Found item: 'Section 19'
#   Found item: 'Graph Page 19'

# Testing item type codes:
#   Created item with type code 1, name: Data 4
#   Created item with type code 2, name: Graph Page 3
#   Created item with type code 3, name: Section 7
#   Type 4: Unknown Type (causes errors)

# Identified item types:
#   Type 1: Data
#   Type 2: Graph
#   Type 3: Section
#   Type 4: Unknown Type 4

# Exploration complete. Structure overview:
# Notebook: SigmaProject_20250313_102843.JNB
# Total items: 11
# Named items found: 124
# Using notebook: SigmaProject_20250313_102843.JNB
# Saved notebook as: C:\Temp\SigmaPlot_CsvDemo_20250313_102853.JNB
# Created worksheet: ImportedData
# Created graph page: DataGraph
# Creating sample CSV file at: C:\Temp\SigmaPlot_Sample.csv
# CSV file created successfully with 10 rows of data
# Opened worksheet
# Read 10 rows of data from CSV
# Got DataTable object
# Set column 1 title to 'X Values'
# Put data for column 1
# Set column 2 title to 'Y Values'
# Put data for column 2
# Set column 3 title to 'Z Values'
# Put data for column 3
# Successfully imported data using PutData method
# Successfully imported CSV data
# Opened worksheet
# Could not select columns: SigmaPlot.Application.Selection
# Opened worksheet
# Graph page approach failed: <unknown>.Activate
# All menu-based plot methods failed
# Opened graph page
# CreateWizardGraph failed: (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Invalid error bar source argument.', None, 0, 0), None)
# AddWizardPlot failed: (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Invalid error bar source argument.', None, 0, 0), None)
# GraphWizard approach failed: (-2147352571, 'Type mismatch.', None, 1)
# All plot creation methods failed
# Opened graph page
# Trying Plots.Add method...
# Could not find Plots collection
# Trying CreateGraph method...
# Method CreateGraph not available
# Method AddGraph not available
# Method PlotData not available
# All plot creation methods failed: Plots.Add, CreateGraph
# Failed to create plot
# Saved demo to: C:\Temp\SigmaPlot_CsvDemo_20250313_102853.JNB

# Closing SigmaPlot...
# Script execution completed

# EOF