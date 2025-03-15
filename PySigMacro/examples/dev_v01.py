#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-13 09:27:13 (ywatanabe)"
# File: /home/ywatanabe/proj/SigMacro/PySigMacro/examples/tmp.py

__THIS_FILE__ = "/home/ywatanabe/proj/SigMacro/PySigMacro/examples/tmp.py"

import os
import subprocess
import time
import sys
import win32com.client


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

    def create_notebook(self, name="NewNotebook"):
        """
        Create a new notebook in SigmaPlot.
        Args:
            name (str): Name of the notebook (default: "NewNotebook").
        Returns:
            object: The created notebook object.
        """
        try:
            # Try to use the existing ActiveDocument
            notebook = self.sp.ActiveDocument

            # Set the name
            try:
                notebook.Name = name
                print(f"Set active document name to: {name}")
            except Exception as e:
                print(f"Could not set document name: {e}")

            return notebook
        except Exception as e:
            print(f"Error in create_notebook: {e}")
            raise

    def create_section(self, notebook, name="NewSection"):
        """
        Create a new section (worksheet) in the specified notebook.
        Args:
            notebook (object): The notebook object to add the section to.
            name (str): Name of the section (default: "NewSection").
        Returns:
            object: The created section (worksheet) object.
        """
        section = notebook.NotebookItems.Add(1)
        section.Name = name
        return section

    def import_csv(self, worksheet, csv_path):
        """
        Import a CSV file into the specified worksheet.
        Args:
            worksheet (object): The worksheet object to import data into.
            csv_path (str): Path to the CSV file.
        Raises:
            FileNotFoundError: If the CSV file does not exist.
        """
        if not os.path.exists(csv_path):
            raise FileNotFoundError(f"CSV file not found: {csv_path}")
        win_path = os.path.normpath(csv_path)
        worksheet.ImportFile(win_path, "CSV")

    def set_column_names_from_first_row(self, worksheet):
        """
        Set column names in the worksheet using values from the first row.
        Args:
            worksheet (object): The worksheet object to modify.
        """
        num_columns = worksheet.Columns.Count
        for col in range(1, num_columns + 1):
            first_row_value = worksheet.Cells(1, col)
            worksheet.Columns(col).Name = str(first_row_value)
        # Optionally delete the first row after setting names
        # worksheet.Rows(1).Delete()  # Uncomment if desired

    def create_scatter_plot(self, page, worksheet, x_col, y_col):
        """
        Create a scatter plot on the specified page using data from the worksheet.
        Args:
            page (object): The graph page object to add the plot to.
            worksheet (object): The worksheet containing the data.
            x_col (int): Column index for X data (1-based).
            y_col (int): Column index for Y data (1-based).
        Returns:
            object: The created plot object.
        """
        plot = page.Plots.Add(1)
        plot.DataSource = (worksheet.Columns(x_col), worksheet.Columns(y_col))
        return plot

    def set_font(self, graph, font_name="Arial"):
        """
        Set the font for various elements of the graph.
        Args:
            graph (object): The graph object to style.
            font_name (str): Name of the font (default: "Arial").
        """
        try:
            graph.Font = font_name
            graph.Title.Font = font_name
            for axis in graph.Axes:
                axis.Label.Font = font_name
                axis.TickLabels.Font = font_name
            if graph.Legends.Count > 0:
                graph.Legends(1).Font = font_name
        except Exception as e:
            print(f"Error setting font: {e}")

    def set_font_sizes(
        self, graph, title_size=8, label_size=8, tick_size=7, legend_size=6
    ):
        """
        Set font sizes for various graph elements.
        Args:
            graph (object): The graph object to style.
            title_size (int): Font size for the title (default: 8).
            label_size (int): Font size for axis labels (default: 8).
            tick_size (int): Font size for tick labels (default: 7).
            legend_size (int): Font size for the legend (default: 6).
        """
        try:
            graph.Title.FontSize = title_size
            for axis in graph.Axes:
                axis.Label.FontSize = label_size
                axis.TickLabels.FontSize = tick_size
            if graph.Legends.Count > 0:
                graph.Legends(1).FontSize = legend_size
        except Exception as e:
            print(f"Error setting font sizes: {e}")

    def set_tick_properties(self, graph, length=0.8, thickness=0.2):
        """
        Set tick length and thickness for the graph axes.
        Args:
            graph (object): The graph object to style.
            length (float): Tick length in mm (default: 0.8).
            thickness (float): Tick thickness in mm (default: 0.2).
        """
        try:
            for axis in graph.Axes:
                axis.MajorTicks.Length = length
                axis.MajorTicks.Thickness = thickness
        except Exception as e:
            print(f"Error setting tick properties: {e}")

    def hide_top_right_axes(self, graph):
        """
        Hide the top X-axis and right Y-axis of the graph.
        Args:
            graph (object): The graph object to modify.
        """
        try:
            if graph.Axes.Count >= 4:
                graph.Axes(3).Visible = False
                graph.Axes(4).Visible = False
        except Exception as e:
            print(f"Error hiding axes: {e}")

    def set_plot_color(self, plot, color):
        """
        Set the color of the plot.
        Args:
            plot (object): The plot object to modify.
            color (int): Color value (e.g., RGB integer like 255 for blue).
        """
        try:
            plot.Color = color
        except Exception as e:
            print(f"Error setting plot color: {e}")

    def set_figure_size(self, page, size="S"):
        """
        Set the size of the graph page.
        Args:
            page (object): The graph page object to resize.
            size (str): Size preset ("S", "M", "L") (default: "S").
        """
        sizes = {"S": (100, 100), "M": (200, 200), "L": (300, 300)}
        width, height = sizes.get(size, (200, 200))
        try:
            page.Width = width
            page.Height = height
        except Exception as e:
            print(f"Error setting figure size: {e}")

    def save_notebook(self, notebook, file_path):
        """
        Save the notebook to a file.
        Args:
            notebook (object): The notebook object to save.
            file_path (str): Path where the notebook will be saved.
        """
        win_path = os.path.normpath(file_path)
        try:
            notebook.SaveAs(win_path)
        except Exception as e:
            print(f"Error saving notebook: {e}")

    def export_graph(self, graph, file_path, format="PNG"):
        """
        Export the graph as an image.
        Args:
            graph (object): The graph object to export.
            file_path (str): Path where the image will be saved.
            format (str): Image format (default: "PNG").
        """
        win_path = os.path.normpath(file_path)
        try:
            graph.Export(win_path, format)
        except Exception as e:
            print(f"Error exporting graph: {e}")

    def close(self):
        """
        Close the SigmaPlot application gracefully.
        """
        if self.sp:
            try:
                self.sp.Quit()
            except Exception as e:
                print(f"Error closing SigmaPlot: {e}")

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
            ("New", []),
            ("Open", ["example.jnb"]),
            ("Save", []),
            ("GetObject", ["Sheet1"]),
            ("Activate", [])
        ]

        print(f"\nTrying common methods for {name}:")
        for method_name, args in common_methods:
            try:
                method = getattr(obj, method_name)
                print(f"  Method {method_name} exists")
                # Only try to call methods with no args or explicit args
                if not args:
                    print(f"  Trying to call {method_name}()")
                    result = method()
                    print(f"  Result: {str(result)[:100]}")
            except AttributeError:
                print(f"  Method {method_name} not available")
            except Exception as e:
                print(f"  Error calling {method_name}: {type(e).__name__}: {str(e)[:100]}")

        return True


    def explore_sigmaplot_specific(self, sp_app):
        """
        Explore SigmaPlot-specific properties and methods that might be different from
        standard Office COM objects.
        """
        print("\nExploring SigmaPlot-specific functionality:")

        # Try to access notebook collection
        try:
            notebooks = sp_app.Notebooks
            print(f"Notebooks property exists, Count: {notebooks.Count}")

            # Try to get information about each notebook
            for i in range(1, notebooks.Count + 1):
                try:
                    nb = notebooks.Item(i)
                    print(f"  Notebook {i} Name: {nb.Name}")
                except Exception as e:
                    print(f"  Error accessing notebook {i}: {e}")
        except Exception as e:
            print(f"Error accessing Notebooks: {e}")

        # Try to access ActiveDocument parts
        try:
            active_doc = sp_app.ActiveDocument
            print(f"\nActiveDocument Name: {active_doc.Name}")

            # Try common notebook collection properties
            collections = [
                "Worksheets", "GraphPages", "Sections", "NotebookItems",
                "Items", "Pages", "Sheets"
            ]

            for collection in collections:
                try:
                    coll_obj = getattr(active_doc, collection)
                    count = getattr(coll_obj, "Count", "unknown")
                    print(f"  Collection '{collection}' exists, Count: {count}")

                    # Try to access first item if available
                    if hasattr(coll_obj, "Item") and hasattr(coll_obj, "Count") and coll_obj.Count > 0:
                        item = coll_obj.Item(1)
                        print(f"    First item type: {type(item)}")
                        if hasattr(item, "Name"):
                            print(f"    First item name: {item.Name}")
                except Exception as e:
                    print(f"  Collection '{collection}' not available: {type(e).__name__}")

            # Try common worksheet methods
            worksheet_methods = [
                "CreateWorksheet", "AddWorksheet", "NewWorksheet",
                "CreateSection", "AddSection", "NewSection"
            ]

            print("\nTrying worksheet creation methods:")
            for method_name in worksheet_methods:
                try:
                    method = getattr(active_doc, method_name)
                    print(f"  Method {method_name} exists")
                except AttributeError:
                    print(f"  Method {method_name} not available")

            # Try notebook save/export methods
            print("\nTrying document save/export methods:")
            save_methods = ["SaveAs", "Export", "ExportGraph"]
            for method_name in save_methods:
                try:
                    method = getattr(active_doc, method_name)
                    print(f"  Method {method_name} exists")
                except AttributeError:
                    print(f"  Method {method_name} not available")

        except Exception as e:
            print(f"Error exploring ActiveDocument details: {e}")

        return True

    def explore_notebook_items(self, active_doc):
        """
        Specifically explore the NotebookItems collection which seems to be available
        but with access issues.
        """
        print("\nExploring NotebookItems in detail:")
        try:
            items = active_doc.NotebookItems
            count = items.Count
            print(f"NotebookItems Count: {count}")

            # Try different ways to access items
            print("\nTrying different access methods:")

            # Method 1: Try to access by index using Item method
            try:
                for i in range(1, count + 1):
                    try:
                        item = items.Item(i)
                        print(f"  Item({i}) accessed successfully")

                        # Try to get common properties
                        try:
                            if hasattr(item, "Name"):
                                print(f"    Name: {item.Name}")
                            if hasattr(item, "Type"):
                                print(f"    Type: {item.Type}")
                            if hasattr(item, "ObjectType"):
                                print(f"    ObjectType: {item.ObjectType}")
                        except Exception as e:
                            print(f"    Error accessing item properties: {e}")

                    except Exception as e:
                        print(f"  Error accessing Item({i}): {e}")
            except Exception as e:
                print(f"  Error in item loop: {e}")

            # Method 2: Try to get by name
            try:
                # Try with some common worksheet names
                for name in ["Sheet1", "Worksheet1", "Graph1"]:
                    try:
                        item = items(name)
                        print(f"  Items('{name}') accessed successfully")
                    except Exception as e:
                        print(f"  Could not access Items('{name}'): {type(e).__name__}")
            except Exception as e:
                print(f"  Error trying named access: {e}")

            # Method 3: Try some known methods that might create notebook items
            methods_to_try = [
                ("Add", [1]),
                ("AddWorksheet", []),
                ("AddSection", []),
                ("Item", [1])
            ]

            print("\nTrying NotebookItems methods:")
            for method_name, args in methods_to_try:
                try:
                    method = getattr(items, method_name)
                    print(f"  Method {method_name} exists")
                    if args:
                        result = method(*args)
                        print(f"  Called {method_name}{tuple(args)} successfully")
                        if hasattr(result, "Name"):
                            print(f"    Result Name: {result.Name}")
                    else:
                        print(f"  Not calling {method_name} (no args specified)")
                except AttributeError:
                    print(f"  Method {method_name} not available")
                except Exception as e:
                    print(f"  Error calling {method_name}: {e}")

        except Exception as e:
            print(f"Error exploring NotebookItems: {e}")

        return True

    def create_demo_plot(self, filename=None):
        """
        Create a demo plot with sample data and optionally save it.

        Args:
            filename (str): Optional path to save the notebook.
        """
        try:
            print("Creating demo plot...")

            # Get active document
            notebook = self.sp.ActiveDocument
            print(f"Using notebook: {notebook.Name}")

            # First, let's try to understand what types are available
            items = notebook.NotebookItems
            print(f"NotebookItems count: {items.Count}")

            # Try to discover what's inside each type
            for type_code in range(1, 5):
                try:
                    print(f"Attempting to create an item with type code {type_code}...")
                    test_item = items.Add(type_code)
                    print(f"Created item with type {type_code}, Name: {test_item.Name}")

                    # Try to find methods and properties
                    self.explore_com_object(test_item, f"Item_Type_{type_code}")

                    # Clean up by removing this test item (if possible)
                    try:
                        test_item.Delete()
                        print(f"Deleted test item with type {type_code}")
                    except:
                        print(f"Could not delete test item with type {type_code}")
                except Exception as e:
                    print(f"Error creating item with type {type_code}: {e}")

            # Create a worksheet for our actual data
            worksheet = items.Add(1)
            worksheet_name = "FinalData"
            worksheet.Name = worksheet_name
            print(f"Created worksheet: {worksheet.Name}")

            # Create a graph page
            graph_page = items.Add(2)
            graph_page.Name = "FinalGraph"
            print(f"Created graph page: {graph_page.Name}")

            # Try to use SendKeys to trigger SigmaPlot actions that we can't access via COM
            try:
                # Activate our worksheet
                worksheet.Activate()

                # Try to use Execute method if available
                try:
                    # Some COM applications have Execute method for running commands
                    commands = [
                        'SetColumnName "X Values", 1',
                        'SetColumnName "Y Values", 2',
                        'SetCellValue 1, 1, 1',
                        'SetCellValue 2, 1, 2',
                        'SetCellValue 3, 1, 3',
                        'SetCellValue 4, 1, 4',
                        'SetCellValue 5, 1, 5',
                        'SetCellValue 1, 2, 1',
                        'SetCellValue 2, 2, 4',
                        'SetCellValue 3, 2, 9',
                        'SetCellValue 4, 2, 16',
                        'SetCellValue 5, 2, 25'
                    ]

                    for cmd in commands:
                        try:
                            notebook.Execute(cmd)
                            print(f"Executed: {cmd}")
                        except:
                            print(f"Could not execute: {cmd}")
                except:
                    print("Execute method not available")

                # Activate graph page to work with it
                graph_page.Activate()

                # Try to create a plot through UI automation if API doesn't work
                try:
                    import win32com.client
                    shell = win32com.client.Dispatch("WScript.Shell")

                    # Wait for UI to stabilize
                    time.sleep(1)

                    # Try a simple keyboard shortcut to create a plot
                    # Alt+G, C might open the Create Graph dialog
                    shell.SendKeys("%g")
                    time.sleep(0.5)
                    shell.SendKeys("c")
                    time.sleep(2)

                    # In the dialog, press Tab to navigate and Enter to confirm
                    shell.SendKeys("{TAB}{TAB}{ENTER}")

                    print("Attempted to create plot using UI automation")
                except Exception as e:
                    print(f"Error using UI automation: {e}")

            except Exception as e:
                print(f"Error manipulating worksheet: {e}")

            # Save the notebook if filename is provided
            if filename:
                try:
                    notebook.SaveAs(os.path.normpath(filename))
                    print(f"Saved notebook to: {filename}")
                except Exception as e:
                    print(f"Error saving notebook: {e}")

            return True

        except Exception as e:
            print(f"Error in create_demo_plot: {e}")
            return False


if __name__ == "__main__":
    automator = None
    try:
        # Initialize the automator
        print("Initializing SigmaPlotAutomator...")
        automator = SigmaPlotAutomator(visible=True, close_others=False)
        sp = automator.sp

        # Run the demo
        automator.create_demo_plot("C:\\Temp\\SigmaPlotDemo.JNB")

        print("\nExploring SigmaPlot interface...")
        automator.explore_com_object(sp, "SigmaPlot.Application")

        # Also explore the ActiveDocument if available
        try:
            active_doc = sp.ActiveDocument
            automator.explore_com_object(active_doc, "ActiveDocument")

            # Explore SigmaPlot-specific functionality
            automator.explore_sigmaplot_specific(sp)

            # Explore notebook items specifically
            automator.explore_notebook_items(active_doc)

        except Exception as e:
            print(f"Error exploring ActiveDocument: {e}")

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Ensure SigmaPlot is closed
        if automator:
            print("\nClosing SigmaPlot...")
            automator.close()
        print("Script execution completed")

# (wsl) PySigMacro $ python.exe examples/tmp.py
# Initializing SigmaPlotAutomator...
# SigmaPlot initialized with 3 open notebooks

# Exploring SigmaPlot interface...

# Deep exploration of SigmaPlot.Application
# Type: <class 'win32com.client.CDispatch'>
# Could not get type info: (-2147352565, 'Invalid index.', None, None)

# Checking common properties for SigmaPlot.Application:
#   Application: <COMObject <unknown>>
#   Parent: <COMObject <unknown>>
#   Name: SigmaPlot 15
#   Type not available: AttributeError
#   Value not available: AttributeError
#   Count not available: AttributeError

# Trying common methods for SigmaPlot.Application:
#   Method New not available
#   Method Open not available
#   Method Save not available
#   Method GetObject not available
#   Method Activate not available

# Deep exploration of ActiveDocument
# Type: <class 'win32com.client.CDispatch'>
# Could not get type info: (-2147352565, 'Invalid index.', None, None)

# Checking common properties for ActiveDocument:
#   Application: <COMObject <unknown>>
#   Parent not available: com_error
#   Name: Notebook1
#   Type not available: AttributeError
#   Value not available: AttributeError
#   Count not available: AttributeError

# Trying common methods for ActiveDocument:
#   Method New not available
#   Method Open not available
#   Method Save exists
#   Trying to call Save()
#   Error calling Save: com_error: (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', "No file name assigned. Use 'SaveAs.'",
#   Method GetObject not available
#   Method Activate exists
#   Trying to call Activate()
#   Result: None

# Closing SigmaPlot...
# Script execution completed
# (wsl) PySigMacro $ python.exe examples/tmp.py
# Initializing SigmaPlotAutomator...
# SigmaPlot initialized with 5 open notebooks

# Exploring SigmaPlot interface...

# Deep exploration of SigmaPlot.Application
# Type: <class 'win32com.client.CDispatch'>
# Could not get type info: (-2147352565, 'Invalid index.', None, None)

# Checking common properties for SigmaPlot.Application:
#   Application: <COMObject <unknown>>
#   Parent: <COMObject <unknown>>
#   Name: SigmaPlot 15
#   Type not available: AttributeError
#   Value not available: AttributeError
#   Count not available: AttributeError

# Trying common methods for SigmaPlot.Application:
#   Method New not available
#   Method Open not available
#   Method Save not available
#   Method GetObject not available
#   Method Activate not available

# Deep exploration of ActiveDocument
# Type: <class 'win32com.client.CDispatch'>
# Could not get type info: (-2147352565, 'Invalid index.', None, None)

# Checking common properties for ActiveDocument:
#   Application: <COMObject <unknown>>
#   Parent not available: com_error
#   Name: Notebook1
# (wsl) PySigMacro $ python.exe examples/tmp.py
# Initializing SigmaPlotAutomator...
# SigmaPlot initialized with 5 open notebooks

# Exploring SigmaPlot interface...

# Deep exploration of SigmaPlot.Application
# Type: <class 'win32com.client.CDispatch'>
# Could not get type info: (-2147352565, 'Invalid index.', None, None)

# Checking common properties for SigmaPlot.Application:
#   Application: <COMObject <unknown>>
#   Parent: <COMObject <unknown>>
#   Name: SigmaPlot 15
#   Type not available: AttributeError
#   Value not available: AttributeError
#   Count not available: AttributeError

# Trying common methods for SigmaPlot.Application:
#   Method New not available
#   Method Open not available
#   Method Save not available
#   Method GetObject not available
#   Method Activate not available

# Deep exploration of ActiveDocument
# Type: <class 'win32com.client.CDispatch'>
# Could not get type info: (-2147352565, 'Invalid index.', None, None)

# Checking common properties for ActiveDocument:
#   Application: <COMObject <unknown>>
#   Parent not available: com_error
#   Name: Notebook1
# (wsl) PySigMacro $ python.exe examples/tmp.py
# Initializing SigmaPlotAutomator...
# SigmaPlot initialized with 5 open notebooks

# Exploring SigmaPlot interface...

# Deep exploration of SigmaPlot.Application
# Type: <class 'win32com.client.CDispatch'>
# Could not get type info: (-2147352565, 'Invalid index.', None, None)

# Checking common properties for SigmaPlot.Application:
#   Application: <COMObject <unknown>>
#   Parent: <COMObject <unknown>>
#   Name: SigmaPlot 15
#   Type not available: AttributeError
#   Value not available: AttributeError
#   Count not available: AttributeError

# Trying common methods for SigmaPlot.Application:
#   Method New not available
#   Method Open not available
#   Method Save not available
#   Method GetObject not available
#   Method Activate not available

# Deep exploration of ActiveDocument
# Type: <class 'win32com.client.CDispatch'>
# Could not get type info: (-2147352565, 'Invalid index.', None, None)

# Checking common properties for ActiveDocument:
#   Application: <COMObject <unknown>>
#   Parent not available: com_error
#   Name: Notebook1
#   Type not available: AttributeError
#   Value not available: AttributeError
#   Count not available: AttributeError

# Trying common methods for ActiveDocument:
#   Method New not available
#   Method Open not available
#   Method Save exists
#   Trying to call Save()
#   Error calling Save: com_error: (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', "No file name assigned. Use 'SaveAs.'",
#   Method GetObject not available
#   Method Activate exists
#   Trying to call Activate()
#   Result: None

# Exploring SigmaPlot-specific functionality:
# Notebooks property exists, Count: 5
#   Error accessing notebook 1: (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)
# (wsl) PySigMacro $ python.exe examples/tmp.py
# Initializing SigmaPlotAutomator...
# SigmaPlot initialized with 5 open notebooks
# Creating demo plot...
# Using notebook: Notebook1
# NotebookItems count: 3
# Attempting to create an item with type code 1...
# Created item with type 1, Name: Data 2

# Deep exploration of Item_Type_1
# Type: <class 'win32com.client.CDispatch'>
# Could not get type info: (-2147352565, 'Invalid index.', None, None)

# Checking common properties for Item_Type_1:
#   Application: <COMObject <unknown>>
#   Parent not available: com_error
#   Name: Data 2
#   Type not available: AttributeError
#   Value not available: AttributeError
#   Count not available: AttributeError

# Trying common methods for Item_Type_1:
#   Method New not available
#   Method Open exists
#   Method Save exists
#   Trying to call Save()
#   Result: None
#   Method GetObject not available
#   Method Activate not available
# Could not delete test item with type 1
# Attempting to create an item with type code 2...
# Created item with type 2, Name: Graph Page 1

# Deep exploration of Item_Type_2
# Type: <class 'win32com.client.CDispatch'>
# Could not get type info: (-2147352565, 'Invalid index.', None, None)

# Checking common properties for Item_Type_2:
#   Application: <COMObject <unknown>>
#   Parent not available: com_error
#   Name: Graph Page 1
#   Type not available: AttributeError
#   Value not available: AttributeError
#   Count not available: AttributeError

# Trying common methods for Item_Type_2:
#   Method New not available
#   Method Open exists
#   Method Save exists
#   Trying to call Save()
#   Result: None
#   Method GetObject not available
#   Method Activate not available
# Could not delete test item with type 2
# Attempting to create an item with type code 3...
# Created item with type 3, Name: Section 3

# Deep exploration of Item_Type_3
# Type: <class 'win32com.client.CDispatch'>
# Could not get type info: (-2147352565, 'Invalid index.', None, None)

# Checking common properties for Item_Type_3:
#   Application: <COMObject <unknown>>
#   Parent not available: com_error
#   Name: Section 3
#   Type not available: AttributeError
#   Value not available: AttributeError
#   Count not available: AttributeError

# Trying common methods for Item_Type_3:
#   Method New not available
#   Method Open exists
#   Method Save exists
#   Trying to call Save()
#   Result: None
#   Method GetObject not available
#   Method Activate not available
# Could not delete test item with type 3
# Attempting to create an item with type code 4...
# Error creating item with type 4: 'NoneType' object has no attribute 'Name'
# Created worksheet: FinalData
# Created graph page: FinalGraph
# Error manipulating worksheet: <unknown>.Activate
# Saved notebook to: C:\Temp\SigmaPlotDemo.JNB

# Exploring SigmaPlot interface...

# Deep exploration of SigmaPlot.Application
# Type: <class 'win32com.client.CDispatch'>
# Could not get type info: (-2147352565, 'Invalid index.', None, None)

# Checking common properties for SigmaPlot.Application:
#   Application: <COMObject <unknown>>
#   Parent: <COMObject <unknown>>
#   Name: SigmaPlot 15
#   Type not available: AttributeError
#   Value not available: AttributeError
#   Count not available: AttributeError

# Trying common methods for SigmaPlot.Application:
#   Method New not available
#   Method Open not available
#   Method Save not available
#   Method GetObject not available
#   Method Activate not available

# Deep exploration of ActiveDocument
# Type: <class 'win32com.client.CDispatch'>
# Could not get type info: (-2147352565, 'Invalid index.', None, None)

# Checking common properties for ActiveDocument:
#   Application: <COMObject <unknown>>
#   Parent not available: com_error
#   Name: SigmaPlotDemo.JNB
#   Type not available: AttributeError
#   Value not available: AttributeError
#   Count not available: AttributeError

# Trying common methods for ActiveDocument:
#   Method New not available
#   Method Open not available
#   Method Save exists
#   Trying to call Save()
#   Result: None
#   Method GetObject not available
#   Method Activate exists
#   Trying to call Activate()
#   Result: None

# Exploring SigmaPlot-specific functionality:
# Notebooks property exists, Count: 5
#   Error accessing notebook 1: (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)
#   Error accessing notebook 2: (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)
#   Error accessing notebook 3: (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)
#   Error accessing notebook 4: (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)
#   Error accessing notebook 5: (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)

# ActiveDocument Name: SigmaPlotDemo.JNB
#   Collection 'Worksheets' not available: AttributeError
#   Collection 'GraphPages' not available: AttributeError
#   Collection 'Sections' not available: AttributeError
#   Collection 'NotebookItems' exists, Count: 10
#   Collection 'NotebookItems' not available: com_error
#   Collection 'Items' not available: AttributeError
#   Collection 'Pages' not available: AttributeError
# (wsl) PySigMacro $ python.exe examples/tmp.py
# Initializing SigmaPlotAutomator...
# SigmaPlot initialized with 5 open notebooks
# Creating demo plot...
# Using notebook: SigmaPlotDemo.JNB
# NotebookItems count: 12
# Attempting to create an item with type code 1...
# Created item with type 1, Name: Data 5

# Deep exploration of Item_Type_1
# Type: <class 'win32com.client.CDispatch'>
# Could not get type info: (-2147352565, 'Invalid index.', None, None)

# Checking common properties for Item_Type_1:
#   Application: <COMObject <unknown>>
#   Parent not available: com_error
#   Name: Data 5
#   Type not available: AttributeError
#   Value not available: AttributeError
#   Count not available: AttributeError

# Trying common methods for Item_Type_1:
#   Method New not available
#   Method Open exists
#   Method Save exists
#   Trying to call Save()
#   Result: None
#   Method GetObject not available
#   Method Activate not available
# Could not delete test item with type 1
# Attempting to create an item with type code 2...
# Created item with type 2, Name: Graph Page 3

# Deep exploration of Item_Type_2
# Type: <class 'win32com.client.CDispatch'>
# Could not get type info: (-2147352565, 'Invalid index.', None, None)

# Checking common properties for Item_Type_2:
#   Application: <COMObject <unknown>>
#   Parent not available: com_error
#   Name: Graph Page 3
#   Type not available: AttributeError
#   Value not available: AttributeError
#   Count not available: AttributeError

# Trying common methods for Item_Type_2:
#   Method New not available
#   Method Open exists
#   Method Save exists
#   Trying to call Save()
#   Result: None
#   Method GetObject not available
#   Method Activate not available
# Could not delete test item with type 2
# Attempting to create an item with type code 3...
# Created item with type 3, Name: Section 7

# Deep exploration of Item_Type_3
# Type: <class 'win32com.client.CDispatch'>
# Could not get type info: (-2147352565, 'Invalid index.', None, None)

# Checking common properties for Item_Type_3:
#   Application: <COMObject <unknown>>
#   Parent not available: com_error
#   Name: Section 7
#   Type not available: AttributeError
#   Value not available: AttributeError
#   Count not available: AttributeError

# Trying common methods for Item_Type_3:
#   Method New not available
#   Method Open exists
#   Method Save exists
#   Trying to call Save()
#   Result: None
#   Method GetObject not available
#   Method Activate not available
# Could not delete test item with type 3
# Attempting to create an item with type code 4...
# Error creating item with type 4: 'NoneType' object has no attribute 'Name'
# Error in create_demo_plot: Property '<unknown>.Name' can not be set.

# Exploring SigmaPlot interface...

# Deep exploration of SigmaPlot.Application
# Type: <class 'win32com.client.CDispatch'>
# Could not get type info: (-2147352565, 'Invalid index.', None, None)

# Checking common properties for SigmaPlot.Application:
#   Application: <COMObject <unknown>>
#   Parent: <COMObject <unknown>>
#   Name: SigmaPlot 15
#   Type not available: AttributeError
#   Value not available: AttributeError
#   Count not available: AttributeError

# Trying common methods for SigmaPlot.Application:
#   Method New not available
#   Method Open not available
#   Method Save not available
#   Method GetObject not available
#   Method Activate not available

# Deep exploration of ActiveDocument
# Type: <class 'win32com.client.CDispatch'>
# Could not get type info: (-2147352565, 'Invalid index.', None, None)

# Checking common properties for ActiveDocument:
#   Application: <COMObject <unknown>>
#   Parent not available: com_error
#   Name: SigmaPlotDemo.JNB
#   Type not available: AttributeError
#   Value not available: AttributeError
#   Count not available: AttributeError

# Trying common methods for ActiveDocument:
#   Method New not available
#   Method Open not available
#   Method Save exists
#   Trying to call Save()
#   Result: None
#   Method GetObject not available
#   Method Activate exists
#   Trying to call Activate()
#   Result: None

# Exploring SigmaPlot-specific functionality:
# Notebooks property exists, Count: 5
#   Error accessing notebook 1: (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)
#   Error accessing notebook 2: (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)
#   Error accessing notebook 3: (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)
#   Error accessing notebook 4: (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)
#   Error accessing notebook 5: (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)

# ActiveDocument Name: SigmaPlotDemo.JNB
#   Collection 'Worksheets' not available: AttributeError
#   Collection 'GraphPages' not available: AttributeError
#   Collection 'Sections' not available: AttributeError
#   Collection 'NotebookItems' exists, Count: 18
#   Collection 'NotebookItems' not available: com_error
#   Collection 'Items' not available: AttributeError
#   Collection 'Pages' not available: AttributeError
#   Collection 'Sheets' not available: AttributeError

# Trying worksheet creation methods:
#   Method CreateWorksheet not available
#   Method AddWorksheet not available
#   Method NewWorksheet not available
#   Method CreateSection not available
#   Method AddSection not available
#   Method NewSection not available

# Trying document save/export methods:
#   Method SaveAs exists
#   Method Export not available
#   Method ExportGraph not available

# Exploring NotebookItems in detail:
# NotebookItems Count: 18

# Trying different access methods:
#   Error accessing Item(1): (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)
#   Error accessing Item(2): (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)
#   Error accessing Item(3): (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)
#   Error accessing Item(4): (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)
#   Error accessing Item(5): (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)
#   Error accessing Item(6): (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)
#   Error accessing Item(7): (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)
#   Error accessing Item(8): (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)
#   Error accessing Item(9): (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)
#   Error accessing Item(10): (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)
#   Error accessing Item(11): (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)
#   Error accessing Item(12): (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)
#   Error accessing Item(13): (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)
#   Error accessing Item(14): (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)
#   Error accessing Item(15): (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)
#   Error accessing Item(16): (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)
#   Error accessing Item(17): (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)
#   Error accessing Item(18): (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)
#   Items('Sheet1') accessed successfully
#   Items('Worksheet1') accessed successfully
#   Items('Graph1') accessed successfully

# Trying NotebookItems methods:
#   Method Add exists
#   Called Add(1,) successfully
#     Result Name: Data 7
#   Method AddWorksheet not available
#   Method AddSection not available
#   Error calling Item: (-2147352567, 'Exception occurred.', (65535, 'SigmaPlot 15', 'Only VT_EMPTY, VT_BSTR, VT_I2 and VT_I4 variants allowed', None, 0, 0), None)

# Closing SigmaPlot...
# Script execution completed

# EOF