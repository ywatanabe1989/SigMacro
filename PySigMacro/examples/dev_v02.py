#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-13 09:37:36 (ywatanabe)"
# File: /home/ywatanabe/proj/SigMacro/PySigMacro/examples/dev.py

__THIS_FILE__ = "/home/ywatanabe/proj/SigMacro/PySigMacro/examples/dev.py"

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


    def create_sigmaplot_notebook(self, notebook_name=None, use_existing=True):
        """
        Create a SigmaPlot notebook with basic structure.

        Args:
            notebook_name (str): Optional name for the notebook
            use_existing (bool): Whether to use the existing active document

        Returns:
            object: The notebook object
        """
        try:
            # Use existing document or create new
            if use_existing:
                notebook = self.sp.ActiveDocument
                print(f"Using existing notebook: {notebook.Name}")
            else:
                # Try to create a new notebook if possible
                # Unfortunately direct method not found
                print("Creating new notebook not directly supported")
                notebook = self.sp.ActiveDocument

            # Set name if provided
            if notebook_name:
                try:
                    notebook.Name = notebook_name
                    print(f"Set notebook name to: {notebook_name}")
                except Exception as e:
                    print(f"Could not set notebook name: {e}")

            return notebook
        except Exception as e:
            print(f"Error in create_sigmaplot_notebook: {e}")
            return None

    def create_sigmaplot_sections(self, notebook, section_names=None):
        """
        Create sections in a SigmaPlot notebook.

        Args:
            notebook (object): The notebook object
            section_names (list): List of section names to create

        Returns:
            dict: Dictionary mapping section names to section objects
        """
        if section_names is None:
            section_names = ["Data", "Graphs", "Results"]

        sections = {}

        try:
            items = notebook.NotebookItems
            print(f"Creating {len(section_names)} sections...")

            for name in section_names:
                try:
                    # Create a section (type code 3)
                    section = items.Add(3)
                    section.Name = name
                    sections[name] = section
                    print(f"Created section: {name}")
                except Exception as e:
                    print(f"Error creating section '{name}': {e}")

            return sections
        except Exception as e:
            print(f"Error in create_sigmaplot_sections: {e}")
            return sections

    def create_sigmaplot_worksheets(self, notebook, worksheet_names=None):
        """
        Create worksheets in a SigmaPlot notebook.

        Args:
            notebook (object): The notebook object
            worksheet_names (list): List of worksheet names to create

        Returns:
            dict: Dictionary mapping worksheet names to worksheet objects
        """
        if worksheet_names is None:
            worksheet_names = ["RawData", "ProcessedData"]

        worksheets = {}

        try:
            items = notebook.NotebookItems
            print(f"Creating {len(worksheet_names)} worksheets...")

            for name in worksheet_names:
                try:
                    # Create a worksheet (type code 1)
                    worksheet = items.Add(1)
                    worksheet.Name = name
                    worksheets[name] = worksheet
                    print(f"Created worksheet: {name}")
                except Exception as e:
                    print(f"Error creating worksheet '{name}': {e}")

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
                    graph_page.Name = name
                    graph_pages[name] = graph_page
                    print(f"Created graph page: {name}")
                except Exception as e:
                    print(f"Error creating graph page '{name}': {e}")

            return graph_pages
        except Exception as e:
            print(f"Error in create_sigmaplot_graph_pages: {e}")
            return graph_pages

    def save_sigmaplot_notebook(self, notebook, file_path):
        """
        Save a SigmaPlot notebook to a file.

        Args:
            notebook (object): The notebook object
            file_path (str): Path to save the notebook

        Returns:
            bool: True if saved successfully, False otherwise
        """
        try:
            win_path = os.path.normpath(file_path)
            notebook.SaveAs(win_path)
            print(f"Saved notebook to: {win_path}")
            return True
        except Exception as e:
            print(f"Error saving notebook: {e}")
            return False

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

    def create_sigmaplot_sections(self, notebook, count=3):
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
                for type_code in range(1, 5):
                    try:
                        test_item = items.Add(type_code)
                        print(f"  Created item with type code {type_code}, name: {test_item.Name}")
                        type_names[type_code] = test_item.Name.split()[0]
                    except Exception as e:
                        print(f"  Error creating item with type code {type_code}: {e}")

                print("\nIdentified item types:")
                for code, name in type_names.items():
                    print(f"  Type {code}: {name}")

                structure['type_map'] = type_names

                # Try to see if items can contain other items
                print("\nChecking if items can contain other items:")
                for name, info in structure['named_items'].items():
                    try:
                        container_test = items.Add(3)
                        container_name = container_test.Name

                        # Try to activate the container
                        try:
                            container_test.Activate()

                            # Now try to add something inside it
                            try:
                                sub_item = items.Add(1)
                                sub_name = sub_item.Name
                                print(f"  Added {sub_name} inside {container_name} - container relationship works")
                                info['can_contain_items'] = True
                            except:
                                print(f"  Could not add items inside {container_name}")
                                info['can_contain_items'] = False
                        except:
                            print(f"  Could not activate {container_name} to test containment")
                    except:
                        print(f"  Error testing container relationship for {name}")

            except Exception as e:
                print(f"Error accessing NotebookItems: {e}")

            return structure
        except Exception as e:
            print(f"Error in explore_notebook_structures: {e}")
            return structure

    def add_data_to_worksheet(self, worksheet, data, column_names=None):
        """
        Add data to a worksheet using SendKeys and UI automation since direct COM access fails.

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

            # Use UI automation to add data
            shell = win32com.client.Dispatch("WScript.Shell")
            time.sleep(1)

            # First, set column names if provided
            if column_names:
                for col_idx, name in enumerate(column_names):
                    # Right-click on column header to open context menu
                    # This is a simplification - actual coordinates would be needed
                    # Instead we'll try keyboard shortcuts

                    # ALT+D, L might open the Column Properties dialog
                    shell.SendKeys("%d")
                    time.sleep(0.5)
                    shell.SendKeys("l")
                    time.sleep(1)

                    # Navigate to column number field and set it
                    shell.SendKeys(str(col_idx + 1))
                    shell.SendKeys("{TAB}")

                    # Set the name
                    shell.SendKeys(name)
                    shell.SendKeys("{ENTER}")

                    print(f"Set column {col_idx+1} name to '{name}'")
                    time.sleep(0.5)

            # Then add data - we'll try to use keyboard to navigate and input values
            for row_idx in range(len(data[0])):
                for col_idx in range(len(data)):
                    # Try keyboard shortcut to select cell
                    # ALT+E, G opens "Go To" dialog
                    shell.SendKeys("%e")
                    time.sleep(0.3)
                    shell.SendKeys("g")
                    time.sleep(0.5)

                    # Enter cell coordinates (row, column)
                    shell.SendKeys(f"{row_idx+1},{col_idx+1}")
                    shell.SendKeys("{ENTER}")
                    time.sleep(0.3)

                    # Enter the value
                    value = data[col_idx][row_idx]
                    shell.SendKeys(str(value))
                    shell.SendKeys("{ENTER}")

                    print(f"Set cell ({row_idx+1},{col_idx+1}) to {value}")

            return True
        except Exception as e:
            print(f"Error in add_data_to_worksheet: {e}")
            return False

    def create_plot_with_ui(self, graph_page, data_worksheet):
        """
        Create a plot on a graph page using UI automation since direct COM access fails.

        Args:
            graph_page (object): The graph page to create plot on
            data_worksheet (object): The worksheet containing data to plot

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

            # Use UI automation to create plot
            shell = win32com.client.Dispatch("WScript.Shell")
            time.sleep(1)

            # Try to use Create Graph dialog through keyboard shortcuts
            # ALT+G, C might open Create Graph dialog in SigmaPlot
            shell.SendKeys("%g")
            time.sleep(0.5)
            shell.SendKeys("c")
            time.sleep(1)

            # In the dialog, we'd need to select the worksheet, plot type, etc.
            # This is a simplified approach - might need adjustments
            # Select plot type (first option)
            shell.SendKeys("{TAB}")
            shell.SendKeys("{TAB}")
            shell.SendKeys(" ")
            time.sleep(0.5)

            # Move to Next button
            shell.SendKeys("{TAB}")
            shell.SendKeys("{TAB}")
            shell.SendKeys("{TAB}")
            shell.SendKeys("{ENTER}")
            time.sleep(0.5)

            # In data selection screen - try to navigate to select first two columns
            # This part is most likely to need customization based on actual UI
            shell.SendKeys("{TAB}")
            shell.SendKeys(" ")
            shell.SendKeys("{DOWN}")
            shell.SendKeys(" ")
            time.sleep(0.5)

            # Finish the wizard
            shell.SendKeys("{TAB}")
            shell.SendKeys("{TAB}")
            shell.SendKeys("{TAB}")
            shell.SendKeys("{TAB}")
            shell.SendKeys("{ENTER}")

            print("Created plot using UI automation")
            return True
        except Exception as e:
            print(f"Error in create_plot_with_ui: {e}")
            return False

    def demo_data_and_plot(self, file_path="C:\\Temp\\SigmaPlot_Demo.JNB"):
        """
        Create a complete demo with data and a plot.

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
            success = self.create_plot_with_ui(graph_page, worksheet)
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

        # Create a demo with data and plot
        automator.demo_data_and_plot("C:\\Temp\\SigmaPlot_DataDemo_" + time.strftime("%Y%m%d_%H%M%S") + ".JNB")

        # # Create a notebook
        # notebook = automator.create_sigmaplot_notebook("MyProject")

        # # Create sections
        # sections = automator.create_sigmaplot_sections(notebook, ["Introduction", "Methods", "Results"])

        # # Create worksheets
        # worksheets = automator.create_sigmaplot_worksheets(notebook, ["Dataset1", "Dataset2"])

        # # Create graph pages
        # graphs = automator.create_sigmaplot_graph_pages(notebook, ["ScatterPlot", "BarChart"])

        # # Save the notebook
        # automator.save_sigmaplot_notebook(notebook, "C:\\Temp\\SigmaPlotProject.JNB")


        # # Run the demo
        # automator.create_demo_plot("C:\\Temp\\SigmaPlotDemo.JNB")

        # print("\nExploring SigmaPlot interface...")
        # automator.explore_com_object(sp, "SigmaPlot.Application")

        # # Also explore the ActiveDocument if available
        # try:
        #     active_doc = sp.ActiveDocument
        #     automator.explore_com_object(active_doc, "ActiveDocument")

        #     # Explore SigmaPlot-specific functionality
        #     automator.explore_sigmaplot_specific(sp)

        #     # Explore notebook items specifically
        #     automator.explore_notebook_items(active_doc)

        # except Exception as e:
        #     print(f"Error exploring ActiveDocument: {e}")

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