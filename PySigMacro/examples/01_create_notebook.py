#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 10:02:42 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/examples/create_notebook.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/examples/create_notebook.py"

import os
import sys
import time
import win32com.client
from pysigmacro.core.connection import connect


sigmaplot = connect(visible=True, launch_if_not_found=True)
dir(sigmaplot)
sigmaplot.Visible = True

# Create a new notebook
notebook = sigmaplot.Notebooks.Add()
notebook.Name = "TestNotebook"



def create_notebook():
    """
    Creates a notebook and attempts to access its structure
    """
    try:
        # Connect to SigmaPlot
        print("Connecting to SigmaPlot...")
        sigmaplot = connect(visible=True, launch_if_not_found=True)
        print("Connected successfully.")

        # Give SigmaPlot a moment to fully initialize
        time.sleep(1)

        # Check if we have a Notebooks collection
        print("\nExploring Notebooks collection...")
        if hasattr(sigmaplot, 'Notebooks'):
            notebooks = sigmaplot.Notebooks
            count = notebooks.Count
            print(f"Found {count} notebooks")

            # Try to get active document instead of ActiveNotebook
            print("\nTrying to access active document...")
            notebook = sigmaplot.ActiveDocument
            if notebook:
                print(f"Active document name: {notebook.Name}")

                # Try to create a new worksheet (section might be called worksheet in SigmaPlot)
                print("\nAttempting to create a new worksheet...")
                try:
                    # Method 1: Using Notebook.CreateWorksheet method
                    if hasattr(notebook, 'CreateWorksheet'):
                        worksheet = notebook.CreateWorksheet()
                        print(f"Created worksheet using CreateWorksheet()")
                    # Method 2: Using Notebook.Worksheets.Add method
                    elif hasattr(notebook, 'Worksheets') and hasattr(notebook.Worksheets, 'Add'):
                        worksheet = notebook.Worksheets.Add()
                        print(f"Created worksheet using Worksheets.Add()")
                    # Method 3: Using Notebook.Sections.Add method
                    elif hasattr(notebook, 'Sections') and hasattr(notebook.Sections, 'Add'):
                        worksheet = notebook.Sections.Add()
                        print(f"Created worksheet using Sections.Add()")
                    else:
                        print("Could not find a method to create a worksheet")

                    # Try to access the collection of worksheets
                    print("\nAttempting to access worksheets...")
                    for collection_name in ['Worksheets', 'Sections', 'Pages', 'Sheets']:
                        if hasattr(notebook, collection_name):
                            collection = getattr(notebook, collection_name)
                            print(f"Found collection: {collection_name} with {collection.Count} items")

                            if collection.Count > 0:
                                item = collection.Item(1)
                                print(f"First {collection_name[:-1]} name: {item.Name}")

                                # Try to find a method to select this item
                                if hasattr(item, 'Select'):
                                    item.Select()
                                    print(f"Selected the first {collection_name[:-1]}")
                                elif hasattr(item, 'Activate'):
                                    item.Activate()
                                    print(f"Activated the first {collection_name[:-1]}")

                except Exception as e:
                    print(f"Error creating/accessing worksheet: {e}")

                # Get all notebook attributes and try to find worksheet-related ones
                print("\nInspecting notebook attributes:")
                for attr_name in dir(notebook):
                    if attr_name.startswith('_'):
                        continue

                    try:
                        attr = getattr(notebook, attr_name)
                        if callable(attr):
                            continue

                        if hasattr(attr, 'Count') and hasattr(attr, 'Item'):
                            # This looks like a collection
                            try:
                                count = attr.Count
                                print(f"Collection: {attr_name} with {count} items")

                                if count > 0:
                                    first_item = attr.Item(1)
                                    if hasattr(first_item, 'Name'):
                                        print(f"  First item name: {first_item.Name}")
                            except:
                                pass

                    except Exception as e:
                        # Ignore errors accessing properties
                        pass

                # Try to access sections directly through the notebook index
                print("\nAttempting to access sections through indexing:")
                try:
                    section = notebook[1]
                    print(f"Successfully accessed section through indexing: {section.Name}")
                except:
                    try:
                        # Try a different indexing style
                        section = notebook.Item(1)
                        print(f"Successfully accessed section through Item(1): {section.Name}")
                    except:
                        print("Could not access sections through indexing")

            else:
                print("Could not access active document")
        else:
            print("No Notebooks collection found")

        print("\nOperation completed")

    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    create_notebook()

# EOF