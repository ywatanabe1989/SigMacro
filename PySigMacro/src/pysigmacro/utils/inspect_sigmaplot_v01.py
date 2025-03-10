#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 09:56:14 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/utils/inspect_sigmaplot.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/utils/inspect_sigmaplot.py"

"""
SigmaPlot-specific COM object inspection utilities.
"""

import sys
import time
import logging
import json
from pysigmacro.utils.inspect_com import (
    explore_com_object,
    to_json_serializable,
    get_com_property,
    call_com_method,
)

# Set up logging
logger = logging.getLogger(__name__)


def inspect_sigmaplot(sigmaplot_obj=None, depth=2, explore_collections=True):
    """
    Main entry point for SigmaPlot COM object inspection.

    Args:
        sigmaplot_obj: SigmaPlot COM object. If None, will try to connect
        depth: Depth of exploration for object structure
        explore_collections: Whether to explore notebook collections

    Returns:
        Dictionary with inspection results
    """
    from pysigmacro.core.connection import connect
    import time
    import win32com.client
    import pythoncom

    results = {}

    # Connect to SigmaPlot if not provided
    if sigmaplot_obj is None:
        logger.info("Connecting to SigmaPlot...")
        sigmaplot_obj = connect(visible=True, launch_if_not_found=True)
        logger.info("Connected successfully.")
        time.sleep(1)

    # Basic object structure
    logger.info("Exploring SigmaPlot application object...")
    results['application'] = explore_com_object(
        sigmaplot_obj, "SigmaPlot", max_depth=min(2, depth)
    )

    # For SigmaPlot specifically, enumerate properties we know are safe
    safe_properties = [
        'Name', 'Version', 'Path', 'ActiveDocument', 'Visible',
        'Height', 'Width', 'Left', 'Top'
    ]

    # Get safe properties
    app_properties = {}
    for prop in safe_properties:
        try:
            if hasattr(sigmaplot_obj, prop):
                value = getattr(sigmaplot_obj, prop)
                app_properties[prop] = value
        except Exception as e:
            logger.debug(f"Could not get property {prop}: {e}")

    results['app_properties'] = app_properties

    # Explore Notebooks collection if available
    if explore_collections and hasattr(sigmaplot_obj, 'Notebooks'):
        logger.info("Exploring Notebooks collection...")
        try:
            notebooks = sigmaplot_obj.Notebooks
            count = notebooks.Count
            results['notebooks'] = {
                'count': count,
                'items': {}
            }

            # Try different ways to access notebooks
            for i in range(1, count + 1):
                try:
                    # Try different access methods
                    notebook = None

                    # Method 1: Try with integer
                    try:
                        notebook = notebooks.Item(i)
                    except:
                        pass

                    # Method 2: Try with VT_I4 variant
                    if notebook is None:
                        try:
                            idx = win32com.client.VARIANT(pythoncom.VT_I4, i)
                            notebook = notebooks.Item(idx)
                        except:
                            pass

                    # Method 3: Try with string index
                    if notebook is None:
                        try:
                            notebook = notebooks.Item(f"Notebook{i}")
                        except:
                            pass

                    # Method 4: Try with indexing syntax
                    if notebook is None:
                        try:
                            notebook = notebooks[i]
                        except:
                            pass

                    if notebook:
                        # Just store minimal info
                        notebook_info = {
                            'type': str(type(notebook))
                        }

                        # Try to get name
                        try:
                            if hasattr(notebook, 'Name'):
                                notebook_info['name'] = notebook.Name
                        except:
                            notebook_info['name'] = f"Notebook{i}"

                        # Get additional properties
                        try:
                            # Test for worksheet access
                            if hasattr(notebook, 'Worksheets'):
                                worksheets = notebook.Worksheets
                                if hasattr(worksheets, 'Count'):
                                    notebook_info['worksheet_count'] = worksheets.Count
                        except Exception as ex:
                            logger.debug(f"Error checking worksheets: {ex}")

                        results['notebooks']['items'][f"Notebook{i}"] = notebook_info
                    else:
                        results['notebooks']['items'][f"Notebook{i}"] = "Could not access notebook"
                except Exception as e:
                    logger.error(f"Error exploring notebook {i}: {e}")
                    results['notebooks']['items'][f"Notebook{i}"] = str(e)
        except Exception as e:
            logger.error(f"Error exploring notebooks: {e}")
            results['notebooks'] = str(e)

    # Try to access active document/notebook with minimum exploration
    logger.info("Checking active document...")
    try:
        active_doc = sigmaplot_obj.ActiveDocument
        if active_doc:
            active_info = {
                'type': str(type(active_doc))
            }

            # Try to get name and other safe properties
            try:
                if hasattr(active_doc, 'Name'):
                    active_info['name'] = active_doc.Name
            except:
                active_info['name'] = "ActiveDocument"

            # Try to get window properties
            for prop in ['Top', 'Left', 'Height', 'Width', 'Visible']:
                try:
                    if hasattr(active_doc, prop):
                        active_info[prop] = getattr(active_doc, prop)
                except:
                    pass

            # Try to get document-specific properties
            document_properties = [
                'Path', 'FullName', 'Modified', 'ReadOnly'
            ]
            for prop in document_properties:
                try:
                    if hasattr(active_doc, prop):
                        active_info[prop] = getattr(active_doc, prop)
                except:
                    pass

            results['active_document'] = active_info

            # Try to explore collections in the active document
            for coll_name in ['Worksheets', 'Sections', 'Pages', 'Graphs']:
                try:
                    if hasattr(active_doc, coll_name):
                        collection = getattr(active_doc, coll_name)
                        if hasattr(collection, 'Count'):
                            count = collection.Count

                            if 'collections' not in results['active_document']:
                                results['active_document']['collections'] = {}

                            results['active_document']['collections'][coll_name] = {
                                'count': count
                            }

                            # Try to access first item if any
                            if count > 0:
                                try:
                                    # Try different methods to access items
                                    item = None

                                    # Try with direct index
                                    try:
                                        item = collection.Item(1)
                                    except:
                                        pass

                                    # Try with VT_I4 variant
                                    if item is None:
                                        try:
                                            idx = win32com.client.VARIANT(pythoncom.VT_I4, 1)
                                            item = collection.Item(idx)
                                        except:
                                            pass

                                    if item:
                                        item_info = {'type': str(type(item))}

                                        # Try to get name
                                        try:
                                            if hasattr(item, 'Name'):
                                                item_info['name'] = item.Name
                                        except:
                                            pass

                                        results['active_document']['collections'][coll_name]['first_item'] = item_info
                                except Exception as ex:
                                    logger.debug(f"Error exploring first item of {coll_name}: {ex}")
                except Exception as e:
                    logger.debug(f"Error exploring collection {coll_name}: {e}")
    except Exception as e:
        logger.error(f"Error getting active document: {e}")
        results['active_document'] = str(e)

    # Try notebook methods
    logger.info("Testing notebook operations...")
    try:
        results['notebook_operations'] = {}

        # Test if we can activate a notebook
        if hasattr(sigmaplot_obj, 'Notebooks') and sigmaplot_obj.Notebooks.Count > 1:
            try:
                # Try to activate notebook 2
                current_active = sigmaplot_obj.ActiveDocument.Name
                notebooks = sigmaplot_obj.Notebooks
                notebook2 = None

                # Try different methods to get notebook 2
                try:
                    notebook2 = notebooks.Item(2)
                except:
                    try:
                        idx = win32com.client.VARIANT(pythoncom.VT_I4, 2)
                        notebook2 = notebooks.Item(idx)
                    except:
                        pass

                if notebook2 and hasattr(notebook2, 'Activate'):
                    notebook2.Activate()
                    new_active = sigmaplot_obj.ActiveDocument.Name
                    results['notebook_operations']['activate'] = {
                        'success': new_active != current_active,
                        'previous': current_active,
                        'current': new_active
                    }
            except Exception as e:
                logger.error(f"Error activating notebook: {e}")
                results['notebook_operations']['activate'] = {'error': str(e)}

        # Test notebook creation methods
        operations = [
            ('FileNew', []),
            ('NewNotebook', []),
            ('AddNotebook', []),
            ('Execute', ['FileNew']),
        ]

        for method_name, args in operations:
            if hasattr(sigmaplot_obj, method_name):
                try:
                    logger.info(f"Trying method: {method_name}")
                    method = getattr(sigmaplot_obj, method_name)
                    method(*args)
                    results['notebook_operations'][method_name] = {'success': True}
                except Exception as e:
                    logger.error(f"Error with method {method_name}: {e}")
                    results['notebook_operations'][method_name] = {'error': str(e)}
    except Exception as e:
        logger.error(f"Error in notebook operations test: {e}")
        results['notebook_operations'] = {'error': str(e)}

    return results

def print_inspection_results(results):
    """
    Print inspection results in a readable format

    Args:
        results: Results from inspect_sigmaplot()

    Returns:
        None
    """
    print("\n===== SigmaPlot COM Inspection Results =====\n")

    # Print application structure
    if "application" in results:
        print("APPLICATION STRUCTURE:")
        print("-" * 50)
        print(f"Type: {results['application'].get('type', 'Unknown')}")

        # Print methods
        methods = results["application"].get("methods", [])
        if methods:
            print(f"\nMethods ({len(methods)}):")
            for method in sorted(methods):
                print(f"  - {method}")

        # Print properties
        properties = results["application"].get("properties", {})
        if properties:
            print(f"\nProperties ({len(properties)}):")
            for prop, prop_type in sorted(properties.items()):
                print(f"  - {prop}: {prop_type}")

        # Print collections
        collections = results["application"].get("collections", {})
        if collections:
            print(f"\nCollections ({len(collections)}):")
            for coll_name, coll_data in sorted(collections.items()):
                print(f"  - {coll_name}: {coll_data.get('count', 0)} items")

    # Print notebooks information
    if "notebooks" in results and isinstance(results["notebooks"], dict):
        print("\nNOTEBOOKS:")
        print("-" * 50)
        count = results["notebooks"].get("count", 0)
        print(f"Count: {count}")

        items = results["notebooks"].get("items", {})
        if items:
            print("\nNotebook items:")
            for name, data in items.items():
                print(f"  - {name}")

                # Print collections in notebook
                if isinstance(data, dict) and "collections" in data:
                    collections = data["collections"]
                    for coll_name, coll_data in sorted(collections.items()):
                        print(
                            f"    * {coll_name}: {coll_data.get('count', 0)} items"
                        )

    # Print active document info
    if "active_document" in results:
        print("\nACTIVE DOCUMENT:")
        print("-" * 50)
        if isinstance(results["active_document"], dict):
            print(f"Name: {results['active_document'].get('name', 'Unknown')}")
            print(f"Type: {results['active_document'].get('type', 'Unknown')}")

            # Print collections in active document
            if "collections" in results["active_document"]:
                collections = results["active_document"]["collections"]
                print("\nCollections:")
                for coll_name, coll_data in sorted(collections.items()):
                    print(
                        f"  - {coll_name}: {coll_data.get('count', 0)} items"
                    )
        else:
            print(f"Error: {results['active_document']}")

    # Print notebook creation test results
    if "notebook_creation" in results:
        print("\nNOTEBOOK CREATION TEST:")
        print("-" * 50)
        if isinstance(results["notebook_creation"], dict):
            success = results["notebook_creation"].get("success", False)
            print(f"Success: {success}")
            if success:
                print(
                    f"Original count: {results['notebook_creation'].get('original_count', 'Unknown')}"
                )
                print(
                    f"New count: {results['notebook_creation'].get('new_count', 'Unknown')}"
                )
        else:
            print(f"Error: {results['notebook_creation']}")

    # Print worksheet creation test results
    if "worksheet_creation" in results:
        print("\nWORKSHEET CREATION TEST:")
        print("-" * 50)
        if isinstance(results["worksheet_creation"], dict):
            success = results["worksheet_creation"].get("success", False)
            method = results["worksheet_creation"].get("method", "Unknown")
            print(f"Success: {success}")
            print(f"Method used: {method}")
        else:
            print(f"Error: {results['worksheet_creation']}")

    print("\n===== End of Inspection Results =====")


# def main():
#     """
#     Command-line entry point for COM inspection
#     """
#     import argparse

#     parser = argparse.ArgumentParser(
#         description="Inspect SigmaPlot COM interface"
#     )
#     parser.add_argument(
#         "--depth", type=int, default=2, help="Depth of exploration"
#     )
#     parser.add_argument(
#         "--no-collections",
#         action="store_true",
#         help="Skip collection exploration",
#     )
#     parser.add_argument("--json", action="store_true", help="Output as JSON")
#     parser.add_argument(
#         "--log-level",
#         default="INFO",
#         choices=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
#         help="Set logging level",
#     )

#     args = parser.parse_args()

#     # Configure logging
#     logging.basicConfig(
#         level=getattr(logging, args.log_level),
#         format="%(levelname)s: %(message)s",
#     )

#     # Run inspection
#     results = inspect_sigmaplot(
#         depth=args.depth, explore_collections=not args.no_collections
#     )

#     # Output
#     if args.json:
#         print(json.dumps(to_json_serializable(results), indent=2))
#     else:
#         print_inspection_results(results)

def main():
    """Main function to run the inspection script."""
    import argparse
    import json
    import os
    import time
    from datetime import datetime

    # Parse command line arguments
    parser = argparse.ArgumentParser(description='Inspect SigmaPlot COM objects')
    parser.add_argument('--depth', type=int, default=2,
                       help='Maximum depth for object exploration')
    parser.add_argument('--no-collections', action='store_true',
                       help='Skip exploring collections (faster)')
    parser.add_argument('--output', '-o', type=str,
                       help='Output file for full JSON results')
    parser.add_argument('--verbose', '-v', action='store_true',
                       help='Enable verbose logging')

    args = parser.parse_args()

    if args.verbose:
        logger.setLevel(logging.DEBUG)

    # Run inspection
    results = inspect_sigmaplot(
        depth=args.depth, explore_collections=not args.no_collections
    )

    # Save results to JSON if requested
    if args.output:
        try:
            # Convert results to serializable format
            serializable_results = convert_to_serializable(results)

            with open(args.output, 'w') as f:
                json.dump(serializable_results, f, indent=2)
            logger.info(f"Results saved to {args.output}")
        except Exception as e:
            logger.error(f"Error saving results: {e}")

    # Print a summary of the results
    print("===== SigmaPlot COM Inspection Results =====")

    print("APPLICATION STRUCTURE:")
    print("--------------------------------------------------")
    print(f"Type: {results['application']['type']}")

    if 'app_properties' in results:
        print("\nAPPLICATION PROPERTIES:")
        print("--------------------------------------------------")
        for prop, value in results['app_properties'].items():
            print(f"{prop}: {value}")

    if 'notebooks' in results:
        print("\nNOTEBOOKS:")
        print("--------------------------------------------------")
        if isinstance(results['notebooks'], dict):
            print(f"Count: {results['notebooks']['count']}")
            print("Notebook items:")
            for key, notebook in results['notebooks']['items'].items():
                name = notebook.get('name', key) if isinstance(notebook, dict) else key
                print(f"- {name}")

    if 'active_document' in results:
        print("\nACTIVE DOCUMENT:")
        print("--------------------------------------------------")
        if isinstance(results['active_document'], dict):
            name = results['active_document'].get('name', 'Unknown')
            print(f"Name: {name}")
            print(f"Type: {results['active_document'].get('type', 'Unknown')}")

            # Print collections if available
            if 'collections' in results['active_document']:
                print("Collections:")
                for coll_name, coll_info in results['active_document']['collections'].items():
                    print(f"- {coll_name}: {coll_info.get('count', 0)} items")

    if 'notebook_operations' in results:
        print("\nNOTEBOOK OPERATIONS:")
        print("--------------------------------------------------")
        if isinstance(results['notebook_operations'], dict):
            for op, info in results['notebook_operations'].items():
                if isinstance(info, dict) and 'success' in info:
                    print(f"{op}: {'Success' if info['success'] else 'Failed'}")
                else:
                    print(f"{op}: {info}")

    print("===== End of Inspection Results =====")

def convert_to_serializable(obj):
    """Convert a nested dict with potentially non-serializable values to JSON-serializable format."""
    if isinstance(obj, dict):
        return {k: convert_to_serializable(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [convert_to_serializable(item) for item in obj]
    elif isinstance(obj, (str, int, float, bool, type(None))):
        return obj
    else:
        return str(obj)

if __name__ == "__main__":
    main()

# EOF