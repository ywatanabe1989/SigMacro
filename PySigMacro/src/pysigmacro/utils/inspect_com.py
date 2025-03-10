#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 09:51:03 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/utils/inspect_com.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/utils/inspect_com.py"
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 10:30:12 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/utils/inspect_com.py

"""
General COM object inspection utilities.
Provides tools to inspect and interact with any COM objects in a more Pythonic way.
"""

import sys
import win32com.client
import pythoncom
import inspect
from typing import Any, Dict, List, Tuple, Optional, Union, Set
import logging
from pprint import pformat
import json

# Set up logging
logger = logging.getLogger(__name__)


def get_com_properties(obj: Any) -> List[str]:
    """
    Get a list of properties available in a COM object.

    Args:
        obj: A COM object

    Returns:
        List of property names
    """
    properties = []

    try:
        # Try to get type info
        if hasattr(obj, "_oleobj_"):
            typeinfo = obj._oleobj_.GetTypeInfo()
            attr = typeinfo.GetTypeAttr()

            # Loop through all functions
            for i in range(attr.cFuncs):
                func_desc = typeinfo.GetFuncDesc(i)
                func_name = typeinfo.GetNames(func_desc[0])[0]
                invkind = func_desc[4]

                # Property getters
                if invkind == pythoncom.INVOKE_PROPERTYGET:
                    properties.append(func_name)
    except Exception as e:
        logger.debug(f"Error getting COM properties: {e}")

    return properties


def get_com_methods(obj: Any) -> List[str]:
    """
    Get a list of methods available in a COM object.

    Args:
        obj: A COM object

    Returns:
        List of method names
    """
    methods = []

    try:
        # Try to get type info
        if hasattr(obj, "_oleobj_"):
            typeinfo = obj._oleobj_.GetTypeInfo()
            attr = typeinfo.GetTypeAttr()

            # Loop through all functions
            for i in range(attr.cFuncs):
                func_desc = typeinfo.GetFuncDesc(i)
                func_name = typeinfo.GetNames(func_desc[0])[0]
                invkind = func_desc[4]

                # Methods
                if invkind == pythoncom.INVOKE_FUNC:
                    methods.append(func_name)
    except Exception as e:
        logger.debug(f"Error getting COM methods: {e}")

    return methods


def explore_com_object(
    obj: Any,
    name: str = "Object",
    max_depth: int = 2,
    current_depth: int = 0,
    visited: Optional[Set] = None,
) -> Dict:
    """
    Recursively explore a COM object and return its structure.

    Args:
        obj: A COM object
        name: Name to give this object in the output
        max_depth: Maximum depth to explore
        current_depth: Current depth (used in recursion)
        visited: Set of already visited objects to prevent loops

    Returns:
        Dictionary representing the object structure
    """
    if visited is None:
        visited = set()

    # Check if we've already visited this object or reached max depth
    if current_depth >= max_depth:
        return {
            "name": name,
            "type": str(type(obj)),
            "max_depth_reached": True,
        }

    # Get object ID for tracking visited objects
    # Use memory address as a proxy for object identity since CDispatch isn't hashable
    obj_id = id(obj)
    if obj_id in visited:
        return {"name": name, "type": str(type(obj)), "already_visited": True}

    # Add to visited set
    visited.add(obj_id)

    result = {
        "name": name,
        "type": str(type(obj)),
        "properties": {},
        "methods": [],
        "collections": {},
    }

    # Get methods and properties
    methods = get_com_methods(obj)
    properties = get_com_properties(obj)

    result["methods"] = methods

    # Try to access properties
    for prop in properties:
        try:
            prop_value = getattr(obj, prop)

            # Check if this is a collection
            if hasattr(prop_value, "Count") and hasattr(prop_value, "Item"):
                count = prop_value.Count
                result["collections"][prop] = {"count": count, "items": {}}

                # Get first few items if any
                if count > 0 and current_depth < max_depth - 1:
                    for i in range(1, min(count + 1, 4)):
                        try:
                            item = prop_value.Item(i)
                            item_name = getattr(item, "Name", f"Item{i}")
                            result["collections"][prop]["items"][f"{i}"] = (
                                explore_com_object(
                                    item,
                                    item_name,
                                    max_depth,
                                    current_depth + 1,
                                    visited,
                                )
                            )
                        except Exception as e:
                            result["collections"][prop]["items"][
                                f"{i}"
                            ] = f"Error: {str(e)}"
            else:
                # For non-collection properties, just show type
                result["properties"][prop] = str(type(prop_value))
        except Exception as e:
            result["properties"][prop] = f"Error: {str(e)}"

    return result


def call_com_method(obj: Any, method_name: str, *args, **kwargs) -> Any:
    """
    Safely call a COM method, with better error handling.

    Args:
        obj: COM object
        method_name: Name of the method to call
        *args: Arguments to pass to the method
        **kwargs: Keyword arguments for the method

    Returns:
        Result of the method call
    """
    try:
        # First try direct attribute access
        if hasattr(obj, method_name):
            method = getattr(obj, method_name)
            return method(*args, **kwargs)

        # If that fails, try COM dispatch
        if hasattr(obj, "_oleobj_"):
            # Get method ID
            method_id = obj._oleobj_.GetIDsOfNames(method_name)[0]

            # Prepare arguments
            dispatch_args = []
            for arg in args:
                dispatch_args.append(arg)

            # Call method
            return obj._oleobj_.Invoke(
                method_id, 0, pythoncom.DISPATCH_METHOD, dispatch_args
            )

        raise AttributeError(f"Method {method_name} not found on object")
    except Exception as e:
        logger.error(f"Error calling {method_name}: {e}")
        raise


def get_com_property(obj: Any, prop_name: str) -> Any:
    """
    Safely get a COM property with better error handling.

    Args:
        obj: COM object
        prop_name: Name of the property

    Returns:
        Property value
    """
    try:
        # First try direct attribute access
        if hasattr(obj, prop_name):
            return getattr(obj, prop_name)

        # If that fails, try COM dispatch
        if hasattr(obj, "_oleobj_"):
            # Get property ID
            prop_id = obj._oleobj_.GetIDsOfNames(prop_name)[0]

            # Get property
            return obj._oleobj_.Invoke(
                prop_id, 0, pythoncom.DISPATCH_PROPERTYGET, []
            )

        raise AttributeError(f"Property {prop_name} not found on object")
    except Exception as e:
        logger.error(f"Error getting property {prop_name}: {e}")
        raise


def set_com_property(obj: Any, prop_name: str, value: Any) -> None:
    """
    Safely set a COM property with better error handling.

    Args:
        obj: COM object
        prop_name: Name of the property
        value: Value to set

    Returns:
        None
    """
    try:
        # First try direct attribute access
        if hasattr(obj, prop_name):
            setattr(obj, prop_name, value)
            return

        # If that fails, try COM dispatch
        if hasattr(obj, "_oleobj_"):
            # Get property ID
            prop_id = obj._oleobj_.GetIDsOfNames(prop_name)[0]

            # Set property
            obj._oleobj_.Invoke(
                prop_id, 0, pythoncom.DISPATCH_PROPERTYPUT, [value]
            )
            return

        raise AttributeError(f"Property {prop_name} not found on object")
    except Exception as e:
        logger.error(f"Error setting property {prop_name}: {e}")
        raise


def print_com_object_structure(
    obj: Any, name: str = "Object", max_depth: int = 2
) -> None:
    """
    Print the structure of a COM object in a readable format.

    Args:
        obj: COM object
        name: Name to give this object in the output
        max_depth: Maximum recursion depth

    Returns:
        None
    """
    structure = explore_com_object(obj, name, max_depth)
    print(pformat(structure, indent=2))


def create_com_wrapper(obj: Any) -> Any:
    """
    Create a more Pythonic wrapper around a COM object.
    This adds better error messages and makes properties/methods more discoverable.

    Args:
        obj: COM object

    Returns:
        Wrapped object
    """
    # Get properties and methods
    properties = get_com_properties(obj)
    methods = get_com_methods(obj)

    class ComWrapper:
        """Wrapper for COM object with better Python integration."""

        def __init__(self, com_obj):
            self._com_obj = com_obj
            self._properties = properties
            self._methods = methods

        def __getattr__(self, name):
            # Check if it's a property
            if name in self._properties:
                value = get_com_property(self._com_obj, name)

                # If it's another COM object, wrap it too
                if hasattr(value, "_oleobj_"):
                    return create_com_wrapper(value)
                return value

            # Check if it's a method
            if name in self._methods:

                def method_wrapper(*args, **kwargs):
                    result = call_com_method(
                        self._com_obj, name, *args, **kwargs
                    )

                    # If result is a COM object, wrap it
                    if hasattr(result, "_oleobj_"):
                        return create_com_wrapper(result)
                    return result

                return method_wrapper

            # Try direct attribute access as a fallback
            try:
                value = getattr(self._com_obj, name)

                # If it's a callable, wrap it
                if callable(value):

                    def wrapped_callable(*args, **kwargs):
                        result = value(*args, **kwargs)
                        if hasattr(result, "_oleobj_"):
                            return create_com_wrapper(result)
                        return result

                    return wrapped_callable

                # If it's another COM object, wrap it
                if hasattr(value, "_oleobj_"):
                    return create_com_wrapper(value)
                return value
            except AttributeError:
                raise AttributeError(
                    f"'{type(self._com_obj).__name__}' has no attribute '{name}'"
                )

        def __setattr__(self, name, value):
            # Handle internal attributes
            if name.startswith("_"):
                super().__setattr__(name, value)
                return

            # Handle COM properties
            if name in self._properties:
                set_com_property(self._com_obj, name, value)
                return

            # Fallback to direct attribute setting
            try:
                setattr(self._com_obj, name, value)
            except AttributeError:
                super().__setattr__(name, value)

        def __dir__(self):
            """Make properties and methods discoverable with dir()"""
            return list(
                set(super().__dir__() + self._properties + self._methods)
            )

    return ComWrapper(obj)


def inspect_com_object(
    obj: Any,
    name: str = "ComObject",
    depth: int = 2,
    explore_collections: bool = True,
) -> Dict:
    """
    General-purpose COM object inspection function.

    Args:
        obj: COM object to inspect
        name: Name to use for the object
        depth: Depth of exploration
        explore_collections: Whether to explore collections in the object

    Returns:
        Dictionary with inspection results
    """
    results = {}

    # Basic object structure
    logger.info(f"Exploring {name} object...")
    results["object"] = explore_com_object(obj, name, max_depth=depth)

    # If we're exploring collections, look for them in the primary object
    if explore_collections:
        collections = results["object"].get("collections", {})
        if collections:
            results["collections"] = {}

            # Explore each collection
            for coll_name, coll_info in collections.items():
                logger.info(f"Exploring collection: {coll_name}")
                try:
                    collection = getattr(obj, coll_name)
                    count = coll_info.get("count", 0)

                    results["collections"][coll_name] = {
                        "count": count,
                        "items": {},
                    }

                    # Explore items in collection
                    if count > 0:
                        for i in range(1, min(count + 1, 11)):
                            try:
                                item = collection.Item(i)
                                item_name = getattr(
                                    item, "Name", f"{coll_name}Item{i}"
                                )
                                results["collections"][coll_name]["items"][
                                    item_name
                                ] = explore_com_object(
                                    item, item_name, max_depth=depth - 1
                                )
                            except Exception as e:
                                logger.error(
                                    f"Error exploring item {i} in {coll_name}: {e}"
                                )
                                results["collections"][coll_name]["items"][
                                    f"Item{i}"
                                ] = str(e)
                except Exception as e:
                    logger.error(
                        f"Error exploring collection {coll_name}: {e}"
                    )
                    results["collections"][coll_name] = str(e)

    return results


def print_com_inspection_results(results: Dict) -> None:
    """
    Print COM inspection results in a readable format.

    Args:
        results: Results from inspect_com_object()

    Returns:
        None
    """
    print("\n===== COM Object Inspection Results =====\n")

    # Print object structure
    if "object" in results:
        print("OBJECT STRUCTURE:")
        print("-" * 50)
        print(f"Name: {results['object'].get('name', 'Unknown')}")
        print(f"Type: {results['object'].get('type', 'Unknown')}")

        # Print methods
        methods = results["object"].get("methods", [])
        if methods:
            print(f"\nMethods ({len(methods)}):")
            for method in sorted(methods):
                print(f"  - {method}")

        # Print properties
        properties = results["object"].get("properties", {})
        if properties:
            print(f"\nProperties ({len(properties)}):")
            for prop, prop_type in sorted(properties.items()):
                print(f"  - {prop}: {prop_type}")

        # Print collections
        collections = results["object"].get("collections", {})
        if collections:
            print(f"\nCollections ({len(collections)}):")
            for coll_name, coll_data in sorted(collections.items()):
                print(f"  - {coll_name}: {coll_data.get('count', 0)} items")

    # Print collections information if explored
    if "collections" in results:
        print("\nCOLLECTIONS DETAILS:")
        print("-" * 50)

        for coll_name, coll_data in sorted(results["collections"].items()):
            if isinstance(coll_data, dict):
                count = coll_data.get("count", 0)
                print(f"\n{coll_name} ({count} items):")

                items = coll_data.get("items", {})
                if items:
                    for item_name, item_data in items.items():
                        if isinstance(item_data, dict):
                            print(
                                f"  - {item_name} ({item_data.get('type', 'Unknown')})"
                            )

                            # Show methods and properties count
                            methods = item_data.get("methods", [])
                            properties = item_data.get("properties", {})
                            print(
                                f"    * {len(methods)} methods, {len(properties)} properties"
                            )

                            # Show collections in this item
                            item_collections = item_data.get("collections", {})
                            if item_collections:
                                print(f"    * Collections:")
                                for item_coll_name, item_coll_data in sorted(
                                    item_collections.items()
                                ):
                                    print(
                                        f"      - {item_coll_name}: {item_coll_data.get('count', 0)} items"
                                    )
                        else:
                            print(f"  - {item_name}: {item_data}")
            else:
                print(f"\n{coll_name}: {coll_data}")

    print("\n===== End of Inspection Results =====")


def to_json_serializable(obj: Any) -> Any:
    """
    Convert inspection results to JSON-serializable format.

    Args:
        obj: Inspection results

    Returns:
        JSON-serializable object
    """
    if isinstance(obj, dict):
        return {k: to_json_serializable(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [to_json_serializable(i) for i in obj]
    elif isinstance(obj, (int, float, str, bool, type(None))):
        return obj
    else:
        return str(obj)


def main():
    """
    Command-line entry point for COM inspection
    """
    import argparse
    import win32com.client

    parser = argparse.ArgumentParser(description="Inspect any COM object")
    parser.add_argument(
        "--progid",
        type=str,
        required=True,
        help='ProgID of the COM object to inspect (e.g. "Excel.Application")',
    )
    parser.add_argument(
        "--depth", type=int, default=2, help="Depth of exploration"
    )
    parser.add_argument(
        "--no-collections",
        action="store_true",
        help="Skip collection exploration",
    )
    parser.add_argument("--json", action="store_true", help="Output as JSON")
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
        help="Set logging level",
    )

    args = parser.parse_args()

    # Configure logging
    logging.basicConfig(
        level=getattr(logging, args.log_level),
        format="%(levelname)s: %(message)s",
    )

    # Create COM object
    try:
        logger.info(f"Creating COM object: {args.progid}")
        obj = win32com.client.Dispatch(args.progid)
        logger.info("COM object created successfully")

        # Run inspection
        results = inspect_com_object(
            obj,
            name=args.progid,
            depth=args.depth,
            explore_collections=not args.no_collections,
        )

        # Output
        if args.json:
            print(json.dumps(to_json_serializable(results), indent=2))
        else:
            print_com_inspection_results(results)

    except Exception as e:
        logger.error(f"Error: {e}")
        import traceback

        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()

# EOF