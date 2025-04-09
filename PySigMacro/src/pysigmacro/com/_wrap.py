#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-26 19:41:31 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_wrap.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_wrap.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import re

from ._BaseCOMWrapper import BaseCOMWrapper
from ._NotebookWrapper import NotebookWrapper
from ._NotebooksWrapper import NotebooksWrapper
from ._NotebookItemsWrapper import NotebookItemsWrapper
from ._WorksheetItemWrapper import WorksheetItemWrapper
from ._MacroItemWrapper import MacroItemWrapper
from ._GraphItemWrapper import GraphItemWrapper
from ._GraphPagesWrapper import GraphPagesWrapper
from ._GraphPageWrapper import GraphPageWrapper
from ._GraphsWrapper import GraphsWrapper
from ._GraphWrapper import GraphWrapper
from ._PlotsWrapper import PlotsWrapper
from ._base import register_wrap_function


def wrap(com_object, access_path="", path=""):
    """
    Wrap a COM object in an appropriate wrapper class.
    """
    try:
        # Ensure access_path is a string
        access_path = str(access_path) if access_path is not None else ""

        # Create and configure the appropriate wrapper
        wrapper = _create_wrapper(com_object, access_path)

        # Set path if provided
        if path:
            wrapper._path = path

        return wrapper
    except Exception as e:
        # print(f"Error wrapping object: {e}")
        # Fall back to base wrapper
        wrapper = BaseCOMWrapper(com_object, access_path)
        if path:
            wrapper._path = path
        return wrapper


def _create_wrapper(com_object, access_path):
    """Create the appropriate wrapper based on access_path pattern"""
    access_path_last = access_path.split(".")[-1]
    # Notebooks
    if re.search(r"Notebooks$", access_path_last):
        return NotebooksWrapper(com_object, access_path)

    # Notebook
    elif re.search(r"Notebooks\[.*\]$", access_path_last):
        return NotebookWrapper(com_object, access_path)

    # NotebookItems
    elif re.search(r"NotebookItems$", access_path_last):
        return NotebookItemsWrapper(com_object, access_path)

    # Item
    elif re.search(r"NotebookItems\[.*\]$", access_path_last):

        # GraphItem
        if hasattr(com_object, "Name") and re.search(
            r".*graph.*", com_object.Name, re.IGNORECASE
        ):
            return GraphItemWrapper(com_object, access_path)
        # WorksheetItem
        elif hasattr(com_object, "Name") and re.search(
            r".*worksheet.*", com_object.Name, re.IGNORECASE
        ):
            return WorksheetItemWrapper(com_object, access_path)
        # MacroItem
        elif hasattr(com_object, "Name") and re.search(
            r".*macro.*", com_object.Name, re.IGNORECASE
        ):
            return MacroItemWrapper(com_object, access_path)
        else:
            return BaseCOMWrapper(com_object, access_path)

        # # GraphItem
        # if hasattr(com_object, "Name") and re.search(
        #     r".*graph.*", com_object.Name
        # ):
        #     return GraphItemWrapper(com_object, access_path)
        # # WorksheetItem
        # elif hasattr(com_object, "Name") and re.search(
        #     r".*worksheet.*", com_object.Name
        # ):
        #     return WorksheetItemWrapper(com_object, access_path)
        # # MacroItem
        # elif hasattr(com_object, "Name") and re.search(
        #     r".*macro.*", com_object.Name
        # ):
        #     return MacroItemWrapper(com_object, access_path)
        # else:
        #     return BaseCOMWrapper(com_object, access_path)

    # GraphPages
    elif re.search(r"GraphPages$", access_path_last):
        return GraphPagesWrapper(com_object, access_path)

    # GraphPage
    elif re.search(r"GraphPages\[.*\]$", access_path_last):
        return GraphPageWrapper(com_object, access_path)

    # Graphs
    elif re.search(r"Graphs$", access_path_last):
        return GraphsWrapper(com_object, access_path)

    # Graph
    elif re.search(r"Graphs\[.*\]$", access_path_last):
        return GraphWrapper(com_object, access_path)

    # Plots
    elif re.search(r"Plots$", access_path_last):
        return PlotsWrapper(com_object, access_path)

    # SigmaPlot root application
    elif access_path == "SigmaPlot":
        return BaseCOMWrapper(com_object, access_path)

    else:
        # Default to the base COM wrapper for unknown types
        return BaseCOMWrapper(com_object, access_path)

# Register the wrap function
register_wrap_function(wrap)

# EOF