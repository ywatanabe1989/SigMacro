#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-26 18:29:39 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_base.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_base.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

wrap_function = None  # Will be populated later

def register_wrap_function(func):
    global wrap_function
    wrap_function = func

def get_wrapper(com_obj, access_path="", path=""):
    """
    Get the appropriate wrapper for a COM object
    """
    try:
        # Ensure access_path is a string
        access_path = str(access_path) if access_path is not None else ""

        from ._wrap import wrap
        return wrap(com_obj, access_path, path)
    except Exception as e:
        print(f"Error in get_wrapper: {e}")
        # Ensure we don't return None or an unwrapped object
        from ._BaseCOMWrapper import BaseCOMWrapper
        wrapper = BaseCOMWrapper(com_obj, access_path)
        if path:
            wrapper._path = path
        return wrapper

# EOF