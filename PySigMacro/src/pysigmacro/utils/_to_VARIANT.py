#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-24 20:43:10 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_to_VARIANT.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_to_VARIANT.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from win32com.client import VARIANT
import pythoncom

def to_VARIANT(value, variant_type=None):
    """Create a properly typed COM variant"""
    if variant_type is None:
        # Auto-determine type
        if isinstance(value, str):
            return VARIANT(pythoncom.VT_BSTR, value)
        elif isinstance(value, int):
            return VARIANT(pythoncom.VT_I4, value)
        elif isinstance(value, float):
            return VARIANT(pythoncom.VT_R8, value)
        elif isinstance(value, list):
            # Handle list by creating a safe array
            return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, value)
        elif hasattr(value, "_oleobj_"):
            # COM object - pass through directly
            return value
        elif isinstance(value, dict):
            # Dictionary - convert to COM dictionary object
            dict_obj = pythoncom.CreateObject("Scripting.Dictionary")
            for k, v in value.items():
                dict_obj.Add(k, to_VARIANT(v))
            return dict_obj
        else:
            return value
    else:
        return VARIANT(variant_type, value)

# EOF