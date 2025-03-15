#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-15 01:43:20 (ywatanabe)"
# File: /home/ywatanabe/proj/SigMacro/PySigMacro/src/pysigmacro/utils/_sigmaplot_inspect.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/proj/SigMacro/PySigMacro/src/pysigmacro/utils/_sigmaplot_inspect.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import pandas as pd
from ._sigmaplot_objects import SIGMAPLOT_OBJECTS
from ._sigmaplot_properties import SIGMAPLOT_PROPERTIES
from ._sigmaplot_methods import SIGMAPLOT_METHODS




def inspect(com_object):
    valid_props = get_valid_properties(com_object)
    valid_methods = get_valid_methods(com_object)
    prop_dict = {}
    object_dict = {}
    for prop in valid_props:
        try:
            value = getattr(com_object, prop)
            if hasattr(value, "_oleobj_"):
                object_dict[prop] = value
            else:
                value_str = str(value)
                prop_dict[prop] = value_str
        except Exception as e:
            value_str = f"Error: {e}"
    method_dict = {"Methods": valid_methods}
    rows = []
    for key, value in prop_dict.items():
        rows.append({"Name": key, "Type": "Property", "Value": value})
    for key, value in object_dict.items():
        rows.append({"Name": key, "Type": "Object", "Value": value})
    for m in valid_methods:
        rows.append({"Name": m, "Type": "Method", "Value": ""})
    summary_df = pd.DataFrame(rows)
    return summary_df
def _get_valid_xxx(com_object, available_list):
    valid_list = []
    for possible in available_list:
        if possible in ("Save", "SaveAs"):
            valid_list.append(possible)
            continue
        try:
            getattr(com_object, possible)
            valid_list.append(possible)
        except Exception as e:
            err_str = str(e)
            if "VT_EMPTY" in err_str or "VT_BSTR" in err_str or "VT_I2" in err_str or "VT_I4" in err_str:
                valid_list.append(possible)
    return valid_list
def get_valid_methods(com_object):
    return _get_valid_xxx(com_object, SIGMAPLOT_METHODS)
def get_valid_properties(com_object):
    return _get_valid_xxx(com_object, SIGMAPLOT_PROPERTIES)

# EOF