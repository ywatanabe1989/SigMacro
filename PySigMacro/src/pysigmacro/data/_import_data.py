#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-09 19:37:24 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/data/_import_data.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/data/_import_data.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import time
import pandas as pd

def import_data(worksheet_item, df=None, csv=None, left=0, top=0):
    """
    Import data into a SigmaPlot worksheet from a pandas DataFrame or CSV file.

    This function takes either a pandas DataFrame or a path to a CSV file and
    imports the data into a SigmaPlot worksheet. Column headers are set based
    on the DataFrame's column names.

    Args:
        worksheet_item (COM object): A SigmaPlot worksheet COM object
        df (pandas.DataFrame, optional): DataFrame containing data to import.
            Defaults to None.
        csv (str, optional): Path to a CSV file to import. Used only if df is None.
            Defaults to None.
        left (int, optional): Left column index where data import starts.
            Defaults to 0.
        top (int, optional): Top row index where data import starts.
            Defaults to 0.

    Returns:
        COM object: The worksheet's DataTable COM object

    Note:
        Either df or csv must be provided. If both are provided, df takes precedence.
    """
    if (df is None) and (csv is None):
        raise ValueError("Either df or csv must be provided")

    # df
    if (df is None) and (csv is not None):
        df = pd.read_csv(csv)

    # datatable object
    datatable_obj = worksheet_item.DataTable_obj

    for ii, column_name in enumerate(df.columns):
        col = df[column_name]
        # Remove NaN for data
        col = col.replace("NONE_STR", "None")
        col = col[~col.isna()]
        if column_name == "symbol":
            col = col.astype(str)
        try:
            datatable_obj.PutData(col.tolist(), left+ii, top)
        except Exception as e:
            print(e)
            __import__("ipdb").set_trace()

    # Header
    for ii, header in enumerate(df.columns):
        datatable_obj.ColumnTitle(ii, str(header)) # This does not work

    return datatable_obj

# EOF