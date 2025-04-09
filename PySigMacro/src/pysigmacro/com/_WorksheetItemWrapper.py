#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-26 19:49:55 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_WorksheetItemWrapper.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_WorksheetItemWrapper.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from ..const import *
from ._BaseCOMWrapper import BaseCOMWrapper

class WorksheetItemWrapper(BaseCOMWrapper):
    """Specialized wrapper for WorksheetItem object"""
    __classname__ = "WorksheetItemWrapper"

    def import_data(self, df=None, csv=None, left=0, top=0):
        """
        Import data into this worksheet from a pandas DataFrame or CSV file.

        This method takes either a pandas DataFrame or a path to a CSV file and
        imports the data into the current worksheet. Column headers are set based
        on the DataFrame's column names.

        Args:
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
        from ..data._import_data import import_data as ps_data_import_data
        self.datatable = ps_data_import_data(self, df=df, csv=csv, left=left, top=top)

# EOF