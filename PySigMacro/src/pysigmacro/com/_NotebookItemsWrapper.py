#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-31 15:06:29 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_NotebookItemsWrapper.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_NotebookItemsWrapper.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from ..const import *
import re
from ._BaseCOMWrapper import BaseCOMWrapper
from ._base import get_wrapper

class NotebookItemsWrapper(BaseCOMWrapper):
    """Specialized wrapper for NotebookItems collection"""

    __classname__ = "NotebookItemsWrapper"

    def clean(self):
        """Clean notebook items with default naming pattern (e.g., Section 1, Section 2, ...)"""
        patterns = [
            re.compile(r"^Section\s+\d+$"), # Matches "Section 1", "Section 2", etc.
            re.compile(r"^Graph\s+Page\s+\d+$"), # Matches "Graph Page 1", "Graph Page 2", etc.
            re.compile(r"^Data\s+\d+$"), # Matches "Data 1", "Data 2", etc.
            # Add other default patterns if needed
        ]
        for pattern in patterns:
            # Need to iterate in reverse because closing affects the collection indices
            for ii in range(self._com_object.Count - 1, -1, -1):
                # print(ii, self._com_object[ii].Name)
                if pattern.match(self._com_object[ii].Name):
                    try:
                        self._com_object[ii].Close(False)
                    except IndexError as e:
                        print(f"Could not close notebook item {self._com_object[ii].Name}: {e}")

    def add_worksheet(self, name=None):
        """Add a new worksheet to the notebook"""
        worksheet_item = get_wrapper(self._com_object.Add(CT_WORKSHEET), self.access_path, self.path)
        if name:
            worksheet_item.Name = name
        return worksheet_item

    def add_graph(self, name=None):
        """Add a new graph to the notebook"""
        graph_item = get_wrapper(self._com_object.Add(CT_GRAPHICPAGE), self.access_path, self.path)
        if name:
            graph_item.Name = name
        return graph_item

    def add_report(self, name=None):
        """Add a new report to the notebook"""
        report_item = get_wrapper(self._com_object.Add(CT_REPORT), self.access_path, self.path)
        if name:
            report_item.Name = name
        return report_item

    def add_section(self, name=None):
        """Add a new section to the notebook"""
        section_item = get_wrapper(self._com_object.Add(CT_FOLDER), self.access_path, self.path)
        if name:
            section_item.Name = name
        return section_item

    def Item(self, key):
        if key == -1 and hasattr(self._com_object, "Count"):
            key = self._com_object.Count - 1
        result = self._com_object(key)
        access_path = f"{self._access_path}({key})" if self._access_path else f"({key})"
        return get_wrapper(result, access_path, self.path)

    def __call__(self, key):
        if key == -1 and hasattr(self._com_object, "Count"):
            key = self._com_object.Count - 1
        result = self._com_object(key)
        access_path = f"{self._access_path}({key})" if self._access_path else f"({key})"
        return get_wrapper(result, access_path, self.path)

    def __getitem__(self, key):
        if key == -1 and hasattr(self._com_object, "Count"):
            key = self._com_object.Count - 1
        result = self._com_object(key)
        access_path = f"{self._access_path}[{key}]" if self._access_path else f"[{key}]"
        return get_wrapper(result, access_path, self.path)

# EOF