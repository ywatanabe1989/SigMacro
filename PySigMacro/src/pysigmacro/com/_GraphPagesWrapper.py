#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-26 18:30:01 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_GraphPagesWrapper.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_GraphPagesWrapper.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from ._BaseCOMWrapper import BaseCOMWrapper
from ._base import get_wrapper
import re

class GraphPagesWrapper(BaseCOMWrapper):
    """Specialized wrapper for GraphPages collection"""

    __classname__ = "GraphPagesWrapper"

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

    def clean(self):
        """Clean graph pages with default naming pattern"""
        print(f"Cleaning graph pages in {self._access_path}")
        pattern = re.compile(r"^Page\s*\d+$")
        # Need to iterate in reverse because closing affects the collection indices
        for ii in range(self._com_object.Count - 1, -1, -1):
            try:
                name = self._com_object[ii].Name
                if pattern.match(name):
                    print(f"  Removing graph page {ii}: {name}")
                    self._com_object[ii].Delete()
            except Exception as e:
                print(f"  Error cleaning graph page {ii}: {e}")
        return self

# EOF