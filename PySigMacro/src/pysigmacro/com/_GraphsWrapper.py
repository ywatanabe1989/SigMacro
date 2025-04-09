#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-26 18:30:16 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_GraphsWrapper.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_GraphsWrapper.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from ._BaseCOMWrapper import BaseCOMWrapper
from ._base import get_wrapper

class GraphsWrapper(BaseCOMWrapper):
    """Specialized wrapper for Graphs collection"""

    __classname__ = "GraphsWrapper"

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
        """Remove default or auto-generated graphs"""
        print(f"Cleaning graphs in {self._access_path}")
        # Need to iterate in reverse because removing affects the collection indices
        for ii in range(self._com_object.Count - 1, -1, -1):
            try:
                # Criteria for graphs to remove
                graph = self._com_object[ii]
                name = getattr(graph, "Name", f"Graph{ii}")
                if name.startswith("Graph") and name[5:].isdigit():
                    print(f"  Removing graph {ii}: {name}")
                    graph.Delete()
            except Exception as e:
                print(f"  Error cleaning graph {ii}: {e}")
        return self

# EOF