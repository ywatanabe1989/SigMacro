#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-26 18:31:01 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_PlotsWrapper.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_PlotsWrapper.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from ._BaseCOMWrapper import BaseCOMWrapper
from ._base import get_wrapper

class PlotsWrapper(BaseCOMWrapper):
    """Specialized wrapper for Plots collection"""

    __classname__ = "PlotsWrapper"

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

    def configure_all(self, **kwargs):
        """Configure all plots with the same properties"""
        for i in range(self._com_object.Count):
            plot = self[i]
            for key, value in kwargs.items():
                try:
                    # Use SetAttribute for plot properties
                    plot.SetAttribute(key, value)
                except Exception as e:
                    print(f"Error setting {key}={value} on plot {i}: {e}")
        return self

    def clean(self):
        """Remove default or auto-generated plots"""
        print(f"Cleaning plots in {self._access_path}")
        # Need to iterate in reverse because removing plots affects the collection indices
        for ii in range(self._com_object.Count - 1, -1, -1):
            try:
                # You might want to customize the criteria for plots to remove
                # For example, removing plots with default names or specific attributes
                plot = self._com_object[ii]
                # Example criteria: check if it has a default name
                # (modify based on your needs)
                try:
                    name = plot.Name
                    if name.startswith("Plot") and name[4:].isdigit():
                        print(f"  Removing plot {ii}: {name}")
                        plot.Delete()
                except:
                    # If we can't get the name, we might want to be cautious
                    pass
            except Exception as e:
                print(f"  Error cleaning plot {ii}: {e}")
        return self

# EOF