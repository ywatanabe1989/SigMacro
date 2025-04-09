#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-24 18:16:43 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_GraphPageWrapper.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_GraphPageWrapper.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from ._BaseCOMWrapper import BaseCOMWrapper

class GraphPageWrapper(BaseCOMWrapper):
    """Specialized wrapper for GraphPage object"""

    __classname__ = "GraphPageWrapper"

    def configure(self, **kwargs):
        """Configure graph page properties"""
        for key, value in kwargs.items():
            try:
                setattr(self._com_object, key, value)
            except Exception as e:
                print(f"Error setting {key}={value}: {e}")
        return self

# EOF