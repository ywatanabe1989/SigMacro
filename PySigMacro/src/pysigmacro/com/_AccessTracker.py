#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-26 18:29:12 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_AccessTracker.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_AccessTracker.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

class AccessTracker:
    def __init__(self):
        self.access_history = []

    def add_access(self, parent_name, child_name, child_obj=None, path=""):
        full_path = f"{parent_name}.{child_name}" if parent_name else child_name
        self.access_history.append(full_path)
        return full_path

    def get_history(self):
        return self.access_history

    def get_current_path(self):
        return self.access_history[-1] if self.access_history else ""

# EOF