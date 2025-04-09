#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-23 10:33:11 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_get_active_document.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_get_active_document.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

def get_active_document(app):
    active_document = app.ActiveDocument_obj
    return active_document

# EOF