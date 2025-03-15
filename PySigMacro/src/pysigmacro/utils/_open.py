#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-15 02:01:11 (ywatanabe)"
# File: /home/ywatanabe/proj/SigMacro/PySigMacro/src/pysigmacro/utils/_open.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/proj/SigMacro/PySigMacro/src/pysigmacro/utils/_open.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import win32com.client

def open():
    return win32com.client.Dispatch("SigmaPlot.Application")

# EOF