#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-14 18:28:09 (ywatanabe)"
# File: /home/ywatanabe/proj/SigMacro/PySigMacro/dev.py

import os

__THIS_FILE__ = (
    "/home/ywatanabe/proj/SigMacro/PySigMacro/dev.py"
)
__THIS_DIR__ = os.path.dirname(__THIS_FILE__)


"""
Development and testing script for the PySigMacro package.

This script demonstrates the basic capabilities of the PySigMacro package
for SigmaPlot automation using the core functionality directly.
"""

import time
import tempfile
from pysigmacro import SigmaPlotAutomator
from pysigmacro.data.csv_handler import create_sample_csv
from pysigmacro.macro.executor import MacroExecutor
from pysigmacro.macro.templates import get_plot_macro
from pysigmacro.utils.explorer import explore_notebook_structures
from pysigmacro.utils.com_helpers import explore_com_object

PATH = os.path.join("C:\\Temp", f"SigmaPlot_Basic_{time.strftime('%Y%m%d_%H%M%S')}.JNB")
sp = SigmaPlotAutomator(visible=True, close_others=True, file_path=PATH)


# Create some worksheet and graph items
 # This creates new section; why?
ws_item = sp.add(item_type="worksheet", name="TestWorksheet")
graph_item = sp.add("graph", "TestGraph")
 # This creates new section; why?
ws_item = sp.add("worksheet", "TestWorksheet")
graph_item = sp.add("graph", "TestGraph2")

# EOF