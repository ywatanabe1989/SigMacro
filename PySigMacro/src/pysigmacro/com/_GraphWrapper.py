#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-24 19:33:14 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_GraphWrapper.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_GraphWrapper.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from ._BaseCOMWrapper import BaseCOMWrapper
from ._base import get_wrapper

class GraphWrapper(BaseCOMWrapper):
    """Specialized wrapper for Graph"""

    __classname__ = "GraphWrapper"

# EOF