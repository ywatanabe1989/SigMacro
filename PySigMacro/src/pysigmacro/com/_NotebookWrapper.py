#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-24 18:42:47 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_NotebookWrapper.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_NotebookWrapper.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from ..const import *
import re
from ._BaseCOMWrapper import BaseCOMWrapper
from ._base import get_wrapper

class NotebookWrapper(BaseCOMWrapper):
    """Specialized wrapper for Notebook Item"""

    __classname__ = "NotebookWrapper"

# EOF