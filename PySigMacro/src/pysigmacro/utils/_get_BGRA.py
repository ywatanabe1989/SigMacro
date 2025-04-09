#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-01 15:58:07 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_get_BGRA.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_get_BGRA.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from ..const._COLORS import BGRA

def get_BGRA(color_str, alpha=1.0):
    bgra = BGRA[color_str]
    bgra[3] = alpha
    return bgra

# EOF