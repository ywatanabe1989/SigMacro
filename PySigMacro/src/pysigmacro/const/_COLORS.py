#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-07 12:28:48 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/const/_COLORS.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/const/_COLORS.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

# Parameters
# --------------------------------------------------
BGR = {
    "black": [0, 0, 0],
    "white": [255, 255, 255],
    "gray": [128, 128, 128],
    "blue": [192, 128, 0],
    "green": [20, 180, 20],
    "red": [50, 70, 255],
    "yellow": [20, 160, 230],
    "purple": [255, 50, 200],
    "pink": [200, 150, 255],
    "light_blue": [200, 200, 20],
    "navy": [100, 0, 0],
    "orange": [50, 94, 228],
    "brown": [0, 0, 128],
}
BGRA = {k: [b, g, r, 1.0] for k, (b, g, r) in BGR.items()}
COLORS = list(BGR.keys())

BGRA_FAKE = ["NONE_STR", "NONE_STR", "NONE_STR", "NONE_STR"]

# EOF