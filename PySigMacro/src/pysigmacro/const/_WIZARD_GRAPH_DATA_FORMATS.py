#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-22 11:12:31 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/const/_WIZARD_GRAPH_DATA_FORMATS.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/const/_WIZARD_GRAPH_DATA_FORMATS.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

_GW_DATAFORMATS_SIMPLE_PLOTS = [
    "XY Pair",
    "Single X",
    "Single Y",
]
_GW_DATAFORMATS_MULTIPLE_PLTOS = [
    "XY Pairs",
    "X Many Y",
    "Y Many X",
    "Many X",
    "Many Y",
]
_GW_DATAFORMATS_POLAR_PLOTS = [
    "ThetaR",
    "XY Pairs",
    "Theta Many R",
    "R Many Theta",
    "Many R",
    "Many Theta",
]
_GW_DATAFORMATS_THREE_DIM_AND_CONTOUR_PLOTS = [
    "XYZ Triplet (not available for bar charts)",
    "Many Z",
    "XY Many Z",
]
_GW_DATAFORMATS_TERNARY_PLOTS = [
    "Ternary Triplets",
    "Ternary XY Pairs",
    "Ternary YZ Pairs",
    "Ternary XZ Pairs",
]
_GW_DATAFORMATS_PIE_PLOTS = [
    "Single Column",
]


GW_DATAFORMATS_DICT = {
    "SIMPLE_PLOTS": _GW_DATAFORMATS_SIMPLE_PLOTS,
    "MULTIPLE_PLOTS": _GW_DATAFORMATS_MULTIPLE_PLTOS,
    "POLAR_PLOTS": _GW_DATAFORMATS_POLAR_PLOTS,
    "THREE_DIM_AND_CONTOUR_PLOTS": _GW_DATAFORMATS_THREE_DIM_AND_CONTOUR_PLOTS,
    "TERNARY_PLOTS": _GW_DATAFORMATS_TERNARY_PLOTS,
    "PIE_PLOTS": _GW_DATAFORMATS_PIE_PLOTS
}

# EOF