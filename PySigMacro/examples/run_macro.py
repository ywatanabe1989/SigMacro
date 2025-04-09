#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-30 10:37:24 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/examples/run_macro.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/examples/run_macro.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

# Environmental variables only set if not already defined
if "SIGMACRO_JNB_PATH" not in os.environ:
    os.environ["SIGMACRO_JNB_PATH"] = rf"C:\Users\{os.getlogin()}\Documents\SigMacro\SigMacro.JNB"
if "SIGMACRO_TEMPLATES_DIR" not in os.environ:
    os.environ["SIGMACRO_TEMPLATES_DIR"] = rf"C:\Users\{os.getlogin()}\Documents\SigMacro\SigMacro\Templates"
if "SIGMAPLOT_BIN_PATH_WIN" not in os.environ:
    os.environ["SIGMAPLOT_BIN_PATH_WIN"] = rf"C:\Program Files (x86)\SigmaPlot\SPW16\Spw.exe"

"""
Scratch for Development
"""

import numpy as np
import pandas as pd

import pysigmacro as ps

# PARAMS
PLOT_TYPE = "line"
CLOSE_OTHERS = True
PATH = ps.path.copy_template("line", rf"C:\Users\wyusu\Downloads")
spw = ps.con.open(PATH)
notebooks = spw.Notebooks_obj
# print(notebooks.list)
notebook = notebooks[notebooks.find_indices(f"{PLOT_TYPE}")[0]]

# From here, templates defines indices and names
notebookitems = notebook.NotebookItems_obj
graphitem_s = notebookitems[
    notebookitems.find_indices(f"{PLOT_TYPE}_graph_S")[0]
]
graphitem_s.rename_xy_labels("aaa", "bbb")
worksheetitem = notebookitems[notebookitems.find_indices(f"{PLOT_TYPE}_worksheet")[0]]

graphitem_m = notebookitems[
    notebookitems.find_indices(f"{PLOT_TYPE}_graph_M")[0]
]
graphitem_l = notebookitems[
    notebookitems.find_indices(f"{PLOT_TYPE}_graph_L")[0]
]

# Data
df = pd.DataFrame(
    columns=[ii for ii in range(30)], data=np.random.rand(100, 30)
)

# spw = ps.con.open(PATH)
# notebooks = spw.Notebooks_obj
# # print(notebooks.list)
# notebook = notebooks[notebooks.find_indices(f"{PLOT_TYPE}")[0]]

# # From here, templates defines indices and names
# notebookitems = notebookitems
# graphitem_s = notebookitems[
#     notebookitems.find_indices(f"{PLOT_TYPE}_graph_S")[0]
# ]
# graphitem_m = notebookitems[
#     notebookitems.find_indices(f"{PLOT_TYPE}_graph_M")[0]
# ]
# graphitem_l = notebookitems[
#     notebookitems.find_indices(f"{PLOT_TYPE}_graph_L")[0]
# ]

# ps.utils.run_macro(
#     graphitem_s, "RenameXYLabels_macro", xlabel="X Label 1", ylabel="Y Label 1"
# )
# ps.utils.run_macro(
#     graphitem_m, "RenameXYLabels_macro", xlabel="X Label", ylabel="Y Label"
# )
# ps.utils.run_macro(
#     graphitem_l, "RenameXYLabels_macro", xlabel="X Label 2", ylabel="Y Label 2"
# )

# EOF