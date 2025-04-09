#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-30 10:37:02 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/examples/run_macros.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/examples/run_macros.py"
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

import pysigmacro as psm
import pandas as pd
import numpy as np


# Parameters
PLOT_TYPE = "dev"
CLOSE_OTHERS = True
TGT_DIR = rf"C:\Users\{os.getlogin()}\Desktop"
TGT_PATH = psm.path.copy_template(PLOT_TYPE, TGT_DIR)
TGT_FILENAME = os.path.basename(TGT_PATH)

# Instanciates SigmaPlot Objects using Wrapper Classses
# Open a JNB notebook
sp = psm.con.open(lpath=TGT_PATH, close_others=True, visible=False)
# print(sp) # <BaseCOMWrapper for SigmaPlot 15 at SigmaPlot>
# print(sp.path) # "C:\Users\YOUR_LOGIN_NAME\Desktop\dev_20250326_193346.JNB"
notebooks = sp.Notebooks_obj
notebook = notebooks[notebooks.find_indices(TGT_FILENAME)[0]]
notebookitems = notebook.NotebookItems_obj
# # print(notebookitems.list)
# ['Notebook',
#  'section',
#  'worksheet',
#  'graph',
#  'SetLabelsMacro',
#  'SetFigureSizeMacro',
#  '_SetScalesMacro',
#  '_SetRangesMacro',
#  '_SetColorsMacro']
## Instanciates item objects
worksheet = notebookitems["worksheet"]
# print(worksheet) # <WorksheetItemWrapper for worksheet at SigmaPlot.Notebooks[4].NotebookItems[worksheet]>
graph = notebookitems["graph"]
# print(graph) # <GraphItemWrapper for graph at SigmaPlot.Notebooks[4].NotebookItems[graph]>
MACRO_NAMES = [
    # Removers
    "RemoveTitleMacro",
    "RemoveLegendMacro",
    "RemoveSpinesMacro",

    # Figure
    "SetFigureSizeMacro",

    # Labels
    "SetLabelSizesMacro",
    "SetLabelsMacro",

    # Ticks
    "SetTickSizesMacro",
    "SetTicksMacro",

    # Axes
    "SetRangesMacro",
    "SetScalesMacro",

    # Plot
    "SetColorsMacro",
]
macros = [notebookitems[macro_name] for macro_name in MACRO_NAMES]
# set_labels_macro = notebookitems["SetLabelsMacro"]
# print(set_labels_macro) # <MacroItemWrapper for SetLabelsMacro at SigmaPlot.Notebooks[4].NotebookItems[SetLabelsMacro]>

# Demo data
xlabel = "XLABEL specified in Python"
xmm = 40
xscale = "linear"
xmin = 0
xmax = 21
xticks = [0, 5, 10, 15, 20]
params_dict = {
    "xlabel": xlabel,
    "xmm": xmm,
    "xscale": xscale,
    "xmin": xmin,
    "xmax": xmax,
    "xticks": xticks,
}
df = psm.data.create_padded_df(params_dict)

# Main
worksheet.import_data(df)
# set_labels_macro.run()
[macro.run() for macro in macros]
graph.export_as_tif(path=None, crop=True, convert_from_bmp=True)
# 'C:\\Users\\wyusu\\Desktop\\dev_20250326_200042_cropped.tiff'

psm.con.close_all()

# EOF