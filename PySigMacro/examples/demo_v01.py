#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-01 23:32:57 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/examples/demo.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/examples/demo.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

# Environmental variables only set if not already defined
if "SIGMACRO_JNB_PATH" not in os.environ:
    os.environ["SIGMACRO_JNB_PATH"] = rf"C:\Users\{os.getlogin()}\Documents\SigMacro\SigMacro.JNB"
if "SIGMACRO_TEMPLATES_DIR" not in os.environ:
    os.environ["SIGMACRO_TEMPLATES_DIR"] = rf"C:\Users\{os.getlogin()}\Documents\SigMacro\templates"
if "SIGMAPLOT_BIN_PATH_WIN" not in os.environ:
    os.environ["SIGMAPLOT_BIN_PATH_WIN"] = rf"C:\Program Files (x86)\SigmaPlot\SPW16\Spw.exe"

import pysigmacro as psm

UNABAILABLE_PLOT_TYPES = ["violin", "contour", "conf_mat", "filled_line"]

for plot_type in psm.const.PLOT_TYPES:

    if plot_type != "violin":
        continue

    # if plot_type in UNABAILABLE_PLOT_TYPES:
    #     continue

    # if plot_type not in ["box", "boxh"]:
    #     continue

    # if plot_type != "violin":
    #     continue

    # CSV data
    psm.demo.gen_csv(plot_type, save=True)

    # JNB and Figures
    try:
        psm.demo.gen_jnb(plot_type)
    except Exception as e:
        print(f"Creating template for {plot_type} failed")
        print(e)

# EOF