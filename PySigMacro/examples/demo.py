#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-09 18:37:49 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/examples/demo.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/examples/demo.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

# Environmental variables only set if not already defined
VARIABLE_DICT = {
    "SIGMACRO_JNB_PATH": rf"C:\Users\{os.getlogin()}\Documents\SigMacro\SigMacro.JNB",
    "SIGMACRO_TEMPLATES_DIR": rf"C:\Users\{os.getlogin()}\Documents\SigMacro\templates",
    "SIGMAPLOT_BIN_PATH_WIN": rf"C:\Program Files (x86)\SigmaPlot\SPW16\Spw.exe",
}

for k, v in VARIABLE_DICT.items():
    if k not in os.environ:
        os.environ[k] = v


import pysigmacro as psm


def run_demo():
    for plot_type in psm.const.PLOT_TYPES:

        if plot_type in [
            # "area",
            # "bar",
            # "barh",
            # "scatter",
            # "box", "boxh",
            # "line",
            # "line_yerr",
            # "lines_y_many_x",
            # "lines_x_many_y",
            # "polar",
            # "contour",
            # "heatmap",
            "violin",
            # "filled_line",
            # "histogram",
            # "jitter"
        ]:
            n_plots = define_n_plots(plot_type)
            plot_types = [plot_type for _ in range(n_plots)]
            psm.demo.gen_csv(plot_types, save=True)
            psm.demo.gen_jnb(plot_types)


def define_n_plots(plot_type):
    return {
        "contour": 1,
        "heatmap": 1,
        "lines_y_many_x": 1,
        "lines_x_many_y": 1,
        "filled_line": 1,
        "histogram": 3,
        "area": 3,
    }.get(plot_type, len(psm.const.COLORS))


def main():
    run_demo()


if __name__ == "__main__":
    run_demo()

# EOF