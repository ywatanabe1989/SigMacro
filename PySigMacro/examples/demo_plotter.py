#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-09 20:49:28 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/examples/demo_plotter.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/examples/demo_plotter.py"
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
import pandas as pd
import numpy as np

def main():
    # 1. Create sample data (replace with your actual data)
    line_dict = {
        'x': np.linspace(0, 10, 20),
        'y': np.sin(np.linspace(0, 10, 20)) * 2 + 5,
        'bgra': psm.utils.get_BGRA("blue", alpha=1.),
    }

    scatter_dict = {
        'x': np.random.rand(15) * 10,
        'y': np.random.rand(15) * 10 + 2,
        'bgra': psm.utils.get_BGRA("green", alpha=1.),
    }

    bar_dict = {
        'x': [i+1 for i in range(5)],
        'y': np.random.rand(5) * 8 + 1,
        'yerr': np.random.rand(5) * 1 + 0.5,
        'bgra': psm.utils.get_BGRA("red", alpha=0.5),
    }


    # 2. Instantiate the Plotter
    plotter = psm.Plotter()

    # 3. Add plot layers
    # Add a line plot using df_line
    plotter.add('line', line_dict)

    # Add a scatter plot using df_scatter
    plotter.add('scatter', scatter_dict)

    # Example: Add bar plot - Note: Bar plots often need specific data format
    # For the demo structure, categorical 'x' and numeric 'y'/'yerr' are common
    plotter.add('bar', bar_dict) # Uncomment if you want to include a bar plot

    # 4. Set visual parameters (optional)
    plotter.set_params(
        xlabel = "X-Axis Label",
        xrot = 45,
        xmm = 40,
        xscale = "linear",
        xmin = "auto",
        xmax = "auto",
        xticks = ["auto"],
        ylabel = "Y-Axis Label",
        yrot = 0,
        ymm = 40 * 0.7,
        yscale = "linear",
        ymin = "auto",
        ymax = "auto",
        yticks = ["auto"],
    )

    # 5. Render the plot
    # This will create 'my_figure.jnb', 'my_figure.tif', etc. in the specified directory
    output_directory = rf"C:\Users\{os.getlogin()}\Downloads" # CHANGE THIS
    base_filename = "demo_plot_using_plotter"

    try:
        jnb_path = plotter.render(output_dir=output_directory, filename_base=base_filename)
        print(f"SigmaPlot JNB generated at: {jnb_path}")
    except Exception as e:
        print(f"An error occurred during rendering: {e}")

if __name__ == '__main__':
    main()

# EOF