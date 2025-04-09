#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-07 21:14:13 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/demo/_gen_csv.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/demo/_gen_csv.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from ..path._to_win import to_win
from ..data._create_padded_df import create_padded_df
from ._gen_data import gen_data
from ._gen_visual_params import gen_visual_params
from ._update_visual_params_with_nice_ticks import (
    update_visual_params_with_nice_ticks,
)
import numpy as np
import pandas as pd
import re

# Demo data generation
# --------------------------------------------------


def gen_csv(plot_types, save=False, plot_type_for_visual_params=None):
    """
    Generate demo data for a given plot type and return as a DataFrame.
    Parameters:
    plot_type (str): The type of plot for which to generate data.
    save (bool): If True, the generated DataFrame will be saved as a CSV file.
    Returns:
    pandas.DataFrame: The generated demo data.
    """

    if not plot_type_for_visual_params:
        plot_type_for_visual_params = plot_types[0]

    # Parameters
    df_visual_params = gen_visual_params(plot_type_for_visual_params)

    # Data
    df_data = gen_data(plot_types)

    # Update when auto specified
    df_visual_params = update_visual_params_with_nice_ticks(
        plot_type_for_visual_params, df_visual_params, df_data
    )

    # Concatenate
    df = create_padded_df(df_visual_params, df_data)

    if save:
        # Saving
        templates_dir = os.getenv(
            "SIGMACRO_TEMPLATES_DIR", os.path.join(__DIR__, "templates")
        )
        templates_csv_dir = os.path.join(templates_dir, "csv")
        fname = "-".join(plot_types)
        spath = to_win(os.path.join(templates_csv_dir, f"{fname}.csv"))
        if os.path.exists(spath):
            os.remove(spath)
        df.to_csv(spath, index=False)
        print(f"Saved to: {spath}")
    return df

# EOF