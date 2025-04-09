#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-09 13:45:48 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/demo/_gen_single_data_filled_line.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/demo/_gen_single_data_filled_line.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import matplotlib
import matplotlib.colors as mcolors
import numpy as np
import pandas as pd

from ..const import BGRA, BGRA_FAKE, COLORS
from ..data._create_padded_df import create_padded_df
from ..data._create_graph_wizard_params import create_graph_wizard_params
from scipy import stats

# Main
# ------------------------------


def _gen_single_data_filled_line(i_plot, alpha=0.1):
    batch_size = 3

    # Upper -> Lower -> Middle
    i_plot_uu = (i_plot * batch_size) + 0
    i_plot_ll = (i_plot * batch_size) + 1
    i_plot_mm = i_plot * batch_size + 2

    # Data
    raw_data_dict = _gen_single_raw_data_filled_line(i_plot, alpha=alpha)
    x = raw_data_dict.pop("x")
    y = raw_data_dict.pop("y")
    y_lower = raw_data_dict.pop("y_lower")
    y_upper = raw_data_dict.pop("y_upper")
    # bgra = raw_data_dict.pop("bgra")
    bgra = BGRA["blue"] # fixme
    bgra[-1] = alpha

    # Split data into filled area and line plot
    uu_dict = {"x": x, "y_upper": y_upper, "bgra": bgra}
    mm_dict = {"x": x, "y": y, "bgra": bgra}
    ll_dict = {"x": x, "y_lower": y_lower, "bgra": BGRA["white"]}

    # Fill using Lines X Many Y
    gw_uu = create_graph_wizard_params("filled_line_uu", i_plot_uu)
    uu_df = create_padded_df(uu_dict)

    gw_mm = create_graph_wizard_params("filled_line_mm", i_plot_mm)
    mm_df = create_padded_df(mm_dict)

    gw_ll = create_graph_wizard_params("filled_line_ll", i_plot_ll)
    ll_df = create_padded_df(ll_dict)

    # Combined (order is matter)
    filled_line_df = create_padded_df(gw_uu, uu_df, gw_ll, ll_df, gw_mm, mm_df)

    return filled_line_df


def _gen_single_raw_data_filled_line(i_plot, alpha=0.5, num_points=50):
    """
    Generates raw data for a filled line plot based on sine waves.

    Parameters:
    ----------
    i_plot : int
        Index to differentiate the plot (affects phase shift and color).
    alpha : float, optional
        Alpha transparency for the fill color (0 to 1). Default 0.5.
    num_points : int, optional
        Number of points along the x-axis. Default 50.

    Returns:
    -------
    dict
        Dictionary containing 'x', 'y', 'y_lower', 'y_upper', and 'bgra'.
        Returns an empty dict if color lookup fails.
    """
    # --- X Coordinates ---
    # Generate points from 0 to 4*pi
    x_coords = np.linspace(0, 4 * np.pi, num_points)

    # --- Y Coordinates (Center Line) ---
    # Base sine wave
    base_frequency = 1.0
    # Shift phase based on i_plot
    phase_shift = (np.pi / 4) * i_plot
    # Add vertical shift based on i_plot
    vertical_shift = i_plot * 0.5
    y_coords = np.sin(base_frequency * x_coords + phase_shift) + vertical_shift

    # --- Y Bounds (Lower and Upper for Fill) ---
    # Define the 'thickness' or 'error band' around the center line
    # Example: Use a smaller amplitude sine wave or a constant offset
    band_amplitude = 0.3 + i_plot * 0.05
    y_lower_coords = y_coords - band_amplitude
    y_upper_coords = y_coords + band_amplitude

    # --- Color ---
    # Get the base color based on i_plot
    bgra = BGRA[COLORS[i_plot % len(COLORS)]]
    bgra[-1] = alpha

    return dict(
        x=x_coords,
        y_lower=y_lower_coords,
        y=y_coords,
        y_upper=y_upper_coords,
        bgra=bgra,
    )


# def _gen_single_raw_data_filled_line(ii, alpha=0.5):
#     # Random Seed
#     np.random.seed(42)
#     # X
#     x = np.linspace(0, 10, 20) + ii
#     # Y
#     y = np.exp(-((x - 5 * (ii % 3)) ** 2) / 10)
#     y_lower = y - np.random.normal(0, 0.05 * (ii + 1), size=len(x))
#     y_upper = y + np.random.normal(0, 0.05 * (ii + 1), size=len(x))

#     # Color
#     bgra = BGRA[COLORS[ii % len(COLORS)]]
#     bgra[-1] = alpha
#     return dict(
#         x=x,
#         y_lower=y_lower,
#         y=y,
#         y_upper=y_upper,
#         bgra=bgra,
#     )

# EOF