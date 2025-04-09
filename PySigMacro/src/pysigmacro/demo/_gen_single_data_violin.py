#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-09 18:57:09 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/demo/_gen_single_data_violin.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/demo/_gen_single_data_violin.py"
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

def _gen_single_data_violin(i_plot, violin_half_width=1., box_alpha=0.5, kde_alpha=0.7):
    batch_size = 2

    i_plot_kde_left = i_plot * batch_size + 0
    i_plot_violin = (i_plot * batch_size) + 1

    # Data Calculation
    raw_data_dict = _gen_single_raw_data_violin(i_plot_violin, box_alpha)
    kde_data_dict = _calculate_kde(
        i_plot_violin,
        raw_data_dict,
        num_kde_points=100,
        violin_half_width=violin_half_width,
        alpha=kde_alpha,
    )

    # Alpha
    bgra_box = raw_data_dict["bgra"]
    bgra_box[-1] = box_alpha
    raw_data_dict["bgra"] = bgra_box

    # Formatting for pysigmaplot
    # Box
    gw_box = create_graph_wizard_params("box_violin", i_plot_violin)
    box_df = create_padded_df(raw_data_dict)

    # KDE
    gw_kde = create_graph_wizard_params("lines_y_many_x_violin", i_plot_kde_left)
    kde_df = create_padded_df(kde_data_dict)

    violin_df = create_padded_df(gw_kde, kde_df, gw_box, box_df)

    return violin_df


def _gen_single_data_violinh(i_plot):
    violin_df = _gen_single_data_violin(i_plot)
    violinh_df = None
    return violinh_df


def _gen_single_raw_data_violin(i_plot, box_alpha=1.0):
    """
    Generates categorical raw data suitable for SigmaPlot's violin plot.
    SigmaPlot handles KDE and box plot generation from these raw points.
    """
    # Random Seed - use the same seed as box plot for comparable data
    np.random.seed(42)
    # X category label
    x_category = f"Category {i_plot}"  # Use a numeric category index for plotting position
    # x_position = i_plot + 1  # Position on the X-axis (1-based)
    x_position = i_plot + 1  # Position on the X-axis (1-based)

    # Y data points (same generation logic as _gen_single_data_box)
    mean = 5 * (i_plot + 1)
    std_dev = 1.5 * (i_plot + 1)
    num_points = 50
    y_values = np.random.normal(loc=mean, scale=std_dev, size=num_points)
    # Add a couple of potential outliers
    y_values = np.concatenate(
        [y_values, np.array([mean - 3 * std_dev, mean + 3 * std_dev])]
    )

    bgra = BGRA[COLORS[i_plot % len(COLORS)]]
    bgra[-1] = box_alpha

    # Return the category identifier (numeric position) and the raw data points
    return dict(
        # Use numeric x for positioning, category label might be handled elsewhere
        x=x_position,
        y=y_values,  # Array of raw data points for this category
        bgra=bgra,
    )


def _calculate_kde(
        i_plot, raw_data_dict, num_kde_points=100, violin_half_width=1, alpha=1.0
):
    """
    Calculates the Kernel Density Estimate (KDE) shape for a given data category.

    Parameters
    ----------
    i_plot : int
        Index for the data category (used to generate data via _gen_single_raw_data_violin).
    num_kde_points : int, optional
        Number of points to evaluate the KDE at, by default 100.
    violin_half_width : float, optional
        The maximum half-width of the violin shape, relative to the category center,
        by default 0.4 (total width would be 0.8).

    Returns
    -------
    dict or None
        A dictionary containing the KDE shape coordinates:
        'y_kde': The y-values where the density was evaluated.
        'x_left': The x-coordinates for the left edge of the violin shape.
        'x_right': The x-coordinates for the right edge of the violin shape.
        'x_center': The center x-position of the violin.
        Returns None if KDE cannot be computed (e.g., insufficient data variance).
    """
    # --- Generate Raw Data ---
    # Uses the same seed logic as the simple generator

    data_points = raw_data_dict["y"]
    x_position = raw_data_dict["x"]

    # --- Validate Data for KDE ---
    if len(data_points) < 2:
        print(
            f"Warning: Insufficient data points ({len(data_points)}) for KDE calculation for index {i_plot}."
        )
        return None
    data_std_dev = np.std(data_points)
    if data_std_dev == 0:
        print(
            f"Warning: Data has zero standard deviation for index {i_plot}. Cannot compute KDE."
        )
        return None

    # --- Calculate KDE ---
    try:
        # Use scipy's built-in bandwidth estimation (e.g., Silverman's rule)
        # You could manually calculate bandwidth as in the original snippet if needed
        kde = stats.gaussian_kde(data_points, bw_method="silverman")

        # Define the range over which to evaluate the KDE
        kde_y_min = (
            np.min(data_points) - 1 * data_std_dev
        )  # Extend slightly beyond data range
        kde_y_max = np.max(data_points) + 1 * data_std_dev
        y_values_kde = np.linspace(kde_y_min, kde_y_max, num_kde_points)

        # Evaluate the KDE to get density values
        density = kde(y_values_kde)

    except Exception as e:
        print(f"Error during KDE calculation for index {i_plot}: {e}")
        return None

    # --- Scale Density for Plotting ---
    max_density = np.max(density)
    if max_density > 0:
        # Scale density values so the max density corresponds to violin_half_width
        scaled_density = (density / max_density) * violin_half_width
    else:
        # Handle case where density is zero everywhere (shouldn't happen with valid data)
        scaled_density = np.zeros_like(density)

    # --- Calculate Violin Edge Coordinates ---
    x_left_edge = x_position - scaled_density
    x_right_edge = x_position + scaled_density

    # --- Return KDE Shape Data ---
    bgra = BGRA[COLORS[i_plot % len(COLORS)]]
    bgra[-1] = alpha
    return dict(
        y=y_values_kde,  # Y-axis coordinates (where density is evaluated)
        x_lower=x_left_edge,  # X-axis coordinates for the left edge
        x_upper=x_right_edge,  # X-axis coordinates for the right edge
        bgra=bgra,
    )


if __name__ == "__main__":
    import pysigmacro as psm
    from pysigmacro.data import create_padded_df
    from pysigmacro.data import create_graph_wizard_params

# EOF