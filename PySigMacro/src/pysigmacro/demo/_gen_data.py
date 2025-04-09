#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-09 18:23:31 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/demo/_gen_data.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/demo/_gen_data.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import numpy as np

from ..const import BGRA, BGRA_FAKE, COLORS
from ..data._create_padded_df import create_padded_df
from ..data._create_graph_wizard_params import create_graph_wizard_params
from scipy import stats
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.colors import to_rgba
from ._gen_data_heatmap import _gen_data_heatmap
from ._gen_single_data_violin import (
    _gen_single_data_violin,
    _gen_single_data_violinh,
)
from ._gen_single_data_filled_line import _gen_single_data_filled_line

# Main
# ------------------------------


def gen_data(plot_types):

    chunks_df = pd.DataFrame()
    for i_plot, plot_type in enumerate(plot_types):
        try:
            gen_single_data_func = {
                "scatter": _gen_single_data_scatter,
                "jitter": _gen_single_data_jitter,
                "line": _gen_single_data_line,
                "line_yerr": _gen_single_data_line_yerr,
                "lines_y_many_x": _gen_single_data_lines_y_many_x,
                "lines_x_many_y": _gen_single_data_lines_x_many_y,
                "bar": _gen_single_data_bar,
                "barh": _gen_single_data_barh,
                "histogram": _gen_single_data_histogram,
                "area": _gen_single_data_area,
                "box": _gen_single_data_box,
                "boxh": _gen_single_data_boxh,
                "polar": _gen_single_data_polar,
                "violin": _gen_single_data_violin,
                "violinh": _gen_single_data_violinh,
                "heatmap": _gen_data_heatmap,
                "filled_line": _gen_single_data_filled_line,
                "contour": _gen_single_data_contour,
            }[plot_type]

            if plot_type == "heatmap":
                chunks_df = gen_single_data_func(i_plot)

            if plot_type in ["violin", "violinh"]:
                single_chunk_df = gen_single_data_func(i_plot)
                chunks_df = create_padded_df(chunks_df, single_chunk_df)

            if plot_type == "filled_line":
                single_chunk_df = gen_single_data_func(i_plot)
                chunks_df = create_padded_df(chunks_df, single_chunk_df)

            else:
                single_plot_dict = gen_single_data_func(i_plot)
                single_plot_df = create_padded_df(single_plot_dict)
                gw_df = create_graph_wizard_params(plot_type, i_plot)
                single_chunk = create_padded_df(gw_df, single_plot_df)
                chunks_df = create_padded_df(chunks_df, single_chunk)
        except Exception as e:
            print(plot_type)
            print(e)

    return chunks_df


# Special
# ------------------------------


def _gen_single_data_lines_y_many_x(i_plot, alpha=0.5):
    # Random Seed for reproducibility based on index
    np.random.seed(42)

    # Generate multiple X arrays
    x_values = {}
    num_x_lines = 6
    num_points = 50  # Increased points for smoother sine wave
    for x_index in range(num_x_lines):
        # Generate x values relative to the index for variation
        x_values[f"x{x_index}"] = (
            np.linspace(0, 4 * np.pi, num_points) + x_index * np.pi / 4
        )

    # Calculate Y based on x0 as a shifted sine curve
    x0 = x_values["x0"]
    # Calculate phase shift based on i_plot
    phase_shift = i_plot * np.pi / 3
    # Calculate sine wave
    y_values = (
        np.sin(x0 + phase_shift) + np.random.rand(num_points) * 0.2
    )  # Add some noise

    # Determine color based on index
    color_name = COLORS[i_plot % len(COLORS)]
    bgra_color = BGRA[
        color_name
    ].copy()  # Use copy to avoid modifying the original
    bgra_color[-1] = alpha  # Set alpha

    return dict(
        y=y_values,
        **x_values,
        bgra=bgra_color,
    )


def _gen_single_data_lines_x_many_y(i_plot, alpha=0.5):
    dd = _gen_single_data_lines_y_many_x(i_plot, alpha=alpha)
    y_dict = {f"y{k[1:]}": v for k, v in dd.items() if k.startswith("x")}
    return dict(
        x=dd["y"],
        **y_dict,
        bgra=dd["bgra"],
    )


def _gen_single_data_contour(ii):
    # Random Seed
    np.random.seed(42)

    # Create grid data
    x = np.linspace(-5, 5, 10)
    y = np.linspace(-5, 5, 10)
    X, Y = np.meshgrid(x, y)

    # Create Z values (peaks with noise)
    sigma_x = 1.0 + 0.2 * ii
    sigma_y = 1.0 + 0.1 * ii

    # Create multiple peaks
    peaks = [
        (
            3,
            3,
            1 + 0.5 * ii,
            sigma_x,
            sigma_y,
        ),
        (-2, 2, 0.8 + 0.3 * ii, sigma_x * 0.8, sigma_y * 1.2),
        (0, -3, 0.9 + 0.4 * ii, sigma_x * 1.2, sigma_y * 0.8),
        (-3, -2, 0.7 + 0.2 * ii, sigma_x * 0.9, sigma_y * 0.9),
    ]

    Z = np.zeros_like(X)
    for x0, y0, height, sx, sy in peaks:
        Z += height * np.exp(
            -((X - x0) ** 2 / (2 * sx**2) + (Y - y0) ** 2 / (2 * sy**2))
        )

    # Add noise
    noise_level = 0.05 * (ii + 1)
    Z += np.random.normal(0, noise_level, Z.shape)

    # Convert to xyz format
    x_flat = X.flatten()
    y_flat = Y.flatten()
    z_flat = Z.flatten()

    return dict(
        x=x_flat,
        y=y_flat,
        z=z_flat,
        bgra=BGRA_FAKE,
    )


# Single
# ------------------------------
def _gen_single_data_bar(ii):
    # Random Seed
    np.random.seed(42)
    # X
    x = f"X {ii}"
    # Y
    y = 1.0 * (ii + 1) + np.random.normal(0, 0.3 * (ii + 1))
    yerr = 0.1 * (ii + 1)
    return dict(
        x=x,
        y=y,
        yerr=yerr,
        bgra=BGRA[COLORS[ii % len(COLORS)]],
    )


def _gen_single_data_barh(ii):
    # Random Seed
    np.random.seed(42)
    vv = _gen_single_data_bar(ii)
    return dict(
        y=vv["x"],
        x=vv["y"],
        xerr=vv["yerr"],
        bgra=BGRA[COLORS[ii % len(COLORS)]],
    )


def _gen_single_data_area(plot_idx, alpha=0.5, x_shift=0):
    """
    Generate area plot data for demonstration using probabilistic distributions.

    Args:
        plot_idx: Index for random seed and distribution selection
        alpha: Transparency value for the area plot
        x_shift: Value to shift all x positions

    Returns:
        Dictionary with x values, y values, and color (bgra) for area plot
    """
    # Set random seed for reproducibility
    np.random.seed(42 + plot_idx)

    # Generate x values as a linear sequence
    x_values = np.linspace(0, 20, 200) + x_shift

    # Generate different distributions based on plot_idx
    if plot_idx % 3 == 0:
        # Normal distribution
        mean = 5 * (plot_idx % 2 + 1) + x_shift
        std_dev = 1.0 + 0.5 * (plot_idx % 3)
        y_values = np.exp(-((x_values - mean) ** 2) / (2 * std_dev**2))

    elif plot_idx % 3 == 1:
        # Bimodal distribution
        mean1 = 5 + x_shift
        mean2 = 12 + x_shift
        std_dev = 1.5
        y_values = 0.6 * np.exp(
            -((x_values - mean1) ** 2) / (2 * std_dev**2)
        ) + 0.4 * np.exp(-((x_values - mean2) ** 2) / (2 * std_dev**2))

    else:
        # Gamma-like distribution
        shape = 2.0 + 0.5 * plot_idx
        rate = 0.5
        # Ensure x values are positive for gamma
        x_for_gamma = np.maximum(x_values - x_shift, 0.01)
        y_values = (x_for_gamma ** (shape - 1)) * np.exp(-rate * x_for_gamma)

    # Normalize sum of y values to 1
    if np.max(y_values) > 0:
        y_values = y_values / np.sum(y_values)

    # Select color based on plot index
    bgra = BGRA[COLORS[(plot_idx + 3) % len(COLORS)]]
    bgra[-1] = alpha

    return dict(
        x=x_values,
        y=y_values,
        bgra=bgra,
    )


def _gen_single_data_box(ii):
    # Random Seed
    np.random.seed(42)
    # X
    x = f"Category {ii}"
    # Y
    # Generate data from uniform distribution to emphasize box plot visualization
    low = 3 * (ii + 1)
    high = 8 * (ii + 1)
    base_data = np.random.uniform(low, high, 30)
    # Add a few outliers outside the uniform range
    outliers_low = np.random.uniform(low - 2, low - 1, 1)
    outliers_high = np.random.uniform(high + 1, high + 2, 1)
    y = np.concatenate([base_data, outliers_low, outliers_high])
    return dict(
        x=x,
        y=y,
        bgra=BGRA[COLORS[ii % len(COLORS)]],
    )


def _gen_single_data_boxh(ii):
    vv = _gen_single_data_box(ii)
    return dict(
        y=vv["x"],
        x=vv["y"],
        bgra=BGRA[COLORS[ii % len(COLORS)]],
    )


def _gen_single_data_line(ii):
    # Random Seed
    np.random.seed(42)
    # X
    x = np.linspace(0, 10, 20)
    # Y
    y = np.sin(x + ii * 0.5) * (ii + 1)
    y += np.random.normal(0, 0.1 * (ii + 1), size=len(x))
    return dict(
        x=x,
        y=y,
        bgra=BGRA[COLORS[ii % len(COLORS)]],
    )


def _gen_single_data_line_yerr(ii):
    # Random Seed
    np.random.seed(42)
    # X
    x = np.linspace(0, 10, 20)
    # Y
    y = np.sin(x + ii * 0.5) * (ii + 1)
    y += np.random.normal(0, 0.1 * (ii + 1), size=len(x))
    yerr = 0.2 * np.ones_like(x) * (1 + 0.1 * ii)
    return dict(
        x=x,
        y=y,
        yerr=yerr,
        bgra=BGRA[COLORS[ii % len(COLORS)]],
    )


def _gen_single_data_polar(ii):
    # Random Seed
    np.random.seed(42)

    # X
    theta = np.linspace(0, 2 * np.pi, 30)
    degree = theta / (2 * np.pi) * 360

    # Y
    r = 0.5 + ii + 0.5 * np.sin(theta * (ii + 1))
    r_fluctuation = np.random.normal(0, 0.1 * (ii + 1), size=len(theta))
    r = r + r_fluctuation
    return dict(
        theta=degree,
        r=r,
        bgra=BGRA[COLORS[ii % len(COLORS)]],
    )


def _gen_single_data_scatter(ii):
    # Random Seed
    np.random.seed(42)
    n_points = 30 + ii * 5
    # X
    center_x = 5 * (ii % 3)
    x = np.random.normal(center_x, 1 + 0.2 * ii, n_points)
    # Y
    center_y = 5 * (ii // 3)
    y = np.random.normal(center_y, 1, n_points) + ii * 0.1 * x
    return dict(
        x=x,
        y=y,
        bgra=BGRA[COLORS[ii % len(COLORS)]],
    )


def _gen_single_data_histogram(
    i_plot, alpha=0.5, bin_count=15, bin_width=1, bin_range=None, x_shift=0
):
    """Generate histogram data for demonstration.

    Args:
        i_plot: Index for random seed and parameter variation
        alpha: Transparency value for the histogram
        bin_count: Number of bins to use (ignored if bin_width is specified)
        bin_width: Exact width for each bin (overrides bin_count)
        bin_range: Tuple of (min, max) to set the range of the histogram
        x_shift: Value to shift all x positions (affects bin positions)

    Returns:
        Dictionary with data for a histogram plot
    """
    # Set random seed for reproducibility
    np.random.seed(42)

    # Generate different distributions based on i_plot
    if i_plot % 3 == 0:
        # Normal distribution
        mean = 5 * (i_plot % 2 + 1)
        std_dev = 1.0 + 0.5 * (i_plot % 3)
        data = np.random.normal(mean, std_dev, 1000)
    elif i_plot % 3 == 1:
        # Bimodal distribution
        mean1 = 2 + i_plot
        mean2 = 8 + i_plot
        std_dev = 1.0
        samples1 = np.random.normal(mean1, std_dev, 500)
        samples2 = np.random.normal(mean2, std_dev, 500)
        data = np.concatenate([samples1, samples2])
    else:
        # Skewed distribution
        shape = 2.0 + 0.5 * i_plot
        scale = 2.0
        data = np.random.gamma(shape, scale, 1000)

    # Apply x_shift to the data if specified
    if x_shift != 0:
        data = data + x_shift

    # Calculate histogram with specified bin parameters
    if bin_width is not None:
        # Create bins with exact width
        if bin_range is None:
            # Auto-detect range if not specified
            bin_min = np.floor(data.min())
            bin_max = np.ceil(data.max())
        else:
            bin_min, bin_max = bin_range

        # Create bins of exact width
        num_bins = int(np.ceil((bin_max - bin_min) / bin_width))
        bins = np.linspace(
            bin_min, bin_min + num_bins * bin_width, num_bins + 1
        )
    else:
        # Use bin count with optional range
        bins = (
            bin_count
            if bin_range is None
            else np.linspace(bin_range[0], bin_range[1], bin_count + 1)
        )

    # Calculate histogram data
    hist, bin_edges = np.histogram(data, bins=bins)

    # Calculate bin centers for plotting
    bin_centers = (bin_edges[:-1] + bin_edges[1:]) / 2

    # Get color based on index
    color_name = COLORS[(i_plot + 3) % len(COLORS)]
    bgra_color = BGRA[color_name].copy()
    bgra_color[-1] = alpha

    return dict(
        x=bin_centers,
        y=hist,
        bgra=bgra_color,
    )


def _gen_single_data_jitter(i_plot, jitter_width=0.2, alpha=0.8):
    """Generate data for a jitter plot (scatter plot with categorical x-axis).

    Args:
        i_plot: Index for plot variation
        jitter_width: Width of the jitter (scatter spread) around each category
        alpha: Transparency of points

    Returns:
        Dictionary with data for a scatter plot with jittered x positions
    """
    # Random seed for reproducibility
    np.random.seed(42 + i_plot)

    # Points per category
    points_per_category = 20

    # Generate data for just one category
    category = f"Category {i_plot+1}"

    # Base position for this category
    x_position = i_plot + 1

    # Generate data points with characteristics based on i_plot
    base_mean = 5 + (i_plot * 2)
    base_std = 1.0 + (0.2 * i_plot)

    # Generate y values from normal distribution
    y_values = np.random.normal(base_mean, base_std, points_per_category)

    # Create jittered x positions around the category position
    x_positions = np.full(points_per_category, x_position)
    jitter = np.random.uniform(
        -jitter_width, jitter_width, points_per_category
    )
    jittered_positions = x_positions + jitter

    # Get color based on index
    color_name = COLORS[i_plot % len(COLORS)]
    bgra_color = BGRA[color_name].copy()
    bgra_color[-1] = alpha

    # Return data with jittered x positions and category mapping
    return dict(
        x=jittered_positions,  # Numerical positions with jitter
        y=y_values,
        bgra=bgra_color,
    )

# EOF