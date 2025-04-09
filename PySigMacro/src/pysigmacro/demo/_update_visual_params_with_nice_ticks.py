#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-09 19:57:05 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/demo/_update_visual_params_with_nice_ticks.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/demo/_update_visual_params_with_nice_ticks.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import pandas as pd
import numpy as np
from ..utils._calculate_nice_ticks import calculate_nice_ticks
from ..const._SCALES import SCALES

AVAILABLE_SCALES = SCALES


# Main
# ------------------------------
def update_visual_params_with_nice_ticks(
    plot_type, df_visual_params, df_data, pad_perc=5
):
    # Nice Ticks when "auto" specified
    is_xticks_auto = df_visual_params["xticks"].iloc[0] == "auto"
    is_yticks_auto = df_visual_params["yticks"].iloc[0] == "auto"

    # Check if plot type is polar
    is_polar = plot_type == "polar"

    # Parameters for nice ticks calculation
    if not is_polar:
        # Update x-axis ticks if auto
        if is_xticks_auto:
            x_nice_ticks, x_padded_min, x_padded_max = _calculate_x_nice_ticks(
                plot_type, df_visual_params, df_data, pad_perc
            )
            df_visual_params = _update_xticks(
                df_visual_params, x_nice_ticks, x_padded_min, x_padded_max
            )

        # Update y-axis ticks if auto
        if is_yticks_auto:
            y_nice_ticks, y_padded_min, y_padded_max = _calculate_y_nice_ticks(
                plot_type, df_visual_params, df_data, pad_perc
            )
            df_visual_params = _update_yticks(
                df_visual_params, y_nice_ticks, y_padded_min, y_padded_max
            )
        return df_visual_params

    # For polar, x and y is inverse.
    # xticks: xaxis ticks | xvalues: degrees | yticks: yaxis ticks | yvalues: radius
    elif is_polar:
        is_theta_ticks_auto = is_xticks_auto
        is_r_ticks_auto = is_yticks_auto

        _calculate_theta_nice_ticks = _calculate_y_nice_ticks
        _calculate_r_nice_ticks = _calculate_x_nice_ticks

        _update_theta_ticks = _update_xticks
        _update_r_ticks = _update_yticks

        # Theta
        if is_theta_ticks_auto:
            theta_nice_ticks, theta_padded_min, theta_padded_max = (
                _calculate_theta_nice_ticks(
                    plot_type, df_visual_params, df_data, pad_perc
                )
            )
            df_visual_params = _update_theta_ticks(
                df_visual_params,
                theta_nice_ticks,
                theta_padded_min,
                theta_padded_max,
            )

        # R
        if is_r_ticks_auto:
            r_padded_min = 0
            r_padded_max = 360
            r_nice_ticks = [0, 90, 180, 270]
            df_visual_params = _update_r_ticks(
                df_visual_params, r_nice_ticks, r_padded_min, r_padded_max
            )

        return df_visual_params


# Core
# ------------------------------


def _calculate_x_nice_ticks(
    plot_type,
    df_visual_params,
    df_data,
    pad_perc,
    numeric_columns=["x", "xerr", "x_lower", "x_upper", "theta"],
):
    """Update x-axis ticks in visual parameters"""
    return _calculate_axis_nice_ticks(
        plot_type,
        df_visual_params,
        df_data,
        pad_perc,
        axis="x",
        numeric_columns=numeric_columns,
    )


def _calculate_y_nice_ticks(
    plot_type,
    df_visual_params,
    df_data,
    pad_perc,
    numeric_columns=["y", "yerr", "y_lower", "y_upper", "r"],
):
    """Update y-axis ticks in visual parameters"""
    return _calculate_axis_nice_ticks(
        plot_type,
        df_visual_params,
        df_data,
        pad_perc,
        axis="y",
        numeric_columns=numeric_columns,
    )


def _update_xticks(df_visual_params, x_nice_ticks, x_padded_min, x_padded_max):
    return _update_axis_ticks(
        df_visual_params,
        x_nice_ticks,
        x_padded_min,
        x_padded_max,
        axis="x",
    )


def _update_yticks(df_visual_params, y_nice_ticks, y_padded_min, y_padded_max):
    return _update_axis_ticks(
        df_visual_params,
        y_nice_ticks,
        y_padded_min,
        y_padded_max,
        axis="y",
    )


# Base
# ------------------------------


def _calculate_axis_nice_ticks(
    plot_type,
    df_visual_params,
    df_data,
    pad_perc,
    axis="x",
    numeric_columns=None,
):
    """
    Calculate nice ticks, min, and max values for a specified axis.

    Args:
        df_visual_params: DataFrame containing visualization parameters
        df_data: DataFrame containing the data to visualize
        pad_perc: Padding percentage to apply around min/max values
        axis: Axis to calculate for ('x' or 'y')
        numeric_columns: List of column names containing numeric data for this axis

    Returns:
        tuple: (nice_ticks, min_padded, max_padded)
    """
    # Default numeric columns for each axis if not specified
    if numeric_columns is None:
        if axis == "x":
            numeric_columns = ["x", "xerr", "x_lower", "x_upper", "theta"]
        if axis == "y":
            numeric_columns = ["y", "yerr", "y_lower", "y_upper", "r"]

    # Define separators and the range for the index
    separators = ["", " ", "-", "."]
    max_index = 64  # Generate indices from 0 to 63
    new_keys = [
        f"{key}{sep}{ii}"
        for sep in separators
        for ii in range(max_index)
        for key in numeric_columns
    ]

    # Add suffixes
    numeric_columns += new_keys

    # Set up axis-specific parameter names
    min_param = f"{axis}min"
    max_param = f"{axis}max"
    scale_param = f"{axis}scale"
    ticks_column = f"{axis}ticks"

    # Check if axis scale is numeric
    scale_specified = df_visual_params.set_index("visual parameter label")[
        "visual parameter value"
    ][scale_param]

    if scale_specified not in AVAILABLE_SCALES:
        print(
            f"Warning: Specified scale {scale_specified} is not available. "
            f"Skipping calculation of {axis} nice ticks."
        )
        return ["auto"], "auto", "auto"

    # Get specified parameters
    min_specified = df_visual_params.set_index("visual parameter label")[
        "visual parameter value"
    ][min_param]

    max_specified = df_visual_params.set_index("visual parameter label")[
        "visual parameter value"
    ][max_param]

    _ticks_specified = df_visual_params[ticks_column]
    ticks_specified = _ticks_specified[~_ticks_specified.isna()].tolist()

    try:
        # Get axis min/max from data
        axis_data = _extract_numeric_values(df_data, numeric_columns)

        if np.isnan(axis_data).all().all():
            print(
                f"All {axis} data is not numeric. Skipping calculation of {axis}_nice_ticks"
            )
            return ticks_specified, min_specified, max_specified

        min_data = float(np.nanmin(axis_data.values))
        max_data = float(np.nanmax(axis_data.values))

        if plot_type == "heatmap":
            if axis == "x":
                nice_ticks = (np.arange(min_data, max_data) + 0.5).tolist()
                return nice_ticks, min_data, max_data
            if axis == "y":
                min_shifted = min_data - 0.5
                max_shifted = max_data + 0.5
                nice_ticks = (
                    np.arange(min_shifted, max_shifted) - 0.5
                ).tolist()
                return nice_ticks, min_shifted, max_data

        # min/max considering actual data range
        if min_specified == "auto":
            data_min = min_data
        else:
            data_min = _prefer_int(np.array([min_specified, min_data]).min())

        if max_specified == "auto":
            data_max = max_data
        else:
            data_max = _prefer_int(np.array([max_specified, max_data]).max())

        # Padding
        pad_amount = (data_max - data_min) * pad_perc / 100.0 / 2
        min_padded = data_min - pad_amount if data_min != 0 else data_min
        max_padded = data_max + pad_amount

        # Calculate nice ticks
        nice_ticks = calculate_nice_ticks(min_padded, max_padded)

        return nice_ticks, min_padded, max_padded

    except Exception as e:
        print(f"Warning: Error calculating {axis}-axis nice ticks:\n{e}")
        return ticks_specified, min_specified, max_specified


def _update_axis_ticks(
    df_visual_params, nice_ticks, padded_min, padded_max, axis="x"
):
    """
    Update the DataFrame of visual parameters with calculated axis ticks, min, and max values.

    Args:
        df_visual_params: DataFrame containing visualization parameters
        nice_ticks: List of calculated nice tick values
        padded_min: Calculated minimum value with padding
        padded_max: Calculated maximum value with padding
        axis: Axis to update ('x' or 'y')

    Returns:
        DataFrame: Updated visual parameters DataFrame
    """
    df = df_visual_params.copy()

    try:
        # Set up axis-specific parameter names
        label_min = f"{axis}min"
        label_max = f"{axis}max"
        label_ticks = f"{axis}ticks"

        # Column Names
        col_name_vis_label = "visual parameter label"
        col_name_vis_value = "visual parameter value"

        # Columns
        col_vis_vals = df.columns.get_loc(col_name_vis_value)
        col_ticks = df.columns.get_loc(label_ticks)

        # Rows
        row_min = df.index[df[col_name_vis_label] == label_min].tolist()[0]

        row_max = df.index[df[col_name_vis_label] == label_max].tolist()[0]

        # Update min
        if padded_min != "auto":
            df.iloc[row_min, col_vis_vals] = _prefer_int(padded_min)

        # Update max
        if padded_max != "auto":
            df.iloc[row_max, col_vis_vals] = _prefer_int(padded_max)

        # Update ticks
        for i_tick, tick in enumerate(nice_ticks):
            if i_tick < len(df):
                df.loc[i_tick, label_ticks] = tick

        return df

    except Exception as e:
        print(f"Error updating {axis}-axis ticks:\n{e}")
        __import__("ipdb").set_trace()
        return df


# Helpers
# ------------------------------


def _prefer_int(float_or_int_value):
    """Convert float to integer if the value close to int"""
    if float(float_or_int_value) == int(float_or_int_value):
        return int(float_or_int_value)
    elif abs(float(float_or_int_value) - round(float_or_int_value)) < 1e-10:
        return int(round(float_or_int_value))
    else:
        return float(float_or_int_value)


def _extract_numeric_values(df, possible_numeric_columns):
    """Extract numeric values from dataframe columns"""
    numeric_data = []
    for i_col, col in enumerate(df.columns):
        try:
            if col in possible_numeric_columns:
                numeric_data.append(
                    pd.to_numeric(df.iloc[:, i_col], errors="coerce")
                )
        except:
            pass

    if len(numeric_data) > 0:
        return pd.concat(numeric_data, axis=1)
    else:
        raise ValueError("Numeric data not found")

# EOF