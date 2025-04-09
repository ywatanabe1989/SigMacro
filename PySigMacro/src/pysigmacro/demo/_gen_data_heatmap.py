#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-09 13:45:49 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/demo/_gen_data_heatmap.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/demo/_gen_data_heatmap.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import matplotlib
import matplotlib.colors as mcolors
import numpy as np
import pandas as pd

from ..const import BGRA_FAKE
from ..data._create_padded_df import create_padded_df
from ..data._create_graph_wizard_params import create_graph_wizard_params

# Main
# ------------------------------


def _gen_data_heatmap(_ii, n_cols=4, n_rows=3, cbar=False):
    assert _ii == 0

    # Random seed for reproducibility
    np.random.seed(42)

    df_2d = _gen_simple_heatmap_2d(n_cols, n_rows).copy()

    df_xyz = _to_xyz(df_2d).copy()

    df_xyz_and_coords = _add_xy_coords_for_category(df_xyz).copy()

    df_fill_coords = _gen_coords_for_filling(n_cols, n_rows)

    df_bgra_fake = pd.DataFrame(
        columns=["bgra_fake"], data=pd.Series(BGRA_FAKE)
    )


    _df_cbar = _gen_colorbar(df_2d, n=100)

    df = create_padded_df(df_xyz_and_coords, _df_cbar, df_fill_coords, df_bgra_fake)
    # df = create_padded_df(df_xyz_and_coords, _df_cbar, df_bgra_fake)

    df, df_cbar = _convert_to_pysigmacro_format(df)

    if cbar:

        return df, df_cbar
    else:
        return df


# Subroutines
# ------------------------------


def _gen_simple_heatmap_2d(ncols=4, nrows=3):
    """
    Generates a Pandas DataFrame with random values for heatmap visualization.

    Args:
        ncols (int): Number of columns for the DataFrame. Defaults to 4.
        nrows (int): Number of rows for the DataFrame. Defaults to 3.

    Returns:
        pd.DataFrame: A DataFrame with random float values,
                      named columns (Col1, Col2, ...),
                      and named rows (Row1, Row2, ...).
    """
    # Generate random data
    data = np.random.rand(nrows, ncols)

    # Create column names
    col_names = [f"Col{i+1}" for i in range(ncols)]

    # Create row names (index)
    row_names = [f"Row{i+1}" for i in range(nrows)]

    # Create DataFrame
    df = pd.DataFrame(data, index=row_names, columns=col_names)

    return df


def _to_xyz(
    df,
    xlabel="x",
    ylabel="y",
    zlabel="z",
):

    # Reset index to make row names a column
    df_reset = df.reset_index()

    # Melt the DataFrame to long format (XYZ)
    df_xyz = pd.melt(
        df_reset, id_vars=["index"], var_name=xlabel, value_name=zlabel
    )

    # Rename columns for clarity (optional, but conventional for XYZ)
    df_xyz = df_xyz.rename(columns={"index": ylabel})

    # Reorder columns if desired (e.g., X, Y, Z)
    df_xyz = df_xyz[[xlabel, ylabel, zlabel]]

    return df_xyz


def _add_xy_coords_for_category(df):
    n_cols = len(df["x"].unique())
    n_rows = len(df["y"].unique())

    x_codes, x_categories = pd.factorize(df["x"])
    # For scatter
    df["x_coord"] = x_codes + 1.0

    y_codes, y_categories = pd.factorize(df["y"])
    # For scatter
    df["y_coord"] = y_codes + 1.0

    df = df.rename(
        columns={
            "x": "xlabel",
            "y": "ylabel",
            "x_coord": "x",
            "y_coord": "y",
        }
    ).copy()
    df = df[["x", "y", "z", "xlabel", "ylabel"]].copy()

    return df


def _gen_colorbar(df, n=100):
    """
    Generates n color steps for a colorbar based on the range
    of values in a DataFrame using the 'Blues' colormap.
    RGB values are scaled to 0-255 and rounded.

    Args:
        df (pd.DataFrame): The DataFrame containing the data values.
        n (int): The number of color steps to generate for the colorbar.

    Returns:
        pd.DataFrame: A DataFrame with columns ['r', 'g', 'b', 'a', 'z'].
                      Each row represents one of the n color steps.
                      'r', 'g', 'b' are integer color components (0-255 range).
                      'a' is the alpha component (0-1 range).
                      'z' is the corresponding data value.
                      Returns an empty DataFrame if n <= 0 or min/max
                      cannot be determined.
    """

    if n <= 0:
        # Return an empty DataFrame if n is not positive
        return pd.DataFrame(columns=["r", "g", "b", "a", "z"])

    # Calculate the minimum and maximum values in the DataFrame
    try:
        # Ensure only numeric data is considered for min/max
        numeric_df = df.select_dtypes(include=np.number)
        if numeric_df.empty:
            raise ValueError("DataFrame contains no numeric data.")
        # Drop NaN values before calculating min/max to avoid warnings/errors
        vmin = np.nanmin(numeric_df.values)
        vmax = np.nanmax(numeric_df.values)
        # Check if min/max calculation resulted in NaN (e.g., all NaNs in df)
        if np.isnan(vmin) or np.isnan(vmax):
            raise ValueError("Could not determine valid min/max (all NaNs?).")
    except (TypeError, ValueError, AttributeError) as e:
        # Handle cases where min/max cannot be computed
        print(
            f"Warning: Could not determine min/max from DataFrame values. Error: {e}"
        )
        return pd.DataFrame(columns=["r", "g", "b", "a", "z"])

    # Handle the edge case where min and max are the same (after handling NaNs)
    if vmin == vmax:
        # Generate n values, all equal to vmin
        values = np.full(n, vmin)
    elif n == 1:
        # Generate a single value (e.g., the minimum) if n is 1
        values = np.array([vmin])
    else:
        # Generate n evenly spaced values between min and max
        values = np.linspace(vmin, vmax, n)

    # Get the 'Blues' colormap
    # cmap = plt.cm.get_cmap('Blues')
    cmap = matplotlib.colormaps["Blues"]

    # Create a Normalize instance to map values to the [0, 1] range
    # Handle the case vmin == vmax for normalization
    if vmin == vmax:
        # If all values are the same, map them consistently (e.g., to the start)
        normalized_values = np.zeros(n)
    else:
        norm = mcolors.Normalize(vmin=vmin, vmax=vmax)
        normalized_values = norm(values)

    # Apply the colormap to the normalized values to get RGBA tuples (range 0-1)
    rgba_colors = cmap(normalized_values)

    # Create the DataFrame
    colorbar_df = pd.DataFrame(
        {
            # Scale RGB to 0-255 and round to nearest integer
            "r": np.round(rgba_colors[:, 0] * 255),
            "g": np.round(rgba_colors[:, 1] * 255),
            "b": np.round(rgba_colors[:, 2] * 255),
            # Keep Alpha in 0-1 range
            "a": rgba_colors[:, 3],
            "z": values,
        }
    )

    # Ensure RGB columns are integer type
    colorbar_df[["r", "g", "b"]] = colorbar_df[["r", "g", "b"]].astype(int)

    colorbar_df = colorbar_df.rename(columns={"z": "ycbar", "a": "alpha"})
    colorbar_df["xcbar"] = 1

    colorbar_df = colorbar_df[["ycbar", "xcbar", "b", "g", "r", "alpha"]]

    return colorbar_df


def _gen_coords_for_filling(ncols, nrows):
    x_lower = np.arange(ncols) + 0.5
    x_upper = x_lower + 1
    y = np.arange(nrows) + 1.5
    df = pd.DataFrame(
        columns=["x_lower_fill", "x_upper_fill", "y_fill"],
        data=np.array(
            [
                (float(xl), float(xu), float(j))
                for xl, xu in zip(x_lower, x_upper)
                for j in y
            ]
        ),
    )
    return df


def _apply_colormap_to_values(values, vmin, vmax, cmap_name="Blues"):
    """
    Maps an array of values to BGRA colors using a specified colormap and range.

    Args:
        values (np.array or pd.Series): The data values to map to colors.
        vmin (float): The minimum value of the color scale.
        vmax (float): The maximum value of the color scale.
        cmap_name (str): Name of the Matplotlib colormap. Defaults to 'Blues'.

    Returns:
        pd.DataFrame: A DataFrame with columns ['b', 'g', 'r', 'alpha']
                      corresponding to the input values. Returns empty DataFrame
                      if input values are empty.
    """
    if len(values) == 0:
        return pd.DataFrame(columns=["b", "g", "r", "alpha"])

    # Get the specified colormap
    try:
        cmap = matplotlib.colormaps[cmap_name]
    except KeyError:
        print(f"Warning: Colormap '{cmap_name}' not found. Using 'viridis'.")
        cmap = matplotlib.colormaps["viridis"]

    # Create a Normalize instance
    # Handle the case vmin == vmax for normalization
    if vmin == vmax:
        # Map all to the start color if min == max
        normalized_values = np.zeros(len(values))
    else:
        norm = mcolors.Normalize(vmin=vmin, vmax=vmax)
        normalized_values = norm(values)

    # Apply the colormap
    rgba_colors = cmap(normalized_values)

    # Create DataFrame with BGRA colors (scaled to 0-255 for BGR)
    color_df = pd.DataFrame(
        {
            "r": np.round(rgba_colors[:, 0] * 255),
            "g": np.round(rgba_colors[:, 1] * 255),
            "b": np.round(rgba_colors[:, 2] * 255),
            "alpha": rgba_colors[:, 3],
        }
    )

    # Ensure BGR columns are integer type
    color_df[["b", "g", "r"]] = color_df[["b", "g", "r"]].astype(int)

    # Reorder to B, G, R, Alpha if that's preferred, otherwise keep R, G, B, Alpha
    # Standard often expects BGR for certain applications, but R,G,B is also common.
    # Sticking to R,G,B,Alpha as generated, can reorder if needed:
    # color_df = color_df[['b', 'g', 'r', 'alpha']]

    return color_df


def _convert_to_pysigmacro_format(df_heatmap):

    df = df_heatmap.copy()

    df["symbol"] = pd.DataFrame(df["z"].copy())

    df = df.sort_values("y_fill", ascending=False).copy()


    # For Graph Wizard Parameters
    i_plot = 0

    # Area Plots
    cols_area = ["x_lower_fill", "x_upper_fill", "y_fill", "z", "bgra_fake"]
    df_areas = df[cols_area].copy()
    df_bgra = _apply_colormap_to_values(
        df_areas["z"], vmin=0, vmax=df_areas["z"].max(), cmap_name="Blues"
    )
    df_areas = df_areas.drop(columns="z")
    df_areas = pd.concat([df_areas, df_bgra], axis=1)
    df = df.drop(columns=cols_area)

    df_areas_agg = []
    for _ii, row in df_areas.iterrows():
        xl = row.x_lower_fill
        xu = row.x_upper_fill
        y = row.y_fill
        bgra = [row.b, row.g, row.r, row.alpha]

        if np.isnan(np.array([xl, xu, y])).all():
            continue

        _df_area_dict = dict(x=[xl, xu], y=[y, y], bgra=bgra)
        _df_gw_area = create_graph_wizard_params("area_heatmap", i_plot)
        i_plot += 1
        _df_area = create_padded_df(_df_gw_area, _df_area_dict)
        df_areas_agg.append(_df_area)
    df_areas = pd.concat(df_areas_agg, axis=1)

    # Scatter
    cols_scatter = ["x", "y", "symbol"]
    df_scatter = df[cols_scatter]
    df_gw_scatter = create_graph_wizard_params("scatter_heatmap", i_plot)
    i_plot += 1
    df_scatter = create_padded_df(df_gw_scatter, df_scatter)
    # df_scatter["symbol"] = df_scatter["symbol"].astype(str)
    # df_scatter["bgra_fake"] = "NONE_STR"
    df_scatter.loc[:, "bgra_fake"] = "NONE_STR"
    df = df.drop(columns=cols_scatter)

    # Colorbar
    cols_cbar = ["ycbar", "xcbar", "b", "g", "r", "alpha"]
    df_cbar = df[cols_cbar]
    df_cbar["bgra_fake"] = "NONE_STR"
    df_gw_barh = create_graph_wizard_params("barh_heatmap", i_plot)
    i_plot += 1
    df_cbar = create_padded_df(df_gw_barh, df_scatter)
    df = df.drop(columns=cols_cbar)

    return create_padded_df(df_areas, df_scatter), df_cbar


if __name__ == "__main__":
    import pysigmacro as psm

    # Heatmap
    df = psm.demo._gen_data_heatmap(0, n_cols=4, n_rows=3, cbar=False)

# EOF