#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-09 13:57:57 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/data/_create_graph_wizard_params.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/data/_create_graph_wizard_params.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import numpy as np
import pandas as pd
from ._create_padded_df import create_padded_df
from ..const._PLOT_PARAMS_BASE import PLOT_PARAMS_BASE
from ..const import BGRA, COLORS


# Main
# ------------------------------


def create_graph_wizard_params(plot_type, i_plot, label=None, **kwargs):

    # Plottype specific Parameters
    create_graph_wizard_func = {
        "bar": _create_graph_wizard_params_bar,
        "barh": _create_graph_wizard_params_barh,
        "barh_heatmap": _create_graph_wizard_params_barh,
        "area": _create_graph_wizard_params_area,
        "area_heatmap": _create_graph_wizard_params_area,
        "areah": _create_graph_wizard_params_areah,
        "box": _create_graph_wizard_params_box,
        "box_violin": _create_graph_wizard_params_box,
        "boxh": _create_graph_wizard_params_boxh,
        "boxh_violin": _create_graph_wizard_params_boxh,
        "line": _create_graph_wizard_params_line,
        "filled_line_uu": _create_graph_wizard_params_line,
        "filled_line_mm": _create_graph_wizard_params_line,
        "filled_line_ll": _create_graph_wizard_params_line,
        "line_yerr": _create_graph_wizard_params_line_yerr,
        "lines_x_many_y": _create_graph_wizard_params_lines_x_many_y,
        "lines_y_many_x": _create_graph_wizard_params_lines_y_many_x,
        "lines_y_many_x_violin": _create_graph_wizard_params_lines_y_many_x,
        "lines_x_many_y_violinh": _create_graph_wizard_params_lines_x_many_y,
        "polar": _create_graph_wizard_params_polar,
        "scatter": _create_graph_wizard_params_scatter,
        "scatter_heatmap": _create_graph_wizard_params_scatter,
        "jitter": _create_graph_wizard_params_scatter,
        "violin": _create_graph_wizard_params_violin,
        "violinh": _create_graph_wizard_params_violinh,
        "filled_line": _create_graph_wizard_params_filled_line,
        "contour": _create_graph_wizard_params_contour,
        "heatmap": _create_graph_wizard_params_heatmap,
        "histogram": _create_graph_wizard_params_histogram,
    }[plot_type]

    graph_wizard_params = create_graph_wizard_func()

    # Reformat
    gw_df = pd.DataFrame(pd.Series(graph_wizard_params)).reset_index()
    gw_df.columns = ["gw_param_keys", "gw_param_values"]

    # Label: the third column
    label = label if label else plot_type
    gw_df.loc[:, "label"] = "NONE_STR"
    gw_df.loc[0, "label"] = label

    gw_df = create_padded_df(gw_df)

    gw_df.columns = [f"{col} {i_plot}"for col in gw_df.columns]

    return gw_df


# Core Functions
# ------------------------------

def _create_graph_wizard_params_bar():
    return {
        "plot_type": "Vertical Bar Chart",
        "plot_style_type": "Simple Error Bars",
        "plot_data_type": "XY Pair",
        "plot_columns_per_plot": "ColumnsPerPlot",
        "plot_plot_columns_count_array": "PlotColumnCountArray",
        "plot_data_source": "Worksheet Columns",
        "plot_polarunits": "NONE_STR",
        "plot_anguleunits": "Degrees",
        "plot_min_angle_row": 0,
        "plot_max_angle_row": 360,
        "plot_unknown1": "NONE_STR",
        "plot_group_style": "NONE_STR",
        "plot_use_automatic_legends": True,
    }


def _create_graph_wizard_params_barh():
    return {
        "plot_type": "Horizontal Bar Chart",
        "plot_style_type": "Simple Error Bars",
        "plot_data_type": "YX Pair",
        "plot_columns_per_plot": "ColumnsPerPlot",
        "plot_plot_columns_count_array": "PlotColumnCountArray",
        "plot_data_source": "Worksheet Columns",
        "plot_polarunits": "NONE_STR",
        "plot_anguleunits": "Degrees",
        "plot_min_angle_row": 0,
        "plot_max_angle_row": 360,
        "plot_unknown1": "NONE_STR",
        "plot_group_style": "NONE_STR",
        "plot_use_automatic_legends": True,
    }

def _create_graph_wizard_params_histogram():
    return {
        "plot_type": "Vertical Bar Chart",
        "plot_style_type": "Simple Bar",
        "plot_data_type": "XY Pair",
        "plot_columns_per_plot": "ColumnsPerPlot",
        "plot_plot_columns_count_array": "PlotColumnCountArray",
        "plot_data_source": "Worksheet Columns",
        "plot_polarunits": "NONE_STR",
        "plot_anguleunits": "Degrees",
        "plot_min_angle_row": 0,
        "plot_max_angle_row": 360,
        "plot_unknown1": "NONE_STR",
        "plot_group_style": "NONE_STR",
        "plot_use_automatic_legends": True,
    }

def _create_graph_wizard_params_area():
    return {
        "plot_type": "Area Plot",
        "plot_style_type": "Simple Area",
        "plot_data_type": "XY Pair",
        "plot_columns_per_plot": "ColumnsPerPlot",
        "plot_plot_columns_count_array": "PlotColumnCountArray",
        "plot_data_source": "Worksheet Columns",
        "plot_polarunits": "Standard Deviation",
        "plot_anguleunits": "Degrees",
        "plot_min_angle_row": 0,
        "plot_max_angle_row": 360,
        "plot_unknown1": "NONE_STR",
        "plot_group_style": "Standard Deviation",
        "plot_use_automatic_legends": True,
    }

def _create_graph_wizard_params_areah():
    return {
        "plot_type": "Area Plot",
        "plot_style_type": "Vertical Area",
        "plot_data_type": "YX Pair",
        "plot_columns_per_plot": "ColumnsPerPlot",
        "plot_plot_columns_count_array": "PlotColumnCountArray",
        "plot_data_source": "Worksheet Columns",
        "plot_polarunits": "Standard Deviation",
        "plot_anguleunits": "Degrees",
        "plot_min_angle_row": 0,
        "plot_max_angle_row": 360,
        "plot_unknown1": "NONE_STR",
        "plot_group_style": "Standard Deviation",
        "plot_use_automatic_legends": True,
    }


def _create_graph_wizard_params_box():
    return {
        "plot_type": "Box Plot",
        "plot_style_type": "Vertical Box Plot",
        "plot_data_type": "X Many Y",
        "plot_columns_per_plot": "ColumnsPerPlot",
        "plot_plot_columns_count_array": "PlotColumnCountArray",
        "plot_data_source": "Worksheet Columns",
        "plot_polarunits": "Standard Deviation",
        "plot_anguleunits": "Degrees",
        "plot_min_angle_row": 0,
        "plot_max_angle_row": 360,
        "plot_unknown1": "NONE_STR",
        "plot_group_style": "NONE_STR",
        "plot_use_automatic_legends": True,
    }


def _create_graph_wizard_params_boxh():
    return {
        "plot_type": "Box Plot",
        "plot_style_type": "Horizontal Box Plot",
        "plot_data_type": "Y Many X",
        "plot_columns_per_plot": "ColumnsPerPlot",
        "plot_plot_columns_count_array": "PlotColumnCountArray",
        "plot_data_source": "Worksheet Columns",
        "plot_polarunits": "Standard Deviation",
        "plot_anguleunits": "Degrees",
        "plot_min_angle_row": 0,
        "plot_max_angle_row": 360,
        "plot_unknown1": "NONE_STR",
        "plot_group_style": "NONE_STR",
        "plot_use_automatic_legends": True,
    }

def _create_graph_wizard_params_line():
    return {
        "plot_type": "Line Plot",
        "plot_style_type": "Simple Straight Line",
        "plot_data_type": "XY Pair",
        "plot_columns_per_plot": "ColumnsPerPlot",
        "plot_plot_columns_count_array": "PlotColumnCountArray",
        "plot_data_source": "Worksheet Columns",
        "plot_polarunits": "NONE_STR",
        "plot_anguleunits": "Degrees",
        "plot_min_angle_row": 0,
        "plot_max_angle_row": 360,
        "plot_unknown1": "NONE_STR",
        "plot_group_style": "NONE_STR",
        "plot_use_automatic_legends": True,
    }

def _create_graph_wizard_params_line_yerr():
    return {
        "plot_type": "Line and Scatter Plot",
        "plot_style_type": "Simple Error Bars",
        "plot_data_type": "XY Pair",
        "plot_columns_per_plot": "ColumnsPerPlot",
        "plot_plot_columns_count_array": "PlotColumnCountArray",
        "plot_data_source": "Worksheet Columns",
        "plot_polarunits": "NONE_STR",
        "plot_anguleunits": "Degrees",
        "plot_min_angle_row": 0,
        "plot_max_angle_row": 360,
        "plot_unknown1": "NONE_STR",
        "plot_group_style": "NONE_STR",
        "plot_use_automatic_legends": True,
    }

def _create_graph_wizard_params_lines_y_many_x():
    return {
        "plot_type": "Line Plot",
        "plot_style_type": "Multiple Straight Lines",
        "plot_data_type": "Y Many X",
        "plot_columns_per_plot": "ColumnsPerPlot",
        "plot_plot_columns_count_array": "PlotColumnCountArray",
        "plot_data_source": "Worksheet Columns",
        "plot_polarunits": "NONE_STR",
        "plot_anguleunits": "Degrees",
        "plot_min_angle_row": 0,
        "plot_max_angle_row": 360,
        "plot_unknown1": "NONE_STR",
        "plot_group_style": "NONE_STR",
        "plot_use_automatic_legends": True,
    }

def _create_graph_wizard_params_lines_x_many_y():
    return {
        "plot_type": "Line Plot",
        "plot_style_type": "Multiple Straight Lines",
        "plot_data_type": "X Many Y",
        "plot_columns_per_plot": "ColumnsPerPlot",
        "plot_plot_columns_count_array": "PlotColumnCountArray",
        "plot_data_source": "Worksheet Columns",
        "plot_polarunits": "NONE_STR",
        "plot_anguleunits": "Degrees",
        "plot_min_angle_row": 0,
        "plot_max_angle_row": 360,
        "plot_unknown1": "NONE_STR",
        "plot_group_style": "NONE_STR",
        "plot_use_automatic_legends": True,
    }

def _create_graph_wizard_params_scatter():
    return {
        "plot_type": "Scatter Plot",
        "plot_style_type": "Simple Scatter",
        "plot_data_type": "XY Pair",
        "plot_columns_per_plot": "ColumnsPerPlot",
        "plot_plot_columns_count_array": "PlotColumnCountArray",
        "plot_data_source": "Worksheet Columns",
        "plot_polarunits": "Standard Deviation",
        "plot_anguleunits": "Degrees",
        "plot_min_angle_row": 0,
        "plot_max_angle_row": 360,
        "plot_unknown1": "NONE_STR",
        "plot_group_style": "NONE_STR",
        "plot_use_automatic_legends": True,
    }

def _create_graph_wizard_params_jitter():
    return {
        "plot_type": "Scatter Plot",
        "plot_style_type": "Simple Scatter",
        "plot_data_type": "XY Pair",
        "plot_columns_per_plot": "ColumnsPerPlot",
        "plot_plot_columns_count_array": "PlotColumnCountArray",
        "plot_data_source": "Worksheet Columns",
        "plot_polarunits": "Standard Deviation",
        "plot_anguleunits": "Degrees",
        "plot_min_angle_row": 0,
        "plot_max_angle_row": 360,
        "plot_unknown1": "NONE_STR",
        "plot_group_style": "NONE_STR",
        "plot_use_automatic_legends": True,
    }


def _create_graph_wizard_params_polar():
    return {
        "plot_type": "Polar Plot",
        "plot_style_type": "Lines",
        "plot_data_type": "Theta R",
        "plot_columns_per_plot": "ColumnsPerPlot",
        "plot_plot_columns_count_array": "PlotColumnCountArray",
        "plot_data_source": "Worksheet Columns",
        "plot_polarunits": "Standard Deviation",
        "plot_anguleunits": "Degrees",
        "plot_min_angle_row": 0,
        "plot_max_angle_row": 360,
        "plot_unknown1": "NONE_STR",
        "plot_group_style": "Standard Deviation",
        "plot_use_automatic_legends": True,
    }


def _create_graph_wizard_params_violin():
    return {
        "plot_type": "Vertical Violin Plot",
        "plot_style_type": "NONE_STR",
        "plot_data_type": "NONE_STR",
        "plot_columns_per_plot": "NONE_STR",
        "plot_plot_columns_count_array": "NONE_STR",
        "plot_data_source": "NONE_STR",
        "plot_polarunits": "NONE_STR",
        "plot_anguleunits": "NONE_STR",
        "plot_min_angle_row": "NONE_STR",
        "plot_max_angle_row": "NONE_STR",
        "plot_unknown1": "NONE_STR",
        "plot_group_style": "NONE_STR",
        "plot_use_automatic_legends": True,
    }

def _create_graph_wizard_params_violinh():
    return {
        "plot_type": "Horizontal Violin Plot",
        "plot_style_type": "NONE_STR",
        "plot_data_type": "NONE_STR",
        "plot_columns_per_plot": "NONE_STR",
        "plot_plot_columns_count_array": "NONE_STR",
        "plot_data_source": "NONE_STR",
        "plot_polarunits": "NONE_STR",
        "plot_anguleunits": "NONE_STR",
        "plot_min_angle_row": "NONE_STR",
        "plot_max_angle_row": "NONE_STR",
        "plot_unknown1": "NONE_STR",
        "plot_group_style": "NONE_STR",
        "plot_use_automatic_legends": True,
    }

def _create_graph_wizard_params_contour():
    return {
        "plot_type": "Contour Plot",
        "plot_style_type": "Filled Contour Plot",
        "plot_data_type": "XYZ Triplet",
        "plot_columns_per_plot": "NONE_STR",
        "plot_plot_columns_count_array": "NONE_STR",
        "plot_data_source": "Worksheet Columns",
        "plot_polarunits": "Standard Deviation",
        "plot_anguleunits": "Degrees",
        "plot_min_angle_row": 0,
        "plot_max_angle_row": 360,
        "plot_unknown1": "NONE_STR",
        "plot_group_style": "Standard Deviation",
        "plot_use_automatic_legends": True,
    }

# Special
# ------------------------------

def _create_graph_wizard_params_filled_line():
    return {
        "plot_type": "Filled Line Plot",
        "plot_style_type": "NONE_STR",
        "plot_data_type": "NONE_STR",
        "plot_columns_per_plot": "NONE_STR",
        "plot_plot_columns_count_array": "NONE_STR",
        "plot_data_source": "NONE_STR",
        "plot_polarunits": "NONE_STR",
        "plot_anguleunits": "NONE_STR",
        "plot_min_angle_row": "NONE_STR",
        "plot_max_angle_row": "NONE_STR",
        "plot_unknown1": "NONE_STR",
        "plot_group_style": "NONE_STR",
        "plot_use_automatic_legends": True,
    }

def _create_graph_wizard_params_heatmap():
    return {
        "plot_type": "Heatmap",
        "plot_style_type": "NONE_STR",
        "plot_data_type": "NONE_STR",
        "plot_columns_per_plot": "NONE_STR",
        "plot_plot_columns_count_array": "NONE_STR",
        "plot_data_source": "NONE_STR",
        "plot_polarunits": "NONE_STR",
        "plot_anguleunits": "NONE_STR",
        "plot_min_angle_row": "NONE_STR",
        "plot_max_angle_row": "NONE_STR",
        "plot_unknown1": "NONE_STR",
        "plot_group_style": "NONE_STR",
        "plot_use_automatic_legends": True,
    }

# EOF