#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-08 23:38:39 (ywatanabe)"
# File: /home/ywatanabe/proj/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/examples/basic_plots.py

THIS_FILE = "/home/ywatanabe/proj/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/examples/basic_plots.py"

"""
Basic example plots for SigmaPlot.
"""

import os
import tempfile
from pysigmacro.core.graph import SigmaPlotGraph
from pysigmacro.data.generators import (
    create_sample_data,
    create_scatter_data,
    create_time_series_data,
    create_categorical_data
)

def sine_wave_example():
    """Create a basic sine wave plot."""
    # Generate sample sine wave data
    csv_path, x_values, y_values = create_sample_data("sine", num_points=100)

    # Create output path
    output_path = os.path.join(tempfile.gettempdir(), "sigmaplot_sine_example.png")

    # Create the graph
    success = SigmaPlotGraph.create_graph_from_data(
        x_values,
        y_values,
        output_path,
        title="Sine Wave Example",
        x_label="X (radians)",
        y_label="sin(x)"
    )

    if success:
        print(f"Sine wave example created at: {output_path}")
        return output_path
    else:
        print("Failed to create sine wave example")
        return None

def scatter_plot_example():
    """Create a scatter plot example."""
    # Generate scatter data
    csv_path, x_values, y_values = create_scatter_data(
        num_points=50,
        correlation=0.7
    )

    # Create output path
    output_path = os.path.join(tempfile.gettempdir(), "sigmaplot_scatter_example.png")

    # Create the graph
    success = SigmaPlotGraph.create_graph_from_data(
        x_values,
        y_values,
        output_path,
        graph_type="Scatter Plot",
        title="Scatter Plot Example",
        x_label="X Values",
        y_label="Y Values"
    )

    if success:
        print(f"Scatter plot example created at: {output_path}")
        return output_path
    else:
        print("Failed to create scatter plot example")
        return None

def bar_chart_example():
    """Create a bar chart example."""
    # Generate categorical data
    categories = ["Group A", "Group B", "Group C", "Group D", "Group E"]
    values = [4.2, 3.8, 5.7, 2.1, 3.5]

    csv_path, categories, values = create_categorical_data(categories, values)

    # Create output path
    output_path = os.path.join(tempfile.gettempdir(), "sigmaplot_bar_example.png")

    # Create the graph (using CSV import since bar charts need special handling)
    success = SigmaPlotGraph.create_graph(
        csv_path,
        output_path,
        graph_type="Bar Chart"
    )

    if success:
        print(f"Bar chart example created at: {output_path}")
        return output_path
    else:
        print("Failed to create bar chart example")
        return None

def time_series_example():
    """Create a time series example."""
    # Generate time series data
    csv_path, dates, values = create_time_series_data(
        start_date="2023-01-01",
        days=30,
        trend=0.1,
        seasonality=0.5
    )

    # Create output path
    output_path = os.path.join(tempfile.gettempdir(), "sigmaplot_time_series_example.png")

    # Create the graph (using CSV import since time series need special handling)
    success = SigmaPlotGraph.create_graph(
        csv_path,
        output_path,
        graph_type="Line Plot"
    )

    if success:
        print(f"Time series example created at: {output_path}")
        return output_path
    else:
        print("Failed to create time series example")
        return None

def multi_series_example():
    """Create a graph with multiple data series."""
    # Generate different data series
    sine_path, sine_x, sine_y = create_sample_data("sine", num_points=50)
    cos_path, cos_x, cos_y = create_sample_data("cosine", num_points=50)

    # Create data series dictionary
    data_series = {
        "Sine": (sine_x, sine_y),
        "Cosine": (cos_x, cos_y)
    }

    # Create output path
    output_path = os.path.join(tempfile.gettempdir(), "sigmaplot_multi_series_example.png")

    # Create the graph
    success = SigmaPlotGraph.create_multi_series_graph(
        data_series,
        output_path,
        title="Sine and Cosine Functions",
        x_label="X (radians)",
        y_label="Value"
    )

    if success:
        print(f"Multi-series example created at: {output_path}")
        return output_path
    else:
        print("Failed to create multi-series example")
        return None

def run_all_examples():
    """
    Run all basic plot examples.

    Returns:
        dict: Dictionary of example names and their output paths
    """
    results = {}

    print("\nRunning sine wave example...")
    results["sine_wave"] = sine_wave_example()

    print("\nRunning scatter plot example...")
    results["scatter_plot"] = scatter_plot_example()

    print("\nRunning bar chart example...")
    results["bar_chart"] = bar_chart_example()

    print("\nRunning time series example...")
    results["time_series"] = time_series_example()

    print("\nRunning multi-series example...")
    results["multi_series"] = multi_series_example()

    # Print summary
    print("\n=== Example Results ===")
    for name, path in results.items():
        status = "Success" if path else "Failed"
        print(f"{name}: {status}")

    return results

def main():
    """Main function to run when executed as script."""
    print("Pysigmacro Basic Examples")
    run_all_examples()

if __name__ == "__main__":
    main()

# EOF