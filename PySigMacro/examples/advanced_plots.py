#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-08 23:15:37 (ywatanabe)"
# File: /home/ywatanabe/proj/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/examples/advanced_plots.py

THIS_FILE = "/home/ywatanabe/proj/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/examples/advanced_plots.py"

"""
Advanced example plots for SigmaPlot.
"""

import os
import tempfile
import math
from pysigmacro.core.graph import SigmaPlotGraph
from pysigmacro.core.connection import connect
from pysigmacro.core.worksheet import SigmaPlotWorksheet
from pysigmacro.data.generators import (
    create_sample_data,
    create_3d_data,
    prepare_multi_series_data
)

def contour_plot_example():
    """Create a contour plot example."""
    # Generate 3D data
    csv_path, x_grid, y_grid, z_values = create_3d_data(
        nx=20,
        ny=20,
        function_type="peaks"
    )

    # Create output path
    output_path = os.path.join(tempfile.gettempdir(), "sigmaplot_contour_example.png")

    # Create the graph (using CSV import since contour plots need special handling)
    success = SigmaPlotGraph.create_graph(
        csv_path,
        output_path,
        graph_type="Contour Plot"
    )

    if success:
        print(f"Contour plot example created at: {output_path}")
        return output_path
    else:
        print("Failed to create contour plot example")
        return None

def surface_3d_example():
    """Create a 3D surface plot example."""
    # Generate 3D data
    csv_path, x_grid, y_grid, z_values = create_3d_data(
        nx=20,
        ny=20,
        function_type="sincos"
    )

    # Create output path
    output_path = os.path.join(tempfile.gettempdir(), "sigmaplot_surface_3d_example.png")

    # Create the graph (using CSV import since 3D plots need special handling)
    success = SigmaPlotGraph.create_graph(
        csv_path,
        output_path,
        graph_type="3D Surface Plot"
    )

    if success:
        print(f"3D surface plot example created at: {output_path}")
        return output_path
    else:
        print("Failed to create 3D surface plot example")
        return None

def error_bars_example():
    """Create a plot with error bars."""
    try:
        # Connect to SigmaPlot
        app = connect(visible=True)
        if not app:
            print("Failed to connect to SigmaPlot")
            return None

        # Create a new worksheet
        app.NewWorksheet()

        # Generate sample data with error values
        x_values = [1, 2, 3, 4, 5]
        y_values = [2.1, 3.2, 2.8, 4.5, 5.2]
        y_error = [0.3, 0.4, 0.2, 0.5, 0.3]

        # Set X values (column 1)
        for i, x in enumerate(x_values):
            app.CurrentWorksheet.Cells(i+1, 1).Value = x

        # Set Y values (column 2)
        for i, y in enumerate(y_values):
            app.CurrentWorksheet.Cells(i+1, 2).Value = y

        # Set error values (column 3)
        for i, err in enumerate(y_error):
            app.CurrentWorksheet.Cells(i+1, 3).Value = err

        # Create a new graph
        app.NewGraph("Line Plot")

        # Set data for the graph (with error bars)
        app.CurrentGraph.SetData(app.CurrentWorksheet, 1, 2)

        # Set error bars (if available in API)
        try:
            app.CurrentGraph.ErrorBars.Column = 3
            app.CurrentGraph.ErrorBars.Visible = True
        except:
            print("Warning: Could not set error bars through API")

        # Customize graph
        app.CurrentGraph.Title = "Data with Error Bars"
        app.CurrentGraph.XAxis.Label = "X Values"
        app.CurrentGraph.YAxis.Label = "Y Values"

        # Create output path
        output_path = os.path.join(tempfile.gettempdir(), "sigmaplot_error_bars_example.png")

        # Export the graph
        app.CurrentGraph.ExportGraph(output_path, "PNG")

        # Close the application
        app.Quit()
        del app

        print(f"Error bars example created at: {output_path}")
        return output_path

    except Exception as e:
        print(f"Error creating error bars example: {e}")
        return None

def polar_plot_example():
    """Create a polar plot example."""
    try:
        # Connect to SigmaPlot
        app = connect(visible=True)
        if not app:
            print("Failed to connect to SigmaPlot")
            return None

        # Create a new worksheet
        app.NewWorksheet()

        # Generate polar data (theta, r)
        theta_values = [i * 2 * math.pi / 36 for i in range(37)]
        r_values = [1 + 0.5 * math.sin(5 * theta) for theta in theta_values]

        # Set theta values (column 1)
        for i, theta in enumerate(theta_values):
            app.CurrentWorksheet.Cells(i+1, 1).Value = theta

        # Set r values (column 2)
        for i, r in enumerate(r_values):
            app.CurrentWorksheet.Cells(i+1, 2).Value = r

        # Create a new polar graph
        app.NewGraph("Polar Plot")

        # Set data for the graph
        app.CurrentGraph.SetData(app.CurrentWorksheet, 1, 2)

        # Customize graph
        app.CurrentGraph.Title = "Polar Plot Example"

        # Create output path
        output_path = os.path.join(tempfile.gettempdir(), "sigmaplot_polar_example.png")

        # Export the graph
        app.CurrentGraph.ExportGraph(output_path, "PNG")

        # Close the application
        app.Quit()
        del app

        print(f"Polar plot example created at: {output_path}")
        return output_path

    except Exception as e:
        print(f"Error creating polar plot example: {e}")
        return None

def regression_example():
    """Create a plot with regression line."""
    try:
        # Generate scatter data with linear trend
        x_values = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
        y_values = [i * 1.5 + 2 + (random.random() - 0.5) * 2 for i in x_values]

        # Connect to SigmaPlot
        app = connect(visible=True)
        if not app:
            print("Failed to connect to SigmaPlot")
            return None

        # Set data in worksheet
        app = SigmaPlotWorksheet.set_data(app, x_values, y_values)
        if not app:
            return None

        # Create a new scatter plot
        app.NewGraph("Scatter Plot")

        # Set data for the graph
        app.CurrentGraph.SetData(app.CurrentWorksheet, 1, 2)

        # Add regression line (if available in API)
        try:
            app.CurrentGraph.AddStatistics("Linear Regression")
        except:
            print("Warning: Could not add regression line through API")

        # Customize graph
        app.CurrentGraph.Title = "Scatter Plot with Regression Line"
        app.CurrentGraph.XAxis.Label = "X Values"
        app.CurrentGraph.YAxis.Label = "Y Values"

        # Create output path
        output_path = os.path.join(tempfile.gettempdir(), "sigmaplot_regression_example.png")

        # Export the graph
        app.CurrentGraph.ExportGraph(output_path, "PNG")

        # Close the application
        app.Quit()
        del app

        print(f"Regression example created at: {output_path}")
        return output_path

    except Exception as e:
        print(f"Error creating regression example: {e}")
        return None

def box_plot_example():
    """Create a box plot example."""
    try:
        import random

        # Connect to SigmaPlot
        app = connect(visible=True)
        if not app:
            print("Failed to connect to SigmaPlot")
            return None

        # Create a new worksheet
        app.NewWorksheet()

        # Generate data for 3 groups
        groups = ["Group A", "Group B", "Group C"]
        data = {
            "Group A": [random.normalvariate(10, 2) for _ in range(20)],
            "Group B": [random.normalvariate(12, 1.5) for _ in range(20)],
            "Group C": [random.normalvariate(9, 2.5) for _ in range(20)]
        }

        # Set column headers
        for i, group in enumerate(groups):
            app.CurrentWorksheet.Cells(0, i+1).Value = group

        # Set data values for each group
        for col, group in enumerate(groups):
            for row, value in enumerate(data[group]):
                app.CurrentWorksheet.Cells(row+1, col+1).Value = value

        # Create a new box plot
        app.NewGraph("Box Plot")

        # Set data for the graph
        app.CurrentGraph.SetDataMultiColumn(app.CurrentWorksheet, 1, len(groups))

        # Customize graph
        app.CurrentGraph.Title = "Box Plot Example"
        app.CurrentGraph.YAxis.Label = "Values"

        # Create output path
        output_path = os.path.join(tempfile.gettempdir(), "sigmaplot_box_plot_example.png")

        # Export the graph
        app.CurrentGraph.ExportGraph(output_path, "PNG")

        # Close the application
        app.Quit()
        del app

        print(f"Box plot example created at: {output_path}")
        return output_path

    except Exception as e:
        print(f"Error creating box plot example: {e}")
        return None

def run_all_examples():
    """
    Run all advanced plot examples.

    Returns:
        dict: Dictionary of example names and their output paths
    """
    import random

    results = {}

    print("\nRunning contour plot example...")
    results["contour_plot"] = contour_plot_example()

    print("\nRunning 3D surface example...")
    results["surface_3d"] = surface_3d_example()

    print("\nRunning error bars example...")
    results["error_bars"] = error_bars_example()

    print("\nRunning polar plot example...")
    results["polar_plot"] = polar_plot_example()

    print("\nRunning regression example...")
    results["regression"] = regression_example()

    print("\nRunning box plot example...")
    results["box_plot"] = box_plot_example()

    # Print summary
    print("\n=== Example Results ===")
    for name, path in results.items():
        status = "Success" if path else "Failed"
        print(f"{name}: {status}")

    return results

def main():
    """Main function to run when executed as script."""
    print("Pysigmacro Advanced Examples")
    run_all_examples()

if __name__ == "__main__":
    main()

# EOF