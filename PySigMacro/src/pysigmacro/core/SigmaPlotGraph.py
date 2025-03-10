#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 10:16:15 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/core/SigmaPlotGraph.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/core/SigmaPlotGraph.py"

"""
Graph creation and manipulation utilities for SigmaPlot.
"""
import os
from typing import List, Dict, Tuple, Optional, Union
from pysigmacro.core.connection import connect
from pysigmacro.core.SigmaPlotWorksheet import SigmaPlotWorksheet

class SigmaPlotGraph:
    """Class to handle SigmaPlot graph operations"""
    # Available graph types
    GRAPH_TYPES = [
        "Line Plot", "Scatter Plot", "Bar Chart", "Pie Chart",
        "Area Plot", "Histogram", "Box Plot", "3D Scatter Plot",
        "3D Line Plot", "Contour Plot", "Polar Plot"
    ]

    def __init__(self, sigmaplot=None, visible=True):
        """
        Initialize a SigmaPlotGraph instance.
        Args:
            sigmaplot: Optional existing SigmaPlot application object
            visible (bool): Make SigmaPlot visible
        """
        self.app = sigmaplot if sigmaplot else connect(visible=visible)
        if not self.app:
            raise ConnectionError("Failed to connect to SigmaPlot")

    def create_graph(
        self,
        csv_path: str = None,
        output_path: str = None,
        graph_type: str = "Line Plot",
        close_app: bool = False
    ) -> bool:
        """
        Create a graph from a CSV file and export it.
        Args:
            csv_path (str): Path to the CSV file with data
            output_path (str): Path where to save the output graph
            graph_type (str): Type of graph to create (default: "Line Plot")
            close_app (bool): Whether to close SigmaPlot after creating the graph
        Returns:
            bool: True if graph was created successfully, False otherwise
        """
        try:
            # Import CSV if path is provided
            if csv_path:
                # Create a new worksheet
                self.app.NewWorksheet()
                # Import CSV data
                self.app.CurrentWorksheet.ImportData(csv_path, 1, 1, ",")
            # Create a new graph
            self.app.NewGraph(graph_type)
            # Set data for the graph
            self.app.CurrentGraph.SetData(self.app.CurrentWorksheet, 1, 2)
            # Export the graph
            if output_path:
                self.app.CurrentGraph.ExportGraph(output_path, "PNG")
            # Close if requested
            if close_app:
                self.app.Quit()
                self.app = None
            return True
        except Exception as e:
            print(f"Error creating graph: {e}")
            return False

    def create_graph_from_data(
        self,
        x_values: List,
        y_values: List,
        output_path: str = None,
        graph_type: str = "Line Plot",
        title: Optional[str] = None,
        x_label: Optional[str] = None,
        y_label: Optional[str] = None,
        close_app: bool = False,
        export_format: str = "PNG"
    ) -> bool:
        """
        Create a graph from direct data values and export it.
        Args:
            x_values (List): X-axis values
            y_values (List): Y-axis values
            output_path (str): Path where to save the output graph
            graph_type (str): Type of graph to create
            title (str, optional): Graph title
            x_label (str, optional): X-axis label
            y_label (str, optional): Y-axis label
            close_app (bool): Whether to close SigmaPlot after creating
            export_format (str): Export format ("PNG", "TIFF", "JPEG", etc.)
        Returns:
            bool: True if graph was created successfully, False otherwise
        """
        try:
            # Set data in worksheet
            self.app = SigmaPlotWorksheet.set_data(self.app, x_values, y_values)
            if not self.app:
                return False
            # Create a new graph
            self.app.NewGraph(graph_type)
            # Set data for the graph
            self.app.CurrentGraph.SetData(self.app.CurrentWorksheet, 1, 2)
            # Customize graph if needed
            if title or x_label or y_label:
                self.customize_graph(
                    self.app.CurrentGraph,
                    title=title,
                    x_label=x_label,
                    y_label=y_label
                )
            # Export the graph
            if output_path:
                self.app.CurrentGraph.ExportGraph(output_path, export_format)
            # Close if requested
            if close_app:
                self.app.Quit()
                self.app = None
            return True
        except Exception as e:
            print(f"Error creating graph from data: {e}")
            return False

    def create_multi_series_graph(
        self,
        data_series: Dict[str, Tuple[List, List]],
        output_path: str = None,
        graph_type: str = "Line Plot",
        title: Optional[str] = None,
        x_label: Optional[str] = None,
        y_label: Optional[str] = None,
        close_app: bool = False
    ) -> bool:
        """
        Create a graph with multiple data series.
        Args:
            data_series (Dict[str, Tuple[List, List]]): Dictionary mapping series names to (x, y) data
            output_path (str): Path where to save the output graph
            graph_type (str): Type of graph to create
            title (str, optional): Graph title
            x_label (str, optional): X-axis label
            y_label (str, optional): Y-axis label
            close_app (bool): Whether to close SigmaPlot after creating
        Returns:
            bool: True if graph was created successfully, False otherwise
        """
        try:
            # Create a new worksheet
            if not hasattr(self.app, 'CurrentWorksheet'):
                self.app.NewWorksheet()
            # Prepare data for multiple series
            column = 1
            series_columns = {}
            # Fill worksheet with all series data
            for series_name, (x_values, y_values) in data_series.items():
                # Set header for X column
                try:
                    self.app.CurrentWorksheet.Cells(0, column).Value = f"{series_name} X"
                except:
                    pass
                # Set X values
                for i, x in enumerate(x_values):
                    try:
                        self.app.CurrentWorksheet.Cells(i+1, column).Value = x
                    except:
                        try:
                            self.app.CurrentWorksheet.SetCell(i+1, column, x)
                        except:
                            pass
                # Store X column
                x_col = column
                column += 1
                # Set header for Y column
                try:
                    self.app.CurrentWorksheet.Cells(0, column).Value = f"{series_name} Y"
                except:
                    pass
                # Set Y values
                for i, y in enumerate(y_values):
                    try:
                        self.app.CurrentWorksheet.Cells(i+1, column).Value = y
                    except:
                        try:
                            self.app.CurrentWorksheet.SetCell(i+1, column, y)
                        except:
                            pass
                # Store Y column
                y_col = column
                column += 1
                # Save column pair for this series
                series_columns[series_name] = (x_col, y_col)
            # Create a new graph
            self.app.NewGraph(graph_type)
            # Add data for each series
            for series_name, (x_col, y_col) in series_columns.items():
                # For first series, use SetData to initialize
                if series_name == list(series_columns.keys())[0]:
                    self.app.CurrentGraph.SetData(self.app.CurrentWorksheet, x_col, y_col)
                else:
                    # For additional series, use AddData
                    try:
                        self.app.CurrentGraph.AddData(self.app.CurrentWorksheet, x_col, y_col)
                    except:
                        # If AddData doesn't work, try alternative methods
                        try:
                            self.app.CurrentGraph.AddSeries(self.app.CurrentWorksheet, x_col, y_col)
                        except Exception as e:
                            print(f"Warning: Could not add series {series_name}: {e}")
            # Customize graph if needed
            if title or x_label or y_label:
                self.customize_graph(
                    self.app.CurrentGraph,
                    title=title,
                    x_label=x_label,
                    y_label=y_label
                )
            # Export the graph
            if output_path:
                self.app.CurrentGraph.ExportGraph(output_path, "PNG")
            # Close if requested
            if close_app:
                self.app.Quit()
                self.app = None
            return True
        except Exception as e:
            print(f"Error creating multi-series graph: {e}")
            return False

    def customize_graph(
        self,
        graph=None,
        title: Optional[str] = None,
        x_label: Optional[str] = None,
        y_label: Optional[str] = None,
        line_color: Optional[str] = None,
        line_width: Optional[float] = None,
        legend: Optional[bool] = None,
        grid: Optional[bool] = None
    ) -> bool:
        """
        Customize a SigmaPlot graph.
        Args:
            graph: SigmaPlot graph object (uses current graph if None)
            title (str, optional): Graph title
            x_label (str, optional): X-axis label
            y_label (str, optional): Y-axis label
            line_color (str, optional): Line color
            line_width (float, optional): Line width
            legend (bool, optional): Show/hide legend
            grid (bool, optional): Show/hide grid
        Returns:
            bool: True if graph was customized successfully, False otherwise
        """
        try:
            if graph is None:
                graph = self.app.CurrentGraph
            if title:
                graph.Title = title
            if x_label:
                graph.XAxis.Label = x_label
            if y_label:
                graph.YAxis.Label = y_label
            if line_color or line_width:
                # Implementation depends on SigmaPlot's API
                try:
                    if hasattr(graph, 'PlotSeries'):
                        if line_color:
                            graph.PlotSeries(0).LineColor = line_color
                        if line_width:
                            graph.PlotSeries(0).LineWidth = line_width
                except Exception as e:
                    print(f"Warning: Could not set line properties: {e}")
            if legend is not None:
                try:
                    graph.Legend.Visible = legend
                except Exception as e:
                    print(f"Warning: Could not set legend visibility: {e}")
            if grid is not None:
                try:
                    graph.XAxis.Grid.Visible = grid
                    graph.YAxis.Grid.Visible = grid
                except Exception as e:
                    print(f"Warning: Could not set grid visibility: {e}")
            return True
        except Exception as e:
            print(f"Error customizing graph: {e}")
            return False

    def close(self):
        """Close the SigmaPlot application"""
        if self.app:
            try:
                self.app.Quit()
                self.app = None
                return True
            except Exception as e:
                print(f"Error closing SigmaPlot: {e}")
                return False
        return False

    # Keep static methods for backward compatibility
    @staticmethod
    def create_graph(
        sigmaplot=None,
        csv_path: str = None,
        output_path: str = None,
        graph_type: str = "Line Plot",
        close_app: bool = True
    ) -> bool:
        """
        Create a graph from a CSV file and export it (static method).
        Args:
            sigmaplot: Optional existing SigmaPlot application object
            csv_path (str): Path to the CSV file with data
            output_path (str): Path where to save the output graph
            graph_type (str): Type of graph to create (default: "Line Plot")
            close_app (bool): Whether to close SigmaPlot after creating the graph
        Returns:
            bool: True if graph was created successfully, False otherwise
        """
        try:
            # Use provided sigmaplot instance or create a new one
            app = sigmaplot if sigmaplot else SigmaPlotWorksheet.import_csv(csv_path)
            if not app:
                return False
            # If sigmaplot was provided but we still need to import CSV
            if sigmaplot and csv_path:
                # Create a new worksheet if needed
                app.NewWorksheet()
                # Import CSV data
                app.CurrentWorksheet.ImportData(csv_path, 1, 1, ",")
            # Create a new graph
            app.NewGraph(graph_type)
            # Set data for the graph
            app.CurrentGraph.SetData(app.CurrentWorksheet, 1, 2)
            # Export the graph
            app.CurrentGraph.ExportGraph(output_path, "PNG")
            # Close if requested (only if we created the instance)
            if close_app and not sigmaplot:
                app.Quit()
                del app
            return True
        except Exception as e:
            print(f"Error creating graph: {e}")
            return False

# EOF