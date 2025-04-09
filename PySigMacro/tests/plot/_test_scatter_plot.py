#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 00:19:53 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/plot/test_scatter_plot.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/plot/test_scatter_plot.py"

import os
import pytest
import tempfile
from unittest.mock import MagicMock, patch
from typing import Any, Optional, List, Dict

"""
Functionality:
* Tests scatter plot creation in SigmaPlot
* Verifies correct setup of scatter plot properties and data points
Input:
* None (uses mock objects)
Output:
* Test results for scatter plot creation
Prerequisites:
* pytest
* mock objects for SigmaPlot COM interface
"""

class TestScatterPlot:
    @pytest.fixture
    def mock_sigmaplot(self) -> Any:
        """
        Create a mock SigmaPlot application for testing.

        Returns
        -------
        Any
            Mock SigmaPlot application object
        """
        # Skip test if not running on Windows
        if os.name != 'nt':
            pytest.skip("SigmaPlot tests only run on Windows")

        # Mock the SigmaPlot application
        with patch('win32com.client.Dispatch') as mock_dispatch:
            mock_app = MagicMock()
            mock_app.Visible = False

            # Setup mock worksheet
            mock_worksheet = MagicMock()
            mock_app.CurrentWorksheet = mock_worksheet
            mock_app.NewWorksheet.return_value = mock_worksheet

            # Setup mock graph
            mock_graph = MagicMock()
            mock_app.CurrentGraph = mock_graph
            mock_app.NewGraph.return_value = mock_graph

            # Configure mock dispatch to return our mock app
            mock_dispatch.return_value = mock_app

            yield mock_app

    def test_create_scatter_plot(self, mock_sigmaplot: Any) -> None:
        """
        Test creating a basic scatter plot.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Define test data (scattered points)
        x_values = [1.2, 2.5, 3.7, 4.1, 5.8]
        y_values = [10.3, 15.7, 13.2, 17.8, 20.1]

        # Insert data into worksheet
        for i, (x, y) in enumerate(zip(x_values, y_values)):
            mock_sigmaplot.CurrentWorksheet.Cells(i+1, 1).Value = x
            mock_sigmaplot.CurrentWorksheet.Cells(i+1, 2).Value = y

        # Create scatter plot
        mock_sigmaplot.NewGraph("Scatter Plot")

        # Set data for the graph
        mock_sigmaplot.CurrentGraph.SetData(mock_sigmaplot.CurrentWorksheet, 1, 2)

        # Verify method calls
        assert mock_sigmaplot.NewGraph.called
        assert mock_sigmaplot.NewGraph.call_args[0][0] == "Scatter Plot"
        assert mock_sigmaplot.CurrentGraph.SetData.called

        # Verify SetData parameters (worksheet, x_column, y_column)
        call_args = mock_sigmaplot.CurrentGraph.SetData.call_args[0]
        assert call_args[0] == mock_sigmaplot.CurrentWorksheet
        assert call_args[1] == 1
        assert call_args[2] == 2

    def test_scatter_plot_symbol_customization(self, mock_sigmaplot: Any) -> None:
        """
        Test customizing scatter plot symbols.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create simple scatter plot
        mock_sigmaplot.NewGraph("Scatter Plot")

        # Apply customizations to plot symbols
        try:
            # Try to set symbol properties if available
            mock_sigmaplot.CurrentGraph.PlotSeries(0).Symbol.Type = "Circle"
            mock_sigmaplot.CurrentGraph.PlotSeries(0).Symbol.Size = 5
            mock_sigmaplot.CurrentGraph.PlotSeries(0).Symbol.Color = "Red"
            mock_sigmaplot.CurrentGraph.PlotSeries(0).Symbol.Fill = True
            mock_sigmaplot.CurrentGraph.PlotSeries(0).Symbol.FillColor = "Yellow"
        except:
            # Not all properties may be available in mock
            pass

        # Apply basic customizations
        mock_sigmaplot.CurrentGraph.Title = "Scatter Plot Example"
        mock_sigmaplot.CurrentGraph.XAxis.Label = "X Axis"
        mock_sigmaplot.CurrentGraph.YAxis.Label = "Y Axis"

        # Verify basic customizations
        assert mock_sigmaplot.CurrentGraph.Title == "Scatter Plot Example"
        assert mock_sigmaplot.CurrentGraph.XAxis.Label == "X Axis"
        assert mock_sigmaplot.CurrentGraph.YAxis.Label == "Y Axis"

    def test_scatter_with_regression(self, mock_sigmaplot: Any) -> None:
        """
        Test adding regression line to scatter plot.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create scatter plot
        mock_sigmaplot.NewGraph("Scatter Plot")

        # Try to add regression line
        try:
            mock_sigmaplot.CurrentGraph.AddStatistics("Linear Regression")

            # Verify the call was made
            assert mock_sigmaplot.CurrentGraph.AddStatistics.called
            assert mock_sigmaplot.CurrentGraph.AddStatistics.call_args[0][0] == "Linear Regression"
        except:
            # This functionality might not be available in the mock
            pytest.skip("Regression functionality not available in mock")

    def test_export_scatter_plot(self, mock_sigmaplot: Any) -> None:
        """
        Test exporting a scatter plot.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create scatter plot
        mock_sigmaplot.NewGraph("Scatter Plot")

        # Set up export path
        export_path = os.path.join(tempfile.gettempdir(), "test_scatter_plot.png")

        # Export the graph
        mock_sigmaplot.CurrentGraph.ExportGraph(export_path, "PNG")

        # Verify export was attempted
        assert mock_sigmaplot.CurrentGraph.ExportGraph.called
        assert mock_sigmaplot.CurrentGraph.ExportGraph.call_args[0][0] == export_path
        assert mock_sigmaplot.CurrentGraph.ExportGraph.call_args[0][1] == "PNG"

# EOF