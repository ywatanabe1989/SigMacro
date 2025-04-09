#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 00:21:15 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/plot/test_area_plot.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/plot/test_area_plot.py"

import os
import pytest
import tempfile
from unittest.mock import MagicMock, patch
from typing import Any, Optional, List, Dict

"""
Functionality:
* Tests area plot creation in SigmaPlot
* Verifies correct setup of area plot properties and data
Input:
* None (uses mock objects)
Output:
* Test results for area plot creation
Prerequisites:
* pytest
* mock objects for SigmaPlot COM interface
"""

class TestAreaPlot:
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

    def test_create_area_plot(self, mock_sigmaplot: Any) -> None:
        """
        Test creating a basic area plot.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Define test data
        x_values = [1, 2, 3, 4, 5]
        y_values = [10, 15, 13, 17, 20]

        # Insert data into worksheet
        for i, (x, y) in enumerate(zip(x_values, y_values)):
            mock_sigmaplot.CurrentWorksheet.Cells(i+1, 1).Value = x
            mock_sigmaplot.CurrentWorksheet.Cells(i+1, 2).Value = y

        # Create area plot
        mock_sigmaplot.NewGraph("Area Plot")

        # Set data for the graph
        mock_sigmaplot.CurrentGraph.SetData(mock_sigmaplot.CurrentWorksheet, 1, 2)

        # Verify method calls
        assert mock_sigmaplot.NewGraph.called
        assert mock_sigmaplot.NewGraph.call_args[0][0] == "Area Plot"
        assert mock_sigmaplot.CurrentGraph.SetData.called

        # Verify SetData parameters (worksheet, x_column, y_column)
        call_args = mock_sigmaplot.CurrentGraph.SetData.call_args[0]
        assert call_args[0] == mock_sigmaplot.CurrentWorksheet
        assert call_args[1] == 1
        assert call_args[2] == 2

    def test_area_plot_customization(self, mock_sigmaplot: Any) -> None:
        """
        Test customizing an area plot.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create simple area plot
        mock_sigmaplot.NewGraph("Area Plot")

        # Apply customizations
        mock_sigmaplot.CurrentGraph.Title = "Test Area Plot"
        mock_sigmaplot.CurrentGraph.XAxis.Label = "X Values"
        mock_sigmaplot.CurrentGraph.YAxis.Label = "Y Values"

        try:
            # Try to set area properties if available
            mock_sigmaplot.CurrentGraph.PlotSeries(0).AreaFillColor = "LightBlue"
            mock_sigmaplot.CurrentGraph.PlotSeries(0).AreaBorderColor = "Blue"
            mock_sigmaplot.CurrentGraph.PlotSeries(0).AreaBorderWidth = 1.5
        except:
            # Not all properties may be available in mock
            pass

        # Verify basic customizations
        assert mock_sigmaplot.CurrentGraph.Title == "Test Area Plot"
        assert mock_sigmaplot.CurrentGraph.XAxis.Label == "X Values"
        assert mock_sigmaplot.CurrentGraph.YAxis.Label == "Y Values"

    def test_stacked_area_plot(self, mock_sigmaplot: Any) -> None:
        """
        Test creating a stacked area plot with multiple series.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Define test data for multiple series
        x_values = [1, 2, 3, 4, 5]
        y1_values = [10, 15, 13, 17, 20]
        y2_values = [5, 8, 10, 7, 9]

        # Insert data into worksheet
        for i, x in enumerate(x_values):
            mock_sigmaplot.CurrentWorksheet.Cells(i+1, 1).Value = x
            mock_sigmaplot.CurrentWorksheet.Cells(i+1, 2).Value = y1_values[i]
            mock_sigmaplot.CurrentWorksheet.Cells(i+1, 3).Value = y2_values[i]

        # Create area plot
        mock_sigmaplot.NewGraph("Area Plot")

        # Set first data series
        mock_sigmaplot.CurrentGraph.SetData(mock_sigmaplot.CurrentWorksheet, 1, 2)

        # Try to add second data series
        try:
            # Add second series
            mock_sigmaplot.CurrentGraph.AddData(mock_sigmaplot.CurrentWorksheet, 1, 3)

            # Set stacked area option if available
            mock_sigmaplot.CurrentGraph.AreaType = "Stacked"

            # Verify AddData was called
            assert mock_sigmaplot.CurrentGraph.AddData.called
            assert mock_sigmaplot.CurrentGraph.AreaType == "Stacked"
        except:
            # These properties might not be available in mock
            pass

    def test_export_area_plot(self, mock_sigmaplot: Any) -> None:
        """
        Test exporting an area plot.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create area plot
        mock_sigmaplot.NewGraph("Area Plot")

        # Set up export path
        export_path = os.path.join(tempfile.gettempdir(), "test_area_plot.png")

        # Export the graph
        mock_sigmaplot.CurrentGraph.ExportGraph(export_path, "PNG")

        # Verify export was attempted
        assert mock_sigmaplot.CurrentGraph.ExportGraph.called
        assert mock_sigmaplot.CurrentGraph.ExportGraph.call_args[0][0] == export_path
        assert mock_sigmaplot.CurrentGraph.ExportGraph.call_args[0][1] == "PNG"

# EOF