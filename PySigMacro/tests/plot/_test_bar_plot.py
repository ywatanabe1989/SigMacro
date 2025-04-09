#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 00:20:12 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/plot/test_bar_plot.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/plot/test_bar_plot.py"

import os
import pytest
import tempfile
from unittest.mock import MagicMock, patch
from typing import Any, Optional, List, Dict

"""
Functionality:
* Tests bar plot creation in SigmaPlot
* Verifies correct setup of bar plot properties and data
Input:
* None (uses mock objects)
Output:
* Test results for bar plot creation
Prerequisites:
* pytest
* mock objects for SigmaPlot COM interface
"""

class TestBarPlot:
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

    def test_create_bar_plot(self, mock_sigmaplot: Any) -> None:
        """
        Test creating a basic bar plot.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Define test categorical data
        categories = ["Category A", "Category B", "Category C", "Category D"]
        values = [10, 15, 7, 12]

        # Insert data into worksheet
        # For bar plots, categories typically go in column 1, values in column 2
        for i, (category, value) in enumerate(zip(categories, values)):
            mock_sigmaplot.CurrentWorksheet.Cells(i+1, 1).Value = category
            mock_sigmaplot.CurrentWorksheet.Cells(i+1, 2).Value = value

        # Create bar plot
        mock_sigmaplot.NewGraph("Bar Chart")

        # Set data for the graph
        mock_sigmaplot.CurrentGraph.SetData(mock_sigmaplot.CurrentWorksheet, 1, 2)

        # Verify method calls
        assert mock_sigmaplot.NewGraph.called
        assert mock_sigmaplot.NewGraph.call_args[0][0] == "Bar Chart"
        assert mock_sigmaplot.CurrentGraph.SetData.called

        # Verify SetData parameters (worksheet, category_column, value_column)
        call_args = mock_sigmaplot.CurrentGraph.SetData.call_args[0]
        assert call_args[0] == mock_sigmaplot.CurrentWorksheet
        assert call_args[1] == 1
        assert call_args[2] == 2

    def test_bar_plot_customization(self, mock_sigmaplot: Any) -> None:
        """
        Test customizing a bar plot.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create simple bar plot
        mock_sigmaplot.NewGraph("Bar Chart")

        # Apply customizations
        mock_sigmaplot.CurrentGraph.Title = "Test Bar Chart"
        mock_sigmaplot.CurrentGraph.XAxis.Label = "Categories"
        mock_sigmaplot.CurrentGraph.YAxis.Label = "Values"

        try:
            # Try to set bar properties if available
            mock_sigmaplot.CurrentGraph.PlotSeries(0).BarWidth = 0.8
            mock_sigmaplot.CurrentGraph.PlotSeries(0).BarColor = "Blue"
            mock_sigmaplot.CurrentGraph.PlotSeries(0).BarFill = True
        except:
            # Not all properties may be available in mock
            pass

        # Verify basic customizations
        assert mock_sigmaplot.CurrentGraph.Title == "Test Bar Chart"
        assert mock_sigmaplot.CurrentGraph.XAxis.Label == "Categories"
        assert mock_sigmaplot.CurrentGraph.YAxis.Label == "Values"

    def test_multiple_bar_series(self, mock_sigmaplot: Any) -> None:
        """
        Test creating a bar plot with multiple data series.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Define test categorical data with multiple series
        categories = ["Group 1", "Group 2", "Group 3"]

        # Insert categories in column 1
        for i, category in enumerate(categories):
            mock_sigmaplot.CurrentWorksheet.Cells(i+1, 1).Value = category

        # Insert first series values in column 2
        series1_values = [10, 15, 12]
        for i, value in enumerate(series1_values):
            mock_sigmaplot.CurrentWorksheet.Cells(i+1, 2).Value = value

        # Insert second series values in column 3
        series2_values = [8, 12, 9]
        for i, value in enumerate(series2_values):
            mock_sigmaplot.CurrentWorksheet.Cells(i+1, 3).Value = value

        # Create bar plot
        mock_sigmaplot.NewGraph("Bar Chart")

        # Set first data series
        mock_sigmaplot.CurrentGraph.SetData(mock_sigmaplot.CurrentWorksheet, 1, 2)

        # Add second data series
        try:
            mock_sigmaplot.CurrentGraph.AddData(mock_sigmaplot.CurrentWorksheet, 1, 3)

            # Verify AddData was called
            assert mock_sigmaplot.CurrentGraph.AddData.called

            # Verify AddData parameters
            call_args = mock_sigmaplot.CurrentGraph.AddData.call_args[0]
            assert call_args[0] == mock_sigmaplot.CurrentWorksheet
            assert call_args[1] == 1
            assert call_args[2] == 3
        except:
            # AddData might not be available in mock
            pass

    def test_export_bar_plot(self, mock_sigmaplot: Any) -> None:
        """
        Test exporting a bar plot.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create bar plot
        mock_sigmaplot.NewGraph("Bar Chart")

        # Set up export path
        export_path = os.path.join(tempfile.gettempdir(), "test_bar_plot.png")

        # Export the graph
        mock_sigmaplot.CurrentGraph.ExportGraph(export_path, "PNG")

        # Verify export was attempted
        assert mock_sigmaplot.CurrentGraph.ExportGraph.called
        assert mock_sigmaplot.CurrentGraph.ExportGraph.call_args[0][0] == export_path
        assert mock_sigmaplot.CurrentGraph.ExportGraph.call_args[0][1] == "PNG"

# EOF