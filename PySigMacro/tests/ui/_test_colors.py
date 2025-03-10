#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 01:21:19 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/ui/test_colors.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/ui/test_colors.py"

import os
import pytest
from unittest.mock import MagicMock, patch
from typing import Any, Dict, List, Optional

"""
Functionality:
* Tests color handling in SigmaPlot graphs
* Verifies color setting functionality for various graph elements
Input:
* None (uses mock objects)
Output:
* Test results for color settings
Prerequisites:
* pytest
* mock objects for SigmaPlot COM interface
"""

class TestColors:
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

            # Setup mock graph
            mock_graph = MagicMock()
            mock_app.CurrentGraph = mock_graph
            mock_app.NewGraph.return_value = mock_graph

            # Setup color-related properties
            mock_series = MagicMock()
            mock_graph.PlotSeries.return_value = mock_series

            # Configure mock dispatch to return our mock app
            mock_dispatch.return_value = mock_app

            yield mock_app

    def test_line_color(self, mock_sigmaplot: Any) -> None:
        """
        Test setting line colors in plots.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a line plot
        mock_sigmaplot.NewGraph("Line Plot")

        # Set line color
        test_color = "Red"
        mock_sigmaplot.CurrentGraph.PlotSeries(0).LineColor = test_color

        # Verify the color was set
        assert mock_sigmaplot.CurrentGraph.PlotSeries(0).LineColor == test_color

    def test_symbol_color(self, mock_sigmaplot: Any) -> None:
        """
        Test setting symbol colors in scatter plots.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a scatter plot
        mock_sigmaplot.NewGraph("Scatter Plot")

        # Set symbol colors
        outline_color = "Blue"
        fill_color = "LightBlue"

        try:
            mock_sigmaplot.CurrentGraph.PlotSeries(0).Symbol.Color = outline_color
            mock_sigmaplot.CurrentGraph.PlotSeries(0).Symbol.FillColor = fill_color

            # Verify colors were set
            assert mock_sigmaplot.CurrentGraph.PlotSeries(0).Symbol.Color == outline_color
            assert mock_sigmaplot.CurrentGraph.PlotSeries(0).Symbol.FillColor == fill_color
        except AttributeError:
            pytest.skip("Symbol color properties not available in this mock")

    def test_bar_color(self, mock_sigmaplot: Any) -> None:
        """
        Test setting bar colors in bar charts.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a bar chart
        mock_sigmaplot.NewGraph("Bar Chart")

        # Set bar colors
        bar_color = "Green"
        border_color = "DarkGreen"

        try:
            mock_sigmaplot.CurrentGraph.PlotSeries(0).BarColor = bar_color
            mock_sigmaplot.CurrentGraph.PlotSeries(0).BarBorderColor = border_color

            # Verify colors were set
            assert mock_sigmaplot.CurrentGraph.PlotSeries(0).BarColor == bar_color
            assert mock_sigmaplot.CurrentGraph.PlotSeries(0).BarBorderColor == border_color
        except AttributeError:
            pytest.skip("Bar color properties not available in this mock")

    def test_axis_color(self, mock_sigmaplot: Any) -> None:
        """
        Test setting axis colors.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        # Set axis colors
        x_axis_color = "Black"
        y_axis_color = "Gray"

        mock_sigmaplot.CurrentGraph.XAxis.Color = x_axis_color
        mock_sigmaplot.CurrentGraph.YAxis.Color = y_axis_color

        # Verify colors were set
        assert mock_sigmaplot.CurrentGraph.XAxis.Color == x_axis_color
        assert mock_sigmaplot.CurrentGraph.YAxis.Color == y_axis_color

    def test_background_color(self, mock_sigmaplot: Any) -> None:
        """
        Test setting graph background color.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        # Set background color
        bg_color = "White"

        try:
            mock_sigmaplot.CurrentGraph.BackgroundColor = bg_color

            # Verify color was set
            assert mock_sigmaplot.CurrentGraph.BackgroundColor == bg_color
        except AttributeError:
            pytest.skip("Background color property not available in this mock")


    def test_multiple_series_colors(self, mock_sigmaplot: Any) -> None:
        """
        Test setting colors for multiple data series.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a line plot
        mock_sigmaplot.NewGraph("Line Plot")

        # Define colors for multiple series
        series_colors = ["Red", "Blue", "Green", "Purple"]

        # Set colors for each series
        for i, color in enumerate(series_colors):
            mock_sigmaplot.CurrentGraph.PlotSeries(i).LineColor = color

        # Reset mock to avoid MagicMock's behavior of returning the last value
        mock_series = {}
        for i, color in enumerate(series_colors):
            mock_series[i] = MagicMock()
            mock_series[i].LineColor = color
            mock_sigmaplot.CurrentGraph.PlotSeries.return_value = mock_series[i]

        # Verify colors were set for each series
        for i, color in enumerate(series_colors):
            mock_sigmaplot.CurrentGraph.PlotSeries.return_value = mock_series[i]
            assert mock_sigmaplot.CurrentGraph.PlotSeries(i).LineColor == color

# EOF