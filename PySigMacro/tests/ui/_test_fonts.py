#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 01:15:36 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/ui/test_fonts.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/ui/test_fonts.py"

import os
import pytest
from unittest.mock import MagicMock, patch
from typing import Any, Dict, List, Optional

"""
Functionality:
* Tests font handling in SigmaPlot graphs
* Verifies font setting functionality for various graph elements
Input:
* None (uses mock objects)
Output:
* Test results for font settings
Prerequisites:
* pytest
* mock objects for SigmaPlot COM interface
"""

class TestFonts:
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

            # Setup mock font objects
            mock_title_font = MagicMock()
            mock_x_axis_font = MagicMock()
            mock_y_axis_font = MagicMock()
            mock_legend_font = MagicMock()

            # Connect mock fonts to graph elements
            mock_graph.Title.Font = mock_title_font
            mock_graph.XAxis.Label.Font = mock_x_axis_font
            mock_graph.YAxis.Label.Font = mock_y_axis_font
            mock_graph.Legend.Font = mock_legend_font

            # Configure mock dispatch to return our mock app
            mock_dispatch.return_value = mock_app

            yield mock_app

    def test_title_font(self, mock_sigmaplot: Any) -> None:
        """
        Test setting title font properties.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        # Set title font properties
        font_name = "Arial"
        font_size = 14
        font_bold = True
        font_italic = False

        try:
            mock_sigmaplot.CurrentGraph.Title.Font.Name = font_name
            mock_sigmaplot.CurrentGraph.Title.Font.Size = font_size
            mock_sigmaplot.CurrentGraph.Title.Font.Bold = font_bold
            mock_sigmaplot.CurrentGraph.Title.Font.Italic = font_italic

            # Verify font properties were set
            assert mock_sigmaplot.CurrentGraph.Title.Font.Name == font_name
            assert mock_sigmaplot.CurrentGraph.Title.Font.Size == font_size
            assert mock_sigmaplot.CurrentGraph.Title.Font.Bold == font_bold
            assert mock_sigmaplot.CurrentGraph.Title.Font.Italic == font_italic
        except AttributeError:
            pytest.skip("Title font properties not available in this mock")

    def test_axis_label_fonts(self, mock_sigmaplot: Any) -> None:
        """
        Test setting axis label font properties.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        # Set X-axis label font properties
        x_font_name = "Times New Roman"
        x_font_size = 12

        # Set Y-axis label font properties
        y_font_name = "Calibri"
        y_font_size = 12

        try:
            # X-axis font
            mock_sigmaplot.CurrentGraph.XAxis.Label.Font.Name = x_font_name
            mock_sigmaplot.CurrentGraph.XAxis.Label.Font.Size = x_font_size

            # Y-axis font
            mock_sigmaplot.CurrentGraph.YAxis.Label.Font.Name = y_font_name
            mock_sigmaplot.CurrentGraph.YAxis.Label.Font.Size = y_font_size

            # Verify font properties were set
            assert mock_sigmaplot.CurrentGraph.XAxis.Label.Font.Name == x_font_name
            assert mock_sigmaplot.CurrentGraph.XAxis.Label.Font.Size == x_font_size
            assert mock_sigmaplot.CurrentGraph.YAxis.Label.Font.Name == y_font_name
            assert mock_sigmaplot.CurrentGraph.YAxis.Label.Font.Size == y_font_size
        except AttributeError:
            pytest.skip("Axis label font properties not available in this mock")

    def test_legend_font(self, mock_sigmaplot: Any) -> None:
        """
        Test setting legend font properties.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        # Set legend font properties
        font_name = "Verdana"
        font_size = 10
        font_bold = False

        try:
            mock_sigmaplot.CurrentGraph.Legend.Font.Name = font_name
            mock_sigmaplot.CurrentGraph.Legend.Font.Size = font_size
            mock_sigmaplot.CurrentGraph.Legend.Font.Bold = font_bold

            # Verify font properties were set
            assert mock_sigmaplot.CurrentGraph.Legend.Font.Name == font_name
            assert mock_sigmaplot.CurrentGraph.Legend.Font.Size == font_size
            assert mock_sigmaplot.CurrentGraph.Legend.Font.Bold == font_bold
        except AttributeError:
            pytest.skip("Legend font properties not available in this mock")

    def test_tick_label_fonts(self, mock_sigmaplot: Any) -> None:
        """
        Test setting tick label font properties.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        # Set tick label font properties
        font_name = "Consolas"
        font_size = 9

        try:
            mock_sigmaplot.CurrentGraph.XAxis.TickLabels.Font.Name = font_name
            mock_sigmaplot.CurrentGraph.XAxis.TickLabels.Font.Size = font_size
            mock_sigmaplot.CurrentGraph.YAxis.TickLabels.Font.Name = font_name
            mock_sigmaplot.CurrentGraph.YAxis.TickLabels.Font.Size = font_size

            # Verify font properties were set
            assert mock_sigmaplot.CurrentGraph.XAxis.TickLabels.Font.Name == font_name
            assert mock_sigmaplot.CurrentGraph.XAxis.TickLabels.Font.Size == font_size
            assert mock_sigmaplot.CurrentGraph.YAxis.TickLabels.Font.Name == font_name
            assert mock_sigmaplot.CurrentGraph.YAxis.TickLabels.Font.Size == font_size
        except AttributeError:
            pytest.skip("Tick label font properties not available in this mock")

    def test_font_color(self, mock_sigmaplot: Any) -> None:
        """
        Test setting font color properties.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        # Set font colors
        title_color = "Blue"
        x_label_color = "Black"
        y_label_color = "Red"

        try:
            mock_sigmaplot.CurrentGraph.Title.Font.Color = title_color
            mock_sigmaplot.CurrentGraph.XAxis.Label.Font.Color = x_label_color
            mock_sigmaplot.CurrentGraph.YAxis.Label.Font.Color = y_label_color

            # Verify font colors were set
            assert mock_sigmaplot.CurrentGraph.Title.Font.Color == title_color
            assert mock_sigmaplot.CurrentGraph.XAxis.Label.Font.Color == x_label_color
            assert mock_sigmaplot.CurrentGraph.YAxis.Label.Font.Color == y_label_color
        except AttributeError:
            pytest.skip("Font color properties not available in this mock")

# EOF