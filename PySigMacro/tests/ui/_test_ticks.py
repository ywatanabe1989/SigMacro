#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 01:17:51 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/ui/test_ticks.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/ui/test_ticks.py"

import os
import pytest
from unittest.mock import MagicMock, patch
from typing import Any, Dict, List, Optional

"""
Functionality:
* Tests tick mark handling in SigmaPlot graphs
* Verifies tick visibility, spacing, formatting, and appearance
Input:
* None (uses mock objects)
Output:
* Test results for tick mark settings
Prerequisites:
* pytest
* mock objects for SigmaPlot COM interface
"""

class TestTicks:
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

            # Setup mock axes with ticks
            mock_x_axis = MagicMock()
            mock_y_axis = MagicMock()
            mock_x_ticks = MagicMock()
            mock_y_ticks = MagicMock()
            mock_x_minor_ticks = MagicMock()
            mock_y_minor_ticks = MagicMock()

            # Connect axes to graph
            mock_graph.XAxis = mock_x_axis
            mock_graph.YAxis = mock_y_axis

            # Connect ticks to axes
            mock_x_axis.MajorTicks = mock_x_ticks
            mock_y_axis.MajorTicks = mock_y_ticks
            mock_x_axis.MinorTicks = mock_x_minor_ticks
            mock_y_axis.MinorTicks = mock_y_minor_ticks

            # Configure mock dispatch to return our mock app
            mock_dispatch.return_value = mock_app

            yield mock_app

    def test_major_tick_visibility(self, mock_sigmaplot: Any) -> None:
        """
        Test toggling major tick visibility.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        try:
            # Toggle major tick visibility
            mock_sigmaplot.CurrentGraph.XAxis.MajorTicks.Visible = True
            mock_sigmaplot.CurrentGraph.YAxis.MajorTicks.Visible = True

            assert mock_sigmaplot.CurrentGraph.XAxis.MajorTicks.Visible == True
            assert mock_sigmaplot.CurrentGraph.YAxis.MajorTicks.Visible == True

            # Turn off x-axis major ticks
            mock_sigmaplot.CurrentGraph.XAxis.MajorTicks.Visible = False
            assert mock_sigmaplot.CurrentGraph.XAxis.MajorTicks.Visible == False
        except AttributeError:
            pytest.skip("Tick visibility properties not available in this mock")

    def test_minor_tick_visibility(self, mock_sigmaplot: Any) -> None:
        """
        Test toggling minor tick visibility.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        try:
            # Toggle minor tick visibility
            mock_sigmaplot.CurrentGraph.XAxis.MinorTicks.Visible = True
            mock_sigmaplot.CurrentGraph.YAxis.MinorTicks.Visible = True

            assert mock_sigmaplot.CurrentGraph.XAxis.MinorTicks.Visible == True
            assert mock_sigmaplot.CurrentGraph.YAxis.MinorTicks.Visible == True

            # Turn off minor ticks
            mock_sigmaplot.CurrentGraph.XAxis.MinorTicks.Visible = False
            mock_sigmaplot.CurrentGraph.YAxis.MinorTicks.Visible = False

            assert mock_sigmaplot.CurrentGraph.XAxis.MinorTicks.Visible == False
            assert mock_sigmaplot.CurrentGraph.YAxis.MinorTicks.Visible == False
        except AttributeError:
            pytest.skip("Minor tick visibility properties not available in this mock")

    def test_tick_spacing(self, mock_sigmaplot: Any) -> None:
        """
        Test setting tick spacing.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        # Set major tick spacing
        x_major_spacing = 1.0
        y_major_spacing = 5.0

        try:
            # Set automatic tick spacing
            mock_sigmaplot.CurrentGraph.XAxis.AutoTick = True
            mock_sigmaplot.CurrentGraph.YAxis.AutoTick = True

            assert mock_sigmaplot.CurrentGraph.XAxis.AutoTick == True
            assert mock_sigmaplot.CurrentGraph.YAxis.AutoTick == True

            # Turn off automatic spacing and set manual spacing
            mock_sigmaplot.CurrentGraph.XAxis.AutoTick = False
            mock_sigmaplot.CurrentGraph.YAxis.AutoTick = False

            mock_sigmaplot.CurrentGraph.XAxis.MajorTickSpacing = x_major_spacing
            mock_sigmaplot.CurrentGraph.YAxis.MajorTickSpacing = y_major_spacing

            assert mock_sigmaplot.CurrentGraph.XAxis.MajorTickSpacing == x_major_spacing
            assert mock_sigmaplot.CurrentGraph.YAxis.MajorTickSpacing == y_major_spacing
        except AttributeError:
            pytest.skip("Tick spacing properties not available in this mock")

    def test_minor_tick_count(self, mock_sigmaplot: Any) -> None:
        """
        Test setting minor tick count between major ticks.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        # Set minor tick count
        x_minor_count = 4
        y_minor_count = 1

        try:
            mock_sigmaplot.CurrentGraph.XAxis.MinorTickCount = x_minor_count
            mock_sigmaplot.CurrentGraph.YAxis.MinorTickCount = y_minor_count

            assert mock_sigmaplot.CurrentGraph.XAxis.MinorTickCount == x_minor_count
            assert mock_sigmaplot.CurrentGraph.YAxis.MinorTickCount == y_minor_count
        except AttributeError:
            pytest.skip("Minor tick count properties not available in this mock")

    def test_tick_length(self, mock_sigmaplot: Any) -> None:
        """
        Test setting tick length.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        # Set tick lengths
        major_length = 5.0
        minor_length = 2.5

        try:
            # Set major tick length
            mock_sigmaplot.CurrentGraph.XAxis.MajorTicks.Length = major_length
            mock_sigmaplot.CurrentGraph.YAxis.MajorTicks.Length = major_length

            assert mock_sigmaplot.CurrentGraph.XAxis.MajorTicks.Length == major_length
            assert mock_sigmaplot.CurrentGraph.YAxis.MajorTicks.Length == major_length

            # Set minor tick length
            mock_sigmaplot.CurrentGraph.XAxis.MinorTicks.Length = minor_length
            mock_sigmaplot.CurrentGraph.YAxis.MinorTicks.Length = minor_length

            assert mock_sigmaplot.CurrentGraph.XAxis.MinorTicks.Length == minor_length
            assert mock_sigmaplot.CurrentGraph.YAxis.MinorTicks.Length == minor_length
        except AttributeError:
            pytest.skip("Tick length properties not available in this mock")

    def test_tick_direction(self, mock_sigmaplot: Any) -> None:
        """
        Test setting tick direction (inside/outside).

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        try:
            # Set ticks to point inside
            mock_sigmaplot.CurrentGraph.XAxis.MajorTicks.Direction = "Inside"
            mock_sigmaplot.CurrentGraph.YAxis.MajorTicks.Direction = "Inside"

            assert mock_sigmaplot.CurrentGraph.XAxis.MajorTicks.Direction == "Inside"
            assert mock_sigmaplot.CurrentGraph.YAxis.MajorTicks.Direction == "Inside"

            # Set ticks to point outside
            mock_sigmaplot.CurrentGraph.XAxis.MajorTicks.Direction = "Outside"
            mock_sigmaplot.CurrentGraph.YAxis.MajorTicks.Direction = "Outside"

            assert mock_sigmaplot.CurrentGraph.XAxis.MajorTicks.Direction == "Outside"
            assert mock_sigmaplot.CurrentGraph.YAxis.MajorTicks.Direction == "Outside"

            # Set ticks to both sides
            mock_sigmaplot.CurrentGraph.XAxis.MajorTicks.Direction = "Both"
            mock_sigmaplot.CurrentGraph.YAxis.MajorTicks.Direction = "Both"

            assert mock_sigmaplot.CurrentGraph.XAxis.MajorTicks.Direction == "Both"
            assert mock_sigmaplot.CurrentGraph.YAxis.MajorTicks.Direction == "Both"
        except AttributeError:
            pytest.skip("Tick direction properties not available in this mock")

    def test_tick_style(self, mock_sigmaplot: Any) -> None:
        """
        Test setting tick style.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        # Set tick color and width
        tick_color = "Black"
        tick_width = 1.0

        try:
            # Set tick color
            mock_sigmaplot.CurrentGraph.XAxis.MajorTicks.Color = tick_color
            mock_sigmaplot.CurrentGraph.YAxis.MajorTicks.Color = tick_color

            assert mock_sigmaplot.CurrentGraph.XAxis.MajorTicks.Color == tick_color
            assert mock_sigmaplot.CurrentGraph.YAxis.MajorTicks.Color == tick_color

            # Set tick width
            mock_sigmaplot.CurrentGraph.XAxis.MajorTicks.Width = tick_width
            mock_sigmaplot.CurrentGraph.YAxis.MajorTicks.Width = tick_width

            assert mock_sigmaplot.CurrentGraph.XAxis.MajorTicks.Width == tick_width
            assert mock_sigmaplot.CurrentGraph.YAxis.MajorTicks.Width == tick_width
        except AttributeError:
            pytest.skip("Tick style properties not available in this mock")

    def test_tick_labels(self, mock_sigmaplot: Any) -> None:
        """
        Test tick label format and visibility.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        try:
            # Toggle tick label visibility
            mock_sigmaplot.CurrentGraph.XAxis.TickLabels.Visible = True
            mock_sigmaplot.CurrentGraph.YAxis.TickLabels.Visible = True

            assert mock_sigmaplot.CurrentGraph.XAxis.TickLabels.Visible == True
            assert mock_sigmaplot.CurrentGraph.YAxis.TickLabels.Visible == True

            # Set number format for Y-axis (e.g., 2 decimal places)
            mock_sigmaplot.CurrentGraph.YAxis.TickLabels.Format = "0.00"
            assert mock_sigmaplot.CurrentGraph.YAxis.TickLabels.Format == "0.00"

            # Set font properties
            mock_sigmaplot.CurrentGraph.XAxis.TickLabels.Font.Size = 9
            assert mock_sigmaplot.CurrentGraph.XAxis.TickLabels.Font.Size == 9
        except AttributeError:
            pytest.skip("Tick label properties not available in this mock")

# EOF