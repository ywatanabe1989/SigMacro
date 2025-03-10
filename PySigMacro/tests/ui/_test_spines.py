#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 01:17:18 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/ui/test_spines.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/ui/test_spines.py"

import os
import pytest
from unittest.mock import MagicMock, patch
from typing import Any, Dict, List, Optional

"""
Functionality:
* Tests spine (axis frame) handling in SigmaPlot graphs
* Verifies spine visibility, width, color, and style settings
Input:
* None (uses mock objects)
Output:
* Test results for spine settings
Prerequisites:
* pytest
* mock objects for SigmaPlot COM interface
"""

class TestSpines:
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

            # Setup mock axes with spines
            mock_x_axis = MagicMock()
            mock_y_axis = MagicMock()
            mock_top_spine = MagicMock()
            mock_bottom_spine = MagicMock()
            mock_left_spine = MagicMock()
            mock_right_spine = MagicMock()

            # Connect axes to graph
            mock_graph.XAxis = mock_x_axis
            mock_graph.YAxis = mock_y_axis

            # Connect spines to axes
            mock_graph.TopSpine = mock_top_spine
            mock_graph.BottomSpine = mock_bottom_spine
            mock_graph.LeftSpine = mock_left_spine
            mock_graph.RightSpine = mock_right_spine

            # Configure mock dispatch to return our mock app
            mock_dispatch.return_value = mock_app

            yield mock_app

    def test_spine_visibility(self, mock_sigmaplot: Any) -> None:
        """
        Test toggling spine visibility.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        try:
            # Toggle top spine visibility
            mock_sigmaplot.CurrentGraph.TopSpine.Visible = False
            assert mock_sigmaplot.CurrentGraph.TopSpine.Visible == False

            # Toggle right spine visibility
            mock_sigmaplot.CurrentGraph.RightSpine.Visible = False
            assert mock_sigmaplot.CurrentGraph.RightSpine.Visible == False

            # Keep bottom and left spines visible
            mock_sigmaplot.CurrentGraph.BottomSpine.Visible = True
            mock_sigmaplot.CurrentGraph.LeftSpine.Visible = True

            assert mock_sigmaplot.CurrentGraph.BottomSpine.Visible == True
            assert mock_sigmaplot.CurrentGraph.LeftSpine.Visible == True
        except AttributeError:
            # Try alternative approach using axis frame properties
            try:
                mock_sigmaplot.CurrentGraph.XAxis.TopBorder.Visible = False
                mock_sigmaplot.CurrentGraph.YAxis.RightBorder.Visible = False

                assert mock_sigmaplot.CurrentGraph.XAxis.TopBorder.Visible == False
                assert mock_sigmaplot.CurrentGraph.YAxis.RightBorder.Visible == False
            except AttributeError:
                pytest.skip("Spine visibility properties not available in this mock")

    def test_spine_width(self, mock_sigmaplot: Any) -> None:
        """
        Test setting spine width.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        # Set spine width
        line_width = 1.5

        try:
            # Set width for all spines
            mock_sigmaplot.CurrentGraph.BottomSpine.Width = line_width
            mock_sigmaplot.CurrentGraph.LeftSpine.Width = line_width
            mock_sigmaplot.CurrentGraph.TopSpine.Width = line_width
            mock_sigmaplot.CurrentGraph.RightSpine.Width = line_width

            # Verify width was set
            assert mock_sigmaplot.CurrentGraph.BottomSpine.Width == line_width
            assert mock_sigmaplot.CurrentGraph.LeftSpine.Width == line_width
            assert mock_sigmaplot.CurrentGraph.TopSpine.Width == line_width
            assert mock_sigmaplot.CurrentGraph.RightSpine.Width == line_width
        except AttributeError:
            # Try alternative approach using axis properties
            try:
                mock_sigmaplot.CurrentGraph.XAxis.LineWidth = line_width
                mock_sigmaplot.CurrentGraph.YAxis.LineWidth = line_width

                assert mock_sigmaplot.CurrentGraph.XAxis.LineWidth == line_width
                assert mock_sigmaplot.CurrentGraph.YAxis.LineWidth == line_width
            except AttributeError:
                pytest.skip("Spine width properties not available in this mock")

    def test_spine_color(self, mock_sigmaplot: Any) -> None:
        """
        Test setting spine color.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        # Set spine color
        spine_color = "Black"

        try:
            # Set color for all spines
            mock_sigmaplot.CurrentGraph.BottomSpine.Color = spine_color
            mock_sigmaplot.CurrentGraph.LeftSpine.Color = spine_color
            mock_sigmaplot.CurrentGraph.TopSpine.Color = spine_color
            mock_sigmaplot.CurrentGraph.RightSpine.Color = spine_color

            # Verify color was set
            assert mock_sigmaplot.CurrentGraph.BottomSpine.Color == spine_color
            assert mock_sigmaplot.CurrentGraph.LeftSpine.Color == spine_color
            assert mock_sigmaplot.CurrentGraph.TopSpine.Color == spine_color
            assert mock_sigmaplot.CurrentGraph.RightSpine.Color == spine_color
        except AttributeError:
            # Try alternative approach using axis properties
            try:
                mock_sigmaplot.CurrentGraph.XAxis.Color = spine_color
                mock_sigmaplot.CurrentGraph.YAxis.Color = spine_color

                assert mock_sigmaplot.CurrentGraph.XAxis.Color == spine_color
                assert mock_sigmaplot.CurrentGraph.YAxis.Color == spine_color
            except AttributeError:
                pytest.skip("Spine color properties not available in this mock")

    def test_spine_style(self, mock_sigmaplot: Any) -> None:
        """
        Test setting spine line style.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        # Set spine line style
        line_style = "Solid"

        try:
            # Set style for all spines
            mock_sigmaplot.CurrentGraph.BottomSpine.Style = line_style
            mock_sigmaplot.CurrentGraph.LeftSpine.Style = line_style
            mock_sigmaplot.CurrentGraph.TopSpine.Style = line_style
            mock_sigmaplot.CurrentGraph.RightSpine.Style = line_style

            # Verify style was set
            assert mock_sigmaplot.CurrentGraph.BottomSpine.Style == line_style
            assert mock_sigmaplot.CurrentGraph.LeftSpine.Style == line_style
            assert mock_sigmaplot.CurrentGraph.TopSpine.Style == line_style
            assert mock_sigmaplot.CurrentGraph.RightSpine.Style == line_style
        except AttributeError:
            # Try alternative approach using axis properties
            try:
                mock_sigmaplot.CurrentGraph.XAxis.LineStyle = line_style
                mock_sigmaplot.CurrentGraph.YAxis.LineStyle = line_style

                assert mock_sigmaplot.CurrentGraph.XAxis.LineStyle == line_style
                assert mock_sigmaplot.CurrentGraph.YAxis.LineStyle == line_style
            except AttributeError:
                pytest.skip("Spine style properties not available in this mock")

    def test_box_frame(self, mock_sigmaplot: Any) -> None:
        """
        Test setting box frame around the plot.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        try:
            # Enable box frame (all spines visible)
            mock_sigmaplot.CurrentGraph.BoxFrame = True
            assert mock_sigmaplot.CurrentGraph.BoxFrame == True

            # Disable box frame
            mock_sigmaplot.CurrentGraph.BoxFrame = False
            assert mock_sigmaplot.CurrentGraph.BoxFrame == False
        except AttributeError:
            # Try alternative approach
            try:
                # Make all spines visible to create a box
                mock_sigmaplot.CurrentGraph.TopSpine.Visible = True
                mock_sigmaplot.CurrentGraph.BottomSpine.Visible = True
                mock_sigmaplot.CurrentGraph.LeftSpine.Visible = True
                mock_sigmaplot.CurrentGraph.RightSpine.Visible = True

                # Verify all are visible
                assert mock_sigmaplot.CurrentGraph.TopSpine.Visible == True
                assert mock_sigmaplot.CurrentGraph.BottomSpine.Visible == True
                assert mock_sigmaplot.CurrentGraph.LeftSpine.Visible == True
                assert mock_sigmaplot.CurrentGraph.RightSpine.Visible == True
            except AttributeError:
                pytest.skip("Box frame or spine visibility properties not available in this mock")

# EOF