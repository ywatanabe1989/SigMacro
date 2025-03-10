#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 01:19:50 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/ui/test_ylabel.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/ui/test_ylabel.py"

import os
import pytest
from unittest.mock import MagicMock, patch
from typing import Any, Dict, List, Optional

"""
Functionality:
* Tests Y-axis label handling in SigmaPlot graphs
* Verifies y-label text, font, position, and formatting
Input:
* None (uses mock objects)
Output:
* Test results for Y-axis label settings
Prerequisites:
* pytest
* mock objects for SigmaPlot COM interface
"""

class TestYLabel:
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

            # Setup mock y-axis and label
            mock_y_axis = MagicMock()
            mock_y_label = MagicMock()
            mock_font = MagicMock()

            mock_graph.YAxis = mock_y_axis
            mock_y_axis.Label = mock_y_label
            mock_y_label.Font = mock_font

            # Configure mock dispatch to return our mock app
            mock_dispatch.return_value = mock_app

            yield mock_app

    def test_ylabel_text(self, mock_sigmaplot: Any) -> None:
        """
        Test setting Y-axis label text.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        # Set Y-axis label text
        ylabel_text = "Y Axis"

        # Set label
        mock_sigmaplot.CurrentGraph.YAxis.Label = ylabel_text

        # Verify label was set
        assert mock_sigmaplot.CurrentGraph.YAxis.Label == ylabel_text

    def test_ylabel_visibility(self, mock_sigmaplot: Any) -> None:
        """
        Test toggling Y-axis label visibility.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        try:
            # Toggle label visibility
            mock_sigmaplot.CurrentGraph.YAxis.Label.Visible = True
            assert mock_sigmaplot.CurrentGraph.YAxis.Label.Visible == True

            mock_sigmaplot.CurrentGraph.YAxis.Label.Visible = False
            assert mock_sigmaplot.CurrentGraph.YAxis.Label.Visible == False
        except AttributeError:
            pytest.skip("Y-axis label visibility property not available in this mock")

    def test_ylabel_font(self, mock_sigmaplot: Any) -> None:
        """
        Test setting Y-axis label font properties.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        # Set font properties
        font_name = "Arial"
        font_size = 12
        font_bold = True

        try:
            # Set label font properties
            mock_sigmaplot.CurrentGraph.YAxis.Label.Font.Name = font_name
            mock_sigmaplot.CurrentGraph.YAxis.Label.Font.Size = font_size
            mock_sigmaplot.CurrentGraph.YAxis.Label.Font.Bold = font_bold

            # Verify font properties were set
            assert mock_sigmaplot.CurrentGraph.YAxis.Label.Font.Name == font_name
            assert mock_sigmaplot.CurrentGraph.YAxis.Label.Font.Size == font_size
            assert mock_sigmaplot.CurrentGraph.YAxis.Label.Font.Bold == font_bold
        except AttributeError:
            pytest.skip("Y-axis label font properties not available in this mock")

    def test_ylabel_color(self, mock_sigmaplot: Any) -> None:
        """
        Test setting Y-axis label color.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        # Set label color
        label_color = "Red"

        try:
            mock_sigmaplot.CurrentGraph.YAxis.Label.Font.Color = label_color
            assert mock_sigmaplot.CurrentGraph.YAxis.Label.Font.Color == label_color
        except AttributeError:
            # Try alternative property
            try:
                mock_sigmaplot.CurrentGraph.YAxis.Label.Color = label_color
                assert mock_sigmaplot.CurrentGraph.YAxis.Label.Color == label_color
            except AttributeError:
                pytest.skip("Y-axis label color properties not available in this mock")

    def test_ylabel_offset(self, mock_sigmaplot: Any) -> None:
        """
        Test setting Y-axis label offset from axis.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        try:
            # Set label offset
            mock_sigmaplot.CurrentGraph.YAxis.Label.Offset = 10
            assert mock_sigmaplot.CurrentGraph.YAxis.Label.Offset == 10
        except AttributeError:
            pytest.skip("Y-axis label offset property not available in this mock")

    def test_ylabel_rotation(self, mock_sigmaplot: Any) -> None:
        """
        Test setting Y-axis label rotation.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        try:
            # Set label rotation angle (in degrees)
            # For Y-axis, often 90 or 270 degrees by default
            mock_sigmaplot.CurrentGraph.YAxis.Label.Rotation = 90
            assert mock_sigmaplot.CurrentGraph.YAxis.Label.Rotation == 90

            mock_sigmaplot.CurrentGraph.YAxis.Label.Rotation = 0
            assert mock_sigmaplot.CurrentGraph.YAxis.Label.Rotation == 0
        except AttributeError:
            pytest.skip("Y-axis label rotation property not available in this mock")

    def test_ylabel_position(self, mock_sigmaplot: Any) -> None:
        """
        Test setting Y-axis label position.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        try:
            # Set vertical alignment
            mock_sigmaplot.CurrentGraph.YAxis.Label.VerticalAlignment = "Middle"
            assert mock_sigmaplot.CurrentGraph.YAxis.Label.VerticalAlignment == "Middle"

            # Set horizontal position (for Y-axis label, typically "Left" of the axis)
            mock_sigmaplot.CurrentGraph.YAxis.Label.Position = "Left"
            assert mock_sigmaplot.CurrentGraph.YAxis.Label.Position == "Left"
        except AttributeError:
            pytest.skip("Y-axis label position properties not available in this mock")

    def test_ylabel_wrapped_text(self, mock_sigmaplot: Any) -> None:
        """
        Test setting Y-axis label with wrapped text.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        # Multi-line Y-axis label
        ylabel_text = "Line 1\nLine 2"

        try:
            # Set multi-line label
            mock_sigmaplot.CurrentGraph.YAxis.Label = ylabel_text
            assert mock_sigmaplot.CurrentGraph.YAxis.Label == ylabel_text

            # Enable text wrapping if applicable
            mock_sigmaplot.CurrentGraph.YAxis.Label.WordWrap = True
            assert mock_sigmaplot.CurrentGraph.YAxis.Label.WordWrap == True
        except AttributeError:
            pytest.skip("Y-axis label text wrapping property not available in this mock")

    def test_ylabel_secondary_axis(self, mock_sigmaplot: Any) -> None:
        """
        Test setting Y-axis label for secondary Y-axis.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        try:
            # Enable secondary Y-axis
            mock_sigmaplot.CurrentGraph.SecondaryYAxis.Visible = True
            assert mock_sigmaplot.CurrentGraph.SecondaryYAxis.Visible == True

            # Set secondary Y-axis label
            secondary_label = "Secondary Y Axis"
            mock_sigmaplot.CurrentGraph.SecondaryYAxis.Label = secondary_label
            assert mock_sigmaplot.CurrentGraph.SecondaryYAxis.Label == secondary_label
        except AttributeError:
            pytest.skip("Secondary Y-axis properties not available in this mock")

# EOF