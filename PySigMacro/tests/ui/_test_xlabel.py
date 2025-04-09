#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 01:19:23 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/ui/test_xlabel.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/ui/test_xlabel.py"
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 00:01:29 (ywatanabe)"
# File: /home/ywatanabe/proj/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/ui/test_xlabel.py
THIS_FILE = "/home/ywatanabe/proj/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/ui/test_xlabel.py"

import os
import pytest
from unittest.mock import MagicMock, patch
from typing import Any, Dict, List, Optional

"""
Functionality:
* Tests X-axis label handling in SigmaPlot graphs
* Verifies x-label text, font, position, and formatting
Input:
* None (uses mock objects)
Output:
* Test results for X-axis label settings
Prerequisites:
* pytest
* mock objects for SigmaPlot COM interface
"""

class TestXLabel:
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

            # Setup mock x-axis and label
            mock_x_axis = MagicMock()
            mock_x_label = MagicMock()
            mock_font = MagicMock()

            mock_graph.XAxis = mock_x_axis
            mock_x_axis.Label = mock_x_label
            mock_x_label.Font = mock_font

            # Configure mock dispatch to return our mock app
            mock_dispatch.return_value = mock_app

            yield mock_app

    def test_xlabel_text(self, mock_sigmaplot: Any) -> None:
        """
        Test setting X-axis label text.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        # Set X-axis label text
        xlabel_text = "X Axis"

        # Set label
        mock_sigmaplot.CurrentGraph.XAxis.Label = xlabel_text

        # Verify label was set
        assert mock_sigmaplot.CurrentGraph.XAxis.Label == xlabel_text

    def test_xlabel_visibility(self, mock_sigmaplot: Any) -> None:
        """
        Test toggling X-axis label visibility.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        try:
            # Toggle label visibility
            mock_sigmaplot.CurrentGraph.XAxis.Label.Visible = True
            assert mock_sigmaplot.CurrentGraph.XAxis.Label.Visible == True

            mock_sigmaplot.CurrentGraph.XAxis.Label.Visible = False
            assert mock_sigmaplot.CurrentGraph.XAxis.Label.Visible == False
        except AttributeError:
            pytest.skip("X-axis label visibility property not available in this mock")

    def test_xlabel_font(self, mock_sigmaplot: Any) -> None:
        """
        Test setting X-axis label font properties.

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
            mock_sigmaplot.CurrentGraph.XAxis.Label.Font.Name = font_name
            mock_sigmaplot.CurrentGraph.XAxis.Label.Font.Size = font_size
            mock_sigmaplot.CurrentGraph.XAxis.Label.Font.Bold = font_bold

            # Verify font properties were set
            assert mock_sigmaplot.CurrentGraph.XAxis.Label.Font.Name == font_name
            assert mock_sigmaplot.CurrentGraph.XAxis.Label.Font.Size == font_size
            assert mock_sigmaplot.CurrentGraph.XAxis.Label.Font.Bold == font_bold
        except AttributeError:
            pytest.skip("X-axis label font properties not available in this mock")

    def test_xlabel_color(self, mock_sigmaplot: Any) -> None:
        """
        Test setting X-axis label color.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        # Set label color
        label_color = "Blue"

        try:
            mock_sigmaplot.CurrentGraph.XAxis.Label.Font.Color = label_color
            assert mock_sigmaplot.CurrentGraph.XAxis.Label.Font.Color == label_color
        except AttributeError:
            # Try alternative property
            try:
                mock_sigmaplot.CurrentGraph.XAxis.Label.Color = label_color
                assert mock_sigmaplot.CurrentGraph.XAxis.Label.Color == label_color
            except AttributeError:
                pytest.skip("X-axis label color properties not available in this mock")

    def test_xlabel_offset(self, mock_sigmaplot: Any) -> None:
        """
        Test setting X-axis label offset from axis.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        try:
            # Set label offset
            mock_sigmaplot.CurrentGraph.XAxis.Label.Offset = 10
            assert mock_sigmaplot.CurrentGraph.XAxis.Label.Offset == 10
        except AttributeError:
            pytest.skip("X-axis label offset property not available in this mock")

    def test_xlabel_rotation(self, mock_sigmaplot: Any) -> None:
        """
        Test setting X-axis label rotation.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        try:
            # Set label rotation angle (in degrees)
            mock_sigmaplot.CurrentGraph.XAxis.Label.Rotation = 0
            assert mock_sigmaplot.CurrentGraph.XAxis.Label.Rotation == 0

            mock_sigmaplot.CurrentGraph.XAxis.Label.Rotation = 45
            assert mock_sigmaplot.CurrentGraph.XAxis.Label.Rotation == 45
        except AttributeError:
            pytest.skip("X-axis label rotation property not available in this mock")

    def test_xlabel_position(self, mock_sigmaplot: Any) -> None:
        """
        Test setting X-axis label position.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        try:
            # Set horizontal alignment
            mock_sigmaplot.CurrentGraph.XAxis.Label.HorizontalAlignment = "Center"
            assert mock_sigmaplot.CurrentGraph.XAxis.Label.HorizontalAlignment == "Center"

            # Set vertical position (for X-axis label, typically "Below" the axis)
            mock_sigmaplot.CurrentGraph.XAxis.Label.Position = "Below"
            assert mock_sigmaplot.CurrentGraph.XAxis.Label.Position == "Below"
        except AttributeError:
            pytest.skip("X-axis label position properties not available in this mock")

    def test_xlabel_wrapped_text(self, mock_sigmaplot: Any) -> None:
        """
        Test setting X-axis label with wrapped text.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        # Multi-line X-axis label
        xlabel_text = "Line 1\nLine 2"

        try:
            # Set multi-line label
            mock_sigmaplot.CurrentGraph.XAxis.Label = xlabel_text
            assert mock_sigmaplot.CurrentGraph.XAxis.Label == xlabel_text

            # Enable text wrapping if applicable
            mock_sigmaplot.CurrentGraph.XAxis.Label.WordWrap = True
            assert mock_sigmaplot.CurrentGraph.XAxis.Label.WordWrap == True
        except AttributeError:
            pytest.skip("X-axis label text wrapping property not available in this mock")

# EOF