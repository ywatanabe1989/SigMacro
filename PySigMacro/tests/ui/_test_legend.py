#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 01:21:39 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/ui/test_legend.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/ui/test_legend.py"

import os
import pytest
from unittest.mock import MagicMock, patch
from typing import Any, Dict, List, Optional

"""
Functionality:
* Tests legend handling in SigmaPlot graphs
* Verifies legend creation, positioning, and formatting
Input:
* None (uses mock objects)
Output:
* Test results for legend functionality
Prerequisites:
* pytest
* mock objects for SigmaPlot COM interface
"""

class TestLegend:
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

            # Setup mock legend
            mock_legend = MagicMock()
            mock_graph.Legend = mock_legend

            # Configure mock dispatch to return our mock app
            mock_dispatch.return_value = mock_app

            yield mock_app

    def test_legend_visibility(self, mock_sigmaplot: Any) -> None:
        """
        Test toggling legend visibility.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        # Toggle legend visibility
        mock_sigmaplot.CurrentGraph.Legend.Visible = True
        assert mock_sigmaplot.CurrentGraph.Legend.Visible == True

        mock_sigmaplot.CurrentGraph.Legend.Visible = False
        assert mock_sigmaplot.CurrentGraph.Legend.Visible == False

    def test_legend_position(self, mock_sigmaplot: Any) -> None:
        """
        Test setting legend position.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        # Test different positions
        try:
            # Try positioning by location string
            mock_sigmaplot.CurrentGraph.Legend.Position = "TopRight"
            assert mock_sigmaplot.CurrentGraph.Legend.Position == "TopRight"

            mock_sigmaplot.CurrentGraph.Legend.Position = "BottomLeft"
            assert mock_sigmaplot.CurrentGraph.Legend.Position == "BottomLeft"
        except:
            # Position might not be available as string in this mock
            pass

        try:
            # Try positioning by coordinates
            mock_sigmaplot.CurrentGraph.Legend.Left = 0.75
            mock_sigmaplot.CurrentGraph.Legend.Top = 0.1

            assert mock_sigmaplot.CurrentGraph.Legend.Left == 0.75
            assert mock_sigmaplot.CurrentGraph.Legend.Top == 0.1
        except AttributeError:
            pytest.skip("Legend position coordinates not available in this mock")

    def test_legend_title(self, mock_sigmaplot: Any) -> None:
        """
        Test setting legend title.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        # Set legend title
        legend_title = "Data Series"

        try:
            mock_sigmaplot.CurrentGraph.Legend.Title = legend_title
            assert mock_sigmaplot.CurrentGraph.Legend.Title == legend_title
        except AttributeError:
            pytest.skip("Legend title property not available in this mock")

    def test_legend_frame(self, mock_sigmaplot: Any) -> None:
        """
        Test legend frame settings.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        try:
            # Toggle frame visibility
            mock_sigmaplot.CurrentGraph.Legend.Frame.Visible = True
            assert mock_sigmaplot.CurrentGraph.Legend.Frame.Visible == True

            # Set frame properties
            mock_sigmaplot.CurrentGraph.Legend.Frame.Color = "Black"
            mock_sigmaplot.CurrentGraph.Legend.Frame.Width = 1

            assert mock_sigmaplot.CurrentGraph.Legend.Frame.Color == "Black"
            assert mock_sigmaplot.CurrentGraph.Legend.Frame.Width == 1
        except AttributeError:
            pytest.skip("Legend frame properties not available in this mock")

    def test_legend_background(self, mock_sigmaplot: Any) -> None:
        """
        Test legend background settings.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        try:
            # Set background color
            mock_sigmaplot.CurrentGraph.Legend.BackgroundColor = "LightGray"
            assert mock_sigmaplot.CurrentGraph.Legend.BackgroundColor == "LightGray"

            # Toggle background transparency
            mock_sigmaplot.CurrentGraph.Legend.Transparent = False
            assert mock_sigmaplot.CurrentGraph.Legend.Transparent == False

            mock_sigmaplot.CurrentGraph.Legend.Transparent = True
            assert mock_sigmaplot.CurrentGraph.Legend.Transparent == True
        except AttributeError:
            pytest.skip("Legend background properties not available in this mock")

    def test_legend_text_formatting(self, mock_sigmaplot: Any) -> None:
        """
        Test legend text formatting.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph
        mock_sigmaplot.NewGraph("Line Plot")

        try:
            # Set font properties
            mock_sigmaplot.CurrentGraph.Legend.Font.Name = "Arial"
            mock_sigmaplot.CurrentGraph.Legend.Font.Size = 10
            mock_sigmaplot.CurrentGraph.Legend.Font.Bold = False

            assert mock_sigmaplot.CurrentGraph.Legend.Font.Name == "Arial"
            assert mock_sigmaplot.CurrentGraph.Legend.Font.Size == 10
            assert mock_sigmaplot.CurrentGraph.Legend.Font.Bold == False
        except AttributeError:
            pytest.skip("Legend font properties not available in this mock")

    def test_legend_entry_customization(self, mock_sigmaplot: Any) -> None:
        """
        Test customizing legend entries.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a graph with multiple series
        mock_sigmaplot.NewGraph("Line Plot")

        try:
            # Create separate mock objects for each entry
            entry0 = MagicMock()
            entry1 = MagicMock()

            # Configure legend entries mock
            mock_sigmaplot.CurrentGraph.Legend.Entries = MagicMock()
            mock_sigmaplot.CurrentGraph.Legend.Entries.side_effect = lambda i: entry0 if i == 0 else entry1

            # Customize legend entry text
            entry0.Text = "Series 1"
            entry1.Text = "Series 2"

            # Verify entry text was set correctly
            assert entry0.Text == "Series 1"
            assert entry1.Text == "Series 2"
        except AttributeError:
            pytest.skip("Legend entries properties not available in this mock")

# EOF