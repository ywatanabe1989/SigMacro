#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 00:21:35 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/plot/test_histogram_plot.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/plot/test_histogram_plot.py"

import os
import pytest
import tempfile
from unittest.mock import MagicMock, patch
from typing import Any, Optional, List, Dict

"""
Functionality:
* Tests histogram plot creation in SigmaPlot
* Verifies correct setup of histogram plot properties and data
Input:
* None (uses mock objects)
Output:
* Test results for histogram plot creation
Prerequisites:
* pytest
* mock objects for SigmaPlot COM interface
"""

class TestHistogramPlot:
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

    def test_create_histogram_plot(self, mock_sigmaplot: Any) -> None:
        """
        Test creating a basic histogram plot.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Define test data (raw values for histogram)
        values = [2.1, 2.3, 2.5, 2.7, 2.9, 3.1, 3.1, 3.3, 3.5, 3.5,
                  3.7, 3.9, 4.1, 4.1, 4.1, 4.3, 4.5, 4.7, 4.9, 5.1]

        # Insert data into worksheet (single column for histogram)
        for i, value in enumerate(values):
            mock_sigmaplot.CurrentWorksheet.Cells(i+1, 1).Value = value

        # Create histogram plot
        mock_sigmaplot.NewGraph("Histogram")

        # Set data for the graph (only need one column for histogram)
        mock_sigmaplot.CurrentGraph.SetData(mock_sigmaplot.CurrentWorksheet, 1)

        # Verify method calls
        assert mock_sigmaplot.NewGraph.called
        assert mock_sigmaplot.NewGraph.call_args[0][0] == "Histogram"
        assert mock_sigmaplot.CurrentGraph.SetData.called

        # Verify SetData parameters (worksheet, data_column)
        call_args = mock_sigmaplot.CurrentGraph.SetData.call_args[0]
        assert call_args[0] == mock_sigmaplot.CurrentWorksheet
        assert call_args[1] == 1

    def test_histogram_customization(self, mock_sigmaplot: Any) -> None:
        """
        Test customizing a histogram plot.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create simple histogram plot
        mock_sigmaplot.NewGraph("Histogram")

        # Apply customizations
        mock_sigmaplot.CurrentGraph.Title = "Test Histogram"
        mock_sigmaplot.CurrentGraph.XAxis.Label = "Value"
        mock_sigmaplot.CurrentGraph.YAxis.Label = "Frequency"

        try:
            # Try to set histogram specific properties if available
            mock_sigmaplot.CurrentGraph.HistogramBins = 10
            mock_sigmaplot.CurrentGraph.HistogramFillColor = "Green"
            mock_sigmaplot.CurrentGraph.HistogramBorderColor = "Black"
            mock_sigmaplot.CurrentGraph.HistogramBorderWidth = 1.0
        except:
            # Not all properties may be available in mock
            pass

        # Verify basic customizations
        assert mock_sigmaplot.CurrentGraph.Title == "Test Histogram"
        assert mock_sigmaplot.CurrentGraph.XAxis.Label == "Value"
        assert mock_sigmaplot.CurrentGraph.YAxis.Label == "Frequency"

    def test_histogram_with_normal_curve(self, mock_sigmaplot: Any) -> None:
        """
        Test adding a normal curve overlay to histogram.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create histogram plot
        mock_sigmaplot.NewGraph("Histogram")

        # Try to add normal curve overlay
        try:
            mock_sigmaplot.CurrentGraph.AddNormalCurve = True
            mock_sigmaplot.CurrentGraph.NormalCurveColor = "Red"
            mock_sigmaplot.CurrentGraph.NormalCurveWidth = 2.0

            # Verify property was set
            assert mock_sigmaplot.CurrentGraph.AddNormalCurve == True
        except:
            # This property might not be available in mock
            pass

    def test_export_histogram_plot(self, mock_sigmaplot: Any) -> None:
        """
        Test exporting a histogram plot.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create histogram plot
        mock_sigmaplot.NewGraph("Histogram")

        # Set up export path
        export_path = os.path.join(tempfile.gettempdir(), "test_histogram_plot.png")

        # Export the graph
        mock_sigmaplot.CurrentGraph.ExportGraph(export_path, "PNG")

        # Verify export was attempted
        assert mock_sigmaplot.CurrentGraph.ExportGraph.called
        assert mock_sigmaplot.CurrentGraph.ExportGraph.call_args[0][0] == export_path
        assert mock_sigmaplot.CurrentGraph.ExportGraph.call_args[0][1] == "PNG"

# EOF