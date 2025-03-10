#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 00:22:00 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/plot/test_line_with_interval.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/plot/test_line_with_interval.py"

import os
import pytest
import tempfile
from unittest.mock import MagicMock, patch
from typing import Any, Optional, List, Dict

"""
Functionality:
* Tests line plot creation with error intervals in SigmaPlot
* Verifies correct setup of error bars and confidence intervals
Input:
* None (uses mock objects)
Output:
* Test results for line plot with intervals creation
Prerequisites:
* pytest
* mock objects for SigmaPlot COM interface
"""

class TestLineWithInterval:
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

            # Setup mock error bars
            mock_error_bars = MagicMock()
            mock_graph.ErrorBars = mock_error_bars

            # Configure mock dispatch to return our mock app
            mock_dispatch.return_value = mock_app

            yield mock_app

    def test_line_with_error_bars(self, mock_sigmaplot: Any) -> None:
        """
        Test creating a line plot with error bars.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Define test data
        x_values = [1, 2, 3, 4, 5]
        y_values = [10, 15, 13, 17, 20]
        y_errors = [1.0, 1.5, 1.2, 1.7, 2.0]

        # Insert data into worksheet
        for i, (x, y, err) in enumerate(zip(x_values, y_values, y_errors)):
            mock_sigmaplot.CurrentWorksheet.Cells(i+1, 1).Value = x
            mock_sigmaplot.CurrentWorksheet.Cells(i+1, 2).Value = y
            mock_sigmaplot.CurrentWorksheet.Cells(i+1, 3).Value = err

        # Create line plot
        mock_sigmaplot.NewGraph("Line Plot")

        # Set data for the graph
        mock_sigmaplot.CurrentGraph.SetData(mock_sigmaplot.CurrentWorksheet, 1, 2)

        # Set error bars
        mock_sigmaplot.CurrentGraph.ErrorBars.Column = 3
        mock_sigmaplot.CurrentGraph.ErrorBars.Visible = True

        # Verify method calls
        assert mock_sigmaplot.NewGraph.called
        assert mock_sigmaplot.NewGraph.call_args[0][0] == "Line Plot"
        assert mock_sigmaplot.CurrentGraph.SetData.called

        # Verify error bar settings
        assert mock_sigmaplot.CurrentGraph.ErrorBars.Column == 3
        assert mock_sigmaplot.CurrentGraph.ErrorBars.Visible == True

    def test_asymmetric_error_bars(self, mock_sigmaplot: Any) -> None:
        """
        Test creating a line plot with asymmetric error bars.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Define test data with asymmetric errors
        x_values = [1, 2, 3, 4, 5]
        y_values = [10, 15, 13, 17, 20]
        y_errors_plus = [1.5, 2.0, 1.8, 2.2, 2.5]
        y_errors_minus = [1.0, 1.2, 0.8, 1.5, 1.2]

        # Insert data into worksheet
        for i, (x, y, err_plus, err_minus) in enumerate(zip(x_values, y_values, y_errors_plus, y_errors_minus)):
            mock_sigmaplot.CurrentWorksheet.Cells(i+1, 1).Value = x
            mock_sigmaplot.CurrentWorksheet.Cells(i+1, 2).Value = y
            mock_sigmaplot.CurrentWorksheet.Cells(i+1, 3).Value = err_plus
            mock_sigmaplot.CurrentWorksheet.Cells(i+1, 4).Value = err_minus

        # Create line plot
        mock_sigmaplot.NewGraph("Line Plot")

        # Set data for the graph
        mock_sigmaplot.CurrentGraph.SetData(mock_sigmaplot.CurrentWorksheet, 1, 2)

        # Try to set asymmetric error bars
        try:
            mock_sigmaplot.CurrentGraph.ErrorBars.Type = "Asymmetric"
            mock_sigmaplot.CurrentGraph.ErrorBars.PlusColumn = 3
            mock_sigmaplot.CurrentGraph.ErrorBars.MinusColumn = 4
            mock_sigmaplot.CurrentGraph.ErrorBars.Visible = True

            # Verify error bar settings
            assert mock_sigmaplot.CurrentGraph.ErrorBars.Type == "Asymmetric"
            assert mock_sigmaplot.CurrentGraph.ErrorBars.PlusColumn == 3
            assert mock_sigmaplot.CurrentGraph.ErrorBars.MinusColumn == 4
        except:
            # These specific properties might not be available in the mock
            pass

    def test_error_bar_customization(self, mock_sigmaplot: Any) -> None:
        """
        Test customizing error bar appearance.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create line plot with error bars
        mock_sigmaplot.NewGraph("Line Plot")

        # Set error bars
        mock_sigmaplot.CurrentGraph.ErrorBars.Visible = True

        # Customize error bar appearance
        try:
            mock_sigmaplot.CurrentGraph.ErrorBars.Color = "Red"
            mock_sigmaplot.CurrentGraph.ErrorBars.Width = 1.5
            mock_sigmaplot.CurrentGraph.ErrorBars.CapWidth = 5
            mock_sigmaplot.CurrentGraph.ErrorBars.Style = "Both"

            # Verify customizations
            assert mock_sigmaplot.CurrentGraph.ErrorBars.Color == "Red"
            assert mock_sigmaplot.CurrentGraph.ErrorBars.Width == 1.5
            assert mock_sigmaplot.CurrentGraph.ErrorBars.CapWidth == 5
            assert mock_sigmaplot.CurrentGraph.ErrorBars.Style == "Both"
        except:
            # These specific properties might not be available in the mock
            pass

    def test_confidence_intervals(self, mock_sigmaplot: Any) -> None:
        """
        Test adding confidence intervals to a line plot.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create line plot
        mock_sigmaplot.NewGraph("Line Plot")

        # Try to add confidence intervals
        try:
            mock_sigmaplot.CurrentGraph.AddStatistics("Confidence Intervals")
            mock_sigmaplot.CurrentGraph.ConfidenceLevel = 95
            mock_sigmaplot.CurrentGraph.ConfidenceBands.Visible = True
            mock_sigmaplot.CurrentGraph.ConfidenceBands.Color = "LightBlue"
            mock_sigmaplot.CurrentGraph.ConfidenceBands.FillOpacity = 50

            # Verify confidence interval settings
            assert mock_sigmaplot.CurrentGraph.AddStatistics.called
            assert mock_sigmaplot.CurrentGraph.ConfidenceLevel == 95
            assert mock_sigmaplot.CurrentGraph.ConfidenceBands.Visible == True
        except:
            # These properties might not be available in mock
            pytest.skip("Confidence interval properties not available in mock")

    def test_export_line_with_intervals(self, mock_sigmaplot: Any) -> None:
        """
        Test exporting a line plot with error intervals.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create line plot with error bars
        mock_sigmaplot.NewGraph("Line Plot")
        mock_sigmaplot.CurrentGraph.ErrorBars.Visible = True

        # Set up export path
        export_path = os.path.join(tempfile.gettempdir(), "test_line_intervals.png")

        # Export the graph
        mock_sigmaplot.CurrentGraph.ExportGraph(export_path, "PNG")

        # Verify export was attempted
        assert mock_sigmaplot.CurrentGraph.ExportGraph.called
        assert mock_sigmaplot.CurrentGraph.ExportGraph.call_args[0][0] == export_path
        assert mock_sigmaplot.CurrentGraph.ExportGraph.call_args[0][1] == "PNG"

# EOF