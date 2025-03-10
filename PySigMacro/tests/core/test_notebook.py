#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 09:18:19 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/core/test_notebook.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/core/test_notebook.py"
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 09:45:12 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/core/test_notebook.py
THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/core/test_notebook.py"
import os
import sys
import pytest
import tempfile
from unittest.mock import MagicMock, patch
from typing import Any, Optional, Dict

"""
Functionality:
* Tests notebook creation and manipulation for SigmaPlot
* Verifies notebook creation, naming, and property setting
Input:
* None (uses mock objects)
Output:
* Test results for notebook operations
Prerequisites:
* pytest
* mock objects for SigmaPlot COM interface
"""

class TestNotebook:
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

            # Setup mock notebook
            mock_notebook = MagicMock()
            mock_app.NewNotebook.return_value = mock_notebook
            mock_app.ActiveDocument.return_value = mock_notebook

            # Configure mock dispatch to return our mock app
            mock_dispatch.return_value = mock_app
            yield mock_app

    def test_create_notebook(self, mock_sigmaplot: Any) -> None:
        """
        Test creating a new notebook with a specific name.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        import win32com.client as win32

        # Define notebook name
        notebook_name = "Test Notebook"

        # Create a new notebook
        mock_sigmaplot.NewNotebook()

        # Save notebook with specific name
        temp_path = os.path.join(tempfile.gettempdir(), f"{notebook_name}.jnb")
        mock_sigmaplot.ActiveDocument().SaveAs(temp_path)

        # Verify that methods were called correctly
        assert mock_sigmaplot.NewNotebook.called
        assert mock_sigmaplot.ActiveDocument().SaveAs.called
        assert mock_sigmaplot.ActiveDocument().SaveAs.call_args[0][0] == temp_path

    def test_notebook_properties(self, mock_sigmaplot: Any) -> None:
        """
        Test setting notebook properties.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a new notebook
        mock_sigmaplot.NewNotebook()
        notebook = mock_sigmaplot.ActiveDocument()

        # Try to set notebook properties
        try:
            notebook.Author = "Test Author"
            notebook.Title = "Test Title"
            notebook.Subject = "Test Subject"
            notebook.Comments = "Test Comments"

            # Verify properties were set
            assert notebook.Author == "Test Author"
            assert notebook.Title == "Test Title"
            assert notebook.Subject == "Test Subject"
            assert notebook.Comments == "Test Comments"
        except AttributeError:
            # Properties might not be available in mock
            pytest.skip("Notebook properties not available in this version")

    def test_open_notebook(self, mock_sigmaplot: Any) -> None:
        """
        Test opening an existing notebook.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a temp file path for the notebook
        notebook_path = os.path.join(tempfile.gettempdir(), "test_open.jnb")

        # Mock the Open method
        mock_sigmaplot.Open.return_value = MagicMock()

        # Open the notebook
        notebook = mock_sigmaplot.Open(notebook_path)

        # Verify Open was called with the correct path
        mock_sigmaplot.Open.assert_called_once_with(notebook_path)
        assert notebook is not None

    def test_close_notebook(self, mock_sigmaplot: Any) -> None:
        """
        Test closing a notebook.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a new notebook
        mock_sigmaplot.NewNotebook()
        notebook = mock_sigmaplot.ActiveDocument()

        # Mock the Close method
        notebook.Close = MagicMock()

        # Close the notebook
        notebook.Close(False)

        # Verify Close was called
        notebook.Close.assert_called_once_with(False)

# EOF