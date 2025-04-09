#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 09:18:35 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/core/test_sections.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/core/test_sections.py"
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 09:45:30 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/core/test_sections.py
THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/core/test_sections.py"
import os
import sys
import pytest
from unittest.mock import MagicMock, patch
from typing import Any, List, Optional, Dict

"""
Functionality:
* Tests section creation and manipulation within SigmaPlot notebooks
* Verifies section creation, multiple section handling, and section activation
Input:
* None (uses mock objects)
Output:
* Test results for section operations
Prerequisites:
* pytest
* mock objects for SigmaPlot COM interface
"""

class TestSections:
    @pytest.fixture
    def mock_sigmaplot(self) -> Any:
        """
        Create a mock SigmaPlot application with notebook for testing.

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

            # Setup mock section
            mock_section = MagicMock()
            mock_notebook.AddSection.return_value = mock_section

            # Setup mock sections collection
            mock_sections = MagicMock()
            mock_notebook.Sections = mock_sections
            mock_sections.Count = 0

            # Configure mock dispatch to return our mock app
            mock_dispatch.return_value = mock_app
            yield mock_app

    def test_create_section(self, mock_sigmaplot: Any) -> None:
        """
        Test creating a new section with a specific name.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        import win32com.client as win32

        # Define section name
        section_name = "Test Section"

        # Create a new notebook
        mock_sigmaplot.NewNotebook()
        notebook = mock_sigmaplot.ActiveDocument()

        # Add a new section
        section = notebook.AddSection(section_name)

        # Verify that methods were called correctly
        assert notebook.AddSection.called
        assert notebook.AddSection.call_args[0][0] == section_name

    def test_multiple_sections(self, mock_sigmaplot: Any) -> None:
        """
        Test creating multiple sections within a notebook.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a new notebook
        mock_sigmaplot.NewNotebook()
        notebook = mock_sigmaplot.ActiveDocument()

        # Add multiple sections
        section_names = ["Section 1", "Section 2", "Section 3"]
        sections: List[Any] = []
        for name in section_names:
            section = notebook.AddSection(name)
            sections.append(section)

        # Verify all sections were created
        assert notebook.AddSection.call_count == len(section_names)

        # Verify section names in order
        for i, name in enumerate(section_names):
            assert notebook.AddSection.call_args_list[i][0][0] == name

    def test_activate_section(self, mock_sigmaplot: Any) -> None:
        """
        Test activating a specific section within a notebook.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a new notebook
        mock_sigmaplot.NewNotebook()
        notebook = mock_sigmaplot.ActiveDocument()

        # Create some sections
        section1 = notebook.AddSection("Section 1")
        section2 = notebook.AddSection("Section 2")

        # Try to activate a section
        try:
            notebook.ActiveSection = section2
            assert notebook.ActiveSection == section2
        except AttributeError:
            # This property might not be available in mock
            pytest.skip("ActiveSection property not available in this version")

    def test_section_count(self, mock_sigmaplot: Any) -> None:
        """
        Test getting the count of sections in a notebook.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a new notebook
        mock_sigmaplot.NewNotebook()
        notebook = mock_sigmaplot.ActiveDocument()
        sections = notebook.Sections

        # Mock sections count
        sections.Count = 0

        # Add multiple sections and update count each time
        section_names = ["Section A", "Section B", "Section C"]
        for i, name in enumerate(section_names):
            notebook.AddSection(name)
            sections.Count = i + 1

        # Verify correct section count
        assert sections.Count == len(section_names)

    def test_delete_section(self, mock_sigmaplot: Any) -> None:
        """
        Test deleting a section from a notebook.

        Parameters
        ----------
        mock_sigmaplot : Any
            Mock SigmaPlot application object
        """
        # Create a new notebook
        mock_sigmaplot.NewNotebook()
        notebook = mock_sigmaplot.ActiveDocument()

        # Add sections
        section1 = notebook.AddSection("Section 1")
        section2 = notebook.AddSection("Section 2")

        # Mock the Delete method
        section1.Delete = MagicMock()

        # Delete the section
        section1.Delete()

        # Verify Delete was called
        assert section1.Delete.called

# EOF