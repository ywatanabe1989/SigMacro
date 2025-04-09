#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 05:00:35 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/vba/Test_VBALibrary.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/vba/Test_VBALibrary.py"

import os
import tempfile
import unittest
from unittest.mock import patch, MagicMock
import shutil

from pysigmacro.vba.VBALibrary import VBALibrary
from pysigmacro.vba.VBAFileManager import VBAFileManager

class TestVBALibrary(unittest.TestCase):
    """Test cases for VBALibrary"""

    def setUp(self):
        # Create a temporary directory for testing
        self.test_dir = tempfile.mkdtemp()

        # Create a library with the test directory
        self.library = VBALibrary(self.test_dir)

    def tearDown(self):
        # Clean up the temporary directory
        shutil.rmtree(self.test_dir)

    def test_init(self):
        """Test initialization creates standard macros"""
        # Verify that standard macros are created
        manager = self.library.vba_manager
        available_macros = manager.get_available_macros()

        # Check that the standard macros exist
        self.assertIn("data_import", available_macros)
        self.assertIn("plotting", available_macros)
        self.assertIn("data_analysis", available_macros)
        self.assertIn("export", available_macros)
        self.assertIn("utility", available_macros)

    def test_get_macro(self):
        """Test getting a macro by name"""
        # Get a macro that should exist
        macro_code = self.library.get_macro("data_import")

        # Verify it contains some VBA code (should start with Option Explicit)
        self.assertIn("Option Explicit", macro_code)

        # Try getting a non-existent macro
        result = self.library.get_macro("nonexistent_macro")
        self.assertIn("not found", result)

    def test_get_all_macro_names(self):
        """Test getting all macro names"""
        names = self.library.get_all_macro_names()

        # Check that all standard macros are included
        self.assertIn("data_import", names)
        self.assertIn("plotting", names)
        self.assertIn("data_analysis", names)
        self.assertIn("export", names)
        self.assertIn("utility", names)

        # Check correct count
        self.assertEqual(len(names), 5)

    @patch('pysigmacro.core.SigmaPlotVBATemplate.SigmaPlotVBATemplate')
    def test_create_template_with_macro(self, mock_template):
        """Test creating a template with a macro"""
        # Setup mock
        mock_instance = MagicMock()
        mock_template.return_value = mock_instance
        mock_instance.create_template_with_macro.return_value = "/path/to/template.JNB"

        # Call the method
        result = self.library.create_template_with_macro("TestTemplate", "data_import")

        # Verify expected result
        self.assertEqual(result, "/path/to/template.JNB")

        # Verify mock was called correctly
        mock_template.assert_called_once()
        mock_instance.create_template_with_macro.assert_called_once()

    @patch('pysigmacro.core.SigmaPlotVBARunner.SigmaPlotVBARunner')
    def test_run_macro(self, mock_runner):
        """Test running a macro"""
        # Setup mock
        mock_instance = MagicMock()
        mock_runner.return_value = mock_instance
        mock_instance.execute_vba_directly.return_value = True

        # Call the method
        result = self.library.run_macro("data_import", ["test_file.csv"])

        # Verify expected result
        self.assertTrue(result)

        # Verify mock was called correctly
        mock_runner.assert_called_once()
        mock_instance.execute_vba_directly.assert_called_once()

if __name__ == '__main__':
    unittest.main()

# EOF