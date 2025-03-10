#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 05:00:41 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/vba/Test_VBAFileManager.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/tests/vba/Test_VBAFileManager.py"

import os
import tempfile
import unittest
from pathlib import Path
import shutil

from pysigmacro.vba.VBAFileManager import VBAFileManager

class TestVBAFileManager(unittest.TestCase):
    """Test cases for VBAFileManager"""

    def setUp(self):
        # Create a temporary directory for testing
        self.test_dir = tempfile.mkdtemp()
        self.manager = VBAFileManager(self.test_dir)

        # Sample VBA code for testing
        self.test_vba_code = """
Option Explicit

Sub Main()
    MsgBox "Hello from test macro!"
End Sub
"""

    def tearDown(self):
        # Clean up the temporary directory
        shutil.rmtree(self.test_dir)

    def test_init(self):
        """Test initialization creates directory"""
        # Check that the directory exists
        self.assertTrue(os.path.exists(self.test_dir))
        # Check that the macro files dictionary is initialized
        self.assertIsInstance(self.manager.macro_files, dict)

    def test_save_and_get_macro(self):
        """Test saving and retrieving a macro"""
        # Save a test macro
        result = self.manager.save_macro_code("test_macro", self.test_vba_code)
        self.assertTrue(result)

        # Check that the file was created
        expected_path = os.path.join(self.test_dir, "test_macro.bas")
        self.assertTrue(os.path.exists(expected_path))

        # Check that the macro is in the available macros
        self.assertIn("test_macro", self.manager.get_available_macros())

        # Get the macro code and verify it matches
        macro_code = self.manager.get_macro_code("test_macro")
        self.assertEqual(macro_code, self.test_vba_code)

    def test_delete_macro(self):
        """Test deleting a macro"""
        # Save a test macro
        self.manager.save_macro_code("test_delete", self.test_vba_code)

        # Delete the macro
        result = self.manager.delete_macro("test_delete")
        self.assertTrue(result)

        # Check that the macro is no longer available
        self.assertNotIn("test_delete", self.manager.get_available_macros())

        # Check that the file was deleted
        expected_path = os.path.join(self.test_dir, "test_delete.bas")
        self.assertFalse(os.path.exists(expected_path))

    def test_create_temp_macro_file(self):
        """Test creating a temporary macro file"""
        temp_path = self.manager.create_temp_macro_file(self.test_vba_code)

        # Check that the temp file exists
        self.assertTrue(os.path.exists(temp_path))

        # Check that the file contains the correct code
        with open(temp_path, 'r') as f:
            content = f.read()
        self.assertEqual(content, self.test_vba_code)

        # Clean up the temp file
        os.remove(temp_path)

    def test_export_and_import_macros(self):
        """Test exporting and importing macros"""
        # Save a few test macros
        self.manager.save_macro_code("test_export1", self.test_vba_code)
        self.manager.save_macro_code("test_export2", self.test_vba_code)

        # Create a directory for export
        export_dir = tempfile.mkdtemp()

        # Export the macros
        count = self.manager.export_macros_to_directory(export_dir)
        self.assertEqual(count, 2)

        # Create a new manager for import testing
        import_dir = tempfile.mkdtemp()
        import_manager = VBAFileManager(import_dir)

        # Import the macros
        count = import_manager.import_macros_from_directory(export_dir)
        self.assertEqual(count, 2)

        # Check that both macros were imported
        self.assertIn("test_export1", import_manager.get_available_macros())
        self.assertIn("test_export2", import_manager.get_available_macros())

        # Clean up
        shutil.rmtree(export_dir)
        shutil.rmtree(import_dir)

if __name__ == '__main__':
    unittest.main()

# EOF