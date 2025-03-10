#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 04:11:10 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/vba/VBAFileManager.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/vba/VBAFileManager.py"

import os
import tempfile
from typing import Dict, List, Optional

class VBAFileManager:
    """
    Manages VBA files for SigmaPlot macros.

    This class handles loading, saving, and managing VBA macro files
    for SigmaPlot automation.
    """

    def __init__(self, vba_directory: Optional[str] = None):
        """
        Initialize VBA file manager.

        Args:
            vba_directory: Directory where VBA files are stored. If None,
                           a default directory in the package is used.
        """
        if vba_directory is None:
            # Use package directory
            current_dir = os.path.dirname(os.path.abspath(__file__))
            self.vba_directory = os.path.join(current_dir, "vba_macros")
        else:
            self.vba_directory = vba_directory

        # Create directory if it doesn't exist
        os.makedirs(self.vba_directory, exist_ok=True)

        # Dictionary of available macro files
        self.macro_files = self._scan_macro_files()

    def _scan_macro_files(self) -> Dict[str, str]:
        """
        Scan available VBA macro files in the directory.

        Returns:
            Dictionary mapping macro names to file paths
        """
        macro_files = {}
        if os.path.exists(self.vba_directory):
            for filename in os.listdir(self.vba_directory):
                if filename.endswith(('.bas', '.vba', '.txt')):
                    macro_name = os.path.splitext(filename)[0]
                    macro_files[macro_name] = os.path.join(self.vba_directory, filename)
        return macro_files

    def get_macro_code(self, macro_name: str) -> str:
        """
        Get VBA code for a specific macro.

        Args:
            macro_name: Name of the macro to retrieve

        Returns:
            VBA code as a string
        """
        if macro_name in self.macro_files:
            file_path = self.macro_files[macro_name]
            try:
                with open(file_path, 'r') as file:
                    return file.read()
            except Exception as e:
                return f"Error reading macro file: {e}"
        else:
            return f"Macro '{macro_name}' not found. Available macros: {', '.join(self.macro_files.keys())}"

    def save_macro_code(self, macro_name: str, vba_code: str) -> bool:
        """
        Save VBA code to a macro file.

        Args:
            macro_name: Name for the macro
            vba_code: VBA code to save

        Returns:
            True if saved successfully, False otherwise
        """
        try:
            file_path = os.path.join(self.vba_directory, f"{macro_name}.bas")
            with open(file_path, 'w') as file:
                file.write(vba_code)

            # Update macro files dictionary
            self.macro_files[macro_name] = file_path
            return True
        except Exception as e:
            print(f"Error saving macro: {e}")
            return False

    def get_available_macros(self) -> List[str]:
        """
        Get list of available macro names.

        Returns:
            List of macro names
        """
        return list(self.macro_files.keys())

    def delete_macro(self, macro_name: str) -> bool:
        """
        Delete a macro file.

        Args:
            macro_name: Name of the macro to delete

        Returns:
            True if deleted successfully, False otherwise
        """
        if macro_name in self.macro_files:
            try:
                file_path = self.macro_files[macro_name]
                os.remove(file_path)
                del self.macro_files[macro_name]
                return True
            except Exception as e:
                print(f"Error deleting macro: {e}")
                return False
        return False

    def create_temp_macro_file(self, vba_code: str) -> str:
        """
        Create a temporary file with VBA code.

        Args:
            vba_code: VBA code to save

        Returns:
            Path to temporary file
        """
        handle, file_path = tempfile.mkstemp(suffix=".bas", prefix="sigmaplot_macro_")
        os.close(handle)

        with open(file_path, 'w') as file:
            file.write(vba_code)

        return file_path

    def export_macros_to_directory(self, export_dir: str) -> int:
        """
        Export all macros to a directory.

        Args:
            export_dir: Directory to export macros to

        Returns:
            Number of macros exported
        """
        os.makedirs(export_dir, exist_ok=True)
        count = 0

        for macro_name, file_path in self.macro_files.items():
            try:
                with open(file_path, 'r') as source_file:
                    vba_code = source_file.read()

                target_path = os.path.join(export_dir, f"{macro_name}.bas")
                with open(target_path, 'w') as target_file:
                    target_file.write(vba_code)

                count += 1
            except Exception as e:
                print(f"Error exporting macro {macro_name}: {e}")

        return count

    def import_macros_from_directory(self, import_dir: str) -> int:
        """
        Import macros from a directory.

        Args:
            import_dir: Directory to import macros from

        Returns:
            Number of macros imported
        """
        if not os.path.exists(import_dir):
            print(f"Directory not found: {import_dir}")
            return 0

        count = 0
        for filename in os.listdir(import_dir):
            if filename.endswith(('.bas', '.vba', '.txt')):
                try:
                    macro_name = os.path.splitext(filename)[0]
                    source_path = os.path.join(import_dir, filename)

                    with open(source_path, 'r') as source_file:
                        vba_code = source_file.read()

                    if self.save_macro_code(macro_name, vba_code):
                        count += 1
                except Exception as e:
                    print(f"Error importing macro {filename}: {e}")

        return count

if __name__ == "__main__":
    # Test VBA file management
    manager = create_standard_vba_files()
    print(f"Created standard VBA files in: {manager.vba_directory}")
    print(f"Available macros: {manager.get_available_macros()}")

# EOF