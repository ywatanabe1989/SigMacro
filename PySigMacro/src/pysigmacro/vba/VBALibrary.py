#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 11:32:40 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/vba/VBALibrary.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/vba/VBALibrary.py"
import os

THIS_DIR = os.path.dirname(THIS_FILE)


import tempfile
from typing import List, Optional

from pysigmacro.utils.load_text import load_text
from pysigmacro.vba.VBAFileManager import VBAFileManager


class VBALibrary:
    """
    A library of useful SigmaPlot VBA macros that can be embedded into JNB files
    or executed directly.

    This implementation uses separate VBA files for better maintainability.
    """

    def __init__(self, vba_directory: Optional[str] = None):
        """
        Initialize VBALibrary with a catalog of available macros

        Args:
            vba_directory: Optional custom directory for VBA files
        """
        # Initialize VBA file manager
        self.vba_manager = VBAFileManager(vba_directory)

        # Available macro mapping - this connects macro names to their file names
        self.available_macros = {
            "data_import": "data_import",
            "plotting": "plotting",
            "data_analysis": "data_analysis",
            "export": "export",
            "utility": "utility",
        }

        # Ensure standard macros exist
        self._ensure_standard_macros()

    def _ensure_standard_macros(self):
        """
        Ensure standard macros exist in the VBA directory.
        If they don't exist, create them with default content.
        """
        # Check which macros need to be created
        missing_macros = []
        for macro_name in self.available_macros.values():
            if macro_name not in self.vba_manager.get_available_macros():
                missing_macros.append(macro_name)

        # If any are missing, create them
        if missing_macros:
            # Create standard VBA files
            self._create_standard_vba_files()

    def _create_standard_vba_files(self):
        """Create standard VBA files with default macro content"""
        # Create data import macro if missing
        if "data_import" not in self.vba_manager.get_available_macros():
            self.vba_manager.save_macro_code(
                "data_import", self._get_default_data_import_macro()
            )

        # Create plotting macro if missing
        if "plotting" not in self.vba_manager.get_available_macros():
            self.vba_manager.save_macro_code(
                "plotting", self._get_default_plotting_macro()
            )

        # Create data analysis macro if missing
        if "data_analysis" not in self.vba_manager.get_available_macros():
            self.vba_manager.save_macro_code(
                "data_analysis", self._get_default_data_analysis_macro()
            )

        # Create export macro if missing
        if "export" not in self.vba_manager.get_available_macros():
            self.vba_manager.save_macro_code(
                "export", self._get_default_export_macro()
            )

        # Create utility macro if missing
        if "utility" not in self.vba_manager.get_available_macros():
            self.vba_manager.save_macro_code(
                "utility", self._get_default_utility_macro()
            )

    def get_macro(self, macro_name: str) -> str:
        """
        Get a specific macro by name

        Args:
            macro_name: Name of the macro to retrieve

        Returns:
            VBA code for the requested macro or error message
        """
        # Check if macro exists in mapping
        if macro_name in self.available_macros:
            file_name = self.available_macros[macro_name]
            # Get the macro code using the VBA file manager
            return self.vba_manager.get_macro_code(file_name)
        else:
            return f"Macro '{macro_name}' not found. Available macros: {', '.join(self.available_macros.keys())}"

    def get_all_macro_names(self) -> List[str]:
        """
        Get names of all available macros

        Returns:
            List of macro names
        """
        return list(self.available_macros.keys())

    def create_template_with_macro(
        self,
        template_name: str,
        macro_name: str,
        output_path: Optional[str] = None,
    ) -> Optional[str]:
        """
        Create a SigmaPlot template notebook with the specified macro embedded

        Args:
            template_name: Name for the template
            macro_name: Name of macro to include
            output_path: Path to save the template (optional)

        Returns:
            Path to created template file or None if failed
        """
        try:
            from pysigmacro.vba.SigmaPlotVBATemplate import (
                SigmaPlotVBATemplate,
            )

            vba_code = self.get_macro(macro_name)
            if vba_code.startswith("Macro"):
                # Error message returned
                print(vba_code)
                return None

            creator = SigmaPlotVBATemplate(
                output_folder=(
                    os.path.dirname(output_path) if output_path else None
                )
            )
            return creator.create_template_with_macro(
                template_name=template_name,
                macro_name=macro_name,
                vba_code=vba_code,
            )
        except ImportError as e:
            print(f"Error importing SigmaPlotVBATemplate: {e}")
            return None

    def run_macro(
        self,
        macro_name: str,
        args: List[str] = None,
        notebook_path: Optional[str] = None,
    ) -> bool:
        """
        Run a macro in SigmaPlot

        Args:
            macro_name: Name of macro to run
            args: List of arguments to pass to macro
            notebook_path: Path to existing notebook (if None, creates temp notebook)

        Returns:
            True if successful, False otherwise
        """
        try:
            from pysigmacro.vba.SigmaPlotVBARunner import SigmaPlotVBARunner

            vba_code = self.get_macro(macro_name)
            if vba_code.startswith("Macro"):
                # Error message returned
                print(vba_code)
                return False

            # If notebook_path provided, run from template, otherwise execute directly
            runner = SigmaPlotVBARunner(template_path=notebook_path)

            if notebook_path:
                return runner.run_vba_from_template(args)
            else:
                return runner.execute_vba_directly(vba_code)
        except ImportError as e:
            print(f"Error importing SigmaPlotVBARunner: {e}")
            return False

    def _get_default_data_import_macro(self) -> str:
        """Returns default VBA code for data import macro"""
        return load_text(
            os.path.join(THIS_DIR, "./default_data_importing.vba")
        )

    def _get_default_plotting_macro(self) -> str:
        """Returns default VBA code for plotting macro"""
        return load_text(os.path.join(THIS_DIR, "./default_plotting.vba"))

    def _get_default_data_analysis_macro(self) -> str:
        """Returns default VBA code for data analysis macro"""
        return load_text(os.path.join(THIS_DIR, "./default_data_analysis.vba"))

    def _get_default_export_macro(self) -> str:
        """Returns default VBA code for export macro"""
        return load_text(os.path.join(THIS_DIR, "./default_exporting.vba"))

    def _get_default_utility_macro(self) -> str:
        """Returns default VBA code for utility macro"""
        return load_text(os.path.join(THIS_DIR, "./default_utility.vba"))


if __name__ == "__main__":
    # Test the VBA file manager with SigmaPlot library
    library = VBALibrary()

    print("Available macros:", library.get_all_macro_names())

    # Create a template with one of the macros
    template_path = library.create_template_with_macro(
        template_name="SigmaPlot_DataImport",
        macro_name="data_import",
        output_path=os.path.join(
            tempfile.gettempdir(), "SigmaPlot_DataImport.JNB"
        ),
    )

    if template_path:
        print(f"Template created at: {template_path}")

    # Example of running a macro directly
    success = library.run_macro("plotting", ["line", "1"])
    print(f"Macro execution {'successful' if success else 'failed'}")

# EOF