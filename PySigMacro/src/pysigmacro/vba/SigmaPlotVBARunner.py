#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 03:05:32 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/core/SigmaPlotVBARunner.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/core/SigmaPlotVBARunner.py"

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-10 10:00:00 (ywatanabe)"
# File: sigmaplot_vba_runner.py

import os
import sys
import tempfile
import time
import subprocess
import win32com.client
from typing import List, Dict, Any, Optional, Union

class SigmaPlotVBARunner:
    """Class to manage SigmaPlot VBA macro execution"""

    def __init__(self, template_path: Optional[str] = None):
        """
        Initialize SigmaPlot VBA runner

        Args:
            template_path: Path to template JNB file with macros (optional)
        """
        self.sigmaplot_path = self._find_sigmaplot()
        self.template_path = template_path
        self.temp_dir = os.path.join(tempfile.gettempdir(), "sigmaplot_vba")
        os.makedirs(self.temp_dir, exist_ok=True)

    def _find_sigmaplot(self) -> str:
        """Find SigmaPlot executable path"""
        possible_paths = [
            r"C:\Program Files\SigmaPlot\SPW12\Spw.exe",
            r"C:\Program Files (x86)\SigmaPlot\SPW12\Spw.exe",
            r"C:\Program Files\SigmaPlot\SPW14\Spw.exe",
            r"C:\Program Files (x86)\SigmaPlot\SPW14\Spw.exe",
        ]

        for path in possible_paths:
            if os.path.exists(path):
                return path

        # Try to find using Windows registry
        try:
            import winreg
            registry_paths = [
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Systat Software\SigmaPlot"),
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Wow6432Node\Systat Software\SigmaPlot")
            ]
            for reg_root, reg_path in registry_paths:
                try:
                    with winreg.OpenKey(reg_root, reg_path) as key:
                        install_dir, _ = winreg.QueryValueEx(key, "InstallDir")
                        exe_path = os.path.join(install_dir, "Spw.exe")
                        if os.path.exists(exe_path):
                            return exe_path
                except:
                    continue
        except:
            pass

        raise FileNotFoundError("SigmaPlot executable not found")

    def create_vba_macro(self, macro_name: str, macro_code: str) -> str:
        """
        Create a VBA macro file (JNB) with the given code

        Args:
            macro_name: Name for the macro
            macro_code: VBA code for the macro

        Returns:
            Path to the created JNB file
        """
        # TODO: Implement actual JNB creation with VBA
        # For now we'll just save the VBA code to a text file
        macro_file = os.path.join(self.temp_dir, f"{macro_name}.txt")
        with open(macro_file, 'w') as f:
            f.write(macro_code)

        print(f"VBA code saved to {macro_file}")
        print("Note: Creating actual JNB files with VBA requires SigmaPlot COM integration")

        return macro_file

    def run_vba_from_template(self, args: List[str] = None) -> bool:
        """
        Run VBA macro from template JNB file

        Args:
            args: List of arguments to pass to the macro

        Returns:
            True if successful, False otherwise
        """
        if not self.template_path:
            raise ValueError("Template path not specified")

        if not os.path.exists(self.template_path):
            raise FileNotFoundError(f"Template file not found: {self.template_path}")

        # Pass arguments via text file if provided
        if args:
            arg_file = os.path.join(os.path.dirname(self.template_path), "arguments.txt")
            with open(arg_file, 'w') as f:
                f.write(','.join(map(str, args)))

        # Run SigmaPlot with template
        cmd = f'"{self.sigmaplot_path}" "{self.template_path}" /runmacro'

        try:
            # Run SigmaPlot with macro
            proc = subprocess.Popen(cmd, shell=True)

            # Wait for process to complete
            proc.communicate()

            return proc.returncode == 0
        except Exception as e:
            print(f"Error running VBA macro: {e}")
            return False

    def execute_vba_directly(self, macro_code: str, notebook_path: Optional[str] = None) -> bool:
        """
        Execute VBA code directly using COM automation

        Args:
            macro_code: VBA code to execute
            notebook_path: Path to existing notebook to use (optional)

        Returns:
            True if successful, False otherwise
        """
        try:
            # Create SigmaPlot application instance
            sigmaplot = win32com.client.Dispatch('SigmaPlot.Application')
            sigmaplot.Visible = True

            # Open notebook if specified, otherwise create new one
            if notebook_path and os.path.exists(notebook_path):
                # Open existing notebook
                try:
                    notebook = sigmaplot.Notebooks.Open(notebook_path)
                except:
                    # Fallback method if direct open fails
                    cmd = f'"{self.sigmaplot_path}" "{notebook_path}"'
                    subprocess.run(cmd, shell=True)

                    # Find notebook by filename
                    notebook_name = os.path.basename(notebook_path)
                    notebook = None
                    for i in range(sigmaplot.Notebooks.Count):
                        nb = sigmaplot.Notebooks(i)
                        if nb.Name == notebook_name:
                            notebook = nb
                            break
            else:
                # Create new notebook
                sigmaplot.NewNotebook()
                notebook = sigmaplot.ActiveDocument

            # Execute the VBA code directly via COM
            # Note: This may not work for all types of VBA code
            # Complex operations should use template approach instead
            try:
                # Try to execute directly
                result = sigmaplot.ExecuteVBCode(macro_code)
                return True
            except:
                print("Direct VBA execution failed")
                return False

        except Exception as e:
            print(f"Error during COM automation: {e}")
            return False


# Example VBA macro for creating a worksheet with data
example_vba = """
Sub Main
    ' Create a new worksheet
    Worksheet.Create 5, 2

    ' Set column names
    Worksheet.SetColName 1, "X Values"
    Worksheet.SetColName 2, "Y Values"

    ' Add worksheet name
    Worksheet.Name = "Sample Data"

    ' Add data
    For i = 1 To 5
        Worksheet.SetCellValue i, 1, i
        Worksheet.SetCellValue i, 2, i * 2
    Next

    ' Create a simple graph
    Graph.Create 1, 2

    ' Set graph properties
    Graph.Axis.Title.Text = "Sample Graph"
    Graph.Axis.Label(1).Text = "X Axis"
    Graph.Axis.Label(2).Text = "Y Axis"

    ' Save the notebook
    filePath = System.GetSpecialFolderLocation(0) & "\SigmaPlot_Sample.JNB"
    Notebook.SaveAs filePath
End Sub
"""

if __name__ == "__main__":
    # Example usage
    runner = SigmaPlotVBARunner()

    # Save example VBA to file
    vba_file = runner.create_vba_macro("create_worksheet", example_vba)

    # Note: To actually run the VBA, you would need either:
    # 1. A template JNB file with the macro: runner.run_vba_from_template(["arg1", "arg2"])
    # 2. Direct COM execution: runner.execute_vba_directly(example_vba)

# EOF