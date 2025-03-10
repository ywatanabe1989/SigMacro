#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 08:44:29 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/core/SigmaPlotVBATemplate.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/core/SigmaPlotVBATemplate.py"

import os
import sys
import tempfile
import subprocess
import win32com.client
from typing import Optional, List, Dict, Any

class SigmaPlotVBATemplate:
    """Create a SigmaPlot notebook with embedded VBA macros"""

    def __init__(self, output_folder: Optional[str] = None):
        """Initialize template creator"""
        self.output_folder = output_folder or os.path.join(os.path.expanduser("~"), "Documents", "SigmaPlot_Templates")
        os.makedirs(self.output_folder, exist_ok=True)
        self.sigmaplot = None

    def connect(self) -> bool:
        """Connect to SigmaPlot application"""
        try:
            # Connect to existing SigmaPlot or create new instance
            try:
                self.sigmaplot = win32com.client.GetActiveObject("SigmaPlot.Application")
            except:
                self.sigmaplot = win32com.client.Dispatch("SigmaPlot.Application")

            # Make SigmaPlot visible
            self.sigmaplot.Visible = True
            return True
        except Exception as e:
            print(f"Error connecting to SigmaPlot: {e}")
            return False

    def create_template_with_macro(self,
                                   template_name: str,
                                   macro_name: str,
                                   vba_code: str) -> Optional[str]:
        """
        Create a SigmaPlot template with embedded VBA macro

        Args:
            template_name: Name for the template notebook
            macro_name: Name for the VBA macro
            vba_code: VBA code to embed

        Returns:
            Path to created JNB file or None if failed
        """
        if not self.connect():
            return None

        try:
            # Create a new notebook
            self.sigmaplot.NewNotebook()

            # Access the notebook
            notebook = self.sigmaplot.ActiveDocument

            # Create a new macro in the notebook
            vba_module = notebook.VBA.Modules.Add(macro_name)

            # Set the macro code
            vba_module.CodeModule.InsertLines(1, vba_code)

            # Save the notebook as a template
            template_path = os.path.join(self.output_folder, f"{template_name}.JNB")
            notebook.SaveAs(template_path)

            print(f"Successfully created template at: {template_path}")
            return template_path

        except Exception as e:
            print(f"Error creating template: {e}")
            return None
        finally:
            # Don't close SigmaPlot - leave it open so user can verify
            pass

    @staticmethod
    def create_sample_macro() -> str:
        """Create a sample VBA macro for testing"""
        return """
Option Explicit

' Main subroutine - entry point for the macro
Sub Main()
    ' Create a new worksheet
    Worksheet.Create 5, 2

    ' Set column names
    Worksheet.SetColName 1, "X Values"
    Worksheet.SetColName 2, "Y Values"

    ' Name the worksheet
    Worksheet.Name = "Sample Data"

    ' Add some sample data
    For i = 1 To 5
        Worksheet.SetCellValue i, 1, i
        Worksheet.SetCellValue i, 2, i * i
    Next i

    ' Create a simple graph
    Graph.Create 1, 2

    ' Set graph properties
    Graph.Axis.Title.Text = "Sample Graph"
    Graph.Axis.Label(1).Text = "X Axis"
    Graph.Axis.Label(2).Text = "Y Axis"

    ' Read arguments if provided
    ReadArguments
End Sub

' Read arguments passed from Python
Sub ReadArguments()
    Dim filePath As String
    Dim fileNum As Integer
    Dim argText As String
    Dim args() As String

    ' Set path to arguments file
    filePath = System.GetSpecialFolderLocation(1) & "\arguments.txt"

    ' Check if file exists
    If System.FileExists(filePath) Then
        ' Open the file
        fileNum = System.OpenFile(filePath, 0)

        ' Read the content
        argText = System.ReadLine(fileNum)

        ' Close the file
        System.CloseFile(fileNum)

        ' Split arguments by comma
        args = Split(argText, ",")

        ' Process arguments
        Dim i As Integer
        For i = 0 To UBound(args)
            ' Do something with each argument
            Debug.Print "Argument " & i & ": " & args(i)
        Next i
    End If
End Sub

' Function that can be called from Main
Function CustomFunction(x As Double) As Double
    CustomFunction = x * x
End Function
"""

def main():
    """Main function to demonstrate template creation"""
    creator = SigmaPlotVBATemplate()
    vba_code = creator.create_sample_macro()

    template_path = creator.create_template_with_macro(
        template_name="Pysigmacro_Template",
        macro_name="SampleMacro",
        vba_code=vba_code
    )

    if template_path:
        print("\nTo use this template from Python:")
        print(f"1. Load the template: {template_path}")
        print("2. Run the macro with: notebook.VBA.Modules('SampleMacro').Run()")
        print("3. Or run with command line: sigmaplot.exe '{template_path}' /runmacro")

if __name__ == "__main__":
    main()

# EOF