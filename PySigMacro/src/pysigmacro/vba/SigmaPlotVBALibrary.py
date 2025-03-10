#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 03:38:09 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/core/SigmaPlotVBALibrary.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/core/SigmaPlotVBALibrary.py"

import os
import tempfile
import win32com.client
from typing import Dict, List, Optional

class SigmaPlotVBALibrary:
    """
    A library of useful SigmaPlot VBA macros that can be embedded into JNB files
    or executed directly.
    """
    @staticmethod
    def get_data_import_macro() -> str:
        """
        Returns a VBA macro for importing data from various file formats.
        """
        return """
Option Explicit

' Main entry point
Sub Main()
    ' Read arguments if provided
    Dim args() As String
    args = ReadArguments()

    ' Process based on arguments or use defaults
    If UBound(args) >= 0 Then
        ImportDataFile args(0)
    Else
        ImportDataFile ""
    End If
End Sub

' Read arguments passed from Python
Function ReadArguments() As String()
    Dim filePath As String
    Dim fileNum As Integer
    Dim argText As String
    Dim args() As String

    ' Default empty array
    ReDim args(-1 To -1)

    ' Set path to arguments file
    filePath = System.GetSpecialFolderLocation(1) & "\\arguments.txt"

    ' Check if file exists
    If System.FileExists(filePath) Then
        ' Open the file
        fileNum = System.OpenFile(filePath, 0)

        ' Read the content
        argText = System.ReadLine(fileNum)

        ' Close the file
        System.CloseFile(fileNum)

        ' Split arguments by comma
        If Len(argText) > 0 Then
            args = Split(argText, ",")
        End If
    End If

    ReadArguments = args
End Function

' Import data file based on extension
Sub ImportDataFile(filePath As String)
    Dim fileType As String

    ' If no file path provided, ask user
    If filePath = "" Then
        filePath = System.GetOpenFileName("All Files (*.*)|*.*|CSV Files (*.csv)|*.csv|Excel Files (*.xlsx)|*.xlsx", "", "Select Data File")
        If filePath = "" Then
            Exit Sub
        End If
    End If

    ' Get file extension
    fileType = LCase(Right(filePath, 4))

    ' Create a new worksheet
    Worksheet.Create(100, 10)

    ' Import based on file type
    If fileType = ".csv" Then
        ImportCSV filePath
    ElseIf fileType = "xlsx" Then
        ImportExcel filePath
    ElseIf fileType = ".txt" Then
        ImportText filePath
    Else
        ' Try as CSV by default
        ImportCSV filePath
    End If
End Sub

' Import CSV file
Sub ImportCSV(filePath As String)
    ' Set import options for CSV
    Notebook.Import.TextOptions.Delimiter = ","
    Notebook.Import.TextOptions.HeaderRow = 1

    ' Import the file
    Notebook.Import.Text filePath
End Sub

' Import Excel file
Sub ImportExcel(filePath As String)
    ' Import the Excel file
    Notebook.Import.Excel filePath
End Sub

' Import text file
Sub ImportText(filePath As String)
    ' Set import options for text
    Notebook.Import.TextOptions.Delimiter = " "
    Notebook.Import.TextOptions.HeaderRow = 0

    ' Import the file
    Notebook.Import.Text filePath
End Sub
"""

    @staticmethod
    def get_plotting_macro() -> str:
        """
        Returns a VBA macro for creating various types of plots.
        """
        # Similar to get_data_import_macro, include the VBA code here
        return """
Option Explicit

' Main entry point
Sub Main()
    ' ... VBA code for plotting macro ...
End Sub
"""
    @staticmethod
    def get_data_analysis_macro() -> str:
        """
        Returns a VBA macro for performing basic statistical analysis.
        """
        # Include the VBA code for data analysis
        return """
Option Explicit

' Main entry point
Sub Main()
    ' ... VBA code for data analysis macro ...
End Sub
"""

    @staticmethod
    def get_export_macro() -> str:
        """
        Returns a VBA macro for exporting graphs and data.
        """
        # Include the VBA code for export macro
        return """
Option Explicit

' Main entry point
Sub Main()
    ' ... VBA code for export macro ...
End Sub
"""

    @staticmethod
    def get_utility_macro() -> str:
        """
        Returns a utility VBA macro with helper functions.
        """
        # Corrected VBA code with proper syntax and function definitions
        return """
Option Explicit

' Main entry point that demonstrates utility functions
Sub Main()
    ' Read arguments if provided
    Dim args() As String
    args = ReadArguments()

    ' Create a debug report
    CreateDebugReport

    ' Process based on arguments
    If UBound(args) >= 0 Then
        ShowNotebookInfo args(0)
    Else
        ShowNotebookInfo ""
    End If
End Sub

' Read arguments passed from Python
Function ReadArguments() As String()
    Dim filePath As String
    Dim fileNum As Integer
    Dim argText As String
    Dim args() As String

    ' Default empty array
    ReDim args(-1 To -1)

    ' Set path to arguments file
    filePath = System.GetSpecialFolderLocation(1) & "\\arguments.txt"

    ' Check if file exists
    If System.FileExists(filePath) Then
        ' Open the file
        fileNum = System.OpenFile(filePath, 0)

        ' Read the content
        argText = System.ReadLine(fileNum)

        ' Close the file
        System.CloseFile(fileNum)

        ' Split arguments by comma
        If Len(argText) > 0 Then
            args = Split(argText, ",")
        End If
    End If

    ReadArguments = args
End Function

' Create a debug report worksheet
Sub CreateDebugReport()
    ' Create a new worksheet for the report
    Worksheet.Create 100, 3

    ' Set column headers
    Worksheet.SetCellValue 1, 1, "Item"
    Worksheet.SetCellValue 1, 2, "Value"
    Worksheet.SetCellValue 1, 3, "Description"

    ' Get system information
    Dim row As Integer
    row = 2

    ' SigmaPlot version
    Worksheet.SetCellValue row, 1, "SigmaPlot Version"
    Worksheet.SetCellValue row, 2, Application.Version
    Worksheet.SetCellValue row, 3, "Current SigmaPlot version"
    row = row + 1

    ' Operating system
    Worksheet.SetCellValue row, 1, "Operating System"
    Worksheet.SetCellValue row, 2, System.GetOSVersion
    Worksheet.SetCellValue row, 3, "Current OS version"
    row = row + 1

    ' Current date and time
    Worksheet.SetCellValue row, 1, "Date/Time"
    Worksheet.SetCellValue row, 2, Now
    Worksheet.SetCellValue row, 3, "Current date and time"
    row = row + 1

    ' Username
    Worksheet.SetCellValue row, 1, "Username"
    Worksheet.SetCellValue row, 2, System.GetUserName
    Worksheet.SetCellValue row, 3, "Current user name"
    row = row + 1

    ' Computer name
    Worksheet.SetCellValue row, 1, "Computer Name"
    Worksheet.SetCellValue row, 2, System.GetComputerName
    Worksheet.SetCellValue row, 3, "Current computer name"
    row = row + 1

    ' Documents folder path
    Worksheet.SetCellValue row, 1, "Documents Folder"
    Worksheet.SetCellValue row, 2, System.GetSpecialFolderLocation(0)
    Worksheet.SetCellValue row, 3, "Path to Documents folder"
    row = row + 1

    ' Temp folder path
    Worksheet.SetCellValue row, 1, "Temp Folder"
    Worksheet.SetCellValue row, 2, System.GetSpecialFolderLocation(2)
    Worksheet.SetCellValue row, 3, "Path to temporary files folder"
    row = row + 1

    ' Notebook information
    Worksheet.SetCellValue row, 1, "Notebook Path"
    Worksheet.SetCellValue row, 2, Notebook.Path
    Worksheet.SetCellValue row, 3, "Path to current notebook"
    row = row + 1

    Worksheet.SetCellValue row, 1, "Notebook Name"
    Worksheet.SetCellValue row, 2, Notebook.Name
    Worksheet.SetCellValue row, 3, "Name of current notebook"
    row = row + 1

    ' Active item information
    Worksheet.SetCellValue row, 1, "Active Item"
    Worksheet.SetCellValue row, 2, Notebook.GetActiveItemType
    Worksheet.SetCellValue row, 3, "Type of active item"
    row = row + 1

    ' Name the debug report
    Worksheet.Name = "System Info"
End Sub

' Show detailed information about notebook contents
Sub ShowNotebookInfo(Optional notebookPath As String = "")
    ' Open the specified notebook if provided
    If notebookPath <> "" Then
        Notebook.Open notebookPath
    End If

    ' Create a new worksheet for notebook info
    Worksheet.Create 100, 4

    ' Set column headers
    Worksheet.SetCellValue 1, 1, "Item #"
    Worksheet.SetCellValue 1, 2, "Type"
    Worksheet.SetCellValue 1, 3, "Name"
    Worksheet.SetCellValue 1, 4, "Details"

    ' ... Rest of the VBA code ...
End Sub
"""

    def __init__(self):
        """Initialize SigmaPlotVBALibrary with a catalog of available macros"""
        self.available_macros = {
            "data_import": self.get_data_import_macro,
            "plotting": self.get_plotting_macro,
            "data_analysis": self.get_data_analysis_macro,
            "export": self.get_export_macro,
            "utility": self.get_utility_macro
        }

    def get_macro(self, macro_name: str) -> str:
        """
        Get a specific macro by name

        Args:
            macro_name: Name of the macro to retrieve

        Returns:
            VBA code for the requested macro or an error message
        """
        if macro_name in self.available_macros:
            return self.available_macros[macro_name]()
        else:
            return (f"Macro '{macro_name}' not found. "
                    f"Available macros: {', '.join(self.available_macros.keys())}")

    def get_all_macro_names(self) -> List[str]:
        """
        Get names of all available macros

        Returns:
            List of macro names
        """
        return list(self.available_macros.keys())

    def create_template_with_macro(self, template_name: str, macro_name: str,
                                   output_path: Optional[str] = None) -> Optional[str]:
        """
        Create a SigmaPlot template notebook with the specified macro embedded

        Args:
            template_name: Name for the template
            macro_name: Name of the macro to include
            output_path: Path to save the template (optional)

        Returns:
            Path to the created template file or None if failed
        """
        try:
            from pysigmacro.vba.SigmaPlotVBATemplate import SigmaPlotVBATemplate

            vba_code = self.get_macro(macro_name)
            if vba_code.startswith("Macro"):
                # Error message returned
                print(vba_code)
                return None

            creator = SigmaPlotVBATemplate(
                output_folder=os.path.dirname(output_path) if output_path else None
            )
            return creator.create_template_with_macro(
                template_name=template_name,
                macro_name=macro_name,
                vba_code=vba_code
            )
        except ImportError as e:
            print(f"Error importing SigmaPlotVBATemplate: {e}")
            return None

    def run_macro(self, macro_name: str, args: List[str] = None,
                  notebook_path: Optional[str] = None) -> bool:
        """
        Run a macro in SigmaPlot

        Args:
            macro_name: Name of macro to run
            args: List of arguments to pass to macro
            notebook_path: Path to an existing notebook (if None, creates temp notebook)

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

            # If notebook_path is provided, run from template; otherwise, execute directly
            runner = SigmaPlotVBARunner(template_path=notebook_path)

            if notebook_path:
                return runner.run_vba_from_template(args)
            else:
                return runner.execute_vba_directly(vba_code)
        except ImportError as e:
            print(f"Error importing SigmaPlotVBARunner: {e}")
            return False

# EOF