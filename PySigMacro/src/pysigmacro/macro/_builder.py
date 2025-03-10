#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 01:47:03 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/macro/builder.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/macro/builder.py"

import os
import tempfile
import time
import subprocess
import threading
from typing import List, Dict, Any, Union, Optional, Tuple

class MacroBuilder:
    """
    Builder class for SigmaPlot macros that provides a more Pythonic way to construct macros
    programmatically instead of using multi-line strings.
    """

    def __init__(self, macro_name: Optional[str] = None):
        """
        Initialize a new macro builder.

        Args:
            macro_name: Optional name for the macro. Used in macro header.
        """
        self.lines: List[str] = []
        self.main_function_defined = False
        self.indent_level = 0
        self.macro_name = macro_name or "PysigmacroMacro"

        # Start with standard macro header
        self._add_header()

    def _add_header(self) -> None:
        """Add standard macro header."""
        self.lines.append("[Main]")
        self.lines.append("Function=Main")
        self.lines.append("")
        self.lines.append("[Main]")
        self.main_function_defined = True

    def line(self, code: str) -> 'MacroBuilder':
        """
        Add a line of code to the macro with proper indentation.

        Args:
            code: The line of SigmaPlot macro code to add

        Returns:
            self for method chaining
        """
        if not code.strip():
            self.lines.append("")
        else:
            # Apply indentation
            indent = "    " * self.indent_level
            self.lines.append(f"{indent}{code}")
        return self

    def comment(self, text: str) -> 'MacroBuilder':
        """
        Add a comment line to the macro.

        Args:
            text: Comment text (without the semicolon)

        Returns:
            self for method chaining
        """
        indent = "    " * self.indent_level
        self.lines.append(f"{indent}; {text}")
        return self

    def blank_line(self) -> 'MacroBuilder':
        """
        Add a blank line for readability.

        Returns:
            self for method chaining
        """
        self.lines.append("")
        return self

    def begin_block(self) -> 'MacroBuilder':
        """
        Begin a new indented block.

        Returns:
            self for method chaining
        """
        self.indent_level += 1
        return self

    def end_block(self) -> 'MacroBuilder':
        """
        End the current indented block.

        Returns:
            self for method chaining
        """
        if self.indent_level > 0:
            self.indent_level -= 1
        return self

    def add_variable(self, name: str, value: Any) -> 'MacroBuilder':
        """
        Add a variable declaration to the macro.

        Args:
            name: Variable name
            value: Variable value (will be converted to appropriate syntax)

        Returns:
            self for method chaining
        """
        if isinstance(value, str):
            # String values need quotes
            self.line(f"{name} = \"{value}\"")
        elif isinstance(value, bool):
            # Boolean values are True/False
            self.line(f"{name} = {str(value).lower()}")
        else:
            # Numbers and other values
            self.line(f"{name} = {value}")
        return self

    def add_for_loop(self, var: str, start: int, end: int) -> 'MacroBuilder':
        """
        Start a for loop.

        Args:
            var: Loop variable name
            start: Start value
            end: End value (inclusive)

        Returns:
            self for method chaining
        """
        self.line(f"For {var} = {start} To {end}")
        self.begin_block()
        return self

    def end_for_loop(self) -> 'MacroBuilder':
        """
        End a for loop.

        Returns:
            self for method chaining
        """
        self.end_block()
        self.line("Next")
        return self

    def add_if_statement(self, condition: str) -> 'MacroBuilder':
        """
        Start an if statement.

        Args:
            condition: The condition expression

        Returns:
            self for method chaining
        """
        self.line(f"If {condition} Then")
        self.begin_block()
        return self

    def add_else(self) -> 'MacroBuilder':
        """
        Add an else clause.

        Returns:
            self for method chaining
        """
        self.end_block()
        self.line("Else")
        self.begin_block()
        return self

    def end_if(self) -> 'MacroBuilder':
        """
        End an if statement.

        Returns:
            self for method chaining
        """
        self.end_block()
        self.line("End If")
        return self

    def add_error_handler(self) -> 'MacroBuilder':
        """
        Add standard error handling.

        Returns:
            self for method chaining
        """
        self.line("On Error Resume Next")
        return self

    def create_notebook(self) -> 'MacroBuilder':
        """
        Add code to create a new notebook.

        Returns:
            self for method chaining
        """
        self.comment("Create a new notebook")
        self.line("Notebook.New()")
        return self

    def create_worksheet(self, rows: int, cols: int) -> 'MacroBuilder':
        """
        Add code to create a worksheet.

        Args:
            rows: Number of rows in the worksheet
            cols: Number of columns in the worksheet

        Returns:
            self for method chaining
        """
        self.comment(f"Create a worksheet with {rows} rows and {cols} columns")
        self.line(f"Worksheet.Create({rows}, {cols})")
        return self

    def set_column_name(self, col: int, name: str) -> 'MacroBuilder':
        """
        Set a column name.

        Args:
            col: Column index (1-based)
            name: Column name

        Returns:
            self for method chaining
        """
        self.line(f"Worksheet.SetColName({col}, \"{name}\")")
        return self

    def set_cell_value(self, row: int, col: int, value: Any) -> 'MacroBuilder':
        """
        Set a cell value.

        Args:
            row: Row index (1-based)
            col: Column index (1-based)
            value: Cell value

        Returns:
            self for method chaining
        """
        if isinstance(value, str):
            self.line(f"Worksheet.SetCellValue({row}, {col}, \"{value}\")")
        else:
            self.line(f"Worksheet.SetCellValue({row}, {col}, {value})")
        return self

    def import_data(self, file_path: str, delimiter: str = "Tab") -> 'MacroBuilder':
        """
        Import data from a file.

        Args:
            file_path: Path to the data file
            delimiter: Delimiter type ("Tab", "Comma", "Space", etc.)

        Returns:
            self for method chaining
        """
        self.comment(f"Import data from {file_path}")

        # Handle delimiter type
        delimiter_code = 9
        if delimiter.lower() == "comma":
            delimiter_code = ","
        elif delimiter.lower() == "space":
            delimiter_code = " "

        # Set import options
        self.line("Notebook.Import.TextOptions.Delimiter = 9")
        self.line("Notebook.Import.TextOptions.HeaderRow = 1")

        # Escape path for SigmaPlot
        escaped_path = file_path.replace("\\", "\\\\")

        # Import the data
        self.line(f"Notebook.Import.Text(\"{escaped_path}\")")
        return self

    def create_graph(self, x_col: int, y_col: int, graph_type: str = "Line") -> 'MacroBuilder':
        """
        Create a graph.

        Args:
            x_col: X data column index (1-based)
            y_col: Y data column index (1-based)
            graph_type: Type of graph to create

        Returns:
            self for method chaining
        """
        self.comment(f"Create a {graph_type} graph")

        # Different ways to create graphs in SigmaPlot
        if graph_type.lower() == "wizard":
            self.line(f"Notebook.Graph.CreateWizardGraph({x_col}, {y_col}, 0)")
        else:
            self.line(f"Graph.Create({x_col}, {y_col})")

        return self

    def set_graph_properties(self, title: str = None, x_label: str = None,
                           y_label: str = None) -> 'MacroBuilder':
        """
        Set basic graph properties.

        Args:
            title: Graph title
            x_label: X-axis label
            y_label: Y-axis label

        Returns:
            self for method chaining
        """
        self.comment("Set graph properties")

        if title:
            self.line(f"Graph.Axis.Title.Text = \"{title}\"")

        if x_label:
            self.line(f"Graph.Axis.Label(1).Text = \"{x_label}\"")

        if y_label:
            self.line(f"Graph.Axis.Label(2).Text = \"{y_label}\"")

        return self

    def export_graph(self, file_path: str, format_type: str = "PNG",
                    resolution: int = 300) -> 'MacroBuilder':
        """
        Export the graph to a file.

        Args:
            file_path: Output file path
            format_type: File format type (PNG, TIFF, JPG, etc.)
            resolution: DPI resolution for raster formats

        Returns:
            self for method chaining
        """
        self.comment(f"Export graph to {file_path}")

        # Escape path for SigmaPlot
        escaped_path = file_path.replace("\\", "\\\\")

        # For better reliability in SPW12, use the GraphPage.Export approach
        self.line(f"Page.Object.Export(\"{escaped_path}\", \"{format_type}\", {resolution})")

        # Alternative direct export method if the above fails
        # self.line(f"Graph.Export(\"{escaped_path}\", \"{format_type}\", {resolution})")

        return self

    def quit_sigmaplot(self, save_changes: bool = False) -> 'MacroBuilder':
        """
        Add code to quit SigmaPlot.

        Args:
            save_changes: Whether to save changes before quitting

        Returns:
            self for method chaining
        """
        self.comment("Exit SigmaPlot")
        if not save_changes:
            self.line("Application.Quit()")
        else:
            self.line("Application.Save()")
            self.line("Application.Quit()")
        return self

    def add_debug_log(self, log_file: str) -> 'MacroBuilder':
        """
        Add debug logging to the macro.

        Args:
            log_file: Path to write debug log

        Returns:
            self for method chaining
        """
        escaped_path = log_file.replace("\\", "\\\\")
        self.comment("Create debug log file")
        self.line(f"LogFile = System.OpenFile(\"{escaped_path}\", 1)")
        self.line("System.WriteFile(LogFile, \"Macro started: \" & Now)")
        return self

    def log_message(self, message: str) -> 'MacroBuilder':
        """
        Add a log message to the debug log.

        Args:
            message: Message to log

        Returns:
            self for method chaining
        """
        self.line(f"System.WriteFile(LogFile, \"{message}: \" & Now)")
        return self

    def close_debug_log(self) -> 'MacroBuilder':
        """
        Close the debug log file.

        Returns:
            self for method chaining
        """
        self.comment("Close debug log")
        self.line("System.WriteFile(LogFile, \"Macro completed: \" & Now)")
        self.line("System.CloseFile(LogFile)")
        return self

    def to_string(self) -> str:
        """
        Convert the macro to a string.

        Returns:
            The complete macro as a string
        """
        return "\n".join(self.lines)

    def save_to_file(self, file_path: Optional[str] = None) -> str:
        """
        Save the macro to a file.

        Args:
            file_path: Path to save the macro. If None, creates a temp file.

        Returns:
            The path to the saved macro file
        """
        if file_path is None:
            # Create a temp file
            handle, file_path = tempfile.mkstemp(suffix=".mac", prefix="sigma_")
            os.close(handle)

        # Write macro to file
        with open(file_path, 'w') as f:
            f.write(self.to_string())

        return file_path

    def run(self, sigmaplot_path: str, timeout: int = 60) -> Tuple[bool, str]:
        """
        Run the macro in SigmaPlot.

        Args:
            sigmaplot_path: Path to SigmaPlot executable
            timeout: Timeout in seconds

        Returns:
            Tuple of (success, macro_path)
        """
        # Save macro to temp file
        macro_path = self.save_to_file()

        # Run SigmaPlot with macro
        success = run_sigmaplot_with_timeout(sigmaplot_path, macro_path, timeout)

        return success, macro_path

def run_sigmaplot_with_timeout(exe_path: str, macro_path: str, timeout: int = 60) -> bool:
    """
    Run SigmaPlot with a macro and timeout.

    Args:
        exe_path: Path to SigmaPlot executable
        macro_path: Path to macro file
        timeout: Timeout in seconds

    Returns:
        True if successful, False otherwise
    """
    proc = None

    def target():
        nonlocal proc
        try:
            proc = subprocess.Popen([exe_path, "/M", macro_path])
            proc.communicate()
        except Exception as e:
            print(f"Error in subprocess: {e}")

    thread = threading.Thread(target=target)
    thread.start()

    # Wait for the thread to complete or timeout
    thread.join(timeout)

    if thread.is_alive():
        print(f"SigmaPlot process timed out after {timeout} seconds")
        if proc:
            try:
                proc.terminate()
                print("SigmaPlot process terminated")
            except Exception as e:
                print(f"Error terminating process: {e}")
        return False

    return True

# EOF