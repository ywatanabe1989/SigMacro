#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-12 07:52:21 (ywatanabe)"
# File: /home/ywatanabe/proj/SigMacro/PySigMacro/examples/00_run_macro.py

__THIS_FILE__ = "/home/ywatanabe/proj/SigMacro/PySigMacro/examples/00_run_macro.py"

import sys
import os
import win32com.client
import subprocess
import time
from pysigmacro.core.connection import connect
from pysigmacro.utils.paths import to_win
import argparse

def main(path_macro_win: str, macro_name: str, *args,) -> None:
    """
    Run a SigmaPlot macro from a specified notebook file.

    Args:
        path_macro_win (str): Path to the SigmaPlot notebook (.JNB) file containing the macro
        macro_name (str, optional): Name of the macro to run. Defaults to "hello_world"
        *args: Additional arguments to pass to the macro via a text file

    Returns:
        None
    """
    # Convert path to Windows format if needed
    try:
        # Connect to SigmaPlot using the connect function
        sigmaplot = connect(
            file_path=path_macro_win, visible=True, launch_if_not_found=True, close_others=False
        )

        # Open the JNB file that implements the macro
        command = f'"{path_macro_win}" /runmacro'
        subprocess.run(command, shell=True)
        print(command)

        # Wait for SigmaPlot to load the file
        time.sleep(2)

        # Get the notebook object by name
        nbVBLib = None
        sFileName = os.path.basename(path_macro_win)

        # Try multiple times to find the notebook
        max_attempts = 5
        for attempt in range(max_attempts):
            try:
                for i in range(0, sigmaplot.Notebooks.Count):
                    nb = sigmaplot.Notebooks(i)
                    if nb.Name == sFileName:
                        nbVBLib = nb
                        break

                if nbVBLib is not None:
                    break

                print(
                    f"Attempt {attempt+1}: Notebook not found yet, waiting..."
                )
                time.sleep(1)
            except Exception as e:
                print(f"Error on attempt {attempt+1}: {e}")
                time.sleep(1)

        if nbVBLib is None:
            print(f"Could not find notebook: {sFileName}")
            return

        # Get the macro object by name
        try:
            nbiMacro = nbVBLib.NotebookItems(macro_name)
            if nbiMacro is not None:
                # Generate text file for arguments
                gen_args_text_file(path_macro_win, *args)
                # Run the macro
                nbiMacro.Run()
                print("Macro execution completed")
            else:
                print(f'Macro "{macro_name}" not found in {path_macro_win}')
        except Exception as e:
            print(f"Error accessing or running macro: {e}")

    except Exception as e:
        print(f"Error in main function: {e}")


def gen_args_text_file(path_macro: str, *args):
    """
    Generate a text file containing comma-separated arguments for the macro.

    The file will be created in the same directory as the notebook file with
    the name "arguments.txt". The macro can read this file to access the arguments.

    Args:
        path_macro_winDir (str): Directory where the arguments file will be saved
        *args: Arguments to be written to the file

    Returns:
        None
    """
    path_macro_dir = os.path.dirname(path_macro)
    # If one or more arguments are provided
    if args:
        # Combine arguments with comma separator
        arguments_text = ",".join(map(str, args))

        # Show error if directory doesn't exist
        if not os.path.exists(path_macro_winDir):
            print(f"Specified directory does not exist: {path_macro_dir}")
            return

        # Create file path in the folder
        file_path = os.path.join(path_macro_dir, "arguments.txt")

        # Write to file
        with open(file_path, "w") as file:
            file.write(arguments_text)

        print(f"Comma-separated text saved to {file_path}.")
    else:
        print("No arguments provided.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Run a SigmaPlot macro with arguments"
    )

    parser.add_argument(
        "args", nargs="*", help="Arguments to pass to the macro"
    )

    parser.add_argument(
        "--path",
        "-p",
        dest="path_macro_win",
        default="C:/Users/wyusu/Documents/SigmaPlot/SPW12/SigMacro.JNB",
        help="Path to the SigmaPlot notebook (.JNB) file",
    )

    parser.add_argument(
        "--macro",
        "-m",
        dest="macro_name",
        default="hello_world",
        help="Name of the macro to run",
    )

    args_parsed = parser.parse_args()
    print(args_parsed)

    main(path_macro_win=args_parsed.path_macro_win, macro_name=args_parsed.macro_name, *args_parsed.args)

# EOF