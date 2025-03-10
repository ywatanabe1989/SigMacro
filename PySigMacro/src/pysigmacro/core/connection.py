#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-10 04:21:31 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/core/connection.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/core/connection.py"

"""
Connection utilities for SigmaPlot.
"""
import os
import subprocess
import time
import sys
import win32com.client
from pysigmacro.utils.paths import to_win


def connect(
        file_path=None, visible=True, launch_if_not_found=True, close_others=False
):
    """
    Improved function to connect to SigmaPlot

    Args:
        visible (bool): Whether to make SigmaPlot visible
        launch_if_not_found (bool): Whether to launch SigmaPlot if not found
        close_others (bool): Whether to close other SigmaPlot instances
        file_path (str): Optional path to a SigmaPlot file to open upon connection

    Returns:
        The SigmaPlot application object
    """
    # 1. Force close any existing SigmaPlot instances if requested
    if close_others:
        try:
            subprocess.run(
                ["taskkill", "/f", "/im", "spw.exe"],
                shell=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
            time.sleep(2)
        except Exception as e:
            print(f"Warning when closing SigmaPlot: {e}")

    # 2. Launch SigmaPlot directly if requested
    if launch_if_not_found:
        possible_paths = [
            r"C:\Program Files\SigmaPlot\SPW12\Spw.exe",
            r"C:\Program Files (x86)\SigmaPlot\SPW12\Spw.exe",
            r"C:\Program Files\SigmaPlot\SPW14\Spw.exe",
            r"C:\Program Files (x86)\SigmaPlot\SPW14\Spw.exe",
        ]

        # Try to launch SigmaPlot with or without a file
        launched = False
        for path in possible_paths:
            if os.path.exists(path):
                try:
                    if file_path:
                        # Convert path if needed
                        win_path = to_win(file_path)

                        if "/runmacro" in win_path:
                            # If path already contains runmacro, use as is
                            cmd = win_path
                        else:
                            # Otherwise, launch with the file and optional runmacro flag
                            cmd = f'"{path}" "{win_path}"'

                        print(f"Launching SigmaPlot with file: {cmd}")
                        subprocess.Popen(cmd, shell=True)
                    else:
                        print(f"Launching SigmaPlot from {path}")
                        subprocess.Popen([path])

                    time.sleep(5)
                    launched = True
                    break
                except Exception as e:
                    print(f"Failed to launch SigmaPlot: {e}")

        if not launched and file_path:
            # If we couldn't launch SigmaPlot with our executables,
            # try to launch just with the file path
            try:
                win_path = to_win(file_path)
                print(f"Attempting to launch via file association: {win_path}")
                subprocess.Popen(f'"{win_path}"', shell=True)
                time.sleep(5)
            except Exception as e:
                print(f"Failed to launch SigmaPlot via file: {e}")

    # 3. Connect to SigmaPlot
    try:
        # First try to connect to running instance
        initial_app = win32com.client.Dispatch("SigmaPlot.Application.1")
        print("Created SigmaPlot instance using SigmaPlot.Application.1")

        # Make it visible if requested
        if visible:
            initial_app.Visible = True

        # Get the actual application object
        try:
            actual_app = initial_app.Application
            print(
                "Successfully accessed the actual SigmaPlot application object"
            )

            # If we have a file path but didn't launch SigmaPlot with it above,
            # open the file now
            if file_path and not launch_if_not_found:
                win_path = to_win(file_path)
                try:
                    # Try to open via Notebooks.Open
                    actual_app.Notebooks.Open(win_path)
                    print(f"Opened file using Notebooks.Open: {win_path}")
                except:
                    # Fall back to shell command
                    cmd = f'"{win_path}" /runmacro'
                    subprocess.run(cmd, shell=True)
                    print(f"Opened file using command: {cmd}")

            return actual_app
        except:
            # If we can't get the Application property, return the initial app
            print("Using initial application connection")

            # Try to open file with this connection too
            if file_path and not launch_if_not_found:
                win_path = to_win(file_path)
                try:
                    initial_app.Notebooks.Open(win_path)
                except:
                    cmd = f'"{win_path}" /runmacro'
                    subprocess.run(cmd, shell=True)

            return initial_app
    except Exception as e:
        print(f"Failed to connect to SigmaPlot: {e}")
        return None


def _launch():
    """
    Launch SigmaPlot executable.

    Returns:
        bool: True if SigmaPlot was launched successfully, False otherwise
    """
    possible_paths = [
        r"C:\Program Files\SigmaPlot\SPW12\Spw.exe",
        r"C:\Program Files (x86)\SigmaPlot\SPW12\Spw.exe",
        r"C:\Program Files\SigmaPlot\SPW14\Spw.exe",
        r"C:\Program Files (x86)\SigmaPlot\SPW14\Spw.exe",
    ]

    for path in possible_paths:
        if os.path.exists(path):
            try:
                print(f"Launching SigmaPlot from {path}")
                subprocess.Popen([path])
                time.sleep(5)
                return True
            except Exception as e:
                print(f"Failed to launch SigmaPlot: {e}")

    print("SigmaPlot executable not found in expected locations")
    return False


def close(app):
    """
    Close SigmaPlot application.

    Args:
        app: SigmaPlot application COM object

    Returns:
        bool: True if closed successfully, False otherwise
    """
    try:
        app.Quit()
        return True
    except Exception as e:
        print(f"Error closing SigmaPlot: {e}")
        return False

# EOF