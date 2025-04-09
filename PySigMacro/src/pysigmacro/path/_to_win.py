#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-30 10:28:05 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/path/_to_win.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/path/_to_win.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import subprocess

def to_win(wsl_path):
    """
    Convert a WSL path to a Windows path.
    """
    # Special handling for your custom path structure
    if wsl_path.startswith("/home/ywatanabe/win/"):
        # Map to your actual Windows user directory
        relative_path = wsl_path[len("/home/ywatanabe/win/") :].replace(
            "/", "\\"
        )
        windows_user = os.environ.get(
            "USERNAME", "wyusu"
        )  # Get Windows username or default
        win_path = f"C:\\Users\\{windows_user}\\{relative_path}"
        win_path = (
            win_path.replace("program_files", "Program Files")
            .replace("program_files_x86", "Program Files (x86)")
            .replace("C:\\Users\\{windows_user}\\template", "C:\\Users\\{windows_user}\\Template")
            .replace("C:\\Users\\{windows_user}\\documents", "C:\\Users\\{windows_user}\\Documents")
            .replace("C:\\Users\\{windows_user}\\desktop", "C:\\Users\\{windows_user}\\Desktop")
        )
        return win_path

    # Rest of your existing function...
    if os.path.isabs(wsl_path):
        try:
            result = subprocess.run(
                ["wslpath", "-w", wsl_path],
                capture_output=True,
                text=True,
                check=True,
            )
            return result.stdout.strip()
        except (subprocess.SubprocessError, FileNotFoundError):
            if wsl_path.startswith("/mnt/"):
                drive = wsl_path[5:6].upper()
                path = wsl_path[7:].replace("/", "\\")
                return f"{drive}:{path}"
            return wsl_path
    else:
        abs_path = os.path.abspath(wsl_path)
        return to_win(abs_path)

# EOF