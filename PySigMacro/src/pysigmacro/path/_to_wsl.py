#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-29 17:53:15 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/path/_to_wsl.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/path/_to_wsl.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import re
import subprocess

def to_wsl(windows_path):
    """
    Convert a Windows path to a WSL path.

    Args:
        windows_path (str): Path in Windows format (e.g., C:\\Users\\user\\file.txt)

    Returns:
        str: Converted WSL path (e.g., /mnt/c/Users/user/file.txt)
    """
    try:
        # Use wslpath command to convert the path
        result = subprocess.run(
            ["wslpath", "-u", windows_path],
            capture_output=True,
            text=True,
            check=True,
        )
        return result.stdout.strip()
    except (subprocess.SubprocessError, FileNotFoundError):
        # Fallback if wslpath doesn't work
        if re.match(r"^[A-Za-z]:", windows_path):
            drive = windows_path[0].lower()
            path = windows_path[2:].replace("\\", "/")
            return f"/mnt/{drive}{path}"
        return windows_path

# EOF