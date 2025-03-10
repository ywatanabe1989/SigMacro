#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-10 08:23:18 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/utils/paths.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/utils/paths.py"

"""
Path utilities for SigmaPlot integration.
"""

import os
import tempfile
import subprocess
from typing import List, Optional
import os

def find_sigmaplot_executable():
    """
    Find the SigmaPlot executable path on the system.

    Returns:
        str: Path to SigmaPlot executable or None if not found
    """
    possible_paths = [
        r"C:\Program Files\SigmaPlot\SPW12\Spw.exe",
        r"C:\Program Files (x86)\SigmaPlot\SPW12\Spw.exe",
        r"C:\Program Files\SigmaPlot\SPW14\Spw.exe",
        r"C:\Program Files (x86)\SigmaPlot\SPW14\Spw.exe",
    ]

    # Check common installation paths
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

    # Not found
    return None

def create_temp_directory() -> str:
    """
    Create a temporary directory for SigmaPlot data files.

    Returns:
        str: Path to the created temporary directory
    """
    base_temp = tempfile.gettempdir()
    temp_dir = os.path.join(base_temp, "sigmaplot_py_temp")

    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)

    return temp_dir

def check_wsl_environment() -> bool:
    """
    Check if running in Windows Subsystem for Linux (WSL).

    Returns:
        bool: True if running in WSL, False otherwise
    """
    # Check for WSL-specific files
    wsl_indicators = ["/proc/sys/fs/binfmt_misc/WSLInterop", "/proc/version"]

    for path in wsl_indicators:
        if os.path.exists(path):
            try:
                with open(path, "r") as f:
                    content = f.read().lower()
                    if "microsoft" in content or "wsl" in content:
                        return True
            except:
                pass

    return False

def to_win(wsl_path):
    """
    Convert a WSL path to a Windows path.

    Args:
        wsl_path (str): Path in WSL format (e.g., /home/user/file.txt)

    Returns:
        str: Converted Windows path (e.g., C:\\Users\\user\\file.txt)
    """
    # Handle absolute paths
    if os.path.isabs(wsl_path):
        try:
            # Use wslpath command to convert the path
            result = subprocess.run(
                ['wslpath', '-w', wsl_path],
                capture_output=True,
                text=True,
                check=True
            )
            return result.stdout.strip()
        except (subprocess.SubprocessError, FileNotFoundError):
            # Fallback if wslpath doesn't work
            # Basic conversion for /mnt/c/ style paths
            if wsl_path.startswith('/mnt/'):
                drive = wsl_path[5:6].upper()
                path = wsl_path[7:].replace('/', '\\')
                return f"{drive}:{path}"
            return wsl_path
    # Handle relative paths by getting the absolute path first
    else:
        abs_path = os.path.abspath(wsl_path)
        return to_win(abs_path)

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
            ['wslpath', '-u', windows_path],
            capture_output=True,
            text=True,
            check=True
        )
        return result.stdout.strip()
    except (subprocess.SubprocessError, FileNotFoundError):
        # Fallback if wslpath doesn't work
        if re.match(r'^[A-Za-z]:', windows_path):
            drive = windows_path[0].lower()
            path = windows_path[2:].replace('\\', '/')
            return f"/mnt/{drive}{path}"
        return windows_path

# EOF