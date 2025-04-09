#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-01 17:24:03 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/con/_open.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/con/_open.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import win32com.client
import subprocess
from ..path import to_win, to_wsl
from ._close_all import close_all
from ..com._wrap import wrap
from ..utils._wait import wait

def open(lpath=None, close_others=False, visible=True):
    """
    Open SigmaPlot application and optionally load a specific notebook file.
    This function launches SigmaPlot and can optionally open a specific notebook
    file. It handles path conversions between WSL and Windows formats and attempts
    to use environment variables to locate the SigmaPlot executable.
    Args:
    lpath (str, optional): Path to the SigmaPlot notebook file (.JNB) to open.
    Defaults to None, which opens SigmaPlot without loading a file.
    close_others (bool, optional): Whether to close all other SigmaPlot instances
    before opening a new one. Defaults to False.
    visible (bool, optional): Whether to make the SigmaPlot application visible.
    Defaults to True.
    Returns:
    object: A wrapped SigmaPlot COM object if successful, None otherwise.
    Raises:
    Prints an error message if an Exception occurs, but returns None instead of raising.
    """
    try:
        if close_others:
            close_all()

        # SigmaPlot bin path
        sp_bin_wsl = os.getenv(
            "SIGMAPLOT_BIN_PATH_WSL",
            "/mnt/c/Program Files (x86)/SigmaPlot/SPW16/Spw.exe",
        )
        sp_bin_win = os.getenv(
            "SIGMAPLOT_BIN_PATH_WIN",
            r"C:\Program Files (x86)\SigmaPlot\SPW16\Spw.exe",
        )

        if lpath:
            # Convert path once at the beginning
            nonlocal_lpath = os.path.abspath(lpath)
            lpath_win = to_win(nonlocal_lpath)
            lpath_wsl = to_wsl(nonlocal_lpath)

            def try_open():
                if close_others:
                    close_all(verbose=False)

                # Try to open the file with SigmaPlot
                for sp_bin in [sp_bin_win, sp_bin_wsl]:
                    if os.path.exists(sp_bin):
                        for path in [lpath_win, lpath_wsl]:
                            try:
                                if os.path.exists(path):
                                    subprocess.Popen([sp_bin, path])
                                    break
                            except Exception as e:
                                pass

            # Wait for notebook to be loaded
            def check_notebook_loaded():
                try:
                    try_open()
                    sp = win32com.client.Dispatch("SigmaPlot.Application")
                    spw = wrap(sp, "SigmaPlot")
                    # Get current notebook name
                    notebooks = spw.Notebooks_obj
                    # Compare with expected path
                    lpath_base = os.path.basename(nonlocal_lpath)
                    return lpath_base in notebooks.list
                except:
                    return False

            wait(
                wait_condition_func=check_notebook_loaded,
                success_msg=f"Notebook successfully loaded: {nonlocal_lpath}",
                failure_msg=f"Failed to load notebook: {nonlocal_lpath}",
            )

        sp = win32com.client.Dispatch("SigmaPlot.Application")
        spw = wrap(sp, "SigmaPlot")

        if not visible:
            spw.Visible = visible

        # Set the path attribute on the wrapper object
        if lpath:
            spw._path = nonlocal_lpath

        return spw
    except Exception as e:
        print(f"Error opening SigmaPlot: {str(e)}")
        return None

# def open(lpath=None, close_others=False, visible=True): # , print_env_vars=True
#     """
#     Open SigmaPlot application and optionally load a specific notebook file.

#     This function launches SigmaPlot and can optionally open a specific notebook
#     file. It handles path conversions between WSL and Windows formats and attempts
#     to use environment variables to locate the SigmaPlot executable.

#     Args:
#         lpath (str, optional): Path to the SigmaPlot notebook file (.JNB) to open.
#             Defaults to None, which opens SigmaPlot without loading a file.
#         close_others (bool, optional): Whether to close all other SigmaPlot instances
#             before opening a new one. Defaults to False.
#         visible (bool, optional): Whether to make the SigmaPlot application visible.
#             Defaults to True.

#     Returns:
#         object: A wrapped SigmaPlot COM object if successful, None otherwise.

#     Raises:
#         Prints an error message if an Exception occurs, but returns None instead of raising.
#     """
#     try:
#         if close_others:
#             close_all()

#         # SigmaPlot bin path
#         sp_bin_wsl = os.getenv(
#             "SIGMAPLOT_BIN_PATH_WSL",
#             "/mnt/c/Program Files (x86)/SigmaPlot/SPW16/Spw.exe",
#         )
#         sp_bin_win = os.getenv(
#             "SIGMAPLOT_BIN_PATH_WIN",
#             r"C:\Program Files (x86)\SigmaPlot\SPW16\Spw.exe",
#         )

#         if lpath:

#             def try_open():
#                 # JNB file path
#                 lpath = os.path.abspath(lpath)
#                 lpath_win = to_win(lpath)
#                 lpath_wsl = to_wsl(lpath)

#                 # Call SigmaPlot with the file as argument
#                 for sp_bin in [sp_bin_win, sp_bin_wsl]:
#                     if os.path.exists(sp_bin):
#                         for path in [lpath_win, lpath_wsl]:
#                             try:
#                                 if os.path.exists(path):
#                                     subprocess.Popen([sp_bin, path])
#                                     break
#                             except Exception as e:
#                                 pass

#             # Wait for notebook to be loaded
#             def check_notebook_loaded():

#                 try:
#                     if close_others:
#                         close_all()

#                     try_open()

#                     sp = win32com.client.Dispatch("SigmaPlot.Application")
#                     spw = wrap(sp, "SigmaPlot")

#                     # Get current notebook name
#                     notebooks = spw.Notebooks_obj

#                     # Compare with expected path
#                     lpath_base = os.path.basename(lpath)

#                     return lpath_base in notebooks.list

#                 except:
#                     return False

#             wait(
#                 wait_condition_func=check_notebook_loaded,
#                 success_msg=f"Notebook successfully loaded: {lpath}",
#                 failure_msg=f"Failed to load notebook: {lpath}",
#             )

#         sp = win32com.client.Dispatch("SigmaPlot.Application")
#         spw = wrap(sp, "SigmaPlot")

#         if not visible:
#             spw.Visible = visible

#         # Set the path attribute on the wrapper object
#         if lpath:
#             spw._path = lpath

#         return spw
#     except Exception as e:
#         print(f"Error opening SigmaPlot: {str(e)}")
#         return None

# EOF