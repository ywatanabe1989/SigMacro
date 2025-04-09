#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-01 13:46:36 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/con/_close_all.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/con/_close_all.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import subprocess
import time

def close_all(verbose=True):
    """
    Force close all instances of SigmaPlot application.

    This function uses Windows' taskkill command to forcefully terminate
    all running SigmaPlot processes. It waits for 2 seconds after sending
    the kill command to ensure processes have time to terminate.

    Returns:
        None

    Raises:
        Prints a warning message if an Exception occurs, but does not raise it.
    """
    from ..utils._wait import wait
    try:
        subprocess.run(
            ["taskkill", "/f", "/im", "spw.exe"],
            shell=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        # Check if all SigmaPlot processes have been terminated
        def check_sigmaplot_closed():
            result = subprocess.run(
                ["tasklist", "/fi", "imagename eq spw.exe"],
                shell=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )
            return "spw.exe" not in result.stdout

        wait(
            wait_condition_func=check_sigmaplot_closed,
            success_msg="All SigmaPlot instances successfully closed",
            failure_msg="Failed to close all SigmaPlot instances",
            verbose=verbose,
        )
    except Exception as e:
        print(f"Warning when closing SigmaPlot: {e}")

# EOF