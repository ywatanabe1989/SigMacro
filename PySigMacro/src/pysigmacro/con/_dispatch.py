#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-26 18:46:25 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/con/_dispatch.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/con/_dispatch.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import win32com.client

def dispatch():
    """
    Create and return a COM object for the SigmaPlot application.

    This function uses the win32com.client.Dispatch method to create
    and return a COM object that represents the SigmaPlot application.

    Returns:
        COM object: A COM object representing the SigmaPlot application.
    """
    return win32com.client.Dispatch("SigmaPlot.Application")

def get_app():
    """
    Get the Application property of the SigmaPlot COM object.

    This function creates a COM object for SigmaPlot and returns its
    Application property, which provides access to the application's
    functionality.

    Returns:
        COM object: The Application property of the SigmaPlot COM object.
    """
    return win32com.client.Dispatch("SigmaPlot.Application").Application

# EOF