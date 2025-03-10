#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 04:07:35 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/utils/load_text.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/utils/load_text.py"

import os

def load_text(file_path: str) -> str:
    """
    Load text from a file.

    Args:
        file_path: Path to the file to load

    Returns:
        Text content of the file or empty string if file not found
    """
    try:
        # Get directory where this module is located
        module_dir = os.path.dirname(os.path.abspath(__file__))
        # Calculate absolute path to the VBA file
        full_path = os.path.join(module_dir, file_path)

        # Check if file exists
        if os.path.exists(full_path):
            with open(full_path, 'r') as f:
                return f.read()
        else:
            print(f"Warning: VBA file not found: {full_path}")
            return ""
    except Exception as e:
        print(f"Error loading VBA file: {e}")
        return ""

# EOF