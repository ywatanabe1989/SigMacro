#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-15 00:13:17 (ywatanabe)"
# File: /home/ywatanabe/proj/SigMacro/PySigMacro/src/pysigmacro/utils/_close_all.py

import os

__THIS_FILE__ = (
    "/home/ywatanabe/proj/SigMacro/PySigMacro/src/pysigmacro/utils/_close_all.py"
)
__THIS_DIR__ = os.path.dirname(__THIS_FILE__)

import subprocess
import time

def close_all():
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

# EOF