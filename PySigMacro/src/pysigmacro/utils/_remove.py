#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-01 17:16:31 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_remove.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_remove.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from ._wait import wait

def remove(path):
    os.remove(path)
    wait(
        wait_condition_func=lambda: not os.path.exists(path),
        success_msg=f"Successfully removed: {path}",
        failure_msg=f"Failed to remove: {path}",
    )

# EOF