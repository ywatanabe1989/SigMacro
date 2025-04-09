#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-23 10:20:41 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/path/_add_timestamp.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/path/_add_timestamp.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from datetime import datetime


def add_timestamp(path):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_path, ext = path.rsplit(".", 1)
    return f"{base_path}_{timestamp}.{ext}"

# EOF