#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-29 17:24:46 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_args.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_args.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

DEFAULT_TEMPLATES_DIR = os.getenv("SIGMACRO_TEMPLATES_DIR")
DEFAULT_ARGS_PATH = os.path.join(DEFAULT_TEMPLATES_DIR, ".args.txt") if DEFAULT_TEMPLATES_DIR else None
SIGMACRO_ARGS_PATH = os.getenv("SIGMACRO_ARGS_PATH", DEFAULT_ARGS_PATH)

def to_args(**kwargs):
    with open(SIGMACRO_ARGS_PATH, 'w') as f:
        for k, v in kwargs.items():
            f.write(f"{k}={v}\n")

def list_args():
    args = {}
    try:
        with open(SIGMACRO_ARGS_PATH, 'r') as f:
            for line in f:
                line = line.strip()
                if '=' in line:
                    k, v = line.split('=', 1)
                    args[k] = v
    except FileNotFoundError:
        pass
    return args

# EOF