#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-26 19:21:46 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_print_pysigmacro_env_vars.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_print_pysigmacro_env_vars.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

def print_envs():
    print("Environmental variables for pysigmacro are set as follows:")
    for env_var_candidate in [
        "SIGMACRO_JNB_PATH",
        "SIGMACRO_TEMPLATES_DIR",
        "SIGMAPLOT_BIN_PATH_WIN",
        "SIGMAPLOT_BIN_PATH_WSL",
    ]:
        env_var = os.getenv(env_var_candidate, "Not Set")
        print(f"{env_var_candidate}: {env_var}")

# EOF