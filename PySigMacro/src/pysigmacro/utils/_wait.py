#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-01 17:22:27 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_wait.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_wait.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import time

# ANSI color codes
GREEN = '\033[92m'
RED = '\033[91m'
RESET = '\033[0m'

def wait(
    wait_condition_func,
    success_msg=None,
    failure_msg=None,
    sleep_sec=3,
    max_trials=5,
    verbose=True,
):
    """Wait for a wait_condition_func to become True, with a timeout."""
    if wait_condition_func():
        if success_msg and verbose:
            print(f"{GREEN}{success_msg}{RESET}")
        return

    n_trials = 0
    while not wait_condition_func():
        time.sleep(sleep_sec)
        n_trials += 1
        if max_trials < n_trials:
            time.sleep(sleep_sec)
            break

    if wait_condition_func():
        if success_msg and verbose:
            print(f"{GREEN}{success_msg}{RESET}")
    else:
        if failure_msg and verbose:
            print(f"{RED}{failure_msg}{RESET}")

# EOF