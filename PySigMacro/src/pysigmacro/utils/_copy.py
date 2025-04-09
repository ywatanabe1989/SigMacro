#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-01 17:16:56 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_copy.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_copy.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import shutil
from ..path._to_win import to_win
from ..path._to_wsl import to_wsl
from ..utils._add_timestamp import add_timestamp
from ._wait import wait


def copy(src_path, dst_path, convert_paths=True):
    """
    Copy a file from source to destination, with optional path conversion.

    Args:
        src_path (str): Source file path
        dst_path (str): Destination file path
        convert_paths (bool): Whether to convert between WSL and Windows paths
                             based on environment detection

    Returns:
        str: Path to the destination file
    """
    # Check if we're in WSL
    in_wsl = os.path.exists("/proc/sys/fs/binfmt_misc/WSLInterop")

    if convert_paths and in_wsl:
        # If in WSL, ensure paths are in WSL format for shutil operations
        src_path_use = to_wsl(src_path)
        dst_path_use = to_wsl(dst_path)
    else:
        src_path_use = src_path
        dst_path_use = dst_path

    # Create destination directory if it doesn't exist
    dst_dir = os.path.dirname(dst_path_use)
    if dst_dir and not os.path.exists(dst_dir):
        os.makedirs(dst_dir, exist_ok=True)

    # Copy the file
    shutil.copy2(src_path_use, dst_path_use)

    wait(
        wait_condition_func=lambda: os.path.exists(dst_path_use),
        success_msg=f"Successfully copied: {src_path_use} to {dst_path_use}",
        failure_msg=f"Failed to copy: {src_path_use} to {dst_path_use}",
    )

    return dst_path


def copy_template(plot_type, tgt_dir, src_dir=None):
    if not src_dir:
        src_dir = os.getenv(
            "SIGMACRO_TEMPLATES_DIR",
            os.path.join("../data/templates")
        )
        # r"C:\Users\wyusu\Documents\SigMacro\SigMacro\Templates"
    src_path = os.path.join(src_dir, f"{plot_type}.JNB")
    tgt_path = add_timestamp(os.path.join(tgt_dir, f"{plot_type}.JNB"))
    copy(src_path, tgt_path)
    assert os.path.exists(tgt_path)
    return tgt_path

# EOF