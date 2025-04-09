#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-01 17:13:18 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/image/_convert_bmp_to_tiff.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/image/_convert_bmp_to_tiff.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from PIL import Image
from ..utils._wait import wait
from ..utils._remove import remove

def convert_bmp_to_tiff(bmp_path, keep_orig=True):
    """
    Convert BMP to TIFF with lossless compression

    This function takes a BMP file path, opens the image using Pillow,
    and saves it as a TIFF file with no compression and maximum quality.

    Args:
        bmp_path (str): Path to the input BMP file

    Returns:
        str: Path to the converted TIFF file
    """
    tiff_path = bmp_path.replace(".bmp", ".tiff")

    # BMP to TIFF conversion
    with Image.open(bmp_path) as img:
        img.save(tiff_path, format="TIFF", compression=None, quality=100)

    # Wait for TIFF file to be created
    wait(
        wait_condition_func=lambda: os.path.exists(tiff_path),
        success_msg=f"TIFF file created: {tiff_path}",
        failure_msg=f"Failed to create TIFF file: {tiff_path}",
        sleep_sec=0.5,
        max_trials=10
    )

    # Remove original BMP
    if (os.path.exists(tiff_path) and (not keep_orig)):
        remove(bmp_path)

    return tiff_path

if __name__ == '__main__':
    # Convert the BMP we just exported
    bmp_path = PATH.replace(".JNB", ".bmp")
    tiff_path = convert_bmp_to_tiff(bmp_path)
    print(f"Converted to TIFF: {tiff_path}")

# EOF