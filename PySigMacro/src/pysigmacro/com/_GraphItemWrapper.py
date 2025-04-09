#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-05 06:01:06 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_GraphItemWrapper.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_GraphItemWrapper.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from ..const import *
from ._BaseCOMWrapper import BaseCOMWrapper
from ..utils._remove import remove


class GraphItemWrapper(BaseCOMWrapper):
    """Specialized wrapper for GraphItem object"""

    __classname__ = "GraphItemWrapper"

    def export_as_gif(self, path=None, crop=True, keep_orig=True):
        return self._export_as_base(
            "GIF", path=path, crop=crop, keep_orig=keep_orig
        )

    def export_as_jpg(self, path=None, crop=True, keep_orig=True):
        return self._export_as_base(
            "JPG", path=path, crop=crop, keep_orig=keep_orig
        )

    def export_as_bmp(self, path=None, crop=True, keep_orig=True):
        return self._export_as_base(
            "BMP", path=path, crop=crop, keep_orig=keep_orig
        )

    def export_as_tif(self, path=None, crop=True, keep_orig=True):
        """
        Export graph item as TIF, with optional cropping
        If convert_from_bmp is True, will first export as BMP then convert to TIFF
        (useful when direct TIFF export doesn't work properly)
        """
        from ..image import convert_bmp_to_tiff

        if path is None:
            path_tiff = os.path.splitext(self.path)[0] + ".tif"
        else:
            path_tiff = path

        # JNB to BMP
        path_bmp = os.path.splitext(path_tiff)[0] + ".bmp"
        self.export_as_bmp(path_bmp, crop=crop, keep_orig=keep_orig)

        # BMP to TIFF
        actual_path_bmp = (
            path_bmp if not crop else path_bmp.replace(".bmp", "_cropped.bmp")
        )
        tiff_path_out = convert_bmp_to_tiff(
            actual_path_bmp, keep_orig=keep_orig
        )

        return tiff_path_out

    def _export_as_base(
        self, image_format, path=None, crop=True, keep_orig=True
    ):
        image_format = image_format.upper()
        assert image_format in [
            "JPG",
            "BMP",
            "GIF",
        ], f"Unsupported image format: {image_format}. Supported formats are JPG, BMP, and GIF."
        ext = f".{image_format.lower()}"
        if path is None:
            path = os.path.splitext(self.path)[0] + ext
        # Export to BMP
        self._com_object.Export(path, image_format)
        # Crop if requested
        if crop:
            from ..image import crop_images

            crop_images([path], keep_orig=keep_orig)
        return path

    def rename_xy_labels(self, xlabel, ylabel):
        from ..utils._run_macro import run_macro

        run_macro(self, "RenameXYLabels_macro", xlabel=xlabel, ylabel=ylabel)

# EOF