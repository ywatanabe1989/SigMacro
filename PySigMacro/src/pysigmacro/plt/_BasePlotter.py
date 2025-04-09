#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-26 12:36:56 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/plt/_BasePlotter.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/plt/_BasePlotter.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

# Usage:
# pysigmacro.plt.area()
# pysigmacro.plt.line()
# pysigmacro.plt.area
# pysigmacro.plt.bar
# pysigmacro.plt.barh
# pysigmacro.plt.box
# pysigmacro.plt.box_h
# pysigmacro.plt.line
# pysigmacro.plt.filled_line
# pysigmacro.plt.polar
# pysigmacro.plt.scatter
# pysigmacro.plt.violin

# pysigmacro.plt.contour
# pysigmacro.plt.confusion_matrix


class _BasePlotter(_MethodMixin, _MacroMixin):
    def __init__(
        self,
        plot_type,
        data_df,
        xlabel,
        ylabel,
        xscale,
        yscale,
        xmin,
        ymin,
        xmax,
        ymax,
        xticks,
        yticks,
    ):
        self.plot_type = plot_type
        self.data_df = data_df
        self.xlabel = xlabel
        self.ylabel = ylabel
        self.xscale = xscale
        self.yscale = yscale
        self.xmin = xmin
        self.ymin = ymin
        self.xmax = xmax
        self.ymax = ymax
        self.xticks = xticks
        self.yticks = yticks

    def validate_plot_type(self,):
    plot_types = [
        "area",
        "bar",
        "barh",
        "box",
        "box_h",
        "confusion_matrix",
        "contour",
        "line",
        "filled_line",
        "polar",
        "scatter",
        "violin"
    ]

    def validate_data_columns(self):
        valid_column_prefixes = ["x", "y", "xerr", "yerr", "rgb"]
        for col in self.data_df.columns:
            is_valid = False
            for prefix in valid_column_prefixes:
                if col.startswith(prefix) and col[len(prefix) :].isdigit():
                    is_valid = True
                    break
            if not is_valid:
                valid_examples = [
                    f"{prefix}N (e.g., {prefix}1, {prefix}2, ...)"
                    for prefix in valid_column_prefixes
                ]
                raise ValueError(
                    f"Invalid column name: {col}. Valid column formats are: {', '.join(valid_examples)}"
                )
        return True

    def _area(self):
        pass

    def _bar(self):
        pass

    def _barh(self):
        pass

    def _box(self):
        pass

    def _box_h(self):
        pass

    def _confusion_matrix(self):
        pass

    def _contour(self):
        pass

    def _line(self):
        pass

    def _filled_line(self):
        pass

    def _polar(self):
        pass

    def _scatter(self):
        pass

    def _violin(self):
        pass



class _MethodMixin:
    def export_as_jpg(self, path=None, crop=False):
        """Export graph item as JPG, with optional cropping"""
        # Use self.path as default if path is not provided
        if path is None:
            path = os.path.splitext(self.path)[0] + ".jpg"
        # Export to JPG
        self._com_object.Export(path, "JPG")
        # Crop if requested
        if crop:
            from ..image import crop_images

            crop_images([path])
        return path

    def export_as_bmp(self, path=None, crop=False):
        """Export graph item as BMP, with optional cropping"""
        # Use self.path as default if path is not provided
        if path is None:
            path = os.path.splitext(self.path)[0] + ".bmp"
        # Export to BMP
        self._com_object.Export(path, "BMP")
        # Crop if requested
        if crop:
            from ..image import crop_images

            crop_images([path])
        return path

    def export_as_tif(self, path=None, crop=True, convert_from_bmp=True):
        """
        Export graph item as TIF, with optional cropping
        If convert_from_bmp is True, will first export as BMP then convert to TIFF
        (useful when direct TIFF export doesn't work properly)
        """
        # Use self.path as default if path is not provided
        if path is None:
            try:
                path = os.path.splitext(self.path)[0] + ".tif"
            except AttributeError:
                raise ValueError("No path provided and self.path is not set")

        if convert_from_bmp:
            # Create temporary BMP path
            bmp_path = os.path.splitext(path)[0] + ".bmp"
            # Export to BMP first
            self.export_as_bmp(bmp_path, crop=crop)
            # Convert to TIFF
            from ..image import convert_bmp_to_tiff

            tiff_path = convert_bmp_to_tiff(
                bmp_path
                if not crop
                else bmp_path.replace(".bmp", "_cropped.bmp")
            )
            # Delete temporary BMP if it was not cropped
            if not crop:
                try:
                    os.remove(bmp_path)
                except:
                    pass
            return tiff_path
        else:
            # Direct TIFF export
            try:
                self._com_object.Export(path, "TIF")
                # Crop if requested
                if crop:
                    from ..image import crop_images

                    crop_images([path])
                return path
            except Exception as e:
                print(f"Direct TIFF export failed: {e}")
                print("Attempting BMP conversion method...")
                return self.export_as_tif(
                    path, crop=crop, convert_from_bmp=True
                )

class _MacroMixin:
    def _run_all_macros(self):
        pass

    def _run_macro(self):
        from ..utils._run_macro import run_macro
        run_macro(
            self, "RenameXYLabels_macro"
        )
        # Validation
        self.validate_property_columns()
        self.validate_data_columns()

    def validate_property_columns(self):
        valid_columns = [
            "xlabel",
            "ylabel",
            "xscale",
            "yscale",
            "xmin",
            "ymin",
            "xmax",
            "ymax",
            "xticks",
            "yticks",
        ]
        for col in self.data_df.columns:
            if col not in valid_columns:
                raise ValueError(
                    f"Invalid column name: {col}. Valid columns are: {valid_columns}"
                )
        return True

# EOF