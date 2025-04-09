#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-09 20:41:18 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/demo/_gen_jnb.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/demo/_gen_jnb.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

"""
Template Creation Class
"""

import time

from ..con._close_all import close_all as ps_con_close_all
from ..con._open import open as ps_con_open
from ..data._import_data import import_data as ps_data_import_data
from ..utils._remove import remove


class JNBGenerator:
    def __init__(
        self,
        plot_types,
        templates_dir=None,
        df=None,
        from_demo_csv=True,
    ):
        self.plot_types = plot_types
        self.df = df
        self.from_demo_csv = from_demo_csv

        # Names
        self.section_name = f"section"
        self.worksheet_name = f"worksheet"
        self.graph_name = f"graph"

        # Path
        if not templates_dir:
            templates_dir = os.getenv(
                "SIGMACRO_TEMPLATES_DIR", os.path.join(__DIR__, "templates")
            )
        self.path_jnb_template = os.path.join(
            templates_dir, "jnb", "template.JNB"
        )
        fname = "-".join(plot_types)

        self.path_jnb = os.path.join(templates_dir, "jnb", f"{fname}.jnb")
        self.path_csv = self.path_jnb.replace("jnb", "csv")
        self.path_tif = self.path_jnb.replace("jnb", "tif")
        self.path_png = self.path_jnb.replace("jnb", "png")
        self.path_jpg = self.path_jnb.replace("jnb", "jpg")
        self.path_gif = self.path_jnb.replace("jnb", "gif")
        self.path_bmp = self.path_jnb.replace("jnb", "bmp")

        # assert os.path.exists(
        #     self.path_csv
        # ), f"CSV path does not exist: {self.path_csv}"
        assert os.path.exists(
            self.path_jnb_template
        ), f"JNB template path does not exist: {self.path_jnb_template}"

    def remove_template(self):
        if os.path.exists(self.path_jnb):
            remove(self.path_jnb)

    def copy_template(self):
        from ..utils._copy import copy

        copy(self.path_jnb_template, self.path_jnb)

    def process_connection(self):
        self.spw = ps_con_open(lpath=self.path_jnb, close_others=True)

    def process_application(self):
        self.app = self.spw.Application_obj

    def process_notebooks(self):
        self.notebooks = self.app.Notebooks_obj
        self.notebooks.clean()

    def process_notebook(self):
        filename = os.path.basename(self.path_jnb)
        try:
            filename in self.notebooks.list
            self.notebook = self.notebooks[
                self.notebooks.find_indices(filename)[0]
            ]
        except Exception as e:
            print(e)
            exit

    def process_notebookitems(
        self,
    ):
        # NotebookItems
        self.notebookitems = self.notebook.NotebookItems_obj
        self.notebookitems.clean()

        try:
            self.sectionitem = self.notebookitems[
                self.notebookitems.find_indices("section")[0]
            ]
            self.worksheetitem = self.notebookitems[
                self.notebookitems.find_indices("worksheet")[0]
            ]
            self.graphitem = self.notebookitems[
                self.notebookitems.find_indices("graph")[0]
            ]
            self.all_in_one_macro = self.notebookitems[
                self.notebookitems.find_indices("all-in-one-macro")[0]
            ]
        except Exception as e:
            print(e)
            __import__("ipdb").set_trace()

    def import_data(self, df=None):
        if self.from_demo_csv:
            csv = self.path_csv
        else:
            csv = None
        self.datatable = ps_data_import_data(
            self.worksheetitem, df=self.df, csv=csv
        )

    def run_all_in_one_macro(self):
        self.worksheetitem.activate()
        self.graphitem.activate()
        self.all_in_one_macro.run()
        self.graphitem.activate()
        time.sleep(1)

    def save_notebook(self):
        self.notebook.Save()

    def _run(self, keep_orig=False):
        ps_con_close_all()
        self.remove_template()
        self.copy_template()
        self.process_connection()
        self.process_application()
        self.process_notebooks()
        self.process_notebook()
        self.process_notebookitems()
        self.import_data()
        self.save_notebook()
        self.run_all_in_one_macro()
        self.save_notebook()
        self.graphitem.export_as_tif(path=self.path_tif, keep_orig=keep_orig)
        self.graphitem.export_as_gif(path=self.path_gif, keep_orig=keep_orig)
        self.save_notebook()
        # ps_con_close_all()
        return self.path_jnb

    def run(self, keep_orig_figures=False):
        max_retry = 3
        last_error = None

        for retry_count in range(max_retry):
            try:
                return self._run(keep_orig=keep_orig_figures)
            except Exception as e:
                last_error = e
                print(
                    f"Attempt {retry_count + 1}/{max_retry} failed: {str(e)}"
                )
                if retry_count < max_retry - 1:
                    print(f"Retrying...")

        print(
            f"All {max_retry} attempts failed. Last error: {str(last_error)}"
        )
        raise last_error


def gen_jnb(plot_types):
    generator = JNBGenerator(plot_types, from_demo_csv=True)
    generator.run()


if __name__ == "__main__":
    from ..const._PLOT_TYPES import PLOT_TYPES

    for plot_type in PLOT_TYPES:
        gen_jnb(plot_type, df=None)

# EOF