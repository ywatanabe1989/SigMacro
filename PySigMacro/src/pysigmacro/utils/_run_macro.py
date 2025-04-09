#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-01 09:48:36 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_run_macro.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_run_macro.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from ._args import to_args

def run_macro(tgt_obj, macroname, **kwargs):
    from ..con._open import open as ps_con_open

    # Specify Macro
    to_args(**kwargs)
    macro_PATH = os.getenv("SIGMACRO_JNB_PATH", rf"C:\Users\{os.getlogin()}\Documents\SigMacro\SigMacro.JNB")
    # Fixme: Macro will be available in the same notebook (= *.jnb file)
    macro_spw = ps_con_open(macro_PATH)
    macro_notebooks = macro_spw.Notebooks_obj
    macro_notebook = macro_notebooks[macro_notebooks.find_indices(os.path.basename(macro_PATH))[0]]
    macro_notebookitems = macro_notebook.NotebookItems_obj
    macroitem = macro_notebookitems[macro_notebookitems.find_indices(macroname)[0]]

    # Activate the target object
    tgt_obj.activate()

    # Run the macro
    macroitem.run(tgt_obj=tgt_obj, **kwargs)

# EOF