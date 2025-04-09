#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-25 04:28:00 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_MacroItemWrapper.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_MacroItemWrapper.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from ..const import *
from ._BaseCOMWrapper import BaseCOMWrapper

class MacroItemWrapper(BaseCOMWrapper):
    """Specialized wrapper for MacroItem object"""

    __classname__ = "MacroItemWrapper"

    def run(self, tgt_obj=None, **kwargs):
        from ..utils import to_args
        if tgt_obj:
            tgt_obj.activate()
        to_args(**kwargs)
        self.RUN()

# EOF