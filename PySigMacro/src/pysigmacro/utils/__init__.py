#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-15 02:00:29 (ywatanabe)"
# File: /home/ywatanabe/proj/SigMacro/PySigMacro/src/pysigmacro/utils/__init__.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/proj/SigMacro/PySigMacro/src/pysigmacro/utils/__init__.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

"""
Utility functions for SigmaPlot automation
"""

from ._com_wrap import com_wrap

from ._sigmaplot_objects import SIGMAPLOT_OBJECTS
from ._sigmaplot_properties import SIGMAPLOT_PROPERTIES
from ._sigmaplot_methods import SIGMAPLOT_METHODS
from ._sigmaplot_inspect import inspect

from ._open import open
from ._close_all import close_all

# EOF