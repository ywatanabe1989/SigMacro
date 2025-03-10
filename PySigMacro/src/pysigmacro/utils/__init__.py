#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 01:49:42 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/utils/__init__.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/utils/__init__.py"
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Time-stamp: "2024-10-22 19:51:47 (ywatanabe)"
# File: __init__.py

import os as __os
import importlib as __importlib
import inspect as __inspect

# Get the current directory
current_dir = __os.path.dirname(__file__)

# Iterate through all Python files in the current directory
for filename in __os.listdir(current_dir):
    if filename.endswith(".py") and not filename.startswith("__"):
        module_name = filename[:-3]
        module = __importlib.import_module(f".{module_name}", package=__name__)

        # Import only functions and classes from the module
        for name, obj in __inspect.getmembers(module):
            if __inspect.isfunction(obj) or __inspect.isclass(obj):
                if not name.startswith("_"):
                    globals()[name] = obj

# Clean up temporary variables
del __os, __importlib, __inspect, current_dir, filename, module_name, module, name, obj

# EOF