#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-08 23:07:14 (ywatanabe)"
# File: /home/ywatanabe/proj/SigmaPlot-v12.0-Pysigmacro/scripts/sigmaplot-py/setup.py

THIS_FILE = "/home/ywatanabe/proj/SigmaPlot-v12.0-Pysigmacro/scripts/sigmaplot-py/setup.py"

from setuptools import setup, find_packages

setup(
    name="pysigmacro",
    version="0.1.0",
    package_dir={"": "src"},
    packages=find_packages(where="src"),
    install_requires=[
        "pywin32>=228",
        "numpy>=1.19.0",
    ],
    python_requires=">=3.7",
)

# EOF