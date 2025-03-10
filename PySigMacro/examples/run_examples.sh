#!/bin/bash
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 12:18:26 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/examples/run_examples.sh

THIS_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
LOG_PATH="$0.log"
touch "$LOG_PATH"

SCRIPTS=(
    # 01_create_notebook.py
    # 02_create_and_fill_worksheet.py
    03_embed_macro.py
    # 03_import_data.py
    # 04_simple_plot_claude.py
    # 03_simple_plot_o1.py
)
for file in "${SCRIPTS[@]}"; do
    cmd="powershell.exe python.exe ./examples/$file"
    echo
    echo $cmd 2>&1 | tee -a $LOG_PATH
    echo
    eval $cmd 2>&1 | tee -a $LOG_PATH
done

# EOF