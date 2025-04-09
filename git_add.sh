#!/bin/bash
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-09 14:10:38 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/git_add.sh

THIS_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
LOG_PATH="$0.log"
touch "$LOG_PATH"


git add vba/ALL-IN-ONE-MACRO.vba
git add templates/jnb/template.JNB
git add templates/gif/jitter*_cropped.gif -f
# git add templates/gif/line_yerr-*_cropped.gif -f
# git add templates/gif/lines_y_many_x_cropped.gif -f
git add README.md

# EOF