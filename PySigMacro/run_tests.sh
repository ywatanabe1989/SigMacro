#!/bin/bash
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-08 23:38:38 (ywatanabe)"
# File: /home/ywatanabe/proj/SigmaPlot-v12.0-Pysigmacro/pysigmacro/run_tests.sh

THIS_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
LOG_PATH="$0.log"
touch "$LOG_PATH"

usage() {
    echo "Usage: $0 [-m|--mark MARK] [-v|--verbose] [-h|--help]"
    echo ""
    echo "Options:"
    echo "  -m, --mark MARK     Run only tests with the specified marker"
    echo "  -v, --verbose       Run tests in verbose mode"
    echo "  -h, --help          Display this help message"
    echo ""
    echo "Example:"
    echo "  $0 -m 'not windows'"
    echo "  $0 -v"
    exit 1
}

MARK=""
VERBOSE=""

while [[ $# -gt 0 ]]; do
    case $1 in
        -m|--mark)
            MARK="$2"
            shift 2
            ;;
        -v|--verbose)
            VERBOSE="-v"
            shift
            ;;
        -h|--help)
            usage
            ;;
        *)
            echo "Unknown option: $1"
            usage
            ;;
    esac
done

# Create an empty log file
> "$LOG_PATH"

# Determine if we're running in WSL
if grep -q Microsoft /proc/version || grep -q WSL /proc/version; then
    echo "Running in WSL environment" | tee -a "$LOG_FILE"
    # When in WSL, skip Windows-only tests by default if no mark specified
    if [ -z "$MARK" ]; then
        MARK="not windows"
    fi
fi

echo "Running pytest with args: ${MARK:+-m \"$MARK\"} $VERBOSE" | tee -a "$LOG_FILE"

# Run pytest with specified options
if [ -n "$MARK" ]; then
    python -m pytest -m "$MARK" $VERBOSE 2>&1 | tee -a "$LOG_FILE"
else
    python -m pytest $VERBOSE 2>&1 | tee -a "$LOG_FILE"
fi

exit ${PIPESTATUS[0]}

# EOF