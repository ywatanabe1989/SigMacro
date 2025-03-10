<!-- ---
!-- Timestamp: 2025-03-09 04:57:38
!-- Author: ywatanabe
!-- File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/vba/README.md
!-- --- -->

# SigmaPlot VBA Utilities

This package provides utilities for managing and executing VBA macros in SigmaPlot through Python. It enables working with SigmaPlot macros in a more modular and maintainable way.

## Features

- Manage VBA macros in separate files
- Load, save, and organize VBA macros
- Create SigmaPlot templates with embedded macros
- Execute macros directly in SigmaPlot
- Standard library of useful macros for common tasks

## Components

- `VBAFileManager`: Manages VBA files on disk
- `VBALibrary`: Library of useful VBA macros with standardized interface
- Pre-built macros for:
  - Data import
  - Plotting
  - Data analysis
  - Exporting results
  - Utility functions

## Requirements

- Python 3.6+
- SigmaPlot 12.0+
- pywin32 package (for COM automation)

## Installation

```bash
pip install pysigmacro
```

## Basic Usage

```python
from pysigmacro.vba import VBALibrary

# Initialize the VBA library
vba_lib = VBALibrary()

# List available macros
print(vba_lib.get_all_macro_names())

# Get a specific macro's code
print(vba_lib.get_macro("data_import"))

# Run a plotting macro with arguments
vba_lib.run_macro("plotting", args=["line", "1"])

# Create a template with an embedded macro
template_path = vba_lib.create_template_with_macro(
    "MyTemplate", 
    "data_analysis", 
    output_path="C:/path/to/save/template.JNB"
)
```

## Extending with Custom Macros

1. Create a .vba or .bas file with your macro code
2. Place it in the VBA macros directory
3. Access it through the VBALibrary interface

## Advanced Usage

See the examples directory for more detailed usage scenarios.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

<!-- EOF -->