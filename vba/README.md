<!-- ---
!-- Timestamp: 2025-03-29 22:49:50
!-- Author: ywatanabe
!-- File: /home/ywatanabe/win/documents/SigMacro/SigMacro/vba/README.md
!-- --- -->

# SigMacro System Architecture

## Data Organization
The architecture follows these principles:
1. Use Python to write data to a worksheet (simple and efficient)
2. VBA macros read this data to create and customize SigmaPlot graphs
3. Data in the worksheet is organized in a structured format:
   - Columns 0-1: Plot type specification with explanations
   - Columns 2-3: Graph parameters with explanations
   - Columns 7 and 14: Reserved for X and Y ticks configuration
   - Column 16 onward: Actual data for plotting

## Data Layout Details
- **Plot Configuration** (cols 0-1): Contains plot type, style, data type and other wizard parameters
- **Graph Parameters** (cols 2-3): Contains axis labels, scales, ranges and other graph settings
- **Plot Data** (col 16+): Organized in chunks of 16 columns per plot series:
  - X values, X errors, X bounds (upper/lower) (cols 0-3 in chunk)
  - Y values, Y errors, Y bounds (upper/lower) (cols 4-7 in chunk)
  - RGBA values for styling (col 8 in chunk)
  - Reserved space for future extensions (cols 9-15 in chunk)

## Benefits
- Self-documenting worksheet with explanation columns
- Compact layout that maximizes worksheet space
- Consistent interface between Python and VBA
- Flexible data organization that supports various plot types
- Easy to extend with additional parameters

## Implementation Notes
The constants defined in the VBA code correspond to this organization, making it easy to access the appropriate data for each plot element. The chunk-based data layout ensures consistent handling of different plot types while maintaining a standardized interface.


``` vba
' ----------------------------------------
' Constants
' ----------------------------------------
Const WORKSHEET_NAME As String = "worksheet"
Const GRAPH_NAME As String = "graph"

' For Plot Wizard
' ----------------------------------------
' Columns
Const _PLOT_TYPE_EXPLANATION_COL As Long = 0
Const PLOT_TYPE_COL As Long = 1
' Rows
Const PLOT_TYPE_ROW As Long = 0
Const PLOT_STYLE_ROW As Long = 1
Const PLOT_DATA_TYPE_ROW As Long = 2
Const _PLOT_COLUMNS_PER_PLOT_ROW As Long = 3 ' Spacer
Const _PLOT_PLOT_COLUMNS_COUNT_ARRAY_ROW As Long = 4 ' Spacer
Const PLOT_DATA_SOURCE_ROW As Long = 5
Const PLOT_POLARUNITS_ROW As Long = 6
Const PLOT_ANGLEUNITS_ROW As Long = 7
Const PLOT_MIN_ANGLE_ROW As Long = 8
Const PLOT_MAX_ANGLE_ROW As Long = 8
Const PLOT_UNKONWN1_ROW As Long = 9
Const PLOT_GROUP_STYLE_ROW As Long = 10
Const PLOT_USE_AUTOMATIC_LEGENDS_ROW As Long = 11

' Graph Parameters
' ----------------------------------------
' Columns
Const _GRAPH_PARAMS_EXPLANATION_COL As Long = 2
Const GRAPH_PARAMS_COL As Long = 3
' Rows
Const X_LABEL_ROW As Long = 0
Const X_MM_ROW As Long = 1
Const X_SCALE_TYPE_ROW As Long = 2
Const X_MIN_ROW As Long = 3
Const X_MAX_ROW As Long = 4
Const _X_TICKS_ROW As Long = 5
Const Y_LABEL_ROW As Long = 6
Const Y_MM_ROW As Long = 7
Const Y_SCALE_TYPE_ROW As Long = 8
Const Y_MIN_ROW As Long = 9
Const Y_MAX_ROW As Long = 10
Const _Y_TICKS_ROW As Long = 11

' Ticks (Not handled by macros but embedded in JNB file)
' ----------------------------------------
Const _X_TICKS_COL As Long = 7
Const _Y_TICKS_COL As Long = 14

' Data Columns
' ----------------------------------------
Const DATA_START_COL As Long = 16
Const DATA_CHUNK_SIZE As Long = 16
Const DATA_MAX_NUM_CHUNKS As Long = 13
CONST DATA_IDX_X AS LONG = 0
CONST DATA_IDX_X_ERR AS LONG = 1
CONST DATA_IDX_X_UPPER AS LONG = 2
CONST DATA_IDX_X_LOWER AS LONG = 3
CONST DATA_IDX_Y AS LONG = 4
CONST DATA_IDX_Y_ERR AS LONG = 5
CONST DATA_IDX_Y_UPPER AS LONG = 6
CONST DATA_IDX_Y_LOWER AS LONG = 7
CONST DATA_IDX_RGBA AS LONG = 8
```

<!-- EOF -->