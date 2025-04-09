Option Explicit

' ========================================
' General Constants
' ========================================
Const GLOBAL_DEBUG_MODE As Boolean = False
' Just for searching
' Const DEBUG_MODE As Boolean = True
' Const DEBUG_MODE As Boolean = False
Const WORKSHEET_NAME As String = "worksheet"
Const GRAPH_NAME As String = "graph"

' ========================================
' Axis Constants
' ========================================
Const AXIS_X As Long = 1
Const AXIS_Y As Long = 2
Const TICK_DIM_H As Long = 1
Const TICK_DIM_V As Long = 2

Const HORIZONTAL As Long = 0
Const VERTICAL As Long = 1
Const MAJOR_TICK_INDEX As Long = 2

' ========================================
' Hide Options
' ========================================
Const HIDE_LEGEND As Long = 0
Const HIDE_TITLE As Long = 0

' ========================================
' Thickness
' ========================================
Const LINE_THICKNESS_INVISIBLE As Variant = &H00000000
Const TICK_THICKNESS_INVISIBLE As Variant = &H00000000
Const TICK_THICKNESS_00008 As Variant = &H00000008
Const POLAR_LINE_THICKNESS As Double = 0.008 * 1000
Const AREA_LINE_THICKNESS As Double = 0
Const LINETYPE_NONE As Long = 1

' ========================================
' Length
' ========================================
Const TICK_LENGTH_00032 As Variant = &H00000020

' ========================================
' Font Size
' ========================================
Const LABEL_PTS_07 As Long = 97
Const LABEL_PTS_08 As Long = 111

' ========================================
' Symbol
' ========================================
Const SSA_SHAPE As Long = &H00000706&

' ========================================
' Colors
' ========================================
Const AREAFILLTYPE_VERTICAL As Long = 1
Const SSA_COLOR_ALPHA As Long = &H000008a7&
Const RGB_WHITE As Long = &H00c0c0c0&
Const RGB_BLACK As Long = &H00000000
Const RGB_LIGHT_GRAY As Long = &H00808080&
Const RGB_DARK_GRAY As Long = &H00c0c0c0&
Const RGB_NONE As Long = &Hff000000&

' ========================================
' Worksheet
' ========================================
Const LABEL_ROW As Long = -1

' Columns
Const _GRAPH_PARAMS_EXPLANATION_COL As Long = 0
Const GRAPH_PARAMS_COL As Long = 1
Const X_TICKS_COL As Long = 2
Const Y_TICKS_COL As Long = 3

' Rows
Const X_LABEL_ROW As Long = 0
Const X_LABEL_ROTATION_ROW As Long = 1
Const X_MM_ROW As Long = 2
Const X_SCALE_TYPE_ROW As Long = 3
Const X_MIN_ROW As Long = 4
Const X_MAX_ROW As Long = 5
Const Y_LABEL_ROW As Long = 6
Const Y_LABEL_ROTATION_ROW As Long = 7
Const Y_MM_ROW As Long = 8
Const Y_SCALE_TYPE_ROW As Long = 9
Const Y_MIN_ROW As Long = 10
Const Y_MAX_ROW As Long = 11

' ========================================
' Graph Wizard-related constants
' ========================================
Const GW_PLOT_TYPE_ROW As Long = 0
Const GW_PLOT_STYLE_ROW As Long = 1
Const GW_DATA_TYPE_ROW As Long = 2
Const _GW_COLUMNS_PER_GW_ROW As Long = 3
Const _GW_GW_COLUMNS_COUNT_ARRAY_ROW As Long = 4
Const GW_DATA_SOURCE_ROW As Long = 5
Const GW_POLARUNITS_ROW As Long = 6
Const GW_ANGLEUNITS_ROW As Long = 7
Const GW_MIN_ANGLE_ROW As Long = 8
Const GW_MAX_ANGLE_ROW As Long = 9
Const GW_UNKONWN1_ROW As Long = 10
Const GW_GROUP_STYLE_ROW As Long = 11
Const GW_USE_AUTOMATIC_LEGENDS_ROW As Long = 12

' For each plot
Const GW_START_COL_BASE_NAME As String = "gw_param_keys "
Const GW_START_COL As Long = -1
Const GW_ID_PARAM_KEYS As Long = 0
Const GW_ID_PARAM_VALUES As Long = 1
Const GW_ID_LABEL As Long = 2
Const GW_ID_RGBA As Long = -1

' ========================================
' Axis Scales
' ========================================
Const SAA_TYPE_LINEAR = 1
Const SAA_TYPE_COMMON = 2
Const SAA_TYPE_LOG = 3
Const SAA_TYPE_PROBABILITY = 4
Const SAA_TYPE_PROBIT = 5
Const SAA_TYPE_LOGIT = 6
Const SAA_TYPE_CATEGORY = 7
Const SAA_TYPE_DATETIME = 8

' ========================================
' Heatmap
' ========================================
Const HEATMAP_SCATTER_ID_Z As Long = 5
Const HEATMAP_SCATTER_ID_SYMBOL As Long = 7

' ========================================
' Violin
' ========================================
Const VIOLIN_BOX_WIDTH_PERC_x_10 As Long = 150

' ========================================
' Histogram
' ========================================
Const HISTOGRAM_BAR_WIDTH_PERC_x_10 As Long = 1000

' ========================================
' Helper Functions
' ========================================
Sub DebugMsg(DEBUG_MODE As Boolean, msg As String)
    If GLOBAL_DEBUG_MODE Or DEBUG_MODE Then
        MsgBox msg, vbInformation, "Debug Info"
    End If
End Sub

Sub DebugType(DEBUG_MODE As Boolean, item)
    If GLOBAL_DEBUG_MODE Or DEBUG_MODE Then
        MsgBox "Type: " & TypeName(item)
    End If
End Sub

Sub Sleep(milliseconds As Long)
    Dim startTime As Double
    Dim endTime As Double

    ' Get start time
    startTime = Timer

    ' Calculate end time
    endTime = startTime + (milliseconds / 1000)

    ' Wait while processing events
    Do
        DoEvents
        ' Check if we've reached the delay time
        If Timer >= endTime Then
            Exit Do
        ElseIf Timer < startTime Then
            ' Handle midnight rollover
            endTime = endTime - 86400
            startTime = 0
        End If
    Loop
End Sub

' Graph
' ----------------------------------------

Function _DoesGraphExist() As Boolean
    Const DEBUG_MODE As Boolean = False
    On Error Resume Next
    Dim graphObj As Object
    Set graphObj = ActiveDocument.NotebookItems(GRAPH_NAME)
    If Not graphObj Is Nothing Then
        graphObj.Open
        Dim tempGraph As Object
        Set tempGraph = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH)
        If Not tempGraph Is Nothing Then
            _DoesGraphExist = True
            Exit Function
        End If
    End If
    _DoesGraphExist = False
End Function

Function _CountPlot() As Long
    Const DEBUG_MODE As Boolean = False
    Dim graphItem As Object
    Set graphItem = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH)
    If Not graphItem Is Nothing Then
        _CountPlot = graphItem.Plots.Count
    Else
        _CountPlot = 0
    End If
    DebugMsg(DEBUG_MODE, "Number of plots: " & _CountPlot)
End Function

Sub _SelectPlot(plotIndex As Long)
    Const DEBUG_MODE As Boolean = False
    ActiveDocument.NotebookItems(GRAPH_NAME).Open
    On Error Resume Next
    Dim plotObj As Object
    Set plotObj = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).Plots(plotIndex)
    If Not plotObj Is Nothing Then
        plotObj.SetObjectCurrent
        If Err.Number <> 0 Then
            DebugMsg(DEBUG_MODE, "Error in _SelectPlot: " & Err.Description)
            Err.Clear
        End If
    Else
        DebugMsg(DEBUG_MODE, "Plot object not found in _SelectPlot for index " & plotIndex)
    End If
End Sub

Function _MmToSigmaPlotUnit(mm As Long)
    Const DEBUG_MODE As Boolean = False
    DebugMsg(DEBUG_MODE, "_MmToSigmaPlotUnit called")
    _MmToSigmaPlotUnit = mm*30
End Function

Sub _ConvertNumToText(numericCol As Long, targetCol As Long)
    Const DEBUG_MODE As Boolean = False
    Dim dataTable As Object
    Dim rowIndex As Long
    Dim nRows As Long
    Dim dataIndex As Long
    Dim maxNRows As Long
    Dim readValue As Variant
    Dim sourceValues As Variant

    Debug.Print "Starting _ConvertNumToText with column: " & numericCol

    ' Open worksheet before accessing data
    ActiveDocument.NotebookItems(WORKSHEET_NAME).Open()
    Set dataTable = ActiveDocument.NotebookItems(WORKSHEET_NAME).DataTable

    ' Find the first empty or last valid row
    maxNRows = 32
    nRows = 4

    For rowIndex = 0 To maxNRows
        readValue = _ReadCell(numericCol, rowIndex)

        If CStr(readValue) = "-1.#QNAN" Then
            DebugMsg(DEBUG_MODE, "readValue: " & readValue)
            DebugType(DEBUG_MODE, "type(readValue): " & readValue)

            nRows = rowIndex
            Exit For
        End If
    Next rowIndex

    ' Get values from the source column
    sourceValues = dataTable.GetData(numericCol, 0, numericCol, nRows - 1)

    ' Create text array with the same dimensions
    Debug.Print "Creating text array"
    Dim textArray() As Variant
    ReDim textArray(0, nRows - 1)

    ' Convert numeric values to formatted text
    For dataIndex = 0 To nRows - 1
        If Not IsEmpty(sourceValues(0, dataIndex)) Then
            textArray(0, dataIndex) = Format(sourceValues(0, dataIndex), "0.00") & " "
            Debug.Print "Row " & dataIndex & ": " & textArray(0, dataIndex)
        End If
    Next dataIndex

    ' Write formatted text to target column
    Debug.Print "Writing data to column " & targetCol
    dataTable.PutData(textArray, targetCol, 0)

    ' Add column title/header
    dataTable.ColumnTitle(targetCol) = "Text"
End Sub

' ========================================
' Reader Functions
' ========================================
Function _ReadCell(columnIndex As Long, rowIndex As Long) As Variant
    Dim dataTable As Object
    Dim cellValue As Variant
    Set dataTable = ActiveDocument.NotebookItems(WORKSHEET_NAME).DataTable
    cellValue = dataTable.GetData(columnIndex, rowIndex, columnIndex, rowIndex)
    _ReadCell = cellValue(0, 0)
End Function

Function _ReadRGB(columnIndex As Long) As Long
    Const DEBUG_MODE As Boolean = False
    ' DebugMsg(DEBUG_MODE, "_ReadRGB called for plot " & columnIndex
    Dim rValue As Variant, gValue As Variant, bValue As Variant
    ' Read RGB values from worksheet (R, G, B values are assumed to be in adjacent columns)
    bValue = _ReadCell(columnIndex, 0)
    gValue = _ReadCell(columnIndex, 1)
    rValue = _ReadCell(columnIndex, 2)
    If (bValue = "None") And (gValue = "None") And (rValue = "None") Then
        _ReadRGB = -1
    Else
        ' Convert to integers and create RGB color
        Dim r As Integer, g As Integer, b As Integer
        b = CInt(bValue)
        g = CInt(gValue)
        r = CInt(rValue)
        ' Standard RGB (VBA default)
        _ReadRGB = RGB(r, g, b)
    End If
End Function

Function _ReadAlphaAsTransparency(columnIndex As Long) As Long
    Const DEBUG_MODE As Boolean = False
    Dim alphaValue As Variant
    Dim transparency As Variant
    alphaValue = _ReadCell(columnIndex, 3)
    If (alphaValue = "None") Then
        _ReadAlphaAsTransparency = -1
    Else
        transparency = (1 - alphaValue) * 100
        _ReadAlphaAsTransparency = transparency
    End If
End Function

Function _ReadPlotTypeStr(iPlot As Long) As String
    Const DEBUG_MODE As Boolean = False
    Dim startCol As Long, valuesCol As Long, labelCol As Long
    Dim plotType As String
    Dim spacePos As Long

    startCol = _FindChunkStartCol(iPlot)
    If startCol <> -1 Then
        labelCol = startCol + GW_ID_LABEL
        plotType = _ReadCell(labelCol, 0)

        ' Extract base type by removing any trailing numbers
        spacePos = InStr(plotType, " ")
        If spacePos > 0 Then
            plotType = Left(plotType, spacePos - 1)
        End If

        _ReadPlotTypeStr = plotType
    Else
        _ReadPlotTypeStr = "line"
    End If
End Function

' ========================================
' Color-related Functions
' ========================================
Function _GenRGB(r As Long, g As Long, b As Long) As Long
    _GenRGB = RGB(r, g, b)
End Function

Sub _CreateColorColumn(rColumn As Long, gColumn As Long, bColumn As Long, resultColumn As Long)
    Dim sep As String
    sep = ListSeparator

    Dim SPTransform As Object
    Set SPTransform = ActiveDocument.NotebookItems.Add(9)
    SPTransform.Open

    ' Simple transform to create color from RGB values
    SPTransform.Text = "col(" & resultColumn & ") = rgbcolor(col(" & rColumn & ")" & _
               sep & "col(" & gColumn & ")" & sep & "col(" & bColumn & "))"

    ' Execute transform
    SPTransform.Execute
    SPTransform.Close(False)

    ' Add column title
    ActiveDocument.CurrentDataItem.DataTable.NamedRanges.Add _
               "Color", resultColumn-1, 0, 1, -1, True
End Sub

Function _ReadGWColumnMapping(plotType As String, startCol As Long, endCol As Long) As Variant
    Const DEBUG_MODE As Boolean = False
    Dim mapping()

    ' Data Columns
    Dim nDataCols As Long
    Const nHeadCols As Long = 3
    Const nTailCols As Long = 1

    nDataCols = (endCol - startCol + 1) - (nHeadCols + nTailCols)

    If plotType = "scatter_heatmap" Then
        nDataCols = 2
    End If

    ReDim mapping(2, nDataCols)

    Dim iCol As Long
    For iCol = 0 To nDataCols
        mapping(0, iCol) = startCol + nHeadCols + iCol
    Next iCol

    ' Fill in the row ranges for all columns
    Dim ii As Integer
    For ii = 0 To UBound(mapping, 2)
        mapping(1, ii) = 0
        mapping(2, ii) = 31999999
    Next ii

    _ReadGWColumnMapping = mapping
End Function

Function _ReadGWPlotCountColumnArray(plotType As String, startCol As Long, endCol As Long) As Variant
    Const DEBUG_MODE As Boolean = False

    Dim countArray()
    ReDim countArray(0)

    ' Data Columns
    Dim nDataCols As Long
    Const nHeadCols As Long = 3
    Const nTailCols As Long = 1

    nDataCols = (endCol - startCol + 1) - (nHeadCols + nTailCols)

    ' Fixme; the third columns is symbol ...
    If plotType = "scatter_heatmap" Then
        nDataCols = 2
    End If

    DebugMsg(DEBUG_MODE, "_ReadGWPlotCountColumnArray called")
    DebugMsg(DEBUG_MODE, "startCol: " & startCol)
    DebugMsg(DEBUG_MODE, "endCol: " & endCol)
    DebugMsg(DEBUG_MODE, "nDataCols: " & nDataCols)

    ' ReDim countArray(0)
    countArray(0) = nDataCols

    _ReadGWPlotCountColumnArray = countArray
End Function

' ========================================
' Finder Functions
' ========================================
Function _FindMaxCol() As Long
    Const DEBUG_MODE As Boolean = False
    Dim maxCol As Long, maxRow As Long, dataTable As Object
    Set dataTable = ActiveDocument.NotebookItems(WORKSHEET_NAME).DataTable
    DataTable.GetMaxUsedSize(maxCol, maxRow)
    _FindMaxCol = maxCol
End Function

Function _FindMaxRow() As Long
    Const DEBUG_MODE As Boolean = False
    Dim maxCol As Long, maxRow As Long, dataTable As Object
    Set dataTable = ActiveDocument.NotebookItems(WORKSHEET_NAME).DataTable
    DataTable.GetMaxUsedSize(maxCol, maxRow)
    _FindMaxRow = maxRow
End Function

Function _FindColIdx(columnName As String) As Long
    Const DEBUG_MODE As Boolean = False
    Dim maxCol As Long, ColIndex As Long, ColName As String, ii As Long
    maxCol = _FindMaxCol()
    ColIndex = -1
    For ii = 0 To maxCol
        ColName = _ReadCell(ii, LABEL_ROW)
        If LCase(ColName) = LCase(columnName) Then
            ColIndex = ii
            Exit For
        End If
    Next ii
    _FindColIdx = ColIndex
End Function

Function _FindChunkStartCol(iPlot As Long) As Long
    Const DEBUG_MODE As Boolean = False
    Dim colName As String
    colName = GW_START_COL_BASE_NAME & iPlot
    _FindChunkStartCol = _FindColIdx(colName)
End Function

Function _FindChunkEndCol(iPlot As Long) As Long
    Const DEBUG_MODE As Boolean = False
    Dim startCol As Long, nextStartCol As Long
    Dim maxCol As Long

    startCol = _FindChunkStartCol(iPlot)
    If startCol = -1 Then
        _FindChunkEndCol = -1
        Exit Function
    End If

    nextStartCol = _FindChunkStartCol(iPlot + 1)
    If nextStartCol = -1 Then
        maxCol = _FindMaxCol()
        _FindChunkEndCol = maxCol - 1
    Else
        _FindChunkEndCol = nextStartCol - 1
    End If
End Function

Function _FindNumPlots() As Long
    Const DEBUG_MODE As Boolean = False
    Dim iCol As Long
    Dim count As Long
    Dim maxCol As Long
    maxCol = _FindMaxCol()
    count = 0
    For iCol = 0 To maxCol
        If _FindChunkStartCol(iCol) <> -1 Then
            count = count + 1
        Else
            ' No more chunks found, exit loop
            Exit For
        End If
    Next iCol
    _FindNumPlots = count
    DebugMsg(DEBUG_MODE, "Found " & count & " plot chunks")
End Function

' ========================================
' Plot
' ========================================
Sub Plot()
    Const DEBUG_MODE As Boolean = False
    ' Open the worksheet
    ActiveDocument.NotebookItems(WORKSHEET_NAME).Open

    ' Get the number of plots
    Dim numPlots As Long
    numPlots = _FindNumPlots()

    Dim graphAlreadyExists As Boolean
    graphAlreadyExists = _DoesGraphExist()

    ' Loop through all plot types
    Dim iPlot As Long
    For iPlot = 0 To numPlots - 1

        ' Find the start and end columns for this plot type
        Dim startCol As Long, endCol As Long
        startCol = _FindChunkStartCol(iPlot)

        ' If no more plot chunks found, exit loop
        If startCol = -1 Then
            DebugMsg(DEBUG_MODE, "No plot chunks found")
            Exit For
        End If

        endCol = _FindChunkEndCol(iPlot)
        DebugMsg(DEBUG_MODE, "Plot " & iPlot & " columns: " & startCol & " to " & endCol)

        ' Read GW parameters for this plot
        Dim gwPlotType As String, gwPlotStyle As String, gwDataType As String
        Dim gwDataSource As String, gwPolarUnits As String, gwAngleUnits As String
        Dim gwMinAngle As Double, gwMaxAngle As Double, gwGroupStyle As String
        Dim gwUseAutomaticLegends As Boolean, gwUnknown1 As Variant

        ' Read parameters from the param_keys and param_values columns
        Dim gwValuesCol As Long
        gwValuesCol = startCol + 1

        Dim plotType As String
        plotType = _ReadPlotTypeStr(iPlot)

        ' Get type and style based on plot index
        gwPlotType = _ReadCell(gwValuesCol, GW_PLOT_TYPE_ROW)
        gwPlotStyle = _ReadCell(gwValuesCol, GW_PLOT_STYLE_ROW)
        gwDataType = _ReadCell(gwValuesCol, GW_DATA_TYPE_ROW)
        gwDataSource = _ReadCell(gwValuesCol, GW_DATA_SOURCE_ROW)
        gwPolarUnits = _ReadCell(gwValuesCol, GW_POLARUNITS_ROW)
        gwAngleUnits = _ReadCell(gwValuesCol, GW_ANGLEUNITS_ROW)
        gwMinAngle = CDbl(_ReadCell(gwValuesCol, GW_MIN_ANGLE_ROW))
        gwMaxAngle = CDbl(_ReadCell(gwValuesCol, GW_MAX_ANGLE_ROW))
        gwUnknown1 = _ReadCell(gwValuesCol, GW_UNKONWN1_ROW)
        gwGroupStyle = _ReadCell(gwValuesCol, GW_GROUP_STYLE_ROW)
        gwUseAutomaticLegends = CBool(_ReadCell(gwValuesCol, GW_USE_AUTOMATIC_LEGENDS_ROW))

        ' Build column mapping based on the plot type
        Dim gwColumnsPerPlot() As Variant
        gwColumnsPerPlot = _ReadGWColumnMapping(plotType, startCol, endCol)

        ' Get the column count array
        Dim gwPlotColumnCountArray() As Variant
        gwPlotColumnCountArray = _ReadGWPlotCountColumnArray(plotType, startCol, endCol)

        ' Create or add plot
        If Not graphAlreadyExists And iPlot = 0 Then
            DebugMsg(DEBUG_MODE, "Creating new graph...")
            ActiveDocument.CurrentPageItem.CreateWizardGraph(gwPlotType, _
                                                             gwPlotStyle, _
                                                             gwDataType, _
                                                             gwColumnsPerPlot, _
                                                             gwPlotColumnCountArray, _
                                                             gwDataSource, _
                                                             gwPolarUnits, _
                                                             gwAngleUnits, _
                                                             gwMinAngle, _
                                                             gwMaxAngle, _
                                                             , _
                                                             gwGroupStyle, _
                                                             gwUseAutomaticLegends)
            graphAlreadyExists = True
        Else
            ActiveDocument.NotebookItems(GRAPH_NAME).Open
            ActiveDocument.CurrentPageItem.AddWizardPlot(gwPlotType, _
                                                         gwPlotStyle, _
                                                         gwDataType, _
                                                         gwColumnsPerPlot, _
                                                         gwPlotColumnCountArray, _
                                                         gwDataSource, _
                                                         gwPolarUnits, _
                                                         gwAngleUnits, _
                                                         gwMinAngle, _
                                                         gwMaxAngle, _
                                                         , _
                                                         gwGroupStyle, _
                                                         gwUseAutomaticLegends)
            DebugMsg(DEBUG_MODE, "Plot added to existing graph")
        End If
    Next iPlot
    ActiveDocument.NotebookItems(GRAPH_NAME).Open
End Sub

' ========================================
' Remover Functions
' ========================================
Sub RemoveExistingGraphs()
    Const DEBUG_MODE As Boolean = False
    On Error Resume Next
    ActiveDocument.NotebookItems(GRAPH_NAME).Open
    ActiveDocument.CurrentItem.SelectAll
    ActiveDocument.CurrentItem.Clear
End Sub

Sub RemoveLegend()
    Const DEBUG_MODE As Boolean = False
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETGRAPHATTR, SGA_AUTOLEGENDSHOW, HIDE_LEGEND)
End Sub

Sub RemoveTopSpine()
    Const DEBUG_MODE As Boolean = False
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_X)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SUB2OPTIONS, TICK_THICKNESS_INVISIBLE)
End Sub

Sub RemoveRightSpine()
    Const DEBUG_MODE As Boolean = False
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_Y)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SUB2OPTIONS, TICK_THICKNESS_INVISIBLE)
End Sub

Sub _RemoveLineButTicks(axisDim As Long)
    Const DEBUG_MODE As Boolean = False
    With ActiveDocument.CurrentPageItem
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, axisDim)
        ' Select the axis Line
        .SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, 1)
        ' Set graph flags (affects how axis components are displayed)
        .SetCurrentObjectAttribute(GPM_SETGRAPHATTR, SGA_FLAGS, &H00010001&)
        ' Set axis color (still needed even with thickness 0)
        .SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_COLOR, &Hff000000&)
        ' Set axis line thickness to 0 (invisible)
        .SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_THICKNESS, 0)
        ' Set position parameters (anchors the axis)
        .SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_POS1PERMILL, 0)
        .SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_POS2PERMILL, 0)
        ' Set sub-options controlling ticks and other elements
        .SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SUB1OPTIONS, &H0000000d&)
        .SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SUB2OPTIONS, &H0000000e&)
    End With
End Sub

Sub RemoveLeftLine()
    _RemoveLineButTicks(AXIS_Y)
End Sub

Sub RemoveBottomLine()
    _RemoveLineButTicks(AXIS_X)
End Sub

Sub RemoveTitle()
    Const DEBUG_MODE As Boolean = False
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETGRAPHATTR, SGA_SHOWNAME, 0)
End Sub

' ========================================
' Color Applying Functions
' ========================================
Sub ApplyColors()
    Const DEBUG_MODE As Boolean = False
    On Error GoTo ErrorHandler
    Dim plotCount As Long
    Dim iPlot As Long
    Dim colorColumn As Long
    Dim RGB_VAL As Long
    Dim transparencyVAL As Long
    Dim graphItem As Object
    Dim plotObj As Object
    Dim plotType As String

    ' Get the graph page
    Set graphItem = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH)
    If graphItem Is Nothing Then
        DebugMsg(DEBUG_MODE, "Error: Graph object not found")
        Exit Sub
    End If

    ' Get the number of plots
    plotCount = graphItem.Plots.Count

    ' Loop through all plots
    For iPlot = 0 To plotCount - 1
        plotType = LCase(_ReadPlotTypeStr(iPlot))
        colorColumn = _FindChunkEndCol(iPlot)

        RGB_VAL = _ReadRGB(colorColumn)
        transparencyVal = _ReadAlphaAsTransparency(colorColumn)

        ' Apply color based on plot type
        Select Case plotType
            Case "area", "area_heatmap"
                _ApplyColorArea(iPlot, RGB_VAL, transparencyVal)
            Case "bar", "barh", "barh_heatmap", "histogram"
                _ApplyColorBar(iPlot, RGB_VAL, transparencyVal)
            Case "box", "boxh"
                _ApplyColorBox(iPlot, RGB_VAL, transparencyVal)
            Case "line", "line_yerr", "lines_y_many_x", "lines_x_many_y"
                _ApplyColorLine(iPlot, RGB_VAL, transparencyVal)

               ' Violine
            Case "box_violin", "box_violinh"
                _ApplyColorViolinBox(iPlot, RGB_VAL, transparencyVal)
            Case "lines_y_many_x_violin", "lines_x_many_y_violinh"
                _ApplyColorViolinLine(iPlot, RGB_VAL, transparencyVal)

            Case "polar"
                _ApplyColorPolar(iPlot, RGB_VAL, transparencyVal)
            Case "scatter", "jitter"
                _ApplyColorScatter(iPlot, RGB_VAL, transparencyVal)
            Case "scatter_heatmap"
                _ApplyColorScatter(iPlot, RGB_VAL, transparencyVal)

            ' == Start Filled Line Handling ==
           Case "filled_line_uu"
                ' _ApplyColorArea(iPlot, RGB_VAL, transparencyVal)

               _ApplyColorFilledLineUpper(iPlot, RGB_VAL, transparencyVal)
           Case "filled_line_mm"
                ' _ApplyColorArea(iPlot, RGB_VAL, transparencyVal)
               _ApplyColorFilledLineMiddle(iPlot, RGB_VAL, transparencyVal)
           Case "filled_line_ll"
                ' _ApplyColorArea(iPlot, RGB_VAL, transparencyVal)
               _ApplyColorFilledLineLower(iPlot, RGB_VAL, transparencyVal)
            ' == End Filled Line Handling ==
           Case "3dscatter"
                _ApplyColor3DScatter(iPlot, RGB_VAL, transparencyVal)
           Case "contour", "heatmap"
                _ApplyColorFake(iPlot, RGB_VAL, transparencyVal)
        End Select
    Next iPlot
    Exit Sub

ErrorHandler:
    DebugMsg(DEBUG_MODE, "Error in Main: " & Err.Description)
End Sub

Sub _ApplyColorArea(iPlot As Long, RGB_VAL As Long, transparencyVal As Long)
    Const DEBUG_MODE As Boolean = False
    DebugMsg(DEBUG_MODE, "_ApplyColorArea called")
    _SelectPlot(iPlot)
    With ActiveDocument.CurrentPageItem
    .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLOR, RGB_VAL)
    .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, RGB_VAL)
    .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_AREAFILLTYPE, AREAFILLTYPE_VERTICAL)
    .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR_ALPHA, transparencyVal)
    End With
End Sub

Sub _ApplyColorBar(iPlot As Long, RGB_VAL As Long, transparencyVal As Long)
    Const DEBUG_MODE As Boolean = False
    DebugMsg(DEBUG_MODE, "_ApplyColorBar called")
    _SelectPlot(iPlot)
    With ActiveDocument.CurrentPageItem
    .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, RGB_VAL)
    .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_EDGECOLOR, RGB_BLACK)
    .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR_ALPHA, transparencyVal)
    End With
End Sub

Sub _ApplyColorBox(iPlot As Long, RGB_VAL As Long, transparencyVal As Long)
    Const DEBUG_MODE As Boolean = False
    DebugMsg(DEBUG_MODE, "_ApplyColorBox called")
    _SelectPlot(iPlot)
    With ActiveDocument.CurrentPageItem
    .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, RGB_VAL)
    .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_EDGECOLOR, RGB_BLACK)
    End With
End Sub

Sub _ApplyColorLine(iPlot As Long, RGB_VAL As Long, transparencyVal As Long)
    Const DEBUG_MODE As Boolean = False
    DebugMsg(DEBUG_MODE, "_ApplyColorLine called")
    _SelectPlot(iPlot)
    With ActiveDocument.CurrentPageItem
        ' Line
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLORCOL, -2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SOA_COLOR, RGB_VAL)
        ' Symbol
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_EDGECOLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLORREPEAT, 2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR_ALPHA, transparencyVal)
        ' Solid
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_EDGECOLOR, RGB_VAL)
        ' Error bar
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_ERRCOLOR, RGB_VAL)
    End With
End Sub

Sub _ApplyColorFilledLineUpper(iPlot As Long, RGB_VAL As Long, transparencyVal As Long)
    Const DEBUG_MODE As Boolean = False
    DebugMsg(DEBUG_MODE, "_ApplyColorFilledLineUpper called")
    _SelectPlot(iPlot)
    With ActiveDocument.CurrentPageItem
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_LINETYPE, LINETYPE_NONE)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_EDGECOLOR, &H00000000&)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_EDGECOLORREPEAT, 2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, &H00000362&, 0)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLORCOL, -2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLORREPEAT, 2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_AREAFILLTYPE, AREAFILLTYPE_VERTICAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_EDGECOLOR, &H00000000&)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, &H000008a7&, 0)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, &H000008a8&, 0)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR_ALPHA, transparencyVal)
    End With
End Sub

Sub _ApplyColorFilledLineMiddle(iPlot As Long, RGB_VAL As Long, transparencyVal As Long)
    Const DEBUG_MODE As Boolean = False
    DebugMsg(DEBUG_MODE, "_ApplyColorFilledLineMiddle called")
    _SelectPlot(iPlot)
    With ActiveDocument.CurrentPageItem
        ' Line
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLORCOL, -2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SOA_COLOR, RGB_VAL)
    End With
End Sub

Sub _ApplyColorFilledLineLower(iPlot As Long, RGB_VAL As Long, transparencyVal As Long)
    Const DEBUG_MODE As Boolean = False
    DebugMsg(DEBUG_MODE, "_ApplyColorFilledLineLower called")
    _SelectPlot(iPlot)
    With ActiveDocument.CurrentPageItem
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_LINETYPE, LINETYPE_NONE)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_EDGECOLOR, &H00000000&)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_EDGECOLORREPEAT, 2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, &H00000362&, 0)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLORCOL, -2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLORREPEAT, 2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_AREAFILLTYPE, AREAFILLTYPE_VERTICAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_EDGECOLOR, &H00000000&)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, &H000008a7&, 0)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, &H000008a8&, 0)
    End With
End Sub

Sub _ApplyColorViolinBox(iPlot As Long, RGB_VAL As Long, transparencyVal As Long)
    Const DEBUG_MODE As Boolean = False
    DebugMsg(DEBUG_MODE, "_ApplyColorViolinBox called")
    _SelectPlot(iPlot)
    With ActiveDocument.CurrentPageItem
       ' .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR, _GenRGB(200,200,200))
       ' .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, _GenRGB(200,200,200))
       .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR, RGB_WHITE)
       .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, RGB_WHITE)
       .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_EDGECOLOR, RGB_BLACK)
       .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR_ALPHA, transparencyVal)
    End With
End Sub

Sub _ApplyColorViolinLine(iPlot As Long, RGB_VAL As Long, transparencyVal As Long)
    Const DEBUG_MODE As Boolean = False
    DebugMsg(DEBUG_MODE, "_ApplyColorLine called")
    _SelectPlot(iPlot)
    With ActiveDocument.CurrentPageItem
       .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, RGB_VAL)
       .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLORCOL, -2)
       .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLORREPEAT, 2)
       .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_AREAFILLTYPE, 1)
       .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_EDGECOLOR, RGB_BLACK)
       .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR_ALPHA, transparencyVal)
    End With
End Sub

Sub _ApplyColorFilledLineFill(iPlot As Long, RGB_VAL As Long, transparencyVal As Long)
Const DEBUG_MODE As Boolean = False
    DebugMsg(DEBUG_MODE, "_ApplyColorFilledLineFill called")
    _SelectPlot(iPlot)
    With ActiveDocument.CurrentPageItem
       .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, RGB_VAL)
       .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLORCOL, -2)
       .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLORREPEAT, 2)
       .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_AREAFILLTYPE, 1)
       .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_EDGECOLOR, RGB_BLACK)
    End With
End Sub

Sub _ApplyColorPolar(iPlot As Long, RGB_VAL As Long, transparencyVal As Long)
    Const DEBUG_MODE As Boolean = False
    DebugMsg(DEBUG_MODE, "_ApplyColorPolar called")
    _SelectPlot(iPlot)
    With ActiveDocument.CurrentPageItem
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLORCOL, -2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SOA_COLOR, RGB_VAL)
    End With
End Sub

Sub _ApplyColorScatter(iPlot As Long, RGB_VAL As Long, transparencyVal As Long)
    Const DEBUG_MODE As Boolean = False
    DebugMsg(DEBUG_MODE, "_ApplyColorScatter called")
    _SelectPlot(iPlot)
    With ActiveDocument.CurrentPageItem
        ' Line attributes
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLORCOL, -2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SOA_COLOR, RGB_VAL)
        ' Symbol attributes
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_EDGECOLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLORREPEAT, 2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR_ALPHA, transparencyVal)
        ' Solid attributes
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_EDGECOLOR, RGB_VAL)
        ' Error bar attributes
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_ERRCOLOR, RGB_VAL)
    End With
End Sub


Sub _ApplyColor3DScatter(iPlot As Long, RGB_VAL As Long, transparencyVal As Long)
    Const DEBUG_MODE As Boolean = False
    DebugMsg(DEBUG_MODE, "_ApplyColor3DScatter called")
    _SelectPlot(iPlot)
    With ActiveDocument.CurrentPageItem
        ' Line attributes
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLORCOL, -2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SOA_COLOR, RGB_VAL)
        ' Symbol attributes
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_EDGECOLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLORREPEAT, 2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR_ALPHA, transparencyVal)
    End With
End Sub

Sub _ApplyColorFake(iPlot As Long, RGB_VAL As Long, transparencyVal As Long)
    Const DEBUG_MODE As Boolean = False
    DebugMsg(DEBUG_MODE, "_ApplyColorFake called")
End Sub

' ========================================
' Figure Size
' ========================================
Sub _SetWidth()
    Const DEBUG_MODE As Boolean = False
    DebugMsg(DEBUG_MODE, "_SetWidth called")
    On Error Resume Next
    Dim xLength_mm As Double
    Dim xLength_sp As Double
    xLength_mm = _ReadCell(GRAPH_PARAMS_COL, X_MM_ROW)
    xLength_sp = _MmToSigmaPlotUnit(xLength_mm)
    With ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH)
        .Width = xLength_sp
    End With
End Sub

Sub _SetHeight()
    Const DEBUG_MODE As Boolean = False
    On Error Resume Next
    Dim yLength_mm As Double
    Dim yLength_sp As Double
    yLength_mm = _ReadCell(GRAPH_PARAMS_COL, Y_MM_ROW)
    yLength_sp = _MmToSigmaPlotUnit(yLength_mm)
    With ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH)
       .Height = yLength_sp
    End With
End Sub

Sub SetFigureSize()
    Const DEBUG_MODE As Boolean = False
    On Error Resume Next
    _SetWidth()
    _SetHeight()
End Sub

' ========================================
' Label Text
' ========================================
Function _SetLabelText(axisDim As Variant, labelCol As Long)
    Const DEBUG_MODE As Boolean = False
    Dim Label As Variant
    label = _ReadCell(GRAPH_PARAMS_COL, labelCol)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, axisDim)
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_RTFNAME, label)
End Function

Sub SetLabelsText()
    Const DEBUG_MODE As Boolean = False
    _SetLabelText(AXIS_X, X_LABEL_ROW)
    _SetLabelText(AXIS_Y, Y_LABEL_ROW)
End Sub

' ========================================
' Range
' ========================================
Sub _SetRange(axisDim As Long, scaleTypeRow As Long, minRow As Long, maxRow As Long)
    Const DEBUG_MODE As Boolean = False
    Dim axisMin As String
    Dim axisMax As String
    Dim axisScaleType As Variant
    Const USE_CONSTANT_VALUE As Integer = 10

    ' Get the scale type for the specified axis
    axisScaleType = _ReadCell(GRAPH_PARAMS_COL, scaleTypeRow)

    ' Skip range setting for category or datetime axes
    Select Case LCase(CStr(axisScaleType))
        Case "category", "7", "datetime", "date", "time", "8"
           DebugMsg(DEBUG_MODE, "Skipping range setting for axis " & axisDim & _
                    " due to scale type: " & axisScaleType)
            Exit Sub
    End Select

    ' Read min and max values from the worksheet
    axisMin = _ReadCell(GRAPH_PARAMS_COL, minRow)
    axisMax = _ReadCell(GRAPH_PARAMS_COL, maxRow)

    DebugMsg(DEBUG_MODE, "Setting range for axis " & axisDim & ": Min=" & axisMin & ", Max=" & axisMax)

    ' Select the correct axis object
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, axisDim)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_OPTIONS, USE_CONSTANT_VALUE)

    ' Temporarily ignore errors during attribute setting
    On Error Resume Next

    ' Set the 'From' value if provided
    If LCase(axisMin) <> "none" And axisMin <> "" Then
        ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTRSTRING, SAA_FROMVAL, CStr(axisMin))
        DebugMsg(DEBUG_MODE, "Attempted to set Min value for axis " & axisDim & " to " & axisMin)
    Else
        DebugMsg(DEBUG_MODE, "Min value 'None' for axis " & axisDim & ", skipping.")
    End If

    ' Set the 'To' value if provided
    If LCase(axisMax) <> "none" And axisMax <> "" Then
        ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTRSTRING, SAA_TOVAL, CStr(axisMax))
        DebugMsg(DEBUG_MODE, "Attempted to set Max value for axis " & axisDim & " to " & axisMax)
    Else
        DebugMsg(DEBUG_MODE, "Max value 'None' for axis " & axisDim & ", skipping.")
    End If

    ' Restore default error handling
    On Error GoTo 0

End Sub

Sub SetRanges()
    Const DEBUG_MODE As Boolean = False
    DebugMsg(DEBUG_MODE, "Setting X and Y ranges...")
    _SetRange(AXIS_X, X_SCALE_TYPE_ROW, X_MIN_ROW, X_MAX_ROW)
    _SetRange(AXIS_Y, Y_SCALE_TYPE_ROW, Y_MIN_ROW, Y_MAX_ROW)
End Sub

' ========================================
' Scales
' ========================================
Function _SetScaleType(axisIndex As Long, scaleType As Long)
    Dim axis As Object
    ' Get the axis object directly
    Set axis = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Axes(axisIndex)
    ' Set scale type
    axis.SetAttribute(SAA_TYPE, scaleType)
End Function

Function _cvtScaleTypeFromVariantToLong(cellValue As Variant) As Long
    Const DEBUG_MODE As Boolean = False
    Const SAA_TYPE_LINEAR = 1
    Const SAA_TYPE_COMMON = 2
    Const SAA_TYPE_LOG = 3
    Const SAA_TYPE_PROBABILITY = 4
    Const SAA_TYPE_PROBIT = 5
    Const SAA_TYPE_LOGIT = 6
    Const SAA_TYPE_CATEGORY = 7
    Const SAA_TYPE_DATETIME = 8
    Dim scaleType As Long
    ' Convert string or number to appropriate scale type constant
    Select Case CStr(LCase(cellValue))
        Case "linear", "1"
            scaleType = SAA_TYPE_LINEAR
        Case "common", "common log", "2"
            scaleType = SAA_TYPE_COMMON
        Case "log", "natural log", "3"
            scaleType = SAA_TYPE_LOG
        Case "probability", "4"
            scaleType = SAA_TYPE_PROBABILITY
        Case "probit", "5"
            scaleType = SAA_TYPE_PROBIT
        Case "logit", "6"
            scaleType = SAA_TYPE_LOGIT
        Case "category", "7"
            scaleType = SAA_TYPE_CATEGORY
        Case "datetime", "date", "time", "8"
            scaleType = SAA_TYPE_DATETIME
        Case Else
            ' Default to linear if unrecognized
            scaleType = SAA_TYPE_LINEAR
    End Select
    _cvtScaleTypeFromVariantToLong = scaleType
End Function

Sub _SetXScale()
    Const DEBUG_MODE As Boolean = False
    On Error Resume Next
    Dim xScaleVariant As Variant
    Dim xScaleLong As Long

    xScaleVariant = _ReadCell(GRAPH_PARAMS_COL, X_SCALE_TYPE_ROW)
    xScaleLong = _cvtScaleTypeFromVariantToLong(xScaleVariant)

    _SetScaleType(HORIZONTAL, xScaleLong)
    On Error GoTo 0
End Sub

Sub _SetYScale()
    Const DEBUG_MODE As Boolean = False
    On Error Resume Next
    Dim yScaleVariant As Variant
    Dim yScaleLong As Long

    yScaleVariant = _ReadCell(GRAPH_PARAMS_COL, Y_SCALE_TYPE_ROW)
    yScaleLong = _cvtScaleTypeFromVariantToLong(yScaleVariant)
    _SetScaleType(VERTICAL, yScaleLong)

    On Error GoTo 0
End Sub

Sub SetScales()
    Const DEBUG_MODE As Boolean = False
    _SetXScale()
    _SetYScale()
End Sub

' ========================================
' Tick Positions
' ========================================
Sub _SetXTickPositions()
    Const DEBUG_MODE As Boolean = False
    Dim xTicksFirstRow As Variant
    xTicksFirstRow = _ReadCell(X_TICKS_COL, 0)
    If Not (xTicksFirstRow = "None" Or xTicksFirstRow = "auto") Then
        ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
        ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_X)
        ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SAA_TICCOLUSED, 1)
        ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SAA_TICCOL, X_TICKS_COL)
    End If
End Sub

Sub _SetYTickPositions()
    Const DEBUG_MODE As Boolean = False
    Dim yTicksFirstRow As Variant
    yTicksFirstRow = _ReadCell(Y_TICKS_COL, 0)
    If Not (yTicksFirstRow = "None" Or yTicksFirstRow = "auto") Then
        ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
        ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_Y)
        ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SAA_TICCOLUSED, 1)
        ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SAA_TICCOL, Y_TICKS_COL)
    End If
End Sub

Sub SetTickPositions()
    Const DEBUG_MODE As Boolean = False
    _SetXTickPositions()
    _SetYTickPositions()
End Sub

' ========================================
' Tick Sizes
' ========================================
Sub _SetTickSize(axisDim As Long)
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent()
    With ActiveDocument.CurrentPageItem
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, axisDim)
        .SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, TICK_DIM_H)
        .SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_THICKNESS, TICK_THICKNESS_00008)
        .SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_TICSIZE, TICK_LENGTH_00032)
        .SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, TICK_DIM_V)
        .SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_THICKNESS, TICK_THICKNESS_00008)
        .SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_TICSIZE, TICK_LENGTH_00032)
    End With
End Sub

Sub SetTickSizes()
    _SetTickSize(AXIS_X)
    _SetTickSize(AXIS_Y)
End Sub

' ========================================
' XY and Tick Sizes
' ========================================
Sub _SetLabelSizes(direction As Long)
    Const DEBUG_MODE As Boolean = False
    Dim oAxis As Object
    Dim oText As Object
    Dim oTextTick As Object

    Set oAxis = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).Axes(direction)
    Set oText = oAxis.AxisTitles(0)
    Set oTextTick = oAxis.TickLabelAttributes(MAJOR_TICK_INDEX)

    oText.SetAttribute(STA_SELECT, -65536)
    oText.SetAttribute(STA_SIZE, LABEL_PTS_08)
    oTextTick.SetAttribute(STA_SIZE, LABEL_PTS_07)
End Sub

Sub SetLabelSizes()
    _SetLabelSizes(VERTICAL)
    _SetLabelSizes(HORIZONTAL)
End Sub

' ========================================
' Label Rotation
' ' ========================================
' Sub _SetTickLabelRotation(axisDim As Long, labelRotationRow As Long)
'     On Error Resume Next
'     Dim oTextTick As Object
'     Dim rotationDegrees As Long

'     rotationDegrees = CLng(_ReadCell(GRAPH_PARAMS_COL, labelRotationRow))
'     Set oTextTick = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).Axes(axisDim).AxisTitles(0).TickLabelAttributes(MAJOR_TICK_INDEX)

'     oTextTick.SetAttribute(STA_ORIENTATION, rotationDegrees * 10)
'     On Error GoTo 0
' End Sub

' Sub SetTickLabelRotation()
'     _SetTickLabelRotation(HORIZONTAL, X_LABEL_ROTATION_ROW)
'     _SetTickLabelRotation(VERTICAL, Y_LABEL_ROTATION_ROW)
' End Sub

Sub SetTickLabelRotation()
    Dim xRotation As Long
    Dim yRotation As Long
    Dim oGraph As Object
    Dim oAxisX As Object
    Dim oAxisY As Object
    Dim oTextXTick As Object
    Dim oTextYTick As Object

    ' Default rotations (0 degrees)
    xRotation = 0
    yRotation = 0

    ' Try to read rotation values from worksheet if available
    On Error Resume Next
    ' Assuming rotation values might be stored in cells next to the axis properties
    xRotation = CLng(_ReadCell(GRAPH_PARAMS_COL, X_LABEL_ROTATION_ROW)) * 10
    yRotation = CLng(_ReadCell(GRAPH_PARAMS_COL, Y_LABEL_ROTATION_ROW)) * 10
    On Error GoTo 0

    ' Set the tick label rotation
    Set oGraph = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH)
    Set oAxisX = oGraph.Axes(HORIZONTAL)
    Set oAxisY = oGraph.Axes(VERTICAL)
    Set oTextXTick = oAxisX.TickLabelAttributes(MAJOR_TICK_INDEX)
    Set oTextYTick = oAxisY.TickLabelAttributes(MAJOR_TICK_INDEX)

    ' Apply rotation values
    oTextXTick.SetAttribute(STA_ORIENTATION, xRotation)
    oTextYTick.SetAttribute(STA_ORIENTATION, yRotation)
End Sub

' ========================================
' For Special Cases
' ========================================
Sub HandleSpecialCases()
    Const DEBUG_MODE As Boolean = False
    Dim plotCount As Long
    Dim iPlot As Long
    Dim plotType As String
    plotCount = _CountPlot()

    For iPlot = 0 To plotCount - 1
        plotType = _ReadPlotTypeStr(iPlot)
        _SelectPlot(iPlot)

        Select Case plotType
            Case "area"
                ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, _
                                                                         SEA_LINETYPE, LINETYPE_NONE)
            Case "scatter_heatmap"
                RemoveBottomLine()
                RemoveLeftLine()
                SetTextAsSymbol(iPlot)

            Case "polar"
                ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, _
                                                                         SEA_THICKNESS, _
                                                                         POLAR_LINE_THICKNESS)
           Case "box_violin"
                ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, _
                                                                         SLA_BARTHICKNESS, _
                                                                         VIOLIN_BOX_WIDTH_PERC_x_10)
           Case "histogram"
              With ActiveDocument.CurrentPageItem
                .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_BARTHICKNESS, HISTOGRAM_BAR_WIDTH_PERC_x_10)
              End With
        End Select
    Next iPlot
End Sub

' ========================================
' To handle text as symbol
' ========================================
Sub SetTextAsSymbol(plotIndex As Long)
    Const DEBUG_MODE As Boolean = False
    Dim dataTable As Object
    Dim rowIndex As Long
    Dim nRows As Long
    Dim tgtCol As Long
    Dim plotStartCol As Long
    Dim numCol As Long
    Dim textCol As Long
    Dim dataIndex As Long
    Dim maxNRows As Long
    Dim readValue As Variant
    Dim plotObj As Object

    ' Open worksheet before accessing data
    ActiveDocument.NotebookItems(WORKSHEET_NAME).Open()
    Set dataTable = ActiveDocument.NotebookItems(WORKSHEET_NAME).DataTable

    ' Find source column for data values
    plotStartCol = _FindChunkStartCol(plotIndex)
    DebugMsg(DEBUG_MODE, "plotStartCol found: " & plotStartCol)

    ' Try to find symbol column - look for the z-value column
    numCol = plotStartCol + HEATMAP_SCATTER_ID_Z
    DebugMsg(DEBUG_MODE, "numCol defined: " & numCol)

    ' Convert numeric column to text-formatted column
    textCol = plotStartCol + HEATMAP_SCATTER_ID_SYMBOL
    _ConvertNumToText(numCol, textCol)

    ' Apply text symbol settings in a specific order
    DebugMsg(DEBUG_MODE, "Applying text symbol settings")
    DebugMsg(DEBUG_MODE, "textCol: " & textCol)

    ActiveDocument.NotebookItems(GRAPH_NAME).Open()

    ' Configure the plot to use text symbols
    ' Set plotObj = _SelectPlot(plotIndex)
    _SelectPlot(plotIndex)

    With ActiveDocument.CurrentPageItem
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_OPTIONS, &H00000201&)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_SIZE, 64)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_SIZEREPEAT, 4)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_SHAPE, textCol)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_SHAPEREPEAT, 4)
    End With
End Sub

' ========================================
' Main
' ========================================
Sub Main()
    Const DEBUG_MODE As Boolean = False

    ActiveDocument.NotebookItems(GRAPH_NAME).Open

    ' Remove any existing graphs
    RemoveExistingGraphs()

    ' Data Plotting
    Plot()

    ' Removers
    RemoveLegend()
    RemoveTopSpine()
    RemoveRightSpine()
    RemoveTitle()

    ' Color
    ApplyColors()

    ' Axes
    SetScales()
    SetRanges()

    ' Size
    SetFigureSize()

    ' Ticks
    SetTickPositions()
    SetTickSizes()

    ' XY Labels
    SetLabelsText()

    ' Ticks and XY Labels
    SetLabelSizes()

    ' Tick label rotation
    SetTickLabelRotation()

    ' Handle special cases
    HandleSpecialCases()

    ' Activate the graph page
    ActiveDocument.NotebookItems(GRAPH_NAME).Open

End Sub