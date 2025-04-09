Option Explicit

Sub main()

    Dim oSigmaPlot As SigmaPlot.Application
    Dim nbNew As SigmaPlot.Notebook
    Dim nbVBALib As SigmaPlot.Notebook
    Dim nbiData As NotebookItem
    Dim oDataTable As SigmaPlot.DataTable
    Dim sPathCSV As String
    Dim i As Long
    Dim lRowMax As Long
    Dim lColMax As Long
    
    sPathCSV = "C:\...\line.csv"

    Set oSigmaPlot = CreateObject("SigmaPlot.Application")
    Set nbVBALib = oSigmaPlot.Notebooks.Item("VBALib.JNB")
    
    Set nbNew = oSigmaPlot.ActiveDocument
    Set nbiData = nbNew.NotebookItems.Item("Data 1")
    Set oDataTable = nbiData.DataTable
    
    Call oDataTable.GetMaxLegalSize(lColMax, lRowMax)
    
    Call ImportCSV(nbiData, sPathCSV)
    
    Dim oGraphItem As GraphItem
    Set oGraphItem = nbNew.NotebookItems.Add(SPWNotebookComponentType.CT_GRAPHICPAGE)


    Dim ColumnsPerPlot()
    ReDim ColumnsPerPlot(2, 1)
    
    ColumnsPerPlot(0, 0) = 0
    ColumnsPerPlot(1, 0) = 0
    ColumnsPerPlot(2, 0) = lRowMax - 1
    
    ColumnsPerPlot(0, 1) = 1
    ColumnsPerPlot(1, 1) = 0
    ColumnsPerPlot(2, 1) = lRowMax - 1
    
    Dim PlotColumnCountArray()
    ReDim PlotColumnCountArray(0)
    PlotColumnCountArray(0) = 2
    Call oGraphItem.CreateWizardGraph("Line Plot", "Simple Straight Line", "XY Pair", ColumnsPerPlot, PlotColumnCountArray) ', "Worksheet Columns", "Standard Deviation", "Degrees", 0.000000, 360.000000, , "Standard Deviation", True

    ColumnsPerPlot(0, 0) = 0
    ColumnsPerPlot(1, 0) = 0
    ColumnsPerPlot(2, 0) = lRowMax - 1
    
    ColumnsPerPlot(0, 1) = 2
    ColumnsPerPlot(1, 1) = 0
    ColumnsPerPlot(2, 1) = lRowMax - 1
    Call oGraphItem.AddWizardPlot("Line Plot", "Simple Straight Line", "XY Pair", ColumnsPerPlot, PlotColumnCountArray) ', "Worksheet Columns", "Standard Deviation", "Degrees", 0#, 360#, , "Standard Deviation", True)

    Dim oGraph As Graph
    Set oGraph = oGraphItem.GraphPages(0).Graphs(0)

    Debug.Print oGraphItem.Name
    Debug.Print oGraph.Name
    
    
    Call ApplyGraphProperty(oGraph)
    
    Call oGraphItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
    Call oGraphItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, 1)
    Call oGraphItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_TICSIZE, &HEA&)


    Set oSigmaPlot = Nothing
    
End Sub

Sub processTicks(oGraphItem As GraphItem, dimension As Long)
    ' Ensure the object is correctly targeted before setting attributes
    oGraphItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
    Call oGraphItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, dimension)
    
    Call oGraphItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, 1)
    Call oGraphItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_THICKNESS, &H8)
    Call oGraphItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_TICSIZE, &H20)
    
    Call oGraphItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, 2)
    Call oGraphItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_THICKNESS, &H8)
    Call oGraphItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_TICSIZE, &H20)
    
End Sub

' Apply Figure Knowledge("KOKOROE") to graph object
Private Sub ApplyGraphProperty(oGraph As Graph)

    Dim oAxisX As SigmaPlot.Axis
    Dim oAxisY As SigmaPlot.Axis
    Dim oText As SigmaPlot.Text
    Dim oLine As SigmaPlot.Line

    Set oAxisX = oGraph.Axes(0)
    Set oAxisY = oGraph.Axes(1)
    
    'Set font name of all texts in graph
    Call oGraph.SetAttribute(STA_FONT, "Arial")
    
    'Set graph size
    oGraph.Width = ConvMmToInch(40)           '40mm
    oGraph.Height = ConvMmToInch(40 * 0.7)    '40mm
    
    'Set font size of title
    Set oText = oGraph.NameObject
    Call oText.SetAttribute(STA_SIZE, ConvPtToInch(8))
    
    'Set font size of axis titles
    Set oText = oAxisX.AxisTitles(0)
    Call oText.SetAttribute(STA_SIZE, ConvPtToInch(8))
    Set oText = oAxisY.AxisTitles(0)
    Call oText.SetAttribute(STA_SIZE, ConvPtToInch(8))

    'Set font size of axis major tick labels
    Set oText = oAxisX.TickLabelAttributes(2)
    Call oText.SetAttribute(STA_SIZE, ConvPtToInch(7))
    Set oText = oAxisY.TickLabelAttributes(2)
    Call oText.SetAttribute(STA_SIZE, ConvPtToInch(7))
    
    
    
    Set oLine = oAxisX.LineAttributes(2)
    Call oLine.SetAttribute(SEA_THICKNESS, ConvMmToInch(0.2))
    
'    Call oAxisX.SetAttribute(SAA_TICSIZE, ConvMmToInch(10))
    


End Sub

Private Function ConvMmToInch(ByVal dValMm As Double) As Double
    ConvMmToInch = dValMm / 0.0254          '(1/1000inch)
End Function
Private Function ConvPtToInch(ByVal dValPt As Double) As Double
    ConvPtToInch = dValPt * (1000 / 72)     '(1/1000inch)
End Function

Private Function ImportCSV(nbiData As NotebookItem, sPathCSV As String)

    Dim sCSV As String
    Dim arrDataCSV() As Variant
    Dim arrHeader() As Variant
    Dim arrData() As Variant
    Dim oDataTable As SigmaPlot.DataTable
    Dim i As Long
    Dim j As Long
    
    sCSV = ReadText(sPathCSV)
    arrDataCSV = CsvToArray2D(sCSV)
    
    Call SplitHeaderAndData(arrDataCSV, arrHeader, arrData)
    
    'Set Header Titles
    Set oDataTable = nbiData.DataTable
    For i = 0 To UBound(arrHeader)
        oDataTable.ColumnTitle(i) = arrHeader(i)
    Next
    
    'Set values in DataTable
    arrData = TransposeArray2D(arrData)
    Call oDataTable.PutData(arrData, 0, 0)


End Function

Private Sub SplitHeaderAndData(ByVal arrSrc As Variant, _
                               ByRef arrHeader As Variant, _
                               ByRef arrData As Variant)
                               
    Dim i As Long
    Dim j As Long
    
    ReDim arrHeader(UBound(arrSrc, 2))
    ReDim arrData(UBound(arrSrc, 1) - 1, UBound(arrSrc, 2))
     
    'Extract Header
    For i = 0 To UBound(arrSrc, 2)
        arrHeader(i) = arrSrc(0, i)
    Next
    
    'Extract Data
    For i = 1 To UBound(arrSrc, 1)
        For j = 0 To UBound(arrSrc, 2)
            arrData(i - 1, j) = CDbl(arrSrc(i, j))
        Next
    Next

End Sub

Function TransposeArray2D(ByVal arrSrc As Variant) As Variant()

    Dim arrTrans() As Variant
    Dim i As Long
    Dim j As Long
    
    ' 転置用の配列を作成
    ReDim arrTrans(UBound(arrSrc, 2), UBound(arrSrc, 1))

    ' 転置処理（行と列を入れ替える）
    For i = 0 To UBound(arrSrc, 1)
        For j = 0 To UBound(arrSrc, 2)
            arrTrans(j, i) = arrSrc(i, j)
        Next
    Next

    ' 転置した配列を返す
    TransposeArray2D = arrTrans
    
End Function


Private Function ReadText(sPathFile As String, _
                          Optional sCharset As String = "UTF-8") As String
    Dim oADO As Object
    Dim sContents As String

    Set oADO = CreateObject("ADODB.Stream")
    With oADO
        .Charset = sCharset
        .Open
        .LoadFromFile sPathFile
        sContents = .ReadText
        .Close
    End With
    Set oADO = Nothing
    
    ReadText = sContents
    
End Function
Function Split(sInput As String, sDelimiter As String) As String()

    Dim arrStr() As String
    Dim sTmp As String
    Dim lCnt As Long
    Dim i As Long
    
    For i = 1 To Len(sInput)
        If Mid(sInput, i, 1) = sDelimiter Then
            lCnt = lCnt + 1
            ReDim Preserve arrStr(lCnt)
            arrStr(lCnt - 1) = sTmp
            sTmp = ""
        Else
            sTmp = sTmp & Mid(sInput, i, 1)
        End If
    Next
    
    lCnt = lCnt + 1
    ReDim Preserve arrStr(lCnt)
    arrStr(lCnt - 1) = sTmp
    ReDim Preserve arrStr(UBound(arrStr) - 1)
    
    Split = arrStr
    
End Function
Private Function CsvToArray2D(sCSV As String) As Variant()

    Dim arrData() As Variant
    Dim arrRow() As String
    Dim arrHeader() As String
    Dim arrCol() As String
    Dim i As Long
    Dim j As Long
    
    arrRow = Split(sCSV, vbCr)
    arrHeader = Split(arrRow(0), ",")

    ReDim arrData(UBound(arrRow), UBound(arrHeader))

    For i = 0 To UBound(arrRow)
        arrCol = Split(arrRow(i), ",")
        For j = 0 To UBound(arrHeader)
            arrData(i, j) = arrCol(j)
        Next
    Next
    
    CsvToArray2D = arrData

End Function

