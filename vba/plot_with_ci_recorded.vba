Option Explicit
Function FlagOn(flag As Long)
  FlagOn = flag Or FLAG_SET_BIT ' Use to set option flag bits on, leaving others unchanged
End Function
Function FlagOff(flag As Long)
  FlagOff = flag Or FLAG_CLEAR_BIT ' Use to set option flag bits off, leaving others unchanged
End Function
Sub Main
ActiveDocument.CurrentPageItem.Select(False, -2241, 1891, -2241, 1891)
ActiveDocument.NotebookItems("Data 2").Open
ActiveDocument.CurrentDataItem.Open
ActiveDocument.CurrentDataItem.Open
ActiveDocument.CurrentDataItem.Open
Dim ColumnsPerPlot()
ReDim ColumnsPerPlot(2, 1)
ColumnsPerPlot(0, 0) = 4
ColumnsPerPlot(1, 0) = 0
ColumnsPerPlot(2, 0) = 31999999
ColumnsPerPlot(0, 1) = 5
ColumnsPerPlot(1, 1) = 0
ColumnsPerPlot(2, 1) = 31999999
Dim PlotColumnCountArray()
ReDim PlotColumnCountArray(0)
PlotColumnCountArray(0) = 2
ActiveDocument.CurrentPageItem.AddWizardPlot("Area Plot", "Simple Area", "XY Pair", ColumnsPerPlot, PlotColumnCountArray, "Worksheet Columns", "Standard Deviation", "Degrees", 0.000000, 360.000000, , "Standard Deviation", True)
ActiveDocument.NotebookItems("Graph Page 1").Open
ActiveDocument.CurrentPageItem.Select(False, -1525, 1888, -1525, 1888)
ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_LINETYPE, 1)
ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_LINETYPE, 1)
ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLOR, &H00ff0000&)
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLORREPEAT, &H00000002&)
ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLOR, &H00ff0000&)
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLORREPEAT, &H00000002&)
ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, &H00ff0000&)
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLORREPEAT, &H00000002&)
ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, &H00ff0000&)
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLORREPEAT, &H00000002&)
ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, &H00ff0000&)
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLORREPEAT, &H00000002&)
ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, &H000008a7&, &H0000001e&)
ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, &H000008a7&, &H00000046&)
ActiveDocument.CurrentPageItem.Select(False, -237, 1487, -237, 1487)
ActiveDocument.CurrentPageItem.Select(False, -1739, 1847, -1739, 1847)
ActiveDocument.CurrentPageItem.Select(False, -1739, 1847, -1739, 1847)
ActiveDocument.NotebookItems("Data 2").Open
ActiveDocument.CurrentDataItem.Open
ActiveDocument.CurrentDataItem.Open
ActiveDocument.CurrentDataItem.Open
ReDim ColumnsPerPlot(2, 1)
ColumnsPerPlot(0, 0) = 4
ColumnsPerPlot(1, 0) = 0
ColumnsPerPlot(2, 0) = 31999999
ColumnsPerPlot(0, 1) = 6
ColumnsPerPlot(1, 1) = 0
ColumnsPerPlot(2, 1) = 31999999
ReDim PlotColumnCountArray(0)
PlotColumnCountArray(0) = 2
ActiveDocument.CurrentPageItem.AddWizardPlot("Area Plot", "Simple Area", "XY Pair", ColumnsPerPlot, PlotColumnCountArray, "Worksheet Columns", "Standard Deviation", "Degrees", 0.000000, 360.000000, , "Standard Deviation", True)
ActiveDocument.NotebookItems("Graph Page 1").Open
ActiveDocument.CurrentPageItem.Select(False, -1432, 1612, -1432, 1612)
ActiveDocument.CurrentPageItem.Select(False, -1453, 1604, -1453, 1604)
ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_LINETYPE, 1)
ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_LINETYPE, 1)
ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLOR, &H00ffffff&)
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLORREPEAT, &H00000002&)
ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLOR, &H00ffffff&)
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLORREPEAT, &H00000002&)
ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, &H00ffffff&)
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLORREPEAT, &H00000002&)
ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, &H00ffffff&)
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLORREPEAT, &H00000002&)
ActiveDocument.CurrentPageItem.Select(False, -541, 1633, -541, 1633)
ActiveDocument.CurrentPageItem.Select(False, -1890, 1854, -1890, 1854)
ActiveDocument.CurrentPageItem.Select(False, -1890, 1854, -1890, 1854)
ActiveDocument.NotebookItems("Data 2").Open
ActiveDocument.CurrentDataItem.Open
ActiveDocument.CurrentDataItem.Open
ActiveDocument.CurrentDataItem.Open
ReDim ColumnsPerPlot(2, 1)
ColumnsPerPlot(0, 0) = 4
ColumnsPerPlot(1, 0) = 0
ColumnsPerPlot(2, 0) = 31999999
ColumnsPerPlot(0, 1) = 7
ColumnsPerPlot(1, 1) = 0
ColumnsPerPlot(2, 1) = 31999999
ReDim PlotColumnCountArray(0)
PlotColumnCountArray(0) = 2
ActiveDocument.CurrentPageItem.AddWizardPlot("Line Plot", "Simple Straight Line", "XY Pair", ColumnsPerPlot, PlotColumnCountArray, "Worksheet Columns", "Standard Deviation", "Degrees", 0.000000, 360.000000, , "Standard Deviation", True)
ActiveDocument.NotebookItems("Graph Page 1").Open
ActiveDocument.CurrentPageItem.Select(False, -1533, 1745, -1533, 1745)
ActiveDocument.CurrentPageItem.Select(False, -362, 1849, -362, 1849)
ActiveDocument.CurrentPageItem.Select(False, -1507, 1748, -1507, 1748)
ActiveDocument.NotebookItems("04.BLU").Close(True)
ActiveDocument.CurrentPageItem.Select(False, -911, 1985, -911, 1985)
ActiveDocument.CurrentPageItem.Select(False, -1835, 2024, -1835, 2024)
End Sub