Option Explicit

Function FlagOn(flag As Long)
  FlagOn = flag Or FLAG_SET_BIT ' Use to set option flag bits on, leaving others unchanged
End Function
Function FlagOff(flag As Long)
  FlagOff = flag Or FLAG_CLEAR_BIT ' Use to set option flag bits off, leaving others unchanged
End Function


Sub CreateLineplotWithConfidenceInterval(startColumn As Long)
    Dim columnsPerPlot(2, 1) As Long
    Dim plotColumnCountArray(0) As Long
    
    plotColumnCountArray(0) = 2

    ' Upper bound area plot
    columnsPerPlot(0, 0) = startColumn: columnsPerPlot(0, 1) = startColumn + 1
    ActiveDocument.CurrentPageItem.AddWizardPlot "Area Plot", "Simple Area", "XY Pair", columnsPerPlot, plotColumnCountArray, "Worksheet Columns", "Standard Deviation", "Degrees", 0, 360, , "Standard Deviation", True
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute GPM_SETPLOTATTR, SEA_COLOR, &H00ff0000&

    ' Lower bound area plot
    columnsPerPlot(0, 1) = startColumn + 2
    ActiveDocument.CurrentPageItem.AddWizardPlot "Area Plot", "Simple Area", "XY Pair", columnsPerPlot, plotColumnCountArray, "Worksheet Columns", "Standard Deviation", "Degrees", 0, 360, , "Standard Deviation", True
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute GPM_SETPLOTATTR, SEA_COLOR, &H00ffffff&

    ' Mean line plot
    columnsPerPlot(0, 1) = startColumn + 3
    ActiveDocument.CurrentPageItem.AddWizardPlot "Line Plot", "Simple Straight Line", "XY Pair", columnsPerPlot, plotColumnCountArray, "Worksheet Columns", "Standard Deviation", "Degrees", 0, 360, , "Standard Deviation", True
End Sub

Sub Main()
    Dim startColumnInput As String
    startColumnInput = InputBox("Enter starting column number:", "Input", "1")
    If IsNumeric(startColumnInput) Then
        CreateLineplotWithConfidenceInterval CLng(startColumnInput)
    Else
        MsgBox "Invalid input. Please enter a valid number.", vbExclamation
    End If
End Sub


' could you understand what they do? I do draw a line plot with confidence interval with transparent colors. data columns are alined as x, y_upper, y_under, y_mean. how can I write a macro in sigmaplot?



