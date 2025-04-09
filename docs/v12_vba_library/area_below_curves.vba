Option Explicit
Public Const ObjectHelp = Path + "\SPW.CHM"
Public CurrentIndex, NewIndex, PlottedCurves, CurveArea, ErrorCheck 
Dim DlgReinitializationFlag As Boolean
Public HelpID As Variant
Dim WorksheetTable As Object
Dim LastColumn As Long
Dim LastRow As Long
Dim CurrentXTuple, CurrentYTuple, CurrentZTuple, Row, CurveCount, SPPlot, PlotType, SPGraph, SPPage
Dim ResultCol As Long
Dim SPTransform As Object
Dim Separator$
Function FlagOn(flag As Long)
  FlagOn = flag Or FLAG_SET_BIT ' Use to set option flag bits on, leaving others unchanged
End Function
Function FlagOff(flag As Long)
    
  FlagOff = flag Or FLAG_CLEAR_BIT ' Use to set option flag bits off, leaving others unchanged
End Function
Sub Main
Separator = ListSeparator
On Error GoTo ErrorHandling
'Authored 12/13/99 John Kuo; original code by Fred Cabasa
'Updated 03/14/02
'This macro integrates under curves using the trapezoidal rule.
'This can be used for equal or unequally spaced x values.  
'The algorithm is: sigma i from 0 To n-1, or
'{yi(xi+1 - xi) + (1/2)(yi+1 - yi)(xi+1 - xi)}

	HelpID = 60200			' Help ID number for this topic in SPW.CHM
	CurrentIndex = 0
	NewIndex = 0
	Row = 1	
	DlgReinitializationFlag  = 0
	'Determine the data range
	ErrorCheck =0
	ActiveDocument.Activate
	Set WorksheetTable = ActiveDocument.CurrentDataItem.DataTable
	LastColumn = 0
	LastRow = 0 
	WorksheetTable.GetMaxUsedSize(LastColumn,LastRow)
'Place Worksheet into Overwrite mode
	ActiveDocument.CurrentDataItem.InsertionMode = False
	ErrorCheck =1	
	CurveCount = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_PLOT).ChildObjects.Count

'Get the current curve information
	Set PlottedCurves = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_PLOT).ChildObjects
	PlottedCurves(CurrentIndex).SetAttribute(SNA_SELECTDIM, DIM_X)
		CurrentXTuple = PlottedCurves(CurrentIndex).GetAttribute(SNA_DATACOL, CurrentXTuple)
	PlottedCurves(CurrentIndex).SetAttribute(SNA_SELECTDIM, DIM_Y)
		CurrentYTuple = PlottedCurves(CurrentIndex).GetAttribute(SNA_DATACOL, CurrentYTuple)
	PlottedCurves(CurrentIndex).SetAttribute(SNA_SELECTDIM, DIM_Z)
		CurrentZTuple = PlottedCurves(CurrentIndex).GetAttribute(SNA_DATACOL, CurrentZTuple)

'Restricts use of macro to plots with X and Y data
	If CurrentXTuple < 0 Or CurrentYTuple < 0 Then
		HelpMsgBox 60200, "This macro only functions on plots with both X and Y data",vbInformation,"XY Plot required"
	End
	End If
'Determine if graph is two or three dimensional and the number of X oy Y variables
	If CurrentZTuple > 0 Then 
		HelpMsgBox 60200, "This macro does not function on 3D plots",vbInformation,"XY Plot required"	
	End
	End If
'Checks for other coordinate systems
	Set SPPage = ActiveDocument.CurrentPageItem
	Set SPGraph = SPPage.GraphPages(0).CurrentPageObject(GPT_GRAPH)
	Set SPPlot = SPGraph.Plots(0)
	PlotType = SPPlot.GetAttribute(SLA_TYPE, PlotType)
	If PlotType > 1 Then
		HelpMsgBox 60200, "This macro only functions on scatter or line plots", vbInformation, "SigmaPlot"
	End
	End If	
MacroDialog:
	Begin Dialog UserDialog 442,115,"Area Below Curves",.DialogFunc ' %GRID:10,7,1,0
		OKButton 322,90,96,18,.OK
		PushButton 10,90,96,18,"Help",.PushButton1
		Text 10,8,123,14,"&Area under curve",.Text2
		TextBox 10,22,132,18,.Area
		PushButton 142,22,132,19,"&Compute",.Compute
		PushButton 10,55,132,20,"<< &Previous Curve",.PreviousCurve
		PushButton 142,55,132,20,"&Next Curve >>",.NextCurve
		Text 285,8,120,14,"&Results column",.Text1
		TextBox 285,22,132,18,.Results
		Text 285,58,150,14,"Curve",.PlottedCols
	End Dialog
	On Error GoTo DialogError
	Dim dlg As UserDialog
	Select Case Dialog(dlg)
		Case 0 'Handles Cancel button
			GoTo Finish
'		Case 1 'Handles Help button
'			Help(ObjectHelp,HelpID	)
			GoTo MacroDialog
	End Select

'Error Handling
ErrorHandling:
If ErrorCheck=0 Then
	HelpMsgBox 60200, "You must have a worksheet open to run this macro",vbExclamation,"SigmaPlot"
	GoTo Finish
End If

If ErrorCheck=1 Then
	HelpMsgBox 60200, "This notebook contains no graphs.  Please create a new graph before running this macro.", vbExclamation,"SigmaPlot"
	GoTo Finish
End If
DialogError:
If IsNumeric(dlg.Results)=False Or dlg.Results="" Then
	MsgBox "You must enter a valid positive integer greater than last data column for your result column",vbExclamation,"Invalid Results Column"
	GoTo MacroDialog
ElseIf IsNumeric(dlg.Results)=True Then
	If CLng(dlg.Results) < 1 Or CDbl(dlg.Results) < (LastColumn + 1) Then
		MsgBox "You must enter a postive integer greater than the last data column for your result column",vbExclamation,"Invalid Results Column"
		GoTo MacroDialog
	End If
End If
Finish:
End Sub
Rem See DialogFunc help topic for more information.
Private Function DialogFunc(DlgItem$, Action%, SuppValue%) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
	If DlgReinitializationFlag = 0 Then
		DlgText "OK", "Cancel"
	Else 
		DlgText "OK", "Close"
	End If
		DlgText "PlottedCols", "Column " + CStr(CurrentXTuple + 1) + " vs. " + CStr(CurrentYTuple + 1)
		DlgText "Results", "First Empty"
		If PlottedCurves.Count > 1 Then'ShowNext = True Then
			DlgVisible "NextCurve", True
			DlgVisible "PreviousCurve", True		
			If CurrentIndex = PlottedCurves.Count - 1 Then 
				DlgEnable "PreviousCurve", True
				DlgEnable "NextCurve", False
			Else
				DlgEnable "NextCurve", True
				If CurrentIndex=0 Then				
					DlgEnable "PreviousCurve", False
				Else
					DlgEnable "PreviousCurve", True
				End If
			End If		
		Else
			DlgVisible "NextCurve", False
			DlgVisible "PreviousCurve", False
		End If
		'Parse the "First Empty" result
		DlgText "Results", "First Empty"
		ResultCol = LastColumn + 1
	Case 2 ' Value changing or button pressed
		Select Case DlgItem$
			Case "PushButton1"
				Help(ObjectHelp,HelpID)
				DialogFunc = True 'do not exit the dialog

			Case "Cancel"
				End
			Case "NextCurve"
				CurrentIndex = CurrentIndex + 1
				DlgEnable "PreviousCurve", True
				If CurrentIndex = PlottedCurves.Count - 1 Then DlgEnable "NextCurve", False
				PlottedCurves(CurrentIndex).SetAttribute(SNA_SELECTDIM, DIM_X)
					CurrentXTuple = PlottedCurves(CurrentIndex).GetAttribute(SNA_DATACOL, CurrentXTuple)
				PlottedCurves(CurrentIndex).SetAttribute(SNA_SELECTDIM, DIM_Y)
					CurrentYTuple = PlottedCurves(CurrentIndex).GetAttribute(SNA_DATACOL, CurrentYTuple)
				DialogFunc=True
				Row=Row+1
				DlgText "PlottedCols", "Column " + CStr(CurrentXTuple + 1) + " vs. " + CStr(CurrentYTuple + 1)
			Case "PreviousCurve"
				CurrentIndex = CurrentIndex - 1
				DlgEnable "NextCurve", True
				If CurrentIndex = 0 Then DlgEnable "PreviousCurve", False
				PlottedCurves(CurrentIndex).SetAttribute(SNA_SELECTDIM, DIM_X)
					CurrentXTuple = PlottedCurves(CurrentIndex).GetAttribute(SNA_DATACOL, CurrentXTuple)
				PlottedCurves(CurrentIndex).SetAttribute(SNA_SELECTDIM, DIM_Y)
					CurrentYTuple = PlottedCurves(CurrentIndex).GetAttribute(SNA_DATACOL, CurrentYTuple)
				DialogFunc=True
				Row=Row-1
				DlgText "PlottedCols", "Column " + CStr(CurrentXTuple + 1) + " vs. " + CStr(CurrentYTuple + 1)
			Case "OK"
				If DlgText("OK") = "Close" Then
					WorksheetTable.NamedRanges.Add("Curves",ResultCol,0,1,-1, True)
					WorksheetTable.NamedRanges.Add("Areas",ResultCol-1,0,1,-1, True)
				End If
				End
			Case "Compute"
				ComputeArea
				CurveArea = WorksheetTable.Cell(ResultCol - 1, Row - 1)
				WorksheetTable.Cell(ResultCol, Row - 1) = DlgText("PlottedCols")
				DlgText "Area", CurveArea
				DialogFunc=True
				DialogFunc = True
				DlgReinitializationFlag = 1
				DlgText "OK", "Close"
		End Select
	Case 3 ' TextBox or ComboBox text changed
		Select Case DlgItem$
			Case "Results"
				ResultCol = CLng(DlgText("Results"))
				If ResultCol < (LastColumn + 1) Then
				Error				
				End If
			End Select
	Case 4 ' Focus changed
	Case 5 ' Idle
	Case 6 ' Function key
	End Select
End Function
Sub ComputeArea
'Area under curve transform using trapezoidal rule
	Set SPTransform = ActiveDocument.NotebookItems.Add(9)
		SPTransform.Open
		SPTransform.AddVariableExpression("x", "col(" + CStr(CurrentXTuple + 1) + ")")
		SPTransform.AddVariableExpression("y", "col(" + CStr(CurrentYTuple + 1) + ")")
		SPTransform.AddVariableExpression("res", ResultCol)
		SPTransform.AddVariableExpression("resrow", Row)
		SPTransform.Text =	"xdif1=diff(x)" + vbCrLf + _
							"n=count(x)" + vbCrLf + _
							"xdif=xdif1[data(2" + Separator + "n)]" + vbCrLf + _
							"ydif1=diff(y)" + vbCrLf + _
							"ydif=ydif1[data(2" + Separator + "n)]" + vbCrLf + _
							"y1=y[data(1" + Separator + "n-1)]" + vbCrLf + _
							"intgrl=y1*xdif+0.5*ydif*xdif" + vbCrLf + _
							"a=total(intgrl)" + vbCrLf + _
							"cell(res" + Separator + " resrow)=a" + vbCrLf
'		SPTransform.RunEditor 'Debug the transform
		SPTransform.Execute
		SPTransform.Close(False)
End Sub
Public Function empty_col(column As Variant, column_end As Variant)
'Determines if a column is empty
	Dim WorksheetTable As Object
	Set WorksheetTable = ActiveDocument.CurrentDataItem.DataTable
	Dim i As Long
	Dim empty_cell As Boolean
	
	For i = 0 To column_end Step 3 'Change the step value to change the sampling interval.  Small sample size = Slow operation
		If WorksheetTable.Cell(column,i) = "-1.#QNAN" Or WorksheetTable.Cell(column,i) = "-1,#QNAN" Then empty_cell = True
		If WorksheetTable.Cell(column,i) <> "-1.#QNAN" And WorksheetTable.Cell(column,i) <> "-1,#QNAN" Then GoTo NotEmpty
	Next i
	empty_col = empty_cell
	GoTo EmptyCol:
	NotEmpty:	
	empty_col = False
	EmptyCol:
End Function
Public Sub HelpMsgBox(ID, Msg, Optional MsgType, Optional MsgTitle)
	MsgBox Msg, MsgType, MsgTitle
	HelpID = ID
	Begin Dialog UserDialog 340,91,"SigmaPlot Help",.udHelpBox ' %GRID:10,7,1,1
		Text 50,14,240,28,"For more information on running this macro, please click help",.Text1,2
		PushButton 200,56,90,21,"Help",.Help
		CancelButton 50,56,90,21
	End Dialog
	Dim dlg As UserDialog
	
	Select Case Dialog(dlg)  
		Case 0 'Handles Cancel button
			End
	End Select
End Sub
Rem See DialogFunc help topic for more information.
Public Function udHelpBox(DlgItem$, Action%, SuppValue&) As Boolean
	Select Case Action%
	Case 2 ' Value changing or button pressed
		Select Case DlgItem$
		Case "Help"
			Help(ObjectHelp,HelpID)
        	udHelpBox = False
        End Select
	End Select
End Function