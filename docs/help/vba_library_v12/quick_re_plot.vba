Option Explicit
Public CurrentIndex, NewIndex, PlottedCurves, ChangeColumn, OKButton
Public NextMarker, PrevMarker, ShowNext, ShowPrevious, AllDone As Boolean
Public Const ObjectHelp = Path + "\SPW.CHM"
Public HelpID As Variant
Dim Separator$
Sub Main
AllDone = False
	On Error GoTo NoGraph
	'Authored by Frederick Cabasa 11/4/98
	'Modified on 11/9/98
	'Updated 12/8/99, 2/14/00 John Kuo
	'Updated on 9/7/01 to include handling of categorical data (Frederick Cabasa)
	'This macro lets you quickly re-pick data columns for 2D and 3D plots.
	
	HelpID = 60209			' Help ID number for this topic in SPW.CHM
	Dim SPPage, SPGraph, SPPlot, CurrentXTuple, CurrentYTuple, CurrentZTuple, CurveCount

	Set SPPage = ActiveDocument.CurrentPageItem
	SPPage.Open
	Set SPGraph = SPPage.GraphPages(0).CurrentPageObject(GPT_GRAPH)
	Set SPPlot = SPPage.GraphPages(0).CurrentPageObject(GPT_PLOT)

	'Check for category plot and if so exit
	Dim CategoryColumn As Variant
	CategoryColumn = SPPlot.GetAttribute(SLA_CATEGORYCOL, CategoryColumn)
	If CategoryColumn >= 0 Then GoTo CategoryPlot

	CurrentIndex = 0
	NewIndex = 0
	CurveCount = SPPlot.ChildObjects.Count

Repeat:
If CurveCount - 1 = CurrentIndex Then
	ShowNext = False
	ShowPrevious = True
End If
If CurrentIndex = 0 And CurveCount > 1 Then
	ShowNext = True
	ShowPrevious = False
End If
If CurveCount - 1 > CurrentIndex And CurrentIndex > 0 Then
	ShowNext = True
	ShowPrevious = True
End If
If CurveCount = 1 Then
	ShowNext = False
	ShowPrevious = False
End If

	Set PlottedCurves = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_PLOT).ChildObjects
	PlottedCurves(CurrentIndex).SetAttribute(SNA_SELECTDIM, DIM_X)
		CurrentXTuple = PlottedCurves(CurrentIndex).GetAttribute(SNA_DATACOL, CurrentXTuple)
	PlottedCurves(CurrentIndex).SetAttribute(SNA_SELECTDIM, DIM_Y)
		CurrentYTuple = PlottedCurves(CurrentIndex).GetAttribute(SNA_DATACOL, CurrentYTuple)
	PlottedCurves(CurrentIndex).SetAttribute(SNA_SELECTDIM, DIM_Z)
		CurrentZTuple = PlottedCurves(CurrentIndex).GetAttribute(SNA_DATACOL, CurrentZTuple)

'Determine the data range and define the first empty column
On Error GoTo NoData
	Dim WorksheetTable As Object
	Set WorksheetTable = ActiveDocument.CurrentDataItem.DataTable
	Dim LastColumn As Long
	Dim LastRow As Long
	LastColumn = 0
	LastRow = 0
	WorksheetTable.GetMaxUsedSize(LastColumn,LastRow)
'Place Worksheet into Overwrite mode
	ActiveDocument.CurrentDataItem.InsertionMode = False

'Sort through columns and create list of columns with values in row 1
Dim Index, UsedColumns$(), ListedColumns(), ListIndex, ColContents, ColTitle
ReDim UsedColumns$(LastColumn -1)
ReDim ListedColumns(LastColumn -1)
ListIndex = 0
For Index = 0 To LastColumn - 1
	ColContents = empty_col(Index, LastRow)
	ColTitle = WorksheetTable.Cell(Index,-1) 'Retrieve column title
	If ColContents = True Then GoTo NextIndex
	If ColContents = False Then   'If the first cell is not empty
		Select Case ColTitle
		Case "-1.#QNAN"
			UsedColumns$(Index) = "Column " + CStr(Index + 1)
			ListedColumns(ListIndex) = CStr(Index + 1)
			ListIndex = ListIndex + 1
		Case "-1,#QNAN"
			UsedColumns$(Index) = "Column " + CStr(Index + 1)
			ListedColumns(ListIndex) = CStr(Index + 1)
			ListIndex = ListIndex + 1
		Case Else
			UsedColumns$(Index) = ColTitle	'If title is present use title
			ListedColumns(ListIndex) = CStr(Index + 1)
			ListIndex = ListIndex + 1
		End Select
	End If
	NextIndex:
Next Index

WorksheetTable.GetMaxUsedSize(LastColumn,LastRow) 're-initialize these values

'Determine if graph is two or three dimensional and the number of X oy Y variables
	If CurrentZTuple < 0 Then
		If CurrentXTuple < 0 Then GoTo YOnly
		If CurrentYTuple < 0 Then GoTo XOnly
		If CurrentXTuple >= 0 And CurrentYTuple >= 0 Then GoTo TwoD
	End If
	If CurrentZTuple >= 0 Then
		If CurrentXTuple >= 0 Then GoTo ThreeD
		If CurrentXTuple < 0 Then GoTo ZOnly
	End If

TwoD:
	Begin Dialog UserDialog 402,130,"Quick Replot - 2D Graph",.RePlot2d ' %GRID:10,7,1,0
		OKButton 295,10,96,21
		CancelButton 295,40,96,21
		PushButton 295,80,96,21,"Help",.PushButton1
		Text 12,8,120,14,"&Current x column",.OldXTitle
		TextBox 12,24,123,19,.CurrentX
		Text 140,8,103,14,"New &x column",.NewXTitle
		DropListBox 140,24,123,81,UsedColumns(),.NewX
		Text 12,54,120,14,"C&urrent y column",.OldYTitle
		TextBox 12,70,123,19,.CurrentY
		Text 140,54,103,14,"New &y column",.NewYTitle
		DropListBox 140,70,123,78,UsedColumns(),.NewY
		PushButton 12,100,123,19,"<< &Previous Curve",.PreviousCurve
		PushButton 140,100,123,20,"&Next Curve >>",.NextCurve
	End Dialog

'Set Graph Defaults
		Dim dlg As UserDialog
		dlg.CurrentX = UsedColumns$(CurrentXTuple)
		dlg.CurrentY = UsedColumns$(CurrentYTuple)
		dlg.NewX = CurrentXTuple
		dlg.NewY = CurrentYTuple
		If dlg.NewY > LastColumn - 1 Then
			dlg.NewX = LastColumn - 2
			dlg.NewY = LastColumn - 1
		End If

		Select Case Dialog(dlg)
			Case 0 'Handles Cancel button
				GoTo Finish
'			Case 1 'Handles Help button
'				Dim ObjectHelp, HelpID As Variant
'				ObjectHelp = Path + "\SPW.CHM"
'				Help(ObjectHelp,HelpID)
'				GoTo TwoD
		End Select

'Flow control for state of controls
If ChangeColumn = True Then GoTo RePlotXYData
If ChangeColumn = False Then GoTo NextXY

RePlotXYData:
'Apply dialog changes
Debug.Print NewIndex
		PlottedCurves(NewIndex).SetAttribute(SNA_SELECTDIM, DIM_X)
		PlottedCurves(NewIndex).SetAttribute(SNA_DATACOL, CLng(ListedColumns(dlg.NewX)-1))
		PlottedCurves(NewIndex).SetAttribute(SNA_SELECTDIM, DIM_Y)
		PlottedCurves(NewIndex).SetAttribute(SNA_DATACOL, CLng(ListedColumns(dlg.NewY)-1))

'Reset the Legend
		Dim SPLegend
		Set SPLegend = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).AutoLegend
		SPLegend.ChildObjects(NewIndex + 1).Name = "Col " + CStr(ListedColumns(dlg.NewX)) + " vs. " +  "Col " + CStr(ListedColumns(dlg.NewY))
		ChangeColumn = False

NextXY:
'Flow control for state of controls
If AllDone = True Then GoTo Finish
If CurrentIndex > PlottedCurves.Count - 1 Then GoTo Finish
If NextMarker = True Then NewIndex = NewIndex + 1
If PrevMarker = True Then NewIndex = NewIndex - 1
If NextMarker = True Or PrevMarker = True Then GoTo Repeat

ThreeD:
	Begin Dialog UserDialog 402,180,"Quick Replot- 3D Graph",.RePlot3d ' %GRID:10,7,1,0
		OKButton 295,10,96,21
		CancelButton 295,40,96,21
		PushButton 295,70,96,21,"Help",.PushButton1
		Text 12,8,120,14,"&Current x column",.OldXTitle
		TextBox 12,24,123,19,.CurrentX
		Text 140,8,103,14,"New &x column",.NewXTitle
		DropListBox 140,24,123,67,UsedColumns(),.NewX
		Text 12,54,120,14,"C&urrent y column",.OldYTitle
		TextBox 12,70,123,19,.CurrentY
		Text 140,54,103,14,"New &y column",.NewYTitle
		DropListBox 140,70,123,62,UsedColumns(),.NewY
		Text 12,100,118,14,"Cu&rrent z column",.Text1
		TextBox 12,116,123,19,.CurrentZ
		Text 140,100,107,14,"New &z column",.Text2
		DropListBox 140,115,123,66,UsedColumns(),.NewZ
		PushButton 12,145,123,18,"<< &Previous Curve",.PreviousCurve
		PushButton 140,145,123,19,"&Next Curve >>",.NextCurve
	End Dialog

Dim Dlg2 As UserDialog
		Dlg2.CurrentX = UsedColumns$(CurrentXTuple)
		Dlg2.CurrentY = UsedColumns$(CurrentYTuple)
		Dlg2.CurrentZ = UsedColumns$(CurrentZTuple)
		Dlg2.NewX = CurrentXTuple
		Dlg2.NewY = CurrentYTuple
		Dlg2.NewZ = CurrentZTuple
		If Dlg2.NewZ > LastColumn - 1 Then
			Dlg2.NewX = LastColumn - 3
			Dlg2.NewY = LastColumn - 2
			Dlg2.NewZ = LastColumn - 1
		End If

		Select Case Dialog(Dlg2)
			Case 0 'Handles Cancel button
				GoTo Finish
'			Case 1 'Handles Help button
'				HelpID = 60209 ' Help ID number for the Quick Re-Plot topic
'				Help(ObjectHelp,HelpID)
'				GoTo ThreeD
		End Select

'Flow control for state of controls
If ChangeColumn = True Then GoTo RePlotXYZData
If ChangeColumn = False Then GoTo NextXYZ

RePlotXYZData:
'Apply dialog changes
		PlottedCurves(NewIndex).SetAttribute(SNA_SELECTDIM, DIM_X)
		PlottedCurves(NewIndex).SetAttribute(SNA_DATACOL, CLng(ListedColumns(Dlg2.NewX)-1))
		PlottedCurves(NewIndex).SetAttribute(SNA_SELECTDIM, DIM_Y)
		PlottedCurves(NewIndex).SetAttribute(SNA_DATACOL, CLng(ListedColumns(Dlg2.NewY)-1))
		PlottedCurves(NewIndex).SetAttribute(SNA_SELECTDIM, DIM_Z)
		PlottedCurves(NewIndex).SetAttribute(SNA_DATACOL, CLng(ListedColumns(Dlg2.NewZ)-1))

'Reset the Legend
		Set SPLegend = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).AutoLegend
		SPLegend.ChildObjects(NewIndex + 1).Name = "Col " + CStr(ListedColumns(Dlg2.NewX)) + " vs. " +  "Col " + CStr(ListedColumns(Dlg2.NewY)) + " vs. " + "Col " + CStr(ListedColumns(Dlg2.NewZ))
		ChangeColumn = False

NextXYZ:
'Flow control for state of controls
If AllDone = True Then GoTo Finish
If CurrentIndex > PlottedCurves.Count - 1 Then GoTo Finish
If NextMarker = True Then NewIndex = NewIndex + 1
If PrevMarker = True Then NewIndex = NewIndex - 1
If NextMarker = True Or PrevMarker = True Then GoTo Repeat

XOnly:
	Begin Dialog UserDialog 402,134,"Quick Replot - 2D Graph, X Only Data",.RePlot2d ' %GRID:10,7,1,0
		OKButton 295,10,96,21
		CancelButton 295,40,96,21
		PushButton 295,80,96,21,"Help",.PushButton1
		Text 12,8,120,14,"&Current x column",.OldXTitle
		TextBox 12,24,123,19,.CurrentX
		Text 140,8,103,14,"New &x column",.NewXTitle
		DropListBox 140,24,123,72,UsedColumns(),.NewX
		PushButton 12,100,123,19,"<< &Previous Curve",.PreviousCurve
		PushButton 140,100,123,20,"&Next Curve >>",.NextCurve
	End Dialog

'Set Graph Defaults
		Dim dlg3 As UserDialog
		dlg3.CurrentX = UsedColumns$(CurrentXTuple)
		dlg3.NewX = CurrentXTuple
		If dlg3.NewX > LastColumn - 1 Then
			dlg3.NewX = LastColumn - 1
		End If

		Select Case Dialog(dlg3)
			Case 0 'Handles Cancel button
				GoTo Finish
			Case 1 'Handles Help button
				HelpID = 60209 ' Help ID number for the Quick Re-Plot topic
				Help(ObjectHelp,HelpID)
				GoTo XOnly
		End Select

'Flow control for state of controls
If ChangeColumn = True Then GoTo RePlotXData
If ChangeColumn = False Then GoTo NextX

RePlotXData:
'Apply dialog changes
		PlottedCurves(NewIndex).SetAttribute(SNA_SELECTDIM, DIM_X)
		PlottedCurves(NewIndex).SetAttribute(SNA_DATACOL, CLng(ListedColumns(dlg3.NewX)-1))

'Reset the Legend
		Set SPLegend = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).AutoLegend
		SPLegend.ChildObjects(NewIndex + 1).Name = "Col " + CStr(ListedColumns(dlg3.NewX))
		ChangeColumn = False

NextX:
'Flow control for state of controls
If AllDone = True Then GoTo Finish
If CurrentIndex > PlottedCurves.Count - 1 Then GoTo Finish
If NextMarker = True Then NewIndex = NewIndex + 1
If PrevMarker = True Then NewIndex = NewIndex - 1
If NextMarker = True Or PrevMarker = True Then GoTo Repeat

YOnly:
	Begin Dialog UserDialog 402,134,"Quick Replot - 2D Graph, Y Only Data",.RePlot2d ' %GRID:10,7,1,0
		OKButton 295,10,96,21
		CancelButton 295,40,96,21
		PushButton 295,80,96,21,"Help",.PushButton1
		Text 12,8,120,14,"&Current y column",.OldYTitle
		TextBox 12,24,123,19,.CurrentY
		Text 140,8,103,14,"New &y column",.NewYTitle
		DropListBox 140,24,123,72,UsedColumns(),.NewY
		PushButton 12,100,123,19,"<< &Previous Curve",.PreviousCurve
		PushButton 140,100,123,20,"&Next Curve >>",.NextCurve
	End Dialog

'Set Graph Defaults
		Dim dlg4 As UserDialog
		dlg4.CurrentY = UsedColumns$(CurrentYTuple)
		dlg4.NewY = CurrentYTuple
		If dlg4.NewY > LastColumn - 1 Then
			dlg4.NewY = LastColumn - 1
		End If

		Select Case Dialog(dlg4)
			Case 0 'Handles Cancel button
				GoTo Finish
			Case 1 'Handles Help button
				HelpID = 60209 ' Help ID number for the Quick Re-Plot topic
				Help(ObjectHelp,HelpID)
				GoTo YOnly
		End Select

'Flow control for state of controls
If ChangeColumn = True Then GoTo RePlotYData
If ChangeColumn = False Then GoTo NextY

RePlotYData:
'Apply dialog changes
		PlottedCurves(NewIndex).SetAttribute(SNA_SELECTDIM, DIM_Y)
		PlottedCurves(NewIndex).SetAttribute(SNA_DATACOL, CLng(ListedColumns(dlg4.NewY)-1))

'Reset the Legend
		Set SPLegend = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).AutoLegend
		SPLegend.ChildObjects(NewIndex + 1).Name = "Col " + CStr(ListedColumns(dlg4.NewY))
		ChangeColumn = False

NextY:
'Flow control for state of controls
If AllDone = True Then GoTo Finish
If CurrentIndex > PlottedCurves.Count - 1 Then GoTo Finish
If NextMarker = True Then NewIndex = NewIndex + 1
If PrevMarker = True Then NewIndex = NewIndex - 1
If NextMarker = True Or PrevMarker = True Then GoTo Repeat

ZOnly:
	Begin Dialog UserDialog 402,134,"Quick Replot - 3D Graph, Z Only Data",.RePlot2d ' %GRID:10,7,1,0
		OKButton 295,10,96,21
		CancelButton 295,40,96,21
		PushButton 295,80,96,21,"Help",.PushButton1
		Text 12,8,120,14,"&Current z column",.OldZTitle
		TextBox 12,24,123,19,.CurrentZ
		Text 140,8,103,14,"New &z column",.NewXTitle
		DropListBox 140,24,123,72,UsedColumns(),.NewZ
		PushButton 12,100,123,19,"<< &Previous Curve",.PreviousCurve
		PushButton 140,100,123,20,"&Next Curve >>",.NextCurve
	End Dialog

'Set Graph Defaults
		Dim dlg5 As UserDialog
		dlg5.CurrentZ = UsedColumns$(CurrentZTuple)
		dlg5.NewZ = CurrentZTuple
		If dlg5.NewZ > LastColumn - 1 Then
			dlg5.NewZ = LastColumn - 1
		End If

		Select Case Dialog(dlg5)
			Case 0 'Handles Cancel button
				GoTo Finish
			Case 1 'Handles Help button
				HelpID = 60209 ' Help ID number for the Quick Re-Plot topic
				Help(ObjectHelp,HelpID)
				GoTo XOnly
		End Select

'Flow control for state of controls
If ChangeColumn = True Then GoTo RePlotZData
If ChangeColumn = False Then GoTo NextZ

RePlotZData:
'Apply dialog changes
		PlottedCurves(NewIndex).SetAttribute(SNA_SELECTDIM, DIM_Z)
		PlottedCurves(NewIndex).SetAttribute(SNA_DATACOL, ListedColumns(CLng(dlg5.NewZ)-1))

'Reset the Legend
		Set SPLegend = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).AutoLegend
		SPLegend.ChildObjects(NewIndex + 1).Name = "Col " + CStr(ListedColumns(dlg5.NewZ))
		ChangeColumn = False

NextZ:
'Flow control for state of controls
If AllDone = True Then GoTo Finish
If CurrentIndex > PlottedCurves.Count - 1 Then GoTo Finish
If NextMarker = True Then NewIndex = NewIndex + 1
If PrevMarker = True Then NewIndex = NewIndex - 1
If NextMarker = True Or PrevMarker = True Then GoTo Repeat

NoGraph:
		HelpMsgBox 60209, "This notebook contains no graphs.  Please create a new graph before running this macro.", vbExclamation,"SigmaPlot"
		GoTo EndMacro

NoData:
	HelpMsgBox 60209, "To run this macro you must have worksheet  open.",vbExclamation,"SigmaPlot"
	GoTo EndMacro

CategoryPlot:
	HelpMsgBox 60209, "Quick Replot does not support category plots.",vbExclamation,"SigmaPlot"
	GoTo EndMacro

Finish:
	'Check for axis types
    Dim CurX, CurY, CurCount As Long
    Dim PlotCurve As Object
    Set PlotCurve = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_PLOT).ChildObjects
    CurCount = PlotCurve.Count
    SPPlot.SetAttribute SLA_SELECTDIM, DIM_X
    CurX = PlotCurve(CurCount - 1).GetAttribute(SNA_DATACOL, CurX)
    SPPlot.SetAttribute SLA_SELECTDIM, DIM_Y
    CurY = PlotCurve(CurCount - 1).GetAttribute(SNA_DATACOL, CurY)

    'Change axis for categorical data
    Dim SPGraphAxisX, SPGraphAxisY As Variant
    Set SPGraphAxisX = SPGraph.Axes(0)
    Set SPGraphAxisY = SPGraph.Axes(1)

    If Categorical(CurX) <> True Then
    	SPGraphAxisX.SetAttribute SAA_TYPE, SAA_TYPE_LINEAR
    Else
        If DateAxis(CurX) = True Then
            SPGraphAxisX.SetAttribute SAA_TYPE, SAA_TYPE_DATETIME
        ElseIf TimeAxis(CurX) = True Then
            SPGraphAxisX.SetAttribute SAA_TYPE, SAA_TYPE_DATETIME
            SPGraphAxisX.SetAttribute SAA_TICLABELDATEFORM, " "
            SPGraphAxisX.SetAttribute SAA_TICLABELTIMEFORM, "HH:mm:ss"
        Else
            SPGraphAxisX.SetAttribute SAA_TYPE, SAA_TYPE_CATEGORY
        End If
    End If

    If Categorical(CurY) <> True Then
    	SPGraphAxisY.SetAttribute SAA_TYPE, SAA_TYPE_LINEAR
    Else
        If DateAxis(CurY) = True Then
            SPGraphAxisY.SetAttribute SAA_TYPE, SAA_TYPE_DATETIME
        ElseIf TimeAxis(CurY) = True Then
            SPGraphAxisY.SetAttribute SAA_TYPE, SAA_TYPE_DATETIME
            SPGraphAxisY.SetAttribute SAA_TICLABELDATEFORM, " "
            SPGraphAxisY.SetAttribute SAA_TICLABELTIMEFORM, "HH:mm:ss"
        Else
            SPGraphAxisY.SetAttribute SAA_TYPE, SAA_TYPE_CATEGORY
        End If
    End If

EndMacro:
End Sub

Rem See DialogFunc help topic for more information.
Private Function RePlot3d(DlgItem$, Action%, SuppValue%) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
		If ShowNext = True Then
			DlgVisible ("NextCurve", True)
		Else
			DlgVisible ("NextCurve", False)
		End If
		If ShowPrevious = True Then
			DlgVisible ("PreviousCurve", True)
		Else
			DlgVisible ("PreviousCurve", False)
		End If
	Case 2 ' Value changing or button pressed
			Select Case DlgItem$
				Case "PushButton1"
					Help(ObjectHelp,HelpID)
					RePlot3d = True 'do not exit the dialog
			Case "Cancel"
				End
			Case "NextCurve"
				CurrentIndex = CurrentIndex + 1
				NextMarker = True
			Case "PreviousCurve"
				CurrentIndex = CurrentIndex - 1
				NextMarker = True
			Case "OK"
				AllDone = True
			Case "NewX"
				ChangeColumn = True
			Case "NewY"
				ChangeColumn = True
			Case "NewZ"
				ChangeColumn = True
			End Select
		Rem RePlot3d = True ' Prevent button press from closing the dialog box
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem RePlot3d = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function
Rem See DialogFunc help topic for more information.
Private Function RePlot2d(DlgItem$, Action%, SuppValue%) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
		If ShowNext = True Then
			DlgVisible ("NextCurve", True)
		Else
			DlgVisible ("NextCurve", False)
		End If
		If ShowPrevious = True Then
			DlgVisible ("PreviousCurve", True)
		Else
			DlgVisible ("PreviousCurve", False)
		End If
	Case 2 ' Value changing or button pressed
		Select Case DlgItem$
			Case "PushButton1"
				Help(ObjectHelp,HelpID)
				RePlot2d = True 'do not exit the dialog
			Case "Cancel"
				End
			Case "NextCurve"
				CurrentIndex = CurrentIndex + 1
				NextMarker = True
				PrevMarker = False
			Case "PreviousCurve"
				CurrentIndex = CurrentIndex - 1
				PrevMarker = True
				NextMarker = False
			Case "OK"
				AllDone = True
			Case "NewX"
				ChangeColumn = True
			Case "NewY"
				ChangeColumn = True
			Case "NewZ"
				ChangeColumn = True
			End Select
		Rem RePlot3d = True ' Prevent button press from closing the dialog box
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem RePlot3d = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function
Public Function empty_col(Column As Variant, column_end As Variant)
'Determines if a column is empty
	Dim WorksheetTable As Object
	Set WorksheetTable = ActiveDocument.CurrentDataItem.DataTable
	Dim i As Long
	Dim empty_cell As Boolean

	For i = 0 To column_end Step 3 'Change the step value to change the sampling interval.  Small sample size = Slow operation
		If WorksheetTable.Cell(Column,i) = "-1.#QNAN" Or WorksheetTable.Cell(Column,i) = "-1,#QNAN" Then empty_cell = True
		If WorksheetTable.Cell(Column,i) <> "-1.#QNAN" And WorksheetTable.Cell(Column,i) <> "-1,#QNAN" Then GoTo NotEmpty
	Next i
	empty_col = empty_cell
	GoTo EmptyCol:
	NotEmpty:
	empty_col = False
	EmptyCol:
End Function
Public Function Categorical(ByVal Column As Long) As Boolean
   Dim objSPWorksheet As Object
   Dim LastCol As Long
   Dim LastRow As Long
   Dim TestCol() As Variant
   Dim TestData As Variant
   Dim Index As Long

   Set objSPWorksheet = ActiveDocument.CurrentDataItem.DataTable

   objSPWorksheet.GetMaxUsedSize LastCol, LastRow

   TestCol() = objSPWorksheet.GetData(Column, 0, Column, LastRow - 1)

   For Index = 0 To LastRow - 1
      TestData = TestCol(0, Index)
      If VarType(TestData) = vbString Then
          Categorical = True
          Exit Function
      End If
   Next Index

   Categorical = False

End Function
Public Function DateAxis(ByVal Column As Long) As Boolean
   Dim objSPWorksheet As Object
   Dim LastCol As Long
   Dim LastRow As Long
   Dim TestCol() As Variant
   Dim TestData As Variant
   Dim Index As Long

   Set objSPWorksheet = ActiveDocument.CurrentDataItem.DataTable

   objSPWorksheet.GetMaxUsedSize LastCol, LastRow

   TestCol() = objSPWorksheet.GetData(Column, 0, Column, LastRow - 1)

   For Index = 0 To LastRow - 1
      TestData = TestCol(0, Index)

      If CharCount("/", TestData) = 2 Then
          DateAxis = True
          Exit Function
      End If
   Next Index

   DateAxis = False

End Function
Public Function TimeAxis(ByVal Column As Long) As Boolean
   Dim objSPWorksheet As Object
   Dim LastCol As Long
   Dim LastRow As Long
   Dim TestCol() As Variant
   Dim TestData As Variant
   Dim Index As Long

   Set objSPWorksheet = ActiveDocument.CurrentDataItem.DataTable

   objSPWorksheet.GetMaxUsedSize LastCol, LastRow

   TestCol() = objSPWorksheet.GetData(Column, 0, Column, LastRow - 1)

   For Index = 0 To LastRow - 1
      TestData = TestCol(0, Index)

      If CharCount(":", TestData) = 2 Then
          TimeAxis = True
          Exit Function
      End If
   Next Index

   TimeAxis = False

End Function
Public Function CharCount(ByVal Char As String, ByVal Expression As String) As Integer
    Dim iCount As Integer
    Dim iPlace As Integer

    iPlace = 1

    Do Until InStr(iPlace, Expression, Char) = 0
        iPlace = InStr(iPlace, Expression, Char) + 1
        iCount = iCount + 1
    Loop

    CharCount = iCount
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
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		Select Case DlgItem$
		Case "Help"
			Help(ObjectHelp,HelpID)
        	udHelpBox = False
        End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem MessageBox = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function