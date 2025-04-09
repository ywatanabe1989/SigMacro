Option Explicit
Public Const ObjectHelp = Path + "\SPW.CHM"
Public HelpID As Variant
Dim sD As Variant
Dim QNAN As String
Dim QNB As String	
Dim QNAN1 As String

Sub Main
'Authored by Frederick Cabasa 11/4/98; Modified 12/7/99 John Kuo; Modified 5/26/04 Dick Mitchell
'Modified on 9/07/01 to include multiple use of macro on same graph  (Frederick Cabasa)
'This macro takes a plot and adds a column as labels for the plot's datapoints,
'offsetting the label.  Note that the label is always centered on the offset so
'that long labels must be positioned above or below the datapoints.

	sD = DecimalSymbol  'international decimal symbol
	GetEmptyValues	
	Dim ErrorCheck
	ErrorCheck = 0
	On Error GoTo FocusMsg
	Dim SPPage, CurrentGraph, PlottedCurves, Tuple1, TupleXCol, TupleYCol, TupleRows, SPPlot, PlotType

	Set SPPage = ActiveDocument.CurrentPageItem
	Set CurrentGraph = SPPage.GraphPages(0).CurrentPageObject(GPT_GRAPH)
	Set PlottedCurves = SPPage.GraphPages(0).CurrentPageObject(GPT_PLOT)
	Set Tuple1 = PlottedCurves.ChildObjects(0)
	Dim Xcol, YCol
	ErrorCheck = 1
	Tuple1.SetAttribute(SNA_SELECTDIM,DIM_X)
	TupleXCol = Tuple1.GetAttribute(SNA_DATACOL,TupleXCol)
	Tuple1.SetAttribute(SNA_SELECTDIM,DIM_Y)
	TupleYCol = Tuple1.GetAttribute(SNA_DATACOL,TupleYCol)
	TupleRows = Tuple1.GetAttribute(SNA_LASTROW, TupleRows)

	'Detect stacked bars (1=scatter&line, 2=bar, 3=stacked bar) and grouped bars (>1 tuple)
	Dim BarType As Variant
	BarType = PlottedCurves.GetAttribute(SLA_TYPE,BarType)
	Dim NumTuples As Variant
	NumTuples = PlottedCurves.ChildObjects.Count
	If BarType = 3 Or NumTuples > 1 Then GoTo GroupedOrStackedChart

	'Get axis extents and compute axis ranges
	Dim XAxis As Variant, YAxis As Variant
	Set XAxis=CurrentGraph.Axes(0)
	Set YAxis=CurrentGraph.Axes(1)
	Dim XFromVal As Variant, XToVal As Variant, YFromVal As Variant, YToVal As Variant
	XAxis.GetAttribute(SAA_FROMVAL, XFromVal)
	XAxis.GetAttribute(SAA_TOVAL, XToVal)
	YAxis.GetAttribute(SAA_FROMVAL, YFromVal)
	YAxis.GetAttribute(SAA_TOVAL, YToVal)
	Dim XAxisRange As Double, YAxisRange As Double
	XAxisRange = Abs(XToVal - XFromVal)
	YAxisRange = Abs(YToVal - YFromVal)

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

'Create list of columns through last column
	Dim Index, UsedColumns$(), ListedColumns(), ListIndex, ColContents, ColTitle
	ListIndex = 0
	For Index = 0 To LastColumn - 1
		ColContents = empty_col(Index, LastRow)
		If ColContents = True Then GoTo NextIndex1
		ReDim Preserve UsedColumns$(ListIndex)
		ReDim Preserve ListedColumns(ListIndex)
		ColTitle = WorksheetTable.Cell(Index,-1) 'Retrieve column title
		If ColContents = False Then   'If the first cell is not empty
			Select Case ColTitle
			Case "-1.#QNAN"
				UsedColumns$(ListIndex) = "Column " + CStr(Index + 1)
				ListedColumns(ListIndex) = CStr(Index + 1)
			Case "-1,#QNAN"
				UsedColumns$(ListIndex) = "Column " + CStr(Index + 1)
				ListedColumns(ListIndex) = CStr(Index + 1)
			Case Else
				UsedColumns$(ListIndex) = ColTitle 'If title is present use title
				ListedColumns(ListIndex) = CStr(Index + 1)
			End Select
			ListIndex = ListIndex + 1
		End If
		NextIndex1:
	Next Index

	Dim Placement$(7)
	Placement$(0) = "Above"
	Placement$(1) = "Below"
	Placement$(2) = "Left"			'Omitted until true label alignment is possible
	Placement$(3) = "Right"
	Placement$(4) = "Upper-Left"
	Placement$(5) = "Upper-Right"
	Placement$(6) = "Lower-Left"
	Placement$(7) = "Lower-Right"

'Get worksheet size
	WorksheetTable.GetMaxUsedSize(LastColumn,LastRow)

'Verify that graph is a scatterplot or bar chart
	Set SPPlot = CurrentGraph.Plots(0)
	PlotType = SPPlot.GetAttribute(SLA_TYPE, PlotType)
	Select Case PlotType
		Case 1
			GoTo MacroDialog
		Case 2
			GoTo MacroDialog
'		Case 13
'			GoTo MacroDialog
		Case Else
			HelpMsgBox 60206, "This macro only operates on 2D Cartesian Graphs",vbExclamation,"SigmaPlot"
			GoTo Finish
	End Select

MacroDialog:
	Begin Dialog UserDialog 360,203,"Label Symbols" ' %GRID:10,7,1,1
		Text 10,14,120,14,"&Label column",.Text1
		DropListBox 120,10,120,70,UsedColumns(),.Column
		Text 10,44,90,14,"&Place labels",.Text2
		DropListBox 120,42,120,70,Placement(),.Placement
		Text 10,74,110,14,"&Offset labels by",.Text3
		TextBox 120,71,72,21,.Offset
		Text 199,74,50,14,"percent",.Text4
		OKButton 256,10,96,21
		CancelButton 256,40,96,21
		PushButton 256,72,96,21,"Help",.PushButton1
		GroupBox 20,98,320,98,"",.GroupBox1
		Text 40,112,290,28,"1. A plot with a single curve and a worksheet",.Text8
		Text 53,125,260,14,"column with text labels is required.",.Text7
		Text 40,143,320,14,"2. The graph must be open and in focus.",.Text5
		Text 40,165,290,14,"3. Select the plot in the graph to label",.Text6
		Text 55,178,260,14,"with text. Then run this macro.",.Text9
	End Dialog
	Dim dlg As UserDialog

'	If dlg.Offset = "" Then dlg.Offset = "3.0"
	If dlg.Offset = "" Then dlg.Offset = "3" +sD+ "0"

'Find index of last text column (0-based) skipping empty columns
	Dim EmptyIndex As Long  '1 based
	EmptyIndex = 0
	Dim IndexSave As Long '0 based
	IndexSave = 0
	Index = -1
	Dim I As Integer
	For I = 0 To LastColumn - 1
		ColContents = empty_col(I, LastRow)
		If ColContents = True Then
			EmptyIndex = EmptyIndex + 1
			GoTo ContinueFor
		End If
		Index = Index + 1
		If VarType(WorksheetTable.Cell(I,0)) = 8 Then  '8 is string
			dlg.Column = Index
			IndexSave = Index
		End If
		ContinueFor:
	Next I
	dlg.Placement = 0

	Select Case Dialog(dlg)
		Case 0 'Handles Cancel button
			GoTo Finish
		Case 1 'Handles Help button
			HelpID = 60206			' Help ID number for this topic in SPW.CHM
			Help(ObjectHelp,HelpID)
			GoTo MacroDialog
	End Select

	Dim WorksheetLabelColumn As Long
	WorksheetLabelColumn = IndexSave + EmptyIndex + 1

	If IsNumeric(CVar(dlg.Offset)) = False Then GoTo NumberNeeded
	
	'Get length of data column(s)
	Dim XColLength As Long, YColLength As Long, TupleLength As Long
	XColLength = 1
	YColLength = 1
	If TupleXCol >= 0 Then XColLength = ColumnLength(TupleXCol+1, LastRow, WorksheetTable)
	If TupleYCol >= 0 Then YColLength = ColumnLength(TupleYCol+1, LastRow, WorksheetTable)
	If XColLength > YColLength Then
		TupleLength = XColLength
	Else
		TupleLength = YColLength
	End If

'	Dim mLastCol, mLastRow, mColumn, mRow
'	mLastCol = LastColumn - 1
'	mLastRow = LastRow - 1
'	If LastRow - 1 <> TupleRows Then
'	MsgBox "Number of labels does not equal number of symbols.  Please reformat data worksheet"
'		GoTo Finish
'	End If

'Copy and paste data and symbol columns
	Dim SPDataX()As Variant
	Dim SPDataY()As Variant
	Dim SPDataS()As Variant
	Dim Only()As Variant
	Dim Index1
'	ReDim Only(LastRow - 1)
	ReDim Only(TupleLength - 1)
'	For Index1 = 0 To LastRow - 1
	For Index1 = 0 To TupleLength - 1
		Only(Index1) = CInt(Index1 + 1)
	Next Index1
'	SPDataX() = WorksheetTable.GetData(TupleXCol,0,TupleXCol,mLastRow)
'	SPDataY() = WorksheetTable.GetData(TupleYCol,0,TupleYCol,mLastRow)
	SPDataX() = WorksheetTable.GetData(TupleXCol,0,TupleXCol,TupleLength-1)
	SPDataY() = WorksheetTable.GetData(TupleYCol,0,TupleYCol,TupleLength-1)

	'Check for too-few text labels
	Dim TooFew_TextLabels As Boolean
	TooFew_TextLabels = False	
	If TupleXCol >= 0 Then 'Single - X or XY - Pair
'		TooFew_TextLabels = TooFewTextLabels (SPDataX, dlg.Column, WorksheetTable)	
		TooFew_TextLabels = TooFewTextLabels (SPDataX, WorksheetLabelColumn, WorksheetTable)	
	Else  'Single - Y 
		TooFew_TextLabels = TooFewTextLabels (SPDataY, dlg.Column, WorksheetTable)	
	End If
	If TooFew_TextLabels = True Then
		GoTo TextLabelCheck
	End If
	ReturnFromTextLabelCheck:

'Re-dim worksheet
	WorksheetTable.GetMaxUsedSize(LastColumn,LastRow)

'X only or Y only data
	If TupleYCol < 0 Then		'X Only
		WorksheetTable.PutData(SPDataX,LastColumn,0)
		WorksheetTable.PutData(Only, LastColumn + 1,0)
	ElseIf TupleXCol < 0 Then	'Y Only
		WorksheetTable.PutData(Only,LastColumn,0)
		WorksheetTable.PutData(SPDataY,LastColumn + 1,0)
	Else
		WorksheetTable.PutData(SPDataX,LastColumn,0)
		WorksheetTable.PutData(SPDataY, LastColumn + 1,0)
	End If

'Add Column Titles to Data and Results
	Dim X As String
	Dim Y As String
	Dim Z As String
	X = "Symbol Data X"
	Y = "Symbol Data Y"
	Z = "Symbol Data"
	WorksheetTable.NamedRanges.Add(X,LastColumn,0,1,-1, True)
	WorksheetTable.NamedRanges.Add(Y,LastColumn+1,0,1,-1, True)
	WorksheetTable.NamedRanges.Add(Z,ListedColumns(CLng(dlg.Column))-1,0,1,-1, True)

'Choose columns for plotting
	Dim PlottedColumns(1) As Variant
	PlottedColumns(0) = LastColumn		'Plot Column X Vs
	PlottedColumns(1) = LastColumn + 1 		'Column Y

'Compute X and Y offsets from user-entered percent of axis range

'Modify symbols based on placement and offset measurement
	mSymbolCol dlg.Placement, PlottedColumns(0), PlottedColumns(1), CDbl(dlg.Offset), XAxisRange, YAxisRange, LastRow, WorksheetTable

'Plot second graph
	SPPage.AddWizardPlot("Scatter Plot", "Simple Scatter", "XY Pair", PlottedColumns)

'Enhance new graph
	Dim SPSymbols, SPLegend
	Set SPPlot = CurrentGraph.Plots(CurrentGraph.Plots.Count-1)
	Set SPSymbols = SPPlot.Symbols
'	Set SPLegend = CurrentGraph.AutoLegend

'Remove the Legend
	SPPage.GraphPages(0).Graphs(0).SetAttribute(SGA_FLAGS, SGA_FLAG_AUTOLEGENDBOX)

	SPPlot.Symbols.SetAttribute(SSA_SHAPEREPEAT, SOA_REPEAT_COLUMN)
	SPPlot.Symbols.SetAttribute(SSA_SHAPECOL, ListedColumns(CLng(dlg.Column))-1)

GoTo Finish

FocusMsg:
	HelpMsgBox 60206, "You must have a graph page and its worksheet open" +vbCrLf+ _
	"and the plot to be labeled selected.",vbExclamation,"SigmaPlot"
	GoTo Finish

GroupedOrStackedChart:
	MsgBox "This macro does not work with plots with multiple curves such as" + vbCrLf+ _
	"multiple scatter or grouped or stacked bar charts.  Select a plot" + vbCrLf+ _
	"with a simple scatter, line, vertical or horizontal bar chart.", vbExclamation,"SigmaPlot
	GoTo Finish

'NoGraph:
'	HelpMsgBox 60206, "This notebook contains no graphs.  Please create a new graph before running this macro.", vbExclamation,"SigmaPlot"
'	GoTo Finish

NoData:
	HelpMsgBox 60206, "To run this macro you must have a worksheet open.", vbExclamation,"SigmaPlot"
	GoTo Finish
NumberNeeded:
	MsgBox "The offset must be a percent of the axis range.", vbExclamation,"SigmaPlot"
	GoTo MacroDialog
	GoTo Finish
	
TextLabelCheck:
	If MsgBox ("There are very few text symbols in your symbol column.  Do you want to continue?", vbYesNo,"SigmaPlot") = vbYes Then
		GoTo ReturnFromTextLabelCheck
	End If

Finish:

End Sub
Public Sub mSymbolCol(Alignment As Integer, SymbolX As Variant, SymbolY As Variant, PercentRange As Double, _
XAxisRange As Double, YAxisRange As Double, LastRow As Long, WorksheetTable As Object)
	Dim Counter
	Counter = 0

	Dim XSpacing As Double, YSpacing As Double
	XSpacing = (PercentRange/100)*XAxisRange
	YSpacing = (PercentRange/100)*YAxisRange

	Select Case Alignment

		Case 0 'Above
			Do While Counter < LastRow
				If IsReallyNumeric(WorksheetTable.Cell(SymbolY, Counter)) = True Then
					WorksheetTable.Cell(SymbolY, Counter) = WorksheetTable.Cell(SymbolY, Counter) + YSpacing
				End If
				Counter = Counter + 1
			Loop

		Case 1 'Below
			Do While Counter < LastRow
				If IsReallyNumeric(WorksheetTable.Cell(SymbolY, Counter)) = True Then
					WorksheetTable.Cell(SymbolY, Counter) = WorksheetTable.Cell(SymbolY, Counter) - YSpacing
				End If
				Counter = Counter + 1
			Loop

		Case 2 'Left
			Do While Counter < LastRow
				If IsReallyNumeric(WorksheetTable.Cell(SymbolX, Counter)) = True Then
					WorksheetTable.Cell(SymbolX, Counter) = WorksheetTable.Cell(SymbolX, Counter) - XSpacing
				End If
				Counter = Counter + 1
			Loop

		Case 3 'Right
			Do While Counter < LastRow
				If IsReallyNumeric(WorksheetTable.Cell(SymbolX, Counter)) = True Then
					WorksheetTable.Cell(SymbolX, Counter) = WorksheetTable.Cell(SymbolX, Counter) + XSpacing
				End If
				Counter = Counter + 1
			Loop

		Case 4 'Upper-Left
			Do While Counter < LastRow
				If IsReallyNumeric(WorksheetTable.Cell(SymbolY, Counter)) = True Then
					WorksheetTable.Cell(SymbolY, Counter) = WorksheetTable.Cell(SymbolY, Counter) + YSpacing
				End If
				If IsReallyNumeric(WorksheetTable.Cell(SymbolX, Counter)) = True Then
					WorksheetTable.Cell(SymbolX, Counter) = WorksheetTable.Cell(SymbolX, Counter) - XSpacing
				End If
				Counter = Counter + 1
			Loop

		Case 5 'Upper-Right
			Do While Counter < LastRow
				If IsReallyNumeric(WorksheetTable.Cell(SymbolY, Counter)) = True Then
					WorksheetTable.Cell(SymbolY, Counter) = WorksheetTable.Cell(SymbolY, Counter) + YSpacing
				End If
				If IsReallyNumeric(WorksheetTable.Cell(SymbolX, Counter)) = True Then
					WorksheetTable.Cell(SymbolX, Counter) = WorksheetTable.Cell(SymbolX, Counter) + XSpacing
				End If
				Counter = Counter + 1
			Loop

		Case 6 'Lower-Left
			Do While Counter < LastRow
				If IsReallyNumeric(WorksheetTable.Cell(SymbolY, Counter)) = True Then
					WorksheetTable.Cell(SymbolY, Counter) = WorksheetTable.Cell(SymbolY, Counter) - YSpacing
				End If
				If IsReallyNumeric(WorksheetTable.Cell(SymbolX, Counter)) = True Then
					WorksheetTable.Cell(SymbolX, Counter) = WorksheetTable.Cell(SymbolX, Counter) - XSpacing
				End If
				Counter = Counter + 1
			Loop

		Case 7 'Lower-Right
			Do While Counter < LastRow
				If IsReallyNumeric(WorksheetTable.Cell(SymbolY, Counter)) = True Then
					WorksheetTable.Cell(SymbolY, Counter) = WorksheetTable.Cell(SymbolY, Counter) - YSpacing
				End If
				If IsReallyNumeric(WorksheetTable.Cell(SymbolX, Counter)) = True Then
					WorksheetTable.Cell(SymbolX, Counter) = WorksheetTable.Cell(SymbolX, Counter) + XSpacing
				End If
				Counter = Counter + 1
			Loop
	End Select
End Sub
Public Function empty_col(column As Variant, column_end As Variant)
'Determines if a column is empty
	Dim WorksheetTable As Object
	Set WorksheetTable = ActiveDocument.CurrentDataItem.DataTable
	Dim I As Long
	Dim empty_cell As Boolean

	For I = 0 To column_end Step 3 'Change the step value to change the sampling interval.  Small sample size = Slow operation
		If WorksheetTable.Cell(column,I) = "-1.#QNAN" Or WorksheetTable.Cell(column,I) = "-1,#QNAN" Then empty_cell = True
		If WorksheetTable.Cell(column,I) <> "-1.#QNAN" And WorksheetTable.Cell(column,I) <> "-1,#QNAN" Then GoTo NotEmpty
	Next I
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
Function IsReallyNumeric(ByRef value As Variant) As Boolean
'Determines if worksheet cell is numeric (Isnumeric considers +inf and blank to be numeric)

    IsReallyNumeric = True
    ' weed out obvious garbage
    If IsNumeric(value) Then
        Dim temp
        Dim length As Long
        Dim I As Long
        length = Len(value)
        For I = 1 To length
            temp = Mid$(value, I, 1)
            If (temp = "-" Or temp = "+") And I = length Then
                IsReallyNumeric = False
                Exit For
            ElseIf Not IsNumeric(temp) Then
                If temp <> "E" And temp <> "e" And temp <> sD And _
                   temp <> "+" And temp <> "-" Then
                    IsReallyNumeric = False
                    Exit For
                End If
            End If
        Next I
        If IsReallyNumeric = False Then
            Exit Function
        End If
    Else
        IsReallyNumeric = False
    End If
    If Left$(value, 1) = "+" Then
        value = Right$(value, length - 1)
    End If
End Function
Public Function TooFewTextLabels(SPDataX As Variant, WorksheetLabelColumn As Variant, WorksheetTable As Object) As Boolean
	'Find the ratio of text values in symbol column to numeric values in data (X) column
	'If less than 0.5 then return True

	TooFewTextLabels = False
	Dim CountNumber As Long
	Dim CountText As Long
	CountNumber = 0
	CountText = 0
	Dim I As Long
	For I = 0 To UBound(SPDataX,2)
		If IsReallyNumeric(SPDataX(0,I)) = True Then CountNumber = CountNumber + 1
'		If IsReallyNumeric(WorksheetTable.Cell(SymbolColumn,I)) = False Then CountText = CountText + 1  'assume string if not number - close enough
'		If VarType(WorksheetTable.Cell(SymbolColumn,I)) = 8 Then CountText = CountText + 1  'assume string if not number - close enough
		If VarType(WorksheetTable.Cell(WorksheetLabelColumn-1,I)) = 8 Then CountText = CountText + 1  'assume string if not number - close enough
	Next I
	If CountNumber <> 0 Then
		If CLng(CountText)/CLng(CountNumber) < 0.5 Then TooFewTextLabels = True
	End If
	If CountNumber = 0 And CountText < 1 Then TooFewTextLabels = True
End Function
Function ColumnLength(ByVal SelectedColumn As Long, ByVal LastRow As Long, _
ByVal WorkSheetTableObject As Object) As Long
'Finds the length of a column in a worksheet, 1-based
'Starts at LastRow cell and looks backward for a non-blank or non-"--" cell

	Dim I As Long
	For I = 0 To LastRow-1
		If BlankCell(WorkSheetTableObject.Cell(SelectedColumn-1,I)) = False Then
			ColumnLength = I+1
		Else
			Exit For   'assumes first blank defines end of data
		End If
	Next  I
End Function
Function BlankCell(ByVal X As Variant) As Boolean
'Determines if X is a blank cell or "--"

	If X = QNAN1 Or X = QNAN Then
        BlankCell = True
    Else
        BlankCell = False
    End If
End Function
Sub GetEmptyValues
	QNAN = "-1" & sD & "#QNAN"
	QNB = "-1" & sD & "#QNB"
	QNAN1 = "1" & sD & "#QNAN"
End Sub