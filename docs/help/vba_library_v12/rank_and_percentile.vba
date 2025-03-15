Option Explicit
Public Const ObjectHelp = Path + "\SPW.CHM"
Public HelpID As Variant
Dim Separator$
Sub Main
Separator = ListSeparator
'Authored by Frederick Cabasa 11/4/98
'Updated 12/22/99 John Kuo

'This macro computes ranks and cumulative percentages for a specified data column,
'and also reports the values for specified percentiles.

	HelpID = 60230			' Help ID number for this topic in SPW.CHM
	'Determine the data range and define the first empty column
	On Error GoTo ErrorMsg
	Dim ErrorCheck
	ErrorCheck = 0
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

If UsedColumns(1) = Empty Then GoTo ErrorMsg

DialogBox:
	ErrorCheck = 1
	Begin Dialog UserDialog 390,154,"Rank and Percentile",.Macrodlg ' %GRID:10,7,1,1
		Text 10,10,120,14,"&Data column",.Text1
		DropListBox 150,7,110,70,UsedColumns(),.First
		Text 10,39,130,14,"&Percentile column",.Text2
		DropListBox 150,35,110,70,UsedColumns(),.Second
		Text 10,66,130,14,"&First results column",.Text3
		TextBox 150,63,110,21,.Results
		GroupBox 10,91,250,56,"Percentile type",.GroupBox1
		OptionGroup .Percentiles
			OptionButton 20,107,220,14,"&Graphing (Cleveland method)",.OptionButton1
			OptionButton 20,124,110,14,"&Numeric",.OptionButton2
		OKButton 280,6,96,21
		CancelButton 280,36,96,21
		PushButton 280,74,96,21,"Help",.PushButton1
	End Dialog

Dim dlg As UserDialog
	'Default settings
	dlg.First = 0
	dlg.Second = 1
	dlg.Percentiles = 1

	Select Case Dialog(dlg)
		Case 0 'Handles Cancel button
			GoTo Finish
'		Case 1 'Handles Help button
'			HelpID = 60230			' Help ID number for this topic in SPW.CHM
'			Help(ObjectHelp,HelpID)
'			GoTo DialogBox
	End Select

'Parse the "First Empty" result
	WorksheetTable.GetMaxUsedSize(LastColumn, LastRow) 'Re-initializes variables
	If dlg.Results = "First Empty" Then
		dlg.Results = CStr(LastColumn + 1)
	Else
		dlg.Results = dlg.Results
	End If

If IsNumeric(dlg.Results)=False Or dlg.Results="" Then
	MsgBox "You must enter a valid number for your result column",vbExclamation,"Invalid Results Column"
	GoTo DialogBox
ElseIf IsNumeric(dlg.Results)=True Then
	If CLng(dlg.Results) < 1 Or CDbl(dlg.Results) < (LastColumn + 1) Then
		MsgBox "You must enter a postive integer greater than the last data column for your result column",vbExclamation,"Invalid Results Column"
		GoTo DialogBox
	End If
End If

	If dlg.Second <> 0 Then
		'Create an array of percentile column chosen
		Dim PColumn()As Variant
		ReDim PColumn(LastRow)
		Dim PCount As Integer
		For PCount = 0 To LastRow
			PColumn(PCount) = WorksheetTable.Cell(ListedColumns(dlg.Second) - 1, PCount)
			If PColumn(PCount) = "-1.#QNAN" Or PColumn(PCount) = "-1,#QNAN" Then
				ReDim Preserve PColumn(PCount - 1)
				GoTo Continue
			End If
		Next PCount
	End If

Continue:
	'Run Rank and Percentile Transform
	Dim SPTransform As Object
	Set SPTransform = ActiveDocument.NotebookItems.Add(9)
	SPTransform.Name = Path + "\Macro Transforms\RankPerc.xfm" 'Retrieves from default path
	SPTransform.Open
	Dim Expression As String
	Dim ColumnArg As String
	If dlg.Second = 0 Then
		Expression = "{" + """" + """" + "}"
		GoTo Transform
	End If
	Dim Limit
	Limit = UBound(PColumn)

	ColumnArg = CStr(ListedColumns(dlg.Second)) + Separator + CStr(1) + Separator + CStr(Limit + 1)
	Expression = "col(" + ColumnArg + ")"
Transform:
	SPTransform.AddVariableExpression("pvec", Expression)
	SPTransform.AddVariableExpression("c1", CStr(dlg.Results))
	SPTransform.AddVariableExpression("cvec", "{" + ListedColumns(dlg.First) + "}")
	SPTransform.AddVariableExpression("ptype", dlg.percentiles)
'	SPTransform.RunEditor 'Debug the transform
	ErrorCheck = 2
	SPTransform.Execute
	SPTransform.Close(False)

'Add Column Titles to Data and Results
	Dim Data_1 As String
	Dim Data_2 As String
	Dim Results_1 As String
	Dim Results_2 As String
	Dim Results_3 As String
	Dim Results_4 As String
	Dim Results_5 As String
	Dim Results_6 As String
	Dim Number As String
	Data_1 = "Data"
	Data_2 = "Percentiles"
	Results_1 = "Position"
	Results_2 = "Sorted Data"
	Results_3 = "Rank"
	Results_4 = "Cumulative %"
	Results_5 = "Percentiles "
	Results_6 = "Value"
	WorksheetTable.NamedRanges.Add(Data_1, ListedColumns(dlg.First)-1,0,1,-1, True)
	WorksheetTable.NamedRanges.Add(Data_2, ListedColumns(dlg.Second)-1,0,1,-1, True)
	Dim Marker
	Marker = CLng(dlg.Results) - 1
	WorksheetTable.NamedRanges.Add(Results_1,Marker,0,1,-1, True)
	WorksheetTable.NamedRanges.Add(Results_2,Marker+1,0,1,-1, True)
	WorksheetTable.NamedRanges.Add(Results_3,Marker+2,0,1,-1, True)
	WorksheetTable.NamedRanges.Add(Results_4,Marker+3,0,1,-1, True)
	If dlg.Second <> 0 Then
		WorksheetTable.NamedRanges.Add(Results_5,Marker+4,0,1,-1, True)
		WorksheetTable.NamedRanges.Add(Results_6,Marker+5,0,1,-1, True)
	End If
	GoTo Finish

ErrorMsg:
		If ErrorCheck = 0 Then
			HelpMsgBox 60230, "A worksheet with one data column and one percentiles column must be open", vbExclamation,"SigmaPlot
			GoTo Finish
		ElseIf ErrorCheck = 1 Then
			HelpMsgBox 60230, "You must enter a valid number for your results column", vbExclamation,"SigmaPlot"
			GoTo DialogBox
		ElseIf ErrorCheck = 2 Then
			SPTransform.Close(False)
			HelpMsgBox 60230, "Transform failed to execute.", vbExclamation,"SigmaPlot"
			GoTo Finish
		End If

Finish:
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
Public Function Macrodlg(DlgItem$, Action%, SuppValue&) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgText "Results", "First Empty"
	Case 2 ' Value changing or button pressed
	Select Case DlgItem$
		Case "PushButton1"
			Help(ObjectHelp,HelpID)
			Macrodlg = True 'do not exit the dialog
	End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem MessageBox = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function
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