Option Explicit
Public Const ObjectHelp = Path + "\SPW.CHM"
Public HelpID As Variant
Function FlagOn(flag As Long)
    FlagOn = flag Or FLAG_SET_BIT
End Function
Function FlagOff(flag As Long)
    FlagOff = flag Or FLAG_CLEAR_BIT
End Function
Sub Main
'Modified by Frederick Cabasa 11/6/98
'Updated 12/22/99 John Kuo

'This macro takes four columns:  X, Y, Length, and Angle, and plots the 
'data as a vector plot.

HelpID = 60213			' Help ID number for this topic in SPW.CHM
Dim ErrorCheck
ErrorCheck = 0
On Error GoTo NoData
ActiveDocument.CurrentDataItem.Open 'Opens/select default worksheet and sets focus

'Determine the data range and define the first empty column
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
On Error GoTo EmptyWorksheet
Dim Index, UsedColumns$(), ListedColumns(), ListIndex, ColContents, ColTitle
ReDim UsedColumns$(LastColumn -1)
ReDim ListedColumns(LastColumn -1)
ListIndex = 0
For Index = 0 To LastColumn - 1
	ColContents = empty_col(Index, LastRow) 
	ColTitle = WorksheetTable.Cell(Index,-1) 'Retrieve column title
	If ColContents = True Then GoTo NextIndex
	If ColContents = False Then   'If column is not empty
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
			UsedColumns$(Index) = ColTitle 'If title is present use title
			ListedColumns(ListIndex) = CStr(Index + 1)
			ListIndex = ListIndex + 1
		End Select
	End If
	NextIndex:
Next Index

Dim Angle$(4)
	Angle$(0) = "15"
	Angle$(1) = "20"
	Angle$(2) = "30"
	Angle$(3) = "45"

MacroDialog: 
ErrorCheck = 1
	Begin Dialog UserDialog 418,143,"Vector Plot",.DialogFunc ' %GRID:10,7,1,0
		OKButton 310,6,96,21
		CancelButton 310,34,96,21
		PushButton 310,72,96,21,"Help",.PushButton1
		Text 12,10,132,14,"First source &column",.SourceTitle
		DropListBox 150,7,140,72,UsedColumns(),.SourceCol
		Text 12,40,135,14,"Arrowhead angle",.ArrowTitle
		DropListBox 150,37,100,72,Angle(),.ArrowAngle
		Text 251,40,24,13,"‹",.Text1
		Text 12,71,135,14,"Arrowhead length",.ArrowTitle2
		TextBox 150,68,100,19,.ArrowLeng
		GroupBox 10,96,398,40,"",.GroupBox1
		Text 20,105,356,12,"Columns must be contiguous and in a specific order:",.Text2
		Text 23,118,374,12,"X data, Y data, Direction (radians), and Magnitude (length).",.Text3
	End Dialog
	Dim dlg As UserDialog
	dlg.ArrowAngle = 1
	If dlg.ArrowLeng = "" Then dlg.ArrowLeng = "0.1"
	
	Select Case Dialog(dlg)  
		Case 0 'Handles Cancel button
			GoTo Finish
'		Case 1 'Handles Help button
'			HelpID = 60213			' Help ID number for this topic in SPW.CHM
'			Help(ObjectHelp,HelpID)
'			GoTo MacroDialog 
	End Select

'Check for four populated columns
If empty_col(ListedColumns(dlg.SourceCol), LastRow) = True Then 
	MsgBox "Your Y data column is empty",vbExclamation,"SigmaPlot"
	GoTo MacroDialog
End If
If empty_col(ListedColumns(dlg.SourceCol) + 1, LastRow) = True Then
	MsgBox "Your length data column is empty",vbExclamation,"SigmaPlot"
	GoTo MacroDialog
End If
If empty_col(ListedColumns(dlg.SourceCol) + 2, LastRow) = True Then
	MsgBox "Your angle data column is empty",vbExclamation,"SigmaPlot"
	GoTo MacroDialog
End If

Dim pi
pi= 3.14159265359
Dim Arrow  'Define arrowhead angle values in radians
Select Case dlg.ArrowAngle
	Case 0
		Arrow = pi/12
	Case 1
		Arrow = pi/9
	Case 2
		Arrow = pi/6
	Case 3
		Arrow = pi/4
End Select

'Run vector.xfm transform
	Dim SPTransform As Object
	Set SPTransform = ActiveDocument.NotebookItems.Add(9)
	SPTransform.Name = Path + "\Macro Transforms\vector.xfm" 'Retrieves from default path
	SPTransform.TrigUnit = 0 'You must initialize trig units if you use trig functions
	SPTransform.Open
	SPTransform.AddVariableExpression("xc", ListedColumns(dlg.SourceCol))
	SPTransform.AddVariableExpression("L", dlg.ArrowLeng)
	SPTransform.AddVariableExpression("Angle", Arrow)
'	SPTransform.RunEditor 'debugging only
	SPTransform.Execute
SPTransform.Close(False)

'Add column titles to results
WorksheetTable.Cell(ListedColumns(dlg.SourceCol) + 3,-1) = "Body X"
WorksheetTable.Cell(ListedColumns(dlg.SourceCol) + 4,-1) = "Body Y"
WorksheetTable.Cell(ListedColumns(dlg.SourceCol) + 5,-1) = "Arrow Body X"
WorksheetTable.Cell(ListedColumns(dlg.SourceCol) + 6,-1) = "Arrow Body Y"
WorksheetTable.Cell(ListedColumns(dlg.SourceCol) + 7,-1) = "Arrowhead X"
WorksheetTable.Cell(ListedColumns(dlg.SourceCol) + 8,-1) = "Arrowhead Y"

'Plot results
Dim SPPage
Set SPPage = ActiveDocument.NotebookItems.Add (2)  'Creates graph page
Dim PlottedColumns(5) As Variant
Index = 0
Do While Index <=5
	PlottedColumns(Index) = ListedColumns(dlg.SourceCol) + 3 + Index
	Index = Index + 1
Loop

'Create the graph
SPPage.CreateWizardGraph("Line Plot","Multiple Straight Lines","XY Pairs",PlottedColumns)
Dim SPGraph, SPPLot
Set SPGraph = SPPage.GraphPages(0).Graphs(0)
SPGraph.Name = "Vector Plot" 'Change the graph title
Set SPPLot = SPGraph.Plots(0)
SPPLot.SetAttribute(SLA_SELECTFUNC,0) 'Deselect all functions to make sure plot routes line messages to line.
SPPLot.SetAttribute(SEA_TYPEREPEAT,SOA_REPEAT_SAME) 'Change the line increments to none

'Set Graph size and position
SPGraph.Top = 3000
SPGraph.Left = -2250
SPGraph.Width = 5000
SPGraph.Height = 5000

'Clear the legend
SPGraph.SetAttribute(SGA_FLAGS, FlagOff(SGA_FLAG_AUTOLEGENDSHOW))

GoTo Finish

NoData:
If ErrorCheck = 0 Then
	HelpMsgBox 60213, "You must have a worksheet open",vbExclamation,"No Open Worksheet"
Else
	HelpMsgBox 60213, "You must enter a postitive number for arrowhead length",vbExclamation
End If
GoTo Finish

EmptyWorksheet:
	HelpMsgBox 60213, "This macro requires four data columns:  X, Y, Length, and Angle",vbExclamation,"SigmaPlot"

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
Public Function udHelpBox(DlgItem$, Action%, SuppValue&) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		Select Case dlgItem$
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
Rem See DialogFunc help topic for more information.
Public Function DialogFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Select Case Action%
	Case 2 ' Value changing or button pressed
		Select Case DlgItem$
		Case "PushButton1"
			Help(ObjectHelp,HelpID)
			DialogFunc = True 'do not exit the dialog
        End Select
	End Select
End Function