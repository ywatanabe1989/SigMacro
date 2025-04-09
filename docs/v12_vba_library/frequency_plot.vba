Option Explicit
Public Const ObjectHelp = Path + "\SPW.CHM"
Public HelpID As Variant
Function FlagOn(flag As Long)
  FlagOn = flag Or FLAG_SET_BIT ' Use to set option flag bits on, leaving others unchanged
End Function
Function FlagOff(flag As Long)
  FlagOff = flag Or FLAG_CLEAR_BIT ' Use to set option flag bits off, leaving others unchanged
End Function
Sub Main
'Authored by Mohammad Younus, 10/15/98 
'Modified on 10/23/98; Updated 12/2/99 John Kuo
'This macro uses the frequecy plot transform to compute locations for the 
'symbols of a frequency (also known as density) plot.  Default binning and
'spacing values are computed from the data.

HelpID = 60203			' Help ID number for this topic in SPW.CHM
Dim ErrorCheck As Integer
ErrorCheck = 0 'Display no open worksheet error message on error
On Error GoTo ErrorMsg

Dim CurrentWorksheet
CurrentWorksheet = ActiveDocument.CurrentDataItem.Name
ActiveDocument.NotebookItems(CurrentWorksheet).Open  'Opens/selects default worksheet and sets focus

'Determine the data range and define the first empty column
Dim WorksheetTable As Object
Set WorksheetTable = ActiveDocument.NotebookItems(CurrentWorksheet).DataTable
Dim LastColumn As Long
Dim LastRow As Long
LastColumn = 0
LastRow = 0 
WorksheetTable.GetMaxUsedSize(LastColumn,LastRow)
'Make sure there is data that starts in column one
If LastColumn = 0 Or empty_col(0,LastRow)=True Then GoTo ErrorMsg
'Place Worksheet into Overwrite mode
ActiveDocument.NotebookItems(CurrentWorksheet).InsertionMode = False

'Mean Lines Array
Dim MeanLine$(2)
MeanLine$(0)="None"
MeanLine$(1)="Mean"
MeanLine$(2)="Median"

MacroDialog:
'Dialog for source and results columns
	Begin Dialog UserDialog 600,255,"Frequency Plot ",.DialogFunc ' %GRID:10,7,1,0
		OKButton 492,10,96,21
		CancelButton 492,38,96,21
		PushButton 492,74,96,21,"Help",.PushButton1
		GroupBox 12,7,245,75,"Column selection",.GroupBox1
		Text 25,28,128,14,"No. &data columns",.Text1
		TextBox 155,25,90,19,.x_data
		Text 25,56,120,14,"First &result column",.Text3
		TextBox 155,53,90,19,.ResultsCol
		GroupBox 12,88,245,75,"Bins",.GroupBox4
		Text 25,109,115,14,"Vertical &interval"
		TextBox 155,106,90,19,.VertIntvl
		Text 25,137,104,14,"Start &value"
		TextBox 155,134,90,19,.BinStart
		GroupBox 12,170,245,75,"Mean/median lines",.GroupBox5
		Text 25,191,63,14,"&Type",.Text6
		DropListBox 155,188,91,72,MeanLine(),.Mean
		Text 25,219,90,14,"Width (&x units)",.Text7
		TextBox 155,217,90,19,.LineWidth
		GroupBox 270,7,206,75,"Graph dimensions",.GroupBox2
		Text 284,28,82,14,"&Height (in)",.Text4
		TextBox 373,25,90,19,.High
		Text 284,56,84,14,"&Width (in)",.Text2
		TextBox 373,53,90,19,.Wide
		GroupBox 270,88,206,75,"Symbols",.GroupBox3
		Text 284,109,58,14,"&Size (in)",.Text5
		TextBox 373,106,91,19,.Size
		Text 284,137,82,14,"&Gap (% size)"
		TextBox 373,134,91,19,.SymbolGap
	End Dialog
Dim dlg As UserDialog

'Computing Default settings
WorksheetTable.GetMaxUsedSize(LastColumn,LastRow)'Reinitialize last column
If dlg.x_data = "" Then dlg.x_data = CStr(LastColumn)'Use all data by default
If dlg.Wide = "" Then dlg.Wide = "5"
If dlg.High ="" Then dlg.High ="3.5"
If dlg.Size ="" Then dlg.Size ="0.08"
If dlg.ResultsCol = "" Then dlg.ResultsCol = "First Empty"
dlg.Mean = 0
If dlg.LineWidth = "" Then dlg.LineWidth = "0.5"

Dim Data_Range, MaxValue, MinValue, RangeSize
WorksheetTable.GetMaxUsedSize(LastColumn,LastRow)
Data_Range = WorksheetTable.GetData(0,0,LastColumn-1,LastRow-1)
MinValue = min_array(Data_Range,LastColumn-1,LastRow-1) 'See Min_Array function at end
'Debug.Print MinValue
MaxValue = max_array(Data_Range,LastColumn-1,LastRow-1) 'See Max_Array function at end
RangeSize = MaxValue - MinValue
'Debug.Print MaxValue
If dlg.VertIntvl = "" Then dlg.VertIntvl = CStr(RangeSize/50)
If dlg.BinStart = "" Then dlg.BinStart = CStr(MinValue)
If dlg.SymbolGap = "" Then dlg.SymbolGap = CStr(10/LastRow*1.5)

'Clear working data
Dim Selection()
ReDim Selection(3)
Selection(0) = LastColumn
Selection(1) = 0
Selection(2) = LastColumn
'Selection(3) = 31999999
Selection(3) = LastRow
ActiveDocument.CurrentDataItem.SelectionExtent = Selection
ActiveDocument.CurrentDataItem.Clear
ActiveDocument.CurrentDataItem.Goto(0,0)

Select Case Dialog(dlg)  
	Case 0 'Handles Cancel button
		GoTo Finish
'	Case 1 'Handles Help button
'		HelpID = 60203			' Help ID number for this topic in SPW.CHM
'		Help(ObjectHelp,HelpID)
'		GoTo MacroDialog
	End Select

'Parse the "First Empty" result
WorksheetTable.GetMaxUsedSize(LastColumn,LastRow) 'Re-initialise variables
If 	dlg.ResultsCol = "First Empty" Then
	dlg.ResultsCol = CStr(LastColumn + 1)
Else
	dlg.ResultsCol = dlg.ResultsCol
End If

'Error Handling
If IsNumeric(dlg.ResultsCol)=False Or dlg.ResultsCol="" Then
	MsgBox "You must enter a valid number for your result column",vbExclamation,"Invalid Results Column"
	GoTo MacroDialog
ElseIf IsNumeric(dlg.ResultsCol)=True Then
	If CDbl(dlg.x_data)<=0 Or CDbl(dlg.x_data)> LastColumn Then
		MsgBox "Please always start from the first column and enter the correct nth column number of the data",vbExclamation,"Incorrect Number"
		GoTo MacroDialog
	End If
	If CLng(dlg.ResultsCol) < 1 Or CDbl(dlg.ResultsCol) < (LastColumn + 1) Then
		MsgBox "You must enter a postive integer greater than the last data column for your result column",vbExclamation,"Invalid Results Column"
		GoTo MacroDialog
	End If
End If



'Limiting width, height and symbol size of the graph
Dim GraphWidth, GraphHeight
GraphWidth=CDbl(dlg.Wide)
GraphHeight=CDbl(dlg.High)

If GraphWidth < 1  Then
	 GraphWidth = 1
ElseIf GraphWidth > 8.5 Then
       GraphWidth=8.5
End If

If GraphHeight < 1 Then
	GraphHeight = 1

ElseIf GraphHeight > 11 Then
	GraphHeight = 11
End If

Dim LimitSize
LimitSize=CDbl(dlg.size)
If LimitSize < .01  Then
	 LimitSize = .01
End If

If LimitSize > 1 Then
	LimitSize = 1
End If

'Open and run Freqplt2.xfm transform
Dim SPTransform As Object
Set SPTransform = ActiveDocument.NotebookItems.Add(9)
ErrorCheck = 1 'Display transform not found error
SPTransform.Name = Path + "\Macro Transforms\FreqPlt2.xfm" 'Retrieves from default path
SPTransform.Open
SPTransform.AddVariableExpression("n", dlg.x_data)
SPTransform.AddVariableExpression("d", LimitSize)
SPTransform.AddVariableExpression("wg", dlg.Wide)
SPTransform.AddVariableExpression("ml",CLng(dlg.Mean))
SPTransform.AddVariableExpression("eml",CDbl(dlg.LineWidth))
SPTransform.AddVariableExpression("w",CDbl(dlg.VertIntvl))
SPTransform.AddVariableExpression("ys",CDbl(dlg.BinStart))
SPTransform.AddVariableExpression("fx",CDbl(dlg.SymbolGap))
SPTransform.AddVariableExpression("result", CInt(dlg.ResultsCol))

SPTransform.Execute
SPTransform.Close(False)

'Add column titles to results
Dim FirstResultColumn, ResultCount
FirstResultColumn = dlg.ResultsCol
ResultCount = 1
Dim total_columns
total_columns = dlg.x_data

Do While total_columns > 0
	WorksheetTable.NamedRanges.Add("Group "+CStr(ResultCount)+" X",CLng(FirstResultColumn)-1,0,1,-1, True)
	WorksheetTable.NamedRanges.Add("Group "+CStr(ResultCount)+" Y",CLng(FirstResultColumn),0,1,-1, True)
	total_columns = total_columns - 1
	ResultCount = ResultCount + 1
	FirstResultColumn = FirstResultColumn + 2
Loop

'Create Frequency Plot 
Dim Ind As Long
Ind=((CLng(dlg.x_data))*2)-1
Dim SPPage
Set SPPage = ActiveDocument.NotebookItems.Add(2)  'Creates graph page
Dim PlottedColumns() As Variant
ReDim PlottedColumns(Ind) As Variant
Dim Index
Index = 0
Do While Index <=Ind
PlottedColumns(Index) = CLng(dlg.ResultsCol) -1 + Index
Index = Index + 1
Loop

'Create graph
SPPage.CreateWizardGraph("Scatter Plot","Multiple Scatter","XY Pairs",PlottedColumns)

'Add Mean Lines
If dlg.Mean=0 Then GoTo AddTitles
ReDim PlottedColumns(1) As Variant
PlottedColumns(0)= CLng(dlg.ResultsCol)+Ind
PlottedColumns(1)= CLng(dlg.ResultsCol)+Ind+1
SPPage.AddWizardPlot("Line Plot", _
	"Simple Straight Line","XY Pair",PlottedColumns)

AddTitles:
'Add Axis titles
Dim SPGraph, XAxis, YAxis
Set SPGraph = 	SPPage.GraphPages(0).CurrentPageObject(GPT_GRAPH)
SPGraph.Name = "Frequency Plot"
Set XAxis = SPGraph.Axes(0)
Set YAxis = SPGraph.Axes(1)
XAxis.Name = "Data"
YAxis.Name = "Percentage"

Dim GroupNumber, NumberCurves
NumberCurves = SPGraph.Plots(0).ChildObjects.Count
GroupNumber = 1
Do While NumberCurves > 0
	SPGraph.SetAttribute(SGA_NTHAUTOLEGEND, GroupNumber - 1)
	SPGraph.SetAttribute(SGA_CURRENTLEGENDTEXT, "Group "+ CStr(GroupNumber))
NumberCurves = NumberCurves - 1
GroupNumber = GroupNumber + 1
Loop

'Set symbol size
Dim  SPSymbols, SPPlot

Set SPPlot = SPGraph.Plots(0)
Set SPSymbols = SPPlot.Symbols

'Modify the following lines to set additional plot attributes
'SPSymbols.SetAttribute(SSA_SHAPE,3)           'shape of symbol
'SPSymbols.SetAttribute(SSA_COLOR,RGB_BLUE)     'color of symbol
'SPSymbols.SetAttribute(SSA_EDGECOLOR,RGB_RED) 'edge Color of symbol  
SPSymbols.SetAttribute(SSA_SIZE,(LimitSize*1000)) 'size of symbol

SPGraph.Width=(GraphWidth)*1000     ' Width of the graph
SPGraph.Height=(GraphHeight)*1000    'Height of the graph 

GoTo Finish

ErrorMsg:
	If ErrorCheck = 0 Then 
		HelpMsgBox 60203, "To use an existing graph, you must have an existing point plot open.",vbExclamation,"Point Plot Required"
	ElseIf ErrorCheck = 1 Then 
		HelpMsgBox 60203, "The frequency plot transform " + Chr(34) + "FreqPlt2.xfm" + Chr(34) + " was not found.",vbExclamation,"SigmaPlot"
		GoTo Finish		
	End If	

Finish:

End Sub
Public Function max_array(A As Variant, maxcolumn As Long, maxrow As Long)
'Computes the maximum value of the array A consisting of maxcolumn number of
'columns and maxrow number of rows.
	Dim i, j As Long
	Dim maxval As Variant
	maxval = A(0,0)
	For i = 0 To maxcolumn
		For j = 0 To maxrow
		If A(i,j) > maxval Then 
			maxval = A(i,j)
		End If
		Next j
	Next i
	max_array = maxval
End Function
Public Function min_array(A As Variant, maxcolumn As Long, maxrow As Long)
'Computes the minimum value of the array A consisting of maxcolumn number of
'columns and maxrow number of rows.
	Dim i, j As Long
	Dim minval As Variant
	minval = A(0,0)
	For i = 0 To maxcolumn
			For j = 0 To maxrow
		If A(i,j) < minval And A(i,j) <> "-1.#QNAN" And A(i,j) <> "-1,#QNAN" Then minval = A(i,j)
		Next j
	Next i
	min_array = minval
End Function
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

	HelpID=ID
	MsgBox Msg, MsgType, MsgTitle
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