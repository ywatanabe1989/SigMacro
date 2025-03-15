Option Explicit
Public Const ObjectHelp = Path + "\SPW.CHM"
Public HelpID As Variant
Dim sL
Function FlagOn(flag As Long)
    FlagOn = flag Or FLAG_SET_BIT
End Function
Function FlagOff(flag As Long)
    FlagOff = flag Or FLAG_CLEAR_BIT
End Function
Public Not_1_or_0 As Boolean

'Modified by Mohammad Younus On 10/30/98
'Modified On 11/02/98
'Macro by John Kuo.  Updated 12/21/99
'This macro computes and graphs a Kaplan-Meier survival curve. Specify the
'column containing the survival data, as well as the column indicating
'censoring of cases. A value of 0 in the censoring column indicates that
'the case is censored, whereas a value of 1 indicates an uncensored case.

Sub Main
HelpID = 60212			' Help ID number for this topic in SPW.CHM
sL = ListSeparator  'international list separator
Dim CurrentWorksheet
On Error GoTo NoData
CurrentWorksheet = ActiveDocument.CurrentDataItem.Name
ActiveDocument.NotebookItems(CurrentWorksheet).Open 'Opens/select default worksheet and sets focus

'Determine the data range and define the first empty column
Dim WorksheetTable As Object
Set WorksheetTable = ActiveDocument.NotebookItems(CurrentWorksheet).DataTable
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
			UsedColumns$(Index) = ColTitle 'If title is present use title
			ListedColumns(ListIndex) = CStr(Index + 1)
			ListIndex = ListIndex + 1
		End Select
	End If
	NextIndex:
Next Index

'Dialog for source and results columns
MacroDialog:
	Begin Dialog UserDialog 440,250,"Survival Curve",.DialogFunc ' %GRID:10,7,1,0
		OKButton 330,8,96,20
		CancelButton 330,36,96,21
		PushButton 330,92,96,21,"Help",.PushButton1
		GroupBox 8,3,300,95,"Survival data",.GroupBox2
		Text 20,21,120,14,"Survival data col",.Text1
		DropListBox 150,16,148,72,UsedColumns(),.SourceCol
		Text 20,46,130,14,"Censored data col",.Text2
		DropListBox 150,42,148,72,UsedColumns(),.CensorCol
		Text 20,73,120,14,"First results col",.Text3
		TextBox 150,70,148,19,.ResultsCol
		GroupBox 8,103,300,91,"Graph titles",.GroupBox3
		Text 20,122,90,13,"&Graph title",.Text5
		TextBox 150,119,148,18,.GraphTitle
		Text 20,145,72,15,"&X axis title",.Text6
		TextBox 150,142,148,18,.X_AxisTitle
		Text 19,169,80,13,"&Y axis title",.Text7
		TextBox 150,166,148,18,.Y_AxisTitle
		GroupBox 8,196,423,46,"",.GroupBox1
		Text 18,208,408,24,"Sort your survival data by time.   Use a 0 to indicate censored data, or 1 for uncensored.  Place censored ties last.",.Text4
	End Dialog

Dim dlg As UserDialog
'Default settings
dlg.SourceCol = 0
dlg.CensorCol = 1
If dlg.Y_AxisTitle="" Then dlg.Y_AxisTitle="Survival Probability"
If dlg.X_AxisTitle="" Then dlg.X_AxisTitle="Survival Time"
If dlg.GraphTitle="" Then dlg.GraphTitle="Survival Curve"
If dlg.ResultsCol = "" Then dlg.ResultsCol = "First Empty"

	Select Case Dialog(dlg)
		Case 0 'Handles Cancel button
			GoTo Finish
'		Case 1 'Handles Help button
'			HelpID = 60212			' Help ID number for this topic in SPW.CHM
'			Help(ObjectHelp,HelpID)
'			GoTo MacroDialog
	End Select

'Parse the "First Empty" result
WorksheetTable.GetMaxUsedSize(LastColumn,LastRow) 'Re-initialize variables
If dlg.ResultsCol = "First Empty" Then
	dlg.ResultsCol = CStr(LastColumn + 1)
Else
	dlg.ResultsCol = dlg.ResultsCol
End If

'Error handling for Results
If IsNumeric(dlg.ResultsCol)=False Or dlg.ResultsCol="" Then
	MsgBox "You must enter a valid number for your result column",vbExclamation,"Invalid Results Column"
	GoTo MacroDialog
ElseIf IsNumeric(dlg.ResultsCol)=True Then
	If CLng(dlg.ResultsCol) < 1 Or CDbl(dlg.ResultsCol) < (LastColumn + 1) Then
		MsgBox "You must enter a postive integer greater than the last data column for your result column",vbExclamation,"Invalid Results Column"
		GoTo MacroDialog
	End If
End If

Dim Censored
Censored = CDbl(ListedColumns(dlg.CensorCol))-1
Not_1_or_0 = False
CensorData(Censored, LastRow)
If Not_1_or_0 = True Then GoTo BadCensorData

'Open and run survival transform
Dim SPTransform As Object
Set SPTransform = ActiveDocument.NotebookItems.Add(9)
SPTransform.Open
SPTransform.Text =	"sur=col("+ ListedColumns(dlg.SourceCol) + ")" + vbCrLf + _
					"cen=col(" + ListedColumns(dlg.CensorCol) + ")" + vbCrLf + _
					"res=" + CStr(dlg.ResultsCol) + vbCrLf + _
					"mv=0/0" + vbCrLf + _
					"i=data(1" + sL + "size(sur))" + vbCrLf + _
					"N=size(sur)" + vbCrLf + _
					"n=N+1" + vbCrLf + _
					"pi=(N-i+1-cen)/(N-i+1)" + vbCrLf + _
					"cs=10^(sum(Log(pi)))" + vbCrLf + _
					"se=cs*sqrt(sum(cen/((N-i)*(N-i+1))))" + vbCrLf + _
					"xdat=If(cen=0" + sL + " sur)" + vbCrLf + _
					"ydat=If(cen=0" + sL + " cs)" + vbCrLf + _
					"col(res)={0" + sL + "sur}" + vbCrLf + _
					"col(res+1)={1" + sL + "cs}" + vbCrLf + _
					"col(res+2)={0" + sL + "se}" + vbCrLf + _
					"col(res+3)=xdat" + vbCrLf + _
					"col(res+4)=ydat" + vbCrLf
'SPTransform.RunEditor 'Debug transform code
SPTransform.Execute
SPTransform.Close(False)

'Add column titles to results
Dim Source, Censor, Results_1, Results_2, Results_3, Results_4, Results_5 As String
Source = "Survival Time"
Censor = "Censor"
Results_1 = "Time"
Results_2="Cum Prob"
Results_3="Cum Prob SE"
Results_4="X symbol"
Results_5="Y symbol"
WorksheetTable.NamedRanges.Add(Source,ListedColumns(CLng(dlg.SourceCol))-1,0,1,-1, True)
WorksheetTable.NamedRanges.Add(Censor,ListedColumns(CLng(dlg.CensorCol))-1,0,1,-1, True)
WorksheetTable.NamedRanges.Add(Results_1,CLng(dlg.ResultsCol)-1,0,1,-1, True)
WorksheetTable.NamedRanges.Add(Results_2,CLng(dlg.ResultsCol),0,1,-1, True)
WorksheetTable.NamedRanges.Add(Results_3,CLng(dlg.ResultsCol)+1,0,1,-1, True)
WorksheetTable.NamedRanges.Add(Results_4,CLng(dlg.ResultsCol)+2,0,1,-1, True)
WorksheetTable.NamedRanges.Add(Results_5,CLng(dlg.ResultsCol)+3,0,1,-1, True)

'Plot survival curve
Dim SPPage
Set SPPage = ActiveDocument.NotebookItems.Add(2)  'Creates graph page
Dim PlottedColumns()As Variant
ReDim PlottedColumns(1)
	PlottedColumns(0) = CLng(dlg.ResultsCol) -1
	PlottedColumns(1) = CLng(dlg.ResultsCol)
SPPage.CreateWizardGraph("Line Plot", _
	"Simple Horizontal Step Plot","XY Pair",PlottedColumns)

'Line Thickness
Dim SPPlot,SPLine
Set SPPlot=SPPage.GraphPages(0).Graphs(0).Plots(0)
Set SPLine=SPPlot.Line
SPLine.SetAttribute(SEA_THICKNESS,15)

'Plot symbols
Dim CensorPlot, CensorSymbols
ReDim PlottedColumns(1)
PlottedColumns(0) = CLng(dlg.ResultsCol) + 2
PlottedColumns(1) = CLng(dlg.ResultsCol) + 3

If WorksheetTable.Cell(PlottedColumns(0),0) <> "-1.#QNAN" Then
	If WorksheetTable.Cell(PlottedColumns(0),0) <> "-1,#QNAN" Then  'If there is no censored data, don't try and plot it
		SPPage.AddWizardPlot("Scatter Plot", "Simple Scatter","XY Pair",PlottedColumns)
		Set CensorPlot = SPPage.GraphPages(0).Graphs(0).Plots(1)
		CensorPlot.SetAttribute(SSA_SHAPE, 9)
	End If
End If

'Clears The Legend
ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).SetAttribute(SGA_FLAGS, _
							FlagOff(SGA_FLAG_AUTOLEGENDSHOW))

'Add Axis titles
Dim SPGraph, XAxis, YAxis
Set SPGraph = SPPage.GraphPages(0).Graphs(0)
SPGraph.Name = dlg.GraphTitle
Set XAxis = SPGraph.Axes(0)
Set YAxis = SPGraph.Axes(1)
XAxis.Name = dlg.X_AxisTitle
YAxis.Name = dlg.Y_AxisTitle

'X and Y-axis Lower and upper range
YAxis.SetAttribute(SAA_OPTIONS,SAA_FLAG_AUTORANGE Or FLAG_CLEAR_BIT)
YAxis.SetAttribute(SAA_FROMVAL, 0)
YAxis.SetAttribute(SAA_TOVAL, 1.05)
YAxis.SetAttribute(SAA_MAJORFREQINDIRECT, 0.2)

XAxis.SetAttribute(SAA_OPTIONS,SAA_FLAG_AUTORANGE Or FLAG_CLEAR_BIT)
XAxis.SetAttribute(SAA_OPTIONS,SAA_FLAG_AUTORANGEMIN Or FLAG_CLEAR_BIT)
XAxis.SetAttribute(SAA_OPTIONS,SAA_FLAG_AUTORANGEMAX Or FLAG_SET_BIT)
XAxis.SetAttribute(SAA_FROMVAL, 0)

GoTo Finish

NoData:
HelpMsgBox 60212, "You must have a worksheet open and in focus",vbExclamation,"No Open Worksheet"
GoTo Finish

BadCensorData:
HelpMsgBox 60212, "The censored data column must consist of only 1's or 0's",vbExclamation,"SigmaPlot"
GoTo Finish

EmptyWorksheet:
HelpMsgBox 60212, "You must have survival and censor data in your worksheet.  The censored data column must consist of only 1's or 0's",vbExclamation,"SigmaPlot"

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
Public Function CensorData(column As Variant, column_end As Variant)
'Determines if a column is only 1's and 0's
	Dim WorksheetTable As Object
	Set WorksheetTable = ActiveDocument.CurrentDataItem.DataTable
	Dim i As Long
	For i = 0 To column_end Step 1 'Change the step value to change the sampling interval.  Small sample size = Slow operation
		If WorksheetTable.Cell(Column,i) <> 0 And WorksheetTable.Cell(Column,i) <> 1 Then Not_1_or_0 = True
	Next i
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