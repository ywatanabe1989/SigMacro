Public Const ObjectHelp = Path + "\SPW.CHM"
Dim HelpID As Variant
Option Explicit

Sub Main
'Macro by Mohmammad Younus, 10/9/98
'Modified on 10/23/98
'Modified on 12/09/98
'Updated 12/8/99 John Kuo
'This macro computes the PSD (power spectral density) for a 1D array
'using the FFT function in the SigmaPlot transform language.  This
'macro can be useful in many signal processing applications.
'This macro also formats the color of the graphs.

HelpID = 60207			' Help ID number for this topic in SPW.CHM
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

Dim Hanning$(1)
Hanning$(0)="Yes"
Hanning$(1)="No"

'Dialog for source and results columns
MacroDialog:
	Begin Dialog UserDialog 395,128,"Power Spectral Density",.DialogFunc ' %GRID:10,7,1,0
		OKButton 300,10,80,21
		CancelButton 300,40,80,21
		PushButton 300,76,80,21,"Help",.PushButton1
		Text 12,14,102,14,"&Data column",.Text2
		DropListBox 140,10,140,57,UsedColumns(),.Y_data
		Text 12,43,120,14,"&First result column",.Text3
		TextBox 140,40,140,18,.ResultsCol
		Text 12,70,135,14,"&Sampling frequency",.Text4
		TextBox 174,68,106,18,.samplefreq
		Text 12,99,114,14,"&Hanning window",.Text1
		DropListBox 140,95,140,47,Hanning(),.hann
	End Dialog

Dim dlg As UserDialog
'Default settings
dlg.Y_data = 0
dlg.hann=0
If dlg.samplefreq = "" Then dlg.samplefreq = "1000"
If dlg.ResultsCol = "" Then dlg.ResultsCol = "First Empty"

Select Case Dialog(dlg)  
	Case 0 'Handles Cancel button
		GoTo Finish
'	Case 1 'Handles Help button
'		Dim ObjectHelp, HelpID As Variant
'		ObjectHelp = Path + "\SPW.CHM"
'		Help(ObjectHelp,HelpID)
'		GoTo MacroDialog
End Select

'Parse the "First Empty" result
WorksheetTable.GetMaxUsedSize(LastColumn,LastRow) 'Re-initialize variables
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
	If CLng(dlg.ResultsCol) < 1 Or CDbl(dlg.ResultsCol) < (LastColumn + 1) Then
		MsgBox "You must enter a postive integer greater than the last data column for your result column",vbExclamation,"Invalid Results Column"
		GoTo MacroDialog
	End If
End If

Dim Samfreq, Result
Samfreq =CVar(dlg.samplefreq)
If IsNumeric(Samfreq) = False Then GoTo note2
If Samfreq <=0 Then GoTo note2

'Open and run Powspec2 transform
Dim SPTransform As Object
Set SPTransform = ActiveDocument.NotebookItems.Add(9)
SPTransform.Name = Path + "\Macro Transforms\Powspec2.xfm" 'Retrieves from default path
SPTransform.Open
SPTransform.AddVariableExpression("ci", ListedColumns(dlg.Y_data))
SPTransform.AddVariableExpression("fs", dlg.samplefreq)
SPTransform.AddVariableExpression("han", CLng(dlg.hann))
SPTransform.AddVariableExpression("co", CInt(dlg.ResultsCol))

'SPTransform.RunEditor  'debug the transform
SPTransform.Execute
SPTransform.Close(False)

'Add column titles to results
Dim Results_1 As String
Dim Results_2 As String
Results_1 = "Frequency"
Results_2 = "Power Spectral Density"
WorksheetTable.NamedRanges.Add(Results_1,CLng(dlg.ResultsCol)-1,0,1,-1, True)
WorksheetTable.NamedRanges.Add(Results_2,CLng(dlg.ResultsCol),0,1,-1, True)

'Plot the graphs
Dim SPPage
Set SPPage = ActiveDocument.NotebookItems.Add(2)  'Creates graph page

'Create dynamic array for Plotted columns to handle problem with memory allocation
Dim PlottedColumns()As Variant
ReDim PlottedColumns(0)
PlottedColumns(0) = ListedColumns(dlg.Y_data) - 1
SPPage.CreateWizardGraph("Line Plot", _
	"Simple Straight Line","Single Y" ,PlottedColumns)
ReDim PlottedColumns(1)
PlottedColumns(0) = CLng(dlg.ResultsCol)-1 
PlottedColumns(1) = CLng(dlg.ResultsCol)
SPPage.CreateWizardGraph("Line Plot", _
	"Simple Straight Line","XY Pair",PlottedColumns)

'Add Axis & titles to the first graph 
Dim SPGraph1, XAxis1, YAxis1, SPLine1
Set SPGraph1 = SPPage.GraphPages(0).Graphs(0)
SPGraph1.Name = ""
Set XAxis1 = SPGraph1.Axes(0)
Set YAxis1 = SPGraph1.Axes(1)
XAxis1.Name = "Time"
YAxis1.Name = "Amplitude"

'Size and Position first graph
SPGraph1.Top=4000
SPGraph1.Left=-2750
SPGraph1.Width=6000
SPGraph1.Height=3000
SPGraph1.SetAttribute(SGA_PLANECOLORXYBACK, &H00e8ffff)

'Line Color of 1st graph
Set SPLine1 = SPPage.GraphPages(0).Graphs(0).Plots(0).Line
SPLine1.SetAttribute(SEA_COLOR, &H00000080)



Dim SPGraph2, XAxis2, YAxis2,SPLine2
Set SPGraph2 = SPPage.GraphPages(0).Graphs(1)
'Size and Position of 2nd graph
SPGraph2.Top=-500
SPGraph2.Left=-2750
SPGraph2.Width=6000
SPGraph2.Height=3000
SPGraph2.SetAttribute(SGA_PLANECOLORXYBACK, &H00e8ffff)

'Changing the font type and size 
'SPGraph2.SetAttribute(STA_BOLD,True)
SPGraph2.SetAttribute(STA_SIZE,200)

'sets 2nd graph Tick Line length
ActiveDocument.CurrentPageItem.Select(False, 533, -1750, 533, -1750)
ActiveDocument.CurrentPageItem.Select(False, 533, -1750, 533, -1750)
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, &H00000308, 1)
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, &H00000308, 1)
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, &H0000040a, 2)
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, &H0000042e, 79)
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, &H00000601, 2)
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, &H00000410, &H0001545c)
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, &H00000411, &H0000000c)
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, &H0000040a, 2)

'Line Thickness and color of 2nd graph
Set SPLine2 = SPPage.GraphPages(0).Graphs(1).Plots(0).Line
SPLine2.SetAttribute(SEA_THICKNESS,15)
SPLine2.SetAttribute(SEA_COLOR, &H00000080)

'Add Axis & titles to the 2nd graph
SPGraph2.Name = ""
Set XAxis2 = SPGraph2.Axes(0)
Set YAxis2 = SPGraph2.Axes(1)
XAxis2.Name = Results_1
YAxis2.Name = Results_2

'Change Y-scale of the 2nd graph to Log Common
Dim SPYAxis
Set SPYAxis = SPPage.GraphPages(0).Graphs(1).Axes(1)
SPYAxis.SetAttribute(SAA_TYPE,SAA_TYPE_COMMON)  'comment this to use a linear scale

'Clear the Legends
Dim SPLegend1, SPLegend2
Set SPLegend1 = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).AutoLegend
SPLegend1.ChildObjects(1).SetAttribute(STA_OPTIONS, 4960)  
Set SPLegend2 = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(1).AutoLegend
SPLegend2.ChildObjects(1).SetAttribute(STA_OPTIONS, 4960)  
GoTo Finish

NotAnInteger:
MsgBox "The result column must be a positive integer",vbExclamation,"Invalid Results Column"
GoTo MacroDialog

note2:
MsgBox "Please enter a positive numeric value for the Sampling Frequency",vbExclamation,"SigmaPlot"
GoTo MacroDialog

NoData:
HelpMsgBox 60207, "You must have a worksheet with at least one column of data open",vbExclamation,"No Open Worksheet"

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
HelpID = ID
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