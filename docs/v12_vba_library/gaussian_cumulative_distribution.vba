Option Explicit
Public Const ObjectHelp = Path + "\SPW.CHM"
Public HelpID As Variant
Dim Separator$
Dim PlotKludgeState As Boolean
Dim DecimalChar$
Function FlagOn(flag As Long)
  FlagOn = flag Or FLAG_SET_BIT ' Use to set option flag bits on, leaving others unchanged
End Function
Sub Main
HelpID = 80501			' Help ID number for this topic in SPW.HLP
Separator = ListSeparator
DecimalChar = DecimalSymbol
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

If LastColumn = 0 Then GoTo ErrorMsg

DialogBox:
	ErrorCheck = 1
	Begin Dialog UserDialog 350,168,"Cumulative Gaussian Distribution",.Macrodlg ' %GRID:10,7,1,1
		Text 10,7,120,14,"&Data column",.Text1
		DropListBox 10,21,140,105,UsedColumns(),.First
		Text 180,7,130,14,"&First results column",.Text3
		TextBox 180,21,140,21,.Results
		CheckBox 10,46,180,14,"&Plot results",.PlotResults
		CheckBox 10,64,290,14,"&Use probablity scale",.Probability
		OKButton 130,140,100,21
		CancelButton 240,140,100,21
		PushButton 10,140,100,21,"Help",.PushButton1
		GroupBox 10,84,330,42,"",.GroupBox1
		Text 20,93,310,28,"Computes the CDF for a single data column and optionally plots the results.",.Text4
	End Dialog

Dim dlg As UserDialog
	'Default settings
	dlg.First = 0
	dlg.Probability = 0
	If dlg.Results = "" Then dlg.Results = "First Empty"
	PlotKludgeState = False

StartAgain:
	Select Case Dialog(dlg)
		Case 0 'Handles Cancel button
			GoTo Finish
'		Case 1 'Handles Help button
'			Help(ObjectHelp,HelpID)
'			GoTo StartAgain
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

Continue:
	'Run CDF Transform
	Dim SPTransform As Object
	Set SPTransform = ActiveDocument.NotebookItems.Add(9)
	SPTransform.Open
	SPTransform.Text = 	"erf(x)=1-(" + DecimalChar + "3480242*terf(x)-" + DecimalChar + "0958798*terf(x)^2+" + DecimalChar + "7478556*terf(x)^3)*Exp(-x^2)" + vbCrLf + _
						"terf(x)=1/(1+" + DecimalChar + "47047*x)" + vbCrLf + _
						"erf1(x)=If(x<0" + Separator + "-erf(-x)" + Separator + "erf(x))" + vbCrLf + _
						"res=" + CStr(dlg.Results) + vbCrLf + _
						"x=sort(col(" + ListedColumns(dlg.First) + "))" + vbCrLf + _
						"m=mean(x)" + vbCrLf + _
						"s=stddev(x)" + vbCrLf + _
						"P(x)=(erf1(x/sqrt(2))+1)/2" + vbCrLf + _
						"col(res)=x" + vbCrLf + _
						"col(res+1)=P((x-m)/(s))" + vbCrLf + _
						"col(res+2)=col(res+1)*100"
'	SPTransform.RunEditor 'Debug the transform
	SPTransform.Execute
	SPTransform.Close(False)

'Add Column Titles to Results
	Dim Marker
	Marker = CLng(dlg.Results)
	WorksheetTable.NamedRanges.Add("Sorted Data",Marker-1,0,1,-1, True)
	WorksheetTable.NamedRanges.Add("CDF",Marker,0,1,-1, True)
	WorksheetTable.NamedRanges.Add("CDF*100",Marker+1,0,1,-1, True)

'Plot the graph
Dim PlottedColumns()As Variant
Dim SPPage As Object

If dlg.PlotResults = 1 Then
	Set SPPage = ActiveDocument.NotebookItems.Add(2)  'Creates graph page
	SPPage.Name = SPPage.Name + ": C.D.F."
	ReDim PlottedColumns(1)
'	PlottedColumns(0) = ListedColumns(dlg.First)
'	PlottedColumns(1) = CLng(dlg.Results)
	PlottedColumns(0) = CLng(dlg.Results)-1
	PlottedColumns(1) = CLng(dlg.Results)+1

	SPPage.CreateWizardGraph("Line Plot","Simple Straight Line","XY Pair" ,PlottedColumns)
	SPPage.GraphPages(0).Graphs(0).Name = "Gaussian Cumulative Distribution"
End If

If dlg.Probability = 1 Then
	SPPage.GraphPages(0).Graphs(0).Axes(1).SetAttribute(SAA_TYPE,SAA_TYPE_PROBABILITY)
End If

GoTo Finish

ErrorMsg:
		If ErrorCheck = 0 Then
			HelpMsgBox HelpID, "A worksheet with at least one data column must be open", vbExclamation,"SigmaPlot"
			GoTo Finish
		Else
			HelpMsgBox HelpID, "You must enter a valid number for your results column", vbExclamation,"SigmaPlot"
			GoTo DialogBox
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
		If PlotKludgeState = True Then
			DlgEnable "Probability", True
		Else
			DlgEnable "Probability", False
        End If
	Case 2 ' Value changing or button pressed
		Select Case DlgItem$
'		Case "Help"
		Case "PushButton1"
			Help(ObjectHelp,HelpID)
			Macrodlg = True 'do not exit the dialog
        Case "PlotResults"
        	If SuppValue = 0 Then
        		DlgEnable "Probability", False
        		DlgValue "Probability", 0
        		PlotKludgeState = False
        	Else
				DlgEnable "Probability", True
				PlotKludgeState = True
        	End If
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