Option Explicit
Dim Separator$
Public ObjectHelp As Variant
Public HelpID As Variant
Sub HelpPath
	ObjectHelp = Path + "\SPW.CHM"
End Sub
Sub Main
Separator = ListSeparator
'Macro by Mohammad Younus, 10/15/98
'Modified on 10/26; Updated 12/7/99, 12/21/99 John Kuo; Updated 4/20/10 Dick Mitchell
'This macro requires that two nonlinear curve fits to have been preformed with the regression wizard,
'and assumes that the equations are different only by the level of parameterization, e.g. first order 
'vs. second order, etc.  The macro automatically searches for the parameter and residual columns
'created by the regression wizard and compares the two fits using an F-test approximation.

HelpID = 60204			' Help ID number for this topic in SPW.CHM
Dim CurrentWorksheet
On Error GoTo NoData
CurrentWorksheet = ActiveDocument.CurrentDataItem.Name
ActiveDocument.NotebookItems(CurrentWorksheet).Open 'Opens/select default worksheet and sets focus
'Place Worksheet into Overwrite mode
ActiveDocument.NotebookItems(CurrentWorksheet).InsertionMode = False

'Determine the data range and define the first empty column
Dim WorksheetTable As Object
Set WorksheetTable = ActiveDocument.NotebookItems(CurrentWorksheet).DataTable
Dim LastColumn As Long
Dim LastRow As Long
LastColumn = 0
LastRow = 0 
WorksheetTable.GetMaxUsedSize(LastColumn,LastRow)


'Sort through columns and create list of columns with values in row 1
On Error GoTo EmptyWorksheet
Dim Index, UsedColumns$(), ListedColumns(), ListIndex, ColContents, ColTitle
ReDim UsedColumns$(LastColumn -1)
ReDim ListedColumns(LastColumn -1)
Dim Res1, Res2, Param1, Param2
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
		Case "Residuals" 'Looks for existence of Residuals columns
			UsedColumns$(Index) = ColTitle'use title
			Res1 = ListIndex
			ListedColumns(ListIndex) = CStr(Index + 1)
			ListIndex = ListIndex + 1
		Case "Residuals 1" 'Looks for existence of Residuals columns
			UsedColumns$(Index) = ColTitle	'use title
			Res2 = ListIndex
			ListedColumns(ListIndex) = CStr(Index + 1)
			ListIndex = ListIndex + 1
		Case "Parameters"
			UsedColumns$(Index) = ColTitle	'use title
			Param1 = column_size(Index,LastRow) 'Sets Param1 to size of Parameters column
			ListedColumns(ListIndex) = CStr(Index + 1)
			ListIndex = ListIndex + 1
		Case "Parameters 1"
			UsedColumns$(Index) = ColTitle	'use title
			Param2 = column_size(Index,LastRow)'Sets Param2 to size of Parameters 1 column
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

'Dialog for source and results columns
MacroDialog:
On Error GoTo Errors
	Begin Dialog UserDialog 475,185,"F-test Comparison of Curves",.DialogFunc ' %GRID:10,7,1,0
		OKButton 382,15,80,20
		CancelButton 382,47,80,21
		PushButton 382,79,80,21,"Help",.PushButton1
		Text 12,14,206,25,"Number of parameters for &1st function (with fewer parameters)",.Text1
		TextBox 220,17,135,19,.Para1
		Text 12,46,192,28,"Number of parameters for &2nd function",.Text2
		TextBox 220,49,135,19,.Para2
		Text 12,85,203,14,"&Residual column for function 1",.Text3
		DropListBox 220,81,138,72,UsedColumns(),.Function1
		Text 12,116,200,14,"R&esidual column for function 2",.Text4
		DropListBox 220,113,138,72,UsedColumns(),.Function2
		Text 12,148,120,14,"&First Result Column",.Text5
		TextBox 220,145,135,18,.ResultsCol
	End Dialog
Dim dlg As UserDialog
'Default settings

If dlg.Para1 = "" Then dlg.Para1 = CStr(Param1)
If dlg.Para2 = "" Then dlg.Para2 = CStr(Param2)
If dlg.ResultsCol = "" Then dlg.ResultsCol = "First Empty"
dlg.Function1 = Res1 
dlg.Function2 = Res2 

Select Case Dialog(dlg)  
	Case 0 'Handles Cancel button
		GoTo Finish
'	Case 1 'Handles Help button
			'Dim ObjectHelp, HelpID As Variant
			'ObjectHelp = Path + "\SPW.CHM"
'			HelpID = 60204			' Help ID number for this topic in SPW.CHM
'			Help(ObjectHelp,HelpID)
'		GoTo MacroDialog
End Select

If dlg.Para1 = "" Or dlg.Para2 = "" Then GoTo note4
'Parse the "First Empty" result
WorksheetTable.GetMaxUsedSize(LastColumn,LastRow) 'Re-initialize variables
If 	dlg.ResultsCol = "First Empty" Then
	dlg.ResultsCol = CStr(LastColumn + 1)
Else
	dlg.ResultsCol = dlg.ResultsCol
End If

If IsNumeric(dlg.ResultsCol)=False Or dlg.ResultsCol="" Then
	MsgBox "You must enter a valid number for your result column",vbExclamation,"Invalid Results Column"
	GoTo MacroDialog
ElseIf IsNumeric(dlg.ResultsCol)=True Then
	If CLng(dlg.ResultsCol) < 1 Or CDbl(dlg.ResultsCol) < (LastColumn + 1) Then
		MsgBox "You must enter a postive integer greater than the last data column for your result column",vbExclamation,"Invalid Results Column"
		GoTo MacroDialog
	End If
End If

'Handle for non-numeric value
	If IsNumeric(dlg.Para1) = False Or IsNumeric(dlg.Para2) = False Then
		MsgBox "Please enter a number for your parameters", vbInformation, "Numeric Parameter"
		GoTo MacroDialog
	End If

'Error handling
Errors:
Dim ParameterOne, ParameterTwo, Result
Result=CDbl(dlg.ResultsCol)
ParameterOne =CDbl(dlg.Para1)
ParameterTwo=CDbl(dlg.Para2)
If ParameterOne<0 Then GoTo note2
If ParameterTwo<1 Then GoTo note3
If ParameterTwo<=ParameterOne Then GoTo note3

'Open and run F_Test transform
Dim SPTransform As Object
Set SPTransform = ActiveDocument.NotebookItems.Add(9)
'SPTransform.Name = Path + "\Macro Transforms\F_Test.xfm" 'No longer used; transform is embedded in macro code
SPTransform.Open

Dim n1, n2, cs1, cs2, cres
n1 = dlg.Para1
n2 = dlg.Para2
cs1 = ListedColumns(dlg.Function1)
cs2 = ListedColumns(dlg.Function2)
cres = CInt(dlg.ResultsCol)
'Approximate P value for F distribution A&S, Eq. 26.6.15, p. 947 'not used 4-20-10, use Jake's fdist()
'Normal distribution approximation for P value A&S, Eq. 26.2.17, p 932
SPTransform.Text = "N=size(col(" +cs1+ "))" + vbCrLf + _
"ss1=total(col(" +cs1+ ")^2)"  + vbCrLf + _
"ss2=total(col(" +cs2+ ")^2)" + vbCrLf + _
"F = ((ss1-ss2)/ss2)*((N-" +n2+ ")/(" +n2+ "-" +n1+ "))" + vbCrLf + _
"N1=" +n2+ "-" +n1+ vbCrLf + _
"N2=N-" +n2+ vbCrLf + _
"p=1-fdist(F" + Separator + n1 + Separator + n2+ ")" + vbCrLf + _
"col(" +cres+ ")={" + Chr(34) + "F =" + Chr(34) + Separator + " " + Chr(34) + "p =" + Chr(34) + "}" + vbCrLf + _
"col(" +cres+ "+1)={F" + Separator + "p}" + vbCrLf + _
"col(" +cres+ "+2)= N" +vbCrLf

'"x=(F^(1/3)*(1-2/(9*N2))-(1-2/(9*N1)))/sqrt(2/(9*N1)+F^(2/3)*2/(9*N2))" + vbCrLf + _
'"pi=3.1415926" + vbCrLf + _
'"z=Exp(-x^2/2)/sqrt(2*pi)" + vbCrLf + _
'"t=1/(1+.2316419*x)" + vbCrLf + _
'"p=z*(.31938153*t-.356563782*t^2+1.781477937*t^3-1.821255978*t^4+1.330274429*t^5)" + vbCrLf + _
'"col(" +cres+ ")={" + Chr(34) + "F =" + Chr(34) + Separator + " " + Chr(34) + "p =" + Chr(34) + "}" + vbCrLf + _
'"col(" +cres+ "+1)={F" + Separator + "p}" + vbCrLf

'SPTransform.RunEditor 'for debugging transform
SPTransform.Execute
SPTransform.Close(False)

'Generating a report for F and p values
Dim N As Long
N = ActiveDocument.NotebookItems(CurrentWorksheet).DataTable.Cell(cres+1,0)
ActiveDocument.NotebookItems(CurrentWorksheet).DataTable.Cell(cres+1,0) = ""
Dim var1,var2,cres1
cres1=CLng(dlg.ResultsCol)
var1= ActiveDocument.NotebookItems(CurrentWorksheet).DataTable.Cell(cres1,0)
var2= ActiveDocument.NotebookItems(CurrentWorksheet).DataTable.Cell(cres1,1)

Dim SPReport As Object
Set SPReport = ActiveDocument.NotebookItems.Add(CT_REPORT)
SPReport.Name = "F-test " + SPReport.Name

If var1< 0.001 Then
	var1 = "< 0.001"
Else
	var1 = " = " + Format$(var1,"0.###")+" ("+CStr(n2-n1)+","+CStr(N-n2)+")"
End If

If var2>=0.05 Then
	var2 = " = " + Format$(var2, "0.####")
	GoTo label2 
End If
If var2<0.05 Then
	If var2<0.0001 Then
		var2 = "< 0.0001"
	Else
		var2 = " = " + Format$(var2, "0.####")
	End If
	GoTo label1
End If

label1:
SPReport.Text = "F(DFn,DFd) " + var1 +vbCrLf+"P "+var2+vbCrLf+"The more complex equation provides a significantly better fit." + _
vbCrLf + vbCrLf + "F: F value" + vbCrLf + "DFn: numerator degrees of freedom" + vbCrLf + "DFd: denominator degrees of freedom"
GoTo finish

label2:
SPReport.Text = "F(DFn,DFd) " + var1 +vbCrLf+"P "+var2+vbCrLf+"The more complex equation does not provide a significantly better fit." + _
vbCrLf + vbCrLf + "F: F value" + vbCrLf + "DFn: numerator degrees of freedom" + vbCrLf + "DFd: denominator degrees of freedom"
GoTo Finish

note2:
MsgBox "Please enter a positive number of parameters for 1st equation with fewer parameters",vbExclamation,"SigmaPlot"
GoTo MacroDialog

note3:
MsgBox "Please enter a positive number of parameters for 2nd equation with higher number of parameters",vbExclamation,"SigmaPlot"
GoTo MacroDialog

note4:
MsgBox "Please enter numbers of parameters.  The 2nd equation must have more parameters than the 1st.",vbExclamation,"SigmaPlot"
GoTo MacroDialog

NoData:
HelpMsgBox 60204, "You must have a worksheet open.",vbExclamation,"No Open Worksheet"
GoTo Finish

EmptyWorksheet:
HelpMsgBox 60204, "You must have the results of two curve fits in your worksheet, with residual and parameter values saved",vbExclamation,"SigmaPlot"

Finish:
End Sub

Public Function column_size(column As Variant, column_end As Variant)
'Returns column size of entries within a range
	Dim WorksheetTable As Object
	Set WorksheetTable = ActiveDocument.CurrentDataItem.DataTable
	Dim i
	Dim size As Long
		i = 0
		size = 0
	Do 
		If WorksheetTable.Cell(column,i) <> "-1.#QNAN" Then
		  If WorksheetTable.Cell(column,i) <> "-1,#QNAN" Then size = size + 1
		End If
		i = i + 1
	Loop Until i = column_end
	column_size = size
End Function
Public Function empty_col(column As Variant, column_end As Variant)
'Determines if a column is empty
	Dim WorksheetTable As Object
	Set WorksheetTable = ActiveDocument.CurrentDataItem.DataTable
	Dim i As Long
	Dim empty_cell As Boolean
	
	For i = 0 To column_end
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