Option Explicit
Public Const ObjectHelp = Path + "\SPW.CHM"
Public HelpID As Variant
Dim Separator$
Sub Main
HelpID = 60210			' Help ID number for this topic in SPW.CHM
Separator = ListSeparator
'Authored by 12/12/01 John Kuo
'Derivative.xfm transform by Dick Mitchell 11/28/01

' This transform computes a numerical first derivative of data.  It
' computes the running average of navg adjacent first order derivatives.
' The SigmaPlot transform language 'diff' function is used to compute
' the first order differences in x and y required for the numerical
' derivative.  The data need not be sorted by x.  Replicate x values
' and the associated y values are rowwise deleted to eliminate zero
' divides.  Using even values of navg will place each derivative at
' the midpoint of the derivatives used in the average.  Using odd
' values will place it at the first point to the left of midpoint.
' For even navg values, navg/2 cells will be empty at the beginning
' and end of the derivative values.

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
	Begin Dialog UserDialog 350,189,"Running Average of 1st Derivative",.DialogFunc ' %GRID:10,7,1,1
		Text 10,7,120,14,"First &data column",.Text1
		DropListBox 10,21,140,105,UsedColumns(),.First
		Text 200,7,130,14,"&First results column",.Text3
		TextBox 200,21,140,21,.Results
		Text 10,45,170,14,"&Length of running average",.Text2
		TextBox 10,59,140,21,.Navg
		CheckBox 200,49,130,14,"&Plot results",.PlotResults
		CheckBox 200,66,130,14,"Plot &original data",.PlotOriginal
		OKButton 130,161,100,21
		CancelButton 240,161,100,21
		PushButton 10,161,100,21,"Help",.PushButton1
		GroupBox 10,84,330,70,"",.GroupBox1
		Text 20,95,310,53,"This macro computes the running average of adjacent numerical derivatives.  The data need not be sorted by x.  Your x and y data columns must be adjacent.",.Text4
	End Dialog

Dim dlg As UserDialog
	'Default settings
	dlg.First = 0
	dlg.Navg = "1"
	If dlg.Results = "" Then dlg.Results = "First Empty"

	Select Case Dialog(dlg)
		Case 0 'Handles Cancel button
			GoTo Finish
'		Case 1 'Handles Help button
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

Continue:
	'Run Derivative Transform
	Dim SPTransform As Object
	Set SPTransform = ActiveDocument.NotebookItems.Add(9)
	SPTransform.Open
	SPTransform.Text = 	"navg=" + CStr(dlg.Navg) + vbCrLf + _
						"co=" + CStr(dlg.Results) + vbCrLf + _
						"ci=" + ListedColumns(dlg.First) + vbCrLf + _
						"wc=co+100" + vbCrLf + _
						"x=col(ci)" + vbCrLf + _
						"y=col(ci+1)" + vbCrLf + _
						"col(wc)=x" + vbCrLf + _
						"col(wc+1)=y" + vbCrLf + _
						"block(wc" + Separator + "1)=sort(block(wc" + Separator + "1)" + Separator + "col(wc))" + vbCrLf + _
						"xsort=col(wc)" + vbCrLf + _
						"ysort=col(wc+1)" + vbCrLf + _
						"xsortu=If(diff(xsort)!=0" + Separator + "xsort)" + vbCrLf + _
						"ysortu=If(diff(xsort)!=0" + Separator + "ysort)" + vbCrLf + _
						"col(wc+2)=xsortu" + vbCrLf + _
						"col(wc+3)=ysortu" + vbCrLf + _
						"uniquex=if(diff(xsort)=0" + Separator + "xsort)" + vbCrLf + _
						"for k = 1 to size(uniquex) do" + vbCrLf + _
						"y_for_uniquex_i=If(xsort=uniquex[k]" + Separator + "ysort)" + vbCrLf + _
						"yimean = mean(y_for_uniquex_i)" + vbCrLf + _
						"cell(wc+4" + Separator + "k) = yimean" + vbCrLf + _
						"end for" + vbCrLf + _
						"nu=size(xsortu)" + vbCrLf + _
						"cell(wc+5" + Separator + "1)=1" + vbCrLf + _
						"for k1 = 1 to nu do" + vbCrLf + _
						"xtemp=xsortu[k1]" + vbCrLf + _
						"if xtemp = uniquex[cell(wc+5" + Separator + "1)] then" + vbCrLf + _
						"cell(wc+6" + Separator + "k1)=cell(wc+4" + Separator + "cell(wc+5" + Separator + "1))" + vbCrLf + _
						"cell(wc+4" + Separator + "1)=cell(wc+5" + Separator + "1) + 1" + vbCrLf + _
						"else" + vbCrLf + _
						"cell(wc+6" + Separator + "k1)=ysortu[k1]" + vbCrLf + _
						"end if" + vbCrLf + _
						"end for" + vbCrLf + _
						"ysortu1=col(wc+6)" + vbCrLf + _
						"dydx={" + Chr(34) + Chr(34) + "" + Separator + "(diff(ysortu1)/diff(xsortu))[data(2" + Separator + "nu)]}" + vbCrLf + _
						"col(co)=xsortu" + vbCrLf + _
						"navg1=If(navg>=nu" + Separator + "nu-1" + Separator + "navg) 'navg must < nu " + vbCrLf + _
						"cell(co+2" + Separator + "1)=navg1" + vbCrLf + _
						"For i=1 To nu-navg1 Do" + vbCrLf + _
						"   n1=i+1" + vbCrLf + _
						"   n2=i+navg" + vbCrLf + _
						"   range=data(n1" + Separator + "n2)" + vbCrLf + _
						"   dydxseg=dydx[range]" + vbCrLf + _
						"   dydxavg=mean(dydxseg)" + vbCrLf + _
						"   cell(co+1" + Separator + "i+Int(navg/2))=dydxavg" + vbCrLf + _
						"End For" + vbCrLf + _
						"For j = 1 To size(col(wc)) Do" + vbCrLf + _
						"   For j1 = 1 to 7 do" + vbCrLf + _
						"      cell(wc+j1-1" + Separator + "j)=" + Chr(34) + Chr(34) + vbCrLf + _
						"   End For" + vbCrLf + _
						"End For"
'	SPTransform.RunEditor 'Debug the transform
	SPTransform.Execute
	SPTransform.Close(False)

'check for navg < nu
Dim nu
nu=WorksheetTable.Cell(CLng(dlg.Results)+1,0)
If dlg.Navg > nu Then MsgBox "Your running average length was adjusted to be less than your x data size",vbInformation,"SigmaPlot"

'Add Column Titles to Data and Results
'	Dim Data_1 As String
'	Dim Data_2 As String
	Dim Results_1 As String
	Dim Results_2 As String
'	Data_1 = "x data"
'	Data_2 = "y data"
	Results_1 = "sorted unique x"
	Results_2 = "dy/dx"
'	On Error GoTo NextTitle
'	WorksheetTable.NamedRanges.Add(Data_1, ListedColumns(dlg.First)-1,0,1,-1, True)
'	NextTitle:
'	On Error GoTo NextTitle2
'	WorksheetTable.NamedRanges.Add(Data_2, ListedColumns(dlg.First),0,1,-1, True)
	NextTitle2:
	Dim Marker
	Marker = CLng(dlg.Results)
	WorksheetTable.NamedRanges.Add(Results_1,Marker-1,0,1,-1, True)
	WorksheetTable.NamedRanges.Add(Results_2,Marker,0,1,-1, True)
	WorksheetTable.NamedRanges.Add("averaging length",Marker+1,0,1,-1, True)


'Plot the graphs
Dim PlottedColumns()As Variant
Dim SPPage As Object

If dlg.PlotOriginal = 1 Then
	Set SPPage = ActiveDocument.NotebookItems.Add(2)  'Creates graph page
	SPPage.Name = SPPage.Name + ": Original Data"
	ReDim PlottedColumns(1)
	PlottedColumns(0) = ListedColumns(dlg.First)
	PlottedColumns(1) = ListedColumns(dlg.First+1)
	SPPage.CreateWizardGraph("Line Plot","Simple Straight Line","XY Pair" ,PlottedColumns)
	SPPage.GraphPages(0).Graphs(0).Name = "Original Data"
End If

If dlg.PlotResults = 1 Then
	Set SPPage = ActiveDocument.NotebookItems.Add(2)  'Creates graph page
	SPPage.Name = SPPage.Name + ": 1st Deriviative"
	ReDim PlottedColumns(1)
	PlottedColumns(0) = CLng(dlg.Results)-1
	PlottedColumns(1) = CLng(dlg.Results)
	SPPage.CreateWizardGraph("Line Plot","Simple Straight Line","XY Pair",PlottedColumns)
	SPPage.GraphPages(0).Graphs(0).Name = "dx/dy"
End If

GoTo Finish

ErrorMsg:
		If ErrorCheck = 0 Then
			HelpMsgBox HelpID, "A worksheet with one x data column and one y data column must be open", vbExclamation,"SigmaPlot"
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