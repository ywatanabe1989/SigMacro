Option Explicit
Public Const ObjectHelp = Path + "\SPW.CHM"
Public HelpID As Variant
Dim qnan As String
Dim sd As String
Dim time_column
Dim date_column
Function FlagOn(flag As Long)
  FlagOn = flag Or FLAG_SET_BIT ' Use to set option flag bits on, leaving others unchanged
End Function
Function FlagOff(flag As Long)
  FlagOff = flag Or FLAG_CLEAR_BIT ' Use to set option flag bits off, leaving others unchanged
End Function
Sub Main
'Macro created by Mark Antro
'Updated 11/28/2001 by John Kuo
HelpID = 80502			' Help ID number for this topic in SPW.HLP

sd = decimalsymbol
qnan = "-1" + sd + "#QNAN"
Dim i As Long
Dim index As Long
Dim listindex As Long
Dim UsedColumns$()
Dim ListedColumns$()
Dim x_min As Long
Dim x_max As Long

Dim ErrorCheck As Integer
ErrorCheck = 0
On Error GoTo NoData

Dim CurrentWorksheet
CurrentWorksheet = ActiveDocument.CurrentDataItem.Name
ActiveDocument.NotebookItems(CurrentWorksheet).Open 'Opens/select default worksheet and sets focus
'Place Worksheet into Overwrite mode
ActiveDocument.NotebookItems(CurrentWorksheet).InsertionMode = False

'Determine the data range and define the first empty column
Dim wskt As Object
Set wskt = ActiveDocument.NotebookItems(CurrentWorksheet).DataTable
Dim LastColumn As Long
Dim LastRow As Long
LastColumn = 0
LastRow = 0
wskt.GetMaxUsedSize(LastColumn,LastRow)

listindex = 0		'number of emtries in arrays, equal to the number of valid columns
For index = 0 To LastColumn - 1
	ReDim Preserve UsedColumns(listindex)
	ReDim Preserve ListedColumns(listindex)
	If wskt.Cell(index,-1) = qnan Then
			UsedColumns$(listindex) = "Column " + CStr(index + 1)
			ListedColumns$(listindex) = CStr(index + 1)
	Else
			UsedColumns$(listindex) = wskt.Cell(index,-1)
			ListedColumns$(listindex) = CStr(index + 1)
	End If
			listindex = listindex + 1
Next index

If UsedColumns(1) = Empty Then GoTo NoData

MacroDialog:
	Begin Dialog UserDialog 310,203,"Merge Columns",.join_dialog ' %GRID:10,7,1,1
		Text 10,7,100,14,"&1st column",.Text2
		ListBox 10,21,140,91,UsedColumns(),.date_col
		Text 160,7,100,14,"&2nd column",.Text1
		ListBox 160,21,140,91,UsedColumns(),.time_col
		OKButton 110,175,90,21
		CancelButton 210,175,90,21
		GroupBox 10,112,290,56,"",.GroupBox1
		Text 20,122,270,35,"Combines two columns (such as dates and times) to a single column in the first empty column",.Text3
		PushButton 10,175,90,21,"Help",.PushButton1
	End Dialog
	Dim dlg As UserDialog
	dlg.time_col=1		'sets date column to second in list ie different to time column

Select Case Dialog(dlg)
	Case 0 'Handles Cancel button
		GoTo endd
	Case 1 'Handles Help button
		Help(ObjectHelp,HelpID)
		GoTo MacroDialog
	End Select

'make selected columns reflect the actual columns
date_column = Val(ListedColumns(date_column))-1

time_column = Val(ListedColumns(time_column))-1

'MsgBox CStr(time_column) + " / " + CStr(date_column)

index = 0		'counter
Dim date_cell, time_cell As Variant
Do
	date_cell =	CStr(wskt.Cell((date_column),index))
		If date_cell = qnan Then date_cell = ""
	time_cell= CStr(wskt.Cell(time_column,index))
		If time_cell = qnan Then time_cell = ""
	wskt.Cell(LastColumn,index)= CStr(date_cell) + " " + CStr(time_cell)
	index = index + 1
Loop Until index = LastRow
GoTo endd:

NoData:
HelpMsgBox HelpID, "You must have a worksheet with at least two columns of data open",vbExclamation,"No Open Worksheet"
endd:

End Sub
Rem See DialogFunc help topic for more information.
Private Function join_dialog(DlgItem$, Action%, SuppValue&) As Boolean
	Dim close_dlg As Boolean
	close_dlg = True		'indicates whether to close dialog box (true keep box open)
	Select Case Action%
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		Select Case DlgItem$
        Case "PushButton1"			' Help button
			Help(ObjectHelp,HelpID)
        Case "OK"
		'Check two different columns selected
		If DlgValue("time_col")=DlgValue("date_col") Then
			MsgBox "You must select two different columns to run this macro."
			GoTo end_case
		End If
		time_column = DlgValue("time_col")	'sets columns for user selections
		date_column = DlgValue("date_col")

		close_dlg = False
		Case "Cancel"
		close_dlg = False
		Case Else

		End Select
		end_case:
		join_dialog = close_dlg ' Prevent button press from closing the dialog box
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem select_options = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
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