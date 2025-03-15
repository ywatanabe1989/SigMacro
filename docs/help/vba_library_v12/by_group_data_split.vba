Option Explicit
Dim ColIndFrom
Dim ColIndTo
Dim ColInd
Dim LastRow As Long
Dim Separator$
Public Const ObjectHelp = Path + "\SPW.CHM"
Public HelpID As Variant
Sub Main
Separator= ListSeparator
'**********************************************************************
'Macro by Alex Ayshel 12/2000
'Split column by group (group is identified as the same cell value) and
'put data related to the group to individual columns. Optionally, data
'within the column could be sorted in ascending order. 
'
' "F-Test Comparison of Curves' written by Mohammad Younus and 
' modified and upgraded by John Kuo was used as a prototype
'
' The empty_col Public Function is written by John Kuo
'
' 2/05/2001 A.Ayshel, verified code for the latest changes
' 2/22/2001 A.Ayshel, fixed bug related to empty cell in group column
' 6/18/2001 F.Cabasa, help files added to error messages
'***********************************************************************

Dim CurrentWorksheet

HelpID = 70202			' Help ID number for this topic in SPW.CHM
Dim iFlagChkBox As Integer        '0 = empty; 1 = there is an entry
Dim iFlagOutputCol As Integer
Dim iFlagErr As Integer

Dim vWField As Variant

Dim iIsNum As Integer
Dim vFirstEmptyCol 
Dim vOutCol
Dim Column
Dim Column_end                     ' last row
Dim ColContents
Dim ColTitle

On Error GoTo NoData
CurrentWorksheet = ActiveDocument.CurrentDataItem.Name
'********************************************************************

'Opens/select default worksheet and sets focus
ActiveDocument.NotebookItems(CurrentWorksheet).Open 
'Place Worksheet into Overwrite mode
ActiveDocument.NotebookItems(CurrentWorksheet).InsertionMode = False

'*********************************************************************
'Determine the data range and define the First Empty column
'*********************************************************************
Dim WorksheetTable As Object
Set WorksheetTable = ActiveDocument.NotebookItems(CurrentWorksheet).DataTable
Dim LastColumn As Long
'Dim LastRow As Long

LastColumn = 0
LastRow = 0

WorksheetTable.GetMaxUsedSize(LastColumn,LastRow)
vFirstEmptyCol = LastColumn + 1

'*********************************************************************
'Set all flags to zero (empty state)
'*********************************************************************
iFlagChkBox = 0
iFlagOutputCol = 0
iFlagErr = 0

'***********************************************************************
' The following code originally written in F-test Comparison of Curves
' macro. Its modified for the purpose of this macro.
' Sort through columns and create list of columns with values in row 1
'***********************************************************************
On Error GoTo EmptyWorksheet
Dim Index, UsedColumns$()
Dim ListIndex, ListedColumns()
Dim ColTitles()
Dim ColId
Dim DatCol 
Dim GrCol
Dim NonEmptyCnt As Integer

ListIndex = 0
NonEmptyCnt = 0

ReDim UsedColumns(LastColumn -1)
ReDim ListedColumns(LastColumn -1)
ReDim ColTitles(LastColumn -1)

For Index = 0 To LastColumn - 1
	ColContents = empty_col(Index, LastRow) 
	If ColContents = True Then
		ColTitles(Index) = WorksheetTable.Cell(Index, -1)
		GoTo NextIndex
	Else
		UsedColumns$(Index) = CStr(Index + 1)
		ColTitles(Index) = WorksheetTable.Cell(Index, -1)
		If DatCol = "" Then
			DatCol = ListIndex
			GrCol  = ListIndex
		End If	
		ListedColumns(ListIndex) = CStr(Index + 1)
		ListIndex = ListIndex + 1
		NonEmptyCnt = NonEmptyCnt + 1
	End If
	NextIndex:
Next Index
On Error GoTo 0

If NonEmptyCnt < 2 Then
	HelpMsgBox 70202, "You must have at least 2 non empty columns to run this macro",vbExclamation, _
	       "Error Message"
	GoTo Finish       
End If

'**********************************************************************
'Dialog for source and results columns
'**********************************************************************
MacroDialog:

	Begin Dialog UserDialog 390,119,"By Group Data Split",.DialogFunc ' %GRID:10,7,1,1
		Text 10,14,100,21,"Data column",.Text1
		DropListBox 120,10,130,105,UsedColumns(),.DataColumn
		Text 10,42,110,21,"Group column",.Text2
		DropListBox 120,38,130,105,UsedColumns(),.GroupColumn
		Text 10,70,100,14,"Output column",.Text4
		TextBox 120,68,100,21,.txtOutputCol
		CheckBox 10,98,200,14,"Sort data within the group",.CheckBox1
		OKButton 280,10,100,21
		CancelButton 280,40,100,21
		PushButton 280,70,100,21,"Help",.HelpButton
	End Dialog
	Dim dlg As UserDialog

'*****************************************************************
'Default settings for Dialog
'*****************************************************************

If iFlagOutputCol = 0 Then
	'dlg.txtOutputCol = CStr(vFirstEmptyCol)
	dlg.txtOutputCol = "First Empty"
End If

If iFlagChkBox = 0 Then
	dlg.CheckBox1 = 0
End If

dlg.GroupColumn = 1 'set index to 1 to display 2nd text in droplist

'****************************************************************
'Displays Dialog 
'****************************************************************
Select Case Dialog(dlg)  
	Case 0 'Handles Cancel button
		    GoTo Finish
'	Case 1 'Handles Help button
'			HelpID = 70202			' Help ID number for this topic in SPW.CHM
'			Help(ObjectHelp,HelpID)
'            GoTo MacroDialog				   
End Select

'**********************************************************************
'Getting input from dialog, validating input
'Check if OutputCol is greater than LastCol
'**********************************************************************

DatCol = ListedColumns(dlg.DataColumn)
GrCol  = ListedColumns(dlg.GroupColumn)

'********************************************************************
' Check if GrCol has not less cells with data than DatCol 
'********************************************************************
Dim R As Long          ' running Row index
Dim GrCol_R As Long
Dim DatCol_R As Long

GrCol_R = 0
DatCol_R = 0

Set WorksheetTable = ActiveDocument.CurrentDataItem.DataTable
  '******************************************************************
  ' count non-empty cells in GrCol
  '******************************************************************
Column = GrCol -1
For R = 0 To LastRow - 2
	If WorksheetTable.Cell(Column,R) <> "-1.#QNAN" And WorksheetTable.Cell(Column,R) <> "-1,#QNAN" Then
	   GrCol_R = GrCol_R + 1
	End If
Next R	

  '******************************************************************
  ' count non-empty cells in DatCol
  '******************************************************************
Column = DatCol -1
For R = 0 To LastRow - 2
	If WorksheetTable.Cell(Column,R) <> "-1.#QNAN" And WorksheetTable.Cell(Column,R) <> "-1,#QNAN" Then
	   DatCol_R = DatCol_R + 1
	End If
Next R

  '******************************************************************
  ' compare GrCol and DatCol
  '******************************************************************
If GrCol_R < DatCol_R Then
   HelpMsgBox 70202, "There is an empty cell(s) in group column", vbExclamation, _
	          "Error Message"
   GoTo Finish	          
End If	          

'If dlg.txtOutputCol <> CStr(vFirstEmptyCol) Then
If dlg.txtOutputCol <> "First Empty" Then
    vWField = dlg.txtOutputCol
    iIsNum = IsNumeric(vWField)
    If iIsNum <> 0 Then               'Affirm that field is numeric
   		If LastColumn > vWField Then
   	 		iFlagErr = 1
       		MsgBox "Output Column must be Greater Than the last data column", vbExclamation, _
			"Error Message"
     		'dlg.txtOutputCol = CStr(vFirstEmptyCol)
     		dlg.txtOutputCol = "First Empty"
     	Else
	     	iFlagOutputCol = 1
	     	vOutCol = vWField
	    End If
	Else
   		iFlagErr = 1
    	MsgBox "Output Column must be numeric", vbExclamation, _
     	"Error Message"
     	'dlg.txtOutputCol = CStr(vFirstEmptyCol)
     	dlg.txtOutputCol = "First Empty"		
	End If
Else
	vOutCol = vFirstEmptyCol
End If

If dlg.CheckBox1 = 1 Then
	iFlagChkBox = 1
End If 

If iFlagErr = 1 Then
    iFlagErr = 0
    vWField = ""
	GoTo MacroDialog
End If

'**********************************************************************
' Copy group and data to temp location
'**********************************************************************
Dim vGrColInd	         ' Gr.Column as a work column
Dim vDataColInd          ' Data Column as a work column

vGrColInd   = vOutCol - 1
vDataColInd = vGrColInd + 1

Dim i As Integer

For i = 1 To 2
	If i = 1 Then
		ColIndFrom = CVar(GrCol - 1)
		ColIndTo   = vGrColInd
		CopyCol
	Else
		ColIndFrom = CVar(DatCol - 1) 
		ColIndTo   = vDataColInd
		CopyCol
	End If
Next i
	ActiveDocument.CurrentDataItem.Goto(0, (ColIndTo -1))

'*********************************************************
' Sort block of group and data by group in work columns
'*********************************************************
Dim SPTransform As Object
Dim GroupCol 
Dim DataCol 
GroupCol = vGrColInd + 1
DataCol	 = GroupCol + 1	

Set SPTransform = ActiveDocument.NotebookItems.Add(9)
SPTransform.Open
SPTransform.Text = "block("+GroupCol+Separator+"1) = sort(block("+GroupCol+Separator+"1"+Separator+DataCol+Separator+"size(col("+GroupCol+")))"+Separator+"col("+GroupCol+"))"
SPTransform.Execute
SPTransform.Close(False)

'********************************************************************
' Determine number of groups and assign temp work column in worksheet
'********************************************************************

Dim vF1				   ' work field 1	
Dim vF2	               ' work field 2
Dim lGrCnt As Long     ' group counter - total # of unique cells

Column = vOutCol - 1

Set WorksheetTable = ActiveDocument.CurrentDataItem.DataTable

'For R = 0 To LastRow -1
For R = 0 To GrCol_R
	If R = 0 Then
		vF1 = WorksheetTable.Cell(Column,R)
		vF2 = vF1
		lGrCnt = lGrCnt + 1
	Else
		vF1 = WorksheetTable.Cell(Column,R)
		If vF1 <> vF2 Then
		   vF2  = vF1
		   lGrCnt = lGrCnt + 1
		End If   
	End If
Next R

'**********************************************************************
' Move group and data to another temp location
'**********************************************************************
For i = 1 To 2
	If i = 1 Then
		ColIndFrom = vGrColInd 
		ColIndTo   = vGrColInd + lGrCnt + 1
		GroupCol   = ColIndTo	
		MoveCol
	Else
		ColIndFrom = vDataColInd
		ColIndTo   = vDataColInd + lGrCnt + 1
		DataCol	   = ColIndTo
		MoveCol
	End If
Next i
	ActiveDocument.CurrentDataItem.Goto(0, (ColIndTo -1))

'********************************************************************
' Read group and data in work area and split data by group.
' Assign column titles
'********************************************************************
Dim j As Integer	 ' flag to indicate change in group; 0 = new group
Dim RR As Long       ' Result Row - work field
Dim GrTitle
Dim GrRowCount()
Dim IGrCnt
Dim GrRowCnt

vF1 = ""
vF2 = ""
j = 0                ' group flag 0 or 1. Zero = new group
i = vOutCol -2
Index = 0

ReDim GrRowCount(lGrCnt)
GrRowCnt = 0         ' row counter while putting data into groups
IGrCnt = 0
'???????
'For R = 0 To LastRow -1
For R = 0 To GrCol_R
If j = 0 Then
    	i = i + 1
    	RR = 0
    	vF1 = WorksheetTable.Cell(GroupCol,R)
		vF2 = vF1
		If vF1 = "-1.#QNAN" Or vF1 = "-1,#QNAN" Then vF1 = "None"
		GrTitle = CStr(vF1 + " ")
		For Index = 0 To (LastColumn -1)
			If GrTitle = ColTitles(Index) Then
				GrTitle = ColTitles(Index) + " "
			End If
		Next Index	
		WorksheetTable.Cell(i, -1) = GrTitle
	    WorksheetTable.Cell(i,RR) = WorksheetTable.Cell(DataCol,R)
		GrRowCnt = GrRowCnt + 1
		GrRowCount(IGrCnt) = GrRowCnt
		j = 1
	Else
	    vF1 = WorksheetTable.Cell(GroupCol,R)
		If vF1 = vF2 Then
		   RR = RR + 1
		   WorksheetTable.Cell(i,RR) = WorksheetTable.Cell(DataCol,R)
		   GrRowCnt = GrRowCnt + 1
		   GrRowCount(IGrCnt) = GrRowCnt
		Else
		   j = 0
		   R = R -1
		   GrRowCnt = 0
		   IGrCnt = IGrCnt + 1
	    End If	
	End If
Next R

'*********************************************************
' Delete data in work area
'*********************************************************
' Delete Group Column
ColInd = ColIndTo - 1
ClearCol
' Delete Data Column
ColInd = ColInd + 1
ClearCol
ActiveDocument.CurrentDataItem.Goto(0, (vOutCol -1))

'*********************************************************
' Sort data in ascending order in group columns, if sort 
' option is selected
'*********************************************************
If iFlagChkBox = 1 Then
	GroupCol = vOutCol 
	IGrCnt = 0
	Set SPTransform = ActiveDocument.NotebookItems.Add(9)
	SPTransform.Open
	For i = GroupCol To (GroupCol + lGrCnt - 1)
		If GrRowCount(IGrCnt) <> 1 Then
    		SPTransform.Text = "col(" +GroupCol+ ") = sort(col(" +GroupCol+ "))"
	    	SPTransform.Execute
	    End If
		GroupCol = GroupCol + 1
		IGrCnt = IGrCnt + 1
	Next i	
	SPTransform.Close(False)
End If
GoTo Finish

'********************************************************************
EmptyWorksheet:
   HelpMsgBox 70202,"You must have data in your worksheet", vbExclamation, _
	"Error Message"
	GoTo Finish

'********************************************************************
NoData:
   HelpMsgBox 70202, "You must have a worksheet open", vbExclamation, _
	"Error Message"
	GoTo Finish
'********************************************************************
	
Finish:
End Sub


'******************************************************************
' The following Public Function is written by John Kuo
' Determines if a column is empty
'******************************************************************
Public Function empty_col(Column As Variant, Column_end As Variant)
	Dim WorksheetTable As Object
	Set WorksheetTable = ActiveDocument.CurrentDataItem.DataTable
	Dim i As Long
	Dim empty_cell As Boolean
	
	For i = 0 To Column_end
		If WorksheetTable.Cell(Column,i) = "-1.#QNAN" Or WorksheetTable.Cell(Column,i) = "-1,#QNAN" Then empty_cell = True
		If WorksheetTable.Cell(Column,i) <> "-1.#QNAN" And WorksheetTable.Cell(Column,i) <> "-1,#QNAN" Then GoTo NotEmpty
	Next i
	empty_col = empty_cell
	GoTo EmptyCol:
	NotEmpty:	
	empty_col = False
	EmptyCol:
End Function
'******************************************************************
' Copy Entire Column
' Dim ColIndFrom   -  declare in Option Explicit
' Dim ColIndTo     -  declare in Option Explicit
'******************************************************************
Public Sub CopyCol()
	ActiveDocument.CurrentDataItem.Open
	Dim CopyFrom()
	ReDim CopyFrom(3)
	CopyFrom(0) = ColIndFrom             ' Col index
	CopyFrom(1) = 0         		     ' Row index - start
	CopyFrom(2) = ColIndFrom             ' Col index
	CopyFrom(3) = LastRow     			 ' Row index - end
	ActiveDocument.CurrentDataItem.SelectionExtent = CopyFrom
	ActiveDocument.CurrentDataItem.Copy

	Dim CopyTo()

	ReDim CopyTo(3)
	CopyTo(0) = ColIndTo             ' Col index
	CopyTo(1) = 0         		     ' Row index - start
	CopyTo(2) = ColIndTo             ' Col index 
	CopyTo(3) = LastRow     		 ' Row index - end
	ReDim CopyTo(3)
	ActiveDocument.CurrentDataItem.SelectionExtent(CopyTo)
	ActiveDocument.CurrentDataItem.Goto(0, ColIndTo)
	ActiveDocument.CurrentDataItem.Paste

End Sub
'******************************************************************
' Move Entire Column
' Dim ColIndFrom   -  declare in Option Explicit
' Dim ColIndTo     -  declare in Option Explicit
'******************************************************************
Public Sub MoveCol()
	ActiveDocument.CurrentDataItem.Open
	Dim MoveFrom()
	ReDim MoveFrom(3)
	MoveFrom(0) = ColIndFrom             ' Col index
	MoveFrom(1) = 0         		     ' Row index - start
	MoveFrom(2) = ColIndFrom             ' Col index 
	MoveFrom(3) = LastRow     			 ' Row index - end
	ActiveDocument.CurrentDataItem.SelectionExtent = MoveFrom
	ActiveDocument.CurrentItem.Copy
	ActiveDocument.CurrentItem.Clear

	Dim CopyTo()

	ReDim CopyTo(3)
	CopyTo(0) = ColIndTo             ' Col index
	CopyTo(1) = 0         		     ' Row index - start
	CopyTo(2) = ColIndTo             ' Col index 
	CopyTo(3) = LastRow     		 ' Row index - end
	ReDim CopyTo(3)
	ActiveDocument.CurrentDataItem.SelectionExtent(CopyTo)
	ActiveDocument.CurrentDataItem.Goto(0, ColIndTo)
	ActiveDocument.CurrentDataItem.Paste

End Sub
'******************************************************************
' Clear Entire Column
' Dim ColInd   -  declare in Option Explicit
'******************************************************************
Public Sub ClearCol()
	ActiveDocument.CurrentDataItem.Open
	Dim ClearData()
	ReDim ClearData(3)
	ClearData(0) = ColInd             ' Col index
	ClearData(1) = 0       		     ' Row index - start
	ClearData(2) = ColInd             ' Col index 
	ClearData(3) = LastRow       	 ' Row index - end
	ActiveDocument.CurrentDataItem.SelectionExtent = ClearData
	ActiveDocument.CurrentItem.Copy
	ActiveDocument.CurrentItem.Clear
End Sub
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
		Case "HelpButton"
			Help(ObjectHelp,HelpID)
			DialogFunc = True 'do not exit the dialog
        End Select
	End Select
End Function