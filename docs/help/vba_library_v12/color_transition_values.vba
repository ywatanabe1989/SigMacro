Option Explicit
Public Const ObjectHelp = Path + "\SPW.CHM"
Public HelpID As Variant
Dim Separator$
Function FlagOn(flag As Long)
    FlagOn = flag Or FLAG_SET_BIT
End Function
Function FlagOff(flag As Long)
    FlagOff = flag Or FLAG_CLEAR_BIT
End Function
Public R1, R2, B1, B2, G1, G2

Sub Main
Separator = ListSeparator
'Macro by M Younus on 10/28/98
'Modified on 1/13/99
'Updated 12/22/99 John Kuo

'This macro assigns colors from a gradient between two selected 
'colors to all datapoints in a selected plot, using a data column 
'as an index.  

On Error GoTo NoGraph

HelpID = 60202			' Help ID number for this topic in SPW.CHM
Dim Index, SPPage, SPGraph, NumberPlots, PlotList$()
Set SPPage = ActiveDocument.CurrentPageItem
SPPage.Open
Set SPGraph = SPPage.GraphPages(0).CurrentPageObject(GPT_GRAPH)
NumberPlots = SPGraph.Plots.Count
ReDim PlotList$(NumberPlots - 1)

For Index = 0 To NumberPlots - 1
	PlotList(Index) = SPGraph.Plots(Index).Name
Next Index

On Error GoTo NoData
Dim CurrentWorksheet
Set CurrentWorksheet = ActiveDocument.CurrentDataItem
CurrentWorksheet.Open 'Opens/select default worksheet and sets focus

'Determine the data range and define the first empty column

Dim WorksheetTable As Object
Set WorksheetTable = CurrentWorksheet.DataTable
Dim LastColumn As Long
Dim LastRow As Long
LastColumn = 0
LastRow = 0 
WorksheetTable.GetMaxUsedSize(LastColumn,LastRow)
'Place Worksheet into Overwrite mode
ActiveDocument.CurrentDataItem.InsertionMode = False

'Sort through columns and create list of columns with values in row 1
Dim UsedColumns$(), ListedColumns(), ListIndex, ColContents, ColTitle
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

Dim Colors$()
ReDim Colors$(16)

Colors(0) = "Black"
Colors(1) = "White"
Colors(2) = "Red"
Colors(3) = "Green"
Colors(4) = "Yellow"
Colors(5) = "Blue"
Colors(6) = "Pink"
Colors(7) = "Cyan"
Colors(8) = "Gray"
Colors(9) = "Dark Red"
Colors(10) = "Dark Green"
Colors(11) = "Dark Yellow"
Colors(12) = "Dark Blue"
Colors(13) = "Dark Pink"
Colors(14) = "Dark Cyan"
Colors(15) = "Dark Gray"
Colors(16) = "Custom..."

'Initial default colors
Dim DefaultColor1, DefaultColor2
DefaultColor1 = 2
DefaultColor2 = 5

'Dialog for source and results columns
MacroDialog:
	Begin Dialog UserDialog 355,214,"Color Transition Values",.CustomColors ' %GRID:10,7,1,0
		OKButton 140,185,96,19
		CancelButton 250,185,96,19
		PushButton 10,185,96,19,"Help",.PushButton1
		Text 10,5,130,14,"&Data column",.Text2
		DropListBox 10,20,140,100,UsedColumns(),.SourceCol
		Text 180,5,128,13,"Color &result column",.Text3
		TextBox 180,20,140,18,.ResultsCol
		Text 10,50,90,14,"&Start color",.Text4
		DropListBox 10,65,140,100,Colors(),.Color1
		Text 180,50,90,14,"&End color",.Text5
		DropListBox 180,65,140,100,Colors(),.Color2
		Text 10,95,131,14,"Apply colors to &plot:",.Text1
		DropListBox 10,110,140,100,PlotList(),.PlotUsed
		CheckBox 180,95,90,14,"&Fills",.Fills
		CheckBox 180,110,90,14,"Ed&ges",.Edges
		CheckBox 180,125,130,14,"S&ymbol fills",.SymbolFill
		CheckBox 180,140,142,14,"Sy&mbol edges",.SymbolEdge
		CheckBox 180,155,90,14,"&Lines",.LineColor
	End Dialog

Dim dlg As UserDialog
'Default settings
dlg.Color1 = DefaultColor1 
dlg.Color2 = DefaultColor2
dlg.SourceCol = 0
dlg.Fills = 1
dlg.Edges = 1
dlg.SymbolFill = 0
dlg.SymbolEdge = 0
dlg.LineColor = 0
If dlg.ResultsCol = "" Then dlg.ResultsCol = "First Empty"

	Select Case Dialog(dlg)  
		Case 0 'Handles Cancel button
			GoTo Finish
'		Case 1 'Handles Help button
'			HelpID = 60202			' Help ID number for this topic in SPW.CHM
'			Help(ObjectHelp,HelpID)
'			GoTo MacroDialog 
	End Select

'Preserve selection
DefaultColor1 = dlg.Color1
DefaultColor2 = dlg.Color2

'Set the RGB values for Color1
Select Case dlg.Color1
Case 0
	R1 = 0
	G1 = 0
	B1 = 0
Case 1
	R1 = 255
	G1 = 255
	B1 = 255
Case 2
	R1 = 255
	G1 = 0
	B1 = 0
Case 3
	R1 = 0
	G1 = 255
	B1 = 0
Case 4
	R1 = 255
	G1 = 255
	B1 = 0
Case 5
	R1 = 0
	G1 = 0
	B1 = 255
Case 6
	R1 = 255
	G1 = 0
	B1 = 255
Case 7
	R1 = 0
	G1 = 255
	B1 = 255
Case 8
	R1 = 192
	B1 = 192
	G1 = 192
Case 9
	R1 = 128
	B1 = 0
	G1 = 0
Case 10
	R1 = 0
	G1 = 128
	B1 = 0
Case 11
	R1 = 128
	G1 = 128
	B1 = 0
Case 12
	R1 = 0
	G1 = 0
	B1 = 128
Case 13
	R1 = 128
	G1 = 0
	B1 = 128
Case 14
	R1 = 0
	G1 = 128
	B1 = 128
Case 15
	R1 = 128
	G1 = 128
	B1 = 128
End Select

'Set the RGB values for Color2
Select Case dlg.Color2
Case 0
	R2 = 0
	G2 = 0
	B2 = 0
Case 1
	R2 = 255
	G2 = 255
	B2 = 255
Case 2
	R2 = 255
	G2 = 0
	B2 = 0
Case 3
	R2 = 0
	G2 = 255
	B2 = 0
Case 4
	R2 = 255
	G2 = 255
	B2 = 0
Case 5
	R2 = 0
	G2 = 0
	B2 = 255
Case 6
	R2 = 255
	G2 = 0
	B2 = 255
Case 7
	R2 = 0
	G2 = 255
	B2 = 255
Case 8
	R2 = 192
	B2 = 192
	G2 = 192
Case 9
	R2 = 128
	B2 = 0
	G2 = 0
Case 10
	R2 = 0
	G2 = 128
	B2 = 0
Case 11
	R2 = 128
	G2 = 128
	B2 = 0
Case 12
	R2 = 0
	G2 = 0
	B2 = 128
Case 13
	R2 = 128
	G2 = 0
	B2 = 128
Case 14
	R2 = 0
	G2 = 128
	B2 = 128
Case 15
	R2 = 128
	G2 = 128
	B2 = 128
End Select

'Parse the "First Empty" result
WorksheetTable.GetMaxUsedSize(LastColumn,LastRow) 'Re-initialize variables
If dlg.ResultsCol = "First Empty" Then
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

'Open and run Rgbcolr2.xfm transform
Dim SPTransform As Object
Set SPTransform = ActiveDocument.NotebookItems.Add(9)
SPTransform.Open
SPTransform.Text =	"d = max(col(" + CStr(ListedColumns(dlg.SourceCol)) + "))-min(col(" + CStr(ListedColumns(dlg.SourceCol)) + "))" + vbCrLf + _
					"range = if( d=0" + Separator + "1" + Separator + " d)" + vbCrLf + _
					"t = (col("+ CStr(ListedColumns(dlg.SourceCol)) + ") - min(col(" + CStr(ListedColumns(dlg.SourceCol)) + ")))/range" + vbCrLf + _
					"r = (" + CStr(R2) + "-" + CStr(R1) + ")*t+" + CStr(R1) + vbCrLf + _
					"g = (" + CStr(G2) + "-" + CStr(G1) + ")*t+" + CStr(G1) + vbCrLf + _
					"b = (" + CStr(B2) + "-" + CStr(B1) + ")*t+" + CStr(B1) + vbCrLf + _
					"col(" + CStr(dlg.ResultsCol) + ") = rgbcolor(r" + Separator + "g" + Separator + "b)" + vbCrLf
'SPTransform.RunEditor 'Debug transform code
SPTransform.Execute
SPTransform.Close(False)

'Add column titles to results
Dim Results_1 As String
Dim Results_2 As String
Results_1 = "Color"
WorksheetTable.NamedRanges.Add(Results_1,CLng(dlg.ResultsCol)-1,0,1,-1, True)

'Set fill color to color column
If dlg.Fills = 1 Then
	SPGraph.Plots(dlg.PlotUsed).SetAttribute(SDA_COLORCOL, CLng(dlg.ResultsCol) - 1)
	SPGraph.Plots(dlg.PlotUsed).SetAttribute(SDA_COLORREPEAT, 4)
End If
If dlg.Edges = 1 Then
	SPGraph.Plots(dlg.PlotUsed).SetAttribute(SDA_EDGECOLORCOL, CLng(dlg.ResultsCol) - 1)
	SPGraph.Plots(dlg.PlotUsed).SetAttribute(SDA_EDGECOLORREPEAT, 4)
End If
If dlg.SymbolFill = 1 Then
	SPGraph.Plots(dlg.PlotUsed).SetAttribute(SSA_COLORCOL, CLng(dlg.ResultsCol) - 1)
	SPGraph.Plots(dlg.PlotUsed).SetAttribute(SSA_COLORREPEAT, 4)
End If
If dlg.SymbolEdge = 1 Then
	SPGraph.Plots(dlg.PlotUsed).SetAttribute(SSA_EDGECOLORCOL, CLng(dlg.ResultsCol) - 1)
	SPGraph.Plots(dlg.PlotUsed).SetAttribute(SSA_EDGECOLORREPEAT, 4)
End If
If dlg.LineColor = 1 Then
	SPGraph.Plots(dlg.PlotUsed).SetAttribute(SEA_COLORCOL, CLng(dlg.ResultsCol) - 1)
	SPGraph.Plots(dlg.PlotUsed).SetAttribute(SEA_COLORREPEAT, 4)
End If

'Clear the legend
SPGraph.SetAttribute(SGA_FLAGS, FlagOff(SGA_FLAG_AUTOLEGENDSHOW))
SPPage.Open
GoTo Finish

NoGraph:
	HelpMsgBox 60202, "You must have a graph page with at least one graph open.",vbExclamation,"No Open Page"

NoData:
	HelpMsgBox 60202, "You must have a worksheet open.",vbExclamation,"No Open Worksheet"

Finish:
End Sub
Public Function empty_col(Column As Variant, column_end As Variant)
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
Private Function CustomColors(DlgItem$, Action%, SuppValue%) As Boolean
Dim CurrentFillColor, CurrentColor
	Select Case Action%
	Case 2
			Select Case DlgItem$
			Case "PushButton1"
				Help(ObjectHelp,HelpID)
				CustomColors = True 'do not exit the dialog
			Case "Cancel"
				End
			Case "Color1"
				CurrentColor = 1
				If SuppValue = 16 Then CurrentFillColor = 1 
				If SuppValue = 16 Then GoTo CustomColor
			Case "Color2"
				CurrentColor = 2
				If SuppValue = 16 Then CurrentFillColor = 2
				If SuppValue = 16 Then GoTo CustomColor
			End Select	
	End Select

GoTo Finish

CustomColor:

	Begin Dialog UserDialog 260,112,"Custom Fill Colors" ' %GRID:10,7,1,1
		GroupBox 10,7,140,91,"Color",.GroupBox1
		Text 20,23,20,14,"&R",.Text2
		TextBox 60,21,70,21,.R
		Text 20,49,20,14,"&G",.Text3
		TextBox 60,47,70,18,.G
		Text 20,73,20,14,"&B",.Text4
		TextBox 60,70,70,21,.B
		OKButton 160,14,90,21
		CancelButton 160,49,90,21
	End Dialog
	Dim Custom_dlg As UserDialog

	'Default settings
	Custom_dlg.R = "255"
	Custom_dlg.G = "255"
	Custom_dlg.B = "128"

	Select Case Dialog(Custom_dlg)  
	Case 0 'Handles Cancel button
		GoTo Finish
	End Select
	
Select Case CurrentColor
Case 1 
	R1 = Custom_dlg.R
	G1 = Custom_dlg.G
	B1 = Custom_dlg.B
Case 2 
	R2 = Custom_dlg.R
	G2 = Custom_dlg.G
	B2 = Custom_dlg.B
End Select

Finish:
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