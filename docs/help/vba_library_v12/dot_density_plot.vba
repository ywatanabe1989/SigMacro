Option Explicit
'Public Const ObjectHelp = Path + "\SPW.HLP"
'Public HelpID As Variant
Public Const LastRow = 31999999
Public Const strQnan = "-1" & decimalsymbol & "#QNAN"		' value for empty cell
Public Const strInf = "1" & decimalsymbol & "#INF"				'positive inf (neg = "-" & inf)

' Prevent textbox editing
Const EM_SETREADONLY = &hCF

Declare Sub SendDlgItemMessage Lib "user32" _
(ByVal hWnd As Long, _
ByVal id As Long, _
ByVal uMsg As Long, _
ByVal wParam As Long, _
ByVal lParam As Long)

Dim iIndex As Integer				' Counter variable
Dim iColumns As Integer
'Dim blnColcontents As Boolean
Dim lngLastColumn As Long
Dim lngLastRow As Long
Dim blnNumeric As Boolean			' Flags column type => Numeric = True
Dim ablnData_columns() As Boolean
Dim alngData_columns() As Long		' List of data columns - numbers
Dim astrData_columns() As String	' List of data columns - column titles
Dim objCnb As Object				' CurrentNoteBook object
Dim objCdi As Object				' CurrentDataItem object
Dim objCdt As Object				' CurrentDataTable object
Dim objGraph_page As Object			' Graph Page
Dim objCdi_name As String			' Name of CurrentDataItem object
Dim dblPrecision As Double
Dim dblMean As Double				' Mean
Dim strGname As String
Dim iNo As Integer					' Size of method 1 column
Dim aMethod1(0) As String
Dim aMethod2(0) As String
Dim strMethod1 As String
Dim strMethod2 As String
Dim lngMethod1 As Long
Dim lngMethod2 As Long
' Graph variables
Dim iGraph As Integer
Dim objPlot As Object
Dim strGraph_title As String
Dim lngXmax
Dim lngXmin
Dim ColumnsPerPlot()
Dim PlotColumnCountArray(0)
Dim blnNewPage As Boolean				' Are new graphs created on a new page (TRUE) or not (FALSE)
Dim blnConf As Boolean					' Are confidence lines plotted (TRUE) or not (FALSE)
Dim blnConf99 As Boolean				' Are confidence lines 99% (TRUE) or not - ie 95% (FALSE)
Dim blnCI99 As Boolean					' Are confidence limits used in stats 99% (TRUE) or not - ie 95% (FALSE)
Dim dblStddev As Double
Dim strDecimals As String
Dim sL$
Dim sD$
Function FlagOn(flag As Long)
  FlagOn = flag Or FLAG_SET_BIT ' Use to set option flag bits on, leaving others unchanged
End Function
Function FlagOff(flag As Long)
  FlagOff = flag Or FLAG_CLEAR_BIT ' Use to set option flag bits off, leaving others unchanged
End Function
Sub Main
' Written by Antro
' Made changes to several bits and pieces 02/18/03
' Modified by Dick Mitchell 8/27/09
' version 1.03

sL = ListSeparator  'internationalize list separator
sD = DecimalSymbol  'internationalize decimal symbol

Dim strErr_Section As String
'HelpID = 80503			' Help ID number for this topic in SPW.HLP

' Default program values
strDecimals = "#.0000"
dblPrecision = 1.96
blnConf = True
blnNewPage = False
'*********************************************
' Set Notebook and Worksheet Variables 
'*********************************************
Set objCnb=ActiveDocument
On Error GoTo no_open_wsk
Set objCdi = objCnb.CurrentDataItem
objCdi.Open							'Opens worksheet when macro runs
objCdi_name = objCdi.Name			'Name of active Worksheet
Set objCdt= objCdi.DataTable

RESUME_FROM_NO_OPEN_WSK:
' Find last column and row in worksheet
lngLastColumn = 0
lngLastRow = 0
objCdt.GetMaxUsedSize(lngLastColumn,lngLastRow)
lngLastColumn = lngLastColumn - 1
lngLastRow = lngLastRow - 1

'***************
' Check for Data
'***************
If lngLastColumn < 0 Or lngLastRow < 0 Then
	MsgBox "There is no data in this worksheet. You must" & vbCrLf & "have two columns of data to compare against" & vbCrLf & "each each other, in order to run this program.",16, "Data Worksheet Error Message"
	Exit All
ElseIf lngLastColumn = 0 Then		' Only column present
	MsgBox "There is not enough data in this worksheet. You" & vbCrLf & "must have more than one column of data present.",16, "Data Worksheet Error Message"
	Exit All
End If

'Sort through columns and create list of columns with values in row 1
On Error GoTo ERROR_MESSAGE
strErr_Section = "finding columns in the worksheet containing data."
ReDim ablnData_columns(lngLastColumn)
Dim lnglistindex As Long
lnglistindex = 0

For iIndex = 0 To lngLastColumn
	If empty_col(iIndex, lngLastRow) = False Then   'If the first cell is not empty
		ablnData_columns(lnglistindex) = blnNumeric
		ReDim Preserve astrData_columns(lnglistindex + 1)		' Column titles
		ReDim Preserve alngData_columns(lnglistindex + 1)		' Column numbers
			Select Case CStr(objCdt.Cell(iIndex,-1))
				Case strQnan
					astrData_columns(lnglistindex) = "Column " + CStr(iIndex + 1)
					alngData_columns(lnglistindex) = CStr(iIndex)
					lnglistindex = lnglistindex + 1
				Case Else
					astrData_columns(lnglistindex) = CStr(objCdt.Cell(iIndex,-1))							'If title is present use title
					alngData_columns(lnglistindex) = CStr(iIndex)
					lnglistindex = lnglistindex + 1
			End Select
	End If
Next iIndex

main_dialog:
On Error GoTo ERROR_MESSAGE
strErr_Section = "creating the main dialog."

	Begin Dialog UserDialog 470,378,"Bland - Altman Graph Settings",.BA_dialog ' %GRID:10,7,1,1
		GroupBox 10,7,450,168,"Data",.GroupBox1
		GroupBox 10,182,310,112,"Plots",.GroupBox2
		GroupBox 10,301,450,66,"Information",.GroupBox3
		ListBox 30,49,150,112,astrData_columns(),.listColumns
		PushButton 200,70,60,21,">>",.pbSelect1
		PushButton 200,133,60,21,">>",.pbSelect2
		Text 30,28,150,14,"Available Columns",.Text1,2
		Text 290,28,150,14,"Selected Columns",.Text4,2
		Text 290,56,150,14,"Method One",.Text2
		Text 290,112,150,14,"Method Two",.Text3
		OKButton 350,189,90,21
		CancelButton 350,217,90,21
		PushButton 350,273,90,21,"Help",.pbHelp
		CheckBox 30,203,270,14,"Plot Method 1 versus Method 2",.CheckBox1
		CheckBox 80,224,210,14,"Add Linear Regression Line",.CheckBox2
		CheckBox 80,245,200,14,"Add Line of Equality",.CheckBox3
		CheckBox 30,266,270,14,"Plot Bland - Altman Graph",.CheckBox4
		Text 20,315,420,49,"Text8",.txtInformation
		TextBox 290,70,150,21,.cbMethod1
		TextBox 290,133,150,21,.cbMethod2
		PushButton 350,245,90,21,"Options",.pbOptions
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg

Dim objTransform As TransformItem
Set objTransform = open_xfm

If check_data(objTransform) = True Then
	' Must have an error, select new data from dialog
	GoTo main_dialog
End If

If dlg.CheckBox4 = 0 And dlg.CheckBox1 = 0 Then
	Set objTransform = open_xfm
	analysis(objTransform)
	Exit All 		' Neither graph type selected, so we exit here after doing analysis above
End If

' Start Graphing Section
On Error GoTo ERROR_MESSAGE
strErr_Section = "starting to create the graph."

If dlg.checkbox1 = 1 Then		' plot Method 1 versus Method 2
	If blnNewPage = True Then
		iGraph = 0				' Number of graphs plotted so far - needed for graph title in create_graph routine
	Else
		iGraph = 1
	End If
	add_graph_page
	strGraph_title = "Method 1 versus Method 2"
	ReDim ColumnsPerPlot(2,1)
	ColumnsPerPlot(0,0) = lngMethod1		' X
	ColumnsPerPlot(1,0) = 0					' first row
	ColumnsPerPlot(2,0) = LastRow			' last row
	ColumnsPerPlot(0,1) = lngMethod2 		' Y
	ColumnsPerPlot(1,1) = 0					' first row
	ColumnsPerPlot(2,1) = LastRow			' last row
	PlotColumnCountArray(0) = 2
	create_graph

	strErr_Section = "adding the regression line to the graph."

	If dlg.checkbox2 = 1 Then 	' Add Regression Line
		objPlot.SetAttribute(SLA_REGRORDER, 1)
		objPlot.SetAttribute(SLA_REGROPTIONS, &H00016018&)		' draws regression line for all data
		Dim objRegressionLine 
		Set objRegressionLine = objPlot.Plots(0).Functions(1) 
		objRegressionLine.SetAttribute(SEA_THICKNESS, 30)

		If blnConf = True Then
			If blnConf99 = False Then
				objPlot.SetAttribute( SLA_REGROPTIONS, &H00016038&)	' draws 95% confidence lines
			Else
				objPlot.SetAttribute( SLA_REGROPTIONS, &H00016138&)	' draws 99% confidence lines
			End If
			' modify them to blue, increased thickness and medium dashed lines
			objPlot.Plots(0).Functions(SLA_FUNC_CONF1).SetAttribute(SEA_LINETYPE,SEA_LINE_MEDD)
			objPlot.Plots(0).Functions(SLA_FUNC_CONF1).Color = RGB_BLUE
			objPlot.Plots(0).Functions(SLA_FUNC_CONF1).SetAttribute(SEA_THICKNESS, 15)
			objPlot.Plots(0).Functions(SLA_FUNC_CONF2).SetAttribute(SEA_LINETYPE,SEA_LINE_MEDD)
			objPlot.Plots(0).Functions(SLA_FUNC_CONF2).Color = RGB_BLUE
			objPlot.Plots(0).Functions(SLA_FUNC_CONF2).SetAttribute(SEA_THICKNESS, 15)
		End If
	End If
	
	' Find min & max X & Y axis range values
	Dim lngYmin, lngYmax
	objPlot.Axes(1).GetAttribute(SAA_FROMVAL,lngYmin)
	objPlot.Axes(1).GetAttribute(SAA_TOVAL, lngYmax)		
	' Now determine which values to use
	' For start, we need smallest value of XMin and YMin
	If lngXmin < lngYmin Then
		lngYmin = lngXmin
	End If
	' For end, we need largest value of XMax and YMax
	If lngXmax > lngYmax Then
		lngYmax = lngXmax
	End If
	
	'Set the X and Y axis ranges to be identical
	objPlot.Axes(0).SetAttribute(SAA_OPTIONS,SAA_FLAG_AUTORANGE Or FLAG_CLEAR_BIT)
	objPlot.Axes(0).SetAttribute(SAA_FROMVAL,lngYmin)
	objPlot.Axes(0).SetAttribute(SAA_TOVAL, lngYmax)
	objPlot.Axes(1).SetAttribute(SAA_OPTIONS,SAA_FLAG_AUTORANGE Or FLAG_CLEAR_BIT)
	objPlot.Axes(1).SetAttribute(SAA_FROMVAL,lngYmin)
	objPlot.Axes(1).SetAttribute(SAA_TOVAL, lngYmax)		

	If dlg.checkbox3 = 1 Then 	' Add Line of Equality
		strErr_Section = "adding the line of equality to the graph."		
		
		' Enter data in to next empty columns
		objCdt.GetMaxUsedSize(lngLastColumn,lngLastRow)
		objCdt.Cell(lngLastColumn, 0) = lngXmin
'		objCdt.Cell(lngLastColumn, 0) = lngXmin
		objCdt.Cell(lngLastColumn, 1) = lngXmax
'		objCdt.Cell(lngLastColumn, 1) = lngXmax

		ReDim ColumnsPerPlot(2,1)
		ColumnsPerPlot(0,0) = lngLastColumn	' X
		ColumnsPerPlot(1,0) = 0				' first row
		ColumnsPerPlot(2,0) = LastRow		' last row
		ColumnsPerPlot(0,1) = lngLastColumn	' Y
		ColumnsPerPlot(1,1) = 0				' first row
		ColumnsPerPlot(2,1) = LastRow		' last row
		PlotColumnCountArray(0) = 2
		objGraph_page.AddWizardPlot("Line Plot", "Simple Straight Line", "XY Pair", ColumnsPerPlot, PlotColumnCountArray)
		
		'Set the line color and thickness.  Send it to back
		Set objPlot = objCnb.NotebookItems(strGname).GraphPages(0).Graphs(0).Plots(1)
		objPlot.Line.SetAttribute(SEA_COLOR, RGB_RED)
		objPlot.Line.SetAttribute(SEA_THICKNESS, 20)
		objPlot.SelectObject
		objPlot.SetAttribute(SPA_SENDTOBACK, 0)
		
	End If
End If					' End of Graph One

If dlg.checkbox4 = 0 Then		' plot standard BA graph? 0 = false, 1 = true
	Exit All					' Only Plot Methods against each other, so exit here
End If

' **************************************
' Start BA Graph, so need to do analysis
' **************************************
On Error GoTo ERROR_MESSAGE
strErr_Section = "starting the Bland-Altman graph."

Set objTransform = open_xfm
analysis(objTransform)

' Plot BA Graph
If dlg.checkbox1 = 0 Or blnNewPage = True Then		' If we haven't plotted methods add graph page
	add_graph_page 
End If
strGraph_title = "Bland-Altman Graph"
' Create graph, get X axis values (min & max) and add plots for limits and symbols
ReDim ColumnsPerPlot(2,1)
ColumnsPerPlot(0,0) = lngLastColumn + 2		' X
ColumnsPerPlot(1,0) = 0					' first row
ColumnsPerPlot(2,0) = LastRow		' last row
ColumnsPerPlot(0,1) = lngLastColumn + 1 	' Y
ColumnsPerPlot(1,1) = 0					' first row
ColumnsPerPlot(2,1) = LastRow		' last row
PlotColumnCountArray(0) = 2
create_graph
objCdt.Cell(lngLastColumn + 6, 0) = lngXmin				' Get X scale for lines and labels
objCdt.Cell(lngLastColumn + 6, 1) = lngXmax
objCdt.Cell(lngLastColumn + 7, 0) = lngXmax - (((lngXmax - lngXmin)/100)*10)
objCdt.Cell(lngLastColumn + 7, 1) = objCdt.Cell(lngLastColumn + 7, 0)
objCdt.Cell(lngLastColumn + 7, 2) = objCdt.Cell(lngLastColumn + 7, 0)
Set objPlot = objCnb.NotebookItems(strGname).GraphPages(0).Graphs(iGraph).Plots(0)
resize_symbols(objPlot)

' Add Horizontal Lines
strErr_Section = "adding horizontal lines to the Bland-Altman graph."
ReDim ColumnsPerPlot(2, 3)
ColumnsPerPlot(0, 0) = lngLastColumn + 6		' X
ColumnsPerPlot(0, 1) = lngLastColumn + 5		' Y1
ColumnsPerPlot(0, 2) = lngLastColumn + 4		' Y2
ColumnsPerPlot(1, 2) = 0
ColumnsPerPlot(2, 2) = LastRow 
ColumnsPerPlot(0, 3) = lngLastColumn + 3		' Y3
ColumnsPerPlot(1, 3) = 0
ColumnsPerPlot(2, 3) = LastRow 
PlotColumnCountArray(0) = 4
objGraph_page.AddWizardPlot("Line Plot", "Multiple Straight Lines", "X Many Y", ColumnsPerPlot, PlotColumnCountArray)
Set objPlot = objCnb.NotebookItems(strGname).GraphPages(0).Graphs(iGraph).Plots(1)
objPlot.Line.SetAttribute(SEA_LINETYPE, 2)
objPlot.Line.SetAttribute(SEA_TYPECOL, -2)
objPlot.Line.SetAttribute(SEA_TYPEREPEAT, 2)
objPlot.Line.SetAttribute(SEA_THICKNESS, 20)
objPlot.Line.SetAttribute(SEA_COLOR, RGB_BLUE)
objPlot.SelectObject
objPlot.SetAttribute(SPA_SENDTOBACK, 0)

' Add text labels
strErr_Section = "adding text labels to the Bland-Altman graph."
ReDim ColumnsPerPlot(2,1)
ColumnsPerPlot(0, 0) = lngLastColumn + 7		' X
ColumnsPerPlot(0, 1) = lngLastColumn + 8		' Y1
PlotColumnCountArray(0) = 2
objGraph_page.AddWizardPlot("Scatter Plot", "Simple Scatter", "XY Pair", ColumnsPerPlot, PlotColumnCountArray)
Set objPlot = objCnb.NotebookItems(strGname).GraphPages(0).Graphs(iGraph).Plots(2)
objPlot.SetAttribute(SSA_OPTIONS, &H00000201&)
objPlot.SetAttribute(SSA_SHAPE, 0)
objPlot.SetAttribute(SSA_SHAPECOL, lngLastColumn + 9)
objPlot.SetAttribute(SSA_SHAPEREPEAT, 4)
objPlot.SetAttribute(SSA_OPTIONS, &H00000200&)
resize_symbols(objPlot)

add_text
If dlg.checkbox1 = 0 Or blnNewPage = True Then
	' centre text, only one graph present
	centre_text
Else
	' format graphs - two on one page
	strErr_Section = "formatting the graphs."	
	objCnb.NotebookItems(strGname).GraphPages(0).Graphs(1).Top = 1000		' modify BA plot to differentiate between the two
	ActiveDocument.CurrentPageItem.Select(False, -1916, 1771, -1916, 1771)	' Select Methods Graph
	Dim XPosAdjust
	XPosAdjust = -200
	Dim Pos()
	ReDim Pos(1)
	Pos(0) = XPosAdjust-1750
	Pos(1) = 4500
	' Set position
	ActiveDocument.CurrentPageItem.SetSelectedObjectsAttribute(SOA_POSEX, Pos)
	Pos(0) = 3500
	Pos(1) = 3500
	ActiveDocument.CurrentPageItem.SetSelectedObjectsAttribute(SOA_SIZEEX, Pos)		
	ActiveDocument.CurrentPageItem.Select(False, -1875, -729, -1875, -729)	' Select BA Plot
	ActiveDocument.CurrentPageItem.SetSelectedObjectsAttribute(SOA_SIZEEX, Pos)	
	Pos(0) = XPosAdjust-1750
	Pos(1) = -650
	' Set position
	ActiveDocument.CurrentPageItem.SetSelectedObjectsAttribute(SOA_POSEX, Pos)	
End If

'Deselect graph by clicking in upper right corner
ActiveDocument.CurrentPageItem.Select(False, 3792, 5061, 3792, 5061)

'Create a redraw to attempt to clear visible selection markers that occurs in some cases
objGraph_page.GraphPages(0).Graphs(0).SetAttribute(SPA_FORCEUPDATE)

' END OF PROGRAM
Exit All

'*********************************************
'	ERROR ROUTINES:
'*********************************************
'*********************************************
' 1. No default worksheet open
'*********************************************
no_open_wsk:
Dim intCa As Integer
Dim intCb As Integer
On Error GoTo no_wsk
For intCb = 0 To Val(objCnb.NotebookItems.Count)
	If CStr(objCnb.NotebookItems(intCa).ItemType) = 1 Then
		Set objCdi = objCnb.NotebookItems(intCa)
		objCdi.Open
		objCdi_name = objCdi.Name
		Set objCdt = objCdi.DataTable
		objCdi.Goto(0,0)
		If MsgBox ("There was no data worksheet open." & vbCrLf & "Continue with this worksheet ? "  & vbCrLf & vbCrLf + _
	"Worksheet:     " & CStr(objCdi_name) & vbCrLf & "Filename:         "  + _
	CStr(objCnb.FullName),vbYesNo,"No Open Worksheet")= vbNo Then
			Exit All
		End If
		Resume RESUME_FROM_NO_OPEN_WSK
	Else
		intCa = intCa + 1
	End If
Next intCb
MsgBox "There is no worksheet present in this notebook.",vbOkOnly,"No Worksheet in Notebook"
Exit All

no_wsk:
MsgBox "There is no worksheet present in this notebook.",vbOkOnly,"No Worksheet in Notebook"
Exit All

ERROR_MESSAGE:
MsgBox "An error occurred while " & strErr_Section & "  Please contact technical support.", 16, "Error Message"
Exit All


End Sub
Private Function BA_dialog(DlgItem$, Action%, SuppValue&) As Boolean
' Prevent editing of cbMethod1
SendDlgItemMessage SuppValue, DlgControlId("cbMethod1"), EM_SETREADONLY, 1, 0
' Prevent editing of cbMethod2
SendDlgItemMessage SuppValue, DlgControlId("cbMethod2"), EM_SETREADONLY, 1, 0

Select Case Action%
	Case 1 ' Dialog box initialization
'		DlgEnable ("pbHelp", False)	' No help yet
		DlgEnable ("OK", False)
		DlgValue("CheckBox1",1)		' Plot methods to true
		DlgValue("CheckBox2",1)		' Regression to true
'		DlgValue("CheckBox3",1)		' Equality to false
		DlgValue("CheckBox4",1)		' Plot BA graph to true
		DlgText("txtInformation", "Select two data columns for the methods to compare and the graph type(s) that you wish to create, and select OK. " + _
"Select Options to change various program settings and to move to the extents of your data in the worksheet.")
	Case 2 ' Value changing or button pressed
		Select Case DlgItem$
			Case "pbSelect1"			' Add selection to method one
				DlgText("cbMethod1", astrData_columns(DlgValue("listColumns")))
				strMethod1 = DlgText("cbmethod1")
				lngMethod1 = DlgValue("listColumns")
				DlgFocus("listColumns")
			Case "pbSelect2"			' Add selection to method two
				DlgText("cbMethod2", astrData_columns(DlgValue("listColumns")))
				strMethod2 = DlgText("cbmethod2")
				lngMethod2 = DlgValue("listColumns")
				DlgFocus("listColumns")
			Case "pbOptions"			' Set other Dialog box for user options
				options_dialog
			Case "pbHelp"				' Open HELP window
				MsgBox " Bland-Altman Analysis"+vbCrLf+vbCrLf+ _
				"The Bland-Altman analysis compares two methods to see if they agree.  It consists of two graphs," +vbCrLf+ _
				"a method comparison graph and the Bland-Altman graph, and some statistics associated with the latter." +vbCrLf+vbCrLf+ _
				"Method Comparison Graph"+vbCrLf+ _
				"The method comparison graph is an XY scatter plot of the two method results values.  Options"+vbCrLf+ _
				"are available to add a linear regression line, it's confidence lines and the line of identity."+vbCrLf+vbCrLf+ _
				"Bland-Altman Graph"+vbCrLf+ _
				"This graph is an XY scatter plot with the difference of the two methods on the Y axis and"+vbCrLf+ _
				"and the average of the two methods on the X axis.  The mean of the differences is displayed"+vbCrLf+ _
				"together with the Limits of Agreement for the difference data."+vbCrLf+vbCrLf+ _
				"Difference Statistics"+vbCrLf+ _
				"      Bias - the mean of the differences."+vbCrLf+ _
				"      Std Dev - the standard deviation of the differences."+vbCrLf+ _
				"      Limits of Agreement - the mean of the differences (bias) +- 1.96 (or 2) times the"+vbCrLf+ _
				"                standard deviation of the differences."+vbCrLf+ _
				"      Confidence Intervals (CI) - these are the 95% or 99% CIs for the "+vbCrLf+ _
				"                1. Bias"+vbCrLf+ _
				"                2. Lower limit of agreement"+vbCrLf+ _
				"                3. Upper limit of agreement"+vbCrLf+vbCrLf+ _
				"Options"+vbCrLf+ _
				"      Number of SDs - used in the Limits of Agreement computation.  Select either 1.96 or 2" & vbCrLf+ _
				"      No. of decimal places - sets the number of decimal places in the difference statistics."+vbCrLf+ _
				"      CIs to be used for B-A plot - used in difference statistic CI computation.  Select either"+vbCrLf+ _
				"                 95% or 99%"+vbCrLf+ _
				"      Add confidence lines - select 95%, 99% or None for the Method Comparison Graph"+vbCrLf+ _
				"                regression confidence lines."+vbCrLf+ _
				"      GoTo Column Position - allows viewing the extents ofyour data while the Bland-Altman graph"+vbCrLf+ _
				"                Dialog is displayed."+vbCrLf+ _
				"      Graph Options - select to place both graphs on either one or two pages."+vbCrLf, "Help"
			
			Case "Cancel"				' Handles Cancel Button
				Exit All
			Case "OK"
				If lngMethod1 = lngMethod2 Then
					MsgBox "The same column has been selected for both" & vbCrLf & "methods. Please change one of" + _
					" your selections.", vbInformation, "Bland-Altman Message"					
				ElseIf DlgValue("CheckBox1") = 0 And DlgValue("CheckBox4") = 0 Then
					If MsgBox("You have not selected a graph type to create. Do you want to continue?",vbOkCancel,"Bland-Altman Plot") = vbYes Then
						Exit Function
					End If
				Else
						Exit Function
				End If
		End Select
		If DlgValue("CheckBox1") = 0 Then	' if not methods plot, can't plot regression nor equality lines
			DlgEnable ("CheckBox2", False)
			DlgEnable ("CheckBox3", False)	
		Else
			DlgEnable ("CheckBox2", True)
			DlgEnable ("CheckBox3", True)	
		End If
		If DlgText("cbmethod1") = "" Or DlgText("cbmethod2") = "" Then
			DlgEnable ("OK", False)
		End If		
		If strMethod1 <> "" And strMethod2 <> "" Then
			DlgEnable ("OK",True)
		End If		
		BA_dialog = True ' Prevent button press from closing the dialog box
End Select
End Function
Public Function empty_col(column As Variant, column_end As Variant)		' Thanks to John Kuo for this function - Determines if a column is empty
	Dim WorksheetTable As Object
	Set WorksheetTable = ActiveDocument.CurrentDataItem.DataTable
	Dim i As Long
	Dim empty_cell As Boolean
	'IDEA HERE - set step value according to lastrow value
	' - theory is is lastrow is large, small steps take too long
	' if lastrow is small, can afford to quick in detail
	For i = 0 To column_end Step 1 'Change the step value to change the sampling interval.  Small sample size = Slow operation
		If WorksheetTable.Cell(column,i) = strQnan Then
			empty_col = True
		Else
			blnNumeric = IsNumeric(WorksheetTable.Cell(column,i))
			empty_col = False
			Exit Function
		End If
	Next i
End Function
Public Sub options_dialog()
Dim strSD_Values(1) As String
strSD_Values(0) = "1.96"
strSD_Values(1) = "2.00"
	Begin Dialog UserDialog 530,217,"Bland-Altman Settings",.options_dlg ' %GRID:10,7,1,1
		GroupBox 20,7,220,161,"Statistics",.GroupBox1
		OptionGroup .ogCI
			OptionButton 60,105,60,14,"95%",.obCI95
			OptionButton 130,105,60,14,"99%",.obCI99
		GroupBox 260,140,250,63,"Graph Options",.GroupBox2
		OKButton 30,182,90,21
		CancelButton 150,182,90,21
		Text 40,28,130,14,"Number of SDs",.Text1
		TextBox 190,49,40,21,.tbDecimals
		OptionGroup .Group1
			OptionButton 280,182,150,14,"Plots on same page",.OptionButton1
			OptionButton 280,161,190,14,"Plots on individual pages",.OptionButton2
		GroupBox 260,7,250,126,"Goto Column Position",.GroupBox3
		ListBox 280,28,120,91,astrData_columns(),.listCols
		PushButton 420,63,70,21,"Top",.pbTop
		PushButton 420,91,70,21,"Bottom",.pbBottom
		OptionGroup .ogConf
			OptionButton 40,147,53,14,"95%",.ob95
			OptionButton 110,147,53,14,"99%",.ob99
			OptionButton 170,147,60,14,"None",.obNone
		Text 40,126,140,14,"Add confidence lines",.Text2
		Text 420,28,70,28,"Position of cursor",.Text3,2
		Text 40,56,140,14,"No. of decimal places",.Text4
		Text 40,84,180,14,"CIs to use for B-A plot",.Text5
		DropListBox 140,28,90,49,strSD_Values(),.lbSD
	End Dialog
				Dim dlg2 As UserDialog
				dlg2.tbDecimals = CStr(Len(strDecimals) - InStr(strDecimals,"0") + 1)
				dlg2.Group1 = 0
				If blnCI99 = False Then
					dlg2.ogCI = 0
				Else
					dlg2.ogCI = 1
				End If
				If blnConf = False Then
					dlg2.ogConf = 2
				Else 
					If blnConf99 = False Then
						dlg2.ogConf = 0
					Else
						dlg2.ogConf = 1
					End If
				End If
				If Dialog(dlg2) <> 0 Then
					' change graph settings
					If dlg2.Group1 = 0 Then
						blnNewPage = False				' Do not create new page for each graph
					Else
						blnNewPage = True				' Do create new page for each graph
					End If
					blnConf =True						' Do add confidence lines to Regression Plot
					If dlg2.ogCI = 0 Then
						blnCI99 = False
					Else
						blnCI99 = True
					End If
					Select Case dlg2.ogConf
						Case 0
							blnConf99 = False
						Case 1
							blnConf99 = True
						Case 2
							blnConf = False				' Do not add confidence lines to Regression Plot
					End Select
				End If
				
End Sub
Private Function options_dlg(DlgItem$, Action%, SuppValue&) As Boolean
Select Case Action%
	Case 2 ' Value changing or button pressed
		Select Case DlgItem$
			Case "pbTop"
				objCdi.Goto(0,alngData_columns(DlgValue("listCols")))
			Case "pbBottom"
				objCdi.Goto(lngLastRow,alngData_columns(DlgValue("listCols")))
			Case "Cancel"
				Exit Function
			Case "OK"
				If DlgValue("lbSD") = 0 Then
					dblPrecision = 1.96
				Else
					dblPrecision = 2
				End If
				Exit Function
		End Select
		options_dlg = True ' Prevent button press from closing the dialog box
	Case 3 ' TextBox or ComboBox text changed
	' check to ensure number of decimal places is numeric NOT text!
		If IsNumeric(DlgText("tbDecimals")) = False Then
			MsgBox "Please enter a positive integer for the" & vbCrLf & "number of decimal places.", vbInformation, "Data Entry Error"
			DlgText("tbDecimals", "4")
		ElseIf CDbl(DlgText("tbDecimals")) < 0 Then
			MsgBox "Please enter a positive integer for the" & vbCrLf & "number of decimal places.", vbInformation, "Data Entry Error"
			DlgText("tbDecimals", "4")
		Else
			If CDbl(DlgText("tbDecimals")) > 10 Then
				MsgBox "Please enter a decimal place value <= 10.", vbInformation, "Data Entry Error"
				DlgText("tbDecimals", "4")
			Else
				' set strDecimals to correct format eg "#.0000" as default tbDecimals = 4
				strDecimals = "#."
				For iIndex = 1 To CDbl(DlgText("tbDecimals"))
					strDecimals = strDecimals & "0"
				Next iIndex
			End If
		End If
	Case 4 ' Focus changed
	Case 5 ' Idle
	End Select
End Function
Sub add_graph_page()										' Add graph page and name it
	objCnb.NotebookItems.Add(CT_GRAPHICPAGE)
	Set objGraph_page = ActiveDocument.CurrentPageItem
	On Error GoTo NAME_ERROR
	Dim iError As Integer
	iError = 1
	Name_graph:
	strGname = "BA Graph: " & iError
	objGraph_page.Name = strGname	
	On Error GoTo 0
	Exit Sub
NAME_ERROR:
	Select Case Err.Number
		Case 65535		'duplicate name
			iError = iError + 1
			Resume name_graph
		Case Else
			MsgBox (Err.Description & " (" & CStr(Err.Number) & ") when naming the graph page",16, "Error")
			Exit All
	End Select	
End Sub
Sub hide_legend(objPlot As Object)
	objPlot.SetAttribute(SGA_FLAGS, FlagOff(SGA_FLAG_AUTOLEGENDSHOW))	' Remove legend
End Sub
Sub create_graph()
objGraph_page.CreateWizardGraph("Scatter Plot", "Simple Scatter", "XY Pair", ColumnsPerPlot, PlotColumnCountArray)
If strGraph_title = "Method 1 versus Method 2" Then
	Set objPlot = objCnb.NotebookItems(strGname).GraphPages(0).Graphs(0)
Else
	Set objPlot = objCnb.NotebookItems(strGname).GraphPages(0).Graphs(iGraph)
End If

'Fill symbol gray
Dim objSymbol As Object
Set objSymbol = objPlot.Plots(0).Symbols	
objSymbol.SetAttribute(SSA_COLOR, &H00c0c0c0&)

objPlot.Name = strGraph_title
hide_legend(objPlot)
objPlot.Axes(0).SetAttribute(SAA_OPTIONS,SAA_FLAG_AUTORANGE Or FLAG_CLEAR_BIT)
objPlot.Axes(0).GetAttribute(SAA_FROMVAL,lngXmin)
objPlot.Axes(0).GetAttribute(SAA_TOVAL, lngXmax)

'Make axis titles bold
Dim XAxis, YAxis
Set XAxis=objPlot.Axes(0)
Set YAxis=objPlot.Axes(1)
Dim SPXAxisTitle, SPYAxisTitle
Set SPXAxisTitle = XAxis.AxisTitles(0)
Set SPYAxisTitle = YAxis.AxisTitles(0)
SPXAxisTitle.SetAttribute(STA_SELECT, 5)
SPXAxisTitle.SetAttribute(STA_BOLD, True)
SPYAxisTitle.SetAttribute(STA_SELECT, 5)
SPYAxisTitle.SetAttribute(STA_BOLD, True)

'Reposition graph name and make it bold
Dim objGraphName As Object
Set objGraphName = objCnb.NotebookItems(strGname).GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject
'Dim GraphName As String
'GraphName = objGraphName.Name  'check to see if I got the object - I did
objGraphName.SelectObject      'select it
objGraphName.SetAttribute(STA_BOLD,True)  'make it bold
Dim Pos()
ReDim Pos(1)
Pos(0) = 0
Pos(1) = 50   'positioni the graph name slightly above the graph
objGraphName.SetAttribute(STA_RELTACKPOINT, Pos)

If strGraph_title = "Method 1 versus Method 2" Then
	objPlot.Axes(0).Name = "Method 1"
	objPlot.Axes(1).Name = "Method 2"
Else
	objPlot.Axes(0).Name = "Average of Method 1 and Method 2"
	objPlot.Axes(1).Name = "Difference Method 2 - Method 1"
End If
End Sub
Function get_dblTvalue95(ByVal Number As Integer)

Number = Number - 2 ' Degrees of Freedom = sample size - 1 (and -1 for array)

Select Case Number
	Case 1
		get_dblTvalue95 = 12.706
	Case 2
		get_dblTvalue95 = 4.303
	Case 3
		get_dblTvalue95 = 3.182
	Case 4
		get_dblTvalue95 = 2.776
	Case 5
		get_dblTvalue95 = 2.571
	Case 6
		get_dblTvalue95 = 2.447
	Case 7
		get_dblTvalue95 = 2.365
	Case 8
		get_dblTvalue95 = 2.306
	Case 9
		get_dblTvalue95 = 2.262
	Case 10
		get_dblTvalue95 = 2.228
	Case 11
		get_dblTvalue95 = 2.201
	Case 12
		get_dblTvalue95 = 2.179
	Case 13
		get_dblTvalue95 = 2.16
	Case 14
		get_dblTvalue95 = 2.145
	Case 15
		get_dblTvalue95 = 2.131
	Case 16
		get_dblTvalue95 = 2.12
	Case 17
		get_dblTvalue95 = 2.11
	Case 18
		get_dblTvalue95 = 2.101
	Case 19
		get_dblTvalue95 = 2.093
	Case 20
		get_dblTvalue95 = 2.086
	Case 21
		get_dblTvalue95 = 2.08
	Case 22
		get_dblTvalue95 = 2.074
	Case 23
		get_dblTvalue95 = 2.069
	Case 24
		get_dblTvalue95 = 2.064
	Case 25
		get_dblTvalue95 = 2.06
	Case 26
		get_dblTvalue95 = 2.056
	Case 27
		get_dblTvalue95 = 2.052
	Case 28
		get_dblTvalue95 = 2.048
	Case 29
		get_dblTvalue95 = 2.045
	Case 30
		get_dblTvalue95 = 2.042
	Case Is <=35
		get_dblTvalue95 = 2.030
	Case Is <=40
		get_dblTvalue95 = 2.021
	Case Is <=45
		get_dblTvalue95 = 2.014
	Case Is <=50
		get_dblTvalue95 = 2.009
	Case Is <=60
		get_dblTvalue95 = 2.000
	Case Is <=70
		get_dblTvalue95 = 1.994
	Case Is <=80
		get_dblTvalue95 = 1.990
	Case Is <=90
		get_dblTvalue95 = 1.987
	Case Is <=100
		get_dblTvalue95 = 1.984
	Case Is <=120
		get_dblTvalue95 = 1.980
	Case Else
		get_dblTvalue95 = 1.960
End Select
End Function
Function get_dblTvalue99(ByVal Number As Integer)

Number = Number - 2 ' Degrees of Freedom = sample size - 1 (and -1 for array)

Select Case Number
	Case 1
		get_dblTvalue99 = 6.314
	Case 2
		get_dblTvalue99 = 2.920
	Case 3
		get_dblTvalue99 = 2.353
	Case 4
		get_dblTvalue99 = 2.132
	Case 5
		get_dblTvalue99 = 2.015
	Case 6
		get_dblTvalue99 = 1.943
	Case 7
		get_dblTvalue99 = 1.895
	Case 8
		get_dblTvalue99 = 1.860
	Case 9
		get_dblTvalue99 = 1.833
	Case 10
		get_dblTvalue99 = 1.812
	Case 11
		get_dblTvalue99 = 1.796
	Case 12
		get_dblTvalue99 = 1.782
	Case 13
		get_dblTvalue99 = 1.771
	Case 14
		get_dblTvalue99 = 1.761
	Case 15
		get_dblTvalue99 = 1.753
	Case 16
		get_dblTvalue99 = 1.745
	Case 17
		get_dblTvalue99 = 1.740
	Case 18
		get_dblTvalue99 = 1.734
	Case 19
		get_dblTvalue99 = 1.729
	Case 20
		get_dblTvalue99 = 1.725
	Case 21
		get_dblTvalue99 = 1.721
	Case 22
		get_dblTvalue99 = 1.717
	Case 23
		get_dblTvalue99 = 1.714
	Case 24
		get_dblTvalue99 = 1.711
	Case 25
		get_dblTvalue99 = 1.708
	Case 26
		get_dblTvalue99 = 1.706
	Case 27
		get_dblTvalue99 = 1.703
	Case 28
		get_dblTvalue99 = 1.701
	Case 29
		get_dblTvalue99 = 1.699
	Case 30
		get_dblTvalue99 = 1.697
	Case Is <=35
		get_dblTvalue99 = 1.690
	Case Is <=40
		get_dblTvalue99 = 1.684
	Case Is <=45
		get_dblTvalue99 = 1.680
	Case Is <=50
		get_dblTvalue99 = 1.676
	Case Is <=60
		get_dblTvalue99 = 1.671
	Case Is <=70
		get_dblTvalue99 = 1.667
	Case Is <=80
		get_dblTvalue99 = 1.664
	Case Is <=90
		get_dblTvalue99 = 1.662
	Case Is <=100
		get_dblTvalue99 = 1.660
	Case Is <=120
		get_dblTvalue99 = 1.658
	Case Else
		get_dblTvalue99 = 1.645
End Select

End Function
Public Function udHelpBox(DlgItem$, Action%, SuppValue&) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		Select Case DlgItem$
		Case "Help"
'			Help(ObjectHelp,HelpID)
'        	udHelpBox = False
MsgBox "Help"
        End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
	Case 6 ' Function key
	End Select
End Function
Public Sub add_text()

	' Precision of Estimated Limits of Agreement
	Dim dblAgreeL As Double
	Dim dblAgreeU As Double
	Dim dblBiasL As Double
	Dim dblBiasU As Double
	Dim dbl95LL As Double
	Dim dbl95LU As Double
	Dim dbl95UL As Double
	Dim dbl95UU As Double
	Dim dblStdErr As Double
	Dim dblT As Double
	Dim dblTvalue As Double
	Dim strPercent As String
	If blnCI99 = False Then
		dblTvalue = get_dblTvalue95(iNo)					' get T value
		strPercent = "95%"
	Else
		dblTvalue = get_dblTvalue99(iNo)					' get T value
		strPercent = "99%"
	End If
	
	dblAgreeL = dblMean - dblPrecision*dblStddev
	dblAgreeU = dblMean + dblPrecision*dblStddev
	
	dblStdErr = ((dblStddev^2) / iNo ) ^ 0.5			' Standard Error
	dblBiasL = Format ((dblMean - (dblTvalue * dblStdErr)), strDecimals)
	dblBiasU = Format ((dblMean + (dblTvalue * dblStdErr)), strDecimals)
	
	dblStdErr = ((3 * (dblStddev^2)) / iNo ) ^ 0.5
	
	dbl95LL = Format ((objCdt.Cell(lngLastColumn + 4, 0) - (dblTvalue * dblStdErr)), strDecimals)
	dbl95LU = Format ((objCdt.Cell(lngLastColumn + 4, 0) + (dblTvalue * dblStdErr)), strDecimals)
	dbl95UL = Format (( objCdt.Cell(lngLastColumn + 5, 0) - (dblTvalue * dblStdErr)), strDecimals)
	dbl95UU = Format ((objCdt.Cell(lngLastColumn + 5, 0) + (dblTvalue * dblStdErr)), strDecimals)
	
	' Add Text...
	Dim strText As String
	Dim Points()
	ReDim Points(1)
	
	Dim XPosAdjust
	XPosAdjust = -200
	Dim YPosAdjust
	YPosAdjust = 200
	Points(0) = XPosAdjust+1902
	Points(1) = YPosAdjust-2785
	strText = "Bias = " & Format (dblMean, strDecimals)
	objGraph_page.GraphPages(0).ChildObjects.Add(GPT_TEXT, strText, Points)
	Points(1) = YPosAdjust-2985
	strText = "Std Dev = " & Format (dblStddev, strDecimals)
	objGraph_page.GraphPages(0).ChildObjects.Add(GPT_TEXT, strText, Points)
	Points(1) = YPosAdjust-3185
	strText = "Limits of Agreement = " & Format (dblAgreeL, strDecimals) & ", " & Format(dblAgreeU, strDecimals)
	objGraph_page.GraphPages(0).ChildObjects.Add(GPT_TEXT, strText, Points)
	Points(1) = YPosAdjust-3385
	strText = "Bias CI"
	objGraph_page.GraphPages(0).ChildObjects.Add(GPT_TEXT, strText, Points)
	Points(1) = YPosAdjust-3785
	strText = "Lower Limit of Agreement CI"
	objGraph_page.GraphPages(0).ChildObjects.Add(GPT_TEXT, strText, Points)
	Points(1) = YPosAdjust-4185
	strText = "Upper Limit of Agreement CI"
	objGraph_page.GraphPages(0).ChildObjects.Add(GPT_TEXT, strText, Points)
	Points(0) = XPosAdjust+1982			' Slight indent for second lines
	Points(1) = YPosAdjust-3585
	strText = strPercent & " CI = " & dblBiasL & " To " & dblBiasU
	objGraph_page.GraphPages(0).ChildObjects.Add(GPT_TEXT, strText, Points)
	Points(1) = YPosAdjust-3985
	strText = strPercent & " CI = " & dbl95LL & " to " & dbl95LU
	objGraph_page.GraphPages(0).ChildObjects.Add(GPT_TEXT, strText, Points)
	Points(1) = YPosAdjust-4385
	strText = strPercent & " CI = " & dbl95UL & " to " & dbl95UU
	objGraph_page.GraphPages(0).ChildObjects.Add(GPT_TEXT, strText, Points)
End Sub
Public Sub centre_text()
	objGraph_page.Select(False, 1500, -2250, 5400, -4500)
	objGraph_page.GraphPages(0).SetAttribute(SPA_GROUP, 0)
	objGraph_page.Select(False, 1500, -2250, 5400, -4500)
	Dim Pos()
	ReDim Pos(1)
	Pos(0) = -961
	Pos(1) = -2374
	objGraph_page.SetSelectedObjectsAttribute(SOA_POSEX, Pos)
End Sub
Public Sub resize_symbols(objPlot As Object)
Debug.Print objPlot.Name
	objPlot.SetAttribute(SSA_SIZE, 80)		' resize labels to smaller size for this graph
	objPlot.SetAttribute(SSA_SIZEREPEAT, 2)
End Sub
Public Function check_data(objXfm As Object) As Boolean
objXfm.Text = 	"If size(col(method1))<>size(col(method2)) then" + vbCrLf + _
						"cell(results_column"+sL+"1)=1/0" + vbCrLf + _
						"else" + vbCrLf + _
						"if missing(col(method1)) or missing(col(method2)) <> 0 then" + vbCrLf + _
						"cell(results_column"+sL+"2)=1/0" + vbCrLf + _
						"end if" + vbCrLf + _
						"end if" + vbCrLf						
'objXfm.RunEditor
objXfm.Execute
objXfm.Close(False)

If objCdt.Cell(lngLastColumn+1, 0) = strInf Then
	objCdt.Cell(lngLastColumn+1, 0) = ""
	If MsgBox("The two columns have a different number of" & vbCrLf & "data values in each. Please reselect your columns, or change your " + _
"data.", vbOkCancel, "Bland-Altman Message")=vbOK Then
		check_data = True
		Exit Function
	Else
		Exit All
	End If
ElseIf objCdt.Cell(lngLastColumn+1, 1) = strInf Then
	objCdt.Cell(lngLastColumn+1, 1) = ""
	If MsgBox("One of the two columns contains one or more missing values" & vbCrLf & "(or text). Please reselect your columns, or change your " + _
"data.", vbOkCancel, "Bland-Altman Message")=vbOK Then
		check_data = True
		Exit Function
	Else
		Exit All
	End If
End If
End Function
Public Function open_xfm() As Object
	Dim objXfm As Object
	Set objXfm = ActiveDocument.NotebookItems.Add(9)
	objXfm.Open
	objXfm.AddVariableExpression("method1", lngMethod1 + 1)
	objXfm.AddVariableExpression("method2", lngMethod2 + 1)
	objXfm.AddVariableExpression("results_column", lngLastColumn + 2)	
	Set open_xfm = objXfm
End Function
Public Sub analysis(objXfm As Object)
objXfm.AddVariableExpression("precision", dblPrecision)
objXfm.Text = 	"col(results_column)=col(method2)-col(method1)" + vbCrLf + _
						"col(results_column+1)=(col(method2)+col(method1))/2" + vbCrLf + _
						"m=mean(col(results_column))" + vbCrLf + _
						"sd=stddev(col(results_column))" + vbCrLf + _
						"cell(results_column+2"+sL+"1)=m" + vbCrLf + _
						"cell(results_column+2"+sL+"2)=m" + vbCrLf + _
						"cell(results_column+3"+sL+"1)=m-(precision*sd)" + vbCrLf + _
						"cell(results_column+3"+sL+"2)=m-(precision*sd)" + vbCrLf + _
						"cell(results_column+4"+sL+"1)=m+(precision*sd)" + vbCrLf + _
						"cell(results_column+4"+sL+"2)=m+(precision*sd)" + vbCrLf + _
						"cell(results_column+5"+sL+"1)=sd" + vbCrLf + _
						"cell(results_column+5"+sL+"3)=m" + vbCrLf + _
						"cell(results_column+5"+sL+"2)=size(col(method1))" + vbCrLf
						
'objXfm.RunEditor
objXfm.Execute
objXfm.Close(False)

' Grab dblStddev from wsksheet
dblStddev = objCdt.Cell(lngLastColumn + 6, 0)
iNo = objCdt.Cell(lngLastColumn + 6, 1)
dblMean = objCdt.Cell(lngLastColumn + 6, 2)
' Add column titles & other graphing values
objCdt.Cell(lngLastColumn + 1, -1) = "Differences"
objCdt.Cell(lngLastColumn + 2, -1) = "Means"
objCdt.Cell(lngLastColumn + 3, -1) = "Mean of Differences"
objCdt.Cell(lngLastColumn + 4, -1) = "mean - " & CStr(dblPrecision) & " SD"
objCdt.Cell(lngLastColumn + 5, -1) = "mean + " & CStr(dblPrecision) & " SD"
objCdt.Cell(lngLastColumn + 6, -1) = "X Axis Scale"
objCdt.Cell(lngLastColumn + 7, -1) = "Symbol X"
objCdt.Cell(lngLastColumn + 8, -1) = "Symbol Y"
objCdt.Cell(lngLastColumn + 9, -1) = "Symbol Labels"
Dim dblPercent1 As Double, dblPercent2 As Double
dblPercent1 = (Sgn(CDbl(objCdt.Cell(lngLastColumn + 5, 0))) * CDbl(objCdt.Cell(lngLastColumn + 5, 0))) 
dblPercent2 = (Sgn(CDbl(objCdt.Cell(lngLastColumn + 4, 0))) * CDbl(objCdt.Cell(lngLastColumn + 4, 0)))
dblPercent1 = ((dblPercent1 + dblPercent2)/100)*3	' 3%
objCdt.Cell(lngLastColumn + 8, 0) = objCdt.Cell(lngLastColumn + 3, 0) + (Sgn(objCdt.Cell(lngLastColumn + 3, 0)) * dblPercent1)
objCdt.Cell(lngLastColumn + 8, 1) = objCdt.Cell(lngLastColumn + 4, 0) + (Sgn(objCdt.Cell(lngLastColumn + 4, 0)) * dblPercent1)
objCdt.Cell(lngLastColumn + 8, 2) = objCdt.Cell(lngLastColumn + 5, 0) + (Sgn(objCdt.Cell(lngLastColumn + 5, 0)) * dblPercent1)
objCdt.Cell(lngLastColumn + 9, 0) = "Mean"
objCdt.Cell(lngLastColumn + 9, 1) = "Mean - " & CStr(dblPrecision) & "SD"
objCdt.Cell(lngLastColumn + 9, 2) = "Mean + " & CStr(dblPrecision) & "SD"
End Sub