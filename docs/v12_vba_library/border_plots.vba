Option Explicit
Dim Separator$, Data_X$, Data_Y$, Results_X$, Results_Y$
Dim SPGraphX, SPGraphY
Dim SPGraphXAxisX, SPGraphXAxisY, SPGraphXTitleX, SPGraphXTitleY, SPTitleX, SPLegendX   'X dimensions
Dim SPGraphYAxisX, SPGraphYAxisY, SPGraphYTitleX, SPGraphYTitleY, SPTitleY, SPLegendY   'Y dimensions
Dim SPPlot, PlotType, SPGraph, SPPage, SPTitle, SPGraphAxisX, SPGraphAxisY 'Graph Objects
Dim CurrentXTuple, CurrentYTuple, PlottedCurves, Bins
Dim LastColumn As Long
Dim LastRow As Long
Dim WorksheetTable As Object
Dim XStart, XEnd, YStart, YEnd
Public DrawBoxPlot As Boolean
Public Const ObjectHelp = Path + "\SPW.CHM"
Public HelpID As Variant
Function FlagOn(flag As Long)
  FlagOn = flag Or FLAG_SET_BIT ' Use to set option flag bits on, leaving others unchanged
End Function
Function FlagOff(flag As Long)
  FlagOff = flag Or FLAG_CLEAR_BIT ' Use to set option flag bits off, leaving others unchanged
End Function
Sub Main
Separator = ListSeparator
	'Border Plots macro
	'Authored by Frederick Cabasa 11/4/98; updated 12/2/1999 John Kuo
	'Updated on 9/07/01 to fix error message with scatter plot and help  (Frederick Cabasa)
	'This macro creates frequency plots in the form of either histograms (of equally 
	'sized bins) or box plots along the top and left axes of 2D scatter plots.  The 
	'macro works for any current and open graph page, and can convert any 2D Cartesian
	'graph into the equivalent scatter plot.
	
	HelpID = 60201			' Help ID number for this topic in SPW.CHM
	On Error GoTo ErrorMsg 'Jumps to error condition checking
	'Error message logic
	Dim ErrorCheck As Integer
	
	'Find first empty column	
	ErrorCheck = 3 'Display no data table Error message 	
	Set WorksheetTable = ActiveDocument.CurrentDataItem.DataTable
	LastColumn = 0
	LastRow = 0 
	WorksheetTable.GetMaxUsedSize(LastColumn,LastRow)
	'Place Worksheet into Overwrite mode
	ActiveDocument.CurrentDataItem.InsertionMode = False	
	'Set column titles for results
	
	ErrorCheck = 0 'Display no open page error message on error

	'Selecting the current page and graph
	Set SPPage = ActiveDocument.CurrentPageItem
	SPPage.Open
	ErrorCheck = 2 'Display no graphs error message on error
	Set SPGraph = SPPage.GraphPages(0).CurrentPageObject(GPT_GRAPH)
	Set SPPlot = SPGraph.Plots(0)
	Set SPTitle = SPGraph.NameObject
	Set SPGraphAxisX = SPGraph.Axes(0)
	Set SPGraphAxisY = SPGraph.Axes(1)
	ErrorCheck = 0 

	'Specifies "hard-coded" number of histogram bins; edit this array to change the values in the list
	Dim NumberOfBins$()
	ReDim NumberOfBins$(7)
	NumberOfBins$(0) = "5"
	NumberOfBins$(1) = "10"
	NumberOfBins$(2) = "15"
	NumberOfBins$(3) = "20"
	NumberOfBins$(4) = "25"
	NumberOfBins$(5) = "30"
	NumberOfBins$(6) = "40"
	NumberOfBins$(7) = "50"
	
	'Verify that graph is a scatterplot
	PlotType = SPPlot.GetAttribute(SLA_TYPE, PlotType)
	Select Case PlotType
		Case 1
			GoTo MacroDialog
		Case 2 
			MsgBox "The selected graph "+ Chr(34) + SPPlot.Parent.Parent.Name + Chr(34) +" is not a Scatter Plot", vbExclamation,"SigmaPlot"
		Case Else
			HelpMsgBox 60201, "This macro only operates on 2D Cartesian Graphs",vbExclamation,"SigmaPlot"
			GoTo Finish
	End Select

	'Convert option to a scatter plot
	Begin Dialog UserDialog 390,70,"Border Plot Macro" ' %GRID:10,7,1,1
		Text 10,7,370,28,"The Border Plots macro requires a scatter plot.  Do you want to convert the current graph to a scatter plot?",.Text1
		OKButton 80,42,90,21
		CancelButton 220,42,90,21
	End Dialog
	Dim dlgConvert As UserDialog
	
		If Dialog(dlgConvert) = 0 Then  'handles cancel button
			GoTo Finish
		Else
			SPPlot.SetAttribute(SLA_TYPE,SLA_TYPE_SCATTER)
		End If

DrawBoxPlot = False
MacroDialog:
	ErrorCheck = 1 'Display positive number of bins required error message on error
	Begin Dialog UserDialog 400,100,"Border Plots",.border ' %GRID:10,7,1,0
		OptionGroup .Choices
			OptionButton 12,12,150,14,"&Histogram border",.histogram
			OptionButton 12,30,176,18,"Box &plot border",.boxplot
		Text 166,12,30,14,"&Bins",.BinTitle
		ComboBox 206,10,60,76,NumberOfBins(),.Bins
		OKButton 294,10,96,21
		CancelButton 294,38,96,21
		PushButton 294,70,96,21,"Help",.PushButton2
	End Dialog
	Dim dlg As UserDialog
	If dlg.bins="" Then dlg.bins = "10"
	
	
	Select Case Dialog(dlg)  
		Case 0 'Handles Cancel button
			GoTo Finish
'		Case 1 'Handles Help button
'			HelpID = 60201			' Help ID number for this topic in SPW.CHM
'			Help(ObjectHelp,HelpID)
'		GoTo MacroDialog
	End Select
	
	'Handles non-numeric entries
	If IsNumeric(dlg.bins) = False Then
		MsgBox "You must enter a positive integer for the number of bins",vbExclamation,"SigmaPlot"
		dlg.bins = "10"		
		GoTo MacroDialog
	End If
	
	'Handles negative integer entries
	If CInt(dlg.bins) < 0 Then
		MsgBox "You must enter a positive integer for the number of bins",vbExclamation,"SigmaPlot"
		dlg.bins = "10"		
		GoTo MacroDialog
	End If
		

Bins=dlg.bins

	'Set the axis ranges to the data range
	SPGraphAxisX.SetAttribute(SAA_OPTIONS, SAA_FLAG_NOAUTOPAD)
	SPGraphAxisX.SetAttribute(SAA_OPTIONS, 12583944)
	SPGraphAxisX.SetAttribute(SAA_SUB1OPTIONS, &H0000000d)
	SPGraphAxisX.SetAttribute(SAA_SUB2OPTIONS, &H0000000d)
	SPGraphAxisY.SetAttribute(SAA_OPTIONS, SAA_FLAG_NOAUTOPAD)
	SPGraphAxisY.SetAttribute(SAA_OPTIONS, 12583944)
	SPGraphAxisY.SetAttribute(SAA_SUB1OPTIONS, &H0000000d)
	SPGraphAxisY.SetAttribute(SAA_SUB2OPTIONS, &H0000000d)

	'Determine X and Y Data Columns
	Set PlottedCurves = SPPlot.ChildObjects
	
	PlottedCurves(0).SetAttribute(SNA_SELECTDIM, DIM_X)
	CurrentXTuple = PlottedCurves(0).GetAttribute(SNA_DATACOL, CurrentXTuple)
	PlottedCurves(0).SetAttribute(SNA_SELECTDIM, DIM_Y)
	CurrentYTuple = PlottedCurves(0).GetAttribute(SNA_DATACOL, CurrentYTuple)
	
	'Set column titles for results	
	Results_X = "X Histogram"
	Results_Y = "Y Histogram"

	'Determine if X only or Y only plot	
	Dim DataFormat$
	If CurrentXTuple < 0 And CurrentYTuple >= 0 Then
		DataFormat = "Y Only"
	ElseIf CurrentXTuple >= 0 And CurrentYTuple < 0 Then
		DataFormat = "X Only"
	ElseIf CurrentXTuple >= 0 And CurrentYTuple >= 0 Then
		DataFormat = "XY"
	End If


Select Case DataFormat
	Case "XY"
	Select Case dlg.Choices
		Case 1 	
			BoxPlot_X
			BoxPlot_Y
		Case Else	
			Histogram_X
			Histogram_Y
		End Select
		Manipulate_X
		Manipulate_Y
		GoTo Finish
	Case "X Only"
		Select Case dlg.Choices
		Case 1 	
			BoxPlot_X
		Case Else
			Histogram_X
		End Select
		Manipulate_X
		GoTo Finish
	Case "Y Only"
		Select Case dlg.Choices
		Case 1 	
			BoxPlot_Y
		Case Else
			Histogram_Y
		End Select
		Manipulate_Y
		GoTo Finish
End Select

ErrorMsg:
	If ErrorCheck = 0 Then 
		HelpMsgBox 60201, "You must have a graph page open to run this macro",vbExclamation,"SigmaPlot"
	ElseIf ErrorCheck = 1 Then 
		MsgBox "You must enter a positive integer for the number of bins",vbExclamation,"SigmaPlot"
		dlg.bins = "10"		
		GoTo MacroDialog
	ElseIf ErrorCheck = 2 Then 
		HelpMsgBox 60201, "This notebook contains no graphs.  Please create a new graph before running this macro.", vbExclamation,"SigmaPlot"
		GoTo Finish	
	ElseIf ErrorCheck = 3 Then
		HelpMsgBox 60201, "This notebook contains no Worksheet.  Please create a new worksheet before running this macro.", vbExclamation,"SigmaPlot"
		GoTo Finish	
	End If
	
Finish:	
End Sub

'This function hides/shows bins control depending on select state of border plot type.
'See DialogFunc help topic for more information.
Private Function border(DlgItem$, Action%, SuppValue%) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
		If DrawBoxPlot = False Then
			DlgVisible ("Bins", True)
			DlgVisible ("BinTitle", True)
		Else
			DlgVisible ("Bins", False)
			DlgVisible ("BinTitle", False)
		End If
	Case 2 ' Value changing or button pressed
		Rem border = True ' Prevent button press from closing the dialog box
		Select Case DlgItem$
			Case "PushButton2"
				Help(ObjectHelp,HelpID)
				border = True 'do not exit the dialog
			Case "Cancel"
				End
			Case "Choices"
				If SuppValue% = 0 Then
					DlgVisible ("Bins", True)
					DlgVisible ("BinTitle", True)
					DrawBoxPlot = False
				Else
					DlgVisible ("Bins", False)
					DlgVisible ("BinTitle", False)
					DrawBoxPlot = True
				End If
		End Select 
    
	End Select
End Function
Public Function Histogram_X
	'Run Histogram Transform for X Values and place values in worksheet
	Dim SPTransformX As Object
	Set SPTransformX = ActiveDocument.NotebookItems.Add(9)
	SPTransformX.Open
	SPTransformX.Text = "col(c)=histogram(col(x)" + Separator + "b)" + vbCrLf
	Dim HistogramParametersX(2)
	HistogramParametersX(0) = CStr(LastColumn + 1)
	HistogramParametersX(1)= CStr(CurrentXTuple + 1)
	HistogramParametersX(2)= Bins
	SPTransformX.AddVariableExpression("c", HistogramParametersX(0))
	SPTransformX.AddVariableExpression("x", HistogramParametersX(1))
	SPTransformX.AddVariableExpression("b", HistogramParametersX(2))
	SPTransformX.Execute
	SPTransformX.Close(False)

	'Add Bar Chart for X Values	
	Dim PlottedColumnX(0)
	PlottedColumnX(0) = LastColumn 

	SPPage.CreateWizardGraph("Vertical Bar Chart", "Simple Bar", "Single Y", PlottedColumnX)	
	Set SPGraphX = SPPage.GraphPages(0).CurrentPageObject(GPT_GRAPH)
	'Label Column
	WorksheetTable.NamedRanges.Add(Results_X,LastColumn,0,1,-1, True)
	'Set bar chart properties:  no edge color, light gray fill, 100% width 
	SPGraphX.Plots(0).SetAttribute(SDA_EDGECOLOR, &Hff000000)
	SPGraphX.Plots(0).SetAttribute(SDA_COLOR, &H00c0c0c0)
	SPGraphX.Plots(0).SetAttribute(SLA_BARTHICKNESS, 1000)
End Function	
Public Function Histogram_Y
	'Run Histogram Transform for Y Values and place values in worksheet
	Dim SPTransformY As Object
	Set SPTransformY = ActiveDocument.NotebookItems.Add(9)
	SPTransformY.Open
	SPTransformY.Text = "col(c)=histogram(col(y)" + Separator + "b)" + vbCrLf
	Dim HistogramParametersY(2)
	HistogramParametersY(0) = CStr(LastColumn + 2)
	HistogramParametersY(1)= CStr(CurrentYTuple + 1)
	HistogramParametersY(2)= Bins
	SPTransformY.AddVariableExpression("c", HistogramParametersY(0))
	SPTransformY.AddVariableExpression("y", HistogramParametersY(1))
	SPTransformY.AddVariableExpression("b", HistogramParametersY(2))
	SPTransformY.Execute
	SPTransformY.Close(False)
	
	'Add Bar Chart for Y Values
	Dim PlottedColumnY(0)
	PlottedColumnY(0) = LastColumn + 1
	SPPage.CreateWizardGraph("Horizontal Bar Chart", "Simple Bar", "Single X", PlottedColumnY) 
	Set SPGraphY =  SPPage.GraphPages(0).CurrentPageObject(GPT_GRAPH)
	'Set bar chart properties:  no edge color, light gray fill, 100% width 
	SPGraphY.Plots(0).SetAttribute(SDA_EDGECOLOR, &Hff000000)
	SPGraphY.Plots(0).SetAttribute(SDA_COLOR, &H00c0c0c0)
	SPGraphY.Plots(0).SetAttribute(SLA_BARTHICKNESS, 1000)

	'Label Column
	WorksheetTable.NamedRanges.Add(Results_Y,LastColumn + 1,0,1,-1, True)
End Function	
Public Function BoxPlot_X
	'Add Box Plot for X Values	
	Dim PlottedBoxX(0)
	PlottedBoxX(0) = CurrentXTuple
	SPPage.CreateWizardGraph("Box Plot", "Horizontal Box Plot", "Many X", PlottedBoxX)
	Set SPGraphX = SPPage.GraphPages(0).CurrentPageObject(GPT_GRAPH)	
End Function
Public Function BoxPlot_Y
	'Add Box Plot for Y Values	
	Dim PlottedBoxY(0)
	PlottedBoxY(0) = CurrentYTuple
	SPPage.CreateWizardGraph("Box Plot", "Vertical Box Plot", "Many Y", PlottedBoxY)
	Set SPGraphY = SPPage.GraphPages(0).CurrentPageObject(GPT_GRAPH)		
End Function
Public Function Manipulate_Y  'Manipulate Y graph
	SPGraphY.Top = SPGraph.Top
	SPGraphY.Left = SPGraph.Left + SPGraph.Width + 100
	SPGraphY.Width = SPGraph.Height/5
	SPGraphY.Height = SPGraph.Height
	SPGraphY.Plots(0).SetAttribute(SSA_SIZE, 100)
	'remove titles
	Set SPGraphYAxisX = SPGraphY.Axes(0)
	Set SPGraphYAxisY = SPGraphY.Axes(1)
	Set SPTitleY = SPGraphY.NameObject
	Set SPLegendY = SPGraphY.AutoLegend
	Set SPGraphYTitleX = SPGraphY.Axes(0).AxisTitles(0)
	Set SPGraphYTitleY = SPGraphY.Axes(1).AxisTitles(0)

	'Set axis ranges to data range
	SPLegendY.ChildObjects(1).SetAttribute(STA_OPTIONS, 4960)	
	SPGraphYAxisX.SetAttribute(SAA_OPTIONS, SAA_FLAG_NOAUTOPAD)
	SPGraphYAxisX.SetAttribute(SAA_OPTIONS, 12583944)
	SPGraphYAxisY.SetAttribute(SAA_OPTIONS, SAA_FLAG_NOAUTOPAD)
	SPGraphYAxisY.SetAttribute(SAA_OPTIONS, 12583944)
	SPGraphY.SetAttribute(SGA_SHOWNAME, 0)
	SPGraphY.SetAttribute(SLA_QCOPTIONS, &H00008044) 'hide zero reference line

	'Move graph title up
	SPTitle.SelectObject
	Dim Pos
	ReDim Pos(1)
	Pos(0)=SPTitle.Width/2 + SPGraph.NameObject.Left
	Pos(1)=SPTitle.Top + (SPGraph.Height/5)/2
	ActiveDocument.CurrentPageItem.SetSelectedObjectsAttribute(SOA_POSEX, Pos)
End Function
Public Function Manipulate_X	'manipulate x graph
	'change location
	SPGraphX.Top = SPGraph.Top + 100 + (SPGraph.Height/5)
	SPGraphX.Left = SPGraph.Left
	SPGraphX.Width = SPGraph.Width
	SPGraphX.Height = SPGraph.Height/5
	
	'remove titles, legend, and axes
	Set SPGraphXAxisX = SPGraphX.Axes(0)
	Set SPGraphXAxisY = SPGraphX.Axes(1)
	Set SPTitleX = SPGraphX.NameObject
	Set SPLegendX = SPGraphX.AutoLegend
	Set SPGraphXTitleX = SPGraphX.Axes(0).AxisTitles(0)
	Set SPGraphXTitleY = SPGraphX.Axes(1).AxisTitles(0)

	'Set axis ranges to data range
	SPLegendX.ChildObjects(1).SetAttribute(STA_OPTIONS, 4960)
	SPGraphXAxisX.SetAttribute(SAA_OPTIONS, SAA_FLAG_NOAUTOPAD)
	SPGraphXAxisX.SetAttribute(SAA_OPTIONS, 12583944)
	SPGraphXAxisY.SetAttribute(SAA_OPTIONS, SAA_FLAG_NOAUTOPAD)	
	SPGraphXAxisY.SetAttribute(SAA_OPTIONS, 12583944)
	SPGraphX.SetAttribute(SGA_SHOWNAME, 0)
	SPGraphX.SetAttribute(SLA_QCOPTIONS, &H00008044) 'hide zero reference line
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