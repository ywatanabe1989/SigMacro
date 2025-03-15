Public Const ObjectHelp = Path + "\Ligand Binding.CHM"
Dim HelpID As Variant
Option Explicit
Dim QNAN As String
Dim QNB As String
Dim sL
Dim sD
Dim CurrentNotebook As Object
Dim FitFile As Object
Dim FitLibrary$
Dim DataSectionName              'name of section containing user's worksheet
Dim DataWorksheetName            'name of worksheet with user's data
Dim LastColumn As Long           'last column in user's data worksheet
Dim LastRow As Long              'last row in user's data worksheet
Dim GraphWorksheetName           'name of worksheet with graph(s)
Dim ReportWorksheetName          'name of report worksheet
Dim NumReplicates As Long        'number of replicates - user entered
Dim KdSave As String             'dialog Kd value for Ki computation
Dim ligandSave As String         'dialog ligand concentration for Ki computation
Dim MinXSave As String           'dialog minimum x for fit line
Dim MaxXSave As String           'dialog maximum x for fit line
Dim Text3Save As String          'information text for each equation
Dim NumDataGroups As Long        'number of groups of replicates - computed in CreateGraph
Dim NumParameters As Double      'number of parameters in fit equation
Dim ParameterNames() As Variant  'names of parameters in fit equation
Dim ParameterStartCol As Long    'first parameter column in graph worksheet
Dim EquationIndex As Long        'index for equation name array
Dim KiCheckBoxState As Boolean   'true if user selects to compute Ki
Dim XRangeCheckBoxState As Boolean 'false if user elects to enter fit line range values
Global Const ResultsStartCol As Long = 1 'first data column in graph worksheet, defined in CreateGraphWorksheet
Function FlagOn(flag As Long)
  FlagOn = flag Or FLAG_SET_BIT ' Use to set option flag bits on, leaving others unchanged
End Function
Function FlagOff(flag As Long)
  FlagOff = flag Or FLAG_CLEAR_BIT ' Use to set option flag bits off, leaving others unchanged
End Function
Sub Main
'This macro fits equations from a library to multiple sets of replicate data
'Dick Mitchell, 6/14/01
'Modified 4/12/06, removed Fitfile.Close in FitReplicateGroups
'Modified 11/27/07, set report column widths, formatted integers in report, 
'named worksheets and graph page, set fit line thickness and sent it to back,
'thicken x and y axis lines, extend lines to axis for certain equations, removed
'run-fit-once to get parameter values, positioned results windows
	HelpID = 1			' Help ID number

	sL = ListSeparator  'international list separator
	sD = DecimalSymbol  'international decimal symbol
	GetEmptyValues
	'Get current notebook and name of user's section and data worksheet
	Dim CurrentWorksheet As Object, CurrentSection As Object
	Set CurrentNotebook = ActiveDocument
	On Error GoTo NoOpenWorksheet
	Set CurrentWorksheet = CurrentNotebook.CurrentDataItem
	DataWorksheetName = CurrentWorksheet.Name
	On Error GoTo OtherError
	Dim i As Long
	For i = 0 To CurrentNotebook.NotebookItems.Count-1
		If CurrentNotebook.NotebookItems(i).Name = DataWorksheetName Then
			Set CurrentSection = CurrentNotebook.NotebookItems(i-1): Exit For
		End If
	Next i
	DataSectionName = CurrentSection.Name

	'Find last column and row in worksheet
	Dim DataWorksheetTable As Object
	Set DataWorksheetTable = CurrentWorksheet.DataTable
	LastColumn = 0
	LastRow = 0
	DataWorksheetTable.GetMaxUsedSize(LastColumn,LastRow)

	'Check if this is a results worksheet from a previous Ligand run
	If (DataWorksheetTable.Cell(0,0) = "Ligand Graph" And DataWorksheetTable.Cell(0,1) = "Worksheet") Or DataWorksheetTable.Cell(0,0) = "Parameter Values" Then
		HelpMsgBox HelpID,"The open worksheet is a previously created Ligand results worksheet.  Please open a worksheet containing your raw data.",vbExclamation, "Ligand Binding Macro"
		GoTo Finish
	End If

	'Check if this is a results worksheet from a previous Exploratory EK run
	If (DataWorksheetTable.Cell(0,0) = "EK Graph" And DataWorksheetTable.Cell(0,1) = "Worksheet") Or _
	(DataWorksheetTable.Cell(0,0) = "[Inhibitor]" And DataWorksheetTable.Cell(1,0) = "Kmapp") Then
		HelpMsgBox HelpID,"The open worksheet is a previously created Exploratory EK results worksheet.  Please open a Ligand worksheet containing your raw data.",vbExclamation, "Ligand Binding Macro"
		GoTo Finish
	End If

	'Check if worksheet contains strings
	'Change to blanks if they want to otherwise exit
	If CheckForStrings(False) = True Then
		If MsgBox("Your worksheet contains text items.  Would you like them to be converted to missing values and continue?", vbYesNo, "Ligand Binding Macro") = vbYes Then
			CheckForStrings(True)
		Else
			GoTo Finish
		End If
	End If

	'Determine the size of the 'cleaned' worksheet
	DataWorksheetTable.GetMaxUsedSize(LastColumn,LastRow)

	'Check for insufficient data
	If LastColumn = 0 And LastRow = 0 Then GoTo EmptyWorksheet
	If LastColumn < 2 Or (LastColumn > 1 And LastRow < 2) Then GoTo InsufficientData
	If LastNonMissingRow(1, LastRow) < 2 Then GoTo InsufficientData
	If LastNonMissingRow(2, LastRow) = 0 And LastColumn > 2 Then GoTo MissingFirstYColumn

	'Check for all zeros in x data
	Dim XData() As Variant
	ReDim XData(0,LastRow-1)
	Dim LastXRow As Long
	LastXRow = LastNonMissingRow(1,LastRow)
	XData = DataWorksheetTable.GetData(0,0,0,LastXRow-1)
	Dim ZeroCount As Long
	ZeroCount = 0
	For i = 0 To LastXRow-1
		If MissingValue(XData(0,i)) = True Or XData(0,i) = 0.0 Then
			ZeroCount = ZeroCount + 1
		End If
	Next i
	If ZeroCount = LastXRow Then GoTo TooManyZerosInX

	'Open the fit library
	Dim Equations$()
	'Defines the equation source.  Edit to use a different fit library
	FitLibrary = "Standard.jfl"
	Notebooks.Open(UserPath + "\" + FitLibrary, ".jfl")
	Set FitFile = Notebooks(UserPath + "\" + FitLibrary)
	FitFile.Visible=False

	'Populate equation list with all equation items in fit library
	Dim NotebookItemType() As Long
	ReDim NotebookItemType(FitFile.NotebookItems.Count)
	Dim SectionIndex As Integer, SectionTwoIndex As Integer
	SectionTwoIndex = -1
	Dim FoundSection As Boolean
	FoundSection = False
	Dim Index As Integer
	Index = 0
	Dim Item As Variant

	'find section indicies for requested section and the next section (if any)
	For Each Item In FitFile.NotebookItems
		If FoundSection = True And FitFile.NotebookItems(Index).ItemType = 3 Then
			SectionTwoIndex = Index
			Exit For
		End If
		If FitFile.NotebookItems(Index).Name = "Ligand Binding" Then
			SectionIndex = Index
			FoundSection = True
		End If
		Index = Index + 1
	Next Item
	If FoundSection = True And SectionTwoIndex = -1 Then SectionTwoIndex = Index  'if last section in notebook

	'Put item names in Equations$() array
	Index = 0
	For i = SectionIndex+1 To SectionTwoIndex-1
		If FitFile.NotebookItems(i).ItemType = CT_FIT Then
			ReDim Preserve Equations$(Index)
			Equations(Index) = FitFile.NotebookItems(i).Name
			Index = Index + 1
		End If
	Next i

	'Populate replicate array
	Dim NumReps$(9)
	For i = 0 To 9
		NumReps(i) = CStr(i+1)
	Next i

	'Bring the data worksheet to the front
	CurrentNotebook.Activate
	ActiveDocument.NotebookItems(DataWorksheetName).Open

	'Initial dialog values
	EquationIndex = 0
	KdSave = ""
	ligandSave = ""
	MinXSave = ""
	MaxXSave = ""
	XRangeCheckBoxState = True
	Text3Save = "Y increases with X."
	NumReplicates = 1
	Repeat:
	Begin Dialog UserDialog 630,232,"Ligand Binding",.DialogFunc '%GRID:10,7,1,0
		Text 10,5,95,14,"No. &Replicates",.Text1
		ComboBox 10,20,76,125,NumReps(),.NumberOfReps
		Text 114,5,60,14,"&Equation",.Text2
		ListBox 114,20,294,127,Equations(),.FitList
		OKButton 400,205,96,18
		CancelButton 522,205,96,18
		PushButton 10,205,96,19,"Help",.PushButton2
		GroupBox 420,5,200,92,"Ki from EC50",.KiGroupBox
		CheckBox 430,22,140,14,"Enter &Kd, [ligand]",.EnterKdCheckBox
		Text 430,46,60,16,"K&d = ",.tKd
		TextBox 525,43,72,18,.Kd
		Text 430,70,60,14,"[&ligand] = ",.tligand
		TextBox 525,68,72,18,.ligand
		GroupBox 10,150,403,45,"",.GroupBox1
		Text 20,161,380,26,"",.Text3
		GroupBox 420,103,200,92,"Fit line range",.FitLineGroupBox
		CheckBox 430,120,140,14,"&Automatic range",.XRangeCheckBox
		Text 430,143,80,14,"Mi&nimum X =",.tMinX
		TextBox 525,140,72,18,.MinX
		Text 430,169,90,14,"Ma&ximum X =",.tMaxX
		TextBox 525,166,72,18,.MaxX
	End Dialog
	Dim dlg As UserDialog
	dlg.NumberOfReps = "1"
	dlg.FitList = EquationIndex
	dlg.Kd = KdSave
	dlg.ligand = ligandSave
	dlg.MinX = MinXSave
	dlg.MaxX = MaxXSave
	dlg.XRangeCheckBox = XRangeCheckBoxState
	Select Case Dialog(dlg)
		Case 0   'handles Cancel button
			End
'		Case 1   'handles Help buttton
'		Help(ObjectHelp,HelpID)
'		GoTo Repeat
	End Select

	'Check for number of replicates errors
	'NumReplicates is obtained from DialogFunc
	If dlg.NumberOfReps = "" Then
		MsgBox "Enter an integer greater than zero for Number of Replicates",vbExclamation, "Ligand Binding Macro"
		NumReplicates = 1
		GoTo Repeat
    ElseIf IsNumeric(dlg.NumberOfReps) = False Then
	    MsgBox "Enter an integer greater than zero for Number of Replicates",vbExclamation, "Ligand Binding Macro"
		NumReplicates = 1
      	GoTo Repeat
	ElseIf CDbl(dlg.NumberOfReps) < 1 Then
	    MsgBox "Enter an integer greater than zero for Number of Replicates",vbExclamation, "Ligand Binding Macro"
		NumReplicates = 1
      	GoTo Repeat
    ElseIf CDbl(dlg.NumberOfReps)-Int(CDbl(dlg.NumberOfReps)) > 0 Then
	    MsgBox "Enter an integer greater than zero for Number of Replicates",vbExclamation, "Ligand Binding Macro"
		NumReplicates = 1
      	GoTo Repeat
    End If

	'Check for inconsistencey between selected number of replicates and the number of y data columns
	'we do not allow the last replicate y column in the last data group to be empty
	Dim ReplicateErrorCheck As Double
	ReplicateErrorCheck = (LastColumn - 1) Mod NumReplicates
	If ReplicateErrorCheck <> 0 Then
		MsgBox "The number of replicates (" + CStr(NumReplicates) + ") is inconsistent with the number of data columns (" + CStr(LastColumn-1) + ")." +vbCrLf+ _
		"The number of data columns should be equally divisable by the number of replicates.",vbExclamation,"Number of Replicates Error"
		GoTo Repeat
	End If

	'Check for Kd, [ligand] errors
	If dlg.EnterKdCheckBox = 1 Then
		If dlg.Kd = "" Or dlg.ligand = "" Then
			MsgBox "Please enter positive numbers for Kd and [ligand].",vbExclamation,"Data Entry Error"
			GoTo Repeat
	    ElseIf IsNumeric(dlg.Kd) = False Or IsNumeric(dlg.ligand) = False Then
			MsgBox "Please enter positive numbers for Kd and [ligand].",vbExclamation,"Data Entry Error"
			GoTo Repeat
		ElseIf CDbl(dlg.Kd) <= 0 Or CDbl(dlg.ligand) <= 0 Then
			MsgBox "Please enter positive numbers for Kd and [ligand].",vbExclamation,"Data Entry Error"
			GoTo Repeat
	    End If
	End If

	'Check for X range error
	If dlg.XRangeCheckBox = 0 Then
		If dlg.MinX = "" And dlg.MaxX = "" Then     'if no entries then X range = data range
			dlg.XRangeCheckBox = 1
		ElseIf dlg.MinX = "" Or dlg.MaxX = "" Then  'do nothing, use logic in graph subroutines
		Else
			If CDbl(dlg.MaxX) <= CDbl(dlg.MinX) Then
				MsgBox("Minimum X must be less than Maximum X",vbExclamation,"MinX,MaxX Entry Error")
				GoTo Repeat
			End If
		End If
	End If
	
	'Determine if equation is type for fit lines to start at x = 0
	Dim FitLineStartsAtXEqualsZero As Boolean
	FitLineStartsAtXEqualsZero = False
	Select Case Equations(dlg.FitList)
		Case "one site saturation", "two site saturation","one site saturation + nonspecific","two site saturation + nonspecific"
			FitLineStartsAtXEqualsZero = True
		Case Else
	End Select
	
	Dim GraphSectionName As Variant
	Dim GraphWorksheet As Object
	CreateGraphWorksheet GraphSectionName, GraphWorksheet
	Dim GraphPageName As String
	CreateDataGraph GraphPageName, dlg.XRangeCheckBox, dlg.MinX, dlg.MaxX, FitLineStartsAtXEqualsZero
	Dim XColumnString As String, YColumnString As String
	Dim ReportArray() As Variant
	ReDim ReportArray(NumDataGroups,1000)      'temporary dimension - gets reset for valid data sets in FitReplicateGroups
	'Will exit from FitReplicateGroups subroutine if too-few-data-points for all data groups
	Dim Offset3 As Long, Offset4 As Long        'offset rows in report
	FitReplicateGroups dlg.FitList, dlg.XRangeCheckBox, Equations, ReportArray, dlg.Kd, dlg.ligand, GraphPageName, GraphSectionName, FitLineStartsAtXEqualsZero, Offset3, Offset4
	CreateFinalGraph dlg.FitList, Equations, dlg.XRangeCheckBox, dlg.MinX, dlg.MaxX, FitLineStartsAtXEqualsZero
	CreateReportWorksheetAndReport ReportArray, Offset3, Offset4

	'Close data and report worksheets to leave graph in view
'	ActiveDocument.NotebookItems(ReportWorksheetName).Close(True) 'close the report worksheet
	ActiveDocument.NotebookItems(DataWorksheetName).Open          'open the data worksheet to allow closing it without a crash (to be fixed v8.0)
	ActiveDocument.NotebookItems(DataWorksheetName).Close(True)  'close user's worksheet
	GoTo Finish

	NoOpenWorksheet:
	HelpMsgBox HelpID, "You must have a worksheet open and in focus",vbExclamation,"No Open Worksheet"
	GoTo Finish

	EmptyWorksheet:
	HelpMsgBox HelpID, "You must have data in your worksheet.",vbExclamation,"Empty Worksheet"
	GoTo Finish

	OtherError:
	MsgBox(Err.Description + " (" + CStr(Err.Number) + ")" + " in Main subroutine", 16, "Ligand Binding Macro")
	GoTo Finish

	InsufficientData:
	HelpMsgBox HelpID, "There is insufficient data in your worksheet.", vbExclamation, "Insufficient Data"
	GoTo Finish

	MissingFirstYColumn:
	HelpMsgBox HelpID, "The first Y data column must be placed in column 2.", vbExclamation, "First Y Data Column Missing"
	GoTo Finish

	TooManyZerosInX:
	HelpMsgBox HelpID, "Your X data is empty or contains too many zeros.", vbExclamation, "Bad X Data"

	Finish:
End Sub
Sub CreateGraphWorksheet(ByRef GraphSectionName As Variant, ByRef GraphWorksheet As Object)
'Creates the graph worksheet and copies data from current worksheet.  Data in graph worksheet
'starts in column 2 for ResultsStartCol = 1 for Graph Worksheet title in column 1

	'Determine the data range and define the first empty column
	Dim DataWorksheetTable As Object
	Set DataWorksheetTable = CurrentNotebook.CurrentDataItem.DataTable
	'Compute the number of data groups
	NumDataGroups = Int((LastColumn - 1)/NumReplicates)
	'Copy data from the data worksheet
	Dim Data() As Variant
	ReDim Data(LastColumn-1, LastRow-1)
	Data = DataWorksheetTable.GetData(0, 0, LastColumn-1, LastRow-1)
	'Copy column titles from the data worksheet
	Dim Titles() As Variant
	ReDim Titles(LastColumn-1, 0)
	Titles = DataWorksheetTable.GetData(0, -1, LastColumn-1, -1)
	
	'Add a graph section and worksheet to the notebook
	Set GraphWorksheet = ActiveDocument.NotebookItems.Add(CT_WORKSHEET)
	GraphWorksheetName = ActiveDocument.CurrentDataItem.Name

	'Define notebook objects and get section name
	Dim CurrentNotebook As Object
	Set CurrentNotebook = ActiveDocument
	Dim CurrentSection As Object
	Dim i As Long
	For i = 0 To CurrentNotebook.NotebookItems.Count-1
		If CurrentNotebook.NotebookItems(i).Name = GraphWorksheetName Then
			Set CurrentSection = CurrentNotebook.NotebookItems(i-1): Exit For
		End If
	Next i

	'Name the section
	On Error GoTo ReplicateSectionNameError
	Dim iError As Long
	iError = 1
	NameSection: If iError = 1 Then
		GraphSectionName = "Graph for " + DataSectionName
	Else
		GraphSectionName = "Graph " + CStr(iError) + " for " + DataSectionName
	End If
	CurrentSection.Name = GraphSectionName
	
	'Name the worksheet
	On Error GoTo ReplicateWorksheetNameError
	iError = 1
	NameWorksheet: If iError = 1 Then
		GraphWorksheetName = "Graph worksheet"
	Else
		GraphWorksheetName = "Graph worksheet " + CStr(iError)
	End If
	GraphWorksheet.Name = GraphWorksheetName
	
	'Put data and column titles from the data worksheet into the graph worksheet starting in column ResultsStartCol
	On Error GoTo PutDataError  'need this?  TBD
	For i = 0 To LastColumn-1
		If Titles(i,0) = QNAN Then
			GraphWorksheet.DataTable.Cell(ResultsStartCol+i, -1) = ""
		Else
			GraphWorksheet.DataTable.Cell(ResultsStartCol+i, -1) = Titles(i,0)
		End If
	Next i
	GraphWorksheet.DataTable.PutData(Data, ResultsStartCol, 0)

	'Change missing values to blanks to clean up worksheet and allow transform 'ape' function to detect correct x length when y length is larger.
	Dim j As Long
	For i = 1 To LastColumn
		For j = 0 To LastRow-1
			If GraphWorksheet.DataTable.Cell(i,j) = QNAN Then GraphWorksheet.DataTable.Cell(i,j) = ""
		Next j
	Next i

	'Name the worksheet in column 1.  Use to disallow running on a results worksheet.
	GraphWorksheet.DataTable.Cell(0,0) = "Ligand Graph"
	GraphWorksheet.DataTable.Cell(0,1) = "Worksheet"
	GoTo Finish

	ReplicateSectionNameError:
	Select Case Err.Number	'evaluate error number.
		Case 65535	        'duplicate section name
			iError = iError + 1
		Case Else           'handle other situations here
			MsgBox(Err.Description + " (" + CStr(Err.Number) + ")" + " in CreateGraphWorksheet subroutine", 16, "Ligand Binding Macro")
	End Select
	Resume NameSection
	
	ReplicateWorksheetNameError:
	Select Case Err.Number	'evaluate error number.
		Case 65535	        'duplicate worksheet name
			iError = iError + 1
		Case Else           'handle other situations here
			MsgBox(Err.Description + " (" + CStr(Err.Number) + ")" + " in CreateGraphWorksheet subroutine", 16, "Ligand Binding Macro")
	End Select
	Resume NameWorksheet
	
	PutDataError:
	MsgBox (Err.Description)
	Resume Next

	Finish:
End Sub
Sub FitReplicateGroups(ByVal FitIndex As Long, ByVal XRangeCheckBox As Long, _
ByRef Equations As Variant, ByRef ReportArray() As Variant, ByVal Kd As String, _
ByVal ligand As String, ByVal GraphPageName As String, ByVal GraphSectionName As Variant, _
ByVal FitLineStartsAtXEqualsZero As Boolean, ByRef Offset3 As Long, ByRef Offset4 As Long)
'Fits the selected equation to multiple groups of replicate data and puts results into the ReportArray

	Dim SPPage, SPGraph, SPPlot
	Set SPPage = CurrentNotebook.NotebookItems(GraphPageName)
	Set SPGraph = SPPage.GraphPages(0).CurrentPageObject(GPT_GRAPH)
	Set SPPlot = SPGraph.Plots(0)
	Dim FitEquationName$
	FitEquationName = Equations(FitIndex)
	FitFile.Visible=False
	Dim FitObject As Object
	Set FitObject = FitFile.NotebookItems(FitEquationName)
	SPPlot.ChildObjects(0).SelectObject   'select first tuple
	FitObject.Open
	'FitObject.Visible = False
	ParameterNames = FitObject.Parameter
	Dim NumParameters As Integer		
	NumParameters = UBound(ParameterNames)+1

	Dim GraphWorksheetTable As Object
	Set GraphWorksheetTable = ActiveDocument.NotebookItems(GraphWorksheetName).DataTable
	Dim XData() As Variant, YData() As Variant
	Dim NumXRows As Long
	NumXRows = LastNonMissingRow(ResultsStartCol+1,LastRow)
	XData = GraphWorksheetTable.GetData(ResultsStartCol,0,ResultsStartCol,NumXRows-1)
	Dim xFirstRow As Long, xLastRow As Long
	FirstAndLastRow xFirstRow, xLastRow, XData(), NumXRows, 1
	Dim XColumn As Long, YColumn As Long
	Dim FirstValidFit As Boolean
	FirstValidFit = True                             'used to write titles in report
	Dim NumGoodnessOfFitItems As Integer
	NumGoodnessOfFitItems = 6                        'number of rows in Goodness of Fit section (including title)
	Dim NumDataItems As Integer
	NumDataItems = 5                                 'number of rows in Data section (including title)

	'data group loop
	Dim NumSuccessiveBadDataSets As Integer          'determines if all data groups have insufficient data
	NumSuccessiveBadDataSets = 0
	Dim i As Long
	For i = 1 To NumDataGroups
		Dim GroupStartCol As Long
		GroupStartCol = ResultsStartCol+1+(i-1)*NumReplicates
		YData = GraphWorksheetTable.GetData(GroupStartCol, 0, GroupStartCol+NumReplicates-1, NumXRows-1)
		Dim yFirstRow As Long, yLastRow As Long
		FirstAndLastRow yFirstRow, yLastRow, YData(), NumXRows, NumReplicates
		
		'Substitute nonoverlapping X data with blanks
		Dim xFirstTemp() As Variant, xLastTemp() As Variant
		ReDim xFirstTemp(xLastRow-1)
		ReDim xLastTemp(xLastRow-1)
		SubstituteAndReplaceXs "Substitute", XData(), xFirstRow, xLastRow, yFirstRow, yLastRow, xFirstTemp(), xLastTemp(), GraphWorksheetTable
		
		Dim NumValidRows As Long, TotalNumMissingValues As Long
		NumberValidRows XData(), YData(), NumXRows, TotalNumMissingValues, NumValidRows
		Dim NotEnoughData As Boolean
	
		'test for sufficient data for the selected equation (num rows >= num parameters + 1)
		If NumValidRows > NumParameters Then
			NumSuccessiveBadDataSets = 0
			NotEnoughData = False
			SPPlot.ChildObjects(i-1).SelectObject   'select each tuple
			FitObject.Open
			FitObject.DatasetType = CF_XREPY
			FitObject.Option("iterations") = 200
			FitObject.Run
			FitObject.OutputReport = False
			FitObject.OutputEquation = False
			FitObject.ResidualsColumn = -2  'none
			FitObject.PredictedColumn = -2  'none
			FitObject.ParametersColumn = LastColumn + ResultsStartCol + i - 1
			
			'Extend fit to axes if user selects axis range or if equation has fit line that starts at x = 0.0
			If XRangeCheckBox = 0 Or FitLineStartsAtXEqualsZero = True Then  'extend fit to axes
				FitObject.ExtendFitToAxes = True	
			Else                                                             'use automatic X axis range
				FitObject.ExtendFitToAxes = False
			End If

			'XY Pairs
			FitObject.XColumn = LastColumn + ResultsStartCol + NumDataGroups +(i-1)*2
			FitObject.YColumn = LastColumn + ResultsStartCol + NumDataGroups + (i-1)*2+1
			FitObject.ZColumn = -2
			FitObject.OutputGraph = False
			FitObject.OutputAddPlot = True
			FitObject.AddPlotGraphIndex = 0

			'Get curve fit statistics for report and store
			Dim Stats As Object
			Set Stats =  FitObject.FitResults
			'Get parameter names, names with "log" prefix and associated name array index
			If NotEnoughData = False And FirstValidFit = True Then
				FirstValidFit = False
				ParameterNames = FitObject.Parameter
				Dim PostLogParameterName() As String
				Dim LogParameterIndex() As Long
				Dim j As Long
				Dim NumLogParameters As Long
				NumLogParameters = 0
				For j = 0 To NumParameters-1
					If Left(ParameterNames(j),3)="log" Then
						ReDim Preserve PostLogParameterName(NumLogParameters)
						PostLogParameterName(NumLogParameters) = Right(ParameterNames(j), Len(ParameterNames(j))-3)
						ReDim Preserve LogParameterIndex(NumLogParameters)
						LogParameterIndex(NumLogParameters)=j
						NumLogParameters = NumLogParameters + 1
					End If
				Next j

				'Generate first column of report
				'Determine offsets for equations with parameters with "log" prefixes and Ki values
'				Dim Offset1 As Long, Offset2 As Long, Offset3 As Long
				Dim Offset1 As Long, Offset2 As Long
				Dim delta1 As Integer, delta2 As Integer
				'Adjust delta for computation of Ki
				If KiCheckBoxState = False Then
					delta1 = NumLogParameters + 1
					delta2 = 0
				Else
					delta1 = 2*NumLogParameters + 1
					delta2 = 2
				End If
				Offset1 = NumParameters + delta1 + delta2 + 1    'offset to std. error title row
				Offset2 = Offset1 + NumParameters + 1            'offset to 95%CI title row
				Offset3 = Offset2 + NumParameters + delta1       'offset to Goodness of Fit title row
				Offset4 = Offset3 + NumGoodnessOfFitItems        'offset to Data title row
				ReDim Preserve ReportArray(NumDataGroups, Offset4+NumDataItems-2)
				'Put report section titles in array
				ReportArray(0,0) = "Parameter Values"
				ReportArray(0, Offset1-1) = "Std. Errors"
				ReportArray(0, Offset2-1) = "95% Confidence Intervals"
				ReportArray(0, Offset3-1) = "Goodness of Fit"
				ReportArray(0, Offset4-1) = "Data"
				For j = 0 To NumParameters-1
					ReportArray(0,j+1) = "   " + ParameterNames(j)
					ReportArray(0,Offset1+j) = "   " +  ParameterNames(j)
					ReportArray(0,Offset2+j) = "   " +  ParameterNames(j)
				Next j
				'Add post-"log" parameter names as row titles (typically EC50)
				If NumLogParameters > 0 Then
					For j = 0 To NumLogParameters-1
						ReportArray(0, NumParameters+1+j) = "   " + PostLogParameterName(j)
						ReportArray(0, Offset2+NumParameters+j) = "   " + PostLogParameterName(j)
					Next j
				End If
				'Add Ki, Kd, [ligand] titles
				If KiCheckBoxState = True Then
					If NumLogParameters > 1 Then
						Dim k As Long
						For k = 1 To NumLogParameters
							ReportArray(0, NumParameters+NumLogParameters+k)="   Ki" + CStr(k)
							ReportArray(0, Offset2+NumParameters+NumLogParameters+k-1)="   Ki" + CStr(k)
						Next k
					Else
						ReportArray(0, NumParameters+NumLogParameters+1)="   Ki"
						ReportArray(0, Offset2+NumParameters+NumLogParameters)="   Ki"
					End If

					ReportArray(0, NumParameters+2*NumLogParameters+1)="   Kd"
					ReportArray(0, NumParameters+2*NumLogParameters+2)="   [ligand]"
				End If
				ReportArray(0,Offset3) = "   Degrees of Freedom"
				ReportArray(0,Offset3+1) = "   R2"
				ReportArray(0,Offset3+2) = "   Residual Sum of Squares"
				ReportArray(0,Offset3+3) = "   Sy.x"
				ReportArray(0,Offset3+4) = "   Fit Status"
				ReportArray(0,Offset4) = "   Number of X Values"
				ReportArray(0,Offset4+1) = "   Number of Y Replicates"
				ReportArray(0,Offset4+2) = "   Total Number of Y Values"
				ReportArray(0,Offset4+3) = "   Number of Missing Values"
			End If
			'Put parameter values, std. errors, 95% CI's in ReportArray
			Dim NumDataPoints As Long
			NumDataPoints = CLng(Stats.DataPointCount)
			Dim ParameterStdError As Double
			Dim ParameterValue
			For j = 0 To NumParameters-1
				ParameterValue = FitObject.FittedParameterValue(ParameterNames(j))
				ReportArray(i,j+1) = ParameterValue
				ParameterStdError = Stats.ParameterStandardError(ParameterNames(j))
				ReportArray(i,Offset1+j) = ParameterStdError
				If NumDataPoints > NumParameters Then
					ReportArray(i,Offset2+j) = FormatNumber(ParameterValue - TValue(NumDataPoints,NumParameters)*ParameterStdError) + " to " + FormatNumber(ParameterValue + TValue(NumDataPoints,NumParameters)*ParameterStdError)
				Else
					ReportArray(i,Offset2+j) = ""
				End If
			Next j
			'Add EC50 and Ki rows
			Dim Param As Double, ParamSE As Double
			Dim EC50 As String
			Dim CIlow As String, CIhigh As String
			If KiCheckBoxState = True Then
				Dim ChengPrusoffDenom As Double
				ChengPrusoffDenom = 1 + CDbl(ligand)/CDbl(Kd)
			End If
			For j = 0 To NumLogParameters-1
				Param = ReportArray(i, LogParameterIndex(j) + 1)
				EC50 = FormatNumber(AntiLog(Param))
				ReportArray(i, NumParameters+1+j) = EC50
				If NumDataPoints > NumParameters Then
					ParamSE = ReportArray(i, Offset1+LogParameterIndex(j))
					CIlow = FormatNumber(AntiLog(Param - TValue(NumDataPoints,NumParameters)*ParamSE))
					CIhigh = FormatNumber(AntiLog(Param + TValue(NumDataPoints,NumParameters)*ParamSE))
					ReportArray(i, Offset2+NumParameters+j) = CIlow + " to " + CIhigh
				Else
					ReportArray(i, Offset2+NumParameters+j) = ""
				End If
				Dim EC50Report As Variant, CIlowReport As Variant, CIhighReport As Variant
				If KiCheckBoxState = True Then
					If EC50 = "+Inf" Then
						EC50Report = "+Inf"
					Else
						EC50Report = FormatNumber(CDbl(EC50)/ChengPrusoffDenom)
					End If
					If CIlow = "+Inf" Then
						CIlowReport = "+Inf"
					Else
						CIlowReport = FormatNumber(CDbl(CIlow)/ChengPrusoffDenom)
					End If
					If CIhigh = "+Inf" Then
						CIhighReport = "+Inf"
					Else
						CIhighReport = FormatNumber(CDbl(CIhigh)/ChengPrusoffDenom)
					End If
					ReportArray(i,NumParameters+NumLogParameters+1+j) = EC50Report
					ReportArray(i, Offset2+NumParameters+NumLogParameters+j) = CIlowReport + " to " + CIhighReport
				End If
			Next j
			'Add Kd and [ligand] rows
			If KiCheckBoxState = True Then
				ReportArray(i, NumParameters+2*NumLogParameters+1) = Kd
				ReportArray(i, NumParameters+2*NumLogParameters+2) = ligand
			End If
			'Goodness of fit results
			'get fit verdicts
			Dim FitVerdict1 As Double, FitVerdict2 As Double, FitVerdict3 As Double
			FitVerdict1 = Stats.FitVerdict(0)
			FitVerdict2 = Stats.FitVerdict(1)
			FitVerdict3 = Stats.FitVerdict(2)
			ReportArray(i,Offset3) = Stats.ResidualDegreesOfFreedom
			ReportArray(i,Offset3+1) = Stats.RSquare
			ReportArray(i,Offset3+2) = Stats.ResidualSumOfSquares
			ReportArray(i,Offset3+3) = Stats.StandardErrorOfEstimate
			Dim FitFailureFlag As Boolean
			FitFailureFlag = False
			'Use simple fitverdict logic for now
			Select Case FitVerdict1
				Case 3,4
				ReportArray(i,Offset3+4) = "Converged"
				Case 6,7
				ReportArray(i,Offset3+4) = "Fit algorithm failure"
				FitFailureFlag = True
				Case Else
				ReportArray(i,Offset3+4) = "Did not converge"
			End Select
			'Data results
			ReportArray(i,Offset4) = NumValidRows
			ReportArray(i,Offset4+1) = NumReplicates
			ReportArray(i,Offset4+2) = NumDataPoints
			ReportArray(i,Offset4+3) = TotalNumMissingValues
			'Put blanks in title rows for valid data groups
			For j = i To NumDataGroups
				ReportArray(j,0) = ""
				ReportArray(j,Offset1-1) = ""
				ReportArray(j,Offset2-1) = ""
				ReportArray(j,Offset3-1) = ""
				ReportArray(j,Offset4-1) = ""
			Next j

			'Trap bad fit errors
	'		If FitVerdict1 = 6 Or FitVerdict1 = 7 Then
	'			MsgBox "FitVerdict =" + CStr(FitVerdict1)+ "  DataGroup = " + CStr(i)
	'		End If
			If FitFailureFlag = False Then
				FitObject.Finish
			End If
			Wait 0.01  'needed to prevent crash for two-site comp. and 8 group data set, will be fixed in 7.0+
			FitObject.Close(False)
		Else
			'insufficient data for this data group
			NumSuccessiveBadDataSets = NumSuccessiveBadDataSets + 1
			'test to see if insufficient data for all data groups
			If NumSuccessiveBadDataSets = NumDataGroups Then
				ActiveDocument.NotebookItems.Delete(GraphSectionName)  'delete the graph worksheet section
				FitFile.Close(False)                                   'Close Ligand.jfl
				Set FitFile = Nothing
				HelpMsgBox HelpID, "There are too few data points for this equation.  You need one more data point than the number of parameters in the equation.", vbExclamation, "Not Enough Data"
			Else
				ReportArray(i,0) = "Not enough data"
				If Offset4 = 0 Then Offset4 = 50  'arbitrarily large value if it never gets set
				For j = 1 To Offset4+NumDataItems-2
					ReportArray(i,j)=""
				Next j
			End If
		End If  'end of 'sufficient data' test
		
		'Replace blanks with original data
		SubstituteAndReplaceXs "Replace", XData(), xFirstRow, xLastRow, yFirstRow, yLastRow,xFirstTemp(), xLastTemp(), GraphWorksheetTable
	Next i      'data group loop	

'	FitFile.Close(False)  'Close Ligand.jfl, removed 4/12/06
	Set FitFile = Nothing
	ActiveDocument.NotebookItems.Delete(GraphPageName)   'delete graph to avoid legend bug (to be fixed in V8.0) TBD
End Sub
Sub XAndYConcatenatedStrings(ByRef XColumnString As String, ByRef YColumnString As String, _
ByVal XColumn As Long, ByVal YColumn As Long)
'Creates X Rep Y concatenated variable strings for the curve fitter
	'Repeat X for each Y column
	'x = {col(1), col(1), col(1)...}
	'Truncate Y columns to size of X column and concatenate
	'y = {col(2,1,size(col(1))), col(3,1,size(col(1))), col(4,1,size(col(1)))...}
	If NumReplicates > 1 Then
		XColumnString = "{col(" + CStr(XColumn) + ")"
		YColumnString = "{col(" + CStr(YColumn) + sL + "1" + sL + "size(col(" + CStr(XColumn) + ")))"
		Dim s As String
		Dim i As Long
		For i = 1 To NumReplicates-1
			s = sL + "col(" + CStr(XColumn) + ")"
			XColumnString = XColumnString + s
			s = sL + "col(" + CStr(YColumn+i) + sL + "1" + sL + "size(col(" + CStr(XColumn) + ")))"
			YColumnString = YColumnString + s
		Next i
		XColumnString = XColumnString + "}"
		YColumnString = YColumnString + "}"
	Else
		XColumnString = "col(" + CStr(XColumn) + ")"
		YColumnString = "col(" + CStr(YColumn) + ")"
	End If
End Sub
Sub CreateDataGraph(ByRef GraphPageName As String, ByVal XRangeCheckBox As Long, ByVal MinX As String, _
ByVal MaxX As String, ByVal FitLineStartsAtXEqualsZero As Boolean)
'Creates X Many Rep Y graph of data in GraphWorksheet

	ActiveDocument.NotebookItems.Add(CT_GRAPHICPAGE)
	'Create X Many Rep Y graph of data
	Dim ColumnsPerPlot()
	ReDim ColumnsPerPlot(2, 2*NumDataGroups)
	ColumnsPerPlot(0, 0) = ResultsStartCol
	ColumnsPerPlot(1, 0) = 0
'	ColumnsPerPlot(2, 0) = 2147483647
	ColumnsPerPlot(2, 0) = LastRow-1
	Dim i As Long
	For i = 1 To NumDataGroups
		ColumnsPerPlot(0, 2*i-1) = ResultsStartCol + 1 + CVar(NumReplicates*(i-1))
		ColumnsPerPlot(1, 2*i-1) = 0
		ColumnsPerPlot(2, 2*i-1) = 2147483647
		ColumnsPerPlot(0, 2*i) = ResultsStartCol +CVar(NumReplicates*i)
		ColumnsPerPlot(1, 2*i) = 0
'		ColumnsPerPlot(2, 2*i) = 2147483647
		ColumnsPerPlot(2, 2*i) = LastRow-1
	Next i
	Dim PlotColumnCountArray()
	ReDim PlotColumnCountArray(0)
	PlotColumnCountArray(0) = 1 + CVar(2*NumDataGroups)
	Dim SPPage As Object
	Set SPPage = ActiveDocument.CurrentPageItem
'	SPPage.Visible = False	
	
	SPPage.CreateWizardGraph("Scatter Plot", "Multiple Error Bars", "X, Many Y Replicates", ColumnsPerPlot, PlotColumnCountArray, "Row Means", "Standard Error", , , , , "Standard Error", True)
	'Get graph name for use in FitReplicateGroups
	GraphPageName = SPPage.Name

	'Set X axis range if XRangeCheckBox unchecked and if XRangeCheckBox is checked and fit line should start at zero
	Dim SPGraph, XAxis, YAxis
	Set SPGraph = SPPage.GraphPages(0).Graphs(0)
	Set XAxis=SPGraph.Axes(0)
	Set YAxis=SPGraph.Axes(1)
	
	'Change the axis range extents if
	'   1) user selects end points
	'   2) equation is the type with fit line that starts at x = 0.0
	If XRangeCheckBox = 1 And FitLineStartsAtXEqualsZero = False Then
	Else
		Dim GraphWorksheetTable As Object
		Set GraphWorksheetTable = ActiveDocument.NotebookItems(GraphWorksheetName).DataTable
		Dim XData() As Variant
		Dim NumXRows As Long
		NumXRows = LastNonMissingRow(ResultsStartCol+1,LastRow)
		XData = GraphWorksheetTable.GetData(ResultsStartCol,0,ResultsStartCol,NumXRows-1)
		Dim XDataMin As Double, XDataMax As Double
		XDataMin = min_array(XData,0,NumXRows-1)
		XDataMax = max_array(XData,0,NumXRows-1)
		'set X range values
		Dim FromVal As Variant
		XAxis.GetAttribute(SAA_FROMVAL, FromVal)                             'validate the graph to fix FROMVAL
		XAxis.SetAttribute(SAA_OPTIONS, FlagOff(SAA_FLAG_AUTORANGE))         'turn off autorange
		
		If XRangeCheckBox = 1 And FitLineStartsAtXEqualsZero = True Then
			XAxis.SetAttribute(SAA_TOVAL, XDataMax)                             'set max x range to max of X data
		Else
			If MinX <> "" Then
				XAxis.SetAttribute(SAA_FROMVAL, MinX)                            'set min x range to MinX
			Else
				XAxis.SetAttribute(SAA_FROMVAL, XDataMin)                        'set min x range to min of X data
			End If
			If MaxX <> "" Then
				XAxis.SetAttribute(SAA_TOVAL, MaxX)                              'set min x range to MaxX
			Else
				XAxis.SetAttribute(SAA_TOVAL, XDataMax)                          'set max x range to max of X data
			End If
		End If
	End If
End Sub
Sub CreateFinalGraph(ByVal FitIndex As Long, ByRef Equations As Variant, ByVal XRangeCheckBox As Long, _
ByVal MinX As String, ByVal MaxX As String, ByVal FitLineStartsAtXEqualsZero As Boolean)
'Creates X Many Rep Y graph of data in GraphWorksheet and an X Many Y graph of fit lines

	Dim GraphWorksheetTable As Object
	Set GraphWorksheetTable = ActiveDocument.NotebookItems(GraphWorksheetName).DataTable
	ActiveDocument.NotebookItems.Add(CT_GRAPHICPAGE)
	
	'Create X Many Rep Y graph of data
	Dim ColumnsPerPlot()
	ReDim ColumnsPerPlot(2, 2*NumDataGroups)
	ColumnsPerPlot(0, 0) = ResultsStartCol
	ColumnsPerPlot(1, 0) = 0
'	ColumnsPerPlot(2, 0) = 2147483647
	ColumnsPerPlot(2, 0) = LastRow-1
	Dim i As Long
	For i = 1 To NumDataGroups
		ColumnsPerPlot(0, 2*i-1) = ResultsStartCol + 1 + CVar(NumReplicates*(i-1))
		ColumnsPerPlot(1, 2*i-1) = 0
'		ColumnsPerPlot(2, 2*i-1) = 2147483647
		ColumnsPerPlot(2, 2*i-1) = LastRow-1
		ColumnsPerPlot(0, 2*i) = ResultsStartCol +CVar(NumReplicates*i)
		ColumnsPerPlot(1, 2*i) = 0
'		ColumnsPerPlot(2, 2*i) = 2147483647
		ColumnsPerPlot(2, 2*i) = LastRow-1
	Next i
	Dim PlotColumnCountArray()
	ReDim PlotColumnCountArray(0)
	PlotColumnCountArray(0) = 1 + CVar(2*NumDataGroups)
	Dim SPPage As Object
	Set SPPage = ActiveDocument.CurrentPageItem
	
	'Name the page
	On Error GoTo ReplicatePageNameError
	Dim iError As Long
	iError = 1
	Dim PageName As String
	NamePage: If iError = 1 Then
		PageName = "Data and fit graph"
	Else
		PageName = "Data and fit graph " + CStr(iError)
	End If
	SPPage.Name = PageName
	
	'Create the scatterplot of the data
	SPPage.CreateWizardGraph("Scatter Plot", "Multiple Error Bars", "X, Many Y Replicates", ColumnsPerPlot, PlotColumnCountArray, "Row Means", "Standard Error", , , , , "Standard Error", True)

	'Add XY Pairs fit line plot
	ReDim ColumnsPerPlot(2, 2*NumDataGroups-1)
	ParameterStartCol = ResultsStartCol + NumDataGroups*NumReplicates + 1
	Dim FitLineStartCol As Long      'first fit line column in graph worksheet
	FitLineStartCol = ParameterStartCol + NumDataGroups
	For i = 0 To 2*NumDataGroups-1
		ColumnsPerPlot(0, i) = FitLineStartCol + i
		ColumnsPerPlot(1, i) = 0
		ColumnsPerPlot(2, i) = 256
	Next i
	ReDim PlotColumnCountArray(0)
	PlotColumnCountArray(0) = 2*NumDataGroups
	SPPage.AddWizardPlot("Line Plot", "Multiple Straight Lines", "XY Pairs", ColumnsPerPlot, PlotColumnCountArray, , , , , , , , True)

	'Change fit lines to solid, width = 0.025 and color black
	Dim SPPlot As Object
	Set SPPlot = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Plots(1)
	SPPlot.SetAttribute(SEA_LINETYPE, SEA_LINE_SOLID)
	SPPlot.SetAttribute(SEA_TYPEREPEAT, 2)
	SPPlot.SetAttribute(SEA_THICKNESS, 25)
'ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLOR, &H00000000&)
'ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLORREPEAT, &H00000002&)
SPPlot.SetAttribute(SEA_COLOR, &H00000000&)
SPPlot.SetAttribute(SEA_COLORREPEAT, &H00000002&)
	
	'Send fit lines to back
	SPPlot.SelectObject
	SPPlot.SetAttribute(SPA_SENDTOBACK) 'send tuple to back

	'Delete the right y axis and top x axis
	Dim SPGraph, XAxis, YAxis
	Set SPGraph = SPPage.GraphPages(0).Graphs(0)
	Set XAxis=SPGraph.Axes(0)
	Set YAxis=SPGraph.Axes(1)
	XAxis.SetAttribute(SAA_SUB2OPTIONS, SAA_SUB_SHOW Or SAA_SUB_SHOWLINE Or FLAG_CLEAR_BIT)
	YAxis.SetAttribute(SAA_SUB2OPTIONS, SAA_SUB_SHOW Or SAA_SUB_SHOWLINE Or FLAG_CLEAR_BIT)

	'Set X and Y axis ranges
	Dim FromVal As Variant
	XAxis.GetAttribute(SAA_FROMVAL, FromVal)                        'validate the graph to fix FROMVAL
'	XAxis.GetAttribute(SAA_TOVAL, ToVal)                            'get the max axis range extent		
	If XRangeCheckBox = 1 Or FitLineStartsAtXEqualsZero = True Then
		If FitLineStartsAtXEqualsZero = True Then
			FromVal = 0.0
			XAxis.SetAttribute(SAA_OPTIONS, FlagOff(SAA_FLAG_AUTORANGE))    'turn off autorange
			XAxis.SetAttribute(SAA_FROMVAL, FromVal)                        'set max x range to extent of axis
		End If
	Else
		XAxis.SetAttribute(SAA_OPTIONS, FlagOff(SAA_FLAG_AUTORANGE))    'turn off autorange
		If MinX <> "" Then
			XAxis.SetAttribute(SAA_FROMVAL, MinX)                       'set min x range to MinX
		End If
		If MaxX <> "" Then
			XAxis.SetAttribute(SAA_TOVAL, MaxX)                         'set min x range to MaxX
		End If
	End If

	'set Y min to zero
	YAxis.GetAttribute(SAA_FROMVAL, FromVal)                            'validate the graph to fix FROMVAL
	YAxis.SetAttribute(SAA_OPTIONS, FlagOff(SAA_FLAG_AUTORANGE))        'turn off autorange
	YAxis.SetAttribute(SAA_FROMVAL, 0)                                  'set min y range = 0
	
	'Thicken the x and y axis lines to 0.015 in
	XAxis.SetAttribute(SAA_SELECTLINE, 1)
	XAxis.SetAttribute(SGA_FLAGS, &H00010001&)
	XAxis.SetAttribute(SEA_THICKNESS, 15)
	YAxis.SetAttribute(SAA_SELECTLINE, 1)
	YAxis.SetAttribute(SGA_FLAGS, &H00010001&)
	YAxis.SetAttribute(SEA_THICKNESS, 15)


	'Title X axis from worksheet column title if it exists
	Dim DataWorksheetTable As Object
	Set DataWorksheetTable = CurrentNotebook.NotebookItems(DataWorksheetName).DataTable
	Dim XAxisTitle
	XAxisTitle = DataWorksheetTable.Cell(0, -1)
	If XAxisTitle <> QNAN Then
		XAxis.Name = XAxisTitle
	End If

	'Create legend
	Dim SPLegend As Object
	Set SPLegend = SPGraph.AutoLegend
	Dim LegendTitle
	For i = 1 To NumDataGroups
		LegendTitle = DataWorksheetTable.Cell(NumReplicates*(i-1)+1, -1)
		If LegendTitle = QNAN Then
			SPLegend.ChildObjects(i+1).Name = "Data " + CStr(i+1)
		Else
			SPLegend.ChildObjects(i+1).Name = LegendTitle
		End If
	Next i
	
	'Turn off legend box outline
	SPGraph.ChildObjects(0).SetAttribute(SGA_FLAGS, FlagOff(SGA_FLAG_AUTOLEGENDBOX))
	
	'Turn off fit line legend item childobject(1) - the fit line
	SPLegend.ChildObjects(1).SetAttribute(STA_OPTIONS, FlagOff(STA_FLAG_VISIBLE))

	'Position legend
	Dim Pos()
	ReDim Pos(1)
	Pos(0) = 1000
	Pos(1) = 3000
	SPLegend.SetAttribute(SOA_POSEX, Pos)

	'Turn off Graph title
	SPGraph.SetAttribute(SGA_SHOWNAME, 0)
	
	'Position graph worksheet and graph
	Dim GraphWorksheet As Object
	Set GraphWorksheet = ActiveDocument.NotebookItems(GraphWorksheetName)
	GraphWorksheet.Left = 450
	GraphWorksheet.Top = 105
	GraphWorksheet.Height = 400
	GraphWorksheet.Width = 500

	SPPage.Left = 450
	SPPage.Top = 105
	SPPage.Height = 400
	SPPage.Width = 500
	GoTo Finish
	
	ReplicatePageNameError:
	Select Case Err.Number	'evaluate error number.
		Case 65535	        'duplicate page name
			iError = iError + 1
		Case Else           'handle other situations here
			MsgBox(Err.Description + " (" + CStr(Err.Number) + ")" + " in CreateReportWorksheetAndReport subroutine", 16, "Ligand Binding Macro")
	End Select
	Resume NamePage
	
	Finish:	
End Sub
Sub CreateReportWorksheetAndReport(ByRef ReportArray() As Variant, _
ByVal Offset3 As Long, ByVal Offset4 As Long)
'Create a worksheet for the report and write the ReportArray to it.

	'Create report section
	Dim ReportSection
	Set ReportSection = ActiveDocument.NotebookItems.Add(CT_FOLDER)
	ReportSection.Open     'make current object so worksheet is added to this section
	'Name the section
	On Error GoTo ReplicateSectionNameError
	Dim iError As Long
	iError = 1
	Dim ReportSectionName As String
	NameSection: If iError = 1 Then
		ReportSectionName = "Fit results for " + DataSectionName
	Else
		ReportSectionName = "Fit results " + CStr(iError) + " for " + DataSectionName
	End If
	ReportSection.Name = ReportSectionName
	'Add a worksheet to the section
	Dim ReportWorksheet
	Set ReportWorksheet = ActiveDocument.NotebookItems.Add(CT_WORKSHEET)
	'Name the worksheet
	On Error GoTo ReplicateWorksheetNameError
	iError = 1
	Dim ReportWorksheetName As String
	NameWorksheet: If iError = 1 Then
		ReportWorksheetName = "Results table"
	Else
		ReportWorksheetName = "Results table " + CStr(iError)
	End If
	ReportWorksheet.Name = ReportWorksheetName
	
	'Write array to worksheet
	ActiveDocument.CurrentDataItem.DataTable.PutData(ReportArray,0,0)
	
	'Set the degrees of freedome number format to integer
    Dim Selection(3)
	Dim i As Long
    For i = 1 To NumDataGroups
	    Selection(0) = i
	    Selection(1) = Offset3
	    Selection(2) = i
	    Selection(3) = Offset3
	    ReportWorksheet.SelectionExtent = Selection
	    ReportWorksheet.NumberFormat = "0"
    Next i

	'Set the Data section number formats to integer
    For i = 1 To NumDataGroups
	    Selection(0) = i
	    Selection(1) = Offset4
	    Selection(2) = i
	    Selection(3) = Offset4+3
	    ReportWorksheet.SelectionExtent = Selection
	    ReportWorksheet.NumberFormat = "0"
	Next i
	
	'Put cursor in 0,0 cell
    Selection(0) = 0
    Selection(1) = 0
    Selection(2) = 0
    Selection(3) = 0
	ReportWorksheet.SelectionExtent = Selection
	
	'Copy column titles from data worksheet to report worksheet
	Dim DataWorksheetTable As Object
	Set DataWorksheetTable = ActiveDocument.NotebookItems(DataWorksheetName).DataTable
	For i = 1 To NumDataGroups
'		If MissingValue(DataWorksheetTable.Cell(NumReplicates*(i-1)+1,-1)) = True Then
		If BlankCell(DataWorksheetTable.Cell(NumReplicates*(i-1)+1,-1)) = True Then
			ReportWorksheet.DataTable.ColumnTitle(i) = ""  'corrected to solve title with : - RRM 5-26-09
'			ReportWorksheet.DataTable.Cell(i, -1) = ""
		Else
'			ReportWorksheet.DataTable.Cell(i, -1) = DataWorksheetTable.Cell(NumReplicates*(i-1)+1, -1)
			ReportWorksheet.DataTable.ColumnTitle(i) = DataWorksheetTable.Cell(NumReplicates*(i-1)+1, -1)    'corrected to solve title with : - RRM 5-26-09
		End If
	Next i

	'Thicken borders between data groups
	ActiveDocument.CurrentDataItem.SetRegionBorderThickness(1,1,NumDataGroups+1)

	'Set column widths
	ReportWorksheet.SetColumnWidth(0,24)
	For i = 1 To NumDataGroups
		ReportWorksheet.SetColumnWidth(i,16)
	Next i
	
	'Set size and position of report worksheet
	ReportWorksheet.Left = 5
	ReportWorksheet.Top = 5
	ReportWorksheet.Height = 500
	ReportWorksheet.Width = 440	
	GoTo Finish
	
	ReplicateSectionNameError:
	Select Case Err.Number	'evaluate error number.
		Case 65535	        'duplicate section name
			iError = iError + 1
		Case Else           'handle other situations here
			MsgBox(Err.Description + " (" + CStr(Err.Number) + ")" + " in CreateReportWorksheetAndReport subroutine", 16, "Ligand Binding Macro")
	End Select
	Resume NameSection

	ReplicateWorksheetNameError:
	Select Case Err.Number	'evaluate error number.
		Case 65535	        'duplicate worksheet name
			iError = iError + 1
		Case Else           'handle other situations here
			MsgBox(Err.Description + " (" + CStr(Err.Number) + ")" + " in CreateReportWorksheetAndReport subroutine", 16, "Ligand Binding Macro")
	End Select
	Resume NameWorksheet

	Finish:
End Sub
Function TValue(ByVal NumDataPoints As Long, ByVal NumParameters As Double) As Double
'Compute t value

    Dim n As Double, Z As Double, v As Double
    Dim t123 As Double, t4 As Double, t5 As Double, t6 As Double, t1 As Double
    t123 = 0
    t1 = 0
    t4 = 0
    t5 = 0
    t6 = 0
    n = NumDataPoints
    Z = 1.96
    v = n - CLng(NumParameters)

    On Error GoTo ErrorHandler
    If v = 1 Then
        TValue = 12.706
    ElseIf v = 2 Then
        TValue = 4.303
    Else
        t123 = Z + (Z ^ 3 + Z) / (4 * v) + (5 * Z ^ 5 + 16 * Z ^ 3 + 3 * Z) / (96 * v ^ 2)
        t4 = (3 * Z ^ 7 + 19 * Z ^ 5 + 17 * Z ^ 3 - 15 * Z) / (384 * v ^ 3)
        t5 = 79 * Z ^ 9 + 776 * Z ^ 7 + 1482 * Z ^ 5 - 1920 * Z ^ 3 - 945 * Z
        t6 = 27 * Z ^ 11 + 339 * Z ^ 9 + 930 * Z ^ 7 - 1782 * Z ^ 5 - 765 * Z ^ 3 + 17955 * Z
        TValue = t123 + t4 + t5 / (92160 * v ^ 4) + t6 / (368640 * v ^ 5)
    End If
    Exit Function

ErrorHandler:
    If Err.Number = 11 Then ' Divide by zero
    	MsgBox "Divide by zero occurred in TValue", vbExclamation, "Zero Divide"
        Resume Next
    Else
		MsgBox(Err.Description + " (" + CStr(Err.Number) + ")" + " in TValue subroutine", 16, "Ligand Binding Macro")
		Resume Next
    End If
End Function
Function FormatNumber(ByVal Number As Variant)As String
	If Number = "+Inf" Then
		FormatNumber = "+Inf"
	Else
		Dim sign As Integer
		sign = Sgn(Number)
		Number = Abs(Number)
		If Number > 10000.0 Then
			FormatNumber = CStr(Format$(Number, "0.000e+0"))
		ElseIf Number > 1000 Then
			FormatNumber = CStr(Format$(Number, "####."))
		ElseIf Number > 100 Then
			FormatNumber = CStr(Format$(Number, "###.#"))
		ElseIf Number > 10 Then
			FormatNumber = CStr(Format$(Number, "##.##"))
		ElseIf Number > 1 Then
			FormatNumber = CStr(Format$(Number, "#.###"))
		ElseIf Number > .1 Then
			FormatNumber = CStr(Format$(Number, "0.####"))
		ElseIf Number > .01 Then
			FormatNumber = CStr(Format$(Number, "0.#####"))
		ElseIf Number > .001 Then
			FormatNumber = CStr(Format$(Number, "0.######"))
		ElseIf Number > .0001 Then
			FormatNumber = CStr(Format$(Number, "0.#######"))
		Else
			FormatNumber = CStr(Format$(Number, "0.000e+0"))
		End If
		If sign = -1 Then FormatNumber = "-" + FormatNumber
	End If
End Function
Function AntiLog(ByVal Number As Double) As Variant
	If Number < 308 Then
		AntiLog = 10^Number
	Else
		AntiLog = "+Inf"
	End If
End Function
Sub GetEmptyValues
	QNAN = "-1" & sD & "#QNAN"
	QNB = "-1" & sD & "#QNB"
End Sub
Public Function max_array(A As Variant, MaxColumn As Long, MaxRow As Long)
	'Computes the maximum value of the array A consisting of maxcolumn number of
	'columns and maxrow number of rows.

	Dim i As Long, j As Long
	Dim maxval As Variant
	maxval = A(0,0)
	For i = 0 To MaxColumn
		For j = 0 To MaxRow
		If A(i,j) > maxval Then
			maxval = A(i,j)
		End If
		Next j
	Next i
	max_array = maxval
End Function
Public Function min_array(A As Variant, MaxColumn As Long, MaxRow As Long)
	'Computes the minimum value of the array A consisting of maxcolumn number of
	'columns and maxrow number of rows.

	Dim i As Long, j As Long
	Dim minval As Variant
	minval = A(0,0)
	For i = 0 To MaxColumn
		For j = 0 To MaxRow
		If A(i,j) < minval And MissingValue(A(i,j)) = False Then
			minval = A(i,j)
		End If
		Next j
	Next i
	min_array = minval
End Function
Public Function LastNonMissingRow(column As Variant, column_end As Variant)
'Returns row number (1-based) of last non-missing value in a column
'column and column_end are 1-based

	Dim WorksheetTable As Object
	Set WorksheetTable = ActiveDocument.CurrentDataItem.DataTable
	Dim i As Long
	For i = 1 To column_end
		If MissingValue(WorksheetTable.Cell(column-1,i-1)) = False Then
			LastNonMissingRow = i
		End If
	Next i
End Function
Function MissingValue(ByVal X As Variant) As Boolean
'Determines if the value X is a missing value

	MissingValue = False
	If VarType(X) = vbString _
		Or X = "1.#INF" Or X = "1,#INF" _
		Or X = "-1.#INF" Or X = "-1,#INF" _
		Or X = "-1.#IND" Or X = "-1,#IND" _
		Or X = "1.#QNAN" Or X = "1,#QNAN" _
		Or X = "-1.#QNAN" Or X = "-1,#QNAN" Then
			MissingValue = True
	End If
End Function
Function BlankCell(ByVal value) As Boolean
    If value = "" Or value = QNAN Or value = QNB Or value = "--" Then
        BlankCell = True
    Else
        BlankCell = False
    End If
End Function
Sub NumberValidRows(ByRef Xdata() As Variant, ByRef YData() As Variant, _
ByVal NumXRows As Long, ByRef TotalNumMissingValues As Long, ByRef NumValidRows As Long)
'Counts rows from X and Y data for which there is a nonmissing X and at least one nonmissing Y
'Also counts the number of missing values in the valid rows ('valid' means the number of rows
'used by the curve fitter).

	Dim i As Long, j As Long, NumMissingValues As Long
	NumValidRows = 0
	TotalNumMissingValues = 0
	For i = 0 To NumXRows-1
		NumMissingValues = 0
		For j = 0 To NumReplicates-1
			If MissingValue(YData(j,i)) = True Then
				NumMissingValues = NumMissingValues + 1
			End If
		Next j
		If MissingValue(XData(0,i)) = False And NumMissingValues < NumReplicates Then
			NumValidRows = NumValidRows + 1
			TotalNumMissingValues = TotalNumMissingValues + NumMissingValues
		End If
	Next i
End Sub
Private Function DialogFunc(DlgItem$, Action%, SuppValue%) As Boolean
'See DialogFunc help topic For more information.

	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgText "Kd", ""
		DlgText "ligand", ""
		DlgText "NumberOfReps", CStr(NumReplicates)
		DlgValue "FitList", EquationIndex
		DlgEnable "KiGroupBox",KiCheckBoxState
		DlgEnable "Kd",KiCheckBoxState
		DlgEnable "tKd",KiCheckBoxState
		DlgEnable "ligand",KiCheckBoxState
		DlgEnable "tligand",KiCheckBoxState
		DlgEnable "EnterKdCheckBox",KiCheckBoxState
		DlgText "Kd", KdSave
		DlgText "ligand", ligandSave
		DlgText "Text3", Text3Save
		DlgText "MinX", MinXSave
		DlgText "MaxX", MaxXSave
		If XRangeCheckBoxState = True Then
			DlgEnable "tMinX", False
			DlgEnable "MinX", False
			DlgEnable "tMaxX", False
			DlgEnable "MaxX", False
		Else
			DlgEnable "tMinX", True
			DlgEnable "MinX", True
			DlgEnable "tMaxX", True
			DlgEnable "MaxX", True
		End If
	Case 2 ' Value changing or button pressed
		Rem Histogram = True ' Prevent button press from closing the dialog box
		Select Case DlgItem$
			Case "PushButton2"
				Help(ObjectHelp,HelpID)
				DialogFunc = True 'do not exit the dialog
			Case "FitList"
				Select Case SuppValue%
					Case 0,1,2,3,9
						DlgEnable "KiGroupBox",False
						DlgEnable "Kd",False
						DlgEnable "tKd",False
						DlgEnable "ligand",False
						DlgEnable "tligand",False
						DlgValue "EnterKdCheckBox", 0
						DlgEnable "EnterKdCheckBox",False
					Case Else
						DlgEnable "KiGroupBox",True
						DlgEnable "EnterKdCheckBox",True
				End Select
				Select Case SuppValue%
					Case 0  'one site saturation
						DlgText "Text3", "Y increases with X."
					Case 1  'two site saturation
						DlgText "Text3", "Y increases with X."
					Case 2  'one site saturation + nonspecific
						DlgText "Text3", "Y increases with X."
					Case 3  'two site saturation + nonspecific
						DlgText "Text3", "Y increases with X."
					Case 4  'sigmoidal dose-response
						DlgText "Text3", "Y increases with X." + vbCrLf +"Use log(X) data."
					Case 5  'sigmoidal dose-response (variable slope)
						DlgText "Text3", "Y increases (Hillslope > 0) or decreases with X." + vbCrLf +"Use log(X) data."
					Case 6  'one site competition
						DlgText "Text3", "Y decreases with X." + vbCrLf +"Use log(X) data."
					Case 7  'two site competition
						DlgText "Text3", "Y decreases with X." + vbCrLf +"Use log(X) data."
					Case 8  'four-parameter logistic function
						DlgText "Text3", "Y increases (B < 0) or decreases with X." + vbCrLf +"Use log(X) data."
					Case 9  'four-parameter logistic function (linear)
						DlgText "Text3", "Y decreases with X."
					Case Else
						DlgText "Text3", "User defined equation."
				End Select
				Text3Save = DlgText("Text3")
			Case "EnterKdCheckBox"
				If SuppValue% = 0 Then      'unchecked
				    DlgEnable "Kd",False
				    DlgEnable "tKd",False
				    DlgEnable "ligand",False
				    DlgEnable "tligand",False
					DialogFunc = True 'keep dialog open
				ElseIf SuppValue% = 1 Then  'checked
				    DlgEnable "Kd",True
				    DlgEnable "tKd",True
				    DlgEnable "ligand",True
				    DlgEnable "tligand",True
					DialogFunc = True 'keep dialog open
				End If
			Case "XRangeCheckBox"
				If SuppValue% = 0 Then      'unchecked
					DlgEnable "tMinX", True
					DlgEnable "MinX", True
					DlgEnable "tMaxX", True
					DlgEnable "MaxX", True
				ElseIf SuppValue% = 1 Then  'checked
					DlgEnable "tMinX", False
					DlgEnable "MinX", False
					DlgEnable "tMaxX", False
					DlgEnable "MaxX", False
				End If
		End Select
		If IsNumeric(CVar(DlgText("NumberOfReps"))) = True Then
			NumReplicates = CLng(DlgText("NumberOfReps"))
		Else
			NumReplicates = 1
		End If
		EquationIndex = DlgValue("FitList")
		KiCheckBoxState = DlgValue("EnterKdCheckBox")
		KdSave = DlgText("Kd")
		XRangeCheckBoxState = DlgValue("XRangeCheckBox")
		MinXSave = DlgText("MinX")
		MaxXSave = DlgText("MaxX")
		ligandSave = DlgText("ligand")
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Histogram = True ' Continue getting idle actions
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
Private Function CheckForStrings(ByVal SetToMissing As Boolean) As Boolean
'Determines if a string (ordate) exists in the worksheet.
'If SetToMissing is True will change strings to blanks

	Dim CurrentWorksheet As Object
	Set CurrentWorksheet = ActiveDocument.CurrentDataItem
	Dim DataWorksheetTable As Object
	Set DataWorksheetTable = CurrentWorksheet.DataTable
	Dim i As Long, j As Long
	CheckForStrings = False
	For i = 0 To LastRow-1
		For j = 0 To LastColumn-1
			If VarType(DataWorksheetTable.Cell(j,i)) = 8 Then
				If SetToMissing = False Then
					CheckForStrings = True
					Exit For
				Else
					DataWorksheetTable.Cell(j,i) = ""
				End If
			End If
		Next j
		If SetToMissing = False Then
			If CheckForStrings = True Then Exit For
		End If
	Next i
End Function
Private Sub FirstAndLastRow(ByRef FirstRow As Long, ByRef LastRow As Long, _
ByRef Data()As Variant, ByVal Nrows As Long, ByVal Ncols As Long)
'Finds the first and last non-missing rows in the Data() array.  If there are multiple
'columns then FirstRow is the minimum first-rows of all columns and LastRow is the
'maximum last-rows of all graphs.

	FirstRow = Nrows
	LastRow = 1
	
	'Find FirstRow
	Dim FirstRowj As Long
	Dim i As Long
	For i = 0 To Ncols-1
		FirstRowj = 1
		Dim j As Long
		For j = 0 To Nrows-1
			If MissingValue(Data(i,j)) = True Then
				FirstRowj = j+2
			Else
				Exit For
			End If
		Next j
		If FirstRowj < FirstRow Then  FirstRow = FirstRowj
		
		'Find LastRow
		Dim LastRowj As Long
		LastRowj = Nrows
		For j = Nrows-1 To 0 Step -1
			If MissingValue(Data(i,j)) = True Then
				LastRowj = j
			Else
				Exit For
			End If
		Next j
		If LastRowj > LastRow Then LastRow = LastRowj
	Next i
End Sub
Private Sub SubstituteAndReplaceXs(ByVal Mode As String, ByRef XData() As Variant, _
ByVal xFirstRow As Long, ByVal xLastRow As Long, ByVal yFirstRow As Long, ByVal yLastRow As Long, _
ByRef xFirstTemp() As Variant, ByRef xLastTemp() As Variant, ByVal GraphWorksheetTable As Object)
'If Mode = "Substitute" then replaces nonoverlapping-with-Ys Xs	with blank cells
'If Mode = "Replace" then replaces original Xs

If Mode = "Substitute" Then
	If yFirstRow > xFirstRow Then
		Dim i As Long
		For i = 0 To yFirstRow-2
			xFirstTemp(i) = XData(0,i)
			XData(0,i) = ""
			GraphWorksheetTable.Cell(1,i) = ""
		Next i
	End If
	If yLastRow < xLastRow Then
		For i = yLastRow To xLastRow-1
			xLastTemp(i) = XData(0,i)
			XData(0,i) = ""
			GraphWorksheetTable.Cell(1,i) = ""			
		Next i
	End If
	
Else 'Replace
	If yFirstRow > xFirstRow Then
		For i = 0 To yFirstRow-2
			XData(0,i) = xFirstTemp(i)
			GraphWorksheetTable.Cell(1,i) = XFirstTemp(i)			
		Next i
	End If
	If yLastRow < xLastRow Then
		For i = yLastRow To xLastRow-1
			XData(0,i) = xLastTemp(i)
			GraphWorksheetTable.Cell(1,i) = xLastTemp(i)			
		Next i
	End If
End If
End Sub