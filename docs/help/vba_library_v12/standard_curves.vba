'Macro created 2/1/2002 by John Kuo
'Several bugs fixed 8/10/04, Dick Mitchell
'Many modifications 1/23/08, Dick Mitchell
'Fixed bugs and improved error checking, added curve fit UI, positioned notebook items 7/20/09 RRM
'Fixed integer X-from-Y on German system and X range problems RRM 9-15-11
Option Explicit
Public Const ObjectHelp = Path + "\Standard Curves.CHM"
Public HelpID As Variant
Dim CurrentNotebook As Object
Dim CurrentWorksheet As Object
Dim FitFile As Object
Dim SPFit As Object
Dim Results As Object
Dim SPGraph As Object
Dim SPPage As Object
Dim XAxis As Object
Dim YAxis As Object
Dim WorksheetTable As Object
Dim SPEquation As Object
Dim Separator$, Equation$, Section$, Model$
Dim UsedColumns$(), ListedColumns(), ColContents, ColTitle, ColumnsPerPlot(),PlotColumnCountArray()
Dim Equations$()
Dim WhichEquation As Integer
Dim LogScale As Integer
Dim Index As Integer
Dim ListIndex As Integer
Dim i As Integer
Dim SolveEq As Integer
Dim PlotPredicted As Integer
Dim YReplicates As Integer
Dim ConvertLog As Integer
Dim a As Double
Dim b As Double
Dim y0 As Double
Dim Min As Double
Dim Max As Double
Dim Hillslope As Double
Dim EC50 As Double
Dim s As Double
Dim Slope1 As Double
'Dim Slope2 As Double
Dim SlopeCon As Double
Dim colx As Long
Dim colxSave As Long
Dim coly As Long
Dim colpredict As Long
Dim LastColumn As Long
Dim LastRow As Long
Dim OldLastRow As Long
Dim XMin, XMax
Dim PrevX As Integer
Dim PrevY As Integer
Dim PrevLog As Integer
Dim SolveX As Boolean
Dim SolveECpct As Boolean
Dim QNAN As String
Dim QNB As String
Dim sD
Dim DynamicFitFlag As Boolean
Dim Iterations As Variant
Dim IterationsSave As Double
Dim ConvergenceTolerance As Variant
Dim ConvergenceToleranceSave As Double
Dim Stepsize As Variant
Dim PageName As String
Dim MakeVisible As Boolean             'invisible if False (except FitNotebook)
Dim NumTicks As Integer                'number of tick marks on log axis
Function FlagOn(flag As Long)
  FlagOn = flag Or FLAG_SET_BIT ' Use to set option flag bits on, leaving others unchanged
End Function
Function FlagOff(flag As Long)
  FlagOff = flag Or FLAG_CLEAR_BIT ' Use to set option flag bits off, leaving others unchanged
End Function
Sub Main
Separator = ListSeparator 'international list separator
sD = DecimalSymbol  'international decimal symbol
GetEmptyValues

Stepsize = 1
HelpID = 1
SolveX=True
SolveECpct=False
MakeVisible = False
'MakeVisible = True

'On Error GoTo NoData
ActiveDocument.CurrentDataItem.Open
Set CurrentNotebook = ActiveDocument

ReDim Equations$(4)
Equations(0) = "Linear Equation"
Equations(1) = "Quadratic Equation"
Equations(2) = "Four Parameter Logistic"
Equations(3) = "Five Parameter Logistic"
Equations(4) = "Five Parameter Logistic - Two Slopes"

WhichEquation = 2

'Determine the data range
Set CurrentWorksheet = ActiveDocument.CurrentDataItem
Set WorksheetTable = CurrentWorksheet.DataTable
LastColumn = 0
LastRow = 0
WorksheetTable.GetMaxUsedSize(LastColumn,LastRow)  'LastColumn and LastRow are 1-based

OldLastRow=LastRow

'Sort through columns and create list of columns with values in row 1
Dim NumberEmptyColumns As Integer
NumberEmptyColumns = 0
ReDim UsedColumns$(LastColumn -1)
ReDim ListedColumns(LastColumn -1)
ListIndex = 0
For Index = 0 To LastColumn - 1
	ColContents = empty_col(Index, LastRow)
	ColTitle = WorksheetTable.Cell(Index,-1) 'Retrieve column title
	If ColContents = True Then
		NumberEmptyColumns = NumberEmptyColumns + 1
		GoTo NextIndex
	End If
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

'Check for inadequate data
If LastColumn - NumberEmptyColumns < 2 Then
	MsgBox("There must be at least two columns of data in your worksheet.",vbOkOnly, "More data needed")
	GoTo Finish2
End If

DialogBox

GoTo Finish:
NoData:
HelpMsgBox HelpID, "You must have a worksheet open with X and Y data columns.",vbExclamation,"No Open Worksheet"

Finish:

'Close the report and position the graphpage and worksheet
ArrangeNotebookItems

Finish2:

End Sub
Sub DialogBox

	Repeat:
	Begin Dialog UserDialog 550,455,"Standard Curve",.DialogFunc ' %GRID:10,7,1,1
		GroupBox 10,7,530,205,"Standard equation",.GroupBox1
		Text 20,21,80,14,"&Equation",.Text0
		DropListBox 20,35,350,133,Equations(),.Equation
		CheckBox 390,38,140,14,"&Log X axis scale",.CheckBox1
		GroupBox 10,212,530,77,"Data",.GroupBox7
		Text 20,224,140,14,"&X data column",.Text1
		DropListBox 20,238,140,105,UsedColumns(),.Xcol
		CheckBox 20,266,150,14,"Lo&g format X data",.LogX
		Text 180,224,170,14,"&Y data column",.Text2
		DropListBox 180,238,140,105,UsedColumns(),.Ycol
		CheckBox 180,266,150,14,"Y &replicates",.Replicates
		GroupBox 10,292,530,127,"Predict values",.GroupBox3
		CheckBox 20,308,200,14,"&Predict unknowns",.SolvePred
		Text 20,329,180,14,"&Solve for X from column:",.Text3
		DropListBox 20,343,140,105,UsedColumns(),.Ycol2
		GroupBox 270,301,250,110,"Predict",.GroupBox2
		OptionGroup .Group1
			OptionButton 280,322,150,14,"Ys &from Xs",.PredictY
			OptionButton 280,343,150,14,"Xs fro&m Ys",.PredictX
			OptionButton 280,364,150,14,"EC%s from %s",.PredictEC
		CheckBox 280,385,160,14,"Plot predicted &values",.PlotPred
		PushButton 10,427,100,21,"Help",.PushButton2
		GroupBox 345,220,185,60,"",.GroupBox4
		Text 355,231,170,35,"Results are always placed starting in the first empty worksheet column",.Text4
		GroupBox 20,56,510,57,"",.GroupBox5
		Text 30,65,495,14,"y = min + (max-min)/(1+(x/EC50)^(-Hillslope))",.EquationText
		Text 30,82,490,28,"The four parameter logistic function is the classic dose-response ""S"" shaped curve, and is most often plotted with a log-x axis.",.EquationDescription
		Text 20,385,240,14,"Results placed in column #",.ResCol
		CancelButton 440,427,100,21,.CancelButton
		OKButton 320,427,100,21,.OKButton
		CheckBox 20,128,160,14,"Dynamic curve fit",.DynamicFit
		GroupBox 170,112,360,48,"",.GroupBox6
		Text 175,119,350,40,"A dynamic curve fit attempts to find the global minimum using a 200 fit intelligent search.  It is particularly useful for the 5 parameter logistic equations.",.Text5
		Text 20,170,140,14,"Curve fit tolerance",.Text6
		TextBox 20,185,120,21,.FitTolerance
		Text 180,170,150,14,"Maximum fit iterations",.Text7
		TextBox 180,185,120,21,.MaxIterations
	End Dialog

Dim dlg As UserDialog
	dlg.Equation = 2  'Four parameter logistic
	dlg.Xcol=0
	dlg.Ycol=1
	dlg.Ycol2=0
	dlg.PlotPred=0
	dlg.SolvePred=0
	dlg.DynamicFit=0
	PrevX=0
	PrevY=1

Select Case Dialog(dlg)
	Case 0 'Handles Cancel button
		End
End Select

End Sub
Rem See DialogFunc help topic for more information.
Private Function DialogFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
		colx=ListedColumns(DlgValue("Xcol"))-1
		colxSave = colx
		coly=ListedColumns(DlgValue("Ycol"))-1
		SolveEq = 0
		PlotPredicted = 0
		YReplicates = 0
		ConvertLog = 0
		DlgText "ResCol", "Results placed in column "+ CStr(LastColumn+6)
		DlgValue "Group1",1
		DlgEnable "Group1", False
		DlgEnable "Ycol2", False
		DlgEnable "PlotPred", False
		DynamicFitFlag = False
		Iterations = 1000
		IterationsSave = Iterations
		ConvergenceTolerance = 1e-10
		ConvergenceToleranceSave = ConvergenceTolerance
		DlgText "MaxIterations", CStr(Iterations)
		DlgText "FitTolerance", CStr(Format(ConvergenceTolerance, "0.0e+00"))

	Case 2 ' Value changing or button pressed
        Select Case DlgItem$
		Case "OKButton"
			ConvergenceTolerance = CVar(DlgText("FitTolerance"))
			Dim ErrorCheck As Boolean
			ErrorCheck = False
			ConvergenceToleranceCheck ErrorCheck
			If ErrorCheck = True Then
				DlgText "FitTolerance", CStr(Format(ConvergenceToleranceSave, "0.0e+00")) 
				GoTo SetDialogFuncTrue
			End If
			Iterations = CVar(DlgText("MaxIterations"))
			MaxIterationsCheck ErrorCheck     'checks for interations checkbox entry errors
			If ErrorCheck = True Then
				DlgText "MaxIterations", CStr(IterationsSave)
				GoTo SetDialogFuncTrue
			End If
       		If ConvertLog = 1 Then  'for Log format X data option
       			LogLinear  'creates antilog X values to be used in the log X plot when have LogX data
       		Else           'get X data max and min and check for negative values
       			Dim DataMax As Double, DataMin As Double, DataMinPlus As Double
       			DataMaxAndMin colx+1, LastRow, DataMax, DataMin, DataMinPlus  'give error message if negative values
'				If DataMin <= 0 Then
				If DataMin < 0 Then
					MsgBox "You have a negative X data value(s).  Please check your data",vbExclamation, "Data Entry Error"
       				GoTo SetDialogFuncTrue
				End If
       		End If
       		
       		'create scatter plot of data points
			If YReplicates = 0 Then
				Plot
			Else
				PlotReplicates
			End If
			
			'perform curve fit
			Compute
			WorksheetTable.GetMaxUsedSize(LastColumn,LastRow)
			DlgText "ResCol", "Results placed in column "+ CStr(LastColumn+1)
			
			'find roots and add predicted values to the graph
			If SolveEq = 1 Then Solve
			GoTo EndOKButton
			
			SetDialogFuncTrue:
			DialogFunc = True
			EndOKButton:
		Case "PushButton2"
			Help(ObjectHelp,HelpID)
			DialogFunc = True
		Case "Equation"
			WhichEquation = SuppValue
			Select Case SuppValue
				Case 0
					DlgText "EquationText", "y = y0 + a*x"
					DlgText "EquationDescription", "A straight line."
					If DlgValue("SolvePred")=1 Then
						DlgEnable "PredictEC", False
						If DlgValue("Group1")=2 Then DlgValue "Group1", 1
					End If
				Case 1
					DlgText "EquationText", "y = y0 + a*x + b*x^2"
					DlgText "EquationDescription", "The quadratic equation."
					If DlgValue("SolvePred")=1 Then
						DlgEnable "PredictEC", False
						If DlgValue("Group1")=2 Then DlgValue "Group1", 1
					End If
				Case 2
'					DlgText "EquationText", "y = min + (max-min)/(1+(EC50/x)^Hillslope)"
					DlgText "EquationText", "y = min + (max-min)/(1+(x/EC50)^(-Hillslope))"
					DlgText "EquationDescription", "The four parameter logistic function is the classic dose-response ""S"" shaped curve, and is most often plotted with a log-x axis."
					If DlgValue("SolvePred")=1 Then DlgEnable "PredictEC", True
				Case 3
'					DlgText "EquationText", "y = min + (max-min)/(1+(xb/x)^Hillslope)^g"
					DlgText "EquationText", "y = min + (max-min)/(1+(x/xb)^(-Hillslope))^s"
					DlgText "EquationDescription", "The five parameter logistic function is an asymmetric version of the four parameter logistic, and is most often plotted with a log-x axis."
'					If DlgValue("SolvePred")=1 Then
'						DlgEnable "PredictEC", False
'						If DlgValue("Group1")=2 Then DlgValue "Group1", 1
'					End If
					If DlgValue("SolvePred")=1 Then DlgEnable "PredictEC", True
				Case 4
'					DlgText "EquationText", "y = min + (max-min)/(1 + fx*(EC50/x)^Slope1 + (1-fx)*(EC50/x)^(Slope1*SCon))"
					DlgText "EquationText", "y = min+(max-min)/(1+fx*(x/EC50)^(-Slope1)+(1-fx)*(x/EC50)^(-Slope1*SCon))"
					DlgText "EquationDescription", "The five parameter with 2 slope logistic function is an asymmetric version of the four parameter logistic, and is most often plotted with a log-x axis."
'					If DlgValue("SolvePred")=1 Then
'						DlgEnable "PredictEC", False
'						If DlgValue("Group1")=2 Then DlgValue "Group1", 1
'					End If
					If DlgValue("SolvePred")=1 Then DlgEnable "PredictEC", True
			End Select
		Case "DynamicFit"
			If SuppValue = 0 Then
				DynamicFitFlag = False
			ElseIf SuppValue = 1 Then
				DynamicFitFlag = True
			End If
		Case "CheckBox1"  'LogX axis scale
        	LogScale = SuppValue
			PrevLog = LogScale
		Case "SolvePred"  'Predict unknowns checkbox is True
			On Error GoTo NoPredicted
			colpredict=ListedColumns(2)
'			OldLastRow = ColumnLength(colpredict,LastRow) + 1
			OldLastRow = ColumnLength(colpredict+1,LastRow) + 1  'RRM 7-8-09
			SolveEq = SuppValue
			If SuppValue = 0 Then
				DlgEnable "Group1", False
				DlgEnable "Ycol2", False
				DlgEnable "PlotPred", False
				If DlgValue("Group1")=2 Then DlgValue "Group1", 1
			Else
				DlgEnable "Group1", True
				DlgEnable "Ycol2", True
				DlgEnable "PlotPred", True
				DlgValue "Ycol2",ListIndex-1
				colpredict=ListedColumns(DlgValue("Ycol2"))-1
				OldLastRow = ColumnLength(colpredict+1,LastRow) + 1     'RRM 7-8-09

				'Lets try commenting this out and see what else it affects
'				If YReplicates = 1 Then	DlgValue "Ycol",ListIndex-2   'RRM 11/29/07 gives bogus results if previous results in worksheet
				If YReplicates = 1 Then
'					DlgValue "Ycol",ListIndex-2
					coly=ListedColumns(DlgValue("Ycol"))-1     'RRM 11/2/04
				End If
				If WhichEquation < 2 Then  'no EC% for linear, quadratic
					DlgEnable "PredictEC", False
					If DlgValue("Group1")=2 Then DlgValue "Group1", 1
				Else
					DlgEnable "PredictEC", True
				End If
			End If
			GoTo DoneSolve
			NoPredicted:
				If MsgBox("There is no column of values from which to predict data.  You must add another column of data if you wish to solve or plot predicted data", 49, "More data needed") = 2 Then End
				DlgValue "SolvePred",0
				DlgEnable "Group1", False
				DlgEnable "Ycol2", False
				DlgEnable "PlotPred", False

			DoneSolve:
		Case "PlotPred"
			PlotPredicted = SuppValue
		Case "Group1"
			If SuppValue = 0 Then
				SolveX = False
				SolveECpct=False
				DlgText "Text3","&Solve for Y from column:"
			ElseIf SuppValue = 1 Then
				SolveX = True
				SolveECpct=False
				DlgText "Text3","&Solve for X from column:"
			Else
				SolveX = True
				SolveECpct=True
				DlgText "Text3","&Solve for EC% from column:"
			End If
		Case "Replicates"
			YReplicates = SuppValue
			If SuppValue = 1 Then
				DlgText "Text2", "Last &Y replicate column"
				DlgText "Text4", "Y replicate columns must begin immedately to the right of the X column"
				If SolveEq=0 Then

'					DlgValue "Ycol", ListIndex-1
					DlgValue "Ycol2", ListIndex
'					coly=ListIndex-2
					coly=ListedColumns(DlgValue("Ycol"))-1  'RRM, 8-10-04
				Else
					DlgValue "Ycol2",ListIndex-1

'					DlgValue "Ycol",ListIndex-2
					coly=ListIndex-2
				End If
				colpredict=ListedColumns(DlgValue("Ycol2"))-1
			Else
				DlgText "Text2", "&Y data column"
				DlgText "Text4", "Results are always placed starting in the first empty worksheet column"
				DlgValue "Ycol", 1
				DlgValue "Ycol2",ListIndex-1
				coly=ListedColumns(DlgValue("Ycol"))-1
			End If
		Case "LogX"  'x data is logarithmic if "LogX" = 1
			ConvertLog = SuppValue
			If SuppValue = 1 Then
				LogScale = 1
				DlgValue "Checkbox1", 1
				DlgEnable "Checkbox1", False
			ElseIf SuppValue = 0 Then
				DlgEnable "Checkbox1", True
				DlgValue "Checkbox1", PrevLog
				LogScale=PrevLog
			End If
        Case "Xcol"
        	If DlgValue("Xcol") <> DlgValue("Ycol") Then
				colx=ListedColumns(DlgValue("Xcol"))-1
				PrevX=DlgValue("Xcol")
			Else
				MsgBox "You cannnot use the same value for x and y", 48, "X and Y Columns Identical"
				DlgValue "Xcol", PrevX
			End If
		Case "Ycol"
			If DlgValue("Xcol") <> DlgValue("Ycol") Then
				coly=ListedColumns(DlgValue("Ycol"))-1
				PrevY=DlgValue("Ycol")
			Else
				MsgBox "You cannnot use the same value for x and y", 48, "X and Y Columns Identical"
				DlgValue "Ycol", PrevY
			End If
		Case "Ycol2"
			colpredict=ListedColumns(DlgValue("Ycol2"))-1
			OldLastRow = ColumnLength(colpredict+1,LastRow) + 1     'RRM 7-8-09
        End Select

	Case 3 ' TextBox or ComboBox text changed
        Select Case DlgItem$
        	Case "FitTolerance  'added both text boxes 7-17-09
        		ConvergenceTolerance = CVar(DlgText("FitTolerance"))
        		ErrorCheck = False
				ConvergenceToleranceCheck ErrorCheck
				If ErrorCheck = True Then
					DlgText "FitTolerance", CStr(Format(ConvergenceToleranceSave, "0.0e+00")) 
				End If
        	Case "MaxIterations"
        		Iterations = CVar(DlgText("MaxIterations"))
        		ErrorCheck = False
        		MaxIterationsCheck ErrorCheck
				If ErrorCheck = True Then
					DlgText "MaxIterations", CStr(IterationsSave)
				End If
        End Select
    Case 5
        DialogFunc = True
	End Select
End Function
Sub ConvergenceToleranceCheck(ByRef ErrorCheck As Boolean)
	'Check for textbox data entry errors
	If ConvergenceTolerance = "" Then
		MsgBox "Please enter a positive number for the fit tolerance.",vbExclamation,"Data Entry Error"
		ErrorCheck = True
	ElseIf IsNumeric(ConvergenceTolerance) = False Then
		MsgBox "Please enter a positive number for the fit tolerance.",vbExclamation,"Data Entry Error"
		ErrorCheck = True
	ElseIf CDbl(ConvergenceTolerance) <= 0 Then
		MsgBox "Please enter a positive number for the fit tolerance.",vbExclamation,"Data Entry Error"
		ErrorCheck = True
	End If
End Sub
Sub MaxIterationsCheck(ByRef ErrorCheck As Boolean)
	'Check for textbox data entry errors
	Iterations = CVar(DlgText("MaxIterations"))
	If Iterations = "" Then
		MsgBox "Please enter a positive number for the maximum number of fit iterations.",vbExclamation,"Data Entry Error"
		ErrorCheck = True
	ElseIf IsNumeric(Iterations) = False Then
		MsgBox "Please enter a positive number for the maximum number of fit iterations.",vbExclamation,"Data Entry Error"
		ErrorCheck = True
	ElseIf CDbl(Iterations) <= 0 Or CDbl(Iterations) > 65000 Then
		MsgBox "Please enter a positive number less than 65,000 for the maximum number of fit iterations.",vbExclamation,"Data Entry Error"
		ErrorCheck = True
	End If
End Sub
Sub LogLinear
'Puts antilog X values into worksheet for X axis scaling 
'Used when X data has log values (ConvertLog = 1)

'	CurrentWorksheet.InsertCells(colx+1,0,colx+1,LastRow,3)
	Dim NewColX As Long
	NewColX = LastColumn+1  'Put antilog X column rightmost rather than next to X column, RRM 7-13-09
	i=0
	Dim XCell As Variant
	For i=0 To ColumnLength(colx+1,LastRow)
'		WorksheetTable.Cell(colx+1,i)=CVar("1e"+CStr(WorksheetTable.Cell(colx,i)))
		XCell = WorksheetTable.Cell(colx,i)
		If MissingValue(XCell) = False Then
'			If Abs(XCell) < 308 Then
			If Abs(XCell) < 290 Then  'prevents bug in Automation log X axis for one data set, RRM 7-20-09
'				WorksheetTable.Cell(colx+1,i)=10^XCell
				WorksheetTable.Cell(NewColX-1,i)=10^XCell  'put antilog values to rightmost column RRM 7-17-09
			Else
'				WorksheetTable.Cell(colx+1,i)="--"
				WorksheetTable.Cell(NewColX-1,i)="--"
			End If
		Else
'			WorksheetTable.Cell(colx+1,i)="--"
			WorksheetTable.Cell(NewColX-1,i)="--"
		End If
	Next i
'	colx=colx+1
'	coly=coly+1
'	colpredict=colpredict+1
	'Column title
	WorksheetTable.ColumnTitle(NewColX-1) = "Antilog X"
	colx=NewColX-1
End Sub
Sub Plot
'Creates scatter plot of the data points

	Set SPPage = ActiveDocument.NotebookItems.Add(CT_GRAPHICPAGE)
'	If MakeVisible = False Then SPPage.Visible = False Else SPPage.Visible = True

	PageName = SPPage.Name

	ReDim ColumnsPerPlot(2, 1)
	ColumnsPerPlot(0, 0) = colx
	ColumnsPerPlot(1, 0) = 0
	ColumnsPerPlot(2, 0) = LastRow
	ColumnsPerPlot(0, 1) = coly
	ColumnsPerPlot(1, 1) = 0
	ColumnsPerPlot(2, 1) = LastRow

	ReDim PlotColumnCountArray(0)
	PlotColumnCountArray(0) = 3
	SPPage.CreateWizardGraph("Scatter Plot", "Simple Scatter", "XY Pair", ColumnsPerPlot, PlotColumnCountArray)
	Set SPGraph = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0)

	'Increase symbol diameter and add fill color
	Dim SPPlot As Object
	Set SPPlot = SPGraph.Plots(0)
	SPPlot.SetAttribute(SSA_OPTIONS, &H00000201&)
	SPPlot.SetAttribute(SSA_EDGETHICKNESS, 10)
	SPPlot.SetAttribute(SSA_SIZE, 150)
	SPPlot.SetAttribute(SSA_SIZEREPEAT, 2)
	SPPlot.SetAttribute(SSA_COLOR, &H00c0c0c0&)  'RGB_GRAY?
	SPPlot.SetAttribute(SSA_COLORREPEAT, 2)

	'Set axis titles
	Set XAxis = SPGraph.Axes(0)
	Set YAxis = SPGraph.Axes(1)
	If WorksheetTable.Cell(colxSave,-1) <> QNAN Then XAxis.Name = WorksheetTable.Cell(colxSave,-1)
	If WorksheetTable.Cell(coly,-1) <> QNAN Then YAxis.Name = WorksheetTable.Cell(coly,-1)

	'Set graph title
	SPGraph.Name = "Standard Curve"

	'Remove legend
	SPGraph.SetAttribute(SGA_FLAGS,FlagOff(SGA_FLAG_AUTOLEGENDSHOW))

	'Change X axis scale type
	Select Case LogScale
		Case 1%
			XAxis.SetAttribute(SAA_TYPE, SAA_TYPE_COMMON)
			FormatLogXAxis
		Case 0%
			XAxis.SetAttribute(SAA_TYPE, SAA_TYPE_LINEAR)
	End Select

	SPPage.Open

End Sub
Sub PlotReplicates
'Creates scatter plot with error bars from replicate data values

	Set SPPage = ActiveDocument.NotebookItems.Add(CT_GRAPHICPAGE)

	ReDim ColumnsPerPlot(2, 2)
	ColumnsPerPlot(0, 0) = colx
'	ColumnsPerPlot(0, 0) = colxSave
	ColumnsPerPlot(1, 0) = 0
	ColumnsPerPlot(2, 0) = LastRow
'	ColumnsPerPlot(0, 1) = colx+1
	ColumnsPerPlot(0, 1) = colxSave+1
	ColumnsPerPlot(1, 1) = 0
	ColumnsPerPlot(2, 1) = LastRow
	ColumnsPerPlot(0, 2) = coly
	ColumnsPerPlot(1, 2) = 0
	ColumnsPerPlot(2, 2) = LastRow

	ReDim PlotColumnCountArray(0)
	PlotColumnCountArray(0) = 3
	SPPage.CreateWizardGraph("Scatter Plot", "Simple Error Bars", "X, Y Replicate", ColumnsPerPlot, PlotColumnCountArray, "Row Means", "Standard Deviation", "Degrees", 0.000000, 360.000000, , "Standard Deviation", True)
	Set SPGraph = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0)

	'Increase symbol diameter and add fill color
	Dim SPPlot As Object
	Set SPPlot = SPGraph.Plots(0)
	SPPlot.SetAttribute(SSA_OPTIONS, &H00000201&)
	SPPlot.SetAttribute(SSA_EDGETHICKNESS, 10)
	SPPlot.SetAttribute(SSA_SIZE, 150)
	SPPlot.SetAttribute(SSA_SIZEREPEAT, 2)
	SPPlot.SetAttribute(SSA_COLOR, &H00c0c0c0&)  'RGB_GRAY?
	SPPlot.SetAttribute(SSA_COLORREPEAT, 2)

	'Set axis titles
	Set XAxis = SPGraph.Axes(0)
	Set YAxis = SPGraph.Axes(1)
	If WorksheetTable.Cell(colxSave,-1) <> QNAN Then XAxis.Name = WorksheetTable.Cell(colxSave,-1)
	If WorksheetTable.Cell(coly,-1) <> QNAN Then YAxis.Name = WorksheetTable.Cell(colxSave+1,-1)

	'Set graph title
	SPGraph.Name = "Standard Curve"

	'Remove legend
	SPGraph.SetAttribute(SGA_FLAGS,FlagOff(SGA_FLAG_AUTOLEGENDSHOW))

	'Set X axis scale type
	Select Case LogScale
		Case 1%
			XAxis.SetAttribute(SAA_TYPE, SAA_TYPE_COMMON)
			FormatLogXAxis
		Case 0%
			XAxis.SetAttribute(SAA_TYPE, SAA_TYPE_LINEAR)
	End Select
End Sub
Sub DataMaxAndMin(ByVal column As Long, ByVal Ndata As Long, ByRef DataMax As Double, _
ByRef DataMin As Double, ByRef DataMinPlus As Double)
'Find the maximum and minimum of a column of data
'column, Ndata are 1-based
'DataMinPlus is used for the non-log data and log X axis case

	Dim i As Long
	Dim Data As Variant
	DataMax = -10^307
	DataMin = 10^307
	DataMinPlus = 10^307
	For i = 1 To Ndata
		Data = WorksheetTable.Cell(column-1,i-1)
		If MissingValue(Data) = False Then
			If Data > DataMax Then DataMax = Data
			If Data < DataMin Then DataMin = Data
			If Data < DataMinPlus And Data <> 0 Then DataMinPlus = Data
		End If
	Next i
End Sub
Sub DataMaxAndMinInRange(ByVal column As Long, ByVal Ndata As Long, _
ByVal min As Double, ByVal max As Double, _
ByRef DataMinInRange As Double, ByRef DataMaxInRange As Double)
'Find the maximum and minimum of a column of data that is within the
'min and max parameters of 4PL and 5PL equations
'column, Ndata are 1-based

	Dim i As Long
	Dim Data As Variant
	DataMaxInRange = -10^307
	DataMinInRange = 10^307
	For i = 1 To Ndata
		Data = WorksheetTable.Cell(column-1,i-1)
		If MissingValue(Data) = False Then
			If Data > DataMaxInRange And Data <= max Then DataMaxInRange = Data
			If Data < DataMinInRange And Data >= min Then DataMinInRange = Data
		End If
	Next i
End Sub
Sub MaxAndMinYForPositiveDiscriminant(ByVal column As Long, ByVal Ndata As Long, _
ByVal a As Double,ByVal b As Double,ByVal y0 As Double, _
ByVal Xmin As Double, ByVal Xmax As Double, _
ByRef XMinInRange As Double, ByRef XMaxInRange As Double)
'Finds the min and max Y values for which a nonnegative discriminant occurs.  Then finds the X
'values for these.  This is a very crude way to deal with possible multiple values.

	Dim i As Integer
	Dim Data As Variant
	Dim YMaxInRange As Double, YMinInRange As Double, Discriminant As Double
	YMaxInRange = -10^307
	YMinInRange = 10^307
	Dim AllDiscriminantNegative As Boolean
	AllDiscriminantNegative = True
	
	For i = 1 To Ndata
			Data = WorksheetTable.Cell(column-1,i-1)
			Discriminant = A*A - 4*b*(y0-Data)
			If Discriminant >= 0 Then AllDiscriminantNegative = False
			If MissingValue(Data) = False Then
				If Discriminant >= 0 And Data > YMaxInRange Then YMaxInRange = Data
				If Discriminant >= 0 And Data < YMinInRange Then YMinInRange = Data
			End If
	Next i
	'This isn't correct for multivalued case but in this case the quadratic function is not good anyway.
	If AllDiscriminantNegative = True Then
		XMinInRange = Xmin
		XmaxInRange = Xmax
	Else
		Dim XMaxInRange1 As Double, XMaxInRange2 As Double, XMinInRange1 As Double, XMinInRange2 As Double
		XMaxInRange1 = (-A + Sqr(A*A - 4*b*(y0-YMaxInRange)))/(2*b)
		XMaxInRange2 = (-A - Sqr(A*A - 4*b*(y0-YMaxInRange)))/(2*b)
		If XMaxInRange2 > XMaxInRange1 Then
			XMaxInRange = XMaxInRange2
		Else
			XMaxInRange = XMaxInRange1
		End If
		XMinInRange1 = (-a + Sqr(a*a - 4*b*(y0-YMinInRange)))/(2*b)
		XMinInRange2 = (-a - Sqr(a*a - 4*b*(y0-YMinInRange)))/(2*b)
		If XMinInRange2 < XMinInRange1 Then
			XMinInRange = XMinInRange2
		Else
			XMinInRange = XMinInRange1	
		End If
	End If
End Sub
Sub DataNonZeroMin(ByVal column As Long, ByVal Ndata As Long, ByRef DataMin As Double)
'Find next largest X value when the minimum value is 0

	Dim i As Long
	Dim Data As Variant
	DataMin = 10^307
	For i = 1 To Ndata
		Data = WorksheetTable.Cell(column-1,i-1)
		If MissingValue(Data) = False Then
			If Data < DataMin And Data <> 0 Then DataMin = Data
		End If
	Next i
End Sub
Sub GetLogAxisRangeAttributes(ByVal DataMin As Double, ByVal DataMax As Double, ByVal DataMinPlus As Double, _
ByRef RangeFromVal As Double, ByRef RangeToVal As Double, ByRef RangeTickInterval As Double)
'Uses 1-2-5 algorithm to find the range attributes

	Dim DeltaData As Double, IntervalLength As Double
	Dim IntervalMagnitude As Double, NormalizedIntervalLength As Double
	Dim DesiredNumberIntervals As Double
	DesiredNumberIntervals = 4   'should produce >=3 and <=9 intervals
	Dim Log10 As Double
	Log10 = Log(10)
	If ConvertLog = 0 And LogScale = 1 Then  'Non-log x data and log scale x axis
		'Check for zero data values and if so then use the next largest value for graph range attributes
		If DataMin = 0 Then DataMin = DataMinPlus
		DataMin = Log(DataMin)/Log10
		DataMax = Log(DataMax)/Log10
	End If
	DeltaData = DataMax-DataMin
	IntervalLength = DeltaData/DesiredNumberIntervals
	'Miminum tick is 1 decade
	If IntervalLength < 1 Then IntervalLength = 1
	IntervalMagnitude = Int(Log(IntervalLength)/Log(10))
	NormalizedIntervalLength = IntervalLength/10^IntervalMagnitude

	'find optimal normalized tick interval
	Dim OptInterval As Integer
	If Abs(NormalizedIntervalLength - 1) < Abs(NormalizedIntervalLength - 2) Then
		OptInterval = 1
	ElseIf Abs(NormalizedIntervalLength - 2) < Abs(NormalizedIntervalLength - 5) Then
		OptInterval = 2
	Else
		OptInterval = 5
	End If
	RangeTickInterval = OptInterval*10^IntervalMagnitude
	

	'find range FromVal and whether it is on the boundary
	Dim epsilon As Double
	epsilon = ((Abs(DataMin)+Abs(DataMax))/2)*10^(-10)
	RangeFromVal = RangeTickInterval*Int(DataMin/RangeTickInterval)
	If Abs(RangeFromVal - DataMin) < epsilon Then
		RangeFromVal = RangeFromVal - RangeTickInterval
	End If
	'find range Toval and whether it is on the boundary
	RangeToVal = RangeTickInterval*(Int(DataMax/RangeTickInterval) +1)
	If Abs(RangeToVal - DataMax) < epsilon Then
		RangeToVal = RangeToVal + RangeTickInterval
	End If
	
End Sub
Sub Compute
'Runs curve fit

'SPPage.Open
'SPGraph.Plots(0).ChildObjects(0).SelectObject

'Check for at least one missing value in antilog data (if any)
Dim FitErrorString As String
If ConvertLog = 1 Then
	Dim NumAntiLogDataPoints As Integer
	NumAntiLogDataPoints = ColumnLength(colxSave+1, LastRow) + 1  'use original column to avoid possible missing values in antilog column  RRM 7-21-09
	For i = 0 To NumAntiLogDataPoints-1
		If MissingValue(WorksheetTable.Cell(LastColumn,i)) = True Then
			FitErrorString = "An invalid antilog of your X data occurred.  Please check" +vbCrLf+ _
			"your X data or do not select the Log format X data option"
			'Delete the antilog and tick location columns
'			CurrentWorksheet.DeleteCells(LastColumn+1, 0, LastColumn+1, 31999999, DeleteLeft)  'causes crash RRM 7-20-09
'			CurrentWorksheet.DeleteCells(LastColumn, 0, LastColumn, 31999999, DeleteLeft)
			LastColumn = 0
			LastRow = 0
			WorksheetTable.GetMaxUsedSize(LastColumn,LastRow)  'LastColumn and LastRow are 1-based
			DeleteWorksheetCells LastColumn-1, 1, LastColumn, LastRow
			WorksheetTable.ColumnTitle(LastColumn-1) = ""
			WorksheetTable.ColumnTitle(LastColumn-2) = ""
			GoTo MissingAntilogError
		End If
	Next i
End If

On Error GoTo NoFitFile
	Notebooks.Open(UserPath + "\Standard.jfl")
	Set FitFile = Notebooks(UserPath + "\Standard.jfl")
	FitFile.Visible = False
On Error GoTo NoEquation
	Select Case WhichEquation
		Case 0
			Section = "Standard Curves"
			Equation = "Linear Curve"
'			Set SPFit = Notebooks(UserPath + "\Standard.jfl").NotebookItems("Linear Curve")
		Case 1
			Section = "Polynomial"
			Equation = "Quadratic"
'			Set SPFit = Notebooks(UserPath + "\Standard.jfl").NotebookItems("Quadratic")
		Case 2
			Section = "Standard Curves"
			Equation = "Four Parameter Logistic Curve"
'			Set SPFit = Notebooks(UserPath + "\Standard.jfl").NotebookItems("Four Parameter Logistic Curve")
		Case 3
			Section = "Standard Curves"
			Equation = "Five Parameter Logistic Curve"
'			Set SPFit = Notebooks(UserPath + "\Standard.jfl").NotebookItems("Five Parameter Logistic Curve")
		Case 4
			Section = "Standard Curves"
			Equation = "Five Parameter Logistic - 2 Slopes"
'			Set SPFit = Notebooks(UserPath + "\Standard.jfl").NotebookItems("Five Parameter Logistic - 2 Slopes")
		End Select
	Set SPFit = Notebooks(UserPath + "\Standard.jfl").NotebookItems(Equation)
	CurrentNotebook.Activate

	SPPage.Open
	SPGraph.Plots(0).ChildObjects(0).SelectObject

	If DynamicFitFlag = True Then
		SPFit.DynamicFitEnabled = True
	End If
	SPFit.Open

	If MakeVisible = False Then SPFit.Visible = False Else SPFit.Visible = True

	SPFit.Option("iterations") = Iterations
	SPFit.Option("tolerance") = ConvergenceTolerance
	SPFit.Option("stepsize") = Stepsize
	If DynamicFitFlag = True Then
		SPFit.NumberOfFits = 200
		SPFit.UseAutoParameterRanges = True
	End If
	SPFit.Run

	Dim Stats As Object
	Set Stats =  SPFit.FitResults

	Dim FitVerdict1 As Double, FitVerdict2 As Double, FitVerdict3 As Double
	FitVerdict1 = Stats.FitVerdict(0)
	FitVerdict2 = Stats.FitVerdict(1)
	FitVerdict3 = Stats.FitVerdict(2)

	'FitVerdict errors
	If FitVerdict1 = 7 Or FitVerdict2 = 9 Then  'unable to evaluate function using initial parameter values
		If ConvertLog = 1 Then  'x data is in log form
			FitErrorString = "You have selected your X data to be logarithmic but the anitlog of this data is out of range." +vbCrLf+ _
			"Please check your data or deselect the Log format X data Option."
			CurrentWorksheet.DeleteCells(LastColumn, 0, LastColumn, 32000000, DeleteLeft)
			GoTo OtherFitError
		Else
			FitErrorString = "The curve fitter was unable to estimate initial parameter values using the data" +vbCrLf+ _
			"in the worksheet.  Please check your data and the equation that was used."
		End If
		GoTo OtherFitError
	End If

	If FitVerdict1 = 2 Then 'inner loop failure
		FitErrorString = "There was an inner loop failure in the curve fitter.  This rarely occurs." +vbCrLf+ _
		"Please check your data and the equation that was used"
		GoTo OtherFitError
	End If

	Dim NumberOfIterations As Long
	NumberOfIterations = Stats.ComputedIterations
	'NumberOfIterations = 50000	'test
	If NumberOfIterations >= Iterations Then
		Dim ExceedIterations As Boolean
		ExceedIterations = SPFit.IterateMore  'try one more time
		'ExceedIterations = False  'test
		'remove antilog column
		If ExceedIterations = False Then
			If LogScale = 1 Then
'				CurrentWorksheet.DeleteCells(colx-1, 0, colx, 32000000, DeleteLeft)
				Dim MaxNumRows As Integer
				If NumTicks+1 > LastRow Then MaxNumRows = NumTicks+1 Else MaxNumRows = LastRow
				DeleteWorksheetCells colx+1, 0, colx+2, MaxNumRows
			End If
			'don't create graph or report
			SPFit.OutputReport = False
			SPFit.OutputEquation = False
			SPFit.ResidualsColumn = -2
			SPFit.PredictedColumn = -2
			SPFit.ParametersColumn = -2
			SPFit.OutputGraph = False
			SPFit.OutputAddPlot = False
			SPFit.ConfidenceBands = False
			SPFit.ExtendFitToAxes = False
			SPFit.AddEquationToTitle = False
			SPFit.AddPlotGraphIndex = 0
			SPFit.DynamicFitGraph = False
			SPFit.XColumn = -2
			SPFit.YColumn = -2
			SPFit.ZColumn = -2
			SPFit.Finish
			SPFit.Close(False)
			FitFile.Close(False)
			Set FitFile = Nothing
			Set XAxis = Nothing
			Set YAxis = Nothing
			Set SPGraph = Nothing
			ActiveDocument.NotebookItems.Delete(SPPage.Name)
			MsgBox "The curve fit did not converge using this equation.  This suggests that" +vbCrLf+ _
			"the fit problem is not well posed with this equation-data combination." +vbCrLf+vbCrLf+ _
			"It might be possible to achieve convergence by increasing the fit tolerance" +vbCrLf+ _
			"to 0.001, say, or, possibly by increasing the number of fit iterations",vbCritical,"Lack of Convergence"
			Exit All
		End If
	End If

	SPFit.OutputReport = True
	SPFit.OutputEquation = False
	SPFit.ResidualsColumn = -1
	SPFit.PredictedColumn = -1
	SPFit.ParametersColumn = -1
	SPFit.OutputGraph = False
	SPFit.OutputAddPlot = True
	SPFit.ConfidenceBands = False
	SPFit.ExtendFitToAxes = True
	SPFit.AddEquationToTitle = False
	SPFit.AddPlotGraphIndex = 0
	SPFit.DynamicFitGraph = False
	SPFit.XColumn = -1
	SPFit.YColumn = -1
	SPFit.ZColumn = -2

	Select Case WhichEquation
		Case 0
			a = SPFit.FittedParameterValue("a")
			y0 = SPFit.FittedParameterValue("y0")
			Model = CStr(y0)+"+"+CStr(a)+"*x"
		Case 1
			a = SPFit.FittedParameterValue("a")
			b = SPFit.FittedParameterValue("b")
			y0 = SPFit.FittedParameterValue("y0")
			Model = CStr(y0)+"+"+CStr(a)+"*x+"+CStr(b)+"*x^2"
		Case 2
			Min = SPFit.FittedParameterValue("min")
			Max = SPFit.FittedParameterValue("max")
			EC50 = SPFit.FittedParameterValue("EC50")
			Hillslope = SPFit.FittedParameterValue("Hillslope")
			Model = CStr(Min)+" + ("+CStr(Max)+"-"+CStr(Min)+")/(1+(x/"+CStr(EC50)+")^"+CStr(-Hillslope)+")"  'RRM 7-8-09
'			Equation = CStr(Min)+" + ("+CStr(Max)+"-"+CStr(Min)+")/(1+(x/"+CStr(EC50)+")^"+CStr(Hillslope)+")"
'			Equation = CStr(Min)+" + ("+CStr(Max)+"-"+CStr(Min)+")/(1+("+CStr(EC50)+"/x)^"+CStr(Hillslope)+")"
		Case 3
			Min = SPFit.FittedParameterValue("min")
			Max = SPFit.FittedParameterValue("max")
			EC50 = SPFit.FittedParameterValue("EC50")
			Hillslope = SPFit.FittedParameterValue("Hillslope")
			s = SPFit.FittedParameterValue("s")
			Dim xb As Double
'			xb=EC50*(10^((1/Hillslope)*Log(2^(1/g)-1)/Log(10)))
'			xb=EC50*(10^((-1/Hillslope)*Log(2^(1/s)-1)/Log(10)))
			xb=EC50*(10^((1/Hillslope)*Log(2^(1/s)-1)/Log(10)))   'RRM 5-21-08

'			Equation = CStr(Min)+" + ("+CStr(Max)+"-"+CStr(Min)+")/(1+(x/"+CStr(EC50)+")^"+CStr(Hillslope)+")"+"^"+CStr(s)
'			Equation = CStr(Min)+" + ("+CStr(Max)+"-"+CStr(Min)+")/(1+("+CStr(xb)+"/x)^"+CStr(Hillslope)+")"+"^"+CStr(s)
			Model = CStr(Min)+" + ("+CStr(Max)+"-"+CStr(Min)+")/(1+(x/"+CStr(xb)+")^"+CStr(-Hillslope)+")"+"^"+CStr(s)
		Case 4
			Min = SPFit.FittedParameterValue("min")
			Max = SPFit.FittedParameterValue("max")
			EC50 = SPFit.FittedParameterValue("EC50")
			Slope1 = SPFit.FittedParameterValue("Slope")
'			Slope2 = SPFit.FittedParameterValue("Slope2")
			SlopeCon = SPFit.FittedParameterValue("SlopeCon")
			Dim Cf As Double
			Cf=2*Slope1*Slope1*SlopeCon/Abs(Slope1+Slope1*SlopeCon)
			Dim fx As String
'			fx="(1/(1 + (x/" + CStr(EC50) +")^"+CStr(Cf)+"))"
			fx="(1/(1 + (x/" + CStr(EC50) +")^"+CStr(Cf)+"))"
'			Equation = CStr(Min)+" + ("+CStr(Max)+"-"+CStr(Min)+")/(1+"+fx+"*(x/"+CStr(EC50)+")^"+CStr(Slope1)+"+(1-"+fx+")*(x/"+CStr(EC50)+")^"+CStr(Slope2)+")"
			Model = CStr(Min)+" + ("+CStr(Max)+"-"+CStr(Min)+")/(1+"+fx+"*(x/"+CStr(EC50)+")^"+CStr(-Slope1)+"+(1-"+fx+")*(x/"+CStr(EC50)+")^"+CStr(-Slope1*SlopeCon)+")"
'			Equation = CStr(Min)+" + ("+CStr(Max)+"-"+CStr(Min)+")/(1+"+fx+"*("+CStr(EC50)+"/x)^"+CStr(Slope1)+"+(1-"+fx+")*("+CStr(EC50)+"/x)^"+CStr(Slope1*SlopeCon)+")"
	End Select

	SPFit.Finish
	SPFit.Close(False)
	FitFile.Close(False)
	Set FitFile = Nothing

	'Increase the fit line thickness
	Dim SPPlot As Object
	Set SPPlot = SPGraph.Plots(1)
	SPPlot.Line.SetAttribute(SLA_SELECTFUNC, 0)            'select line
	SPPlot.Line.SetAttribute(SEA_THICKNESS, 25)            'change thickness

	'Send fit lines to back
	SPPlot.SelectObject
	SPPlot.SetAttribute(SPA_SENDTOBACK) 'send tuple to back

	'Deselect plot
	SPPage.Select(False, 3958, 5159, 3958, 5159)

	'Name the page
	On Error GoTo ReplicatePageNameError
	Dim iError As Long
	iError = 1
	NamePage: If iError = 1 Then
		PageName = "Standard Curve"
	Else
		PageName = "Standard Curve " + CStr(iError)
	End If
	SPPage.Name = PageName

GoTo Done

ReplicatePageNameError:
Select Case Err.Number	'evaluate error number.
	Case 65535	        'duplicate page name
		iError = iError + 1
	Case Else           'handle other situations here
		MsgBox(Err.Description + " (" + CStr(Err.Number) + ")" + " in CreateReportWorksheetAndReport subroutine", 16, "Standard Curve")
End Select
Resume NamePage

NoFitFile:
MsgBox "No Fit Library Notebook found",vbCritical,"No Fit Library"
End

NoEquation:
MsgBox "The selected equation was not found in your curve fit library",vbCritical,"No Equation in Fit Library"
End

OtherFitError:
MsgBox FitErrorString, "Fit Error"
SPFit.Finish
SPFit.Close(False)
FitFile.Close(False)
Set FitFile = Nothing
GoTo Done

MissingAntilogError:
MsgBox FitErrorString, "Antilog Error"

'There is a bug which prevents deleting a page object  RRM 7-20-09
'SPPage.Open
'CurrentNotebook.NotebookItems.Open(PageName)
'CurrentNotebook.NotebookItems(PageName).Close(False)
'CurrentNotebook.NotebookItems(PageName).Activate
'Dim CurrentItem As Object, ItemName As String
'Set CurrentItem = ActiveDocument.CurrentItem
'ItemName = CurrentItem.Name
'CurrentNotebook.NotebookItems.Delete(PageName)
'CurrentNotebook.NotebookItems(PageName).Delete
'End

Exit All
Done:

End Sub
Sub Solve
'Finds roots or values of fitted equations for user-entered column
'Optionally adds droplines to graph

Set SPEquation = ActiveDocument.CurrentDataItem.PlotEquation
SPEquation.EquationRHS = Model

WorksheetTable.GetMaxUsedSize(LastColumn,LastRow)

If SolveX=False Then
	i=0
	For i=0 To OldLastRow-1
		WorksheetTable.Cell(LastColumn,i)=SPEquation.Evaluate(WorksheetTable.Cell(colpredict,i))
		If WorksheetTable.Cell(colpredict,i)="" Then WorksheetTable.Cell(LastColumn,i)="No solution"
		If WorksheetTable.Cell(colpredict,i) = QNAN Then
			WorksheetTable.Cell(LastColumn,i)=""
		End If
	Next i
	WorksheetTable.NamedRanges.Add("Predicted Y Values",LastColumn,0,1,-1, True)

	If PlotPredicted = 1 Then

		ReDim ColumnsPerPlot(2, 1)
		ColumnsPerPlot(0, 0) = colpredict
		ColumnsPerPlot(1, 0) = 0
'		ColumnsPerPlot(2, 0) = 31999999
		ColumnsPerPlot(2, 0) = LastRow
		ColumnsPerPlot(0, 1) = LastColumn
		ColumnsPerPlot(1, 1) = 0
'		ColumnsPerPlot(2, 1) = 31999999
		ColumnsPerPlot(2, 1) = LastRow

		ReDim PlotColumnCountArray(0)
		PlotColumnCountArray(0) = 3

		'Scatterplot for predicted values
		SPPage.AddWizardPlot("Scatter Plot", "Simple Scatter", "XY Pair", ColumnsPerPlot, PlotColumnCountArray)
		SPGraph.SetAttribute(SLA_PLOTOPTIONS, 548864)
		SPGraph.SetAttribute(SLA_SELECTDROP, 1)
		SPGraph.Plots(2).DropLines(3).SetAttribute(SEA_LINETYPE, 6)
		SPGraph.Plots(2).DropLines(3).SetAttribute(SEA_THICKNESS, 20)
		SPGraph.SetAttribute(SLA_SELECTDROP, 2)
		SPGraph.Plots(2).DropLines(2).SetAttribute(SEA_LINETYPE, 6)
		SPGraph.Plots(2).DropLines(2).SetAttribute(SEA_THICKNESS, 20)
		SPGraph.Plots(2).SetAttribute(SSA_COLOR, &H00ffffff&)

	End If

ElseIf SolveX=True Then

	XAxis.GetAttribute(SAA_FROMVAL,XMin)
	XAxis.GetAttribute(SAA_TOVAL,XMax)

	SPEquation.ChangeNotebook(UserPath + "\Standard.jfl")
	SPEquation.SetSection(Section)
	SPEquation.Open(Equation)

	Select Case WhichEquation
		Case 0
			SPEquation.SetParameter("y0", y0)
			SPEquation.SetParameter("a", a)
		Case 1
			SPEquation.SetParameter("y0", y0)
			SPEquation.SetParameter("a", a)
			SPEquation.SetParameter("b",b)
		Case 2
			SPEquation.SetParameter("min", Min)
			SPEquation.SetParameter("max", Max)
			SPEquation.SetParameter("EC50",EC50)
			SPEquation.SetParameter("Hillslope", Hillslope)
		Case 3
			SPEquation.SetParameter("min", Min)
			SPEquation.SetParameter("max", Max)
			SPEquation.SetParameter("EC50",EC50)
			SPEquation.SetParameter("Hillslope", Hillslope)
			SPEquation.SetParameter("s", s)
		Case 4
			SPEquation.SetParameter("min", Min)
			SPEquation.SetParameter("max", Max)
			SPEquation.SetParameter("EC50",EC50)
			SPEquation.SetParameter("Slope", Slope1)
			SPEquation.SetParameter("SlopeCon", SlopeCon)
	End Select
	
	'Get the minimum X and maximum X from both the data and from the function limits
	If SolveECpct = False Then
		Dim XMinFromY As Double, XMaxFromY As Double, DataMinPlus As Double
		GetXRangeForXFromY WhichEquation, y0, a, b, Min, Max, EC50, Hillslope, s, Slope1, SlopeCon, OldLastRow, _
			WorksheetTable, XMin, XMax, XMinFromY, XMaxFromY, DataMinPlus
	Else
		GetXRangeForXFromECpct WhichEquation, y0, a, b, Min, Max, EC50, Hillslope, s, Slope1, SlopeCon, OldLastRow, _
			WorksheetTable, XMin, XMax, XMinFromY, XMaxFromY, DataMinPlus
	End If
	If SolveECpct = False Then
		If XMinFromY < XMin Then XMin = XMinFromY
	Else
		XMin = XMinFromY
	End If
	If XMaxFromY > XMax Then XMax = XMaxFromY

	'Positive minimum for the 5PL equations
	Select Case WhichEquation
	Case 2,3,4 '4PL and 5PL
		If XMin <= 0 Then XMin = 10^(-100)  'not sure how solver will work with this
	End Select

	SPEquation.SetUpSolver
	SPEquation.SetSolverRange(XMin,XMax,1)

	Dim Root As Variant
	Dim Roots As String
	i=0
	For i=0 To OldLastRow-1

		Dim DataForPredict As Variant
		DataForPredict = WorksheetTable.Cell(colpredict,i)
		If MissingValue(DataForPredict) = False Then
			'Modify Ys if finding EC%s
			If SolveECpct = True Then
				DataForPredict = Min + (Max-Min)*DataForPredict/100
			End If

			SPEquation.EquationLHS(DataForPredict)
			SPEquation.Solve
			Roots = SPEquation.EquationRoots
			If Roots <> "" Then
'				Root=Val(Roots)
				Root=CDbl(Roots)
				If SolveECpct = True Then
					WorksheetTable.Cell(LastColumn,i)=DataForPredict
					WorksheetTable.Cell(LastColumn+1,i)=WorksheetTable.Cell(colpredict,i)
					WorksheetTable.Cell(LastColumn+2,i)=Root
				Else
					WorksheetTable.Cell(LastColumn,i)=Root
				End If
			ElseIf Roots = "" Then
				If DataForPredict = QNAN Then
					WorksheetTable.Cell(LastColumn,i)=""
				Else
					WorksheetTable.Cell(LastColumn,i)="No solution"
				End If
			End If
		Else
			If SolveECpct = True Then
				WorksheetTable.Cell(LastColumn,i) = DataForPredict
				WorksheetTable.Cell(LastColumn+1,i)=WorksheetTable.Cell(colpredict,i)
				WorksheetTable.Cell(LastColumn+2,i)="No Predicted Value"
			Else
				WorksheetTable.Cell(LastColumn,i)="No Predicted Value"
			End If
		End If
	Next i

Set SPEquation = Nothing

	If SolveECpct = True Then
		WorksheetTable.NamedRanges.Add("Y Values",LastColumn,0,1,-1, True)
		WorksheetTable.NamedRanges.Add("% Values",LastColumn+1,0,1,-1, True)
		WorksheetTable.NamedRanges.Add("Predicted EC% Values",LastColumn+2,0,1,-1, True)
	Else
		WorksheetTable.NamedRanges.Add("Predicted X Values",LastColumn,0,1,-1, True)
	End If

	If PlotPredicted = 1 Then
		ReDim ColumnsPerPlot(2, 1)
		If SolveECpct = True Then
			ColumnsPerPlot(0, 0) = LastColumn+2
		Else
			ColumnsPerPlot(0, 0) = LastColumn
		End If
		ColumnsPerPlot(1, 0) = 0
'		ColumnsPerPlot(2, 0) = 31999999
		ColumnsPerPlot(2, 0) = LastRow
		If SolveECpct = True Then
			ColumnsPerPlot(0, 1) = LastColumn
		Else
			ColumnsPerPlot(0, 1) = colpredict
		End If
		ColumnsPerPlot(1, 1) = 0
'		ColumnsPerPlot(2, 1) = 31999999
		ColumnsPerPlot(2, 1) = LastRow

		ReDim PlotColumnCountArray(0)
		PlotColumnCountArray(0) = 3

		'Scatterplot for predicted values
		SPPage.AddWizardPlot("Scatter Plot", "Simple Scatter", "XY Pair", ColumnsPerPlot, PlotColumnCountArray)
		SPGraph.SetAttribute(SLA_PLOTOPTIONS, 548864)
		SPGraph.SetAttribute(SLA_SELECTDROP, 1)  'X drop lines V11 (Y drop lines V12)
		SPGraph.Plots(2).DropLines(3).SetAttribute(SEA_LINETYPE, 6)
		SPGraph.Plots(2).DropLines(3).SetAttribute(SEA_THICKNESS, 20)		
		SPGraph.Plots(2).DropLines(1).SetAttribute(SEA_LINETYPE, 6)
		SPGraph.Plots(2).DropLines(1).SetAttribute(SEA_THICKNESS, 20)
		SPGraph.SetAttribute(SLA_SELECTDROP, 2)  'Y drop lines V11 (X drop lines V12)
		SPGraph.Plots(2).DropLines(2).SetAttribute(SEA_LINETYPE, 6)
		SPGraph.Plots(2).DropLines(2).SetAttribute(SEA_THICKNESS, 20)
		SPGraph.Plots(2).SetAttribute(SSA_COLOR, &H00ffffff&)
	End If
End If
End Sub
Public Sub GetXRangeForXFromY(ByVal WhichEquation As Integer,ByVal y0 As Double, _
ByVal a As Double,ByVal b As Double,ByVal Min As Double,ByVal Max As Double, _
ByVal EC50 As Double,ByVal Hillslope As Double,ByVal s As Double, ByVal Slope1 As Double, ByVal SCon As Double, _
ByVal Ndata As Long, ByVal WorksheetTable As Object, ByVal Xmin As Double, ByVal Xmax As Double, _
ByRef XMinFromY As Double,ByRef XMaxFromY As Double, ByRef DataMinPlus As Double)
'determines X range from Y values

	'Get min and max of Y values
	Dim DataMin As Double, DataMax As Double
	DataMaxAndMin colpredict+1, Ndata, DataMax, DataMin, DataMinPlus
	
	'Find X values for the min and max Ys
	Select Case WhichEquation
		Case 0 'linear y = y0 + a*x
			If a <> 0 Then
				XminFromY = (DataMin - y0)/a
				XMaxFromY = (DataMax - y0)/a
			Else
				XminFromY = Xmin
				XmaxFromY = Xmax
			End If
		Case 1 'quadratic y = bx^2 + ax + y0
			'Look for max and min of Y values for which discriminant >= 0
			'get x values for these
			Dim XMinInRange As Double, XMaxInRange As Double
			MaxAndMinYForPositiveDiscriminant colpredict+1, Ndata, a, b, y0, Xmin, Xmax, XMinInRange, XMaxInRange

		Case 2 '4PL
			'Get minimum and maximum Ys that are within {min, max} range
			Dim DataMinInRange As Double, DataMaxInRange As Double
			DataMaxAndMinInRange colpredict+1, Ndata, min, max, DataMinInRange, DataMaxInRange	
		
			'Special case two slope 5PL
			If WhichEquation = 4 Then Hillslope = Slope1
			If EC50 > 0 Then
				Dim Ratio As Double
				Ratio = (max-DataMinInRange)/(DataMinInRange-min)
				If Ratio > 0 Then
					XminFromY = EC50*(Ratio)^(1/(-Hillslope))
				Else
					XminFromY = XMin
				End If
				Ratio = (max-DataMaxInRange)/(DataMaxInRange-min)
				If Ratio > 0 Then
					XMaxFromY = EC50*((max-DataMaxInRange)/(DataMaxInRange-min))^(1/(-Hillslope))
				Else
					XmaxFromY = Xmax
				End If
			Else
				XminFromY = Xmin
				XmaxFromY = Xmax
			End If
			
			'switch if XminFromY > XmaxFromY due to negative Hillslope
			Dim XminFromYTemp As Double
			XminFromYTemp = XminFromY
			If XminFromY > XmaxFromY Then XminFromY = XmaxFromY
			XmaxFromY = XminFromYTemp			
			
		Case 3,4 '5PL
			'Get minimum and maximum Ys that are within {min, max} range
			DataMaxAndMinInRange colpredict+1, Ndata, min, max, DataMinInRange, DataMaxInRange	
		
			'Special case two slope 5PL
			If WhichEquation = 4 Then Hillslope = Slope1
			
			'Check for negative Hillslope
			Dim SignChange As Double
			If Hillslope < 0 Then
				SignChange = -1
			Else
				SignChange = 1
			End If
			
			'Increase (decrease if Hillslope <  0)by factor of 10 until exceed max EC%max
			Dim Log10 As Double
			Log10 = Log(10)
			Dim LogEC50 As Double
			LogEC50 = Log(Abs(EC50))/Log10
			Dim IntLogEC50 As Integer
			IntLogEC50 = Int(LogEC50)
			For i = 0 To 20
				Dim XTest As Double, xb As Double, PctTest As Double
				XTest = 10^(IntLogEC50+SignChange*i)
				If WhichEquation = 3 Then
					xb=EC50*(10^((1/Hillslope)*Log(2^(1/s)-1)/Log(10)))
					PctTest = min + (max-min)/(1+(XTest/xb)^(-Hillslope))^s
				Else
					Dim Cf As Double
					Cf=2*Slope1*Slope1*SCon/Abs(Slope1+Slope1*SCon)
					Dim fx As Double
					fx=1/(1+(XTest/EC50)^Cf)
					PctTest = 1/(1+fx*(XTest/EC50)^(-Slope1)+(1-fx)*(XTest/EC50)^(-Slope1*SCon))
				End If
				If PctTest > DataMaxInRange Then
					XMaxFromY = XTest
					Exit For
				End If
			Next i
			If i > 20 Then
				XmaxFromY = Xmax
			End If
			
			'Decrease by factor of 10 till less than EC%min
			For i = 0 To 20
				XTest = 10^(IntLogEC50-SignChange*i)
				If WhichEquation = 3 Then
					xb=EC50*(10^((1/Hillslope)*Log(2^(1/s)-1)/Log(10)))
					PctTest = min + (max-min)/(1+(XTest/xb)^(-Hillslope))^s
				Else
					Cf=2*Slope1*Slope1*SCon/Abs(Slope1+Slope1*SCon)
					fx=1/(1+(XTest/EC50)^Cf)
					PctTest = 1/(1+fx*(XTest/EC50)^(-Slope1)+(1-fx)*(XTest/EC50)^(-Slope1*SCon))
				End If
				If PctTest < DataMinInRange Then
					XMinFromY = XTest
					Exit For
				End If
			Next i
			If i > 20 Then
				XminFromY = Xmin
			End If
			
			'switch if XminFromY > XmaxFromY due to negative Hillslope
			XminFromYTemp = XminFromY
			If XminFromY > XmaxFromY Then XminFromY = XmaxFromY
			XmaxFromY = XminFromYTemp			
	End Select
		
	XminFromY = (1-0.1*Sgn(XminFromY))*XminFromY
	XmaxFromY = (1+0.1*Sgn(XmaxFromY))*XmaxFromY
End Sub
Public Sub GetXRangeForXFromECpct(ByVal WhichEquation As Integer,ByVal y0 As Double, _
ByVal a As Double,ByVal b As Double,ByVal min As Double,ByVal max As Double, _
ByVal EC50 As Double,ByVal Hillslope As Double,ByVal s As Double, ByVal Slope1 As Double, ByVal SCon As Double, _
ByVal Ndata As Long, ByVal WorksheetTable As Object, ByVal Xmin As Double, ByVal Xmax As Double, _
ByRef XminFromY As Double,ByRef XMaxFromY As Double, ByRef DataMinPlus As Double)
'determines X range from Y values

	'Get min and max of Y values
	Dim DataMin As Double, DataMax As Double
	DataMaxAndMin colpredict+1, Ndata, DataMax, DataMin, DataMinPlus
	
	'Find X values for the min and max Ys
	Select Case WhichEquation
		Case 2 '4PL
			'Get minimum and maximum Ys that are within {min, max} range
			Dim DataMinInRange As Double, DataMaxInRange As Double
			DataMaxAndMinInRange colpredict+1, Ndata, 0, 100, DataMinInRange, DataMaxInRange	
		
			If EC50 > 0 Then
				Dim Ratio As Double
				Ratio = (1-DataMinInRange/100)/(DataMinInRange/100)
				If Ratio > 0 Then
					XminFromY = EC50*(Ratio)^(1/(-Hillslope))	
				Else
					XminFromY = Xmin
				End If
				Ratio = (1-DataMaxInRange/100)/(DataMaxInRange/100)
				If Ratio > 0 Then
					XmaxFromY = EC50*(Ratio)^(1/(-Hillslope))
				Else
					XmaxFromY = Xmax
				End If
				
				'switch if XminFromY > XmaxFromY due to negative Hillslope
				Dim XminFromYTemp As Double
				XminFromYTemp = XminFromY
				If XminFromY > XmaxFromY Then XminFromY = XmaxFromY
				XmaxFromY = XminFromYTemp
			Else
				XminFromY = Xmin
				XmaxFromY = Xmax
			End If
		Case 3,4 '5PL
			'Special case two slope 5PL
			If WhichEquation = 4 Then Hillslope = Slope1
			
			'Check for negative Hillslope
			Dim SignChange As Double
			If Hillslope < 0 Then
				SignChange = -1
			Else
				SignChange = 1
			End If			
			
			'Increase (decrease if Hillslope < 0) by factor of 10 until exceed max EC%max
			Dim Log10 As Double
			Log10 = Log(10)
			Dim LogEC50 As Double
			LogEC50 = Log(Abs(EC50))/Log10
			Dim IntLogEC50 As Integer
			IntLogEC50 = Int(LogEC50)
			For i = 0 To 20
				Dim XTest As Double, xb As Double, PctTest As Double
				XTest = 10^(IntLogEC50+SignChange*i)
				If WhichEquation = 3 Then
					xb=EC50*(10^((1/Hillslope)*Log(2^(1/s)-1)/Log(10)))
					PctTest = 1/(1+(XTest/xb)^(-Hillslope))^s
				Else
					Dim Cf As Double
					Cf=2*Slope1*Slope1*SCon/Abs(Slope1+Slope1*SCon)
					Dim fx As Double
					fx=1/(1+(XTest/EC50)^Cf)
					PctTest = 1/(1+fx*(XTest/EC50)^(-Slope1)+(1-fx)*(XTest/EC50)^(-Slope1*SCon))
				End If
				If PctTest > DataMax/100 Then
					XMaxFromY = XTest
					Exit For
				End If
			Next i
			If i > 20 Then
				XmaxFromY = Xmax
			End If
			
			'Decrease (increase if Hillslope < 0) by factor of 10 till less than EC%min
			For i = 0 To 20
				XTest = 10^(IntLogEC50-SignChange*i)
				If WhichEquation = 3 Then
					xb=EC50*(10^((1/Hillslope)*Log(2^(1/s)-1)/Log(10)))
					PctTest = 1/(1+(XTest/xb)^(-Hillslope))^s
				Else
					Cf=2*Slope1*Slope1*SCon/Abs(Slope1+Slope1*SCon)
					fx=1/(1+(XTest/EC50)^Cf)
					PctTest = 1/(1+fx*(XTest/EC50)^(-Slope1)+(1-fx)*(XTest/EC50)^(-Slope1*SCon))
				End If
				If PctTest < DataMin/100 Then
					XMinFromY = XTest
					Exit For
				End If
			Next i
			If i > 20 Then
				XminFromY = Xmin
			End If
			
			'switch if XminFromY > XmaxFromY due to negative Hillslope
			XminFromYTemp = XminFromY
			If XminFromY > XmaxFromY Then XminFromY = XmaxFromY
			XmaxFromY = XminFromYTemp			
	End Select
	XminFromY = (1-0.1*Sgn(XminFromY))*XminFromY
	XmaxFromY = (1+0.1*Sgn(XmaxFromY))*XmaxFromY
End Sub
Public Function empty_col(column As Variant, column_end As Variant)
'Determines if a column is empty
	Dim WorksheetTable As Object
	Set WorksheetTable = ActiveDocument.CurrentDataItem.DataTable
	Dim i As Long
	Dim empty_cell As Boolean

	For i = 0 To column_end Step 3 'Change the step value to change the sampling interval.  Small sample size = Slow operation
		If WorksheetTable.Cell(column,i) = QNAN Then empty_cell = True
		If WorksheetTable.Cell(column,i) <> QNAN Then GoTo NotEmpty
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
Sub GetEmptyValues
	QNAN = "-1" & sD & "#QNAN"
	QNB = "-1" & sD & "#QNB"
End Sub
Function ColumnLength(ByVal SelectedColumn As Long, ByVal MaxRow As Long)
	'Find length of data column
	'SelectedColumn and MaxRow are 1-based
	'ColumnLength is 0-based

	Dim i As Long
	Dim WorksheetCell As Variant
	For i = 0 To MaxRow-1
		WorksheetCell = WorksheetTable.Cell(SelectedColumn-1,i)
		If WorksheetCell <> "-1.#QNAN" And WorksheetCell <>"-1,#QNAN" Then ColumnLength = i  'added "-1,#QNAN" check 8/24/04
	Next i
End Function
Function MissingValue(ByVal x As Variant) As Boolean
'Determines if the value X is a missing value

	MissingValue = False
	If VarType(x) = vbString _
		Or x = "1.#INF" Or x = "1,#INF" _
		Or x = "-1.#INF" Or x = "-1,#INF" _
		Or x = "-1.#IND" Or x = "-1,#IND" _
		Or x = "1.#QNAN" Or x = "1,#QNAN" _
		Or x = "-1.#QNAN" Or x = "-1,#QNAN" Then
			MissingValue = True
	End If
End Function
Function OppositeSign(ByVal a As Double, ByVal b As Double) As Boolean
	If (a>0 And b<0) Or (a<0 And b>0) Then
		OppositeSign = True
	Else
		OppositeSign = False
	End If
End Function
Sub ArrangeNotebookItems
'Closes the report and positions the graphpage and worksheet

'Close the report
'Get the current item (graph page presumably)
Dim CurrentItem As Object, ItemName As String
ItemName = CurrentNotebook.CurrentItem.Name
'Find its index number
Dim Index As Integer
Index = 0
Dim Item
Dim CurrentItemIndex As Integer
For Each Item In ActiveDocument.NotebookItems
	If CurrentNotebook.NotebookItems(Index).Name = ItemName Then
		CurrentItemIndex = Index
		Exit For
	Else
		Index=Index+1
	End If
Next Item
'Get report name from its index (assuming it is the next item)
Dim ReportIndex As Integer
ReportIndex = CurrentItemIndex+1
Dim ReportName As String
ReportName = CurrentNotebook.NotebookItems(ReportIndex).Name
'Close the report
CurrentNotebook.NotebookItems(ReportName).Close(False)

'Make graph page visible
SPPage.Visible = True

'Position report worksheet and graph
CurrentWorksheet.Left = 20
CurrentWorksheet.Top = 100
CurrentWorksheet.Height = 500
CurrentWorksheet.Width =1000

SPPage.Left = 30
SPPage.Top = 50
SPPage.Height = 400
SPPage.Width = 500

End Sub
Sub FormatLogXAxis
'Used only when X axis is common log (LogScale = 1)
'Sets the min, max and tick intervals and specifies base & exponent for large numbers format

Dim DataMax As Double, DataMin As Double, DataMinPlus As Double, _
RangeFromVal As Double, RangeToVal As Double, RangeTickInterval As Double
'DataMaxAndMin colxSave+1, LastRow, DataMax, DataMin, WorksheetTable
DataMaxAndMin colxSave+1, LastRow, DataMax, DataMin, DataMinPlus  'other code uses colx TBD
GetLogAxisRangeAttributes DataMin, DataMax, DataMinPlus, RangeFromVal, RangeToVal, RangeTickInterval
Dim LogFromVal As Double, LogToVal As Double
LogFromVal = 10^RangeFromVal
LogToVal = 10^RangeToVal

'Put tick locations into a worksheet column
'Dim NumTicks As Integer
NumTicks = CInt((RangeToVal - RangeFromVal)/RangeTickInterval)    'could be roundoff issues here TBD
For i = 0 To NumTicks
	Dim TickValue As Double
	TickValue = RangeFromVal + i*RangeTickInterval
	If Abs(TickValue) < 308 Then
		WorksheetTable.Cell(LastColumn+1,i)=10^TickValue
	Else
		WorksheetTable.Cell(LastColumn+1,i)= ""
	End If
Next i

'Add column title
WorksheetTable.ColumnTitle(LastColumn+1) = "Log Scale Ticks"

'Get tick marks from column
SPPage.Open
XAxis.SelectObject
XAxis.SetAttribute(SAA_SELECTLINE, 2)
XAxis.SetAttribute(SAA_SUB1OPTIONS, &H0001545c&)
XAxis.SetAttribute(SAA_TICCOLUSED, 1)
XAxis.SetAttribute(SAA_TICCOL, LastColumn+1)

'Scale the X axis
Dim XMaxTick As Variant
XAxis.SetAttribute(SAA_OPTIONS, FlagOff(SAA_FLAG_AUTORANGE))		'set range attributes to manual
XAxis.SetAttribute(SAA_OPTIONS, FlagOn(SAA_FLAG_ADVANCEDRANGEOPTS))	'turn on advanced range options
XAxis.SetAttribute(SAA_OPTIONS, FlagOff(SAA_FLAG_AUTORANGEMIN))		'set range minimum to manual
XAxis.SetAttribute(SAA_FROMVAL, LogFromVal)								'set minimum of y range = 0
XAxis.GetAttribute(SAA_TOVAL, XMaxTick)								'get new max range value
XAxis.SetAttribute(SAA_OPTIONS, FlagOff(SAA_FLAG_AUTORANGEMAX))		'set max range to manual
XAxis.SetAttribute(SAA_TOVAL, LogToVal)								'set max range value
XAxis.SetAttribute(SAA_OPTIONS, FlagOff(SAA_FLAG_ADVANCEDRANGEOPTS))'set range attributes to manual

'Change to base & exponent for large numbers
XAxis.SetAttribute(SAA_OLDSTYLEDATELABELON, 0)
XAxis.SetAttribute(SAA_TICLABELNOTATION, 5)
XAxis.SetAttribute(SAA_TICLABELAUTOPREC, 1)
XAxis.SetAttribute(SAA_TICLABELPLACES, 0)
XAxis.SetAttribute(SAA_TICLABELFACTOR, "1")

End Sub
Sub DeleteWorksheetCells(ByVal LeftMostColumn As Long, ByVal TopMostRow As Long, _
ByVal RightMostColumn As Long, ByVal BottomMostRow As Long)
'Puts blanks in worksheet
'LeftMostColumn, TopMostRow, RightMostColumn and BottomMostRow are 1-based

Dim i As Long, j As Long
For i = LeftMostColumn-1 To RightMostColumn-1
	For j = TopMostRow-1 To BottomMostRow-1
		WorksheetTable.Cell(i,j) = ""
	Next j
Next i
End Sub