Option Explicit
Dim QNAN As String
Dim QNB As String
Dim sL
Dim sD
Function FlagOn(flag As Long)
  FlagOn = flag Or FLAG_SET_BIT ' Use to set option flag bits on, leaving others unchanged
End Function
Function FlagOff(flag As Long)
  FlagOff = flag Or FLAG_CLEAR_BIT ' Use to set option flag bits off, leaving others unchanged
End Function
Sub Main
'This macro uses Ancova to determine if XY Pairs lines are parallel and subsequently
'if they have the same intercept.
'Mohammad Younus, Dick Mitchell, 11/2204

	sL = ListSeparator  'international list separator
	sD = DecimalSymbol  'international decimal symbol
	GetEmptyValues
	'Get current notebook and name of user's section and data worksheet
	Dim CurrentWorksheet As Object, CurrentSection As Object
	Dim CurrentNotebook As Object
	Set CurrentNotebook = ActiveDocument
'	On Error GoTo NoOpenWorksheet
	Set CurrentWorksheet = CurrentNotebook.CurrentDataItem
	Dim DataWorksheetName As String
	DataWorksheetName = CurrentWorksheet.Name
	Dim i As Long
	For i = 0 To CurrentNotebook.NotebookItems.Count-1
		If CurrentNotebook.NotebookItems(i).Name = DataWorksheetName Then
			Set CurrentSection = CurrentNotebook.NotebookItems(i-1): Exit For
		End If
	Next i
	Dim DataSectionName As String
	DataSectionName = CurrentSection.Name

	'Find last column and row in worksheet
	Dim DataWorksheetTable As Object
	Set DataWorksheetTable = CurrentWorksheet.DataTable
	Dim LastColumn As Long, LastRow As Long
	LastColumn = 0
	LastRow = 0
	DataWorksheetTable.GetMaxUsedSize(LastColumn,LastRow)

	'Get estimate of number of data sets
	Dim NumberOfDataSets As Integer
	NumberOfDataSets = Int(LastColumn/2)

'	Dim ListArray() As String
'	ReDim ListArray(4)
'	ListArray(0)="2"
'	ListArray(1)="3"
'	ListArray(2)="4"
'	ListArray(3)="5"
'	ListArray(4)="6"

	Dim PValueSpecified() As Double
	ReDim PValueSpecified(1)
	PValueSpecified(0) = 0.05
	PValueSpecified(1) = 0.25
	Dim PValueArray() As String
	ReDim PValueArray(1)
	PValueArray(0) = CStr(PValueSpecified(0))
	PValueArray(1) = CStr(PValueSpecified(1))

	Dim DialogTextString As String
	DialogTextString = "There are "  + CStr(NumberOfDataSets) + " Data Sets to be Analyzed."
	Repeat:
	Begin Dialog UserDialog 240,168,"Parallel Line Anallysis",.DialogFunc ' %GRID:10,7,1,1
		GroupBox 10,0,210,42,"",.GroupBox2
		Text 20,11,190,28,DialogTextString,.Text1
		OKButton 165,140,65,21
		CancelButton 85,140,65,21
		PushButton 10,140,65,21,"Help",.PushButton1
		GroupBox 10,42,210,42,"",.GroupBox1
		Text 20,50,180,28,"Enter data as left adjusted XY pairs.",.Text2
		Text 10,91,130,14,"Critical P Value",.Text3
		DropListBox 10,105,95,56,PValueArray(),.PValueDropListBox
	End Dialog
	Dim dlg As UserDialog

	'Default dialog settings
	dlg.PValueDropListBox = 0

	Select Case Dialog(dlg)
		Case 0   'handles Cancel button
			End
		Case 1   'handles Help buttton

			MsgBox " Parallel Lines Test " + vbCrLf + vbCrLf + _
			"This macro tests XY data pairs for equality" + vbCrLf + _
			"of slopes.  If the slopes are found to be insignificantly" + vbCrLf + _
			"different then it tests for the equality of y intercepts." + vbCrLf + _
			"A pooled slope (intercept) is computed if the individual slopes" + vbCrLf + _
			"(intercepts) are insignificantly different."  + vbCrLf + vbCrLf + _
			"The XY data pairs must be the only data in the worksheet" + vbCrLf + _
			"and be left adjusted." + vbCrLf + vbCrLf + _
			"XY Pairs data format is required." + vbCrLf,"Help"
			GoTo Repeat
	End Select

	'Check for odd number of data columns
	If LastColumn Mod 2 > 0 Then
		MsgBox "There is an odd number of data columns in your worksheet." + vbCrLf + _
			   "XY Pairs data format has an even number of data columns.","SigmaPlot"
		GoTo Repeat
	End If

	Dim NumberOfX As Long
	NumberOfX = 0
	For i = 0 To NumberOfDataSets-1
		NumberOfX = NumberOfX + ColumnLength(2*i+1,LastRow,DataWorksheetTable)
	Next i

	'Open and run Ancova transform
	Dim SPTransform As Object
	Set SPTransform = ActiveDocument.NotebookItems.Add(9)
'	SPTransform.Name = UserPath + "\Ancova - Compare Lines.xfm"
	SPTransform.Open
	SPTransform.AddVariableExpression("a", NumberOfDataSets)
	SPTransform.AddVariableExpression("n", NumberOfX)
	SPTransform.Text="cl=2*a" + vbCrLf + _
					"p=cl+2       ' For Calculation" + vbCrLf + _
					"q=cl+1       ' For Text" + vbCrLf + _
					"s=2*cl+1    ' For transpose rows" + vbCrLf + _
					"For i=0 To cl-2 Step 2 Do" + vbCrLf + _
					"   sCol1=total(col(i+1))" + vbCrLf + _
					"   sCol2=total(col(i+2))" + vbCrLf + _
					"   spC12=total(col(i+1)*col(i+2))" + vbCrLf + _
					"  ssCol1=total(col(i+1)*col(i+1))" + vbCrLf + _
					"   ssCol2=total(col(i+2)*col(i+2))" + vbCrLf + _
					"   y2=ssCol2-sCol2*sCol2/size(col(i+1))" + vbCrLf + _
					"   xy=spC12-sCol1*sCol2/size(col(i+1))" + vbCrLf + _
					"   x2=ssCol1-sCol1*sCol1/size(col(i+1))" + vbCrLf + _
					"   b1=xy/x2" + vbCrLf + _
					"   ybar2=xy*xy/x2" + vbCrLf + _
					"   d2yx=y2-ybar2" + vbCrLf + _
					"   s2yx1=d2yx/(size(col(i+1))-2)" + vbCrLf + _
					"   Fs1=ybar2/s2yx1" + vbCrLf + _
					"   col(i+p"+sL+"1"+sL+"14)={y2 "+sL+" xy "+sL+" x2 "+sL+" b1 "+sL+" ybar2 "+sL+" size(col(i/2+1))-2 "+sL+" d2yx "+sL+" s2yx1 "+sL+" Fs1 "+sL+" total(sCol1) "+sL+" total(sCol2)"+sL+"total(spC12)"+sL+"total(ssCol1)"+sL+"total(ssCol2)}" + vbCrLf + _
					"   cell(s"+sL+"i/2+2)=y2" + vbCrLf + _
					"   cell(s"+sL+"1)=""   W""" + vbCrLf + _
					"   cell(s+1"+sL+"i/2+2)=xy" + vbCrLf + _
					"  cell(s+1"+sL+"1)=""   O""" + vbCrLf + _
					"   cell(s+2"+sL+"i/2+2)=x2" + vbCrLf + _
					"   cell(s+2"+sL+"1)=""   R""" + vbCrLf + _
					"   cell(s+3"+sL+"1)=""   K""" + vbCrLf + _
					"   cell(s+3"+sL+"i/2+2)=d2yx" + vbCrLf + _
					"   cell(s+4"+sL+"1)="" """ + vbCrLf + _
					"   cell(s+4"+sL+"i/2+2)=total(sCol1)" + vbCrLf + _
					"   cell(s+5"+sL+"1)=""   A""" + vbCrLf + _
					"   cell(s+5"+sL+"i/2+2)=total(sCol2)" + vbCrLf + _
					"   cell(s+6"+sL+"1)=""  R""" + vbCrLf + _
					"   cell(s+6"+sL+"i/2+2)=total(spC12)" + vbCrLf + _
					"   cell(s+7"+sL+"1)=""   E""" + vbCrLf + _
					"   cell(s+7"+sL+"i/2+2)=total(ssCol1)" + vbCrLf + _
					"   cell(s+8"+sL+"1)=""   A""" + vbCrLf + _
					"   cell(s+8"+sL+"i/2+2)=total(ssCol2)" + vbCrLf + _
					"End For" + vbCrLf + _
					"last=2*cl      '=q+cl-1" + vbCrLf + _
					"For k=q To last Step 2 Do" + vbCrLf + _
					"   col(k)={""  SSy = """+sL+"""  SPxy = """+sL+" ""  SSx ="""+sL+"""  byx = """+sL+"""  SSy^ ="""+sL+"""  df ="""+sL+"""  SSyx ="""+sL+"""  MSyx ="""+sL+"""  Fs =""" + vbCrLf + _
					"          "+sL+"""Total x="""+sL+" ""Total y="""+sL+" ""Total xy="""+sL+" ""Total x^2="""+sL+"""Total y^2=""}" + vbCrLf + _
					"End For" + vbCrLf + _
					"'ANCOVA" + vbCrLf + _
					"s1=2*cl+1" + vbCrLf + _
					"s2=s1+9" + vbCrLf + _
					"b=total(col(s1))     'sum y^2 within" + vbCrLf + _
					"c=total(col(s1+1))     'sum xy  within" + vbCrLf + _
					"d=total(col(s1+2))     'sum x^2 within" + vbCrLf + _
					"m=total(col(s1+3))    'sum, sum d^2yx" + vbCrLf + _
					"o=total(col(s1+4))     'sum of x" + vbCrLf + _
					"p1=total(col(s1+5))     'sum of y" + vbCrLf + _
					"q1=total(col(s1+6))     'sum of xy" + vbCrLf + _
					"r=total(col(s1+7))      'sum of x^2" + vbCrLf + _
					"s3=total(col(s1+8))       'sum of y^2" + vbCrLf + _
					"term1=r-o^2/n" + vbCrLf + _
					"term2=s3-p1^2/n" + vbCrLf + _
					"bwith=b/c" + vbCrLf + _
					"sy2bw=c^2/d" + vbCrLf + _
					"sd2yxw=b-sy2bw" + vbCrLf + _
					"s2yxw=sd2yxw/(n-a-1)" + vbCrLf + _
					"Fs=sy2bw/s2yxw" + vbCrLf + _
					"S2yx=m/(n-2*a)" + vbCrLf + _
					"ssamgb=sd2yxw-m" + vbCrLf + _
					"Msamgb=ssamgb/(a-1)" + vbCrLf + _
					"Fsamgb=Msamgb/S2yx" + vbCrLf + _
					"sxyt=q1-o*p1/n" + vbCrLf + _
					"sxyg=sxyt-c" + vbCrLf + _
					"sd2yxtotal=term2-sxyt^2/term1" + vbCrLf + _
					"sd2yxadj=sd2yxtotal-sd2yxw" + vbCrLf + _
					"S2yxadj=sd2yxadj/(a-1)" + vbCrLf + _
					"Fslast=S2yxadj/s2yxw" + vbCrLf + _
					"col(s2)={"" b-within ="""+sL+" "" SSy^ ="" "+sL+""" SSyx ="""+sL+""" df ="""+sL+" ""MSxy ="""+sL+""" Fs=""}" + vbCrLf + _
					"col(s2+1)={bwith"+sL+"sy2bw"+sL+"sd2yxw"+sL+"n-a-1"+sL+"s2yxw"+sL+"sy2bw/s2yxw}" + vbCrLf + _
					"col(s2+2"+sL+"3"+sL+"5)={m"+sL+"n-2*a"+sL+"S2yx}" + vbCrLf + _
					"col(s2+3"+sL+"3"+sL+"6)={ssamgb"+sL+"a-1"+sL+" Msamgb"+sL+"Fsamgb}" + vbCrLf + _
					"col(s2+4"+sL+"4"+sL+"7)={"" df="""+sL+""" SS="""+sL+""" MS="""+sL+""" Fs(final)=""}" + vbCrLf + _
					"Col(s2+5"+sL+"4"+sL+"7)={a-1"+sL+"sd2yxadj"+sL+"S2yxadj"+sL+"Fslast}" + vbCrLf + _
					"Col(s2+6"+sL+"4"+sL+"6)={n-a-1"+sL+"sd2yxw"+sL+"s2yxw}" + vbCrLf + _
					"col(s2+7)={""b1_pooled="""+sL+"""b0_pooled=""}" + vbCrLf + _
					"col(s2+8)={c/d"+sL+"(p1-(c/d)*o)/n}" + vbCrLf
	'SPTransform.RunEditor  'debug the transform
	SPTransform.Execute
	SPTransform.Close(False)

	'Get F values, degrees of freedom and pooled slope and intercept
	Dim FSlopeColumn As Integer, FInterceptColumn As Integer
	FSlopeColumn = 4*NumberOfDataSets + 13
	FInterceptColumn = FSlopeColumn + 2
	Dim FSlope As Double, FIntercept As Double
	Dim DF1Slope As Double, DF2Slope As Double
	Dim DF1Intercept As Double, DF2Intercept As Double
	DF1Slope = DataWorksheetTable.Cell(FSlopeColumn-1,3)
	DF2Slope = DataWorksheetTable.Cell(FSlopeColumn-2,3)
	FSlope = DataWorksheetTable.Cell(FSlopeColumn-1,5)
	DF1Intercept = DataWorksheetTable.Cell(FInterceptColumn-1,3)
	DF2Intercept = DataWorksheetTable.Cell(FInterceptColumn,3)
	FIntercept = DataWorksheetTable.Cell(FInterceptColumn-1,6)
	Dim b1Pooled As Double, b0Pooled As Double
	b1Pooled = DataWorksheetTable.Cell(FInterceptColumn+2,0)
	b0Pooled = DataWorksheetTable.Cell(FInterceptColumn+2,1)

	'Check for perfect data
	Dim S2yx, Msamgb, S2yxadj As Double
	S2yx = DataWorksheetTable.Cell(FSlopeColumn-2,4)
	Msamgb = DataWorksheetTable.Cell(FSlopeColumn-1,4)
	S2yxadj = DataWorksheetTable.Cell(FInterceptColumn-1,5)

	Dim PValueSlope As Double, PValueIntercept As Double
	If Abs(S2yx) < 1.0e-14 Then
		If Abs(Msamgb) < 1.0e-14 Then
			PValueSlope = 1
		Else
			PValueSlope = 0
		End If
		If Abs(S2yxadj) < 1.0e-14 Then
			PValueIntercept = 1
		Else
			PValueIntercept = 0
		End If
	Else
		PValueSlope = PValue(DF1Slope, DF2Slope, FSlope)
		PValueIntercept = PValue(DF1Intercept, DF2Intercept, FIntercept)
	End If

	'Create the report
	Dim PCritical As Double
	PCritical = PValueSpecified(dlg.PValueDropListBox)
	Dim SlopesTitle As String, InterceptsTitle As String
	SlopesTitle = "Test for Equality of Slopes"
	InterceptsTitle = "Test for Equality of Intercepts"
	Dim FAndPSlopeText As String, FAndPInterceptText As String
	FAndPSlopeText = "F = " + CStr(Format(FSlope, "###0.0###")) + "  DFnum = " + CStr(DF1Slope) + "  DFdenom = " + CStr(DF2Slope) + vbCrLf + _
		FormatPValue(PValueSlope)
	FAndPInterceptText = "F = " + CStr(Format(FIntercept,"###0.0###")) + "  DFnum = " + CStr(DF1Intercept) + "  DFdenom = " + CStr(DF2Intercept) + vbCrLf + _
		FormatPValue(PValueIntercept)
	Dim SignificantSlopesText As String, SignificantInterceptsText As String
	Dim NotSignificantSlopesText As String, NotSignificantInterceptsText As String
	
	SignificantSlopesText = "The line slopes are significantly different, " + FormatPValue(PValueSlope) + "." + _
	"  There is" + FormatPercentString(PValueSlope) + "chance that" +vbCrLf+ "you will be incorrect in saying that the intercepts are significantly different." + vbCrLf+vbCrLf + "Since this is the case, the line y intercepts can not " + _
	"be tested for significant difference."
	
	NotSignificantSlopesText = "The line slopes are not significantly different, " + FormatPValue(PValueSlope) + "." + _
	"  There is a " + Format(100*PValueSlope, "##")+ "%" +vbCrLf + "chance that you will be incorrect in saying that the slopes are significantly different."
	
	
	SignificantInterceptsText = "The line y intercepts are significantly different, " + FormatPValue(PValueIntercept) + "." + "There is" + FormatPercentString(PValueIntercept)+ _
	"chance that you will be incorrect In saying that the intercepts are significantly different."
	
	
	NotSignificantInterceptsText = "The line y intercepts are not significantly different, " + FormatPValue(PValueIntercept) + "." + _
	"  There is a " + Format(100*PValueIntercept, "##")+ "%" +vbCrLf+ "chance that you will be incorrect in saying that the slopes are significantly different."
	
	Dim PooledSlopeText As String, PooledInterceptText As String
	PooledSlopeText = "The data can now be pooled since the slopes are not significantly" +vbCrLf+ "different.  The slope for the pooled data is " + FormatNumber(b1Pooled)
	PooledInterceptText = "The data can now be pooled since the intercepts are not significantly" +vbCrLf+ "different.  The y intercept for the pooled data is " + FormatNumber(b0Pooled)
	Dim ReportText As String
	Dim SlopeReportTextLength As Integer
	'Text for perfect data
	Dim DifferentSlopesText_Perfect, NotDifferentSlopesText_Perfect, DifferentInterceptsText_Perfect, NotDifferentInterceptsText_Perfect As String
	DifferentSlopesText_Perfect = "All data sets are fit perfectly, or nearly so, by straight lines." + vbCrLf + _
									"The slopes are different."
	NotDifferentSlopesText_Perfect = "All data sets are fit perfectly, or nearly so, by straight lines." + vbCrLf + _
									"The slopes are identical."
	DifferentInterceptsText_Perfect = "The line y intercepts are different."
	NotDifferentInterceptsText_Perfect = "The line y intercepts are identical."

	'Check for perfect data and then
	'Test for significant slopes and intercepts
	If Abs(S2yx) < 1.0e-14 Then  'data is perfect
		If Abs(Msamgb) > 1.0e-14 Then  'slopes different
			ReportText = SlopesTitle + vbCrLf + _
						 DifferentSlopesText_Perfect
		Else                           'slopes are identical
			ReportText = SlopesTitle + vbCrLf + _
						 NotDifferentSlopesText_Perfect + vbCrLf + vbCrLf + _
						 PooledSlopeText + vbCrLf + vbCrLf
			SlopeReportTextLength = Len(ReportText)
			If PValueIntercept = 1 Then  'intercepts are identical
				ReportText = ReportText + _
							 InterceptsTitle + vbCrLf + _
							 NotDifferentInterceptsText_Perfect + vbCrLf + vbCrLf + _
							 PooledInterceptText
			Else
				ReportText = ReportText + _
							 InterceptsTitle + vbCrLf + _
							 DifferentInterceptsText_Perfect
			End If
		End If
	Else
		If PValueSlope <= PCritical Then  'different slopes
			ReportText = SlopesTitle  + vbCrLf + _
						 FAndPSlopeText + vbCrLf + vbCrLf + _
						 SignificantSlopesText
			SlopeReportTextLength = Len(ReportText)
		Else                              'not different slopes
			ReportText = SlopesTitle + vbCrLf + _
						 FAndPSlopeText + vbCrLf + vbCrLf + _
						 NotSignificantSlopesText + vbCrLf + vbCrLf + _
						 PooledSlopeText + vbCrLf + vbCrLf
			SlopeReportTextLength = Len(ReportText)
			ReportText = ReportText + _
						 InterceptsTitle + vbCrLf + _
						 FAndPInterceptText + vbCrLf + vbCrLf
			'Test for significant intercepts
			If PValueIntercept <= PCritical Then  'different intercepts
				ReportText = ReportText + SignificantInterceptsText
			Else                                  'not different intercepts
				ReportText = ReportText + NotSignificantInterceptsText + vbCrLf + vbCrLf + _
				PooledInterceptText
			End If
		End If
	End If

	ActiveDocument.NotebookItems(DataWorksheetName).Open  'add report to this section
	Dim SPReport As Object
	Set SPReport = ActiveDocument.NotebookItems.Add(CT_REPORT)
	'Name the report
	On Error GoTo ReplicateReportNameError
	Dim iError As Long
	iError = 1
	Dim ReportName As String
	NameReport: If iError = 1 Then
		ReportName = "Parallel Lines Analysis " + CStr(iError)
	Else
		ReportName = "Parallel Lines Analysis " + CStr(iError)
	End If
	SPReport.Name = ReportName

	'Bold and underline the report titles
	SPReport.ChangeDefaultFont
	SPReport.Text = ReportText
	Dim Selection() As Variant
	ReDim Selection(1)
	Selection(0) = 0
	Selection(1) = 27
	SPReport.SelectionExtent = Selection
	SPReport.BoldFont
	SPReport.UnderlineFont

	Dim Fudge As Integer
	If Abs(S2yx) < 1.0e-14 Then
		Fudge = 7
	Else
		Fudge = 10
	End If
	If PValueSlope <= PCritical Then
	Else
		'Test for significant intercepts
		If PValueIntercept <= PCritical Then
			Selection(0) = SlopeReportTextLength - Fudge
			Selection(1) = SlopeReportTextLength + 31 - Fudge
			SPReport.SelectionExtent = Selection
			SPReport.BoldFont
			SPReport.UnderlineFont
		Else
			Selection(0) = SlopeReportTextLength - Fudge
			Selection(1) = SlopeReportTextLength + 31 - Fudge
			SPReport.SelectionExtent = Selection
			SPReport.BoldFont
			SPReport.UnderlineFont
		End If
	End If
	SPReport.Close(True)

	'Clear the workspace
	Set SPTransform = ActiveDocument.NotebookItems.Add(9)
	SPTransform.Open
	SPTransform.AddVariableExpression("ndatasets", NumberOfDataSets)
	SPTransform.AddVariableExpression("firstcolumn", LastColumn+1)
	SPTransform.Text="For i = 1 To 2*ndatasets Do" + vbCrLf + _
					"   For j = 1 To 14 Do" + vbCrLf + _
					"      cell(firstcolumn+i-1"+sL+"j)=""""" + vbCrLf + _
					"   End For" + vbCrLf + _
					"End For" + vbCrLf + _
					"For i1 = 1 To 9 Do" + vbCrLf + _
					"   For j1 = 1 To ndatasets+1 Do" + vbCrLf + _
					"      cell(firstcolumn+2*ndatasets+i1-1"+sL+"j1)=""""" + vbCrLf + _
					"   End For" + vbCrLf + _
					"End For" + vbCrLf + _
					"For i2 = 1 To 9 Do" + vbCrLf + _
					"   For j2 = 1 To 7 Do" + vbCrLf + _
					"      cell(firstcolumn+2*ndatasets+9+i2-1"+sL+"j2)=""""" + vbCrLf + _
					"   End For" + vbCrLf + _
					"End For" + vbCrLf
	'SPTransform.RunEditor  'debug the transform
	SPTransform.Execute
	SPTransform.Close(False)

	'Create XY Pairs graph with regression lines
	ActiveDocument.NotebookItems.Add(CT_GRAPHICPAGE)
	Dim ColumnsPerPlot()
	ReDim ColumnsPerPlot(2, LastColumn-1)
	For i = 0 To LastColumn-1
		ColumnsPerPlot(0, i) = i
		ColumnsPerPlot(1, i) = 0
		ColumnsPerPlot(2, i) = LastRow
	Next i
	Dim PlotColumnCountArray()
	ReDim PlotColumnCountArray(0)
	PlotColumnCountArray(0) = LastColumn
	ActiveDocument.CurrentPageItem.CreateWizardGraph("Scatter Plot", "Multiple Scatter", "XY Pairs", ColumnsPerPlot, PlotColumnCountArray)

	'Add linear regression lines
	Dim SPGraph As Object
	Set SPGraph = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0)
	SPGraph.Plots(0).Line.SetAttribute(SLA_SELECTFUNC, 1)            'select line
	SPGraph.Plots(0).Line.SetAttribute(SLA_REGRORDER, 1)
	SPGraph.Plots(0).Line.SetAttribute(SLA_SELECTFUNC, 1)
	SPGraph.Plots(0).Functions(SLA_FUNC_REGR).SetAttribute(SEA_THICKNESS, 25)
	SPGraph.Plots(0).Functions(SLA_FUNC_REGR).SetAttribute(SLA_REGROPTIONS, &H00016014&)
	SPGraph.SetAttribute(SLA_PLOTOPTIONS, &H00000021&)

	'Change the graph title
	SPGraph.Name = "Parallel Line Analysis"

	'Name the graph item in the notebook
	On Error GoTo ReplicateGraphNameError
	iError = 1
	Dim GraphName As String
	NameGraph: If iError = 1 Then
		GraphName = "Parallel Lines Graph " + CStr(iError)
	Else
		GraphName = "Parallel Lines Graph " + CStr(iError)
	End If
	ActiveDocument.CurrentPageItem.Name = GraphName

	GoTo Finish

	NoOpenWorksheet:
	MsgBox "You must have a worksheet open and in focus",vbExclamation,"No Open Worksheet"
	GoTo Finish

	NumericEntryNeeded:
	MsgBox "Enter the number of XY data sets.", vbExclamation,"SigmaPlot"
	GoTo Repeat

	ReplicateReportNameError:
	Select Case Err.Number	'evaluate error number.
		Case 65535	        'duplicate report name
			iError = iError + 1
		Case Else           'handle other situations here
			MsgBox(Err.Description + " (" + CStr(Err.Number) + ")" + " in Main", 16, "Parallel Lines Analysis Macro")
	End Select
	Resume NameReport

	ReplicateGraphNameError:
	Select Case Err.Number	'evaluate error number.
		Case 65535	        'duplicate report name
			iError = iError + 1
		Case Else           'handle other situations here
			MsgBox(Err.Description + " (" + CStr(Err.Number) + ")" + " in Main", 16, "Parallel Lines Analysis Macro")
	End Select
	Resume NameGraph

	Finish:

End Sub
Sub GetEmptyValues
	QNAN = "-1" & sD & "#QNAN"
	QNB = "-1" & sD & "#QNB"
End Sub
Function ColumnLength(ByVal SelectedColumn As Long, ByVal MaxRow As Long, ByVal DataWorksheetTable As Object) As Long
	'Find length of data column
	'SelectedColumn, MaxRow are 1-based

	Dim i As Long
	For i = 0 To MaxRow-1
		If DataWorksheetTable.Cell(SelectedColumn-1,i) <> "-1.#QNAN" Then ColumnLength = i
	Next i
	ColumnLength = ColumnLength + 1  '1-based
End Function
Function PValue(ByVal DF1 As Integer, ByVal DF2 As Integer, ByVal F As Double) As Double

	'One sided test
	Dim ErrorCode As Integer

	'Clamp F if negative (avoids "bad argument..." error message
	If F < 0 Then F = 0

	PValue = BetaI(0.5*DF2, 0.5*DF1, DF2/(DF2+DF1*F), ErrorCode)
	If PValue > 1.0 Then PValue = 2.0 - PValue
End Function
Function BetaI(ByVal A As Double, ByVal B As Double, ByVal X As Double, ByRef ErrorCode As Integer) As Double

	'Returns the incomplete beta function
	ErrorCode = 0
	If X < 0.0 Or X > 1.0 Then
		MsgBox("bad argument X in betaI")
		ErrorCode = 1
	End If
	If ErrorCode = 1 Then GoTo Finish
	Dim BT As Double
	If X = 0.0 Or X = 1.0 Then
		BT = 0.0
	Else
		BT = Exp(GammaLn(A + B) - GammaLn(A) - GammaLn(B) + A*Log(X) + B*Log(1.0 - X))
	End If
	If X < (A + 1.0)/(A + B + 2.0) Then
		BetaI = BT*BetaCf(A, B, X)/A
	Else
		BetaI = 1.0 - BT*BetaCf(B, A, 1.0 - X)/B
	End If
	Finish:
End Function
Function BetaCf(ByVal A As Double,ByVal B As Double,ByVal X As Double) As Double

	Dim ITMax As Integer
	ITMax = 200
	Dim Eps As Double
	Eps = 0.0000003
	Dim AM As Double, BM As Double, AZ As Double
	Dim QAB As Double, QAP As Double, QAM As Double
	Dim BZ As Double, EM As Double, TEM As Double
	Dim AOld As Double, D As Double
	Dim AP As Double, BP As Double, APP As Double, BPP As Double
	AM = 1.0
	BM = 1.0
	AZ = 1.0
	QAB = A + B
	QAP = A + 1.0
	QAM = A - 1.0
	BZ = 1.0 - QAB*X/QAP
	Dim M As Integer
	For M = 1 To ITMax
		EM = Int(M)
		TEM = EM + EM
		D = EM*(B - M)*X/((QAM + TEM)*(A + TEM))
		AP = AZ + D*AM
		BP = BZ + D*BM
		D = -(A + EM)*(QAB + EM)*X/((A + TEM)*(QAP + TEM))
		APP = AP + D*AZ
		BPP = BP + D*BZ
		AOld = AZ
		AM = AP/BPP
		BM = BP/BPP
		AZ = APP/BPP
		BZ = 1.0
		If Abs(AZ - AOld) < Eps*Abs(AZ) Then Exit For
	Next M
	If Abs(AZ - AOld) >= Eps*Abs(AZ) Then
		MsgBox("A or B too big, or ITMax too small")
	Else
		BetaCf = AZ
	End If
End Function
Function GammaLn(ByVal XX As Double) As Double

	Dim Coeff() As Variant
	ReDim Coeff(5)      '6 coefficients
	Coeff = Array(76.18009173d0, -86.50532033d0, 24.01409822d0, -1.231739516d0, 0.120858003d-2, -0.536382d-5)
	Dim STP As Double, Half As Double, One As Double, FPF As Double
	STP = 2.50662827465d0
	Half = 0.5d0
	One = 1.0d0
	FPF = 5.5d0
	Dim X As Double, TMP As Double, SER As Double
	X = XX - One
	TMP = X + FPF
	TMP = (X + Half)*Log(TMP) - TMP
	SER = One
	Dim j As Integer
	For j = 0 To 5
		X = X + One
		SER = SER + Coeff(j)/X
	Next j
	GammaLn = TMP + Log(STP*SER)
End Function
Function FormatPValue(ByVal p As Double) As String

'	Dim pstring As String, pleft As String, pright As String
	p = Format(p,"0.00000")
	If p < 0.0001 Then
		FormatPValue = "P < " + CStr(0.0001)
	Else
		FormatPValue = "P = " + CStr(Format(P,"0.0000"))
	End If
End Function
Function FormatPercentString(ByVal P As Double) As String
	
	P = Format(P,"0.00000")
	If P < 0.0001 Then
		FormatPercentString = " less than a 0.01% "
	Else
		FormatPercentString = " a " + CStr(Format(100*P,"0.00")) + "% "
	End If
End Function
Rem See DialogFunc help topic For more information.
Private Function DialogFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Select Case Action%
		Case 1 ' Dialog box initialization
'			DlgEnable "NoSets",False
		Case 2 ' Value changing or button pressed
		Case 3 ' TextBox or ComboBox text changed
		Case 4 'Focus changed
		Case 5 ' Idle
	End Select
End Function
Function FormatNumber(ByVal Number As Double) As String

	'Changes format to exponential for large and small numbers
	If Abs(Number) > 1.0e-06 And Number < 1.0e06 Then
		FormatNumber = CStr(Format(Number,"###0.0###"))
	Else
		FormatNumber = CStr(Format(Number,"0.0###e+00"))
	End If
End Function