Option Explicit
Dim sL
Dim sD
Global PPPath As String
Public Const msoFalse=0
Public Const msoTrue=-1
Public Const ObjectHelp = Path + "\SPW.CHM"
Public HelpID As Variant
Dim PasteFormat As Integer
Function FlagOn(flag As Long)
  FlagOn = flag Or FLAG_SET_BIT ' Use to set option flag bits on, leaving others unchanged
End Function
Function FlagOff(flag As Long)
  FlagOff = flag Or FLAG_CLEAR_BIT ' Use to set option flag bits off, leaving others unchanged
End Function
Sub Main
'Authored by Mohammad Younus, 2024
'Updated 5/28/24 (MY)
'Updated 5/28/24 (MY)
'Updated 7/09/24 (MY)
'Updated 7/30/24 (MY)
'Updated 08/06/24 (MY)
'Updated 09/03/24 (MY)
'Last updated 9/23/24 (MY)
'
sL = ListSeparator  'internationalize list separator
sD = DecimalSymbol  'internationalize decimal symbol
'HelpID = 12345			' Help ID number for this topic in SPW.CHM
Dim ErrorCheck As Integer
ErrorCheck = 0 'Display no open worksheet error message on error
On Error GoTo ErrorMsg

Dim CurrentWorksheet
CurrentWorksheet = ActiveDocument.CurrentDataItem.Name
ActiveDocument.NotebookItems(CurrentWorksheet).Open  'Opens/selects default worksheet and sets focus

'Determine the data range and define the first empty column
Dim WorksheetTable As Object
Set WorksheetTable = ActiveDocument.NotebookItems(CurrentWorksheet).DataTable
Dim LastColumn As Long
Dim LastRow As Long
LastColumn = 0
LastRow = 0 
WorksheetTable.GetMaxUsedSize(LastColumn,LastRow)
'Make sure there is data that starts in column one
If LastColumn = 0 Or empty_col(0,LastRow)=True Then GoTo ErrorMsg
'Place Worksheet into Overwrite mode
ActiveDocument.NotebookItems(CurrentWorksheet).InsertionMode = False

Dim FillColors$(2), EdgeColors$(2)

FillColors$(0)="WHITE"
FillColors$(1)="GRAY"
FillColors$(2)="DARK GRAY"

EdgeColors$(0)="BLACK"
EdgeColors$(1)="GRAY"
EdgeColors$(2)="DARK GRAY"

MacroDialog:
'Dialog for source and results columns
	Begin Dialog UserDialog 598,161,"Violin Plot " ',.DialogFunc ' %GRID:10,7,1,0
		OKButton 492,10,96,21
		CancelButton 492,40,96,21
		PushButton 492,70,96,21,"Help",.PushButton1
		GroupBox 12,7,245,75,"Column selection",.GroupBox1
		Text 25,28,128,14,"No. &data columns",.Text1
		TextBox 155,25,90,19,.x_data
		Text 25,56,120,14,"First &result column",.Text3
		TextBox 155,53,90,19,.ResultsCol
		GroupBox 270,7,206,75,"Graph dimensions",.GroupBox2
		GroupBox 12,85,454,70,"Voilin",.GroupBox3
		Text 284,28,82,14,"&Height (in)",.Text4
		TextBox 373,25,90,19,.High
		Text 284,56,84,14,"&Width (in)",.Text2
		TextBox 373,53,90,19,.Wide
		Text 22,101,152,19,"Violin gap (in) ",.Text5
		TextBox 183,101,62,18,.Space
		Text 262,101,124,15,"Edge thickness (in)",.Text6
		TextBox 399,101,62,18,.Edge
		DropListBox 140,128,105,31,FillColors(),.Colors
		Text 26,129,66,13,"Fill color:",.Text7
		Text 259,129,90,13,"Edge color:",.Text8
		DropListBox 371,128,90,31,EdgeColors(),.EdgeColors
		CheckBox 482,131,104,13,"Add Boxplot ",.CheckBox1
		
		'Text 284,97,90,14,"work column",.Text6
	End Dialog

Dim dlg As UserDialog


'Computing Default settings
WorksheetTable.GetMaxUsedSize(LastColumn,LastRow)'Reinitialize last column
If dlg.x_data = "" Then dlg.x_data = CStr(LastColumn)'Use all data by default
If dlg.Wide = "" Then dlg.Wide = "5"
If dlg.High ="" Then dlg.High ="3" +sD+ "5"
If dlg.Space ="" Then dlg.Space ="0" +sD+ "05"
If dlg.Edge ="" Then dlg.Edge ="0" +sD+ "025"
dlg.CheckBox1=0



	If dlg.ResultsCol = "" Then dlg.ResultsCol = "First Empty"

Select Case Dialog(dlg)
		Case 0 'Handles Cancel button
			GoTo Finish
		
		Case 1 'Handles Help button
		   MsgBox " This macro creates violin plots for multiple data columns. " + vbCrLf + _
	"  "+ vbCrLf + _
	"Data should contain minimum two numeric-data values in each column, without missing"+ _
    " or empty cells within the data in each column. "+ _
    "Number of data columns for plot should be less than or equal to 16."+ _
    " Violin plot(s) will not be created for column(s) with the missing data.  "+ vbCrLf + _
	" "+ vbCrLf + _
	"Column Selection. Specify the number of data columns. The data includes all columns between" + _
	" the first and inclusive the specified column." + vbCrLf + _
    "  "+ vbCrLf + _
	"Each column corresponds to a group. In addition, define the column in which to begin placing" + _
	" the macro results, otherwise macro results " + _
	"will be placed starting from the first empty column in the worksheet.  " + vbCrLf + _
	"  "+ vbCrLf + _
	"Graph Dimensions. Set the height and width of the graph in inches, using drop down list." +vbCrLf+ _
		"  "+ vbCrLf + _
	"Gap between the violins can be changed (0"+sD+"01 to 1) inch, default value is  0"+sD+"05. "+ vbCrLf + _
	"  "+ vbCrLf + _
	"Edge thickness of violins can be changed (0"+sD+"01 To 0"+sD+"1) inch, Default value Is  0"+sD+"025. "+ vbCrLf + _
	"  "+ vbCrLf + _
	"Fill and edge colors of violins can be changed using the drop downlists. "+ vbCrLf + _
	"  "+ vbCrLf + _
	"Box plot can be added as overlay plot(s) by checking the checkbox." +vbCrLf+ _
	"  "+ vbCrLf + _
	"Graph Properties can be used for desired graph customization." +vbCrLf+ _
	 "    ", vbOkOnly, "Help"
			GoTo MacroDialog
	End Select


'MsgBox  "   ker(x)=exp(-0"+sD+"5*(x/bw)^2)"

Dim Data_Range, MaxValue, MinValue, RangeSize
WorksheetTable.GetMaxUsedSize(LastColumn,LastRow)
Data_Range = WorksheetTable.GetData(0,0,LastColumn-1,LastRow-1)
MinValue = min_array(Data_Range,LastColumn-1,LastRow-1) 'See Min_Array function at end
'Debug.Print MinValue
MaxValue = max_array(Data_Range,LastColumn-1,LastRow-1) 'See Max_Array function at end
RangeSize = MaxValue - MinValue
Debug.Print MaxValue
Debug.Print RangeSize
Debug.Print MinValue
Dim DataOrder
DataOrder=CDbl(dlg.x_data)*CDbl(LastRow)


	'Clear working data
Dim Selection()
ReDim Selection(3)
Selection(0) = LastColumn
Selection(1) = 0
Selection(2) = LastColumn
Selection(3) = LastRow
ActiveDocument.CurrentDataItem.SelectionExtent = Selection
ActiveDocument.CurrentDataItem.Clear
ActiveDocument.CurrentDataItem.Goto(0,0)

'Parse the "First Empty" result
WorksheetTable.GetMaxUsedSize(LastColumn,LastRow) 'Re-initialise variables
If 	dlg.ResultsCol = "First Empty" Then
	dlg.ResultsCol = CStr(LastColumn + 1)
Else
	dlg.ResultsCol = dlg.ResultsCol
End If

'Error Handling
If IsNumeric(dlg.ResultsCol)=False Or dlg.ResultsCol="" Then
	MsgBox "You must enter a valid number for your result column",vbExclamation,"Invalid Results Column"
	GoTo MacroDialog
ElseIf IsNumeric(dlg.ResultsCol)=True Then
	If CDbl(dlg.x_data)<=0 Or CDbl(dlg.x_data)> LastColumn Then
		MsgBox "Please always start from the first column and enter the correct nth column number of the data",vbExclamation,"Incorrect Number"
		GoTo MacroDialog
	End If
	If CLng(dlg.ResultsCol) < 1 Or CDbl(dlg.ResultsCol) < (LastColumn + 1) Then
		MsgBox "You must enter a postive integer greater than the last data column for your result column",vbExclamation,"Invalid Results Column"
		GoTo MacroDialog
	End If
End If


'Limiting width, height, gap and edge of the graph
Dim GraphWidth, GraphHeight
GraphWidth=CDbl(dlg.Wide)
GraphHeight=CDbl(dlg.High)

If GraphWidth < 1  Then
	 GraphWidth = 1

ElseIf GraphWidth > (85/10) Then
       GraphWidth= (85/10)
End If

If GraphHeight < 1 Then
	GraphHeight = 1

ElseIf GraphHeight > 11 Then
		GraphHeight = 11
End If

Dim GapSize
	GapSize=CDbl(dlg.Space)
If 	GapSize < (1/100)  Then
	 GapSize =(1/100)
ElseIf 	GapSize > 1 Then
	GapSize = 1


End If

Dim  EdgeSize
     EdgeSize=CDbl(dlg.Edge)
If   EdgeSize < (1/100)   Then
	 EdgeSize =(1/100)
ElseIf EdgeSize > (1/10) Then
	EdgeSize = (1/10)

End If


Dim DataColumns

	DataColumns=CDbl(dlg.x_data)
If DataColumns > 16 Then
	MsgBox "Maximum 16 data columns are allowed",vbExclamation,"Invalid Data Selection"
	GoTo MacroDialog

End If

	 If dlg.x_data > 10 Or LastRow > 500  Then
MsgBox "Large data, it wall take a bit longer to run",vbExclamation,"large Data Selection"
	End If

'Open and run transform

Dim SPTransform As Object
Set SPTransform = ActiveDocument.NotebookItems.Add(9)
'ErrorCheck = 1 'Display transform not found error
SPTransform.Open
SPTransform.AddVariableExpression("en", dlg.x_data)
SPTransform.AddVariableExpression("space", GapSize)
'SPTransform.AddVariableExpression("work", dlg.WorkCol)
SPTransform.AddVariableExpression("res", CInt(dlg.ResultsCol)+1)
SPTransform.Text =    "st=1" + vbCrLf + _
    "'en=3" + vbCrLf + _
    "'res=8" + vbCrLf + _
    "'space=0.4 'enter between 0.2 and 1" + vbCrLf + _
    "work=600 'index of column for intermediate computations during execution " + vbCrLf + _
    "startoutputindex = res" + vbCrLf + _
    "Column_Index= data(st" +sL+ "en)" + vbCrLf + _
    "numcols = size(Column_Index)" + vbCrLf + _
    "n=256            'number of intervals for kernel density function" + vbCrLf + _
    "pi=4*arctan(1)" + vbCrLf + _
    "for j=1 to numcols do" + vbCrLf + _
    "  'column indices for locations of input and output" + vbCrLf + _
    "  var1=  Column_Index[j]             'sampled data" + vbCrLf + _
    "  var2= startoutputindex  +3*( j -1) 'computed x-values" + vbCrLf + _
    "  var3= var2 + 1                     'computed y-values" + vbCrLf + _
    "  'other parameters needed" + vbCrLf + _
    "  a=sort(col(var1))" + vbCrLf + _
    "  sigma = stddev(a)" + vbCrLf + _
    "  m = size(a)" + vbCrLf + _
    "  'compute Inter Quartile Range" + vbCrLf + _
    "   r = mod(m" +sL+ "2)" + vbCrLf + _
    "   m1 = if(r=0" +sL+ "m/2" +sL+ "(m-1)/2)" + vbCrLf + _
    "   m2 = if(r=0" +sL+ "m/2+1" +sL+ "(m-1)/2 +2)" + vbCrLf + _
    "   b1 = a[data(1" +sL+ "m1)]   " + vbCrLf + _
    "   b2 = a[data(m2" +sL+ "m)]" + vbCrLf + _
    "   IQR = median(b2)-median(b1)" + vbCrLf + _
    "  'Compute bandwidth" + vbCrLf + _
    "   ScaledIQR = IQR/1"+sD+"34" + vbCrLf + _
    "   MinVal = min({sigma" +sL+ " ScaledIQR})" + vbCrLf + _
    "   bandwidth = 0"+sD+"9*MinVal/m^(0"+sD+"2)  " + vbCrLf + _
    "   'Define gaussian kernel and range limits" + vbCrLf + _
    "   kernel(x)=exp(-0"+sD+"5*(x/bandwidth)^2)" + vbCrLf + _
    "   minimumdata = min(a)" + vbCrLf + _
    "   maximumdata = max(a)" + vbCrLf + _
    "   ZStart = minimumdata - 3"+sD+"5*bandwidth" + vbCrLf + _
    "   ZEnd = maximumdata + 3"+sD+"5*bandwidth" + vbCrLf + _
    "   DeltaZ = (ZEnd - ZStart)/n" + vbCrLf + _
    "   'Compute values of kernel density at values of data variable z" + vbCrLf + _
    "   for i =1 to n+1 do" + vbCrLf + _
    "      z = ZStart +(i-1)*DeltaZ" + vbCrLf + _
    "     cell(var2" +sL+ "i) = z" + vbCrLf + _
    "     cell(work" +sL+ "1) = 0"+sD+"0" + vbCrLf + _
    "     for k=1 to m do" + vbCrLf + _
    "        cell(work" +sL+ "1)=kernel(z-cell(var1" +sL+ "k)) + cell(work" +sL+ "1)" + vbCrLf + _
    "     end for" + vbCrLf + _
    "     KE = cell(work" +sL+ "1)/(m*bandwidth*sqrt(2*pi))" + vbCrLf + _
    "     cell(var3" +sL+ "i)=(KE*bandwidth)+(j*space)" + vbCrLf + _
    "     cell(var3+1" +sL+ "i)=(-1*KE*bandwidth)+(j*space)   " + vbCrLf + _
    "end for" + vbCrLf + _
    "end for" + vbCrLf + _
 	"cell(work" +sL+ "1)=""""" + vbCrLf
SPTransform.Execute
SPTransform.Close(False)


'Add column titles to results
Dim FirstResultColumn, ResultCount
FirstResultColumn = CLng(dlg.ResultsCol)+1
ResultCount = 1
Dim total_columns
total_columns = dlg.x_data

Do While total_columns > 0
	WorksheetTable.NamedRanges.Add("Group "+CStr(ResultCount)+" Y",CLng(FirstResultColumn)-1,0,1,-1, True)
	WorksheetTable.NamedRanges.Add("Group "+CStr(ResultCount)+" X1",CLng(FirstResultColumn),0,1,-1, True)
	WorksheetTable.NamedRanges.Add("Group "+CStr(ResultCount)+" X2",CLng(FirstResultColumn)+1,0,1,-1, True)
	total_columns = total_columns - 1
	ResultCount = ResultCount + 1
	FirstResultColumn = FirstResultColumn + 3
Loop


	Dim MyColor, MyEdColor As Long

If dlg.Colors=0 Then MyColor= RGB_WHITE
If dlg.Colors=1 Then MyColor= RGB_GRAY
If dlg.Colors=2 Then MyColor= RGB_DKGRAY

If dlg.EdgeColors=0 Then MyEdColor= RGB_Black
If dlg.EdgeColors=1 Then MyEdColor= RGB_GRAY
If dlg.EdgeColors=2 Then MyEdColor= RGB_DKGRAY


'Create Indices 

Dim Ind, Ind2 As Long
	 Ind= 2'CLng(dlg.ResultsCol)

Dim SPPage, SPGraphPage, PageItem
	Set SPPage = ActiveDocument.NotebookItems.Add(2)  'Creates graph page
	Dim PlottedColumns() As Variant
	ReDim PlottedColumns(Ind) As Variant
	Dim Index
	Index = 0
Do While Index <=Ind
	PlottedColumns(Index) = CLng(dlg.ResultsCol) + Index
	Index = Index + 1
Loop

Set SPGraphPage = 	ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0)
Set PageItem = ActiveDocument.CurrentPageItem

'Create graph
SPPage.CreateWizardGraph("Line Plot","Multiple Straight Lines","Y Many X",PlottedColumns)
PageItem.SetCurrentObjectAttribute(GPM_SETGRAPHATTR, SLA_RENDERED, 1)
PageItem.SetCurrentObjectAttribute(GPM_SETGRAPHATTR, SGA_SELECTPLOT, 1)
PageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_THICKNESS, EdgeSize*1000)
PageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLOR, MyEdColor)
PageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_TYPEREPEAT, 2)
PageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_EDGECOLORREPEAT, 2)
PageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, MyColor)
PageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLORREPEAT, 2)
PageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_AREAFILLTYPE, 1)
'PageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, &H000008a6&, MyColor)


'Adding plots iteratively

	Dim AddPIndex, Extr
	Extr=3


		AddPIndex=0

	Do While AddPIndex < (CLng(dlg.x_data))-1

		ReDim PlottedColumns(2) As Variant
			Dim Index2
		Index2 = 0

		Do While Index2 <=Ind
			PlottedColumns(Index2) =  CLng(dlg.ResultsCol) +Extr+ Index2+ 3*AddPIndex
			Index2 = Index2 + 1
		Loop

		SPPage.AddWizardPlot("Line Plot","Multiple Straight Lines","Y Many X",PlottedColumns)

        PageItem.SetCurrentObjectAttribute(GPM_SETGRAPHATTR, SLA_RENDERED, 1)
		PageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_LINETYPE, 2)
		PageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_THICKNESS, EdgeSize*1000)
		PageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLOR, MyEdColor)
		PageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_TYPEREPEAT, 2)
		PageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, &H00000362&, 0)
		PageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, MyColor)
		PageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLORREPEAT,2)
		PageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_AREAFILLTYPE, 1)
		'PageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, &H000008a6&, MyColor)

		AddPIndex=AddPIndex+1

	Loop


AddTitles:
'Add Axis titles
	Dim SPGraph, XAxis, YAxis
	Set SPGraph = 	SPPage.GraphPages(0).CurrentPageObject(GPT_GRAPH)
	SPGraph.Name = "Violin Plot"
	Set XAxis = SPGraph.Axes(0)
	Set YAxis = SPGraph.Axes(1)
	XAxis.Name = "X-Title"
	YAxis.Name = "Y-Title"



'Parse the "First Empty" result
WorksheetTable.GetMaxUsedSize(LastColumn,LastRow) 'Re-initialise variables

  ActiveDocument.CurrentDataItem.DataTable.Cell(CDbl(LastColumn), 0) = CDbl(dlg.Space)

If dlg.x_data= 1 And dlg.CheckBox1=1  Then
	

Dim ColumnsPerPlot()
ReDim ColumnsPerPlot(2, CDbl(dlg.x_data))
ColumnsPerPlot(0, 0) = CDbl(LastColumn)
ColumnsPerPlot(1, 0) = 0
ColumnsPerPlot(2, 0) = LastRow
ColumnsPerPlot(0, 1) = 0
ColumnsPerPlot(1, 1) = 0
ColumnsPerPlot(2, 1) = LastRow
Dim PlotColumnCountArray()
Redim PlotColumnCountArray(0)
PlotColumnCountArray(0) = 2
PageItem.AddWizardPlot("Box Plot", "Vertical Box Plot", "X Many Y", ColumnsPerPlot, PlotColumnCountArray)

Else

	Set SPTransform = ActiveDocument.NotebookItems.Add(9)
	'ErrorCheck = 1 'Display transform not found error
		SPTransform.Open
		SPTransform.AddVariableExpression("st1", CDbl(GapSize))
		SPTransform.AddVariableExpression("en1", CDbl(GapSize)*CDbl(dlg.x_data)+CDbl(GapSize)*(1/20))
		SPTransform.AddVariableExpression("inc1", CDbl(GapSize))
		SPTransform.AddVariableExpression("res1", CDbl(LastColumn)+1)
		SPTransform.Text =    "'st1=0.05 'Space" + vbCrLf + _
    	"'en1=3*0.05 'X-Data*Space " + vbCrLf + _
    	"'inc1=0.05  'Space " + vbCrLf + _
    	"'res1= 6' lastcolumn+1" + vbCrLf + _
   		"col(res1)=data(st1" +sL+ "en1" +sL+ "inc1)" + vbCrLf
		SPTransform.Execute
		SPTransform.Close(False)


		If dlg.CheckBox1=1 Then

			Dim i As Long
			i=1

			ReDim ColumnsPerPlot(2, CDbl(dlg.x_data))

		Do While i < CDbl(dlg.x_data)+1
			
			ColumnsPerPlot(0, 0) = CDbl(LastColumn)
			ColumnsPerPlot(1, 0) = 0
			ColumnsPerPlot(2, 0) = LastRow
			ColumnsPerPlot(0, i) = i-1
			ColumnsPerPlot(1, i) = 0
			ColumnsPerPlot(2, i) = LastRow

			i=i+1

		Loop

		ReDim PlotColumnCountArray(0)
		PlotColumnCountArray(0) = CDbl(dlg.x_data)+1
		PageItem.AddWizardPlot("Box Plot", "Vertical Box Plot", "X Many Y", ColumnsPerPlot, PlotColumnCountArray)

		PageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_BARTHICKNESS, 200)

		End If

End If

	'Add title to the X-tics column

WorksheetTable.NamedRanges.Add("All Groups X-tics",CLng(LastColumn),0,1,-1, True)

PageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, 1)
PageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, &H000006ba&, 0)
PageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_OPTIONS, 1)
PageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_TICCOLUSED, 1)
PageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_TICCOL, CDbl(LastColumn))


SPGraph.Width=(GraphWidth)*1000     ' Width of the graph

SPGraph.Height=(GraphHeight)*1000    'Height of the graph 


		'Clear the Legends
Dim SPLegend
Set SPLegend = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0)
SPLegend.SetAttribute(SGA_FLAGS, FlagOff(SGA_FLAG_AUTOLEGENDSHOW))


		'X-Axis Tick Labels

PageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_TICLABELCOLUSED, 0)
PageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_OLDSTYLEDATELABELON, 1)
PageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_OLDSTYLEDATELABEL, 4)
PageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_OLDSTYLEDATELABELFROM, 0)
PageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_OLDSTYLEDATELABELTO, 25)
PageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_OLDSTYLEDATELABELGO, 0)
PageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_OLDSTYLEDATELABELBY, 1)
PageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_OLDSTYLEDATELABELLEN, 9)


GoTo Finish

ErrorMsg:
	If ErrorCheck = 0 Then 
		MsgBox "Please use numeric data only",vbExclamation,"Text in worksheet"
	ElseIf ErrorCheck = 1 Then
		'HelpMsgBox 1234, "" + Chr(34) + "" + Chr(34) + " was not found.",vbExclamation,"SigmaPlot"
		GoTo Finish		
	End If	

Finish:

End Sub
Public Function max_array(A As Variant, maxcolumn As Long, maxrow As Long)
'Computes the maximum value of the array A consisting of maxcolumn number of
'columns and maxrow number of rows.
	Dim i, j As Long
	Dim maxval As Variant
	maxval = A(0,0)
	For i = 0 To maxcolumn
		For j = 0 To maxrow
		If A(i,j) > maxval Then 
			maxval = A(i,j)
		End If
		Next j
	Next i
	max_array = maxval
End Function
Public Function min_array(A As Variant, maxcolumn As Long, maxrow As Long)
'Computes the minimum value of the array A consisting of maxcolumn number of
'columns and maxrow number of rows.
	Dim i, j As Long
	Dim minval As Variant
	minval = A(0,0)
	For i = 0 To maxcolumn
			For j = 0 To maxrow
		If A(i,j) < minval And A(i,j) <> "-1.#QNAN" And A(i,j) <> "-1,#QNAN" Then minval = A(i,j)
		Next j
	Next i
	min_array = minval
End Function
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
HelpID = ID
	MsgBox Msg, MsgType, MsgTitle
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