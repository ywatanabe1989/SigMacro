Option Explicit
Public Const ObjectHelp = Path + "\SHELF LIFE.CHM"
Dim HelpID As Variant
Dim sL$
Dim sD$
Function FlagOn(flag As Long)
    FlagOn = flag Or FLAG_SET_BIT
End Function
Function FlagOff(flag As Long)
    FlagOff = flag Or FLAG_CLEAR_BIT
End Function
'Shelf Life Macro by Mohammad Younus on 10/23/98
'Modified on 10/29/98
'Modified on 11/2/98
'Modified on 11/8/2000
'Transform added to code 2/6/01 RRM
'Internationalized 2/28/01 RRM
'Error checking added 2-13-02
'Fixed perfect data and Xmax bugs, added var2, var3 missing value checking 3-23-06 RRM
Sub Main

HelpID = 1			' Help ID number
sL = ListSeparator  'internationalize list separator
sD = DecimalSymbol  'internationalize decimal symbol

Dim CurrentWorksheet
On Error GoTo NoData
CurrentWorksheet = ActiveDocument.CurrentDataItem.Name
ActiveDocument.NotebookItems(CurrentWorksheet).Open  'Opens/select default worksheet and sets focus

'Determine the data range and define the first empty column
Dim WorksheetTable As Object
Set WorksheetTable = ActiveDocument.NotebookItems(CurrentWorksheet).DataTable
Dim LastColumn As Long
Dim LastRow As Long
LastColumn = 0
LastRow = 0
WorksheetTable.GetMaxUsedSize(LastColumn,LastRow)
If LastColumn = 0 Then GoTo EmptyWorksheet  'exit if no data

'Sort through columns and create list of columns with values in row 1
Dim Index, Column, UsedColumns$(), ColContents
ReDim UsedColumns$(LastColumn)
Index = 0
Column = 0
Do While LastColumn >=0
ColContents = WorksheetTable.Cell(Column,0)
If ColContents <> "-1"+ sD +"#QNAN" Then   'If the first cell is not empty
	UsedColumns$(Index) = CStr(Column + 1)
Else
	Index = Index - 1
End If
Column = Column + 1
Index = Index+1
LastColumn = LastColumn - 1
Loop

'Check for existence of only one data column in worksheet
If UsedColumns$(1) = "" Then GoTo OnlyOneColumn

MacroDialog:
'Dialog for source and results columns
	Begin Dialog UserDialog 371,210,"Shelf Life",.ShelfDialog ' %GRID:10,7,1,0
		GroupBox 10,4,190,84,"Specification designs",.GroupBox1
		OptionGroup .Group1
			OptionButton 20,21,160,14,"&Lower specification",.Opt1
			OptionButton 20,36,160,14,"&Upper specification",.Opt2
			OptionButton 20,51,160,14,"Lower &and upper",.Opt3
			OptionButton 20,66,100,14,"&Degradant",.Opt4
		Text 219,4,136,13,"Lo&wer limit (%):",.Text4
		TextBox 218,19,144,18,.Lower
		Text 219,45,126,13,"U&pper limit (%):",.Text5
		TextBox 216,60,144,18,.Upper
		Text 20,96,120,13,"&Time data column:",.Text1
		DropListBox 20,111,144,72,UsedColumns(),.X1_data
		Text 20,137,139,13,"Activit&y data column:",.Text2
		DropListBox 20,152,144,72,UsedColumns(),.Y1_data
		Text 219,96,120,13,"&First result column:",.Text3
		TextBox 219,111,136,18,.ResultsCol
		OKButton 162,185,90,18
		CancelButton 264,185,90,18
		PushButton 10,185,90,18,"Help",.PushButton1
	End Dialog

Dim dlg As UserDialog

'Default settings
dlg.X1_data = 0
dlg.Y1_data = 1
dlg.Lower="90"
dlg.Upper="Not Applicable"
dlg.ResultsCol = "First Empty"

Select Case Dialog(dlg)
	Case 0 'Handles Cancel button
		GoTo Finish
'	Case 1 'Handles Help button
'		'Shelf Life Help
'		Help(ObjectHelp,HelpID)
'	GoTo MacroDialog:
End Select

'Check if user selected same column
If dlg.X1_data = dlg.Y1_data Then
	MsgBox("Please select different columns for time and percent label claim", vbExclamation, "Same Column Selected")
	GoTo MacroDialog:
End If

'Check if data columns contain strings
'Change to blanks if they want, otherwise exit
If CheckForStrings(False, CVar(UsedColumns(dlg.X1_data)), CVar(UsedColumns(dlg.Y1_data)), LastRow) = True Then
	If MsgBox("Your data columns contain text items.  Would you like them to be converted to missing values and continue?", vbYesNo, "Text In Data") = vbYes Then
		CheckForStrings(True, CVar(UsedColumns(dlg.X1_data)), CVar(UsedColumns(dlg.Y1_data)), LastRow)
	Else
		GoTo Finish
	End If
End If

'Check number of valid rows >= 3
If CheckMaxRows(CVar(UsedColumns(dlg.X1_data)), CVar(UsedColumns(dlg.Y1_data)), LastRow) < 3 Then
	HelpMsgBox HelpID, "There are too few data points (< 3)",vbExclamation,"Too Few Data Points"
	GoTo Finish
End If

'Choice for Specification bound
Dim conf As Double
If dlg.Group1=0 Then     ' Lower Specification
	conf=1
ElseIf dlg.Group1=1 Then ' Upper Specification
	conf=2
ElseIf dlg.Group1=2 Then ' Lower and Upper Specifications
	conf=3
ElseIf dlg.Group1=3 Then ' Degradant Specification
	conf=2
End If

'Parse the "First Empty" result
WorksheetTable.GetMaxUsedSize(LastColumn,LastRow) 'Re-initialise variables
If 	dlg.ResultsCol = "First Empty" Then
	dlg.ResultsCol = CStr(LastColumn + 1)
Else
	dlg.ResultsCol = dlg.ResultsCol
End If

'Error handling for Results
Dim Result
Result=CDbl(dlg.ResultsCol)
If Result<=0 Or Result<(LastColumn+1)Then
	GoTo note1
End If

' Transform to Compute Shelf Life
' modified 5-27-98 to allow a non-negative regression slope
' modified 7-23-98 to allow two-sided confidence intervals
   'x data (time) is placed in column x_col (3 or more data points)
   'y data (activity) is placed in column y_col
   'results are placed in columns res through res+10 (res+11 is a work column)
      'cols res & res+1 contain the regression line
      'col res+2 contains the lower confidence line (conf_type = 1 or 3)
      'col res+3 contains the upper confidence line (conf_type = 2 or 3)
      'cols res+4 & res+5 contains the t90 value
      'col res+6 contains caution messages
      'cols res+7 - res+10 contains specification and drop lines
      'col res+11 is a working column
Dim SPTransform As Object
Set SPTransform = ActiveDocument.NotebookItems.Add(9)
'SPTransform.Name = Path + "\Macro Transforms\shelfli3.xfm" 'Retrieves from default path
SPTransform.Open
SPTransform.Text = "z1=1" + sD + "645"+vbCrLf+ _
"z3=1" + sD + "96"+vbCrLf+ _
"z=If(conf_type=3" + sL + " z3" + sL + " z1)"+vbCrLf+ _
"tol=1e-13"+vbCrLf+ _
"x1=col(x_col)"+vbCrLf+ _
"y1=col(y_col)"+vbCrLf+ _
"col(LastColumn+1)=x1"+vbCrLf+ _
"col(LastColumn+2)=y1"+vbCrLf+ _
"For row = 1 To size(x1) Do"+vbCrLf+ _
"   cell(res+11" + sL + "row)=missing(block(LastColumn+1" + sL + " row" + sL + " LastColumn+2" + sL + " row))"+vbCrLf+ _
"End For"+vbCrLf+ _
"For column = 1 to 2 do"+vbCrLf+ _
"   For row1 = 1 to size(x1) do"+vbCrLf+ _
"      cell(LastColumn+column" + sL + " row1)="""""+vbCrLf+ _
"   End For"+vbCrLf+ _
"End For"+vbCrLf+ _
"x=If(col(res+11)=0" + sL + "x1)"+vbCrLf+ _
"y=If(col(res+11)=0" + sL + "y1)"+vbCrLf+ _
"For row2 = 1 To size(x1) Do"+vbCrLf+ _
"   cell(res+11" + sL + "row2)="""""+vbCrLf+ _
"End For"+vbCrLf+ _
"For speccol = 1 To 4 Do"+vbCrLf+ _
"   For specrow = 1 To 3 Do"+vbCrLf+ _
"      cell(res+6+speccol" + sL + " specrow)="""""+vbCrLf+ _
"   End For"+vbCrLf+ _
"End For"+vbCrLf+ _
"n=size(x)"+vbCrLf+ _
"v=n-2"+vbCrLf+ _
"xbar=mean(x)"+vbCrLf+ _
"denom=total((x-xbar)^2)"+vbCrLf+ _
"alpha=total(x^2)/(n*denom)"+vbCrLf+ _
"beta=-xbar/denom"+vbCrLf+ _
"delta=1/denom"+vbCrLf+ _
"r1=total(y)"+vbCrLf+ _
"r2=total(x*y)"+vbCrLf+ _
"b0=alpha*r1+beta*r2"+vbCrLf+ _
"b1=beta*r1+delta*r2"+vbCrLf+ _
"t123=z+(z^3+z)/(4*v)+(5*z^5+16*z^3+3*z)/(96*v^2)"+vbCrLf+ _
"t4=(3*z^7+19*z^5+17*z^3-15*z)/(384*v^3)"+vbCrLf+ _
"t5=79*z^9+776*z^7+1482*z^5-1920*z^3-945*z"+vbCrLf+ _
"t6=27*z^11+339*z^9+930*z^7-1782*z^5-765*z^3+17955*z"+vbCrLf+ _
"t1=t123+t4+t5/(92160*v^4)+t6/(368640*v^5)"+vbCrLf+ _
"t=If(v=1" + sL + " If(conf_type=3" + sL + " 12" + sD + "706" + sL + " 6" + sD + "314)" + sL + " If(v=2" + sL + " If(conf_type=3" + sL + " 4" + sD + "303" + sL + " 2" + sD + "920)" + sL + " t1))"+vbCrLf+ _
"s=sqrt(total(((y-(b0+b1*x))^2))/v)"+vbCrLf+ _
"tsdel=t*s*sqrt(delta)"+vbCrLf+ _
"q=If(conf_type=1" + sL + " y0_l" + sL + " y0_u)"+vbCrLf+ _
"delta0(q) = b0-q"+vbCrLf+ _
"A = delta - (b1/(t*s))^2"+vbCrLf+ _
"B(q) = 2*beta - 2*b1*delta0(q)/(t*s)^2"+vbCrLf+ _
"C(q) = alpha - (delta0(q)/(t*s))^2"+vbCrLf+ _
"b24ac(q)=If(s<tol" + sL + "1" + sD + "0" + sL + "B(q)^2-4*A*C(q))"+vbCrLf+ _
"root1(q)=(-B(q) + sqrt(b24ac(q)))/(2*A)"+vbCrLf+ _
"root2(q)=(-B(q) - sqrt(b24ac(q)))/(2*A)"+vbCrLf+ _
"r(q)={root1(q)" + sL + "root2(q)}"+vbCrLf+ _
"rootl(q)=If(b1<0" + sL + " If(s<tol" + sL + " (q-b0)/b1" + sL + " max(If(r(q)<(q-b0)/b1" + sL + "r(q))) )" + sL + " max(r(q)))"+vbCrLf+ _
"rootu(q)=If(b1>0" + sL + " If(s<tol" + sL + " (q-b0)/b1" + sL + " max(If(r(q)<(q-b0)/b1" + sL + "r(q))) )" + sL + " max(r(q)))"+vbCrLf+ _
"maxroot=If(conf_type=1" + sL + " rootl(y0_l)" + sL + ""+vbCrLf+ _
"               If(conf_type=2" + sL + " rootu(y0_u)" + sL + " max({rootl(y0_l)" + sL + "rootu(y0_u)})))"+vbCrLf+ _
"minroot=If(conf_type=1" + sL + " rootl(y0_l)" + sL + ""+vbCrLf+ _
"               If(conf_type=2" + sL + " rootu(y0_u)" + sL + ""+vbCrLf+ _
"               If(conf_type=3 And b1>=tsdel" + sL + " rootu(y0_u)" + sL + ""+vbCrLf+ _
"               If(conf_type=3 And b1<=-tsdel" + sL + " rootl(y0_l)" + sL + " min({rootl(y0_l)" + sL + "rootu(y0_u)})))))"+vbCrLf+ _
"minx=0"+vbCrLf+ _
"maxx(q)=If((conf_type=1 And b1>=tsdel) Or (conf_type=2 And b1<=-tsdel) Or b24ac(q)<0" + sL + " 3*max(x)" + sL + " 1" + sD + "1*maxroot)"+vbCrLf+ _
"xreg=data(minx" + sL + "maxx(q)" + sL + "(maxx(q)-minx)/40)"+vbCrLf+ _
"yreg=b0+b1*xreg"+vbCrLf+ _
"term=alpha+2*beta*xreg+delta*xreg^2"+vbCrLf+ _
"conf_lim=sqrt(term)"+vbCrLf+ _
"low_conf=yreg-t*s*conf_lim"+vbCrLf+ _
"up_conf=yreg+t*s*conf_lim"+vbCrLf+ _
"col(res)=If(n<3" + sL + " """"" + sL + " xreg)"+vbCrLf+ _
"col(res+1)=If(n<3" + sL + """""" + sL + " yreg)"+vbCrLf+ _
"col(res+2)=If(n<3" + sL + " """"" + sL + " If(conf_type=1 Or conf_type=3" + sL + " low_conf))"+vbCrLf+ _
"col(res+3)=If(n<3" + sL + " """"" + sL + " If(conf_type=2 Or conf_type=3" + sL + " up_conf))"+vbCrLf+ _
"cell(res+4" + sL + "1)=If(n<3" + sL + " """"" + sL + " ""   t90  =  "")"+vbCrLf+ _
"cell(res+5" + sL + "1) = If(n<3" + sL + " """"" + sL + " If((conf_type=1 And b1>=tsdel) Or (conf_type=2 And b1<=-tsdel)" + sL + " ""+infinity""" + sL + ""+vbCrLf+ _
"                        If(b24ac(q)<0" + sL + " "" no solution""" + sL + " minroot)))"+vbCrLf+ _
"col(res+6)=If(conf_type=1" + sL + " If(b1>=0" + sL + " {""caution:""" + sL + """positive""" + sL + """slope""}" + sL + """"")" + sL + ""+vbCrLf+ _
"                  If(conf_type=2" + sL + " If(b1<=0" + sL + " {""caution:""" + sL + """negative""" + sL + """slope""}" + sL + """"")))"+vbCrLf+ _
"xrangel={-10" + sL + " rootl(y0_l)" + sL + "rootl(y0_l)}"+vbCrLf+ _
"xrangeu={-10" + sL + " rootu(y0_u)" + sL + "rootu(y0_u)}"+vbCrLf+ _
"yrangel={y0_l" + sL + " y0_l" + sL + " y0_l-40}"+vbCrLf+ _
"yrangeu={y0_u" + sL + " y0_u" + sL + " y0_u+40}"+vbCrLf+ _
"no_lowline(q)=If(conf_type=2 Or b1>=tsdel Or b24ac(q)<0" + sL + " 1" + sL + " 0)"+vbCrLf+ _
"no_upline(q)=If(conf_type=1or b1<=-tsdel Or b24ac(q)<0" + sL + " 1" + sL + " 0)"+vbCrLf+ _
"col(res+7)=If(n<3" + sL + " ""n must > 2""" + sL + " If(no_lowline(y0_l)=1" + sL + " """"" + sL + " xrangel))"+vbCrLf+ _
"col(res+8)=If(n<3" + sL + " """"" + sL + " If(no_lowline(y0_l)=1" + sL + " """"" + sL + " yrangel))"+vbCrLf+ _
"col(res+9)=If(n<3" + sL + " """"" + sL + " If(no_upline(y0_u)=1" + sL + " """"" + sL + " xrangeu))"+vbCrLf+ _
"col(res+10)=If(n<3" + sL + " """"" + sL + " If(no_upline(y0_u)=1" + sL + " """"" + sL + " yrangeu))"+vbCrLf+ _
SPTransform.AddVariableExpression("conf_type", conf)
If dlg.Group1=0  Then
SPTransform.AddVariableExpression("y0_l", dlg.Lower)
SPTransform.AddVariableExpression("y0_u", "0")
End If
If dlg.Group1=1 Then
SPTransform.AddVariableExpression("y0_l", "0")
SPTransform.AddVariableExpression("y0_u", dlg.Upper)
End If
If dlg.Group1=2 Then
SPTransform.AddVariableExpression("y0_l", dlg.Lower)
SPTransform.AddVariableExpression("y0_u", dlg.Upper)
End If
If dlg.Group1=3 Then
SPTransform.AddVariableExpression("y0_l", "0")
SPTransform.AddVariableExpression("y0_u", dlg.Upper)
End If
SPTransform.AddVariableExpression("x_col", UsedColumns(dlg.X1_data))
SPTransform.AddVariableExpression("y_col", UsedColumns(dlg.Y1_data))
SPTransform.AddVariableExpression("res", dlg.ResultsCol)
SPTransform.AddVariableExpression("LastColumn", LastColumn)
'SPTransform.RunEditor  'debug transform (opens transform dialog)
SPTransform.Execute
SPTransform.Close(False)

'Check for infinite shelf life time
'get shelf life time from worksheet
Dim Time As Double
Time = ActiveDocument.NotebookItems(CurrentWorksheet).DataTable.Cell(CLng(dlg.ResultsCol)+4,0)
If IsReallyNumeric(Time) = False Then
	HelpMsgBox HelpID, "The shelf life time is infinite.  Please check your data.", vbExclamation, "Infinite Shelf Life Time"
	GoTo Finish
End If

'Not sure this is the best method here
If Time <= 0 Then
	HelpMsgBox HelpID, "An invalid shelf life time occurred.  Please check your data.", vbExclamation, "Negative Shelf Life Time"
	GoTo Finish
End If

'Change the shelf life time subscript in the worksheet
Select Case dlg.Group1
	Case 0 'lower
		ActiveDocument.NotebookItems(CurrentWorksheet).DataTable.Cell(CLng(dlg.ResultsCol)+3,0) = "   t" + CStr(dlg.Lower) + " = "
	Case 1 'upper
		ActiveDocument.NotebookItems(CurrentWorksheet).DataTable.Cell(CLng(dlg.ResultsCol)+3,0) = "   t" + CStr(dlg.Upper) + " = "
	Case 2 'lower & upper
		If Abs(Time - ActiveDocument.NotebookItems(CurrentWorksheet).DataTable.Cell(CLng(dlg.ResultsCol)+7,1)) < 1.0e-13 Then
			ActiveDocument.NotebookItems(CurrentWorksheet).DataTable.Cell(CLng(dlg.ResultsCol)+3,0) = "   t" + CStr(dlg.Upper) + " = "
		Else
			ActiveDocument.NotebookItems(CurrentWorksheet).DataTable.Cell(CLng(dlg.ResultsCol)+3,0) = "   t" + CStr(dlg.Lower) + " = "
		End If
	Case 3 'degradant
		ActiveDocument.NotebookItems(CurrentWorksheet).DataTable.Cell(CLng(dlg.ResultsCol)+3,0) = "   t" + CStr(dlg.Upper) + " = "
End Select

'Warn about wrong slope and ask to continue
Dim Slope As String
Slope = WorksheetTable.Cell(CInt(dlg.ResultsCol)+5,1)
'for degradant
If dlg.Group1 = 3 And Slope = "negative" Then
	If MsgBox("The slope of your degradant data is negative.  Would you like to continue?", vbYesNo, "Negative Slope") = vbNo Then
		GoTo Finish
	End If
End If
'for other cases
If dlg.Group1 <> 3 And Slope = "positive" Then
	If MsgBox("The slope of your shelf life data is positive.  Would you like to continue?", vbYesNo, "Postive Slope") = vbNo Then
		GoTo Finish
	End If
End If

'Add column titles to results
Dim Results_1 As String
Dim Results_2 As String
Dim Results_3 As String
Dim Results_4 As String
Dim Results_5 As String
Dim Results_6 As String
Dim Results_7 As String
Dim Results_8 As String
Dim Results_9 As String

Results_1 = "X-Linear Reg"
Results_2 = "Y-Linear Reg"
Results_3 = "95% Conf1"
Results_4 = "95% Conf2"
Results_5 = "95% Conf3"
Results_6 = "Specif Line1"
Results_7 = "Specif Line2"
Results_8 = "Specif Line3"
Results_9 = "Specif Line4"

WorksheetTable.NamedRanges.Add(Results_1,CLng(dlg.ResultsCol)-1,0,1,-1,True)
WorksheetTable.NamedRanges.Add(Results_2,CLng(dlg.ResultsCol),0,1,-1,True)
'WorksheetTable.NamedRanges.Add(Results_3,CLng(dlg.ResultsCol)-1,0,1,-1,True)
WorksheetTable.NamedRanges.Add(Results_4,CLng(dlg.ResultsCol)+1,0,1,-1,True)
WorksheetTable.NamedRanges.Add(Results_5,CLng(dlg.ResultsCol)+2,0,1,-1,True)
WorksheetTable.NamedRanges.Add(Results_6,CLng(dlg.ResultsCol)+6,0,1,-1,True)
WorksheetTable.NamedRanges.Add(Results_7,CLng(dlg.ResultsCol)+7,0,1,-1,True)
WorksheetTable.NamedRanges.Add(Results_8,CLng(dlg.ResultsCol)+8,0,1,-1,True)
WorksheetTable.NamedRanges.Add(Results_9,CLng(dlg.ResultsCol)+9,0,1,-1,True)

'Shelf Life Plot
Dim SPPage
Set SPPage = ActiveDocument.NotebookItems.Add(2)  'Creates graph page
Dim PlottedColumns() As Variant

ReDim PlottedColumns(1)

'Plot1
PlottedColumns(0) = CLng(UsedColumns(dlg.X1_data))-1
PlottedColumns(1) = CLng(UsedColumns(dlg.Y1_data))-1

SPPage.CreateWizardGraph("Scatter Plot", _
	"Simple Scatter","XY Pair",PlottedColumns)

'Plot2
PlottedColumns(0) = CLng(dlg.ResultsCol)-1
PlottedColumns(1) = CLng(dlg.ResultsCol)

SPPage.AddWizardPlot("Line Plot", _
	"Simple Straight Line","XY Pair",PlottedColumns)

'Plot3
ReDim PlottedColumns(2)

PlottedColumns(0) = CLng(dlg.ResultsCol)-1
PlottedColumns(1) = CLng(dlg.ResultsCol)+1
PlottedColumns(2) = CLng(dlg.ResultsCol)+2

SPPage.AddWizardPlot("Line Plot", _
	"Multiple Straight Lines","X Many Y",PlottedColumns)

'Line Type Solid and Line Color  Dark-Gray for plot3
Dim SPPLot3,SPLine3
Set SPPLot3=SPPage.GraphPages(0).Graphs(0).Plots(2)
Set SPLine3=SPPLot3.Line

SPPLot3.SetAttribute(SLA_PLOTOPTIONS,SLA_FLAG_LINEON Or FLAG_SET_BIT)
SPPLot3.SetAttribute(SEA_LINETYPE,SEA_LINE_SOLID)
SPPLot3.SetAttribute(SEA_TYPEREPEAT,SOA_REPEAT_SAME)

SPLine3.SetAttribute(SEA_COLOR,RGB_DKGRAY) ' Line Color
SPLine3.SetAttribute(SEA_THICKNESS,20)

'Plot4
ReDim PlottedColumns(3)

PlottedColumns(0) = CLng(dlg.ResultsCol)+6
PlottedColumns(1) = CLng(dlg.ResultsCol)+7
PlottedColumns(2) = CLng(dlg.ResultsCol)+8
PlottedColumns(3) = CLng(dlg.ResultsCol)+9

SPPage.AddWizardPlot("Line Plot", _
	"Multiple Straight Lines","XY Pairs",PlottedColumns)

'Line Type Solid  and Line Color Gray for Plot4
Dim SPPLot4,SPLine4
Set SPPLot4=SPPage.GraphPages(0).Graphs(0).Plots(3)
Set SPLine4=SPPLot4.Line

SPPLot4.SetAttribute(SLA_PLOTOPTIONS,SLA_FLAG_LINEON Or FLAG_SET_BIT)
SPPLot4.SetAttribute(SEA_LINETYPE,SEA_LINE_SOLID)
SPPLot4.SetAttribute(SEA_TYPEREPEAT,SOA_REPEAT_SAME)

SPLine4.SetAttribute(SEA_COLOR,RGB_GRAY)' Line Color
SPLine4.SetAttribute(SEA_THICKNESS,20)

'Calculating Maximum Range of X-Axis

Dim var1 As Double
Dim var2, var3 As Variant
Dim t90_1,t90_2,t90_3 As Long
t90_1=CLng(dlg.ResultsCol)+4
var1= ActiveDocument.NotebookItems(CurrentWorksheet).DataTable.Cell(t90_1,0)
Dim Xmax As Double
Xmax=10*(Int((var1/10)+0.5)+1)
If Xmax<10 Then
	Xmax=10
Else
	Xmax=Xmax
End If

' Maximum Range of X-Axis in case "both 'Lower and Upper' Specifications"

If dlg.Group1=2 Then
	t90_2=CLng(dlg.ResultsCol)+8
	t90_3 = CLng(dlg.ResultsCol)+6
	var2=ActiveDocument.NotebookItems(CurrentWorksheet).DataTable.Cell(t90_2,1)
	var3 = ActiveDocument.NotebookItems(CurrentWorksheet).DataTable.Cell(t90_3,1)
	If var2 <> "-1.#QNAN" Then
		If var3 <> "-1.#QNAN" And var3 > var2 Then
			Xmax=10*(Int((var3/10)+0.5)+1)
		Else
			Xmax=10*(Int((var2/10)+0.5)+1)
		End If
	End If
End If

'Setting the minimum and maximum range of X-Axis and Y-Axis
Dim SPGraph, XAxis, YAxis
Dim X_Interval As Double
Set SPGraph = SPPage.GraphPages(0).Graphs(0)
Set XAxis = SPGraph.Axes(0)
Set YAxis = SPGraph.Axes(1)

'X-axis tick interval
If Xmax<20 Then
	X_Interval=5
Else
	X_Interval=10
End If

'X and Y-axis Lower and upper range
YAxis.SetAttribute(SAA_OPTIONS,SAA_FLAG_AUTORANGE Or FLAG_CLEAR_BIT)
XAxis.SetAttribute(SAA_OPTIONS,SAA_FLAG_AUTORANGE Or FLAG_CLEAR_BIT)

If dlg.Group1=0 Then ' In case Lower is selected
    YAxis.SetAttribute(SAA_FROMVAL, 85)
	YAxis.SetAttribute(SAA_TOVAL, 105)
	XAxis.SetAttribute(SAA_FROMVAL, -2)
	XAxis.SetAttribute(SAA_TOVAL, Xmax)
	XAxis.SetAttribute(SAA_MAJORFREQINDIRECT, X_Interval)
	YAxis.SetAttribute(SAA_MAJORFREQINDIRECT, 5)
ElseIf dlg.Group1=1 Then ' In case Upper is selected
    YAxis.SetAttribute(SAA_FROMVAL, 85)
	YAxis.SetAttribute(SAA_TOVAL, 115)
	XAxis.SetAttribute(SAA_FROMVAL,-2)
	XAxis.SetAttribute(SAA_TOVAL, Xmax)

	XAxis.SetAttribute(SAA_MAJORFREQINDIRECT, X_Interval)
	YAxis.SetAttribute(SAA_MAJORFREQINDIRECT, 5)

'Upper tick labels for Upper Specification Graph
    XAxis.SetAttribute(SAA_SUB2OPTIONS,SAA_SUB_MAJOR Or _
    									SAA_SUB_MAJOROUT Or _
    									SAA_SUB_MAJORLABEL Or _
    									FLAG_SET_BIT)
ElseIf dlg.Group1=2 Then  ' In case bothe Lower and Upper are seleceted
    YAxis.SetAttribute(SAA_FROMVAL, 85)
    YAxis.SetAttribute(SAA_TOVAL, 115)
	XAxis.SetAttribute(SAA_FROMVAL, -2)
	XAxis.SetAttribute(SAA_TOVAL, Xmax)

	XAxis.SetAttribute(SAA_MAJORFREQINDIRECT, X_Interval)
	YAxis.SetAttribute(SAA_MAJORFREQINDIRECT, 5)

'Upper tick labels for both Upper and Lower Specification Graph
    XAxis.SetAttribute(SAA_SUB2OPTIONS,SAA_SUB_MAJOR Or _
    									SAA_SUB_MAJOROUT Or _
    									SAA_SUB_MAJORLABEL Or _
    									FLAG_SET_BIT)
ElseIf dlg.Group1=3 Then   ' In case degradant is selected
    YAxis.SetAttribute(SAA_FROMVAL, 0)
 	YAxis.SetAttribute(SAA_TOVAL, 20)
	XAxis.SetAttribute(SAA_FROMVAL,-2)
	XAxis.SetAttribute(SAA_TOVAL, Xmax)

	XAxis.SetAttribute(SAA_MAJORFREQINDIRECT, X_Interval)
	YAxis.SetAttribute(SAA_MAJORFREQINDIRECT,5)

'Upper tick labels for Degradant Graph
    XAxis.SetAttribute(SAA_SUB2OPTIONS,SAA_SUB_MAJOR Or _
    									SAA_SUB_MAJOROUT Or _
    									SAA_SUB_MAJORLABEL Or _
    									FLAG_SET_BIT)
End If

'Clears The Legend
ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).SetAttribute(SGA_FLAGS, _
							FlagOff(SGA_FLAG_AUTOLEGENDSHOW))

'Add Graph and Axis Tiltles
SPGraph.Name = "Shelf Life Analysis"
XAxis.Name = "Time"
YAxis.Name = "Percent of Label Claim"

GoTo Finish
note1:
HelpMsgBox HelpID, "Please enter a positive column number after all occupied columns of the worksheet",vbExclamation,"Incorrect Number"
GoTo Finish

NoData:
HelpMsgBox HelpID, "You must have a worksheet open and in focus",vbExclamation,"No Open Worksheet"
GoTo Finish

EmptyWorksheet:
HelpMsgBox HelpID, "You must have shelf life data in your worksheet.",vbExclamation,"Empty Worksheet"
GoTo Finish

OnlyOneColumn:
HelpMsgBox HelpID, "You must have two columns of shelf life data in your worksheet.",vbExclamation,"Not Enough Data"

Finish:
End Sub
'Dialog Function defined
Private Function ShelfDialog(DlgItem$, Action%, SuppValue%) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgEnable "Upper",False
	Case 2 ' Value changing or button pressed
		Select Case DlgItem$
			Case "PushButton1"
				Help(ObjectHelp,HelpID)
				ShelfDialog = True 'do not exit the dialog
			Case "Group1"
				If SuppValue% = 0 Then
					DlgText "Lower", CStr(90)
					DlgEnable "Lower",True
					DlgText "Upper", "Not Applicable"'CStr(0)
					DlgEnable "Upper",False
					ShelfDialog = True ' Prevent button press from closing the dialog box
				ElseIf SuppValue% = 1 Then
					DlgText "Lower", "Not Applicable"'CStr(0)
					DlgEnable "Lower",False
					DlgText "Upper", CStr(110)
					DlgEnable "Upper",True
					ShelfDialog = True ' Prevent button press from closing the dialog box
				ElseIf SuppValue% = 2 Then
					DlgText "Lower", CStr(90)
					DlgEnable "Lower",True
					DlgText "Upper", CStr(110)
					DlgEnable "Upper",True
					ShelfDialog = True ' Prevent button press from closing the dialog box
				ElseIf SuppValue% = 3 Then
					DlgText "Lower", "Not Applicable"'CStr(0)
					DlgEnable "Lower",False
					DlgText "Upper", CStr(15)
					DlgEnable "Upper",True
					ShelfDialog = True ' Prevent button press from closing the dialog box

					'DlgText "Lower", CStr(90)
					'ShelfDialog = True
				End If

                  ShelfDialog = True

			Case "Cancel"
				End
		End Select

	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem ShelfDialog = True ' Continue getting idle actions
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
Private Function CheckForStrings(ByVal SetToMissing As Boolean, ByVal X_column As Variant, _
ByVal Y_column As Variant, ByVal LastRow As Long) As Boolean
'Determines if a string (or date) exists in the worksheet.
'If SetToMissing is True will change strings to blanks

	Dim CurrentWorksheet As Object
	Set CurrentWorksheet = ActiveDocument.CurrentDataItem
	Dim DataWorksheetTable As Object
	Set DataWorksheetTable = CurrentWorksheet.DataTable
	Dim i As Long, j As Long
	CheckForStrings = False
	Dim WorksheetColumn As Long
	For i = 0 To 1
		If i = 0 Then
			WorksheetColumn = X_column -1
		Else
			WorksheetColumn = Y_column -1
		End If
		For j = 0 To LastRow-1
			If VarType(DataWorksheetTable.Cell(WorksheetColumn,j)) = 8 Then
				If SetToMissing = False Then
					CheckForStrings = True
					Exit For
				Else
					DataWorksheetTable.Cell(WorksheetColumn,j) = ""
				End If
			End If
		Next j
		If SetToMissing = False Then
			If CheckForStrings = True Then Exit For
		End If
	Next i
End Function
Private Function CheckMaxRows(ByVal X_column As Variant, ByVal Y_column As Variant, ByVal LastRow As Long) As Long
'Determines the number of rows (1 based) after pairwise deletion

	Dim CurrentWorksheet As Object
	Set CurrentWorksheet = ActiveDocument.CurrentDataItem
	Dim DataWorksheetTable As Object
	Set DataWorksheetTable = CurrentWorksheet.DataTable
	Dim i As Long
	Dim NumValidRows As Long
	NumValidRows = 0
	For i = 0 To LastRow-1
		If IsReallyNumeric(DataWorksheetTable.Cell(X_column-1,i)) = True And IsReallyNumeric(DataWorksheetTable.Cell(Y_column-1,i)) = True Then
		NumValidRows = NumValidRows + 1
		End If
	Next i
	CheckMaxRows = NumValidRows
End Function
Function IsReallyNumeric(ByRef value As Variant) As Boolean
'Determines if worksheet cell is numeric (Isnumeric considers +inf and blank to be numeric)

    IsReallyNumeric = True
    ' weed out obvious garbage
    If IsNumeric(value) Then
        Dim temp
        Dim length As Long
        Dim i As Long
        length = Len(value)
        For i = 1 To length
            temp = Mid$(value, i, 1)
            If (temp = "-" Or temp = "+") And i = length Then
                IsReallyNumeric = False
                Exit For
            ElseIf Not IsNumeric(temp) Then
                If temp <> "E" And temp <> "e" And temp <> sD And _
                   temp <> "+" And temp <> "-" Then
                    IsReallyNumeric = False
                    Exit For
                End If
            End If
        Next i
        If IsReallyNumeric = False Then
            Exit Function
        End If
    Else
        IsReallyNumeric = False
    End If
    If Left$(value, 1) = "+" Then
        value = Right$(value, length - 1)
    End If
End Function