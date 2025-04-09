Option Explicit
Public Const ObjectHelp = Path + "\SPW.CHM"
Public HelpID As Variant
'Macro created by Leland Jacobs, 5/20/2000. 
'This macro provides an example of using the PlotEquation (JAutoPlotEquation) object
'and its properties and methods. 
'Declare Plot Equation Here
Dim AddToGraph As Integer
Dim CreateGraph As Integer
Dim Varname As String
Dim Min As Double
Dim Max As Double
Dim Intervals As Integer
Dim Equation1 As String
Dim Equation2 As String 
Dim CoordSys As Integer
Dim CurveType As Integer
Dim PlotEquation As Object
Function FlagOn(flag As Long)
  FlagOn = flag Or FLAG_SET_BIT ' Use to set option flag bits on, leaving others unchanged
End Function
Function FlagOff(flag As Long)
  FlagOff = flag Or FLAG_CLEAR_BIT ' Use to set option flag bits off, leaving others unchanged
End Function
Sub Main
On Error GoTo NoWorksheet
	Begin Dialog UserDialog 350,280,"Polar and Parametric Equations",.DialogFunc ' %GRID:5,5,0,1
		GroupBox 10,5,150,50,"Coordinate system",.GroupBox1
		OptionGroup .Group1
			OptionButton 20,20,130,15,"&Rectangular",.OptionButton1
			OptionButton 20,35,130,15,"P&olar",.OptionButton2
		GroupBox 10,60,330,65,"Equations",.GroupBox2
		Text 30,80,50,15,"&x",.XLabel,1
		TextBox 90,75,240,20,.xVariable
		Text 30,105,50,15,"&y",.YLabel,1
		TextBox 90,100,240,20,.YVariable
		GroupBox 10,130,330,55,"Independent variable",.GroupBox3
		Text 20,145,50,12,"&Name",.NameLabel
		TextBox 20,160,55,20,.Name
		Text 90,145,80,12,"M&inimum",.MinimumLabel
		TextBox 90,160,80,20,.Minimum
		Text 185,145,90,12,"M&aximum",.MaximumLabel
		TextBox 185,160,80,20,.Maximum
		Text 280,145,55,12,"In&tervals",.IntervalsLabel
		TextBox 280,160,50,20,.Intervals
		GroupBox 170,5,170,50,"Curve description",.GroupBox4
		OptionGroup .Group2
			OptionButton 180,20,130,15,"&Single equation",.OptionButton3
			OptionButton 180,35,120,15,"&Parametric",.OptionButton4
		GroupBox 10,190,330,55,"Graphing options",.GroupBox5
		CheckBox 20,207,175,15,"&Create graph",.CheckBox1
		CheckBox 20,225,175,15,"&Add to current graph",.CheckBox2
		PushButton 130,255,100,20,"P&lot",.PushButton1
		PushButton 10,255,100,20,"Help",.PushButton2
		OKButton 240,255,100,20,.OKButton
	End Dialog
	Dim dlgvar As UserDialog
	dlgvar.Group1 = 0
	dlgvar.Group2 = 1	
	dlgvar.CheckBox1 = 1
	dlgvar.CheckBox2 = 0
	dlgvar.xVariable = "cos(2*t)*sin(t)"
	Rem sin(2*t)
	dlgvar.YVariable = "sin(3*t)*cos(t)"
	Rem cos(2*t)*sin(t)
	dlgvar.Name = "t"
	dlgvar.Minimum = CStr(-3.14159)
	dlgvar.Maximum = CStr(3.14159)
	dlgvar.Intervals = CStr(100)
    Set PlotEquation = ActiveDocument.CurrentDataItem.PlotEquation

	Dialog dlgvar
GoTo Finish:
NoWorksheet:
	HelpMsgBox 60500, "You must have a worksheet open",vbExclamation,"No Open Worksheet"
Finish:
End Sub

Rem See DialogFunc help topic for more information.
Private Function DialogFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization]
	    DlgText "OKButton","Close"
	    DlgEnable "CheckBox2",False 
	Case 2 ' Value changing or button pressed
        Select Case DlgItem$
		Case "PushButton2"
			HelpID = 60500			' Help ID number for this topic in SPW.CHM
			Help(ObjectHelp,HelpID)
			DialogFunc = True 'do not exit the dialog
		Case "PushButton1"
		    On Error GoTo ChangeValues
		    Err=0
            AddToGraph = DlgValue("CheckBox2")
            CreateGraph = DlgValue("CheckBox1")
            Varname = DlgText("Name")
            Min = DlgText("Minimum")
            Max = DlgText("Maximum")
            Intervals = DlgText("Intervals")
            Equation1 = DlgText("xVariable")
            Equation2 = DlgText("YVariable")
            CoordSys = DlgValue("Group1")
            CurveType = DlgValue("Group2") 
            Plot
            DlgEnable "CheckBox2",True
            GoTo NextPlot
            ChangeValues:
            MsgBox Err.Description, vbExclamation
            NextPlot:    
            DialogFunc = True 'do not exit the dialog
        Case "Group1"
            If DlgValue("Group1")=0 Then
              If DlgValue("Group2")=0 Then
                 DlgText "XLabel","y(x)"
              Else
                 DlgText "XLabel","x"
                 DlgText "YLabel","y"
              End If   
            Else
              If DlgValue("Group2")=0 Then
                 DlgText "XLabel","r(theta)"
              Else
                 DlgText "XLabel","r"
                 DlgText "YLabel","theta"
              End If   
            End If   
            DialogFunc = True 'do not exit the dialog
        Case "Group2"
            If DlgValue("Group2")=0 Then
              DlgEnable "YVariable",False
              DlgVisible "YLabel",False
              If DlgValue("Group1")=0 Then
                 DlgText "XLabel","y(x)"
              Else
                 DlgText "XLabel","r(theta)"
              End If
            Else
              DlgEnable "YVariable",True
              DlgVisible "YLabel",True
              If DlgValue("Group1")=0 Then
                 DlgText "XLabel","x"
                 DlgText "YLabel","y"
              Else
                 DlgText "XLabel","r"
                 DlgText "YLabel","theta"
              End If   
            End If   
        Case "CheckBox1"
            If DlgValue("CheckBox1")=0 And DlgValue("CheckBox2")=0 Then
               DlgEnable "PushButton1",False
            Else
               DlgEnable "PushButton1",True
            End If
        Case "CheckBox2"
            If DlgValue("CheckBox1")=0 And DlgValue("CheckBox2")=0 Then
               DlgEnable "PushButton1",False
            Else
               DlgEnable "PushButton1",True
            End If            
        End Select
    Case 5
        DialogFunc = True 
	End Select
End Function
Sub Plot
'Set Attributes and plot
PlotEquation.Dimension = 2
PlotEquation.EquationRHS = Equation1
PlotEquation.XEquationRHS = Equation1
PlotEquation.YEquationRHS = Equation2
PlotEquation.XVarName = Varname
PlotEquation.XRange(Min, Max)
PlotEquation.XIntervals = Intervals
PlotEquation.AddToGraph = AddToGraph
PlotEquation.CreateGraph = CreateGraph
PlotEquation.XColumn = 0
PlotEquation.YColumn = 0
PlotEquation.ZColumn = 0
PlotEquation.TrigUnit = 0
PlotEquation.SaveOption = 0
PlotEquation.CoordSystem = CoordSys
PlotEquation.CurveType = CurveType
PlotEquation.Plot
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