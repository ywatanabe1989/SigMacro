' Modified by Frederick Cabasa 5/15/01
' Updated 8/31/01 John Kuo
' Modified 9/13/01 Barbara Althoff
Public Const ObjectHelp = path + "\SPW.CHM"
Public HelpID As Variant
Dim NotebookList$(), GraphPages$()
Dim CurrentNotebook, SPPage, SPGraph, WordApp, WordDoc, WordRange, WordTable, WordFrame As Object
Dim i, Index As Integer
Dim UserHeight, UserWidth, UserTop, UserLeft As Double
Dim Unit As Double
Dim Inches, Square, InFrame, BadValue As Boolean
Option Explicit
Function FlagOn(flag As Long)
  FlagOn = flag Or FLAG_SET_BIT ' Use to set option flag bits on, leaving others unchanged
End Function
Function FlagOff(flag As Long)
  FlagOff = flag Or FLAG_CLEAR_BIT ' Use to set option flag bits off, leaving others unchanged
End Function
Sub Main
'start Word and instantiate Word Application Object
	On Error GoTo NoWordDoc
	Set WordApp=Word.Application

	' if the doc is from Outlook do not continue - Word is confused
	If WordApp.ActiveDocument.Kind <> 2 Then
		Set WordDoc = WordApp.ActiveDocument
	Else
		GoTo OutlookDocError
	End If

	' got the doc so turn off the NoWordDoc error handler
	On Error GoTo 0

	'Initialize Variables
	Inches = True
	Square =  True
	Unit = 1
	Dim Frame$()
	ReDim Frame$(1)
'	Frame(0)="at cursor position"
'	Frame(1)="in text frame"
	Frame(0)="in floating frame"
	Frame(1)="in-line with text"

	Set CurrentNotebook=ActiveDocument
	GetPages 'Initialize page list

MacroDialog:	
	Begin Dialog UserDialog 390,351,"Insert Graphs into Word",.DialogFunc ' %GRID:10,7,0,0
		CancelButton 290,323,90,20,.CancelButton
		PushButton 191,323,90,20,"&Insert",.Insert
		PushButton 11,323,90,20,"Help",.Help
		GroupBox 10,6,370,159,"Figure attributes",.GroupBox2
		GroupBox 20,21,170,77,"Size",.GroupBox3
		Text 35,40,80,14,"Max &width",.Text2
		TextBox 108,37,64,20,.Width
		Text 35,70,80,14,"Min &height",.Text3
		TextBox 108,67,64,21,.Height
		GroupBox 200,21,170,77,"Offset from position",.GroupBox4
		Text 212,40,62,14,"From &top",.Text4
		TextBox 288,37,64,21,.Top
		Text 212,70,62,14,"From &left",.Text5
		TextBox 288,67,64,21,.Left
		GroupBox 20,101,170,54,"Units",.UnitGroup
		OptionGroup .Units
			OptionButton 30,115,130,14,"I&nches",.inches
			OptionButton 30,133,136,14,"&Centimeters",.centimeters
		Text 202,107,130,14,"Place graph:",.Text10
		DropListBox 200,122,170,59,Frame(),.Target
'		Text 10,114,90,14,"&Notebook",.Text10
'		DropListBox 120,112,260,70,NotebookList(),.SelectedNotebook
		Text 10,172,270,14,"&Graph pages in current notebook:",.Text11
		ListBox 10,188,370,65,GraphPages(),.Pages
		GroupBox 10,255,370,59,"",.GroupBox1
		Text 19,266,350,42,"All graphs and objects on the selected SigmaPlot page will be placed in a frame at the beginning of the current Word line, offset by the specified distance.",.Text1
	End Dialog
	Dim dlg As UserDialog
	
	Select Case Dialog(dlg)
	Case 0 'Handles Cancel button
		GoTo Finish
	End Select

GoTo Finish

OutlookDocError:
HelpMsgBox 80001, "You have an e-mail document open. Please close it and open a Word document.",vbExclamation,"SigmaPlot"
GoTo Finish

NoWordDoc:
HelpMsgBox 80001, "You must have a Word document open.",vbExclamation,"SigmaPlot"
'MsgBox "You must have a Word document open.",vbExclamation,"SigmaPlot"
GoTo Finish
NoPage:
HelpMsgBox 80001, "You must have a graph page open.",vbExclamation,"SigmaPlot"
'MsgBox "You must have a graph page open.",vbExclamation,"SigmaPlot"
Finish:
End Sub
Rem See DialogFunc Help topic for more information.
Private Function DialogFunc(DlgItem$, Action%, SuppValue%) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
		InFrame = True
		DlgValue "Pages",0
		Set SPPage	= CurrentNotebook.NotebookItems(DlgText("Pages"))

		If WordApp.Version = "8.0" Or WordApp.Version = "8.0a" Then	DlgEnable "Target", False 'Check for Word97
		Dim defw, defh As Double
		defw=5
		defh=3.5
		If ListSeparator = ";" Then 'Check for international version
			DlgValue "Units",1
			Unit = 2.54
			Inches = False
			defh=3.543307087
			defw=5.118110237
		End If

	'*********************************************
	'* Change this setting to change the default *
	'* size and position of the figure           *
	'*********************************************
	DlgText "Width", CStr(defw*Unit)
	UserWidth = CStr(defw*Unit)
	DlgText "Height", CStr(defh*Unit)
	UserHeight = CStr(defh*Unit)
	DlgText "Top", "0"
	UserTop = 0
	DlgText "Left", "0"
	UserLeft = 0

	Case 2 ' Value changing or button pressed
        Select Case DlgItem$
		Case "Help"
			HelpID = 80001			' Help ID number for this topic in SPW.CHM
			Help(ObjectHelp,HelpID)
			DialogFunc = True 'do not exit the dialog
'		Case "SelectedNotebook"
'			CurrentNotebook=NotebookList(DlgValue("SelectedNotebook"))
'			GetPages
'			DialogFunc = True 'do not exit the dialog
		Case "Pages"
			Set SPPage = CurrentNotebook.NotebookItems(GraphPages(DlgValue("Pages")))
		Case "Units"
				If SuppValue% = 1 Then   'Centimeters
					Unit = 2.54
					If Inches = True Then
						DlgText "Width", CStr(UserWidth*Unit)
						UserWidth = UserWidth*Unit
						DlgText "Height", CStr(UserHeight*Unit)
						UserHeight = UserHeight*Unit
						DlgText "Top", CStr(UserTop*Unit)
						UserTop = UserTop*Unit
						DlgText "Left", CStr(UserLeft*Unit)
						UserLeft = UserLeft*Unit
					End If
					Inches = False
				End If
				If SuppValue% = 0 Then   'Inches
					If Inches = False Then
						DlgText "Width", CStr(UserWidth/Unit)
						UserWidth = UserWidth/Unit
						DlgText "Height", CStr(UserHeight/Unit)
						UserHeight = UserHeight/Unit
						DlgText "Top", CStr(UserTop/Unit)
						UserTop = UserTop/Unit
						DlgText "Left", CStr(UserLeft/Unit)
						UserLeft = UserLeft/Unit
					End If
					Unit = 1
					Inches = True
				End If
'		Case "Wrap"
'				If SuppValue% = 1 Then    'Inline
'					Square = False
'				End If
		Case "Target"
			If SuppValue% = 1 Then
				InFrame = False
'				DlgText "Text1", "All graphs and objects on the selected SigmaPlot page will be placed at the current Word cursor location, offset by the specified distance."
				DlgText "Text1", "All graphs and objects on the selected SigmaPlot page will be placed in-line at the current Word cursor location."
				DlgEnable "Text4", False
				DlgEnable "Text5", False
				DlgEnable "Top", False
				DlgEnable "Left", False
			Else
				InFrame = True
'				DlgText "Text1", "All graphs and objects on the selected SigmaPlot page will be placed into a text frame at the specified distance from the upper left corner of your Word page."
				DlgText "Text1", "All graphs and objects on the selected SigmaPlot page will be placed in a frame at the beginning of the current Word line, offset by the specified distance."
				DlgEnable "Text4", True
				DlgEnable "Text5", True
				DlgEnable "Top", True
				DlgEnable "Left", True
			End If
		Case "Insert"
			'Copy Graph
			SPPage.Open
			Set SPGraph = SPPage.GraphPages(0)
			If SPGraph.ChildObjects.Count = 0 Then
				MsgBox "Your graph page is empty.",vbExclamation,"SigmaPlot"
				GoTo NextPage
			End If
			SPPage.SelectAll
			SPPage.Copy

			'Re-instantiate Word in case doc focus has changed
			Set WordApp=Word.Application
			Set WordDoc = WordApp.ActiveDocument

		If WordApp.Version <> "8.0" And WordApp.Version <> "8.0a" Then
			Dim CurrentStart, CurrentEnd As Long
		    Selection.MoveRight Unit:=wdCharacter, Count:=1
		    Selection.MoveLeft Unit:=wdCharacter, Count:=1
			CurrentStart = WordApp.Selection.Start
			CurrentEnd = WordApp.Selection.End
			Set WordRange = WordDoc.Range(Start:=CurrentStart,End:=CurrentStart)
			Set WordFrame = WordDoc.Frames.Add(WordRange)
			WordFrame.Width = CDbl(DlgText("Width"))*72/Unit
			WordFrame.Height = CDbl(DlgText("Height"))*72/Unit
			WordFrame.VerticalPosition = CDbl(DlgText("Top"))*72/Unit
			WordFrame.HorizontalPosition = CDbl(DlgText("Left"))*72/Unit
			WordFrame.Select
		End If

		If WordApp.Version = "8.0" Or WordApp.Version = "8.0a" Then
			'Insert and size textbox
			Dim FrameWidth, FrameHeight, FrameTop, FrameLeft As Long
			FrameWidth = CDbl(DlgText("Width"))*72/Unit
			FrameHeight = CDbl(DlgText("Height"))*72/Unit
			FrameTop = CDbl(DlgText("Top"))*72/Unit
			FrameLeft = CDbl(DlgText("Left"))*72/Unit
			Set WordFrame = WordDoc.Shapes.AddTextbox(1, FrameLeft, FrameTop, FrameWidth, FrameHeight)
			WordFrame.Select
			WordFrame.Line.Visible = 0
		End If

			WordApp.ActiveWindow.ActivePane.View.Type = 3
			WordApp.Selection.Paste

		If InFrame = False Then
			'Define current cursor position
			WordFrame.Delete
			WordApp.Selection.Delete
'		Else
			'Insert and size textbox
'			Dim FrameWidth, FrameHeight, FrameTop, FrameLeft As Long
'			FrameWidth = CDbl(DlgText("Width"))*72/Unit
'			FrameHeight = CDbl(DlgText("Height"))*72/Unit
'			FrameTop = CDbl(DlgText("Top"))*72/Unit
'			FrameLeft = CDbl(DlgText("Left"))*72/Unit
'			Set WordFrame = WordDoc.Shapes.AddTextbox(1, FrameLeft, FrameTop, FrameWidth, FrameHeight)
'			WordFrame.Select
'			WordFrame.Line.Visible = 0
	    End If

			DlgText "CancelButton", "Close"
		    NextPage:
		    DialogFunc = True
	    End Select

			'Change Background Format to Inline
			'If Square = False Then
			'	WordDoc.Visible = True
			'	WordDoc.Shapes(1).WrapFormat.Type = wdWrapTight
			'End If
	Case 3
        Select Case DlgItem$
        'Data validations
        Case "Width"
        	If DlgText("Width") <> "" And IsNumeric(DlgText("Width"))=True Then
				Validate(CDbl(DlgText("Width")))
				If BadValue = True Then DlgText "Width",CStr(5*Unit)
			ElseIf IsNumeric(DlgText("Width"))=False Then
				MsgBox "You must enter a number greater than 0 or less then 14in/36cm", vbInformation, "Bad value"
				DlgText "Width",CStr(5*Unit)
			End If
			BadValue = False
			DialogFunc = True
			UserWidth=CLng(DlgText("Width"))
		Case "Height"
        	If DlgText("Height") <> "" And IsNumeric(DlgText("Height"))=True Then
				Validate(CDbl(DlgText("Height")))
				If BadValue = True Then DlgText "Height",CStr(3.5*Unit)
			ElseIf IsNumeric(DlgText("Height"))=False Then
				MsgBox "You must enter a positive number less then 14in/36cm", vbInformation, "Bad value"
				DlgText "Height",CStr(3.5*Unit)
			End If
			BadValue = False
			DialogFunc = True
			UserHeight=CDbl(DlgText("Height"))
		Case "Top"
        	If DlgText("Top") <> "" And IsNumeric(DlgText("Top"))=True Then
				Validate(CDbl(DlgText("Top")))
				If BadValue = True Then DlgText "Top","0"
			ElseIf IsNumeric(DlgText("Top"))=False Then
				MsgBox "You must enter a positive number less then 14in/36cm", vbInformation, "Bad value"
				DlgText "Top","0"
			End If
			BadValue = False
			DialogFunc = True
			UserTop=CDbl(DlgText("Top"))
		Case "Left"
        	If DlgText("Left") <> "" And IsNumeric(DlgText("Left"))=True Then
				Validate(CDbl(DlgText("Left")))
				If BadValue = True Then DlgText "Left","0"
			ElseIf IsNumeric(DlgText("Left"))=False Then
				MsgBox "You must enter a positive number less then 14in/36cm", vbInformation, "Bad value"
				DlgText "Left","0"
			End If
			BadValue = False
			DialogFunc = True
			UserLeft=CDbl(DlgText("Left"))
		End Select
	Case 4
        DialogFunc = True
	Case 5
        DialogFunc = False 'No idle processing necessary
	End Select
End Function
Private Function GetPages
Index = 0
ReDim GraphPages(0)
For i = 3 To CurrentNotebook.NotebookItems.Count - 1
	If CurrentNotebook.NotebookItems(i).ItemType = 2 Then
		ReDim Preserve GraphPages$(Index)
		GraphPages(Index) = CurrentNotebook.NotebookItems(i).Name
		Index = Index + 1
	End If
Next i
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

Rem See DialogFunc Help topic for more information.
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
Public Function Validate(EnteredValue&)
	BadValue=False
	If EnteredValue/Unit < 0 Then
		MsgBox "You must enter a positive value", vbInformation, "Positive value required"
		BadValue=True
	ElseIf EnteredValue/Unit > 14 Then
		MsgBox "Please enter a value of less than 14 inches or 36cm", vbInformation, "Smaller value required"
		BadValue=True
	End If
End Function