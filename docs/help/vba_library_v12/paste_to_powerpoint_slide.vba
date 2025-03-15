Option Explicit
Global PPPath As String
Public Const msoFalse=0
Public Const msoTrue=-1
Public Const ObjectHelp = Path + "\SPW.CHM"
Public HelpID As Variant
Dim PasteFormat As Integer '0 = Embedded, 1 = EMF
Function FlagOn(flag As Long)
  FlagOn = flag Or FLAG_SET_BIT ' Use to set option flag bits on, leaving others unchanged
End Function
Function FlagOff(flag As Long)
  FlagOff = flag Or FLAG_CLEAR_BIT ' Use to set option flag bits off, leaving others unchanged
End Function
Sub Main
'Original PP Macro By Keith Kroeger
'Modified 01/02/01 John Kuo
	Dim ErrorCheck As Integer

	On Error GoTo ErrorMsg

    Dim SourceDoc As Object
	Set SourceDoc = ActiveDocument

	ErrorCheck = 1 'Display no worksheet message
	Dim WorksheetTable As Object
	Set WorksheetTable = ActiveDocument.CurrentDataItem.DataTable

	ErrorCheck = 0
	Dim SPGraph, CurveCount
	Set SPGraph = SourceDoc.CurrentPageItem

	Dim pagename,currentname As String
	currentname=SPGraph.Name

	SourceDoc.NotebookItems(currentname).IsCurrentBrowserEntry = True
	Clipboard "empty"
	SourceDoc.CurrentPageItem.Copy

	Dim TargetDoc As Object

	Dim IsEmbeddedDoc As Boolean
	IsEmbeddedDoc = SourceDoc.IsEmbeddedDoc
	If IsEmbeddedDoc = False Then
		Set TargetDoc = ActiveDocument
	Else
		Set TargetDoc = Application.Notebooks.Add( )
	End If
	TargetDoc.NotebookItems.Add(CT_GRAPHICPAGE)
	pagename=TargetDoc.CurrentBrowserItem.Name
	TargetDoc.CurrentBrowserItem.Name=pagename + " (PowerPoint format)"
	TargetDoc.CurrentItem.Paste
'if no graph selected, end macro
	If TargetDoc.CurrentPageItem.GraphPages(0).Graphs.Count<1 Then GoTo NoGraph

	Dim Colors$()
	ReDim Colors(7)
	Colors(0) = "White"
	Colors(1) = "Red"
	Colors(2) = "Orange"
	Colors(3) = "Yellow"
	Colors(4) = "Green"
	Colors(5) = "Blue"
	Colors(6) = "Indigo"
	Colors(7) = "Violet"

	Begin Dialog UserDialog 479,273,"Insert Graph into PowerPoint",.DialogFunc ' %GRID:10,7,1,0
		OKButton 380,7,90,21
		CancelButton 380,35,90,21
		PushButton 380,70,90,21,"Help",.PushButton1
		CheckBox 19,10,244,16,"&Transparent graph background",.CheckBox1
		CheckBox 19,30,100,16,"&Bold text",.CheckBox2
		CheckBox 19,50,162,16,"Change text &color to",.CheckBox3
		DropListBox 186,47,135,124,Colors(),.DropListBox1
		CheckBox 19,80,140,14,"T&hicken lines to",.CheckBox4
		TextBox 186,77,71,18,.TextBox1
		CheckBox 19,108,160,14,"Change &line color to",.CheckBox5
		DropListBox 186,104,135,124,Colors(),.DropListBox2
		GroupBox 10,210,454,55,"",.GroupBox1
		Text 19,220,436,40,"To use this macro, open your PowerPoint presentation, view the slide where you want to place your figure, then open and select the graph you want to place on the slide.",.Text1
		GroupBox 10,135,204,72,"PowerPoint Graph Format",.GroupBox2
		OptionGroup .PasteFormatGroup
			OptionButton 19,153,175,16,"Embedded",.OptionButton1
			OptionButton 19,181,175,14,"Enhanced Metafile (EMF)",.OptionButton2
	End Dialog

	Dim dlg As UserDialog
	dlg.CheckBox1=True 'transparency applied by default
	dlg.TextBox1 = CStr(.02)
	dlg.PasteFormatGroup = 0
	If Dialog(dlg)=0 Then
		TargetDoc.NotebookItems.Delete(pagename+" (PowerPoint format)")
		GoTo Finish 'if cancel button pressed, end macro
	End If

'if transparent graph background selected, set the plane color to none
	Dim planecolor, LineThickness, linecolor, textbold, textcolor As Long
	If dlg.CheckBox1=1 Then planecolor=&Hff000000
	TargetDoc.NotebookItems(pagename+" (PowerPoint format)").IsCurrentBrowserEntry = True

'page and graph backgrounds
	TargetDoc.CurrentPageItem.GraphPages(0).Color = -16777216 'sets page background to none (transparent)

Dim graphnum
For graphnum=0 To TargetDoc.CurrentPageItem.GraphPages(0).Graphs.Count-1
	TargetDoc.CurrentPageItem.GraphPages(0).Graphs(graphnum).SelectObject
	If dlg.CheckBox1=1 Then 'if user selects transparent background
		TargetDoc.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETGRAPHATTR, SGA_PLANECOLORXYBACK,&Hff000000 ) 'sets graph background color to transparent
	End If
	'text bolded
	If dlg.CheckBox2=1 Then 'if user selects bold text
		TargetDoc.CurrentPageItem.SelectAll
		TargetDoc.CurrentPageItem.SetSelectedObjectsAttribute(STA_BOLD, 700)
	End If
'text color
	Select Case dlg.DropListBox1
		Case 0 'White
			textcolor=&H00FFFFFF& 'white
		Case 1 'Red
			textcolor=&H000000FF& 'red
		Case 2 'Orange
			textcolor=&H004080ff& 'orange
		Case 3 'Yellow
			textcolor=&H0000ffff& 'yellow
		Case 4 'Green
			textcolor=&H0000FF00& 'green
		Case 5 'Blue
			textcolor=&H00FF0000& 'blue
		Case 6 'Indigo
			textcolor=&H00800000& 'dark blue
		Case 7 'Violet
			textcolor=&H00600060& 'violet
	End Select
	If dlg.CheckBox3=1 Then 'if user selects colored text
		'TargetDoc.CurrentPageItem.SetSelectedObjectsAttribute(STA_SELECT, -65536)
		TargetDoc.CurrentPageItem.SelectAll
		TargetDoc.CurrentPageItem.SetSelectedObjectsAttribute(STA_COLOR, textcolor)
	End If
'line thickness
	If dlg.TextBox1 <> "" Then 'Handles empty textbox
		LineThickness = CDbl(dlg.TextBox1)*1000
		If dlg.CheckBox4=1 Then TargetDoc.CurrentPageItem.SetSelectedObjectsAttribute(SEA_THICKNESS, LineThickness)
	End If
'line color
	Select Case dlg.DropListBox2
		Case 0 'White
			linecolor=&H00FFFFFF& 'white
		Case 1 'Red
			linecolor=&H000000FF& 'red
		Case 2 'Orange
			linecolor=&H004080ff& 'orange
		Case 3 'Yellow
			linecolor=&H0000ffff& 'yellow
		Case 4 'Green
			linecolor=&H0000FF00& 'green
		Case 5 'Blue
			linecolor=&H00FF0000& 'blue
		Case 6 'Indigo
			linecolor=&H00800000& 'dark blue
		Case 7 'Violet
			linecolor=&H00600060& 'violet
	End Select
	If dlg.CheckBox5=1 Then
		TargetDoc.CurrentPageItem.SelectAll
		TargetDoc.CurrentPageItem.SetSelectedObjectsAttribute(SOA_COLOR, linecolor)
	End If
'legend background and border
	Dim legendOn As Long
	Dim param As Variant 'Added by Jake in the case that a legend is not included with graph
	legendOn = TargetDoc.CurrentPageItem.GraphPages(0).Graphs(graphnum).GetAttribute(SGA_AUTOLEGENDSHOW,param)
    If legendOn Then
		TargetDoc.CurrentPageItem.GraphPages(0).Graphs(graphnum).AutoLegend.SetObjectCurrent
		TargetDoc.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETOBJECTATTR, SDA_COLOR, &Hff000000)
		If dlg.CheckBox5=1 Then
			TargetDoc.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETOBJECTATTR, SDA_EDGECOLOR, linecolor)
		End If
	End If
Next graphnum

'select and copy the reformatted graph
	TargetDoc.NotebookItems(pagename+" (PowerPoint format)").SelectAll
	TargetDoc.NotebookItems(pagename+" (PowerPoint format)").Copy
	SourceDoc.NotebookItems(currentname).Open

	If IsEmbeddedDoc Then
		TargetDoc.Close(False)
	End If

'start PowerPoint
	Dim PPApp As Object
	On Error GoTo NoPowerPoint
	Set PPApp=CreateObject("PowerPoint.Application")
	PPApp.Visible=True

'adds SigmaPlot graph to PowerPoint presentation
	On Error GoTo NoSlide
'	Dim PasteSpecialEMF As Boolean
'	PasteSpecialEMF = False
	If PasteFormat = 0 Then  'Embedded graph
		'	PPApp.ActiveWindow.ViewType = 1 'ppViewSlide=1
			'PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count).Shapes.Paste
		'	PPApp.ActivePresentation.Slides(PPApp.ActiveWindow.Selection.SlideRange.Name).Shapes.Paste
			PPApp.ActiveWindow.View.Paste
			PPApp.Activate
		'	PPApp.ActiveWindow.WindowState = 3 'ppWindowMaximized
	Else  'EMF
		Dim DataType As Integer
		DataType = 2 'DataType:=ppPasteEnhancedMetafile = 2

		'adds an enhanced metafile to a PowerPoint presentation
		PPApp.ActiveWindow.View.PasteSpecial DataType
		PPApp.Activate
	End If

GoTo Finish

ErrorMsg:
	If ErrorCheck = 0 Then
		HelpMsgBox 60208, "You must have a graph open and in focus",vbExclamation,"SigmaPlot"
	ElseIf ErrorCheck = 1 Then
	   	HelpMsgBox 60208, "You must have a worksheet open to run this macro",vbExclamation,"SigmaPlot"
	End If
	GoTo Finish

NoSlide:
	HelpMsgBox 60208, "You must have a PowerPoint presentation open and a slide selected",vbExclamation,"SigmaPlot"
	GoTo Finish

NoPowerPoint:
	HelpMsgBox 60208, "The PowerPoint application cannot be started",vbExclamation,"SigmaPlot"
	GoTo Finish

NoGraph:
	TargetDoc.NotebookItems.Delete(pagename+" (PowerPoint format)")
	HelpMsgBox 60208, "You have not selected a graph to reformat. Please select a graph before running this macro.",vbInformation,"SigmaPlot"
	GoTo Finish

Finish:

End Sub
Function DialogFunc%(DlgItem$, Action%, SuppValue%)

	Select Case Action%
    	Case 1 ' Dialog box initialization
    	Case 2 ' Value changing or button pressed
        	Select Case DlgItem$
        		Case "PushButton1"				' Help button
					HelpID = 60208			' Help ID number for this topic in SPW.CHM
					Help(ObjectHelp,HelpID)
        			DialogFunc% = True 'do not exit the dialog
      			Case "DropListBox1"
      				DlgValue "CheckBox3",1
        			DialogFunc% = True 'do not exit the dialog
      			Case "DropListBox2"
	       			DlgValue "CheckBox5",1
        			DialogFunc% = True 'do not exit the dialog
				Case "OK"
        			If DlgValue("PasteFormatGroup") = 0 Then
        				PasteFormat = 0
        			Else
        				PasteFormat = 1
        			End If
    		End Select
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