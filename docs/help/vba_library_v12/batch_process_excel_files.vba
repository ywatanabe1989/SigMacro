Option Explicit
Public Const ObjectHelp = Path + "\SPW.CHM"
Public HelpID As Variant
Dim SelectedFiles$(),SaveFile$
Dim i, Index, ReportIndex As Integer
Function FlagOn(flag As Long)
  FlagOn = flag Or FLAG_SET_BIT ' Use to set option flag bits on, leaving others unchanged
End Function
Function FlagOff(flag As Long)
  FlagOff = flag Or FLAG_CLEAR_BIT ' Use to set option flag bits off, leaving others unchanged
End Function
Sub Main
'Macro created 01/19/2000 John Kuo
'This macro is an example of automating a batch process.  You can create a
'list of any number of Excel files, then import a specified block of data from
'each file.  The data from each Excel file is imported into a separate 
'worksheet.  You can then either plot and/or curve fit the first two columns
'of imported data, and save the results to a specified notebook.

'Note that only the data for the first sheet is currently imported, and there
'is no way to specify a different sheet.  This import property will be forthcoming

'***************************************
'* Set default save file; edit this to *
'* change the default save path        *
'***************************************
SaveFile = path + "\" + "BatchFile.jnb"

'Create Graph Type list
Dim GraphTypes$()
ReDim GraphTypes(1)
GraphTypes(0)="Simple Scatter Plot"
GraphTypes(1)="Simple Bar Chart"
Dim eqIndex As Integer
eqIndex=0
'Initialize file list
ReDim SelectedFiles$(0)
	SelectedFiles(0)=Empty
Dim Equations$()

Dim FitLibrary$
'Defines the equation source.  Edit to use a different fit library
FitLibrary = "Standard.jfl"
'Open the fit library
Notebooks.Open(UserPath + "\" + FitLibrary, ".jfl")
Dim FitFile As Object
Set FitFile = Notebooks(UserPath + "\" + FitLibrary)
FitFile.Visible=False

'Populate equation list with all equation items in fit library
i=0
Index = 0
For i = 0 To FitFile.NotebookItems.Count - 1
	If FitFile.NotebookItems(i).ItemType = 6 Then
		ReDim Preserve Equations$(Index)
		Equations(Index) = FitFile.NotebookItems(i).Name
		If FitFile.NotebookItems(i).Name ="sigmoidal dose-response" Then
	        eqIndex = Index
	    End If
	    Index = Index + 1
	End If
Next i

i=-1
MacroDialog:
	Begin Dialog UserDialog 470,434,"Batch Process Excel Files",.DialogFunc ' %GRID:10,7,1,1
		PushButton 270,7,90,21,"Add &File...",.AddFile 'Click Add File to add excel files to the list
		PushButton 370,7,90,21,"Delete File",.DeleteFile
		Text 10,21,90,14,"&Excel files:",.Text1
		ListBox 10,35,450,126,SelectedFiles(),.Files
		OKButton 270,406,90,21
		CancelButton 370,406,90,21
		PushButton 10,406,90,21,"Help",.Help
		GroupBox 10,161,450,70,"Import Range",.Range
		Text 20,178,80,21,"&Start column",.Text2
		TextBox 110,175,60,21,.startcol
		Text 200,178,80,14,"&End column",.Text4
		TextBox 290,175,60,21,.endcol
		Text 20,206,90,14,"Start &row",.Text3
		TextBox 110,203,60,21,.startrow
		Text 200,206,60,14,"End ro&w",.Text5
		TextBox 290,203,60,21,.endrow
		GroupBox 10,238,450,105,"Process",.GroupBox1
		CheckBox 20,259,80,14,"&Plot data",.PlotData
		Text 110,259,20,14,"&as:",.Text6
		DropListBox 180,256,260,63,GraphTypes(),.GraphList
		CheckBox 20,287,110,14,"&Curve fit data",.FitData
		Text 136,287,90,14,"&using:",.Text7
		DropListBox 180,284,260,154,Equations(),.FitList
		Text 20,308,420,28,"Note:  only the first two columns of imported data will be plotted and/or fit",.Text8
		Text 10,364,120,14,"Sa&ve notebook to:",.Text9
		TextBox 10,378,450,21,.SavePath
		PushButton 370,350,90,21,"Browse...",.BrowseButton
		CheckBox 120,10,140,14,"Single-step mode",.stepmode
		
	End Dialog
	Dim dlg As UserDialog

'Default settings
'SaveFile
dlg.SavePath = SaveFile
'*********************************************
'* Change this setting to change the default *
'* location of the source data block         *
'*********************************************
dlg.startcol = "1"
dlg.startrow = "1"
dlg.endcol = "2"
dlg.endrow = "256"  
'********************************************
'* You can also change whether the data are *
'* plotted or fitted by default             *
'********************************************
dlg.PlotData = 1
dlg.FitData = 1
'*********************************************
'* Sets the default Graph Type               * 
'* 0=Simple Scatter Plot, 1=Simple Bar Chart *
'*********************************************
dlg.GraphList = 0
'*********************************************
'* Sets the default Fit Equation             * 
'*********************************************
dlg.FitList = eqIndex 'sigmoidal dose-response fix for issue 877
'See the end of the file for a list of all built-in equations by number
'These numbers only apply to the factory default Standard.jfl library

Select Case Dialog(dlg)  
	Case 0 'Handles Cancel button
		GoTo Finish
End Select

'Error if no Excel files picked
If SelectedFiles(0) = Empty Then 
	MsgBox "You have not selected any Excel Files",vbExclamation,"No Files Selected"
	GoTo MacroDialog
End If

Dim CurrentNotebook
Set CurrentNotebook = Notebooks.Add

'Iterate through each selected Excel file
'You can change the extension to import files of different types
Index = 0
ReportIndex = 0
For Index = 0 To UBound(SelectedFiles)
	CurrentNotebook.CurrentDataItem.Open
	CurrentNotebook.CurrentItem.Import(SelectedFiles(Index), 0, 0, CLng(dlg.startcol)-1, CLng(dlg.startrow)-1, CLng(dlg.endcol)-1, CLng(dlg.endrow)-1, ".XLS")
	If dlg.stepmode = 1 Then MsgBox("The data is imported from the Excel Worksheet...",vbInformation,"SigmaPlot")

'Plot the graph
If dlg.PlotData = 1 Then
	Dim SPPage
	Set SPPage = CurrentNotebook.NotebookItems.Add(2)  'Creates graph page
	Dim ColumnsPerPlot()
	ReDim ColumnsPerPlot(2, 1)
	ColumnsPerPlot(0, 0) = 0
	ColumnsPerPlot(1, 0) = 0
'	ColumnsPerPlot(2, 0) = 31999999
	ColumnsPerPlot(2, 0) = CVar(dlg.endrow)
	ColumnsPerPlot(0, 1) = 1
	ColumnsPerPlot(1, 1) = 0
'	ColumnsPerPlot(2, 1) = 31999999
	ColumnsPerPlot(2, 1) = CVar(dlg.endrow)
	Dim PlotColumnCountArray()
	ReDim PlotColumnCountArray(0)
	PlotColumnCountArray(0) = 2
	Select Case dlg.GraphList
		Case 0 'Simple Scatter Plot
			SPPage.CreateWizardGraph("Scatter Plot", "Simple Scatter", "XY Pair", ColumnsPerPlot, PlotColumnCountArray)
		Case 1 'Simple Bar Chart
			SPPage.CreateWizardGraph("Vertical Bar Chart", "Simple Bar", "XY Pair", ColumnsPerPlot, PlotColumnCountArray)
	End Select
	SPPage.GraphPages(0).Graphs(0).Plots(0).SelectObject 'Curve needs to be selected in order to plot curve fit
	SPPage.Open
	If dlg.stepmode = 1 Then MsgBox("The data is plotted...",vbInformation,"SigmaPlot")
End If

'Fit the data; modify the fit options to suit your needs
On Error GoTo FitFailed
If dlg.FitData = 1 Then
	Dim FitEquation$
	FitEquation = Equations(dlg.FitList)
	Dim FitObject As Object
	Set FitObject = FitFile.NotebookItems(FitEquation)
	FitObject.Open
	FitObject.DatasetType = CF_XYPAIR
	FitObject.Variable("x") = "col(1)"
	FitObject.Variable("y") = "col(2)"
	FitObject.Run
	FitObject.OutputReport = True
	FitObject.OutputEquation = False
	FitObject.ResidualsColumn = -1
	FitObject.PredictedColumn = -1
	FitObject.ParametersColumn = -1
	FitObject.OutputGraph = False
	FitObject.OutputAddPlot = True
	FitObject.ExtendFitToAxes = True
	FitObject.AddPlotGraphIndex = 0
	FitObject.XColumn = -1
	FitObject.YColumn = -1
	FitObject.ZColumn = -2
	FitObject.Finish
	If dlg.stepmode = 1 Then MsgBox("The data is curve fitted and a report is generated...",vbInformation,"SigmaPlot")
End If

Wait 1

'Close the document windows
CurrentNotebook.CurrentDataItem.Close(True)
If dlg.FitData = 1 Then CurrentNotebook.NotebookItems("Report " + CStr(ReportIndex + 1)).Close(True)
GoTo Skip
GoTo Finish
FitFailed:
MsgBox("Error(s) have occurred in fitting your data.",vbExclamation,"SigmaPlot")
ReportIndex = ReportIndex - 1
Skip:
If dlg.PlotData = 1 Then CurrentNotebook.NotebookItems("Graph Page " + CStr(Index + 1)).Close(True)
If dlg.stepmode = 1 Then MsgBox("The results windows are closed...",vbInformation,"SigmaPlot")

'Create a new worksheet for the next file
If Index <> UBound(SelectedFiles) Then 
	CurrentNotebook.NotebookItems.Add(1)
	If dlg.stepmode = 1 Then MsgBox("A worksheet is created for the next Excel worksheet...",vbInformation,"SigmaPlot")
End If

ReportIndex = ReportIndex + 1
Next Index
FitFile.Close(False) 'Close standard.jfl

'Save the file
If dlg.stepmode = 1 Then MsgBox("You are prompted to save the file...",vbInformation,"SigmaPlot")
CurrentNotebook.SaveAs(SaveFile)

Finish:
End Sub
Rem See DialogFunc help topic for more information.
Private Function DialogFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgEnable ("DeleteFile",False)
	Case 2 ' Value changing or button pressed
        Select Case DlgItem$
        Case "CancelButton"
        	DlgEnd 1000 'Handles Cancel button from file dialog
		Case "Help"
			HelpID = 70200			' Help ID number for this topic in SPW.CHM
			Help(ObjectHelp,HelpID)
			DialogFunc = True 'do not exit the dialog
		Case "AddFile" 'Adds file to list
			Dim SelectedFile$
			SelectedFile = GetFilePath (,"XLS",,"Select Excel File")'You can change the extension to import files of different types
			If SelectedFile <> "" Then
				i=i+1
				ReDim Preserve SelectedFiles$(i)
				SelectedFiles(i) = SelectedFile
				DlgListBoxArray "Files",SelectedFiles 
			End If
			DialogFunc = True 'do not exit the dialog
		Case "DeleteFile" 'Removes files from list
			SelectedFiles(DlgValue("Files"))=Empty
			If DlgValue("Files") < UBound(SelectedFiles) Then 'Re-indexes array if index removed from middle of array
				For Index = DlgValue("Files") To UBound(SelectedFiles)-1
					If SelectedFiles(Index)=Empty Then
						SelectedFiles(Index)=SelectedFiles(Index+1)
						SelectedFiles(Index+1)=Empty
					End If
				Next Index 
			End If
			If i >= 1 Then i=i-1
			If i <= 0 Then i=0
			
			DlgListBoxArray "Files",SelectedFiles 
			ReDim Preserve SelectedFiles$(i)
			If i = 0 Then i = -1
			DlgEnable("DeleteFile",False)
			DialogFunc = True 'do not exit the dialog
		Case "BrowseButton" 'Set save path
			SaveFile = GetFilePath(,"JNB",,"Select Notebook File",1)
			If SaveFile <> "" Then DlgText("SavePath",SaveFile)
			DialogFunc = True 'do not exit the dialog
       	Case "Files" 'Enables Delete button if a file is selected
			If DlgValue("Files") <> -1 Then DlgEnable("DeleteFile",True)
		End Select
	Case 3
        Select Case DlgItem$
			Case "SavePath"
				SaveFile=DlgText("SavePath")
		End Select
	Case 4
        Select Case DlgItem$
       	Case "Files" 'Enables Delete button if a file is selected
			If DlgValue("Files") <> -1 Then DlgEnable("DeleteFile",True)
		End Select
	Case 5
        DialogFunc = True 
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
'Standard.jfl equation list
'Use the following values to set the default equation number:
'Polynomial
'	0'	Linear
'	1'	Quadratic
'	2'	Cubic
'	3'	Inverse First Order
'	4'	Inverse Second Order
'	5'	Inverse Second Order
'Peak
'	6'	Gaussian, 3 Parameter
'	7'	Gaussian, 4 Parameter
'	8'	Modified Gaussian, 4 Parameter
'	9'	Modified Gaussian, 5 Parameter
'	10'	Lorentzian, 3 Parameter
'	11'	Lorentzian, 4 Parameter
'	12'	Pseudo-Voigt, 4 Parameter
'	13'	Pseudo-Voigt, 5 Parameter
'	14'	Log Normal, 3 Parameter
'	15'	Log Normal, 4 Parameter
'	16'	Weibull, 4 Parameter
'	17'	Weibull, 5 Parameter
'Sigmoidal
'	18'	Sigmoid, 3 Parameter
'	19'	Sigmoid, 4 Parameter
'	20'	Sigmoid, 5 Parameter
'	21'	Logistic, 3 Parameter
'	22'	Logistic, 4 Parameter
'	23'	Weibull, 4 Parameter
'	24'	Weibull, 5 Parameter
'	25'	Gompertz, 3 Parameter
'	26'	Gompertz, 4 Parameter
'	27'	Hill, 3 Parameter
'	28'	Hill, 4 Parameter
'	29'	Chapman, 3 Parameter
'	30'	Chapman, 4 Parameter
'Exponential Decay	
'	31'	Single, 2 Parameter
'	32'	Single, 3 Parameter
'	33'	Double, 4 Parameter
'	34'	Double, 5 Parameter
'	35'	Triple, 6 Parameter
'	36'	Triple, 7 Parameter
'	37'	Modified Single, 3 Parameter
'	38'	Exponential Linear Combination
'Exponential Rise To Maximum	
'	39'	Single, 2 Parameter 
'	40'	Single, 3 Parameter 
'	41'	Double, 4 Parameter 
'	42'	Double, 5 Parameter 
'	43'	Simple Exponent, 2 Parameter
'	44'	Simple Exponent, 3 Parameter 
'Exponential Growth	
'	45'	Single, 1 Parameter
'	46'	Single, 2 Parameter  
'	47'	Single, 3 Parameter  
'	48'	Double, 4 Parameter  
'	49'	Double, 5 Parameter  
'	50'	Modified Single, 1 Parameter
'	51'	Modified Single, 2 Parameter
'	52'	Stirling Model
'	53'	Simple Exponent, 2 Parameter 
'	54'	Simple Exponent, 3 Parameter
'	55'	Modified Simple Exponent, 2 Parameter
'Hyperbola	
'	56'	Single Rectangular, 2 Parameter
'	57'	Single Rectangular i, 3 Parameter
'	58'	Single Rectangular II, 3 Parameter
'	59'	Double Rectangular, 4 Parameter
'	60'	Double Rectangular, 5 Parameter
'	61'	Hyperbolic Decay, 2 Parameter
'	62'	Hyperbolic Decay, 3 Parameter
'	63'	Modified Hyperbola i
'	64'	Modified Hyperbola II
'	65'	Modified Hyperbola III
'Waveform	
'	66'	Sine, 3 Parameter
'	67'	Sine, 4 Parameter
'	68'	Sine Squared, 3 Parameter
'	69'	Sine Squared, 4 Parameter
'	70'	Damped Sine, 4 Parameter
'	71'	Damped Sine, 5 Parameter
'	72'	Modified Sine
'	73'	Modified Sine Squared
'	74'	Modified Damped Sine
'Power	
'	75'	2 Parameter 
'	76'	3 Parameter  
'	77'	Pareto Function
'	78'	Symmetric, 3 Parameter
'	79'	Symmetric, 4 Parameter
'	80'	2 Parameter Modified i
'	81'	2 Parameter Modified II
'	82'	Modified Pareto Function
'Rational	
'	83'	1 Parameter i
'	84'	1 Parameter II
'	85'	2 Parameter i
'	86'	2 Parameter II
'	87'	3 Parameter i
'	88'	3 Parameter II
'	89'	3 Parameter III
'	90'	3 Parameter IV
'	91'	4 Parameter
'	92'	5 Parameter
'	93'	6 Parameter
'	94'	7 Parameter
'	95'	8 Parameter
'	96'	9 Parameter
'	97'	10 Parameter
'	98'	11 Parameter
'Logarithm	
'	99'	2 Parameter i 
'	100'	2 Parameter II 
'	101'	2 Parameter III 
'	102'	3 Parameter 
'	103'	2nd Order
'	104'	3rd Order
'3D	
'	105'	Plane
'	106'	Paraboloid
'	107'	Gaussian
'	108'	Lorentzian
'User-Defined	
'	109'	Untitled
'Standard Curves	
'	110'	Linear Curve
'	111'	Four Parameter Logistic Curve