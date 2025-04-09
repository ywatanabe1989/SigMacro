Option Explicit
Public Const ObjectHelp = Path + "\SPW.CHM"
Dim Separator$
Public HelpID As Variant
Dim usedcs()
Dim dwskt As Object
Dim a, b, i, counter, number		'misc counter variables
Dim numeric_cell As Boolean			'if cell is numeric this is true
Dim testvar							'dialog value of selected column
Dim Tex()							'array of columns picked used in dialog box
Dim UsedColumns$()					'list of all usedcolumns in worksheet
Dim ListedColumns()						'list of all columns (both used and empty) in data area
Dim Lcol()							'array of column numbers for the selected columns
Dim unit$()							'array of unit types
Dim totals							'value of units used to work out %'s
Dim results()						'array of mg/l,mmol/l and %
Dim sL, sD							'international list separator & international decimal symbol
Dim Qnan							'text string for empty cell (-1.#qnan or -1,#qnan)
Dim posinf, neginf					'text string for +ve and -ve values
Dim textlist()
'Dim chosencols$()
Dim change As Boolean
Function FlagOn(flag As Long)
  FlagOn = flag Or FLAG_SET_BIT ' Use to set option flag bits on, leaving others unchanged
End Function
Function FlagOff(flag As Long)
  FlagOff = flag Or FLAG_CLEAR_BIT ' Use to set option flag bits off, leaving others unchanged
End Function
Sub Main
HelpID = 80503			' Help ID number for this topic in SPW.HLP

Dim ion()
Dim gdata()
Dim gspace			'variable used to create gap between the bottom two plots
gspace = 12
testvar = 0	'first item in columns list is highlighted in dialog box as default
ReDim Tex(6)
'ReDim chosencols(0)
Dim datatype$(2)
datatype$(2) = "mmol/1"
datatype$(1) = "mg/l"
datatype$(0) = "percentages"
ReDim unit(6)
unit(0)="Ca 2+"
unit(1)="Mg 2+"
unit(2)="Na +"
unit(3)="K +"
unit(4)="HCO 3-"
unit(5)="SO4 2-"
unit(6)="Cl -"
counter = 0
Dim cnb, dwskn
Dim lastcolumn As Long	'variable to define last used column in worksheet
Dim lastrow As Long		'variable to define last used row in worksheet
sL = ListSeparator  'international list separator
sD = DecimalSymbol  'international decimal symbol
GetsystemValues			'define an empty, +/- infinity cell

'Get current notebook and name of user's section and data worksheet
Dim Cwsk As Object, CurrentSection As Object
Set cnb=ActiveDocument
Dim Cnbn As String
Cnbn = cnb.FullName		'full name of file

continue:
Dim nn As String
nn = TypeName(cnb.CurrentDataItem)
If CStr(nn)="Nothing" Then GoTo nowsk
Set Cwsk = cnb.CurrentDataItem
dwskn = Cwsk.Name		'Name of active Worksheet
'Find last column and row in worksheet
Set dwskt= Cwsk.DataTable
lastcolumn = 0
lastrow = 0
dwskt.GetMaxUsedSize(lastcolumn,lastrow)		'gets worksheet dimensions
If lastcolumn<1 Then GoTo nodata				'if no data in worksheet
Cwsk.InsertionMode = False						'Put worksheet in to overwrite mode

If lastcolumn>9 Then insert_dialog
	'HelpMsgBox HelpID, "You have data beyond column 9.  This macro requires columns ten and above to be empty of all data.",vbCritical,"Data in Required Empty Columns"
	'End
'End If

'Check if this is a results worksheet from a previous Piper run, and
'Sort through columns and create list of columns with values in row 1
Dim tempstring, Index, ListIndex, ColContents, ColTitle
ReDim UsedColumns$(lastcolumn -1)
ReDim ListedColumns(lastcolumn -1)
ReDim Lcol(lastcolumn -1)
ListIndex = 0
For Index = 0 To lastcolumn - 1
tempstring = CStr(dwskt.Cell(Index,-1))
If InStr(tempstring,"Cation") <> 0 Or InStr(tempstring,"Anion") <> 0 Then
		HelpMsgBox HelpID, "The open worksheet is a previously created Piper results worksheet." + vbCrLf +  "Please open a worksheet containing your raw data.",vbExclamation, "Piper Plot Data Present"
		GoTo endd
End If
	ColContents = empty_col(Index, lastrow)
	ColTitle = dwskt.Cell(Index,-1) 'Retrieve column title
	If ColContents = True Then GoTo NextIndex
	If ColContents = False Then   'If the first cell is not empty
		If CStr(ColTitle) = CStr(Qnan) Then	'if column title is empty then
			UsedColumns$(Index) = "Column " + CStr(Index + 1)
			ListedColumns(ListIndex) = CStr(Index)
			ListIndex = ListIndex + 1
		Else
			UsedColumns$(Index) = ColTitle	'If title is present use title
			ListedColumns(ListIndex) = CStr(Index)
			ListIndex = ListIndex + 1
		End If
	End If
	NextIndex:
Next Index
If UBound(UsedColumns$)<6 Then GoTo usederr
ReDim textlist(6)
Tex(0) = UsedColumns(0)
	Begin Dialog UserDialog 590,284,"Piper Plot",.select_data ' %GRID:10,4,1,1
		GroupBox 10,158,360,88,"",.GroupBox8
		Text 20,167,340,72,"Assign the columns for concentrations of each anion and cation using the Assign button.  You can re-assign data, but you cannot assign the same columns to different ions/cations.  Use the Units drop down to select your plot units.  Results are placed into your worksheet beginning in column 10.",.Text1
		PushButton 160,36,90,20,"&Assign >>",.assign
		Text 10,4,140,16,"&Worksheet columns:",.Col_title
		ListBox 10,20,140,140,UsedColumns(),.usedcols
		Text 260,4,110,16,"Assign &to:",.Text2
		ListBox 260,20,110,92,unit(),.units
		Text 260,112,90,16,"&Units:",.unit_title
		DropListBox 260,128,110,72,datatype(),.unitlist
		OKButton 380,256,90,20
		CancelButton 490,256,90,20
		PushButton 10,256,80,20,"Help",.PushButton1
		Text 380,4,180,16,"Assigned columns:",.Text3
		GroupBox 380,48,200,28,"",.GroupBox1
		GroupBox 380,16,200,28,"",.GroupBox2
		GroupBox 380,80,200,28,"",.GroupBox3
		GroupBox 380,112,200,28,"",.GroupBox4
		GroupBox 380,144,200,28,"",.GroupBox5
		GroupBox 380,176,200,28,"",.GroupBox6
		GroupBox 380,208,200,28,"",.GroupBox7
		Text 390,26,50,12,"Ca 2+",.Text4
		Text 390,58,50,16,"Mg 2+",.Text5
		Text 390,90,50,16,"Na +",.Text6
		Text 390,122,50,16,"K +",.Text7
		Text 390,154,60,16,"HCO 3-",.Text8
		Text 390,186,60,16,"SO4 2-",.Text9
		Text 390,218,40,16,"Cl -",.Text10
		Text 470,26,90,12,"",.Text11
		Text 470,58,90,16,"",.Text12
		Text 470,90,90,16,"",.Text13
		Text 470,122,90,16,"",.Text14
		Text 470,154,90,16,"",.Text15
		Text 470,186,90,16,"",.Text16
		Text 470,218,90,16,"",.Text17
	End Dialog
Dim dlg As UserDialog

dlg.unitlist = 1

If Dialog(dlg)= 0 Then GoTo endd	'If user presses CANCEL button
GetLastRow	'redefine last row value for just those columns picked
If b = 1 Then GoTo nodatacols	'if none of the selected columns have any data
lastrow = b - 2
ReDim results(20,lastrow-1)	'results array contains results data
Dim countvar, pie, sins, tans, coss	'sets variables as pi, and sins as sin(60degrees) etc
pie=3.1415926535898
sins=Sin(pie/3)
tans=Tan(pie/3)
coss=Cos(pie/3)
selectunit:
Select Case CStr(dlg.unitlist)
Case 1	'Second unit type selected mg/l
	For i = 0 To lastrow-1
	b = i
	numeric_cell = True		'}
	checkcellvalue			'}Check if a cell contains non-numeric data (True = numeric)
	If numeric_cell = False Then GoTo non_numeric
	Next i					'}
	For i = 0 To lastrow-1		'}
	b=i							'}
	workoutmmol					'}work out mmol/l data
	Next i						'}
	For i = 0 To lastrow-1	'}
	b=i						'}
	workoutpercent			'}convert mmol/l data In To percentages for plotting
	Next i					'}
Case 2	'Third unit type selected mmol/l
	For i = 0 To lastrow-1
	b = i
	numeric_cell = True
	checkcellvalue	'Check if a cell contains non-numeric data (True = numeric)
	If numeric_cell = False Then GoTo non_numeric
	Next i
	For i = 0 To lastrow-1
	b=i
	workoutmgl		'convert mmol/l data back in to mg/l
	workoutpercent	'convert mmol/l data In To percentages for plotting
	Next i
Case Else	'Percentages can be treated differently
	a = lastrow
	checkpercents
	If number = -1 Then GoTo percentage_error
	ReDim gdata(1,lastrow-1)	'X,Y co-ordinates for plot
	ReDim ion(5,lastrow-1)
	For i = 0 To lastrow-1
	ion(0,i) = CDbl(((100-dwskt.Cell(Lcol(0),i))*tans) - (dwskt.Cell(Lcol(1),i)*sins))/tans  'Cat X data
	ion(1,i) = dwskt.Cell(Lcol(1),i) * sins	'Cat Y data
	ion(2,i) = 100 + gspace + dwskt.Cell(Lcol(6),i) + ((dwskt.Cell(Lcol(5),i)*sins)/tans)	'An X data
	ion(3,i) = dwskt.Cell(Lcol(6),i) * sins		'An Y data
	ion(4,i) = ion(1,i) - (tans*ion(0,i))	'Cat_con
	ion(5,i) = tans*((100 + gspace) + dwskt.Cell(Lcol(6),i) + 2*((dwskt.Cell(Lcol(5),i)*sins)/tans)) 'An_con
	gdata(0,i) = (ion(5,i)-ion(4,i))/(2*tans)
	gdata(1,i) = (gdata(0,i)*tans) + ion(4,i)
	Next i
	'Enter column titles
	For i = 0 To 6
	dwskt.Cell(Lcol(i),-1)=unit(i) + " %"
	Next i
	'check for used data columns
	ReDim usedcs(0)
	For Index = 30 To 38	'Check columns 31 - 37
	ColContents = empty_col(Index, lastrow)
	If ColContents = True Then GoTo NextIndex3
	If ColContents = False Then   'If column contains data
		ReDim Preserve usedcs(i)
		usedcs(i) = Index
		i = i + 1
	End If
	NextIndex3:
	Next Index
	If UBound(usedcs) = 0 Then GoTo enterperdata
	'columns in use, set insert columns range to a and b
	a = 31
	b = 42
	enterperdata:
	dwskt.PutData(ion,31,0)			'Enter ion data in to worksheet for two lower plots
	dwskt.PutData(gdata,35,0)		'Enter data in to worksheet for top plot
	GoTo plot_section
End Select

'Work out Cation and Anion Plotting Data for mg/l and mmol/l
ReDim gdata(1,lastrow-1)	'X,Y co-ordinates for plot
ReDim ion(5,lastrow-1)
For i = 0 To lastrow-1
ion(0,i) = CDbl(((100-results(2,i))*tans) - (results(5,i)*sins))/tans  'Cat X data
ion(1,i) = results(5,i) * sins	'Cat Y data
ion(2,i) = 100 + gspace + results(20,i) + ((results(17,i)*sins)/tans)	'An X data
ion(3,i) = results(17,i) * sins		'An Y data
ion(4,i) = ion(1,i) - (tans*ion(0,i))	'Cat_con
ion(5,i) = tans*((100 + gspace) + results(20,i) + 2*((results(17,i)*sins)/tans)) 'An_con
gdata(0,i) = (ion(5,i)-ion(4,i))/(2*tans)	'X plot data
gdata(1,i) = (gdata(0,i)*tans) + ion(4,i)	'Y plot data
Next i

'Else enter results array (mg/mmol/%) and add column titles
i = 0
ReDim usedcs(0)
For Index = 9 To 35
	ColContents = empty_col(Index, lastrow)
	If ColContents = True Then GoTo NextIndex2
	If ColContents = False Then   'If column contains data
		ReDim Preserve usedcs(i)
		usedcs(i) = Index
		i = i + 1
	End If
	NextIndex2:
Next Index
If UBound(usedcs) = 0 Then GoTo enterdata
a = 9
b = 42
'insert_dialog
enterdata:
coltitles						'Enter column titles
dwskt.PutData(results,9,0)		'Enter results array (mg/l,mmol/l,%)
dwskt.PutData(ion,31,0)			'Enter ion data in to worksheet for two lower plots
dwskt.PutData(gdata,35,0)		'Enter data in to worksheet for top plot

plot_section:
'add polar co-ordinates to stop "dataset is empty" error message being displayed
dwskt.Cell(39,0)="0"
dwskt.Cell(40,0)="0"
dwskt.Cell(41,0)="100"
'add axes labels for third piper
For i=0 To 10
dwskt.Cell(37,10+i)=CDbl(i*10)
dwskt.Cell(38,0+i)=100-CDbl(i*10)
Next i
ActiveDocument.NotebookItems.Add(CT_GRAPHICPAGE)
ActiveDocument.CurrentPageItem.ApplyPageTemplate("Piper Plot")

On Error GoTo NameError
Dim iError As Long
iError = 1
Dim GraphPageName As String
NameWorksheet: If iError = 1 Then
	GraphPageName = "Piper Plot 1"
Else
	GraphPageName = "Piper Plot " + CStr(iError)
End If
ActiveDocument.CurrentPageItem.Name = GraphPageName
'Enter column titles
dwskt.NamedRanges.Add("Cation",31,0,1,-1,True)
dwskt.NamedRanges.Add("Data " + CStr (iError),32,0,1,-1,True)
dwskt.NamedRanges.Add("Anion",33,0,1,-1,True)
dwskt.NamedRanges.Add("Data " + CStr (iError),34,0,1,-1,True)
dwskt.NamedRanges.Add("Centre",35,0,1,-1,True)
dwskt.NamedRanges.Add("Cation",36,0,1,-1,True)
dwskt.NamedRanges.Add("Axes",37,0,1,-1,True)
dwskt.NamedRanges.Add("Labels",38,0,1,-1,True)
dwskt.NamedRanges.Add("Scale 1",39,0,1,-1,True)
dwskt.NamedRanges.Add("Scale 2",40,0,1,-1,True)
dwskt.NamedRanges.Add("Scale 3",41,0,1,-1,True)
GoTo endd:

'Error Messages...
nodata:										'There is no data in this worksheet
HelpMsgBox HelpID, "There is no data in this worksheet.",vbExclamation,"Empty Data Worksheet"
GoTo endd

nodatacols:									'There is no data in the columns selected
HelpMsgBox HelpID, "There is no data in the columns selected.",vbExclamation,"Empty Column Selected"
GoTo endd

nowsk:
'Dim Items, iIndex
'iIndex = 0
'While iIndex< cnb.NotebookItems.Count
'	Items = CStr(cnb.NotebookItems(iIndex).ItemType)
'	If Items = 1 Or Items = 8 Then
'		Cnbn = cnb.NotebookItems(iIndex).Name
'		If MsgBox ("There was no datasheet open. Continue with the worksheet..." +CStr(Cnbn)+ "?",vbYesNo,"No Worksheet Open")=vbNo Then
'			GoTo endd
'		End If
'		cnb.NotebookItems(iIndex).Open
'		GoTo continue
'	Else
'		iIndex = iIndex + 1
'	End If
'Wend
HelpMsgBox HelpID, "There is no worksheet",vbExclamation,"No Data Worksheet"
GoTo endd

non_numeric:							'Non-numeric cell encountered
	Begin Dialog UserDialog 340,189,"Non-numeric Data Cell" ' %GRID:10,7,1,1
		Text 10,7,330,42,"The following cell contains a non-numeric value. Please enter a new value and press Continue, or press the Cancel button to quit.",.Text1
		GroupBox 10,56,110,42,"Units",.GroupBox1
		Text 20,74,90,14," " + CStr(unit(Lcol(counter))) + " ",.Text6
		GroupBox 130,56,200,42,"Cell reference:",.GroupBox2
		Text 140,74,170,14,"Column: " + CStr((i)+1)  + "   Row: " + CStr(CDbl(counter)+1) + "",.Text4
		GroupBox 10,112,110,42,"Current value",.GroupBox4
		Text 20,130,90,14," " + CStr(dwskt.Cell(Lcol(counter),b)) + " ",.Text7
		GroupBox 130,112,200,42,"&New value",.GroupBox3
		TextBox 140,126,170,21,.TextBox1
		PushButton 140,161,90,21,"&Continue",.PushButton1
		CancelButton 240,161,90,21
	End Dialog
	Dim dlgnn As UserDialog
	dlgnn.textbox1 = ""
If Dialog(dlgnn)= 0 Then GoTo endd	'If user presses CANCEL button
If CStr(dlgnn.textbox1) = "" Then
	If MsgBox("New value is empty. Enter a new value or Cancel to stop this macro.",vbOkCancel,"Data Entry Error") = vbOK Then
    	GoTo non_numeric
    Else
    	GoTo endd
    End If
End If
dwskt.Cell(counter,i)= CDbl(dlgnn.textbox1)
GoTo selectunit

usederr:							'Worksheet has less than the seven columns required
HelpMsgBox HelpID, "There are not enough columns of data to create a Piper plot. You must have one" + vbCrLf + "column for each data type, with a minimum of 7 columns.",64,"Not Enough Data"
GoTo endd

percentage_error:
HelpMsgBox HelpID, "The values do not add up to 100%. Please amend the values and run the macro again",vbExclamation,"Not Percentage Data"
GoTo endd

NameError:
Select Case Err.Number	'evaluate error number.
	Case 65535	        'duplicate worksheet name
		iError = iError + 1
	Case Else           'handle other situations here
		MsgBox(Err.Description + " (" + CStr(Err.Number) + ")" + " in CreateGraphWorksheet subroutine", 16, "Unknown Error")
End Select
Resume NameWorksheet

endd:
End Sub
Sub GetsystemValues
	Qnan = "-1" & sD & "#QNAN"
	posinf = "-1" & sD & "#INF"
	neginf = "1" &sD & "#INF"
End Sub
Public Function empty_col(column As Variant, column_end As Variant)
'Determines if a column is empty
	Dim WorksheetTable As Object
	Set WorksheetTable = ActiveDocument.CurrentDataItem.DataTable
	Dim i As Long
	Dim empty_cell As Boolean
	For i = 0 To column_end
		If CStr(WorksheetTable.Cell(column, i)) <> Qnan Then
			empty_col = False: Exit For	'column contains data
		Else
			empty_col = True	'current cell is empty, so set column as empty
		End If
	Next i
End Function
Private Function select_data(DlgItem$, Action%, SuppValue&) As Boolean
	Dim stayindlg As Boolean	'var to determine whether to keep dlg open(TRUE)
	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgEnable "ok",False
	Case 2 ' Value changing or button pressed
		Select Case DlgItem$
	        Case "PushButton1"			' Help button
				Help(ObjectHelp,HelpID)
	        	stayindlg = True 'do not exit the dialog
   			Case "units"
				counter = SuppValue
				change = True
			Case "usedcols"
				testvar = DlgValue("usedcols")
				change = True
			Case "Cancel"
				stayindlg = False	'handles cancel button
			Case "assign"	'User selects options in listboxes
				If change = True Then
					change = False
				End If

				DlgText "TEXT" + CStr(DlgValue("units")+11),UsedColumns(ListedColumns(DlgValue("usedcols")))'set text to new value

				If DlgText("Text11")<>"" And DlgText("Text12")<>"" And DlgText("Text13")<>"" And DlgText("Text14")<>"" And DlgText("Text15")<>"" And DlgText("Text16")<>"" And DlgText("Text17")<>"" Then
					DlgEnable "OK",True
				End If

				If counter < 7 Then Lcol(counter)=ListedColumns(testvar)

				stayindlg = True
				counter = counter + 1
				testvar=testvar + 1
				DlgFocus (2)	'set dialog focus to list of data columns
				DlgValue "units",counter
				DlgValue "usedcols",testvar

			Case Else

			If CStr(DlgItem$)="OK" Then
					i = 0 'counting var
					b = 1 'counting var
					start:
					For i = b To 6
					If CStr(DlgText ("TEXT"+CStr(CDbl(b+10)))) = (DlgText ("Text"+CStr(CDbl(i+11)))) Then
						MsgBox "Entry " + CStr(unit$(b-1)) + " is the same as " + CStr(unit(i))
						counter = 6
						stayindlg = True
						GoTo exit_loop
					End If
					Next i
					b = b + 1
					If b<6 Then
						GoTo start
					End If
					If CStr(DlgText ("TEXT16")) = CStr(DlgText ("Text17")) Then
						MsgBox "Entry Six is the same as 7"
						counter = 6
						stayindlg = True
						GoTo exit_loop
					End If	'Check last two values <>
					'end of checking procedure
'					For i = 0 To 6	'loop to collect selected columns
'					chosencols$(i)=CStr(DlgText ("TEXT"+CStr(CDbl(i+11))))
'					Next i
'					stayindlg = False
'					GoTo exit_loop
			End If
			exit_loop:

			End Select
		select_data = stayindlg ' Keep dialog box open(true) else close box(false)
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem select_data = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function
Sub GetLastRow
	b = 0
	For i = 0 To 6
	a = 0
    glr:
    counter = CStr(dwskt.Cell(Lcol(i),a))
    While counter <> Qnan
	    counter = CStr(dwskt.Cell(Lcol(i),a))
    	a = a + 1
    Wend
    	If CStr(dwskt.Cell(Lcol(i),a+1)) <> Qnan Then	'make sure not just one cell is missing
    		'a = a +2	'this remmed line checks the next two cells not just the next one
    		a = a + 1
    		GoTo glr
    	Else
    		a = a + 1
    	End If
    If a> b Then
    	b = a
    End If
    Next i
End Sub
Sub checkcellvalue	'set numeric_cell to false if cell contains non numeric content
	For counter = 0 To 6
	If IsNumeric (dwskt.Cell(Lcol(counter),b)) = True Then
		'cell is numeric but check for +/- inf and missing value
		If CStr(dwskt.Cell(Lcol(counter),b)) = posinf Or CStr(dwskt.Cell(Lcol(counter),b)) = neginf Then
			numeric_cell = False: Exit For
		End If
 		'cell is numeric and not +/- INF -> convert any empty cells or missing values to zero
 		If CStr(dwskt.Cell(Lcol(counter),b)) = Qnan Then
 			dwskt.Cell(Lcol(counter),b) = "0"
 		End If
 		results(0+(3*counter),b)=dwskt.Cell(Lcol(counter),b)
	Else
		numeric_cell = False: Exit For
	End If
	next_loop:
	Next counter
End Sub
 Sub workoutmmol
	Dim mmol, charge
 	For counter = 0 To 6
 	charge = 1
 	Select Case counter
 			Case "0"
 			mmol = 40.078	'Ca
 			charge = 2
 			Case "1"
 			mmol = 24.305	'Mg
 			charge = 2
 			Case "2"
 			mmol = 22.9898	'Na
 			Case "3"
 			mmol = 39.0983	'K
 			Case "4"
 			mmol = 61.0171	'HCO
 			Case "5"
 			mmol = 96.0636		'SO4
 			charge = 2
 			Case "6"
 			mmol = 35.4527	'Cl
 		End Select
 	results(1+(3*counter),b)=(dwskt.Cell(Lcol(counter),b)/mmol)*charge
	Next counter
End Sub
Sub workoutpercent
	For counter = 0 To 6
	If counter <= 3 Then
		totals = CDbl(results(1,b))+CDbl(results(4,b))+CDbl(results(7,b))+CDbl(results(10,b))
		If totals = 0 Then
			results(2+(3*counter),b) = 0
			GoTo value_is_zero
		End If
	Else
		totals = CDbl(results(13,b))+CDbl(results(16,b))+CDbl(results(19,b))
		If totals = 0 Then
			results(2+(3*counter),b) = 0
			GoTo value_is_zero
		End If
	End If
	results(2+(3*counter),b) = (results(1+(3*counter),b)/totals)*100
	value_is_zero:
	Next counter
End Sub
Sub workoutmgl
	Dim mmol, charge
 	For counter = 0 To 6
 	charge = 1
 	Select Case counter
 			Case "0"
 			mmol = 40.078	'Ca
 			charge = 2
 			Case "1"
 			mmol = 24.305	'Mg
 			charge = 2
 			Case "2"
 			mmol = 22.9898	'Na
 			Case "3"
 			mmol = 39.0983	'K
 			Case "4"
 			mmol = 61.0171	'HCO
 			Case "5"
 			mmol = 96.0636		'SO4
 			charge = 2
 			Case "6"
 			mmol = 35.4527	'Cl
 		End Select
 	results(1+(3*counter),i) = dwskt.Cell(Lcol(counter),b)
 	results(0+(3*counter),b)=(dwskt.Cell(Lcol(counter),b)*mmol)/charge
	Next counter
End Sub
Sub coltitles
	For i = 0 To 18 Step 3
	dwskt.NamedRanges.Add(CStr(unit(i/3)) + " mg/l",9+i,0,1,-1, True)
	dwskt.NamedRanges.Add(CStr(unit(i/3)) + " mmol/l",i+10,0,1,-1, True)
	dwskt.NamedRanges.Add(CStr(unit(i/3)) + " %",i+11,0,1,-1, True)
	Next i
End Sub
Sub insert_dialog
	Begin Dialog UserDialog 370,210,"Data in Required Columns" ' %GRID:10,7,1,1
		Text 10,7,340,42,"The macro needs to enter the results data in to the worksheet to create the Piper plot. However, the required columns contain data. ",.Text2
		Text 10,49,230,14,"To continue please select an option.",.Text1
		PushButton 10,70,100,21,"&Overwrite",.PushButton1
		PushButton 10,126,100,21,"&Insert",.PushButton2
		Text 120,70,240,56,"Overwrite the data already in the worksheet. (Note: You will lose this data permanently. Please ensure you do not need it before using this option)",.Text3
		Text 120,126,240,28,"Move existing data to new columns after entering these macros results.",.Text4
		CancelButton 260,175,90,21
	End Dialog
	Dim dlgdata As UserDialog
		Select Case Dialog(dlgdata)
		Case 0 'Handles Cancel button
			End
		Case 2 'insert number of columns so existing data is not lost
			ActiveDocument.CurrentDataItem.InsertCells(9, 0, 43, 32000000, InsertRight)
	End Select
End Sub
Sub checkpercents
	For counter = 0 To a-1
	number = 0
	For i = 0 To 3
	number = number + dwskt.Cell(Lcol(i),counter)
	Next i
	If number <> 100 Then GoTo percent_error
	number = 0
	For i = 4 To 6
	number = number + dwskt.Cell(Lcol(i),counter)
	Next i
	If number <> 100 Then GoTo percent_error
	Next counter
	GoTo endofsub
	percent_error:
	number = -1		'indicates that there is an error
	endofsub:
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