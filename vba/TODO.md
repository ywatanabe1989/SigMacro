<!-- ---
!-- Timestamp: 2025-03-10 03:45:18
!-- Author: ywatanabe
!-- File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/src/pysigmacro/vba/TODO.md
!-- --- -->

Basic
- Object Browser
  - The Object Browser displays all SigmaPlot object classes. The methods and properties associated with each SigmaPlot macro object class are listed.
  - By clicking F1, you can access extensive Help that includes example code for the individual properties and methods. The Paste feature lets you insert generic code based on your selection into a macro.
  - For full details on using the Object Browser, press F1 from anywhere in the Macro window.

Macros to implement

- Create
  - [ ] core_create_or_open_jnb_file
    - [ ] args: path_jnb
  - [ ] core_create_notebook
    - [ ] args: sigmaplot_object, notebook_name
  - [ ] core_create_section
    - [ ] args: notebook_object, section_name
  - [ ] core_create_worksheet
    - [ ] args: notebook_object, notebook_name
  - [ ] core_create_graphpage
    - [ ] args: notebook_object, graphpage_name
  - [ ] core_create_macro
    - [ ] args: notebook_object, macro_name

- Get
  - [ ] core_get_notebooks
  - [ ] core_get_sections
  - [ ] core_get_worksheets
  - [ ] core_get_graphpages
  - [ ] core_get_macros

- Activate/Select
  - [ ] core_activate_notebook
  - [ ] core_activate_section
  - [ ] core_activate_worksheet
  - [ ] core_activate_graphpage
  - [ ] core_activate_macros

- Save
  - [ ] core_save_notebook
  - [ ] core_save_section
  - [ ] core_save_worksheet
  - [ ] core_save_graphpage
  - [ ] core_save_macros

- Graphing
  - [ ] Plotters
    - [ ] graph-create_scatter
    - [ ] graph-create_line
    - [ ] graph-create_bar
    - [ ] graph-create_box
    - [ ] graph-(create_area)
    - [ ] graph-(create_contour)
    - [ ] graph-(create_polar)
    - [ ] graph-(create_rader)

  - [ ] Title and Labels
    - [ ] graph-add_title
    - [ ] graph-add_xlabel
    - [ ] graph-add_ylabel
    - [ ] graph-add_legend

  - [ ] Spines
    - [ ] graph-remove_top_spine
    - [ ] graph-remove_right_spine

  - [ ] Ticks
    - [ ] graph-set_x_ticks
    - [ ] graph-set_x_tick_length
    - [ ] graph-set_x_tick_width
    - [ ] graph-set_y_ticks
    - [ ] graph-set_y_tick_length
    - [ ] graph-set_y_tick_width

  - [ ] Scale
    - [ ] graph-set_x_scale
    - [ ] graph-set_y_scale

  - [ ] Range
    - [ ] graph-set_x_range
    - [ ] graph-set_y_range

  - [ ] Scatter(s)
    - [ ] graph-add_scatter
    - [ ] graph-set_symbol_size
    - [ ] graph-set_symbol_color
- 
  - [ ] Line(s)
    - [ ] graph-add_line
    - [ ] graph-set_linewidth
    - [ ] graph-set_linelength
    - [ ] graph-set_linecolor

  - [ ] Bar(s)
    - [ ] graph-add_bar
    - [ ] graph-set_barwidth
    - [ ] graph-set_barcolor

  - [ ] Box(es)
    - [ ] graph-add_box
    - [ ] graph-set_boxwidth
    - [ ] graph-set_boxcolor

  - [ ] graph_export_graph

- Data
  - [ ] data_import_csv
  - [ ] data_name_columns
  - [ ] data_select_columns


``` vba
Option Explicit

' =====================================================
' Create Macros
' =====================================================

Function core_create_or_open_jnb_file(path_jnb As String) As Object
    Dim SigmaPlot As Object
    Dim Notebooks As Object
    Dim Notebook As Object
    
    On Error Resume Next
    Set SigmaPlot = GetObject(, "SigmaPlot.Application")
    If SigmaPlot Is Nothing Then
        Set SigmaPlot = CreateObject("SigmaPlot.Application")
        If SigmaPlot Is Nothing Then
            MsgBox "Failed to launch SigmaPlot.", vbCritical
            Set core_create_or_open_jnb_file = Nothing
            Exit Function
        End If
    End If
    On Error GoTo 0
    
    SigmaPlot.Visible = True
    Set Notebooks = SigmaPlot.Notebooks
    
    If Dir(path_jnb) = "" Then
        Set Notebook = Notebooks.Add
        Notebook.SaveAs path_jnb
    Else
        Set Notebook = Notebooks.Open(path_jnb)
    End If
    
    Set core_create_or_open_jnb_file = Notebook
End Function

Function core_create_notebook(sigmaplot_object As Object, notebook_name As String) As Object
    Dim Notebook As Object
    Set Notebook = sigmaplot_object.Notebooks.Add
    ' Optionally set the notebook name if supported
    ' Notebook.Name = notebook_name
    Set core_create_notebook = Notebook
End Function

Function core_create_section(notebook_object As Object, section_name As String) As Object
    Dim Section As Object
    Set Section = notebook_object.Sections.Add
    ' Optionally set the section name if supported
    ' Section.Name = section_name
    Set core_create_section = Section
End Function

Function core_create_worksheet(notebook_object As Object, worksheet_name As String) As Object
    Dim Worksheet As Object
    Set Worksheet = notebook_object.Worksheets.Add
    ' Optionally set the worksheet name if supported
    ' Worksheet.Name = worksheet_name
    Set core_create_worksheet = Worksheet
End Function

Function core_create_graphpage(notebook_object As Object, graphpage_name As String) As Object
    Dim GraphPage As Object
    Set GraphPage = notebook_object.GraphPages.Add
    ' Optionally set the graph page name if supported
    ' GraphPage.Name = graphpage_name
    Set core_create_graphpage = GraphPage
End Function

Function core_create_macro(notebook_object As Object, macro_name As String) As Object
    Dim Macro As Object
    Set Macro = notebook_object.Macros.Add
    ' Optionally set the macro name if supported
    ' Macro.Name = macro_name
    Set core_create_macro = Macro
End Function

' =====================================================
' Get Macros
' =====================================================

Function core_get_notebooks() As Object
    Dim SigmaPlot As Object
    Dim Notebooks As Object
    
    On Error Resume Next
    Set SigmaPlot = GetObject(, "SigmaPlot.Application")
    If SigmaPlot Is Nothing Then
        Set SigmaPlot = CreateObject("SigmaPlot.Application")
    End If
    On Error GoTo 0
    
    Set Notebooks = SigmaPlot.Notebooks
    Set core_get_notebooks = Notebooks
End Function

Function core_get_sections(notebook_object As Object) As Object
    Set core_get_sections = notebook_object.Sections
End Function

Function core_get_worksheets(notebook_object As Object) As Object
    Set core_get_worksheets = notebook_object.Worksheets
End Function

Function core_get_graphpages(notebook_object As Object) As Object
    Set core_get_graphpages = notebook_object.GraphPages
End Function

Function core_get_macros(notebook_object As Object) As Object
    Set core_get_macros = notebook_object.Macros
End Function

' =====================================================
' Activate/Select Macros
' =====================================================

Sub core_activate_notebook(notebook_object As Object)
    notebook_object.Activate
End Sub

Sub core_activate_section(section_object As Object)
    section_object.Activate
End Sub

Sub core_activate_worksheet(worksheet_object As Object)
    worksheet_object.Activate
End Sub

Sub core_activate_graphpage(graphpage_object As Object)
    graphpage_object.Activate
End Sub

Sub core_activate_macros(macro_object As Object)
    macro_object.Activate
End Sub

' =====================================================
' Save Macros
' =====================================================

Sub core_save_notebook(notebook_object As Object)
    notebook_object.Save
End Sub

Sub core_save_section(section_object As Object)
    section_object.Save
End Sub

Sub core_save_worksheet(worksheet_object As Object)
    worksheet_object.Save
End Sub

Sub core_save_graphpage(graphpage_object As Object)
    graphpage_object.Save
End Sub

Sub core_save_macros(macro_object As Object)
    macro_object.Save
End Sub

' =====================================================
' Graphing Macros - Plotters
' =====================================================

Function graph_create_scatter(graphpage_object As Object, dataRange As Object) As Object
    Dim Plot As Object
    Set Plot = graphpage_object.Plotters.AddScatter(dataRange)
    Set graph_create_scatter = Plot
End Function

Function graph_create_line(graphpage_object As Object, dataRange As Object) As Object
    Dim Plot As Object
    Set Plot = graphpage_object.Plotters.AddLine(dataRange)
    Set graph_create_line = Plot
End Function

Function graph_create_bar(graphpage_object As Object, dataRange As Object) As Object
    Dim Plot As Object
    Set Plot = graphpage_object.Plotters.AddBar(dataRange)
    Set graph_create_bar = Plot
End Function

Function graph_create_box(graphpage_object As Object, dataRange As Object) As Object
    Dim Plot As Object
    Set Plot = graphpage_object.Plotters.AddBox(dataRange)
    Set graph_create_box = Plot
End Function

Function graph_create_area(graphpage_object As Object, dataRange As Object) As Object
    MsgBox "graph_create_area not implemented", vbExclamation
    Set graph_create_area = Nothing
End Function

Function graph_create_contour(graphpage_object As Object, dataRange As Object) As Object
    MsgBox "graph_create_contour not implemented", vbExclamation
    Set graph_create_contour = Nothing
End Function

Function graph_create_polar(graphpage_object As Object, dataRange As Object) As Object
    MsgBox "graph_create_polar not implemented", vbExclamation
    Set graph_create_polar = Nothing
End Function

Function graph_create_rader(graphpage_object As Object, dataRange As Object) As Object
    MsgBox "graph_create_rader not implemented", vbExclamation
    Set graph_create_rader = Nothing
End Function

' =====================================================
' Graphing Macros - Title and Labels
' =====================================================

Sub graph_add_title(graphpage_object As Object, titleText As String)
    graphpage_object.Title.Text = titleText
End Sub

Sub graph_add_xlabel(graphpage_object As Object, xlabelText As String)
    graphpage_object.XAxis.Label.Text = xlabelText
End Sub

Sub graph_add_ylabel(graphpage_object As Object, ylabelText As String)
    graphpage_object.YAxis.Label.Text = ylabelText
End Sub

Sub graph_add_legend(graphpage_object As Object)
    graphpage_object.AddLegend
End Sub

' =====================================================
' Graphing Macros - Spines
' =====================================================

Sub graph_remove_top_spine(graphpage_object As Object)
    graphpage_object.YAxis.RemoveTopSpine
End Sub

Sub graph_remove_right_spine(graphpage_object As Object)
    graphpage_object.XAxis.RemoveRightSpine
End Sub

' =====================================================
' Graphing Macros - Ticks
' =====================================================

Sub graph_set_x_ticks(graphpage_object As Object, tickValues As Variant)
    graphpage_object.XAxis.SetTicks tickValues
End Sub

Sub graph_set_x_tick_length(graphpage_object As Object, lengthValue As Double)
    graphpage_object.XAxis.TickLength = lengthValue
End Sub

Sub graph_set_x_tick_width(graphpage_object As Object, widthValue As Double)
    graphpage_object.XAxis.TickWidth = widthValue
End Sub

Sub graph_set_y_ticks(graphpage_object As Object, tickValues As Variant)
    graphpage_object.YAxis.SetTicks tickValues
End Sub

Sub graph_set_y_tick_length(graphpage_object As Object, lengthValue As Double)
    graphpage_object.YAxis.TickLength = lengthValue
End Sub

Sub graph_set_y_tick_width(graphpage_object As Object, widthValue As Double)
    graphpage_object.YAxis.TickWidth = widthValue
End Sub

' =====================================================
' Graphing Macros - Scale
' =====================================================

Sub graph_set_x_scale(graphpage_object As Object, scaleType As String)
    graphpage_object.XAxis.ScaleType = scaleType
End Sub

Sub graph_set_y_scale(graphpage_object As Object, scaleType As String)
    graphpage_object.YAxis.ScaleType = scaleType
End Sub

' =====================================================
' Graphing Macros - Range
' =====================================================

Sub graph_set_x_range(graphpage_object As Object, minVal As Double, maxVal As Double)
    graphpage_object.XAxis.Minimum = minVal
    graphpage_object.XAxis.Maximum = maxVal
End Sub

Sub graph_set_y_range(graphpage_object As Object, minVal As Double, maxVal As Double)
    graphpage_object.YAxis.Minimum = minVal
    graphpage_object.YAxis.Maximum = maxVal
End Sub

' =====================================================
' Graphing Macros - Scatter(s)
' =====================================================

Function graph_add_scatter(graphpage_object As Object, dataRange As Object) As Object
    Dim Scatter As Object
    Set Scatter = graphpage_object.Plotters.AddScatter(dataRange)
    Set graph_add_scatter = Scatter
End Function

Sub graph_set_symbol_size(scatterObject As Object, sizeValue As Double)
    scatterObject.SymbolSize = sizeValue
End Sub

Sub graph_set_symbol_color(scatterObject As Object, colorValue As Long)
    scatterObject.SymbolColor = colorValue
End Sub

' =====================================================
' Graphing Macros - Line(s)
' =====================================================

Function graph_add_line(graphpage_object As Object, dataRange As Object) As Object
    Dim LinePlot As Object
    Set LinePlot = graphpage_object.Plotters.AddLine(dataRange)
    Set graph_add_line = LinePlot
End Function

Sub graph_set_linewidth(lineObject As Object, widthValue As Double)
    lineObject.LineWidth = widthValue
End Sub

Sub graph_set_linelength(lineObject As Object, lengthValue As Double)
    lineObject.LineLength = lengthValue
End Sub

Sub graph_set_linecolor(lineObject As Object, colorValue As Long)
    lineObject.LineColor = colorValue
End Sub

' =====================================================
' Graphing Macros - Bar(s)
' =====================================================

Function graph_add_bar(graphpage_object As Object, dataRange As Object) As Object
    Dim BarPlot As Object
    Set BarPlot = graphpage_object.Plotters.AddBar(dataRange)
    Set graph_add_bar = BarPlot
End Function

Sub graph_set_barwidth(barObject As Object, widthValue As Double)
    barObject.BarWidth = widthValue
End Sub

Sub graph_set_barcolor(barObject As Object, colorValue As Long)
    barObject.BarColor = colorValue
End Sub

' =====================================================
' Graphing Macros - Box(es)
' =====================================================

Function graph_add_box(graphpage_object As Object, dataRange As Object) As Object
    Dim BoxPlot As Object
    Set BoxPlot = graphpage_object.Plotters.AddBox(dataRange)
    Set graph_add_box = BoxPlot
End Function

Sub graph_set_boxwidth(boxObject As Object, widthValue As Double)
    boxObject.BoxWidth = widthValue
End Sub

Sub graph_set_boxcolor(boxObject As Object, colorValue As Long)
    boxObject.BoxColor = colorValue
End Sub

Sub graph_export_graph(graphpage_object As Object, exportPath As String)
    graphpage_object.Export exportPath
End Sub

' =====================================================
' Data Macros
' =====================================================

Function data_import_csv(filePath As String) As Object
    Dim DataObj As Object
    MsgBox "data_import_csv not implemented", vbExclamation
    Set DataObj = Nothing
    Set data_import_csv = DataObj
End Function

Sub data_name_columns(dataObject As Object, names As Variant)
    MsgBox "data_name_columns not implemented", vbExclamation
End Sub

Sub data_select_columns(dataObject As Object, columnIndices As Variant)
    MsgBox "data_select_columns not implemented", vbExclamation
End Sub
```

<!-- EOF -->