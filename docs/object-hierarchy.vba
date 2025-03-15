Option Explicit

' Flag manipulation functions
Function FlagOn(flag As Long)
    FlagOn = flag Or FLAG_SET_BIT ' Use to set option flag bits on, leaving others unchanged
End Function

Function FlagOff(flag As Long)
    FlagOff = flag Or FLAG_CLEAR_BIT ' Use to set option flag bits off, leaving others unchanged
End Function

' Main execution function
Sub Main()
    Dim SPApp As Object
    Dim SPNotebook As Object
    
    ' Initialize application and create notebook
    Set SPApp = InitializeApplication()
    Set SPNotebook = CreateNotebook(SPApp)
    
    ' Create worksheet and populate with data
    PopulateWorksheet SPNotebook
    
    ' Create and customize graph
    CreateGraph SPNotebook
    
    ' Save the notebook
    SaveNotebook SPNotebook
    
    ' Show completion message
    MsgBox "SigmaPlot demonstration macro is complete.", vbInformation, "Done"
End Sub

' Level 1: Initialize SigmaPlot application
Function InitializeApplication() As Object
    Dim SPApp As Object
    
    ' Launch SigmaPlot if not already running
    Set SPApp = CreateObject("SigmaPlot.Application.1")
    SPApp.Visible = True
    
    InitializeApplication = SPApp
End Function

' Level 2: Create a new notebook
Function CreateNotebook(SPApp As Object) As Object
    Dim SPNotebook As Object
    
    Set SPNotebook = SPApp.Notebooks.Add
    
    ' Note: Trying to set Name may cause issues as it might be read-only
    ' Uncomment if your version allows this
    SPNotebook.Name = "DemoNotebook"
    
    CreateNotebook = SPNotebook
End Function

' Level 3: Create and populate worksheet with data
Sub PopulateWorksheet(SPNotebook As Object)
    Dim SPWorksheet As Object
    Dim SPDataTable As Object
    Dim i As Long
    
    ' Add a worksheet to the notebook
    Set SPWorksheet = SPNotebook.NotebookItems.Add(1)  ' 1 = CT_WORKSHEET
    SPWorksheet.Name = "ExampleData"
    SPWorksheet.Open
    
    ' Access its DataTable for reading/writing cells
    Set SPDataTable = SPWorksheet.DataTable
    
    ' Insert sample data (X and Y)
    For i = 0 To 9
        SPDataTable.Cell(0, i) = i
        SPDataTable.Cell(1, i) = i * i
    Next i
    
    ' Assign column titles
    SPDataTable.Cell(0, -1) = "X Values"
    SPDataTable.Cell(1, -1) = "Y Values"
End Sub

' Level 4: Create and customize a graph
Sub CreateGraph(SPNotebook As Object)
    Dim SPPage As Object
    Dim SPGraph As Object
    Dim PlotColumns(2) As Variant
    Dim XAxis As Object, YAxis As Object
    
    ' Create a new graph page
    Set SPPage = SPNotebook.NotebookItems.Add(2) ' 2 = CT_GRAPHICPAGE
    SPPage.Name = "MyGraphPage"
    
    ' Set up columns for plotting
    PlotColumns(0) = 0  ' X
    PlotColumns(1) = 1  ' Y
    
    ' Create the graph using wizard
    SPPage.CreateWizardGraph "Scatter Plot", "Simple Scatter", "XY Pair", PlotColumns
    
    ' Access the graph object
    Set SPGraph = SPPage.GraphPages(0).Graphs(0)
    SPGraph.Name = "MyScatter"
    
    ' Access and customize the axes
    Set XAxis = SPGraph.Axes(0)   ' 0 = X-axis
    Set YAxis = SPGraph.Axes(1)   ' 1 = Y-axis
    
    ' Set axis titles
    XAxis.Name = "X Axis Title"
    YAxis.Name = "Y Axis Title"
    
    ' Remove the legend for clarity
    SPGraph.SetAttribute SGA_FLAGS, FlagOff(SGA_FLAG_AUTOLEGENDSHOW)
    
    ' Customize X axis range
    XAxis.SetAttribute SAA_OPTIONS, FlagOff(SAA_FLAG_AUTORANGE)
    XAxis.SetAttribute SAA_FROMVAL, -1
    XAxis.SetAttribute SAA_TOVAL, 10
    
    ' Customize Y axis range
    YAxis.SetAttribute SAA_OPTIONS, FlagOff(SAA_FLAG_AUTORANGE)
    YAxis.SetAttribute SAA_FROMVAL, -5
    YAxis.SetAttribute SAA_TOVAL, 90
End Sub

' Level 5: Save the notebook
Sub SaveNotebook(SPNotebook As Object)
    Dim savePath As String
    
    savePath = "C:\Temp\DemoMacroNotebook.jnb"
    SPNotebook.SaveAs savePath
    
    ' Optional close operations (commented out)
    'SPNotebook.Close True
    'SPApp.Quit
End Sub
