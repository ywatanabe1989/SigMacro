<!-- ---
!-- Timestamp: 2025-03-12 07:14:24
!-- Author: ywatanabe
!-- File: /home/ywatanabe/proj/SigMacro/docs/object-hierarchy.md
!-- --- -->

# Hierarchy
``` plaintext
Application (SigmaPlot.Application)
├── Notebooks (Collection)
│    ├── Notebook (Object = a .JNB file)
│    │    ├── NotebookItems (Collection)
│    │    │    ├── WorksheetItem (CT_WORKSHEET = 1)
│    │    │    ├── GraphPageItem (CT_GRAPHICPAGE = 2)
│    │    │    ├── FolderItem (CT_FOLDER = 3)
│    │    │    ├── StatTestItem (CT_STATTEST = 4)
│    │    │    ├── ReportItem (CT_REPORT = 6)
│    │    │    ├── FitItem (CT_FIT = 7)
│    │    │    ├── Notebook (CT_NOTEBOOK = 8)
│    │    │    └── ExcelWorksheet (CT_EXCELWORKSHEET = 9)
│    │    ├── Properties (e.g., Name, Saved, FullName, Path, etc.)
│    │    └── Methods (e.g., Save, SaveAs, Close, etc.)
│    └── … (other Notebook objects)
├── Properties (e.g., Visible, DefaultPath, StatusBar, etc.)
└── Methods (e.g., Quit, Help, etc.)
```

- A “WorksheetItem” typically gives access to a “DataTable” object for reading/writing cells.  
- A “GraphPageItem” provides the container for Graph objects (Axes, Plots, Legends, etc.).

1) Application Object  
• Represents the SigmaPlot application itself.  
• Provides access to global settings (Visible, DefaultPaths, StatusBar, etc.) and the open notebooks through the Notebooks collection.

2) Notebooks Collection  
• Contains all open Notebook objects.  
• A Notebook corresponds to a SigmaPlot file (.JNB).  
• Use Notebooks.Add to create a new Notebook or Notebooks(index) to reference an existing one.

3) Notebook Object  
• Represents a single SigmaPlot notebook file.  
• Holds multiple NotebookItem objects such as worksheets, graphs, reports, etc. within its NotebookItems collection.  
• Common properties include Name, Saved, and FullName; methods include Save and Close.

4) NotebookItems Collection  
• Contains all items (e.g., Worksheet, GraphPage, Report) in a given Notebook.  
• Use NotebookItems.Add to create new items, and NotebookItems(index) to reference an existing one.  
• Each item has an ItemType (e.g., 1 = worksheet, 2 = graph, etc.).

5) Worksheet Objects  
• Two main types: NativeWorksheetItem (SigmaPlot worksheet) and ExcelItem (embedded Excel sheet).  
• Use DataTable to read/write cell values.  
• Methods like Open, Import, Export, Paste, and transformations.

6) GraphItem / GraphPage / Plot Objects  
• GraphItem: Represents a top-level graph page (added via NotebookItems).  
• GraphPage: A container for one or more Graph objects.  
• Graph: Contains Plots, Axes, titles, and legends.  
• Plot: Can be lines, symbols, bars, error bars, etc. Each Plot references X, Y, and optionally Z data.

7) Axis, Legend, Text, and Other Graph Elements  
• Axis objects manage scaling and labeling (linear, log, date/time, categories).  
• Legend objects show symbolic references; can be turned on/off or modified.  
• Text, line, and symbol objects provide annotations within a graph.

8) Methods and Properties  
• Each object has methods (actions) and properties (settings/attributes).  
• Common actions include Open, Save, Close, Copy, Paste, and transformations (e.g., TransposePaste, SetSelectedObjectsAttribute).  
• Properties return or modify configuration details—for example, Graph.Name, Axis.Type, or Worksheet.Paste.

In short, SigmaPlot’s automation model organizes everything around the Application and Notebook objects, each Notebook containing multiple NotebookItems (worksheets, graphs, etc.). Graph objects contain Plots and Axes, while worksheets provide the data source. This hierarchy allows scripted manipulation of all aspects of SigmaPlot—creating new items, reading/writing worksheet cells, constructing graphs, and customizing plots or axes.

Below is an example VBA macro that illustrates how to use SigmaPlot’s object hierarchy to create a new workbook, add a worksheet, insert data, create a graph, customize the axes, and save the file. This example can be used inside SigmaPlot’s built-in macro editor or via an external VBA-capable application that references SigmaPlot’s automation library.

--------------------------------------------------------------------------------
``` vba
' Example SigmaPlot VBA Macro Demonstrating Object Hierarchies

Sub Main()
    Dim SPApp As Object          ' Top-level Application object
    Dim SPNotebook As Object     ' A SigmaPlot notebook (file)
    Dim SPWorksheet As Object    ' One of the items (worksheet) in the notebook
    Dim SPDataTable As Object    ' DataTable object inside the worksheet
    Dim SPPage As Object         ' A graph page item
    Dim SPGraph As Object        ' The actual graph object
    Dim PlotColumns(2) As Variant
    
    '--------------------------------------------------------------------------
    ' 1. Launch SigmaPlot (if not already running) and create a Notebook
    '--------------------------------------------------------------------------
    
    Set SPApp = CreateObject("SigmaPlot.Application.1")   ' or GetObject if already open
    SPApp.Visible = True                                  ' Make SigmaPlot visible
    
    Set SPNotebook = SPApp.Notebooks.Add()
    ' FIXED: Cannot set read-only property, comment out or initialize differently
    ' But how can I rename the notebook?
    MsgBox "SPNotebook.Name" & SPNotebook.Name
    MsgBox "SPNotebook.Title" & SPNotebook.Title
    ' Set SPNotebook.Title = "DemoNotebook"
    
    '--------------------------------------------------------------------------
    ' 2. Add a worksheet to the Notebook
    '--------------------------------------------------------------------------
    
    Set SPWorksheet = SPNotebook.NotebookItems.Add(1)  ' 1 = CT_WORKSHEET
    SPWorksheet.Name = "ExampleData"
    SPWorksheet.Open
    
    ' Access its DataTable for reading/writing cells
    Set SPDataTable = SPWorksheet.DataTable
    
    '--------------------------------------------------------------------------
    ' 3. Insert sample data (e.g., X and Y) into the worksheet
    '--------------------------------------------------------------------------
    
    ' For simplicity, put X in column 1, Y in column 2:
    Dim i As Long
    For i = 0 To 9
        ' Let X = i, Y = i^2, for demonstration
        SPDataTable.Cell(0, i) = i
        SPDataTable.Cell(1, i) = i * i
    Next i
    
    ' Assign column titles
    SPDataTable.Cell(0, -1) = "X Values"
    SPDataTable.Cell(1, -1) = "Y Values"
    
    '--------------------------------------------------------------------------
    ' 4. Create a new graph page and add a plot
    '--------------------------------------------------------------------------
    
    Set SPPage = SPNotebook.NotebookItems.Add(2) ' 2 = CT_GRAPHICPAGE
    SPPage.Name = "MyGraphPage"
    
    ' Build a column array to pass to SigmaPlot’s CreateWizardGraph
    ' We want to plot col 1 vs col 0, so that’s:
    '   column 0 => X; column 1 => Y
    PlotColumns(0) = 0  ' X
    PlotColumns(1) = 1  ' Y
    
    ' The wizard: CreateWizardGraph(GraphType, GraphStyle, DataFormat, ColumnList)
    ' FIXED: Cannot set read-only property, comment out or initialize differently
    SPPage.CreateWizardGraph "Scatter Plot", "Simple Scatter", "XY Pair", PlotColumns '
    
    ' Grab the Graph object
    Set SPGraph = SPPage.GraphPages(0).Graphs(0)
    SPGraph.Name = "MyScatter"
    
    '--------------------------------------------------------------------------
    ' 5. Customize the graph
    '--------------------------------------------------------------------------
    
    ' Access the X and Y axes
    Dim XAxis As Object, YAxis As Object
    Set XAxis = SPGraph.Axes(0)   ' 0 = X-axis
    Set YAxis = SPGraph.Axes(1)   ' 1 = Y-axis
    
    ' Name the axes
    XAxis.Name = "X Axis Title"
    YAxis.Name = "Y Axis Title"
    
    ' Remove the legend for clarity
    SPGraph.SetAttribute SGA_FLAGS, FlagOff(SGA_FLAG_AUTOLEGENDSHOW)
    
    ' Example: Manually set the X axis from -1 to 10
    XAxis.SetAttribute SAA_OPTIONS, FlagOff(SAA_FLAG_AUTORANGE)
    XAxis.SetAttribute SAA_FROMVAL, -1
    XAxis.SetAttribute SAA_TOVAL, 10
    
    ' Example: manually set Y axis from -5 to 90
    YAxis.SetAttribute SAA_OPTIONS, FlagOff(SAA_FLAG_AUTORANGE)
    YAxis.SetAttribute SAA_FROMVAL, -5
    YAxis.SetAttribute SAA_TOVAL, 90
    
    '--------------------------------------------------------------------------
    ' 6. Save the Notebook to disk
    '--------------------------------------------------------------------------
    
    Dim savePath As String
    savePath = "C:\Temp\DemoMacroNotebook.jnb"
    SPNotebook.SaveAs savePath
    
    '--------------------------------------------------------------------------
    ' 7. Optional: Close SigmaPlot or keep it open
    '--------------------------------------------------------------------------
    
    ' If you want to keep SigmaPlot open for further interaction, do nothing more.
    ' Otherwise, close the file and quit:
    
    'SPNotebook.Close True          ' True = Prompt to save if not already
    'SPApp.Quit                     ' Exits SigmaPlot entirely
    
    MsgBox "SigmaPlot demonstration macro is complete.", vbInformation, "Done"
End Sub
```

' End of Macro
--------------------------------------------------------------------------------

Explanation of Key Parts:

• CreateObject("SigmaPlot.Application.1"): Starts SigmaPlot automation and returns the Application object (SPApp).  
• SPApp.Notebooks.Add: Adds a new Notebook (the top-level SigmaPlot file).  
• SPWorksheet = SPNotebook.NotebookItems.Add(1): Inserts a new worksheet item into the current notebook; item type 1 means a SigmaPlot worksheet.  
• SPDataTable.Cell(Col, Row) = Value: Accesses or writes a value to the worksheet at the given zero-based column/row.  
• SPPage = SPNotebook.NotebookItems.Add(2): Adds a new graph page (type 2).  
• SPPage.CreateWizardGraph("Scatter Plot", "Simple Scatter", ...): Creates a new scatter-plot style graph on the page, referencing the assigned columns.  
• SPGraph.Axes(0) and SPGraph.Axes(1): Returns the X and Y axis objects (Axis(0) = X, Axis(1) = Y).  
• Setting attributes with XAxis.SetAttribute or YAxis.SetAttribute modifies scale, ticks, etc.  

This example shows creating data in a worksheet, making a graph page, customizing axes, and saving the notebook, demonstrating many of the common SigmaPlot objects in a logical sequence.

## Constants
## NotebookItem Type Constants (CT_XXX; Content Type)

These constants identify the item type when calling “NotebookItems.Add” or examining the “ItemType” property:

• CT_WORKSHEET = 1  
• CT_GRAPHICPAGE = 2  
• CT_FOLDER = 3  
• CT_STATTEST = 4  
• CT_REPORT = 6  
• CT_FIT = 7  
• CT_NOTEBOOK = 8  
• CT_EXCELWORKSHEET = 9  

(Note that some enumerations or intermediate values may differ in older documentation—these are the commonly used ones in recent SigmaPlot versions.)

--------------------------------------------------------------------------------
## Graph Child-Object Type Constants (GPT_XXX; Graph Page Type)

When navigating graph objects, SigmaPlot uses GPT_XXX constants for specifying child-object types (plots, axes, lines, etc.):

• GPT_GRAPH = 2  
• GPT_PLOT = 3  
• GPT_AXIS = 4  
• GPT_TEXT = 5  
• GPT_LINE = 6  
• GPT_SYMBOL = 7  
• GPT_SOLID = 8  
• GPT_TUPLE = 9  
• GPT_FUNCTION = 10  
• GPT_EXTERNAL = 11  
• GPT_BAG = 12  

These values appear for each “ChildObjects” item under a Graph or Plot, letting you loop through or discern specific subtypes (e.g., lines, symbols, text objects on the page).


<!-- ## VBA Data Type
 !-- • Byte
 !-- • Boolean
 !-- • Integer
 !-- • Long
 !-- • Currency
 !-- • Single
 !-- • Double
 !-- • Date
 !-- • String
 !-- • Variant
 !-- • Object (generic reference to an OLE object)
 !-- • User-defined classes or structures
 !-- • (In particular contexts) Decimal, LongLong, etc.
 !-- 
 !-- In SigmaPlot’s macro language (which is VBA-like), you typically see:
 !-- • As Variant
 !-- • As Object
 !-- • As String
 !-- • As Long, Double, etc.
 !-- • Or you might see Dim without a specific “As” (in which case it defaults to Variant).
 !-- 
 !-- 
 !-- Below are answers to each of your questions in order, focusing on SigmaPlot’s VBA-like macro environment, typical Visual Basic (VB) or VBA syntax, and the SigmaPlot-specific constants and object model. -->

--------------------------------------------------------------------------------
## Valid Data Types
• Byte
• Boolean
• Integer
• Long
• Currency
• Single
• Double
• Date
• String
• Variant (= nearly any kind of data (numbers, strings, objects, etc.))
• Object (generic reference to an OLE object)
• User-defined classes or structures
• (In particular contexts) Decimal, LongLong, etc.

In SigmaPlot’s macro language (which is VBA-like), you typically see:
• As Variant
• As Object
• As String
• As Long, Double, etc.
• Or you might see Dim without a specific “As” (in which case it defaults to Variant).

--------------------------------------------------------------------------------
## What is "SigmaPlot.Application.1"?

“SigmaPlot.Application.1” is the **Programmatic ID (ProgID)** that Windows uses to reference the SigmaPlot application’s OLE Automation server. It’s effectively a string used in CreateObject or GetObject calls:

–––––––––––––––––––––––––––––––
Dim SPApp As Object
Set SPApp = CreateObject("SigmaPlot.Application.1")
–––––––––––––––––––––––––––––––

This instructs Windows to launch (or connect to) SigmaPlot’s automation server. The “.1” typically indicates a major version or matching legacy ID that SigmaPlot registered in the Windows registry.

--------------------------------------------------------------------------------
3) What is SigmaPlot.Application."1"?

It’s simply part of the same ProgID string. Sometimes you see variations like “SigmaPlot.Application.14” for version 14, etc. “SigmaPlot.Application.1” is the canonical older ID. The "1" is not an index but part of the entire ProgID string. Some other OLE applications follow the pattern “Word.Application” or “Word.Application.16,” etc.

--------------------------------------------------------------------------------
4) Could you list internal keywords, like Dim, Set, As, For, To, Next, MsgBox, and is FlagOff one of these?

Common built-in VB (or SigmaPlot macro) keywords include:
• Dim, Set, As, For, To, Next, If, Then, ElseIf, Else, End If, Select Case, Do, Loop, While, Wend, Function, Sub, Exit, With, etc.
• MsgBox is a built-in VB function for showing message boxes.
• Each keyword has a special meaning in the language syntax.

By contrast, FlagOff (or FlagOn) is NOT a built-in VB keyword. It’s simply a custom function or procedure that you might see in SigmaPlot macros:

––––––––––––––––––––––––––––––
Function FlagOff(flag As Long) As Long
    ' Custom code
End Function
––––––––––––––––––––––––––––––

Hence, FlagOff is not a standard VB keyword but a user-defined or library-defined function.

--------------------------------------------------------------------------------
## Logging

``` vba
Sub LogToFileExample()
    Dim f As Integer
    f = FreeFile()
    Open "C:\MyLog.txt" For Append As #f
    Print #f, "This is a log message."
    Close #f
End Sub
```

––––––––––––––––––––––––––––––

• FreeFile() finds an available file handle number.  
• Open … For Append As #f opens (or creates, if nonexistent) a file for appending text.  
• Print #f, "Message" writes the text.  
• Close #f closes the file.  

Or you can open for Output (which overwrites the file each time). In SigmaPlot’s macro environment, the same approach generally works:

––––––––––––––––––––––––––––––

``` vba
Dim FileName As String
FileName = "C:\MyLog.txt"
Dim f As Integer
f = FreeFile()
Open FileName For Append As #f
Print #f, "Log line or data here"
Close #f
```

––––––––––––––––––––––––––––––

This is a typical way to do textual logging in older VB-like environments.

--------------------------------------------------------------------------------
## SGA_ constants (Set Graph Attribute)

SigmaPlot uses SGA_ constants in methods like .SetAttribute for certain graph-level properties. Examples commonly encountered:

• SGA_FLAGS  
• SGA_FLAG_AUTOLEGENDSHOW  
• SGA_FLAG_AUTOLEGENDBOX  
• SGA_PLANECOLORXYBACK  
• SGA_NTHAUTOLEGEND  
• SGA_CURRENTLEGENDTEXT  
• SGA_SHOWNAME  

Each SGA_ constant manages a particular high-level Graph attribute (like turning the auto legend on/off, setting the graph background color, etc.).

--------------------------------------------------------------------------------

## SAA_ constants (Set Axis Attribute)

SAA_ stands for “Set Axis Attribute.” SigmaPlot uses SAA_ constants for properties and behaviors of an Axis object. Common SAA_ constants include:

• SAA_TYPE (e.g., SAA_TYPE_LINEAR, SAA_TYPE_COMMON, SAA_TYPE_LOG, SAA_TYPE_DATETIME, etc.)  
• SAA_FROMVAL, SAA_TOVAL (manually set min/max scale on an axis)  
• SAA_FLAG_AUTORANGE, SAA_FLAG_NOAUTOPAD  
• SAA_MAJORFREQINDIRECT, SAA_MAJORFREQ, SAA_MINORFREQ, etc.  
• SAA_SUB1OPTIONS, SAA_SUB2OPTIONS (sub-axis or second scale line options)  
• SAA_AXISOPTIONS (various axis-level attributes)

By calling something like:

––––––––––––––––––––––––––––––
AxisObject.SetAttribute SAA_TYPE, SAA_TYPE_LOG
AxisObject.SetAttribute SAA_FROMVAL, 0.1
AxisObject.SetAttribute SAA_TOVAL, 100
––––––––––––––––––––––––––––––

…you would change the axis to a log scale from 0.1 to 100.

--------------------------------------------------------------------------------
8) Could you elaborate the tree object by adding comments with available properties, methods, and sub-objects?

Below is a more detailed tree, with short mentions of common properties/methods. (Not an exhaustive list, but a structured reference.)

```markdown
Application (SigmaPlot.Application)
├── [Properties]
│    ├── Visible (Boolean)
│    ├── DefaultPath (String)
│    ├── StatusBar (String)
│    └── … (Ventions like Notebooks, etc.)
├── [Methods]
│    ├── Quit()
│    ├── Help(ContextID, Kword, FileName)
│    └── … (CreateNewNotebook, etc.)
├── Notebooks (Collection)
│    ├── Item(index Or "Name") -> Notebook
│    ├── Add() -> Notebook
│    └── Count -> number of open notebooks
└── Notebook
     ├── [Properties]
     │    ├── Name (String)
     │    ├── FullName (String)
     │    ├── Saved (Boolean)
     │    ├── NotebookItems (Collection)
     │    └── Visible (Boolean, toggles the entire doc window)
     ├── [Methods]
     │    ├── Save(), SaveAs(FileName), Close(SaveChanges)
     │    ├── Open(), Import(), Export(), etc.
     │    └── … 
     ├── NotebookItems (Collection)
     │    ├── Item(index Or "Name") -> WorksheetItem / GraphPageItem / etc.
     │    ├── Add(CT_WORKSHEET or CT_GRAPHICPAGE, etc.)
     │    └── Count
     └── (subclasses of NotebookItem with different CT_XXX Type)
          ├── WorksheetItem
          │    ├── DataTable (for cells)
          │    ├── [Properties: Name, IsOpen, etc.]
          │    ├── [Methods: Open, Paste, Clear, etc.]
          │    └── … 
          ├── GraphPageItem
          │    ├── GraphPages(0) -> Graph
          │    ├── GraphPages.Count
          │    ├── [CreateWizardGraph, AddWizardPlot, etc.]
          │    └── … 
          ├── FitItem
          ├── ReportItem
          ├── ExcelItem
          └── etc.
```

Where each major object has:
• Properties that get or set attributes (like Name, Visible).  
• Methods that do things (Open, Save, Paste, etc.).  
• Sub-objects or collections referencing lower-level items (e.g., DataTable for a Worksheet, GraphPages for a GraphPageItem, and so on).

This breakdown highlights both the primary nodes (Application → Notebooks → Notebook → NotebookItems) and the specialized items (Worksheet, Graph, Fit, etc.), including the relevant constants (CT_WORKSHEET, CT_GRAPHICPAGE, …) or GPT_### for sub-graph objects.

<!-- EOF -->