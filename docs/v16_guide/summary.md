<!-- ---
!-- Timestamp: 2025-03-10 08:53:33
!-- Author: ywatanabe
!-- File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/docs/from_Guide/summary.md
!-- --- -->

**Summary of SigmaPlot VBA Properties, Methods, and Functions**

SigmaPlot is a powerful scientific graphing and data analysis software that supports automation through a VBA-like macro language. This allows users to automate routine tasks, manipulate data, and customize graphs programmatically. Below is an overview of the key properties, methods, functions, and objects used in SigmaPlot's macro language to help you get started.

---

### **1. Understanding the SigmaPlot Object Model**

The SigmaPlot object model is hierarchical, consisting of various objects and collections that represent different components within SigmaPlot. Here's a breakdown of the main objects:

- **Application Object**: Represents the SigmaPlot application itself.
- **Notebooks Collection**: Contains all open notebooks.
- **Notebook Object**: Represents a single notebook, which may contain worksheets, graphs, reports, etc.
- **NotebookItems Collection**: Holds all items (worksheets, graphs, etc.) within a notebook.
- **Worksheet Objects**:
  - **NativeWorksheetItem**: Represents a SigmaPlot worksheet.
  - **ExcelItem**: Represents an embedded Excel worksheet.
- **DataTable Object**: Represents data within a worksheet or graph.
- **GraphItem Object**: Represents a graph page.
- **GraphPage Object**: Contains graphs and page-level settings.
- **Graph Object**: Represents a single graph within a graph page.
- **Plot Object**: Represents a data plot within a graph.
- **Axis Object**: Represents an axis on a graph.
- **Other Objects**: Include Text, Line, Symbol, Solid, Function, Tuple, etc.

---

### **2. Key Properties and Methods**

Below are some of the most commonly used properties and methods for these objects.

#### **Application Object**

- **Properties**:
  - `Visible`: Sets or returns the visibility of the SigmaPlot application.
  - `Notebooks`: Returns the collection of open notebooks.
  - `DefaultPath`: Gets or sets the default file path for opening and saving files.
  - `StatusBar`: Sets the text displayed in the status bar.

- **Methods**:
  - `Quit()`: Closes the SigmaPlot application.
  - `Help(ContextID, Kword, FileName)`: Opens the help file to a specific topic.

**Example**:

```vb
Dim SPApp As Object
Set SPApp = CreateObject("SigmaPlot.Application.1")
SPApp.Visible = True
```

#### **Notebook Object**

- **Properties**:
  - `Name`: Gets or sets the name of the notebook.
  - `FullName`: Returns the full path and name of the notebook.
  - `NotebookItems`: Returns the collection of items within the notebook.
  - `Saved`: Indicates whether the notebook has unsaved changes.
  - `Author`, `Comments`, `Keywords`, `Subject`, `Title`: Metadata properties.

- **Methods**:
  - `Save()`: Saves the notebook.
  - `SaveAs(FileName)`: Saves the notebook with a new name.
  - `Close(SaveChanges As Boolean, FileName As String)`: Closes the notebook.

**Example**:

```vb
Dim nb As Object
Set nb = SPApp.Notebooks.Add()
nb.Name = "MyNotebook"
```

#### **NotebookItems Collection**

- **Methods**:
  - `Add(ItemType As Integer)`: Adds a new item (e.g., worksheet, graph) to the notebook.
    - `1`: Worksheet (`CT_WORKSHEET`)
    - `2`: Graph (`CT_GRAPHICPAGE`)
    - `8`: Excel Worksheet (`CT_EXCELWORKSHEET`)
  - `Item(Index As Variant)`: Returns the item at the specified index or with the specified name.

**Example**:

```vb
Dim ws As Object
Set ws = nb.NotebookItems.Add(1) ' Adds a new SigmaPlot worksheet
ws.Name = "DataSheet"
ws.Open()
```

#### **NativeWorksheetItem and ExcelItem Objects**

- **Properties**:
  - `DataTable`: Accesses the data table within the worksheet.
  - `Name`: Gets or sets the worksheet name.
  - `IsOpen`: Indicates whether the worksheet is open.

- **Methods**:
  - `Open()`: Opens the worksheet.
  - `Close(SaveChanges As Boolean)`: Closes the worksheet.
  - `Import(FileName As String)`: Imports data from a file.
  - `Export(FileName As String)`: Exports data to a file.
  - `Paste()`: Pastes clipboard content into the worksheet.
  - `Select(Left, Top, Right, Bottom)`: Selects a range of cells.

**Example**:

```vb
ws.Import("C:\Data\experiment.txt")
```

#### **DataTable Object**

- **Properties**:
  - `Cell(Column As Long, Row As Long)`: Gets or sets the value of a specific cell.
  - `NamedRanges`: Accesses named data ranges (e.g., column titles).

- **Methods**:
  - `GetData(Left As Long, Top As Long, Right As Long, Bottom As Long) As Variant`: Retrieves data from specified cells.
  - `PutData(Left As Long, Top As Long, Data As Variant)`: Inserts data into specified cells.
  - `GetMaxUsedSize(ByRef LastColumn As Long, ByRef LastRow As Long)`: Retrieves the used range.

**Example**:

```vb
Dim data As Variant
data = ws.DataTable.GetData(1, 1, 3, 100) ' Gets data from columns 1-3, rows 1-100

ws.DataTable.PutData(1, 1, data) ' Puts data back into worksheet starting at cell (1,1)
```

#### **GraphItem Object**

- **Properties**:
  - `Name`: Gets or sets the graph name.
  - `GraphPages`: Returns the collection of pages within the graph item.
  - `IsOpen`: Indicates whether the graph is open.

- **Methods**:
  - `Open()`: Opens the graph item.
  - `Close(SaveChanges As Boolean)`: Closes the graph item.
  - `CreateWizardGraph(GraphType, GraphStyle, DataFormat, ColumnArray)`: Creates a graph using the Graph Wizard.
  - `ApplyPageTemplate(TemplateName As String, SourceNotebook As Object)`: Applies a template to the graph.

**Example**:

```vb
Dim gr As Object
Set gr = nb.NotebookItems.Add(2) ' Adds a new graph item
gr.Name = "ResultsGraph"

' Create a line plot using the Graph Wizard
gr.CreateWizardGraph "Line", "Simple 2D Line", "Y Column(s) vs X Column", Array("XData", "YData")
gr.Open()
```

#### **Graph Object**

- **Properties**:
  - `Plots`: Returns the collection of plots on the graph.
  - `Axes`: Returns the collection of axes.
  - `AutoLegend`: Accesses the graph's legend object.
  - `Name`: Gets or sets the graph name.

**Accessing the Graph Object**:

```vb
Dim graph As Object
Set graph = gr.GraphPages(0).Graphs(0) ' Accesses the first graph on the first page
```

#### **Plot Object**

- **Properties**:
  - `Line`: Accesses line properties (e.g., color, style).
  - `Symbols`: Accesses symbol properties (e.g., type, size).
  - `Fill`: Accesses fill properties (e.g., for bar charts).

- **Methods**:
  - `SetAttribute(Attribute As Long, Value As Variant)`: Sets plot attributes.
  - `GetAttribute(Attribute As Long) As Variant`: Gets plot attribute values.

**Example**:

```vb
Dim plot As Object
Set plot = graph.Plots(0)
plot.SetAttribute SLA_LINE_COLOR, vbBlue ' Sets the line color to blue
```

*Note: `SLA_LINE_COLOR` is a constant representing the line color attribute.*

#### **Axis Object**

- **Methods**:
  - `SetAttribute(Attribute As Long, Value As Variant)`: Sets axis attributes.
  - `GetAttribute(Attribute As Long) As Variant`: Gets axis attribute values.

**Example**:

```vb
Dim xAxis As Object
Set xAxis = graph.Axes(1) ' 1 corresponds to the X-axis
xAxis.SetAttribute SLA_SCALETYPE, 1 ' Sets the scale type to linear (1)
xAxis.SetAttribute SLA_AXISRANGE_LOWER, 0
xAxis.SetAttribute SLA_AXISRANGE_UPPER, 100
```

---

### **3. Commonly Used Constants**

SigmaPlot uses constants to represent various attributes and object types. Some examples include:

- **Item Types**:
  - `1` (`CT_WORKSHEET`): Worksheet
  - `2` (`CT_GRAPHICPAGE`): Graph
  - `8` (`CT_EXCELWORKSHEET`): Excel Worksheet

- **Scale Types** (`SLA_SCALETYPE`):
  - `1`: Linear
  - `2`: Logarithmic (Base 10)

- **Axis Attributes**:
  - `SLA_SCALETYPE`: Sets the axis scale type.
  - `SLA_AXISRANGE_LOWER`: Sets the lower limit of the axis.
  - `SLA_AXISRANGE_UPPER`: Sets the upper limit of the axis.

- **Plot Attributes**:
  - `SLA_LINE_COLOR`: Sets the line color.
  - `SLA_SYMBOL_TYPE`: Sets the symbol type.
  - `SLA_FILL_PATTERN`: Sets the fill pattern for bars.

- **Color Constants**:
  - `vbBlack`, `vbBlue`, `vbRed`, `vbGreen`, etc.

---

### **4. Writing Macros and Procedures**

Macros in SigmaPlot are similar to VBA macros in other applications.

**Creating a Macro**:

```vb
Sub Main()
    ' Your automation code here
End Sub
```

**Running a Macro**:

- Use the Macro Recorder available in SigmaPlot to record actions and generate macro code.
- Access your macros via the Tools > Macro menu.

**Example Macro**:

```vb
Sub Main()
    ' Create a new workbook
    Dim nb As Object
    Set nb = Application.Notebooks.Add()
    
    ' Add a worksheet
    Dim ws As Object
    Set ws = nb.NotebookItems.Add(1)
    ws.Name = "DataSheet"
    ws.Open()
    
    ' Import data
    ws.Import("C:\Data\experiment.txt")
    
    ' Add a graph
    Dim gr As Object
    Set gr = nb.NotebookItems.Add(2)
    gr.Name = "ResultsGraph"
    
    ' Create a line plot
    gr.CreateWizardGraph "Line", "Simple 2D Line", "Y Column(s) vs X Column", Array("Time", "Measurement")
    gr.Open()
End Sub
```

---

### **5. Tips for Effective Automation**

- **Use the Macro Recorder**: Record repetitive tasks to generate macro code and learn how SigmaPlot commands translate into code.
- **Explore the Object Browser**: The Object Browser in the Macro Editor provides a list of objects, methods, properties, and constants.
- **Leverage Help Documentation**: SigmaPlot's help files offer detailed explanations and examples.
- **Debugging Tools**: Use the debugging features in the Macro Editor, such as breakpoints, the Immediate window, and stepping through code.
- **Comment Your Code**: Use comments (') to document your macros for future reference.
- **Error Handling**: Incorporate error handling to manage unexpected situations.

**Example of Error Handling**:

```vb
Sub Main()
    On Error GoTo ErrorHandler
    ' Your code here
    
    Exit Sub
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub
```

---

### **6. Additional Functions and Methods**

#### **Data Manipulation**

- `GetMaxUsedSize(ByRef LastColumn As Long, ByRef LastRow As Long)`: Retrieves the last used row and column in the worksheet.
- `NormalizeTernaryData(Column1, Column2, Column3, Optional Total As Double = 100)`: Normalizes data for ternary plots.

#### **Graph Customization**

- `AddWizardAxis(GraphItem, AxisType, AxisPosition)`: Adds an axis to a graph.
- `ModifyWizardPlot(GraphItem, GraphType, GraphStyle, DataFormat, ColumnArray)`: Modifies an existing plot.

#### **Exporting and Printing**

- `Export(FileName As String, ExportFormat As String)`: Exports graphs or worksheets to various formats (e.g., JPG, PNG, TXT).
- `Print()`: Sends the current item to the printer.

---

### **7. Resources for Learning More**

- **SigmaPlot Help Files**: Comprehensive documentation is available within SigmaPlot under the Help menu.
- **Macro Samples**: SigmaPlot provides sample macros that illustrate common automation tasks.
- **Community and Support**: Engage with the SigmaPlot user community through forums and support channels.
- **Books and Tutorials**: Consider learning resources on VBA programming to extend your automation skills.

---

By familiarizing yourself with these objects, properties, methods, and functions, you can harness the full power of SigmaPlot's automation capabilities to streamline your workflow and customize your data analysis and visualization tasks.

<!-- EOF -->