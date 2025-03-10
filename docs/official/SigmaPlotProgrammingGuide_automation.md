<!-- ---
!-- Timestamp: 2025-03-09 05:34:10
!-- Author: ywatanabe
!-- File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-SigMacro/SigMacro v1.3/references/SigmaPlotProgrammingGuide_automation.md
!-- --- -->

# SigmaPlot Automation with Pywin32: Summary

This document summarizes the key aspects of automating SigmaPlot using Pywin32, based on the previous detailed outputs. It provides a concise overview of the essential objects, properties, methods, and practical tips for Python developers looking to automate SigmaPlot tasks.

## Introduction to SigmaPlot Automation with Pywin32

SigmaPlot automation allows developers to control SigmaPlot programmatically using its COM interface. With Pywin32 (win32com.client), Python can interact with SigmaPlot to automate tasks like data manipulation, graph creation, and file management. The Application object is the entry point for all automation tasks.

## Key SigmaPlot Objects

- **Application**: Represents the SigmaPlot program itself. It provides access to notebooks and global settings.
- **Notebooks**: A collection of all open notebook files in SigmaPlot.
- **Notebook**: Represents a single SigmaPlot notebook file (e.g., .jnb), containing items like worksheets and graphs.
- **NotebookItems**: A collection of items within a notebook, such as worksheets, graph pages, and reports.
- **NativeWorksheetItem**: Represents a SigmaPlot worksheet for data entry and manipulation.
- **DataTable**: Allows direct access to worksheet data for reading and writing cell values.
- **GraphItem**: Represents a SigmaPlot graph page, used to create and modify graphs.

## Automatable Properties and Methods

### Application Object
**Properties**:
- Notebooks: Access to open notebooks.
- ActiveDocument: The currently active notebook.
- DefaultPath: Default file path for saving/loading.
- Visible: Controls application visibility.

**Methods**:
- Help(topic): Opens help for a specific topic.

### Notebooks Collection
**Properties**:
- Count: Number of open notebooks.
- Item(index): Access a specific notebook by index or name.

**Methods**:
- Add(): Creates a new notebook.
- Open(filename): Opens an existing notebook.

### Notebook Object
**Properties**:
- Name: Notebook name.
- NotebookItems: Collection of items in the notebook.
- Saved: Indicates if changes are saved.

**Methods**:
- Save(): Saves the notebook.
- SaveAs(filename): Saves to a new file.
- Close(): Closes the notebook.

### NotebookItems Collection
**Properties**:
- Count: Number of items.
- Item(index): Access a specific item.

**Methods**:
- Add(type): Adds a new item (e.g., worksheet, graph).
- Delete(index): Deletes an item.

### NativeWorksheetItem Object
**Properties**:
- DataTable: Access to worksheet data.
- Name: Worksheet name.

**Methods**:
- Open(): Opens the worksheet.
- Close(): Closes the worksheet.
- Import(filename): Imports data from a file.

### DataTable Object
**Properties**:
- Cell(column, row): Gets/sets a specific cell value.

**Methods**:
- GetData(left, top, right, bottom): Retrieves a range of data.
- PutData(left, top, data): Inserts data into the worksheet.

### GraphItem Object
**Properties**:
- GraphPages: Access to the graph page (currently one per graph item).
- Name: Graph page name.

**Methods**:
- CreateWizardGraph(options): Creates a graph using specified options.
- ApplyPageTemplate(template): Applies a template to the graph.

## Practical Example

Below is a simple automation script that creates a notebook, adds a worksheet, inputs data, and generates a scatter plot:

```python
import win32com.client

# Initialize SigmaPlot
sp = win32com.client.Dispatch("SigmaPlot.Application.1")
sp.Visible = True

# Create a new notebook
notebook = sp.Notebooks.Add()
notebook.Name = "TestNotebook"

# Add a worksheet
worksheet = notebook.NotebookItems.Add(1)  # 1 = Worksheet
worksheet.Name = "DataSheet"
data_table = worksheet.DataTable

# Input sample data
data_table.Cell(1, 1) = 1
data_table.Cell(1, 2) = 2
data_table.Cell(2, 1) = 3
data_table.Cell(2, 2) = 4

# Add a graph page
graph_item = notebook.NotebookItems.Add(2)  # 2 = Graph Page
graph_item.Name = "ScatterGraph"

# Create a scatter plot
graph_item.CreateWizardGraph({
    "graph type": "Scatter",
    "graph style": "Simple Scatter",
    "data format": "XY Pair",
    "columns plotted": [1, 2]
})

# Save the notebook
notebook.SaveAs("C:\\Data\\test_notebook.jnb")
```

## Tips for Automation

- **Object Hierarchy**: Always navigate from Application → Notebooks → Notebook → NotebookItems to access worksheets or graphs.
- **Constants**: Use SigmaPlot constants (e.g., for graph attributes) with methods like SetAttribute. Refer to SigmaPlot's Automation Help for details.
- **Error Handling**: Use try-except blocks to handle COM errors gracefully.
- **Indexing**: Assume 1-based indexing for properties like Cell unless specified otherwise.

This summary provides a quick reference for automating SigmaPlot with Pywin32, focusing on the most commonly used objects, properties, and methods. For more detailed information, refer to the full SigmaPlot Automation Reference or the original outputs.

<!-- EOF -->