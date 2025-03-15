#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-15 04:25:37 (ywatanabe)"
# File: /home/ywatanabe/proj/SigMacro/PySigMacro/dev.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/proj/SigMacro/PySigMacro/dev.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

"""
Main SigmaPlot automation class
"""
import subprocess
import time
import pysigmacro as ps
import pandas as pd
import numpy as np

# Params
FILENAME = f"SigmaPlot_Basic_{time.strftime('%Y%m%d_%H%M%S')}.JNB"
PATH = os.path.join("C:\\Temp", FILENAME)

# Initialization
ps.utils.close_all()

# Connection
sp = ps.utils.open()

# Python Wrapper
spw = ps.utils.com_wrap(sp)
app = spw.Application_obj
notebooks = app.Notebooks_obj
notebooks.Count
active_doc = app.ActiveDocument_obj
notebook_items = active_doc.NotebookItems_obj

notebook_item = notebook_items.Add(ps.const.CT_NOTEBOOK)
worksheet_item = notebook_items.Add(ps.const.CT_WORKSHEET)
datatable_obj = worksheet_item.DataTable_obj
graph_item = notebook_items.Add(ps.const.CT_GRAPHICPAGE)

graph_pages = graph_item.GraphPages_obj
graph_pages.Add(1)


# Import CSV Data
# df = pd.read_csv('your_file.csv')
df = pd.DataFrame(columns=["aaa", "bbb"], data=np.random.rand(3, 2))
header_list = [list(df.columns)]
df_T = df.T
data_list = df_T.values.tolist()
datatable_obj.PutData(data_list, 0, 0)

# Set Column names; FIXME


# Plotting
plotted_columns = [0, 1]
# Call CreateWizardGraph with required parameters: graph type, graph style, data format, and plotted columns
result = graph_item.CreateWizardGraph(
    "Vertical Bar Chart",
    "Simple Bar",
    "XY Pair",
    plotted_columns
)
print("Vertical Bar Chart creation result:", result)

# Example 2: Create a Scatter Plot with error bars and regression lines
plotted_columns = [0, 1, 2, 3, 4, 6, 7, 8, 9, 10]
# Define the columns per plot as a list
columns_per_plot = [5, 5]
# Call CreateWizardGraph with additional optional parameters:
# columns per plot, error bar source ("Column Means") and error bar computation ("Standard Deviation")
result = graph_page2.CreateWizardGraph(
    "Scatter Plot",
    "Multiple Error Bars & Regression",
    "X Many Y",
    plotted_columns,
    columns_per_plot,
    "Column Means",
    "Standard Deviation"
)
print("Scatter Plot creation result:", result)


# # Access the GraphPages collection of the graph item
# graph_pages = graph_item.GraphPages_obj
# Retrieve the first graph page using 1-based indexing with bracket notation

graph_pages._com_object.Item("0")
graph_pages._com_object.Item("1")
graph_page = graph_pages.Item(0)
graph_page = graph_pages["1"]
graph_page = graph_pages("1")
graph_page = graph_pages.Item("1")
# Retrieve the first graph on that page using 1-based indexing
graph = graph_page.Graphs_obj.Item(1)
# Plot data from the DataTable using the first two columns for x and y
graph.PlotXY(datatable_obj, 0, 1, 0, 2)

from win32com.client import VARIANT
import pythoncom

from win32com.client.util import Iterator
graph_page = next(Iterator(graph_pages))

graph_page = graph_pages._com_object.Item(VARIANT(pythoncom.VT_I2, 1))


notebook_items.Aunthor

notebooks.Author(1)  # NoneType object is not subscriptable
notebooks.Author[1]  # NoneType object is not subscriptable
notebooks[1]
cur_obj = notebooks.CurrentItem_obj

# # AddOnLocation
# app.AddOnLocation("Enzyme Kinetics")


# # Graph
# # GraphObject.AutoLegend
# # ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).AutoLegend

# # Axes
# # GraphObject.Axes
# # AxisObject.AxisTitles
# # AxisObject.AxisTitles(0).Name = "Bottom X Axis Title"
# # AxisObject.AxisTitles(1).Name = "Left Y Axis Title"
# # AxisObject.AxisTitles(0).SetAttribute(STA_ORIENTATION, 0)

# # Data
# # Datatable
# # Cell
# # DataTableObject.Cell(i_col, i_row)
# # ActiveDocument.NotebookItems("Data 1").DataTable.Cell(0,0)
# # PutData method is much faster for placing operation

# # ChildObjects Prop
# # PageObj.ChildObjects

# # Color Prop

# # Comments Prop
# # Count Prop

# # NotebookObj.CurrentBrowserItem Prop
# # If ActiveDocument.CurrentBrowserItem.Saved=True Then

# # CurrentDataItem Prop
# # NotebookObj.CurrentDataItem
# # ActiveDocument.CurrentDataItem.Interpolate3DMesh(1,2,3) # Creates interpolated mesh data for columns 1, 2 and 3 and places them in the first empty column.

# # GetMaxusedsize Prop
# # DataTableObj.GetMaxUsedSize(i_col, i_row)

# # CurrentDateString Prop
# # ApplicationObj.CurrentDateString(DatePicture)
# # Ex.) MsgBox(Application.CurrentDateString("MMMM d, yyyy"),0+64,"Today's Date")

# # CurrentItem Prop
# # NotebookObj.CurrentItem
# # NotebookObj.CurrentItem.Name = "XXX"

# # CurrentPageItem Prop.
# # Syntax: NotebookObj.CurrentPageItem

# # ApplyPageTemplate
# # Syntax: ActiveDoccument.CurrentPageItem.ApplyPageTemplate("Scatter Plot")

# # ActiveDocument Prop.

# # NotebookItems.Add

# # Line Prop.
# # PlotObj.Line
# # Ex.)
# # Dim SPLine As Object Set SPLine = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Plots(0).Line
# # SPLine.SetAttribute(SEA_THICKNESS,50)
# # SPLine.Color = RGB_DKRED
# # Changes the line color for the first plot to dark red and the line thickness to 0.05 inches.


# # ----------------------------------------
# # Axis
# # ----------------------------------------
# # Line Attributes Prop.
# # AxisObj.LineAttributes
# # 1: Axis Lines
# # 2: Major Ticks
# # 3: Minor Ticks
# # 4: Major Grid
# # 5: Minor Grid
# # 6: Axis Break
# # Ex.)
# # Dim SPHoriz,SPVert Set SPHoriz = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Axes(0).LineAttributes(1)
# # Set SPVert = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Axes(1).LineAttributes(1) SPHoriz.Color(RGB_BLUE) SPVert.Color(RGB_RED) Set SPHoriz = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Axes(0).LineAttributes(4)
# # Set SPVert = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Axes(1).LineAttributes(4) SPHoriz.SetAttribute(SEA_LINETYPE,6) SPVert.SetAttribute(SEA_LINETYPE,6) SPHoriz.Color(RGB_GRAY) SPVert.Color(RGB_GRAY) Dim i,breakstatus,brkparam(2) For i=0 To 1 Set SPHoriz = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Axes(i) breakstatus=SPHoriz.GetAttribute(SAA_BREAKON,brkparam(i)) If breakstatus=1 Then SPHoriz.LineAttributes(6).Color(RGB_BLACK) SPHoriz.SetAttribute(SAA_BREAKTYPE,2) SPHoriz.LineAttributes(6).SetAttribute(SEA_LINETYPE,6) End If Next i Changes the horizontal axis lines to blue and the vertical axis lines to red. Gridlines for both axes are set to a gray, dotted style. In addition, if either axis contains a break, the break appears as two black, diagonal, dotted, parallel lines.


# app.Notebooks(4).NotebookItems
# app.Notebooks(4).NotebookItems.Item(2)
# app.Notebooks("Notebook 2").NotebookItems.Item(1)


# app.Notebooks(ii).Name = "My Notebook"
# app.Notebooks("Notebook2").Name  # Notebook2
# app.Notebooks("Notebook2").Title = "My Notebook"
# sp.Notebooks.Close(3)
# sp.Notebooks.Close(1)  # all closed


# sp.Open
# sp.SaveAs


# # notebook = app.Notebooks(5)
# # notebook.Name  # Notebook 2
# # notebook.Close(0)  # Notebook 2 is closed
# # app.ActiveDocument.Name  # ''

# # notebook = app.Notebooks(3)
# # notebook.Name  # Notebook 1
# # notebook.Close(0)  # Notebook 1 is closed and no notebook opened

# # notebook.Name
# # PATH = os.path.join(
# #     "C:\\Temp", f"SigmaPlot_Basic_{time.strftime('%Y%m%d_%H%M%S')}.JNB"
# # )
# # notebook.SaveAs()
# # notebook.Close(0)
# # app.Notebooks.Open()
# # app.Notebooks.Save()
# # app.ActiveDocument.Name
# # app.Notebooks.Close()

# # EOF

# CreateWizardGraph Method

# Objects

# Type: Function
# Results: Boolean

# Syntax: GraphItem object.CreateWizardGraph(required parameters variants, optional parameters variants)

# Creates a graph in the specified GraphItem object using the Graph Wizard options. These options are expressed using the following parameters:

# Parameter
#  Values
#  Optional

# graph type
#  any valid type name
#  no

# graph style
#  any valid style name
#  no

# data format
#  any valid data format name
#  no

# columns plotted
#  any column number/title array
#  no

# columns per plot
#  array of columns in each plot
#  yes

# error bar source
#  any valid source name
#  error bar plots only

# upper error bar computation
#  any valid computation name
#  error bar plots only

# anglular axis units
#  any valid angle unit name
#  polar plots only

# lower range bound
#  any valid degree value
#  polar plots only

# upper range bound
#  any valid degree value
#  polar plots only

# ternary units
#  upper range of ternary axis scale
#  ternary plots only

# lower error bar computation
#  any valid computation name
#  error bar plots only

# row selection
#  Boolean: True allows selection of a row range for y-replicate (row-summary) plots. Use False to support pre-y replicate data format macros.
#  Row summary plots only

# Examples


# --------------------------------------------------------------------------------
# ActiveDocument.NotebookItems.Add(2) 'Adds a new graph page
# Dim PlottedColumns(1) As Variant
# PlottedColumns(0) = 0
# PlottedColumns(1) = 1
# ActiveDocument.NotebookItems("Graph Page 1").CreateWizardGraph("Vertical Bar Chart", _
# "Simple Bar","XY Pair",PlottedColumns)

# Plots columns 1 and 2 as a simple bar chart

# Dim GraphPage As Object
# Set GraphPage = ActiveDocument.NotebookItems.Add(CT_GRAPHICPAGE) 'Adds a new graph page
# Dim PlottedColumns(9) As Variant
# PlottedColumns(0) = 0
# PlottedColumns(1) = 1
# PlottedColumns(2) = 2
# PlottedColumns(3) = 3
# PlottedColumns(4) = 4
# PlottedColumns(5) = 6
# PlottedColumns(6) = 7
# PlottedColumns(7) = 8
# PlottedColumns(8) = 9
# PlottedColumns(9) = 10
# Dim ColumnsPerPlot(1) As Variant
# ColumnsPerPlot(0) = 5
# ColumnsPerPlot(1) = 5 'remaining columns are automatically plotted
# GraphPage.CreateWizardGraph("Scatter Plot", _
# "Multiple Error Bars & Regression","X Many Y",PlottedColumns,ColumnsPerPlot, _
# "Column Means","Standard Deviation")

# Plots columns 1-5 and 7-11 as column averaged scatter plots with error bars and regression lines.

# EOF