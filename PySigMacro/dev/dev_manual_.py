#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-14 23:08:49 (ywatanabe)"
# File: /home/ywatanabe/proj/SigMacro/PySigMacro/dev_manual.py

import os

__THIS_FILE__ = (
    "/home/ywatanabe/proj/SigMacro/PySigMacro/dev_manual.py"
)
__THIS_DIR__ = os.path.dirname(__THIS_FILE__)

"""
Main SigmaPlot automation class
"""
import subprocess
import time

import win32com.client

def close_all_sigmaplot_processes():
    try:
        subprocess.run(
            ["taskkill", "/f", "/im", "spw.exe"],
            shell=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        time.sleep(2)
    except Exception as e:
        print(f"Warning when closing SigmaPlot: {e}")


from pysigmacro.utils import com_wrap, com_dir

# Initialization
close_all_sigmaplot_processes()

# Connection
sp = win32com.client.Dispatch("SigmaPlot.Application")
app = sp.Application

# Visibility
sp.Visible = False
app.Visible = True

# Preparation
FILENAME = f"SigmaPlot_Basic_{time.strftime('%Y%m%d_%H%M%S')}.JNB"
PATH = os.path.join(
    "C:\\Temp", FILENAME
)

# SigmaPlot
sp.FullName  # C:\\Program Files (x86)\\SigmaPlot\\SPW16\\Spw.exe
sp.Name  # SigmaPlot 15
sp.Notebooks
# sp.Quit() # Quit

# Application
app.Name
app.FullName
app.Notebooks
app.ActiveDocument
app.DefaultPath

# Notebooks
sp.Notebooks.Add
app.Notebooks.Count  # 4
for ii in range(app.Notebooks.Count):
    print("----------------------------------------")
    print(ii)
    print(app.Notebooks(ii).Name)
    print(app.Notebooks(ii).Title)
    print("----------------------------------------")

# Notebook
app.Notebooks(4).Author # ywatanabe@alumni.u-tokyo.ac.jp
app.Notebooks(4).SaveAs(PATH)
app.Notebooks(4).Save()

# Active Document
sp.ActiveDocument.FullName # 'C:\\Temp\\SigmaPlot_Basic_20250314_191623.JNB'
sp.ActiveDocument.Name # 'SigmaPlot_Basic_20250314_191623.JNB'
sp.ActiveDocument.NotebookItems(2).Author
sp.ActiveDocument.NotebookItems(2).Name
sp.ActiveDocument.NotebookItems(2).Close(True)

# AddOnLocation
app.AddOnLocation("Enzyme Kinetics")


# Graph
# GraphObject.AutoLegend
# ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).AutoLegend

# Axes
# GraphObject.Axes
# AxisObject.AxisTitles
# AxisObject.AxisTitles(0).Name = "Bottom X Axis Title"
# AxisObject.AxisTitles(1).Name = "Left Y Axis Title"
# AxisObject.AxisTitles(0).SetAttribute(STA_ORIENTATION, 0)

# Data
# Datatable
# Cell
# DataTableObject.Cell(i_col, i_row)
# ActiveDocument.NotebookItems("Data 1").DataTable.Cell(0,0)
# PutData method is much faster for placing operation

# ChildObjects Prop
# PageObj.ChildObjects

# Color Prop

# Comments Prop
# Count Prop

# NotebookObj.CurrentBrowserItem Prop
# If ActiveDocument.CurrentBrowserItem.Saved=True Then

# CurrentDataItem Prop
# NotebookObj.CurrentDataItem
# ActiveDocument.CurrentDataItem.Interpolate3DMesh(1,2,3) # Creates interpolated mesh data for columns 1, 2 and 3 and places them in the first empty column.

# GetMaxusedsize Prop
# DataTableObj.GetMaxUsedSize(i_col, i_row)

# CurrentDateString Prop
# ApplicationObj.CurrentDateString(DatePicture)
# Ex.) MsgBox(Application.CurrentDateString("MMMM d, yyyy"),0+64,"Today's Date")

# CurrentItem Prop
# NotebookObj.CurrentItem
# NotebookObj.CurrentItem.Name = "XXX"

# ItemType Prop
# NotebookObj.CurrentItem.ItemType
# CT_WORKSHEET; SigmaPlot Worksheet
# CT_GRAPHICPAGE; Graph Page
# CT_FOLDER: Section
# CT_STATTEST: SigmaStat Reoprt
# CT_REPORT: SigmaPlot Report
# CT_FIT: Equation
# CT_NOTEBOOK: Notebook
# CT_EXCELWORKSHEET: Excel Worksheet
# CT_TRANSFORMTransform
# Macro

# CurrentPageItem Prop.
# Syntax: NotebookObj.CurrentPageItem

# ApplyPageTemplate
# Syntax: ActiveDoccument.CurrentPageItem.ApplyPageTemplate("Scatter Plot")

# ActiveDocument Prop.

# NotebookItems.Add

# Line Prop.
# PlotObj.Line
# Ex.)
# Dim SPLine As Object Set SPLine = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Plots(0).Line
# SPLine.SetAttribute(SEA_THICKNESS,50)
# SPLine.Color = RGB_DKRED
# Changes the line color for the first plot to dark red and the line thickness to 0.05 inches.


# ----------------------------------------
# Axis
# ----------------------------------------
# Line Attributes Prop.
# AxisObj.LineAttributes
# 1: Axis Lines
# 2: Major Ticks
# 3: Minor Ticks
# 4: Major Grid
# 5: Minor Grid
# 6: Axis Break
# Ex.)
# Dim SPHoriz,SPVert Set SPHoriz = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Axes(0).LineAttributes(1)
# Set SPVert = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Axes(1).LineAttributes(1) SPHoriz.Color(RGB_BLUE) SPVert.Color(RGB_RED) Set SPHoriz = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Axes(0).LineAttributes(4)
# Set SPVert = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Axes(1).LineAttributes(4) SPHoriz.SetAttribute(SEA_LINETYPE,6) SPVert.SetAttribute(SEA_LINETYPE,6) SPHoriz.Color(RGB_GRAY) SPVert.Color(RGB_GRAY) Dim i,breakstatus,brkparam(2) For i=0 To 1 Set SPHoriz = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Axes(i) breakstatus=SPHoriz.GetAttribute(SAA_BREAKON,brkparam(i)) If breakstatus=1 Then SPHoriz.LineAttributes(6).Color(RGB_BLACK) SPHoriz.SetAttribute(SAA_BREAKTYPE,2) SPHoriz.LineAttributes(6).SetAttribute(SEA_LINETYPE,6) End If Next i Changes the horizontal axis lines to blue and the vertical axis lines to red. Gridlines for both axes are set to a gray, dotted style. In addition, if either axis contains a break, the break appears as two black, diagonal, dotted, parallel lines.



app.Notebooks(4).NotebookItems
app.Notebooks(4).NotebookItems.Item(2)
app.Notebooks("Notebook 2").NotebookItems.Item(1)



app.Notebooks(ii).Name = "My Notebook"
app.Notebooks("Notebook2").Name # Notebook2
app.Notebooks("Notebook2").Title = "My Notebook"
sp.Notebooks.Close(3)
sp.Notebooks.Close(1)  # all closed


sp.Open
sp.SaveAs


# notebook = app.Notebooks(5)
# notebook.Name  # Notebook 2
# notebook.Close(0)  # Notebook 2 is closed
# app.ActiveDocument.Name  # ''

# notebook = app.Notebooks(3)
# notebook.Name  # Notebook 1
# notebook.Close(0)  # Notebook 1 is closed and no notebook opened

# notebook.Name
# PATH = os.path.join(
#     "C:\\Temp", f"SigmaPlot_Basic_{time.strftime('%Y%m%d_%H%M%S')}.JNB"
# )
# notebook.SaveAs()
# notebook.Close(0)
# app.Notebooks.Open()
# app.Notebooks.Save()
# app.ActiveDocument.Name
# app.Notebooks.Close()

# EOF