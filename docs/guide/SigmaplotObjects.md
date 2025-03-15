<!-- ---
!-- Timestamp: 2025-03-09 10:35:01
!-- Author: ywatanabe
!-- File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/docs/SigmaplotObjects.md
!-- --- -->

SigmaPlot Automation Reference
OLE Automation is a technology that lets other applications, development tools, and macro languages use a program. SigmaPlot Automation allows you to integrate SigmaPlot with the applications you have developed. It also provides an effective tool to customize or automate frequent tasks you want to perform. Automation uses objects to manipulate a program. Objects are the fundamental building block of macros; nearly all macro programs involve modifying objects. Every item in SigmaPlot-graphs, worksheets, axes, tick marks, reports, notebooks, etc.-can be represented by an object. SigmaPlot uses a VBAÆ-like macro language to access automation internally. For more information on recording and editing SigmaPlot macros, see About Macros.
This chapter contains the following topics:
ä “Opening SigmaPlot from Microsoft Word or Excel” on page 255 ä “SigmaPlot Objects and Collections” on page 256 ä “SigmaPlot Properties” on page 266 ä “SigmaPlot Methods” on page 275
Opening SigmaPlot from Microsoft Word or ExcelSigmaPlot Automation Reference
SigmaPlot Objects and Collections Application Object
' SigmaPlot Macro ' ' Dim SPApp as Object Set SPApp = CreateObject("SigmaPlot.Application.1") SPApp.Visible = True
SPApp.Application.Notebooks.Add End Sub
4. Choose Run/Run Sub/User Form to run the macro.
SigmaPlot appears with an empty worksheet and notebook window.
To open SigmaPlot from Word or Excel in the future:
1. Choose Tools/Macro/Macros to open the Macros dialog box.
2. Select SigmaPlot.
3. Click Run.
0
An object represents any type of identifiable item in SigmaPlot. Graphs, axes, notebooks, worksheets, and worksheet columns are all objects. A collection is an object that contains several other objects, usually of the same type; for example, all the items in a notebook are contained in a single collection object. Collections can have methods and properties that affect the all objects in the collection. Properties and methods are used to modify objects and collections of objects. To specify the properties and methods for an object that is part of a collection, you need to return that individual object from the collection first. For more information, refer to SigmaPlot Automation Help from the SigmaPlot Help menu.
An Application object represents the SigmaPlot application, within which all other objects are found. (Most other objects must exist inside higher-level objects. You access objects by applying properties and methods on these higher-level objects.) It is a "user-creatable" object, that is, outside programs can run SigmaPlot and access its properties directly, and will be registered in registry as SPW32.Application. The Application object properties and methods return or manipulate attributes of the SigmaPlot application and main window, and access the list of notebooks and from there all other objects.
256 SigmaPlot Objects and CollectionsNotebooks Collection Object
Notebook Object
NotebookItems Collection ObjectSigmaPlot Automation Reference
NotebookItem Object SectionItem Object
or collection index, and created using the NotebookItems Add method. The MacroItem object has an ItemType property and NotebookItems.Add method value of 0.
Represents the notebook item in the notebook window. You can use this object to rename the notebook item. The notebook item can always be reference with NotebookItems(0).
To use the NotebookItem Object
The NotebookItem object has most of the standard notebook item properties and methods, and created using the NotebookItems Add method. The NotebookItem object has an ItemType property and NotebookItems.Add method value of 7.
Represents the section folders within a SigmaPlot notebook. To use the SectionItem Object
The SectionItem object most of the standard notebook item properties and methods. A SectionItem is returned from the NotebookItems collection using the Item property or collection index, and created using the NotebookItems Add method. The SectionItem object has an ItemType property and NotebookItems.Add method value of 3.

<!-- EOF -->