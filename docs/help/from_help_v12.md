<!-- ---
!-- Timestamp: 2025-03-12 07:29:57
!-- Author: ywatanabe
!-- File: /home/ywatanabe/proj/SigMacro/docs/from_help/from_help.md
!-- --- -->

## Parts of the Macro Programming Language 
The following topics list the parts of the macro programming language: 
 •  **Statements** are instructions to SigmaPlot to perform an action(s). Statements can consist of **keywords**, **operators**, **variables**, and **procedure calls**.  
 •  Keywords are terms that have special meaning in SigmaPlot. For example, the **Sub** and **End Sub** keywords mark the beginning and end of a macro. By default, keywords appears as blue text on color monitors. To find out more about a specific keyword in a macro, **select the keyword and press F1**. When you do this, a topic in the SigmaPlot on-line reference appears and presents information about the term.  
 •  You can add optional comments to describe a macro command or function, and how it interacts in the script. When the macro is running, comment lines are ignored. Indicate a comment by beginning a line with **an apostrophe**. Comments always must end the line they’re on. The next program line must go on a new line. By default, comment lines appear as green text. 


## Creating Custom Dialog Boxes
Design and customize your own dialog boxes using the UserDialog Editor. When you are designing and creating SigmaPlot macros, you can automatically create the necessary dialog box code and dialog monitor function code. Like the other automated coding features in SigmaPlot, the code may require further customizing. 
To create a custom dialog box: 
 1. In the Macro Window, place the insertion point where you want to put the code for the dialog box. For more information, see Editing Macros.  
 2. On the Macro Window toolbar click the User Dialog button. The blank grid in the User Dialog Editor appears. 
 3. On the left hand side of the User Dialog Editor there is a Toolbox. You can select a tool, such as a button or check boxes, from the Toolbox. The cursor changes to a cross when you move it over the grid.  
 4. To place a tool on the dialog box, click a position on the grid. A default tool will be added to the dialog grid.  
 5. Resize the dialog box by dragging the handles on the sides and the corners.  
 6. Right-click any of the controls that you have placed on the dialog surface (after selecting the control) and enter a name for the control.  
 7. Right-click the dialog form (with no control selected) and enter a name for the dialog monitor function in the DialogFunc field.  
 8. To finish, click OK. The code for the dialog box with controls will be written to the Macro Window. 
Finally, and in most cases, you must edit the code for dialog box monitor function to define the specific behavior of the elements in your dialog box.  

## Using the Object Browser 
The Object Browser displays all SigmaPlot object classes. The methods and properties associated with each SigmaPlot macro object class are listed. A short description of each object appears in the dialog box as you select them from the list.
To view the Object Browser, the Macro Window must first be in view. For more information, see Creating Macros Using the Macro Language. 
To open the Object Browser:
 1. On the Macro Window toolbar, click the Object Browser button.  
 2. Use Paste to insert generic code based on your selection into a macro. 
Tip
Press F1 at any time for full details on using the Object Browser. 

## Using the Add Procedure Dialog Box 
Organizing your code in procedures makes it easier to manage and reuse. SigmaPlot macros, like Visual Basic programs, must have at least one procedure (the main subroutine) and often they have several. The main procedure may contain only a few statements, aside from calling subroutines that do the work. You add procedures using the Add Procedure dialog box.
To add a procedure:
 1. On the Macro Window toolbar, click the New Procedure button. 
 2. In the Add Procedure dialog box, define a **sub**, **function**, or **property** using the **Name**, **Type**, and **Scope** boxes. 
 3. Click OK to paste the code for a new procedure. The new procedure appears at the bottom of the macro.
Tip
For full details on using the Add Procedure dialog box, press F1 from anywhere in the Macro Window. 

## About User-Defined Functions
A **user-defined function** is a combination of **math expressions** and **Basic code**. The function always requires **input data values** and always **returns a value**. **You supply the function with a value**; it performs calculations on the values and returns a new value as the answer. Functions can work with **text**, **dates**, and **codes**, not just numbers. A user-defined function is similar to a macro but there are differences. Some of the differences are listed in the following table.
| Recorded Macro                                                                                     | User-Defined Functions                                                                   |
| Performs a SigmaPlot action, such as creating a new chart. Macros change the state of the program. | Returns a value; cannot perform actions. Functions return answers based on input values. |
| Can be recorded.                                                                                   | Must be created in Macro code.                                                           |
| Are enclosed in **the Sub and End Sub** keywords.                                                  | Are enclosed in the keywords **Function and End Function**.                              |

## Creating User-Defined Functions 
A user-defined function is like any of the built-in SigmaPlot function. Because you create the user-defined function, however, you have control over exactly what it does. A single user-defined function can replace database and spreadsheet data manipulation with a single program that you call from inside SigmaPlot. It is a lot easier to remember a single program than it is to remember several spreadsheet macros. 

## Using the Debug Window
The Debug Window contains a group of features that are helpful when you are trying to locate and resolve errors in your macro code. The debugging tools in SigmaPlot will be familiar if you have used one of the modern visual programming languages or Microsoft Visual Basic for Applications. Essentially, the Debug Window gives you incremental control over the execution of your program so that you can sleuth errors in your programs. **The Debug Window also gives you a precise way to determine the contents of your variables.** Again, a series of buttons is used to select the operation mode of the Debug Window. 

## Debug Window Tabs 
The output from the Debug Window is organized in four tabs that allow you to type in statements, observe program execution responses, and iteratively modify your code using this feedback. If you have never used a debugging tool and are new to programming, it would be a good idea to supplement the following description with further study. 

## Immediate Tab 
The Immediate Tab lets you evaluate an expression, assign a specific value to a variable or call a subroutine and evaluate the results. Trace mode prints the code in the tab when the macro is running. 
 •  Type "?expr" and press Enter to show the value of "expr".  
 •  Type "var = expr" and press Enter to change the value of "var".  
 •  Type "set var = expr" and press Enter to change the reference of "var" for object vars.  
 •  Type "subname args" and press Enter to call a subroutine or built-in expression "subname" with arguments "args".  
 •  Type "trace" and press Enter to toggle trace mode. Trace mode prints each statement in the Immediate Tab when a macro is running.  
 
## Watch Tab 
The Watch Tab lists variables, functions, and expressions that are calculated during execution of the program. 
 •  Each time program execution pauses, the value of each line in the window is updated.  
 •  The expression to the left of the "->" may be edited.  
 •  Pressing Enter updates all the values immediately.  
 •  Pressing Ctrl+Y deletes the line.  
 
## Stack Tab 
The output from the Stack Tab lists the program lines that called the current statement. This is a macro command audit and is helpful to determine the order of statements in you program.
 •  The first line is the current statement. The second line is the one that called the first, and so on.  
 •  Clicking a line brings that macro into a sheet and highlights the line in the edit window. 

## Streamlining Procedures with Macros
Use SigmaPlot macros to help streamline your workflow. For example, you can create macros in Microsoft Word or Excel that allow you to open SigmaPlot from within either application. You can place macros that you create yourself on the main menu. You can even run a SigmaPlot macro by specifying its path in your command prompt without ever having to open SigmaPlot. Examples of these macro applications appear in the following topics. For more information, see Using SigmaPlot’s Macros. 

## Opening SigmaPlot from Microsoft Word or Excel
You can create a macro in either Microsoft Word or Microsoft Excel that can open SigmaPlot directly from either application.
To create this macro:
 1. In either Microsoft Word or Excel, click the Microsoft Office Button, and then click Excel Options. 
 2. In the Popular category, under Top options for working with Excel, select the Show Developer tab in the Ribboncheck box, and then click OK. 
 3. Click the Developer tab, and then in the Codegroup, click Visual Basic. 
 4. Type (or copy and paste): 
Sub SigmaPlot()
’
SigmaPlot Objects and Collections
’
SigmaPlot Macro
’


’

``` vba
Dim SPApp as Object
Set SPApp = CreateObject("SigmaPlot.Application.1")
SPApp.Visible = True
SPApp.Application.Notebooks.Add 
End Sub 
```

4.
 
 5. Click Run. SigmaPlot appears with an empty worksheet and notebook window.  
 6. To run the macro, in Excel, on the Developer tab, in the Code group, click Macros. 
 7. Click Run. SigmaPlot appears with an empty worksheet and notebook window.  

## Running SigmaPlot Macros from the Command Prompt
You can run SigmaPlot macros directly from your command prompt, saving you valuable time. Suppose you need to produce the same graph report of a data set week after week. Rather than going through the trouble of starting up SigmaPlot, opening a file, and then running a macro, you can run the entire macro from a run command on the Start menu instead.
 1. In your command prompt type: **c:\spw "filename" /runmacro:"macroname"**.
For example, if you want to run a macro you created called "ErrorBars", and it is stored in a notebook file called "MyNotebook.jnb", you would type **c:\spw MyNotebook.jnb\runmacro:ErrorBar**.
Tip
**You can also create a batch file or script that runs SigmaPlot from the DOS command prompt** as part of the batch file’s set of operations. 

## SigmaPlot’s Macros
SigmaPlot’s available macros are:
 •  Area Below Curves. Integrates under curves using the trapezoidal rule. For more information, see Area Below Curves.  
 •  Batch Process Excel Files. Imports data from multiple Excel Files into individual SigmaPlot worksheets, then plots and curve fits the imported data automatically. For more information, see Batch Process Excel Files .  
 •  Bland Altman Analysis. A technique for comparing two methods. For more information, see Bland-Altman Analysis.  
 •  Border Plots. Draws a histogram or box plot along the top and right axes of a scatter plot. For more information, see Border Plots.  
 •  By Group Data Split. Splits data contained in one column into groups of data sorted into multiple data columns within one SigmaPlot worksheet. For more information, see By Group Data Split.  
 •  Color Transition Values. Creates a column of colors changing smoothly in intensity as the data changes from its minimum value to its maximum value. For more information, see Color Transition Values.  
 •  Compute 1st Derivative. Computes a numerical first derivative of a pair of data columns. For more information, see Compute 1st Derivative.  
 •  Dot Density Plot. Creates a display of symbols arranged to show their spatial density.  
 •  Frequency Plot. Creates frequency plots with mean bars for multiple data columns. For more information, see Frequency Plot.  
 •  F-test Comparison of Curves. Compares the fits of two equations to determine if the more complicated equation provides a significantly better fit. For more information, see F Test Comparison of Curves.  
 •  Gaussian Cumulative Distribution. Returns the results of a Gaussian Cumulative Distribution function (CDF) for a single column of data, and optionally plots the results with a probability Y axis scale. For more information, see Gaussian Cumulative Distribution.  
 •  Insert Graphs into Microsoft Word. Inserts a SigmaPlot graph into an open Microsoft Word document. For more information, see Insert Graphs into Microsoft Word.  
 •  Label Symbols. Labels a plot with symbols or text from a specified column. For more information, see Label Symbols .  
 •  Merge Columns. Merges two separate worksheet columns into one single text column. For more information, see Merge Columns .  
 •  Paste to PowerPoint Slide. Creates PowerPoint slides from selected SigmaPlot graphs. For more information, see Paste to PowerPoint Slide.  
 •  Piper Plots. Creates a Piper Plot. For more information, see Piper Plots.  
 •  Plotting Polar and Parametric Equations. Creates curves in either Cartesian or polar coordinate systems. For more information, see Plotting Polar and Parametric Equations.  
 •  Power Spectral Density. Computes the power spectral density (psd) for a data column. For more information, see Power Spectral Density.  
 •  Quick Re-Plot. Re-assigns the columns that are plotted for the current curve in the current two- or three-dimensional plot. For more information, see Quick Re-Plot.  
 •  Rank and Percentile. Computes ranks and cumulative percentages for a specified data column. For more information, see Rank and Percentile.  
 •  ROC Curve Analysis. For more information, see ROC Curve Analysis.  
 •  Standard Curves. Creates a standard curve by fitting one of five functions to instrument data. Will also generate X-from-Y and ECnn-from-EC% values. For more information, see Standard Curve.  
 •  Survival Curve. Computes and graphs a Kaplan-Meier survival curve using the SurvlMod transform. For more information, see Survival Curve.  
 •  Vector Plot. Uses the vector transform to plot X,Y, angle and magnitude data as vectors with arrowheads. For more information, see Vector Plot.  

## SigmaPlot Automation Reference
OLE Automation is a technology that lets other applications, development tools, and macro languages use a program. Using SigmaPlot Automation, you can integrate SigmaPlot with the applications you have developed. Automation can also be an effective tool to customize or automate frequent tasks you want to perform. 
Automation uses **objects** to manipulate a program. For more information, see About Objects and Collections. Objects are the fundamental building block of macros; nearly all macro programs involve modifying objects. Every item in SigmaPlot - **graphs, worksheets, axes, tick marks, reports, notebooks, and so on - can be represented by an object**.

## About Objects and Collections
**An object** represents any type of identifiable item in SigmaPlot. **Graphs**, **axes**, **notebooks**, **worksheets**, and **worksheet columns** are all objects. 
A **collection** is **an object that contains several other objects, usually of the same type**; for example, **all the items in a notebook are contained in a single collection object**. Collections can have **methods and properties that affect the all objects in the collection**.
Use properties and methods to modify objects and collections of objects. **To specify the properties and methods for an object that is part of a collection, you need to return that individual object from the collection first**. 

## About Properties
A **property** is **a setting or other attribute of an object**. Think of a property as **an "adjective."** **For example, properties of a graph include the size, location, type and style of plot, and the data that is plotted.** 
To change the settings of an object, change the properties settings. **Properties are also used to access the objects that are below the current object in the hierarchy.**
**To change a property setting, type the object reference followed with a period, then type the property name, an equal sign (=), and the property value.**

Example
**Set Notebook.Title = "My Notebook"** Sets the name of the referenced SigmaPlot notebook to "My Notebook". 
Note that **some properties cannot be set, and only retrieved**. The Help topic for each property indicates whether you can both set and retrieve that property (read-write), only retrieve the property (read-only), or only set the property (write-only). 
**You can get information about an object by returning the values of its properties.**
Example
**Set CurrentDoc = ActiveDocument.NotebookItems(3)**
**The fourth item in the current notebook (specified by ActiveDocument) is assigned to the variable CurrentDoc (item counts start with 0).**

## About Methods
A Method is an action that can be performed on or by an object. Think of methods as verbs. For example, the **WorksheetEditItem object** has Copy and Clear methods. Methods can have parameters that specify the action (adverbs). 

Example
**Notebooks(0).NotebookItems(2).Close(True)**

## Returning Objects
In order to work with an object, **you must be able to define the specific object by returning it**. **In general, most objects are returned using a property of the object above it in the object tree.**

## Returning Objects from Collections
Other objects are returned by specifying a single object from a collection. Once you define the collection, **you can return a specific object by using an index value (as you would with an array)**. You can use either the Item method shared by all collections, or use the index directly. **The index can be the item name or a number**. For example: 
**Set Worksheet = Notebooks("My Notebook").NotebookItems.Item(2)**
The collection index value returns the notebook "My Notebook" from the Notebooks collection, then the Item property and index number returns the third item from the **NotebookItems collection** as the variable Worksheet. 
The **Notebooks collection** contains **a list of all the open notebooks** in SigmaPlot, and the **NotebookItems collection contains all items in the specified notebook**. 

## Defining Variables
You can also return and use objects by defining the object to be a variable, generally using the **Dim (dimension) statement**. 
**Although you can implicitly declare variables just by using the variable for the first time, **you can avoid bugs caused by typos using Option Explicit**. For example, the script: 
**Option Explicit**

``` vba
Sub Main 
Dim ItemCount
Dim SPWorksheets$()
ItemCount = ActiveDocument.NotebookItems.Count 
ReDim SPWorksheets$(ItemCount) 
Dim SPItems 
Set SPItems = ActiveDocument.NotebookItems 
Dim Index 
Index = 0 
Dim Item 
For Each Item In SPItems 
If SPItems(Index).ItemType = 1 Then 	
      SPWorksheets$(Index) = SPItems(Index).Name 
End If 
Index = Index + 1 
Next Item  

Begin Dialog UserDialog 320,119,"Worksheets in Active Notebook" ’ %GRID:10,7,1,1 
      OKButton 210,14,90,21 	
      ListBox 20,14,170,91,SPWorksheets(),.ListBox1 
End Dialog 
Dim dlg As UserDialog 
Dialog dlg 
End SubUses 
```

the Dim (Dimension) statement to define several variables, and uses the **Set instruction** to **define a declared variable as an object**.
Related Topics

<!-- ## Variable Definitions
 !-- Variable definitions use the form: 
 !-- **variable = range**
 !-- You can use any valid variable name, but short, single letter names are recommended for the sake of simplicity (for example, x and y). **The range can either be the column number for the data associated with each variable, or a manually entered range.**
 !-- Most typically, **the range is data read from a worksheet**. The curve fitter uses SigmaPlot’s transform language, so the notation for a column of data is: 
 !-- 
 !-- ``` vba
 !-- col(column,top,bottom)
 !-- ```
 !-- 
 !-- **The column argument determines the column number or title. To use a column title for the column argument, enclose the column title in quotation marks**. **The top and bottom arguments specify the first and last row numbers and can be omitted.** The default row numbers are 1 and the end of the column, respectively. If both are omitted, the entire column is used. For example, to define the variable x to be column 1, enter:
 !-- 
 !-- ``` vba
 !-- x = col(1) 
 !-- ```
 !-- 
 !-- Data may also be entered directly in the variables section. For example, you can define y and z variables by entering:
 !-- 
 !-- ``` vba
 !-- y = {1,2,4,8,16,32,64}
 !-- z = data(1,100)
 !-- ```
 !-- 
 !-- This method can have some advantages. For example, in the example above the data function was used to automatically generate z values of 1 through 100, which is simpler than typing the numbers into the worksheet. 
 !-- 
 !-- ## Iterations
 !-- Setting the number of iterations, or the maximum number of repeated regression attempts, is useful if you do not want to regression to proceed beyond a certain number of iterations, or if the regression exceeds the default number of iterations. 
 !-- The default iteration value is 200. To change the number of iterations, simply enter the maximum number of iterations in the Iterations edit box. 
 !-- 
 !-- ## Evaluating Parameter Values Using 0 Iterations
 !-- Iterations must be non-negative. However, setting Iterations to 0 causes no iterations occur; instead, the regression evaluates the function at all values of the independent variables using the parameter values entered under the Initial Parameters section and returns the results.
 !-- **If you are trying to evaluate the effectiveness of automatic parameter estimation function, setting Iterations to 0 allows you to view what initial parameter values were computed by your algorithms.**
 !-- Using zero iterations can be very useful for evaluating the effect of changes in parameter values. For example, once you have determined the parameters using the regression, you can enter these values plus or minus a percentage, run the regression with zero iterations, then graph the function results to view the effect of the parameter changes.  -->

<!-- EOF -->