<!-- ---
!-- Timestamp: 2025-03-09 10:35:39
!-- Author: ywatanabe
!-- File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/docs/SigmaplotProperties.md
!-- --- -->

SigmaPlot Properties 0
Application Property Σ
A property is a setting or other attribute of an objectóthink of a property as an "adjective." For example, properties of a graph include the size, location, type and style of plot, and the data that is plotted. To change the settings of an object, you change the properties settings. Properties are also used to access the objects that are below the current object in the hierarchy. To change a property setting, type the object reference followed with a period, then type the property name, an equal sign (=), and the property value. For more information, refer to SigmaPlot Automation Help from the SigmaPlot Help menu.
Used without an object qualifier, this property returns an Application object that represents the SigmaPlot application. Used with an object qualifier, this property returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return that object's application). Use the CreateObject and GetObject functions give you access to an Automation object.
266 SigmaPlot PropertiesSigmaPlot Automation Reference
Returns the Notebook object used as the template source file. The template is used for new page creation. To create a graph page using a template file, use the ApplyPageTemplate method.
Specifies the text for the report, transform or macro code. The text is unformatted, plain text. Use the vbCrLf string data constant to insert a carriage-return and linefeed string. Transforms:   To change the value of a transform variable, use the AddVariableExpression method. Run transforms using the Execute method.
Returns the tick label Text objects for the specified Axis object.
SigmaPlot Methods
A Notebook object property. Sets the Name of the NotebookItem object of the Notebook file, and the Title field under the Summary tab of the Windows 95/98 file Properties dialog box. Does not affect the file name; to change the file name, use either the Name or FullName property.
Sets or returns the top coordinate of the application window or specified notebook document window.
A property common to the Application, Notebook, and NotebookItems document objects. Sets or returns a Boolean indicating whether or not the application or specified document window is visible. Do not set the Application property to False from within SigmaPlot or you will lose access to the application. Note that hidden document windows will still appear in the notebook window tree. Setting Visible=False for a notebook object hides all document windows for the notebook as well.
Sets or returns the width of the application window or specified notebook document window.

<!-- EOF -->