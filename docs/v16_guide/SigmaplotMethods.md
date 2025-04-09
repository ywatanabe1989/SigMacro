<!-- ---
!-- Timestamp: 2025-03-09 10:36:18
!-- Author: ywatanabe
!-- File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-Pysigmacro/pysigmacro/docs/SigmaplotMethods.md
!-- --- -->

SigmaPlot MethodsSigmaPlot Automation Reference
Activate Method Makes the specified notebook the object specified by the ActiveDocument property.
Add Method
The Add method is used in collections to add a new item to the collection. The parameters depend on the collection type:
Collection Value Parameters Notebooks None NotebookItems 1 CT_WORKSHEET 2 CT_GRAPHICPAGE 2CT_FOLDER 4 CT_STATTEST 5CT_REPORT 6 CT_FIT 7 CT_NOTEBOOK 8CT_EXCELWORKSHEET 9 CT_TRANSFORMTransformItem 10 Graph Objects 2 GPT_GRAPH, more... 3 GPT_PLOT, more... 4 GPT_AXIS, more... 5 GPT_TEXT, more... 6 GPT_LINE, more... 7 GPT_SYMBOL, more... 8 GPT_SOLID, more... 9 GPT_TUPLE, more... 10 GPT_FUNCTION, more... 11 GPT_EXTERNAL, more... 12 GPT_BAG, more...
NamedRanges Name string, Left long, Top long, Width long, Height long, NamedRange
276 SigmaPlot MethodsAddVariable Expression Method
ApplyPageTemplate Method
AddWizardAxis Method
SigmaPlot Automation Reference
The GraphObjects collection uses the CreateGraphFromTemplate and CreateWizardGraph methods to create new GraphObject objects.
Allows the substitution of any transform variable with a value.
Overwrites the current GraphItem using a new page template specified by the template name. Optionally, you can specify the notebook file to use as the source of the template page. If no template file is specified, the default template notebook is used, as returned by the Template property.
Adds an additional axis to the current graph and plot on the specified GraphItem object, using the AddWizardAxis options. If there is only one plot for the current graph, SigmaPlot will return an error. Use the following parameters to specify the type of scale, the dimension, and the position for the new axis:
Scale TYPE
SAA_TYPE_LINEAR
SAA_TYPE_COMMON (Base 10)
SAA_TYPE_LOG (Base e)
SAA_TYPE_PROBABILITY
SAA_TYPE_PROBIT
SAA_TYPE_LOGIT
Dimension DIM_X 1 The X dimension DIM_Y 2 The Y dimension DIM_Z 3 The Z dimension (if applicable) Position AxisPosRightNormal 0 AxisPosRightOffset 1 AxisPosTopNormal 2 AxisPosTopOffset 3SigmaPlot Automation Reference
SetObjectCurrent Method
SetSelectedObjects Attribute Method
TransposePaste Method
Undo Method
ä Use the SigmaPlot menus (e.g. “Select Graph”). ä Use the SetObjectCurrent method. If the specified GraphItem is not open or there is no current object of the appropriate type on the page, the method will fail. Sets the specified object to the “current” object for the purpose of the “SetCurrentObjectAttribute” command. It the specified GraphItem is not open, the method will fail. Changes the attribute specified by “Attribute” for all the selected objects on the graphics page. Select graphics page objects using one of the following two techniques: ä Click the object with the mouse. ä Use the SelectObject method.
Pastes the data in the clipboard into the worksheet, transposing the row and column indices of the data such that rows and columns are swapped. If there is nothing in the clipboard or the data is not of the right type, nothing will happen.
Undoes the last performed action for the specified object. If undo has been disabled in SigmaPlot for either the worksheet or page, this method has no effect.

<!-- EOF -->