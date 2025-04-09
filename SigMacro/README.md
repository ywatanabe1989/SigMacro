<!-- ---
!-- Timestamp: 2025-03-10 16:01:02
!-- Author: ywatanabe
!-- File: /home/ywatanabe/proj/SigMacro/SigMacro/README.md
!-- --- -->

# SigMacro Installation

## Automatic Configuration
  1. Extract the `SigmaPlotConfig - Extract and Place it under Documents.zip`
  2. Copy the `SigmaPlot` folder to `C:\Users\<YOUR_USER_NAME>\Documents\SigmaPlot\SPW12\`

## Manual Configuration
  1. Copy (or create shortcuts of) `Confusion_Matrix.JNB` and `SigMacro.JNB` to user configuration directory (`C:\Users\<YOUR_USER_NAME>\Documents\SigmaPlot\SPW12\`)
    
  2. Configurations

     Right Click the Icon at the top left corner-> Options

     - Macro Tab
       - Macro library
         - Update from `C:\Users\<YOUR_USER_NAME>\Documents\SigmaPlot\SPW12\SigmaPlot Macro Library.jnb` to `C:\Users\<YOUR_USER_NAME>\Documents\SigmaPlot\SPW12\SigMacro.JNB`
     - Page Tab
       - Page Units mm
     - General Tab
       - Uncheck Novice prompting
       - Uncheck Startup Screen
       - Template file
         - C:\Users\<YOUR_USER_NAME>\Documents\SigmaPlot\SPW12\TEMPLATE.JNT
       - Layout file
         - C:\Users\<YOUR_USER_NAME>\Documents\SigmaPlot\SPW12\Layout.JNT
       - Gallery file
         - C:\Users\<YOUR_USER_NAME>\Documents\SigmaPlot\SPW12\Gallery.jgg
       - Author
         - YOUR-NAME-OR-EMAIL
       - Macro
         - Macro library
           - Default
             - C:\Users\<YOUR_USER_NAME>\Documents\SigmaPlot\SPW12\SigmaPlot Macro Library.jnb
       - Customize Tool Bar
         - Right click
         - Choose commands from:
           - ToolBox
             - Enzyme Kinetics Wizard -> Remove

<!-- EOF -->