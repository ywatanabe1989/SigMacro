<!-- ---
!-- Timestamp: 2025-03-26 18:44:17
!-- Author: ywatanabe
!-- File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/README.md
!-- --- -->

## SigmaPlot Object Hierarchy

``` plaintext
**Application**
└── **Notebooks** (collection)
    └── **Notebook**
        └── **NotebookItems** (collection)
            ├── **NativeWorksheetItem**
            │   ├── **DataTableNamedDataRanges** (collection)
            │   │   └── **NamedDataRange**
            │   ├── Smoother
            │   ├── PlotEquation
            │   └── **GraphWizard**
            ├── ExcelItem
            │   ├── DataTableNamedDataRanges (collection)
            │   │   └── NamedDataRange
            │   ├── Smoother
            │   ├── PlotEquation
            │   └── **GraphWizard**
            ├── FitItem
            │   └── FitResults
            ├── TransformItem
            ├── ReportItem
            ├── **MacroItem**
            ├── **NotebookItem**
            ├── **SectionItem**
            └── **GraphItem**
                └── **Pages** (collection)
                    └── **GraphObjects (Page)** (collection)
                        ├── Text
                        ├── **Line**
                        ├── **Solid**
                        ├── **GraphObject**
                        ├── Group
                        ├── Smoother
                        ├── PlotEquation
                        └── **Graph**
                            ├── **Graph Objects (Axis)** (collection)
                            │   └── **Axis**
                            ├── **Line** (collection)
                            ├── Text (collection)
                            │   └── Text
                            ├── Group (AutoLegend)
                            │   ├── Solid
                            │   └── Text
                            ├── **Graph Objects (Plots)** (collection)
                            │   └── **Plot**
                            │       ├── Symbol
                            │       ├── **Line**
                            │       ├── **Solid**
                            │       └── Text
                            ├── GraphObjects (Tuple) (collection)
                            │   └── Tuple
                            ├── Graph Objects (DropLines) (collection)
                            │   └── Line
                            └── Graph Objects (Function) (collection)
                                ├── Function (Line)
                                └── Text
```

<!-- EOF -->