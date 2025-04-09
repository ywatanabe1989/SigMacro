<!-- ---
!-- Timestamp: 2025-03-30 10:47:25
!-- Author: ywatanabe
!-- File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/README.md
!-- --- -->

# PySigMacro

A Python library for controlling SigmaPlot

## Install Python on Windows

https://www.python.org/ftp/python/3.13.2/python-3.13.2-amd64.exe

``` ps1
where.exe python
```

## Install PySigMacro via pip
``` ps1
cd C:\path\to\SigMacro\pysigmacro
python.exe -m pip install pip -U
python.exe -m pip install -r requirements.txt
python.exe -m pip install -e .
```

## (Optional) Environmental Variable

``` ps1
$env:SIGMACRO_JNB_PATH = "C:\Users\wyusu\Documents\SigMacro\SigMacro.JNB"
$env:SIGMACRO_TEMPLATES_DIR = "C:\Users\wyusu\Documents\SigMacro\Templates"
$env:SIGMAPLOT_BIN_PATH_WIN = "C:\Program Files (x86)\SigmaPlot\SPW16\Spw.exe"
# $env:SIGMAPLOT_BIN_PATH_WSL = "/mnt/c/Program Files (x86)/SigmaPlot/SPW16/Spw.exe" # Optional
```

## (Optional) Work on WSL

``` bash
## Directory
mkdir -p ~/proj/
ln -s /mnt/c/Users/<YOUR-USER-NAME>/path/to/SigMacro/PySigMacro ~/proj/PySigMacro

## Executables
mkdir -p ~/.win-bin
ln -s /mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe ~/.win-bin/powershell.exe
ln -s /mnt/c/Program Files (x86)/SigmaPlot/SPW12/Spw.exe ~/.win-bin/sigmaplot.exe
export PATH:$PATH:~/.win-bin

## Aliases
alias 'kill-sigmaplot'='powershell.exe -File "$(wslpath -w /home/<YOUR-USER-NAME>/win/program_files_x86/<YOUR-USER-NAME>/kill-sigmaplot.ps1)"'
alias 'python.exe'='powershell.exe python.exe'
alias 'ipython.exe'='powershell.exe ipython.exe --no-autoindent'
```

<!-- ## Environmental Variables
 !-- ```powershell
 !-- $env:SIGMACRO_PATH = "C:/Users/<YOUR-USER-NAME>/Documents/SigmaPlot/SPW12/SigMacro.JNB"
 !-- ``` -->

## Python (such as ipython.exe)
``` ps1
import os
import pysigmacro as psm
import pandas as pd
import numpy as np

# Parameters
PLOT_TYPE = "dev" # "line"
CLOSE_OTHERS = True
TGT_DIR = rf"C:\Users\{os.getlogin()}\Desktop"
TGT_PATH = psm.path.copy_template(PLOT_TYPE, TGT_DIR)
TGT_FILENAME = os.path.basename(TGT_PATH)

# Open a JNB notebook
sp = psm.con.open(lpath=TGT_PATH, close_others=True) # sp is SigmaPlot COM Object
# print(sp) # <BaseCOMWrapper for SigmaPlot 15 at SigmaPlot>
# print(sp.path) # "C:\Users\YOUR_LOGIN_NAME\Desktop\dev_20250326_193346.JNB"
notebooks = sp.Notebooks_obj
notebook = notebooks[notebooks.find_indices(TGT_FILENAME)[0]]
notebookitems = notebook.NotebookItems_obj
# # print(notebookitems.list)
# ['Notebook',
#  'section',
#  'worksheet',
#  'graph',
#  'SetLabelsMacro',
#  'SetFigureSizeMacro',
#  '_SetScalesMacro',
#  '_SetRangesMacro',
#  '_SetColorsMacro']

## Instanciates item objects
worksheet = notebookitems["worksheet"]
# print(worksheet) # <WorksheetItemWrapper for worksheet at SigmaPlot.Notebooks[4].NotebookItems[worksheet]>
graph = notebookitems["graph"]
# print(graph) # <GraphItemWrapper for graph at SigmaPlot.Notebooks[4].NotebookItems[graph]>
set_labels_macro = notebookitems["SetLabelsMacro"]
# print(set_labels_macro) # <MacroItemWrapper for SetLabelsMacro at SigmaPlot.Notebooks[4].NotebookItems[SetLabelsMacro]>

# Demo data
df = pd.DataFrame(
    columns=[ii for ii in range(30)], data=np.random.rand(100, 30)
)

worksheet.import_data(df)
set_labels_macro.run()
```

<!-- EOF -->