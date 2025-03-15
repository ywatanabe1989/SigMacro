<!-- ---
!-- Timestamp: 2025-03-10 16:29:39
!-- Author: ywatanabe
!-- File: /home/ywatanabe/proj/SigMacro/PySigMacro/README.md
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

<!-- EOF -->