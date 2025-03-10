# run_tests.ps1
# Timestamp: "2025-03-08 23:31:07 (ywatanabe)"

param(
    [string]$Mark,
    [switch]$Verbose,
    [switch]$Help
)

# Get script directory and set log path
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogFile = "$($MyInvocation.MyCommand.Path).log"

# Initialize log file
"" | Out-File -FilePath $LogFile -Encoding utf8 -Force

# Help function
if ($Help) {
    Write-Host "Usage: $($MyInvocation.MyCommand.Name) [-Mark <MARK>] [-Verbose] [-Help]"
    Write-Host ""
    Write-Host "Options:"
    Write-Host "  -Mark <MARK>     Run only tests with the specified marker"
    Write-Host "  -Verbose         Run tests in verbose mode"
    Write-Host "  -Help            Display this help message"
    Write-Host ""
    Write-Host "Example:"
    Write-Host "  .\$($MyInvocation.MyCommand.Name) -Mark 'not windows'"
    Write-Host "  .\$($MyInvocation.MyCommand.Name) -Verbose"
    exit 1
}

$VerboseArg = if ($Verbose) { "-v" } else { "" }

# Check if running in WSL
if ($env:WSL_DISTRO_NAME -or (Test-Path /proc/version -ErrorAction SilentlyContinue)) {
    if ((Get-Content /proc/version -ErrorAction SilentlyContinue) -match "Microsoft|WSL") {
        $Message = "Running in WSL environment"
        Write-Host $Message
        Add-Content -Path $LogFile -Value $Message
        
        # When in WSL, skip Windows-only tests by default if no mark specified
        if (-not $Mark) {
            $Mark = "not windows"
        }
    }
}

# Build pytest command
$PytestArgs = @("-m", "pytest")
if ($Mark) {
    $PytestArgs += "-m"
    $PytestArgs += "$Mark"
}
if ($Verbose) {
    $PytestArgs += "-v"
}

$Message = "Running pytest with args: $($PytestArgs -join ' ')"
Write-Host $Message
Add-Content -Path $LogFile -Value $Message

# Navigate to script directory and run pytest
Push-Location $ScriptDir
try {
    $Output = & python $PytestArgs 2>&1
    $ExitCode = $LASTEXITCODE
    
    # Log the output
    $Output | Out-File -FilePath $LogFile -Append
    # Display the output
    $Output | ForEach-Object { Write-Host $_ }
    
    exit $ExitCode
}
finally {
    Pop-Location
}
