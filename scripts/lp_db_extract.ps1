<#
    Wrapper for pe-sub-jobs/data/lp_db_extract.py — the standalone LP DB Export -> seed extract,
    for Windows PowerShell. Resolves its own location so it runs from any working directory.

    There are NO arguments: set the export to process by editing the EXPORT_FILE variable near the
    top of lp_db_extract.py, then run this wrapper.

        # from anywhere:
        powershell -ExecutionPolicy Bypass -File pe-sub-jobs\data\lp_db_extract.ps1

    If PowerShell blocks the script ("running scripts is disabled"), the -ExecutionPolicy Bypass
    above runs it without changing any machine setting. To allow .ps1 files for your user once:
        Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
#>
$ErrorActionPreference = 'Stop'

$ScriptDir = $PSScriptRoot
$PyScript  = Join-Path $ScriptDir 'lp_db_extract.py'

if (-not (Test-Path -LiteralPath $PyScript)) {
    Write-Error "cannot find $PyScript"
    exit 1
}

# Resolve a Python interpreter: prefer the 'py' launcher, then python / python3.
$PyExe  = $null
$PyArgs = @()
if (Get-Command py -ErrorAction SilentlyContinue) {
    $PyExe = 'py'; $PyArgs = @('-3')
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
    $PyExe = 'python'
} elseif (Get-Command python3 -ErrorAction SilentlyContinue) {
    $PyExe = 'python3'
} else {
    Write-Error "Python not found on PATH (tried 'py', 'python', 'python3')."
    exit 1
}

# openpyxl is the only third-party dependency.
& $PyExe @PyArgs -c 'import openpyxl' 2>$null
if ($LASTEXITCODE -ne 0) {
    $how = (@($PyExe) + $PyArgs) -join ' '
    Write-Error "the 'openpyxl' package is required. Install it with:`n  $how -m pip install openpyxl"
    exit 1
}

& $PyExe @PyArgs $PyScript
exit $LASTEXITCODE
