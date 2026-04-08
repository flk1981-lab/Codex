param(
    [Parameter(Position = 0)]
    [string]$Action = "check",

    [Parameter(Position = 1)]
    [string]$DocxPath = ".\input\working\thesis.docx"
)

$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Tool = Join-Path $Root "tools\thesis_tools.py"
$Terms = Join-Path $Root "config\terms.csv"
$Chapters = Join-Path $Root "output\chapters"
$Outline = Join-Path $Root "output\outline\outline.md"
$Stats = Join-Path $Root "output\reports\docx-stats.md"
$TermReport = Join-Path $Root "output\reports\term-check.md"
$Backups = Join-Path $Root "output\backups"

function Resolve-PythonExe {
    $localAppData = $env:LOCALAPPDATA
    if (-not $localAppData) {
        $localAppData = Join-Path $env:USERPROFILE "AppData\Local"
    }

    $candidates = @(
        (Join-Path $localAppData "Python\bin\python.exe"),
        (Join-Path $localAppData "Python\pythoncore-3.14-64\python.exe"),
        (Join-Path $localAppData "Programs\Python\Python314\python.exe"),
        (Join-Path $localAppData "Programs\Python\Python313\python.exe")
    )

    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path $candidate)) {
            return $candidate
        }
    }

    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python) {
        return $python.Source
    }

    throw "Python executable not found. Please confirm Python is installed and restart the terminal."
}

function Invoke-Tool {
    param([string[]]$Args)
    $pythonExe = Resolve-PythonExe
    & $pythonExe $Tool @Args
}

switch ($Action.ToLowerInvariant()) {
    "check" {
        Invoke-Tool @("env-check", "--root", $Root)
    }
    "stats" {
        Invoke-Tool @("docx-stats", $DocxPath, "--output", $Stats)
    }
    "outline" {
        Invoke-Tool @("outline", $DocxPath, "--output", $Outline)
    }
    "split" {
        Invoke-Tool @("split-docx", $DocxPath, "--outdir", $Chapters)
    }
    "terms" {
        Invoke-Tool @("term-check", $DocxPath, "--terms", $Terms, "--output", $TermReport)
    }
    "backup" {
        Invoke-Tool @("backup", $DocxPath, "--backup-dir", $Backups)
    }
    default {
        throw "Unsupported action: $Action"
    }
}
