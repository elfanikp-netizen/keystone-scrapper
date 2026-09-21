param(
    [int]$limit = 0,
    [switch]$show,
    [switch]$fresh,
    [switch]$manualLogin,
    [string]$input = "parts.xlsx",
    [string]$column = "",
    [string]$output = "keystone_results.xlsx"
)

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$pythonCandidates = @(
    "c:\Users\administrator\AppData\Local\Python\pythoncore-3.14-64\python.exe",
    "python",
    "py"
)

$python = $null
foreach ($candidate in $pythonCandidates) {
    if ($candidate -eq "python" -or $candidate -eq "py") {
        try {
            $null = Get-Command $candidate -ErrorAction Stop
            $python = $candidate
            break
        }
        catch {
            continue
        }
    }
    elseif (Test-Path $candidate) {
        $python = $candidate
        break
    }
}

if (-not $python) {
    throw "Python was not found on this machine. Install Python 3.10+ and try again."
}

& $python -m pip install -r requirements.txt
& $python -m playwright install chromium

$argsList = @("keystone_crawler.py")
if ($show) { $argsList += "--show" }
if ($fresh) { $argsList += "--fresh" }
if ($manualLogin) { $argsList += "--manual-login" }
if ($limit -gt 0) { $argsList += "--limit"; $argsList += $limit }
if ($input -ne "parts.xlsx") { $argsList += "--input"; $argsList += $input }
if ($column) { $argsList += "--column"; $argsList += $column }
if ($output -ne "keystone_results.xlsx") { $argsList += "--output"; $argsList += $output }

& $python @argsList
