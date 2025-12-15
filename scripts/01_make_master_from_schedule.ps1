param(
  [Parameter(Mandatory=$true)]
  [string]$ScheduleFile,

  [Parameter(Mandatory=$true)]
  [string]$SummaryFile,

  # Optional: when omitted, Python default ("集計表") is used.
  [string]$TargetSheet = "",
  [switch]$NoClearScores
)

$ErrorActionPreference = "Stop"

# Avoid mojibake on Windows PowerShell 5.1
$utf8 = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = $utf8
$OutputEncoding = $utf8

# Resolve repo root from this script location
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptDir

Set-Location $repoRoot

$python = Join-Path $repoRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $python)) {
  throw "Python venv not found: $python`nRun setup: python -m venv .venv; .\.venv\Scripts\Activate.ps1; pip install -e ."
}

if (-not (Test-Path $ScheduleFile)) {
  $sample = (Get-ChildItem -File -Name "*.xlsx" -ErrorAction SilentlyContinue | Select-Object -First 10) -join "`n  - "
  throw "Schedule file not found: $ScheduleFile`nRepo root: $repoRoot`nExamples of .xlsx in this folder:`n  - $sample"
}
if (-not (Test-Path $SummaryFile)) {
  $sample = (Get-ChildItem -File -Name "*.xlsm" -ErrorAction SilentlyContinue | Select-Object -First 10) -join "`n  - "
  throw "Summary file not found: $SummaryFile`nRepo root: $repoRoot`nExamples of .xlsm in this folder:`n  - $sample"
}

$clearFlag = "--clear-scores"
if ($NoClearScores) {
  $clearFlag = "--no-clear-scores"
}

$args = @(
  "fill-summary-grid-from-xlsx",
  "--schedule-file", $ScheduleFile,
  "--summary-file", $SummaryFile,
  $clearFlag
)
if ($TargetSheet -and $TargetSheet.Trim().Length -gt 0) {
  $args += @("--target-sheet", $TargetSheet)
}

& $python -m badminton_program.scheduler @args

Write-Host "" 
Write-Host "Next (Excel): open the generated *_filled_from_xlsx file and run the macro to refresh the short match list." -ForegroundColor Yellow
