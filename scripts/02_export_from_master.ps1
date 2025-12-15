param(
  [Parameter(Mandatory=$true)]
  [string]$MasterSummaryFile,

  # Optional: when omitted, Python default ("対戦一覧_短縮") is used.
  [string]$SheetName = "",
  [switch]$WallHtml,
  [int]$WallCourtsPerPage = 3,
  [string]$HtmlPasscode = "",
  [switch]$NoMembers
)

$ErrorActionPreference = "Stop"

# Avoid mojibake on Windows PowerShell 5.1
$utf8 = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = $utf8
$OutputEncoding = $utf8

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptDir
Set-Location $repoRoot

$python = Join-Path $repoRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $python)) {
  throw "Python venv not found: $python`nRun setup: python -m venv .venv; .\.venv\Scripts\Activate.ps1; pip install -e ."
}

if (-not (Test-Path $MasterSummaryFile)) {
  $sample = (Get-ChildItem -File -Name "*.xlsm" -ErrorAction SilentlyContinue | Select-Object -First 10) -join "`n  - "
  throw "Master summary file not found: $MasterSummaryFile`nRepo root: $repoRoot`nExamples of .xlsm in this folder:`n  - $sample"
}

$exportArgs = @(
  "export-from-summary",
  "--input-file", $MasterSummaryFile,
  "--html-passcode", $HtmlPasscode
)
if (-not $NoMembers) {
  $exportArgs += "--include-members"
} else {
  $exportArgs += "--include-members=False"
}
if ($SheetName -and $SheetName.Trim().Length -gt 0) {
  $exportArgs += @("--sheet-name", $SheetName)
}
if ($WallHtml) {
  $exportArgs += @("--wall-html", "--wall-courts-per-page", "$WallCourtsPerPage")
}

& $python -m badminton_program.scheduler @exportArgs

$scoreArgs = @(
  "score-sheets-from-summary",
  "--input-file", $MasterSummaryFile
)
if ($SheetName -and $SheetName.Trim().Length -gt 0) {
  $scoreArgs += @("--sheet-name", $SheetName)
}

& $python -m badminton_program.scheduler @scoreArgs
