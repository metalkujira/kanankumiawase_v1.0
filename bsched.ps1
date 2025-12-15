param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Args
)

$ErrorActionPreference = 'Stop'

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$python = Join-Path $here '.venv\Scripts\python.exe'

if (-not (Test-Path $python)) {
    Write-Error "Python venv not found: $python`nRun: python -m venv .venv ; .\.venv\Scripts\Activate.ps1 ; pip install -r requirements.txt"
    exit 1
}

# Ensure local sources are used even if a global bsched is on PATH.
$env:PYTHONPATH = (Join-Path $here 'src')

& $python -m badminton_program.scheduler @Args
exit $LASTEXITCODE
