param(
    [string] $TesseractCmd = $env:TESSERACT_CMD,
    [string] $Tessdata = $env:TESSDATA_PREFIX,
    [int] $Port = 80
)

$ErrorActionPreference = "Stop"
$python = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
if (!(Test-Path -LiteralPath $python)) {
    throw "Missing .venv. Follow README.md to install the Python libraries."
}
if ($TesseractCmd) { $env:TESSERACT_CMD = $TesseractCmd }
if ($Tessdata) { $env:TESSDATA_PREFIX = $Tessdata }
$env:PYTHONDONTWRITEBYTECODE = "1"
& $python (Join-Path $PSScriptRoot "app.py") --port $Port
exit $LASTEXITCODE
