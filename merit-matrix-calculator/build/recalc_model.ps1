# Recalculates the Excel model with the live engine and verifies the result.
#
# openpyxl only writes formulas. What they actually compute is shown only by
# real Excel. Without this step the file holds no computed value at all, and
# "agrees to the kopeck" is backed by nothing.
#
# It does three things:
#   1. CalculateFullRebuild + Save - values appear in the workbook;
#   2. looks for formula errors (#REF!, #VALUE!, #DIV/0! and the rest);
#   3. runs check_model.py: totals and every employee row against reference_calc.py.

$ErrorActionPreference = "Stop"

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$book = Join-Path (Split-Path $here -Parent) "merit-model.xlsx"
if (-not (Test-Path $book)) { throw "No workbook at $book - run python build_model.py first" }

Write-Host "Starting Excel and rebuilding all formulas..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$errors = @()
try {
    $wb = $excel.Workbooks.Open($book)
    $excel.CalculateFullRebuild()
    $wb.Save()
    foreach ($ws in $wb.Worksheets) {
        try {
            # SpecialCells(-4123, 16) - cells whose formulas evaluate to an error
            $bad = $ws.UsedRange.SpecialCells(-4123, 16)
            if ($bad) {
                foreach ($cell in $bad) {
                    $errors += ("{0}!{1} = {2}" -f $ws.Index, $cell.Address(0,0), $cell.Text)
                }
            }
        } catch {
            # SpecialCells throws when there are no such cells - that is normal
        }
    }
    $wb.Close($true)
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

if ($errors.Count -gt 0) {
    Write-Host "FORMULA ERRORS (sheet number!cell):" -ForegroundColor Red
    $errors | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
}

$env:PYTHONIOENCODING = "utf-8"
Push-Location $here
try {
    python check_model.py
    $checkFailed = $LASTEXITCODE -ne 0
} finally {
    Pop-Location
}

if ($checkFailed -or $errors.Count -gt 0) {
    Write-Host "CHECK FAILED." -ForegroundColor Red
    exit 1
}
Write-Host "Workbook recalculated, no formula errors, totals and rows agree with reference_calc.py." -ForegroundColor Green
