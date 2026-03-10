$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $projectRoot

$python = '.\.venv\Scripts\python.exe'
$chrome = 'C:\Program Files\Google\Chrome\Application\chrome.exe'

if (-not (Test-Path $python)) {
    throw "Python venv executable not found: $python"
}
if (-not (Test-Path $chrome)) {
    throw "Chrome not found: $chrome"
}

$ordersDir = 'docflow\objects\x5-ufa-e2_logistics_park\01_orders_and_appointments'
$outputDir = "$ordersDir\print_pdf_ready"
$files = @(
    "$ordersDir\20260310_ORDER_08_installers_general_v01.md",
    "$ordersDir\20260310_ORDER_09_two_foremen_assignment_v01.md",
    "$ordersDir\20260310_ORDER_10_tb_responsibility_grichushnikov_v01.md",
    "$ordersDir\20260310_ORDER_11_tb_responsibility_yakupov_v01.md",
    "$ordersDir\20260310_PERMIT_12_general_work_permit_2weeks_v01.md"
)

& $python scripts\export_orders_pdf.py --chrome-path $chrome --output-dir $outputDir @files

$printDir = Join-Path $projectRoot $outputDir
if (-not (Test-Path $printDir)) {
    throw "Print folder not found: $printDir"
}

Write-Host "PDF pack ready:" $printDir
Start-Process explorer.exe $printDir
