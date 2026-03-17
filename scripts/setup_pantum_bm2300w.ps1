param(
    [string]$DriverInstallerPath = "",
    [switch]$InstallTesseract
)

$ErrorActionPreference = "Stop"

Write-Host "== Pantum BM2300W setup =="

Write-Host "[1/5] Ensuring WIA (stisvc) service is enabled"
Set-Service -Name stisvc -StartupType Automatic
Start-Service -Name stisvc
Get-Service -Name stisvc | Select-Object Name, Status, StartType | Format-Table -AutoSize

Write-Host "[2/5] Checking connected Pantum devices"
$imgDevices = Get-PnpDevice -Class Image | Where-Object { $_.FriendlyName -match "Pantum|BM2300" }
$printDevices = Get-PnpDevice -Class Printer | Where-Object { $_.FriendlyName -match "Pantum|BM2300" }

if ($imgDevices) {
    Write-Host "Scanner device found:" -ForegroundColor Green
    $imgDevices | Select-Object Status, Class, FriendlyName, InstanceId | Format-Table -AutoSize
}
else {
    Write-Host "Scanner device in class Image not found" -ForegroundColor Yellow
}

if ($printDevices) {
    Write-Host "Printer device found:" -ForegroundColor Green
    $printDevices | Select-Object Status, Class, FriendlyName, InstanceId | Format-Table -AutoSize
}

Write-Host "[3/5] Optional driver installer"
if ($DriverInstallerPath) {
    if (Test-Path $DriverInstallerPath) {
        Write-Host "Launching driver installer: $DriverInstallerPath"
        Start-Process -FilePath $DriverInstallerPath -Verb RunAs
    }
    else {
        Write-Host "Driver installer path does not exist: $DriverInstallerPath" -ForegroundColor Yellow
    }
}
else {
    Write-Host "No DriverInstallerPath provided; skipping installer launch"
}

Write-Host "[4/5] Optional Tesseract OCR installation"
if ($InstallTesseract) {
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $winget) {
        Write-Host "winget not found. Install Tesseract manually." -ForegroundColor Yellow
    }
    else {
        winget install -e --id UB-Mannheim.TesseractOCR --accept-package-agreements --accept-source-agreements
    }
}
else {
    Write-Host "InstallTesseract switch not set; skipping OCR engine installation"
}

Write-Host "[5/5] Final checks"
where.exe tesseract
& .\.venv\Scripts\python.exe scripts/scanner_control.py diagnose
& .\.venv\Scripts\python.exe scripts/scanner_control.py list

Write-Host "Setup completed. If scanner is listed, run scanner_control.py scan-to-inbox next."
