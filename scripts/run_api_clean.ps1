param(
    [int]$Port = 8000,
    [switch]$Reload
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$workspaceRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
Set-Location $workspaceRoot

function Stop-UvicornProcesses {
    param([int]$TargetPort)

    $uvicornProcs = Get-CimInstance Win32_Process | Where-Object {
        $_.Name -match "^python(w)?\.exe$" -and
        $_.CommandLine -match "uvicorn" -and
        $_.CommandLine -match "app.main:app"
    }

    foreach ($proc in $uvicornProcs) {
        Stop-Process -Id $proc.ProcessId -Force -ErrorAction SilentlyContinue
    }

    $listeners = Get-NetTCPConnection -LocalPort $TargetPort -State Listen -ErrorAction SilentlyContinue
    if ($listeners) {
        $listenerPids = $listeners | Select-Object -ExpandProperty OwningProcess -Unique
        foreach ($listenerPid in $listenerPids) {
            Stop-Process -Id $listenerPid -Force -ErrorAction SilentlyContinue
        }
    }
}

Stop-UvicornProcesses -TargetPort $Port
Start-Sleep -Milliseconds 500

$pythonExe = Join-Path $workspaceRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $pythonExe)) {
    throw "Python executable not found: $pythonExe"
}

$uvicornArgs = @(
    "-m",
    "uvicorn",
    "app.main:app",
    "--host",
    "127.0.0.1",
    "--port",
    "$Port"
)

if ($Reload) {
    $uvicornArgs += "--reload"
}

& $pythonExe @uvicornArgs
exit $LASTEXITCODE
