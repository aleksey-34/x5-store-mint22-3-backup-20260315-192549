param(
    [int]$Port = 8000,
    [switch]$Reload
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$workspaceRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
Set-Location $workspaceRoot

function Test-OllamaHealthy {
    try {
        $resp = Invoke-WebRequest -UseBasicParsing -Uri "http://127.0.0.1:11434/api/version" -TimeoutSec 2
        return $resp.StatusCode -eq 200
    }
    catch {
        return $false
    }
}

function Start-OllamaIfNeeded {
    if (Test-OllamaHealthy) {
        return
    }

    $ollama = Get-Command ollama.exe -ErrorAction SilentlyContinue
    if ($null -eq $ollama) {
        Write-Warning "Ollama executable not found in PATH. Continuing without local LLM startup."
        return
    }

    Start-Process -FilePath $ollama.Source -ArgumentList "serve" -WindowStyle Hidden | Out-Null

    for ($i = 0; $i -lt 10; $i++) {
        if (Test-OllamaHealthy) {
            return
        }
        Start-Sleep -Milliseconds 800
    }
}

Start-OllamaIfNeeded

$apiRunner = Join-Path $PSScriptRoot "run_api_clean.ps1"
if ($Reload) {
    & $apiRunner -Port $Port -Reload
}
else {
    & $apiRunner -Port $Port
}
exit $LASTEXITCODE
