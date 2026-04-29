<#
.SYNOPSIS
  Daily launcher — opens WSL Ubuntu in the project dir and starts Claude Code.
#>

$ErrorActionPreference = 'Stop'

function Write-Info ($m) { Write-Host "[INFO]  $m" -ForegroundColor Cyan }
function Write-Err  ($m) { Write-Host "[ERROR] $m" -ForegroundColor Red }

try {
    Write-Info "Launching Claude Code in ~/knowledge-base ..."
    $null = wsl --status 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Err "WSL not available. Run deploy\01-setup-wsl.ps1 first."
        exit 1
    }
    wsl -d Ubuntu-24.04 --cd "~/knowledge-base" -- bash -lic "claude --dangerously-skip-permissions"
    if ($LASTEXITCODE -ne 0) {
        Write-Err "Claude Code exited with code $LASTEXITCODE."
        exit $LASTEXITCODE
    }
} catch {
    Write-Err $_.Exception.Message
    exit 1
}
