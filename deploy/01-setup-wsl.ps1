#Requires -RunAsAdministrator
<#
.SYNOPSIS
  Step 1 — Enable WSL2 and install Ubuntu 24.04. Reboot required after.
#>

$ErrorActionPreference = 'Stop'

function Write-Info  ($m) { Write-Host "[INFO]  $m" -ForegroundColor Cyan }
function Write-Ok    ($m) { Write-Host "[OK]    $m" -ForegroundColor Green }
function Write-Warn  ($m) { Write-Host "[WARN]  $m" -ForegroundColor Yellow }
function Write-Err   ($m) { Write-Host "[ERROR] $m" -ForegroundColor Red }

try {
    Write-Info "Knowledge Base — Step 1: WSL2 + Ubuntu 24.04 setup"
    Write-Info "==================================================="

    # Verify admin
    $isAdmin = ([Security.Principal.WindowsPrincipal] `
        [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
        [Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Err "This script must be run as Administrator. Right-click PowerShell -> Run as Administrator."
        exit 1
    }
    Write-Ok "Running with Administrator privileges."

    # Enable Windows features
    Write-Info "Enabling Windows feature: Microsoft-Windows-Subsystem-Linux"
    dism.exe /online /enable-feature /featurename:Microsoft-Windows-Subsystem-Linux /all /norestart | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "Failed to enable WSL feature (dism exit $LASTEXITCODE)." }
    Write-Ok "WSL feature enabled."

    Write-Info "Enabling Windows feature: VirtualMachinePlatform"
    dism.exe /online /enable-feature /featurename:VirtualMachinePlatform /all /norestart | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "Failed to enable VirtualMachinePlatform (dism exit $LASTEXITCODE)." }
    Write-Ok "Virtual Machine Platform enabled."

    # Install Ubuntu 24.04
    Write-Info "Installing Ubuntu-24.04 via WSL (this downloads ~500 MB)..."
    try {
        wsl --set-default-version 2 2>$null | Out-Null
        wsl --install -d Ubuntu-24.04 --no-launch
        if ($LASTEXITCODE -ne 0) {
            Write-Warn "wsl --install returned $LASTEXITCODE. The WSL feature may not be active until reboot — that's OK, Ubuntu will install on first run after reboot."
        } else {
            Write-Ok "Ubuntu-24.04 install command succeeded."
        }
    } catch {
        Write-Warn "wsl --install raised an exception: $($_.Exception.Message). This is usually fine before reboot."
    }

    Write-Host ""
    Write-Host "==================================================="  -ForegroundColor Green
    Write-Host "  STEP 1 COMPLETE."                                   -ForegroundColor Green
    Write-Host "  Restart your PC, then run: 02-setup-project.ps1"    -ForegroundColor Green
    Write-Host "==================================================="  -ForegroundColor Green
} catch {
    Write-Err $_.Exception.Message
    exit 1
}
