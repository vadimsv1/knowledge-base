#Requires -RunAsAdministrator
<#
.SYNOPSIS
  Step 2 — Stage the project files into C:\dev\knowledge-base and run the WSL Ubuntu setup.
#>

$ErrorActionPreference = 'Stop'

function Write-Info  ($m) { Write-Host "[INFO]  $m" -ForegroundColor Cyan }
function Write-Ok    ($m) { Write-Host "[OK]    $m" -ForegroundColor Green }
function Write-Warn  ($m) { Write-Host "[WARN]  $m" -ForegroundColor Yellow }
function Write-Err   ($m) { Write-Host "[ERROR] $m" -ForegroundColor Red }

try {
    Write-Info "Knowledge Base — Step 2: project staging + Ubuntu setup"
    Write-Info "========================================================"

    $projectRoot = "C:\dev\knowledge-base"
    $deployDir   = Join-Path $projectRoot "deploy"
    $sourceDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
    $sourceRoot  = Split-Path -Parent $sourceDir

    Write-Info "Source root: $sourceRoot"
    Write-Info "Target root: $projectRoot"

    # Verify WSL is available
    Write-Info "Checking WSL availability..."
    $null = wsl --status 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Err "WSL is not available. Did you run 01-setup-wsl.ps1 and reboot?"
        exit 1
    }
    Write-Ok "WSL is available."

    # Verify Ubuntu-24.04 distro is installed
    $distros = (wsl --list --quiet) -join "`n"
    if ($distros -notmatch "Ubuntu-24\.04") {
        Write-Warn "Ubuntu-24.04 is not yet registered. Installing now (you may be prompted to set a Linux username/password)..."
        wsl --install -d Ubuntu-24.04
        if ($LASTEXITCODE -ne 0) { throw "wsl --install -d Ubuntu-24.04 failed." }
    }
    Write-Ok "Ubuntu-24.04 is registered."

    # Create target directory tree
    $subdirs = @('rawdocs','sorted','sorted\pdf','sorted\docx','sorted\xlsx','sorted\pptx','sorted\image','sorted\visio','sorted\text','sorted\unsupported','md_ready','wiki','config','logs','scripts','deploy')
    foreach ($d in $subdirs) {
        $full = Join-Path $projectRoot $d
        if (-not (Test-Path $full)) {
            New-Item -ItemType Directory -Path $full -Force | Out-Null
            Write-Ok "Created $full"
        }
    }

    # Copy project files (only if source != target)
    if ($sourceRoot -ne $projectRoot) {
        Write-Info "Copying project files into $projectRoot ..."
        $itemsToCopy = @('CLAUDE.md','scripts','deploy')
        foreach ($item in $itemsToCopy) {
            $src = Join-Path $sourceRoot $item
            if (Test-Path $src) {
                Copy-Item -Path $src -Destination $projectRoot -Recurse -Force
                Write-Ok "Copied $item"
            } else {
                Write-Warn "Source not found, skipping: $src"
            }
        }
    } else {
        Write-Info "Source already at target — skipping copy."
    }

    # Set ANTHROPIC_API_KEY as a Windows system environment variable
    $existingKey = [System.Environment]::GetEnvironmentVariable("ANTHROPIC_API_KEY", "Machine")
    if ($existingKey) {
        Write-Ok "ANTHROPIC_API_KEY already set as system environment variable — skipping."
    } else {
        Write-Host ""
        Write-Info "Enter your ANTHROPIC_API_KEY (starts with 'sk-ant-'). Leave blank to skip:"
        $apiKey = Read-Host -AsSecureString "API Key"
        $plainKey = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [Runtime.InteropServices.Marshal]::SecureStringToBSTR($apiKey))
        if ($plainKey -and $plainKey.StartsWith("sk-ant-")) {
            [System.Environment]::SetEnvironmentVariable("ANTHROPIC_API_KEY", $plainKey, "Machine")
            Write-Ok "ANTHROPIC_API_KEY saved as Windows system environment variable."
            Write-Warn "Restart your terminal/Claude Code session for the change to take effect."
        } else {
            Write-Warn "No valid API key entered. Set it later with:"
            Write-Warn '  [System.Environment]::SetEnvironmentVariable("ANTHROPIC_API_KEY", "sk-ant-...", "Machine")'
        }
    }

    # Run Ubuntu setup script
    Write-Info "Running Ubuntu setup inside WSL (this can take 5–10 minutes)..."
    $bashScript = "/mnt/c/dev/knowledge-base/deploy/03-setup-ubuntu.sh"
    wsl -d Ubuntu-24.04 -e bash -c "chmod +x '$bashScript' && '$bashScript'"
    if ($LASTEXITCODE -ne 0) { throw "Ubuntu setup script failed (exit $LASTEXITCODE)." }
    Write-Ok "Ubuntu setup completed."

    Write-Host ""
    Write-Host "========================================================" -ForegroundColor Green
    Write-Host "  SETUP COMPLETE."                                       -ForegroundColor Green
    Write-Host "========================================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Usage:" -ForegroundColor Cyan
    Write-Host "  Drop documents into: C:\dev\knowledge-base\rawdocs\"
    Write-Host "  Launch Claude Code:  .\04-launch.ps1"
    Write-Host "  Quick conversion:    .\quick-convert.ps1"
    Write-Host ""
} catch {
    Write-Err $_.Exception.Message
    exit 1
}
