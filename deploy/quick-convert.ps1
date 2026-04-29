<#
.SYNOPSIS
  Run scripts/convert.py inside WSL without launching Claude Code.
#>

$ErrorActionPreference = 'Stop'

function Write-Info ($m) { Write-Host "[INFO]  $m" -ForegroundColor Cyan }
function Write-Ok   ($m) { Write-Host "[OK]    $m" -ForegroundColor Green }
function Write-Err  ($m) { Write-Host "[ERROR] $m" -ForegroundColor Red }

try {
    Write-Info "Running conversion pipeline in WSL..."
    $null = wsl --status 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Err "WSL not available. Run deploy\01-setup-wsl.ps1 first."
        exit 1
    }
    wsl -d Ubuntu-24.04 -e bash -lic "cd ~/knowledge-base && python3 scripts/convert.py"
    if ($LASTEXITCODE -ne 0) {
        Write-Err "convert.py exited with code $LASTEXITCODE."
        exit $LASTEXITCODE
    }
    Write-Ok "Conversion complete. Output in C:\dev\knowledge-base\md_ready\"
} catch {
    Write-Err $_.Exception.Message
    exit 1
}
