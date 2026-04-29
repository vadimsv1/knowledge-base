#!/usr/bin/env bash
# Step 3 — Ubuntu-side setup. Runs inside WSL Ubuntu-24.04.
# Installs Python deps, Node + Claude Code CLI, fixes line endings, prompts for API key.

set -euo pipefail

C_INFO='\033[0;36m'
C_OK='\033[0;32m'
C_WARN='\033[0;33m'
C_ERR='\033[0;31m'
C_OFF='\033[0m'

info() { echo -e "${C_INFO}[INFO]${C_OFF}  $*"; }
ok()   { echo -e "${C_OK}[OK]${C_OFF}    $*"; }
warn() { echo -e "${C_WARN}[WARN]${C_OFF}  $*"; }
err()  { echo -e "${C_ERR}[ERROR]${C_OFF} $*" >&2; }

trap 'err "Setup failed at line $LINENO"' ERR

PROJECT=/mnt/c/dev/knowledge-base
LINK=$HOME/knowledge-base

info "Knowledge Base — Step 3: Ubuntu-side setup"
info "==========================================="

# Sudo wrapper (no password prompt friendliness)
SUDO=""
if [ "$(id -u)" -ne 0 ]; then
    SUDO="sudo"
    info "Will use sudo for system package installation."
fi

info "Updating apt package lists..."
$SUDO apt-get update -y
ok "apt updated."

info "Installing system packages: python3, python3-pip, nodejs, npm, curl..."
$SUDO apt-get install -y python3 python3-pip nodejs npm curl ca-certificates
ok "System packages installed."

info "Installing Python dependencies (pymupdf, pymupdf4llm, mammoth, openpyxl, python-pptx, Pillow, anthropic)..."
pip3 install --break-system-packages --upgrade \
    pymupdf pymupdf4llm mammoth openpyxl python-pptx Pillow anthropic
ok "Python dependencies installed."

info "Installing Claude Code CLI globally via npm..."
$SUDO npm install -g @anthropic-ai/claude-code
ok "Claude Code CLI installed."

info "Creating symlink: $LINK -> $PROJECT"
ln -sfn "$PROJECT" "$LINK"
ok "Symlink created."

info "Normalizing line endings on scripts (CRLF -> LF)..."
if [ -d "$PROJECT/scripts" ]; then
    find "$PROJECT/scripts" -type f -name "*.py" -exec sed -i 's/\r$//' {} \;
    find "$PROJECT/scripts" -type f -name "*.sh" -exec sed -i 's/\r$//' {} \;
    ok "Line endings normalized."
else
    warn "$PROJECT/scripts not found — skipping line-ending fix."
fi

if [ -f "$PROJECT/scripts/run.sh" ]; then
    chmod +x "$PROJECT/scripts/run.sh"
    ok "Marked run.sh executable."
fi

# Prompt for API key
echo
if grep -q "ANTHROPIC_API_KEY" "$HOME/.bashrc" 2>/dev/null; then
    warn "ANTHROPIC_API_KEY already present in ~/.bashrc — skipping prompt."
else
    info "Enter your ANTHROPIC_API_KEY (starts with 'sk-ant-'). Leave blank to skip:"
    read -r -s API_KEY
    echo
    if [ -n "${API_KEY:-}" ]; then
        printf '\nexport ANTHROPIC_API_KEY="%s"\n' "$API_KEY" >> "$HOME/.bashrc"
        ok "API key saved to ~/.bashrc"
    else
        warn "No API key entered. Add it later by editing ~/.bashrc:"
        warn '  echo '"'"'export ANTHROPIC_API_KEY="sk-ant-..."'"'"' >> ~/.bashrc'
    fi
fi

echo
echo -e "${C_OK}===========================================${C_OFF}"
echo -e "${C_OK}  Ubuntu setup complete.${C_OFF}"
echo -e "${C_OK}===========================================${C_OFF}"
echo
echo "Quick commands (run from WSL Ubuntu):"
echo "  cd ~/knowledge-base"
echo "  python3 scripts/convert.py            # full conversion pipeline"
echo "  python3 scripts/convert.py --status   # show state"
echo "  claude --dangerously-skip-permissions # launch Claude Code"
echo
