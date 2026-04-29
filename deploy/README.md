# Knowledge Base — Deployment Guide

One-click deployment of the Knowledge Base Document Converter onto a fresh Windows 10/11 machine. Sets up WSL2 + Ubuntu 24.04 + Python deps + Claude Code CLI.

## Prerequisites

- Windows 10 (build 19041+) or Windows 11
- Administrator account
- ~5 GB free disk space
- Internet connection
- An Anthropic API key (`sk-ant-...`)

## Steps

### Step 1 — Enable WSL2 + install Ubuntu (run as Admin)

Open **PowerShell as Administrator** and run:

```powershell
cd C:\dev\knowledge-base\deploy
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\01-setup-wsl.ps1
```

When it finishes: **restart your PC**.

### Step 2 — Install project + Python deps + Claude Code (run as Admin)

After reboot, open **PowerShell as Administrator** again:

```powershell
cd C:\dev\knowledge-base\deploy
.\02-setup-project.ps1
```

This will:
- Create `C:\dev\knowledge-base\` with all subdirectories
- Copy `scripts/`, `CLAUDE.md`, etc. into place
- Invoke `03-setup-ubuntu.sh` inside WSL to install Python deps + Claude Code CLI
- Prompt you for your `ANTHROPIC_API_KEY` and save it to `~/.bashrc`

### Step 3 — Daily use

To launch Claude Code in the project:

```powershell
C:\dev\knowledge-base\deploy\04-launch.ps1
```

Or to just run the converter without Claude Code:

```powershell
C:\dev\knowledge-base\deploy\quick-convert.ps1
```

## Manual conversion

Inside WSL:

```bash
cd ~/knowledge-base
python3 scripts/convert.py            # full pipeline
python3 scripts/convert.py --status   # show state
python3 scripts/convert.py --reset    # wipe state and start fresh
```

## File layout after deployment

```
C:\dev\knowledge-base\
├── rawdocs/          # drop source files here
├── sorted/           # auto-sorted by type
├── md_ready/         # converted Markdown output
├── wiki/             # compiled knowledge base
├── config/           # state JSON
├── logs/             # per-run logs
├── scripts/          # convert.py, run.sh
└── deploy/           # this folder
```

## Troubleshooting

- **`wsl: command not found`** — Step 1 didn't finish. Re-run `01-setup-wsl.ps1` and reboot.
- **`Ubuntu-24.04` not in `wsl --list`** — re-run `wsl --install -d Ubuntu-24.04` manually and complete the first-launch prompt to set a Linux user.
- **`ANTHROPIC_API_KEY not set`** — open WSL, edit `~/.bashrc`, ensure `export ANTHROPIC_API_KEY="sk-ant-..."` line is present, then `source ~/.bashrc`.
- **CRLF errors when running shell scripts** — re-run `03-setup-ubuntu.sh` (it normalizes line endings).
