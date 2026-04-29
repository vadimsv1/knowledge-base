#!/usr/bin/env bash
set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
case "${1:-run}" in
    status)  python3 "$SCRIPT_DIR/convert.py" --status ;;
    reset)   python3 "$SCRIPT_DIR/convert.py" --reset ;;
    retry)   python3 "$SCRIPT_DIR/convert.py" --retry-errors ;;
    run|*)   python3 "$SCRIPT_DIR/convert.py" "$@" ;;
esac