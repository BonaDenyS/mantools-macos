#!/bin/bash
# Remove Mantools (macOS, per-user).
set -euo pipefail
rm -rf "$HOME/Applications/Mantools.app"
rm -rf "$HOME/Library/Application Support/Mantools"
echo "Mantools uninstalled."
