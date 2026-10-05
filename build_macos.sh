#!/usr/bin/env bash
# Build the macOS app bundle with PyInstaller.
# Run on macOS (PyInstaller cannot cross-build for other OSes).
set -euo pipefail
cd "$(dirname "$0")"

python3 -m venv venv
./venv/bin/pip install --upgrade pip
./venv/bin/pip install -r requirements.txt pyinstaller

# onedir + windowed produces a proper .app bundle (the recommended mode on
# macOS; --onefile is discouraged for .app bundles and clashes with Gatekeeper).
#
# --collect-submodules truststore: its backend is chosen behind a platform test,
# so PyInstaller's static analysis cannot see the per-OS submodule.
./venv/bin/pyinstaller \
  --name "PDPC函釋更新工具" \
  --windowed \
  --noconfirm \
  --collect-submodules truststore \
  app.py

echo
echo "Done. The app is at: dist/PDPC函釋更新工具.app"
