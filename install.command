#!/bin/bash
# Mantools installer for macOS (per-user, no sudo).
# Builds ~/Applications/Mantools.app backed by a private virtual environment.
#
# Run it:  double-click in Finder (after `chmod +x install.command`),
#          or in Terminal:  bash install.command
set -euo pipefail

DIR="$(cd "$(dirname "$0")" && pwd)"
SRC=""
for d in "$DIR" "$(dirname "$DIR")"; do
    if [ -f "$d/mantools.py" ]; then SRC="$d"; break; fi
done
[ -n "$SRC" ] || { echo "ERROR: mantools.py not found next to the installer."; exit 1; }

APP="$HOME/Applications/Mantools.app"
SUPPORT="$HOME/Library/Application Support/Mantools"

echo "Mantools installer (macOS)"
echo "  Source : $SRC"
echo "  App    : $APP"

PY="$(command -v python3 || true)"
[ -n "$PY" ] || { echo "ERROR: python3 not found. Install from https://www.python.org/downloads/macos/"; exit 1; }
"$PY" -c 'import sys; raise SystemExit(0 if sys.version_info[:2] >= (3, 8) else 1)' \
    || { echo "ERROR: Python 3.8+ required."; exit 1; }
if ! "$PY" -c 'import tkinter' 2>/dev/null; then
    echo "WARNING: tkinter is missing from this Python."
    echo "         Homebrew users: brew install python-tk"
    echo "         Or install Python from python.org (bundles Tk)."
fi

mkdir -p "$SUPPORT"
if [ ! -x "$SUPPORT/venv/bin/python" ]; then
    echo "  Creating virtual environment..."
    "$PY" -m venv "$SUPPORT/venv"
fi
VPY="$SUPPORT/venv/bin/python"

if [ "${SKIP_DEPS:-0}" != "1" ]; then
    echo "  Installing dependencies (this can take a few minutes)..."
    "$VPY" -m pip install --upgrade pip
    "$VPY" -m pip install -r "$SRC/requirements.txt"
fi

cp -f "$SRC/mantools.py" "$SUPPORT/mantools.py"
[ -f "$SRC/mantools.png" ] && cp -f "$SRC/mantools.png" "$SUPPORT/mantools.png" || true

echo "  Building Mantools.app ..."
rm -rf "$APP"
mkdir -p "$APP/Contents/MacOS" "$APP/Contents/Resources"

# Icon: use the shipped .icns, else build one from the PNG with macOS tools.
if [ -f "$DIR/mantools.icns" ]; then
    cp "$DIR/mantools.icns" "$APP/Contents/Resources/mantools.icns"
elif command -v iconutil >/dev/null && command -v sips >/dev/null && [ -f "$SRC/mantools.png" ]; then
    ISET="$(mktemp -d)/mantools.iconset"; mkdir -p "$ISET"
    for s in 16 32 128 256 512; do
        d=$((s * 2))
        sips -z "$s" "$s" "$SRC/mantools.png" --out "$ISET/icon_${s}x${s}.png"    >/dev/null
        sips -z "$d" "$d" "$SRC/mantools.png" --out "$ISET/icon_${s}x${s}@2x.png" >/dev/null
    done
    iconutil -c icns "$ISET" -o "$APP/Contents/Resources/mantools.icns" || true
fi

cat > "$APP/Contents/MacOS/Mantools" <<EOF
#!/bin/bash
exec "$SUPPORT/venv/bin/python" "$SUPPORT/mantools.py" "\$@"
EOF
chmod +x "$APP/Contents/MacOS/Mantools"

cat > "$APP/Contents/Info.plist" <<'EOF'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
    <key>CFBundleName</key><string>Mantools</string>
    <key>CFBundleDisplayName</key><string>Mantools</string>
    <key>CFBundleIdentifier</key><string>com.mantools.app</string>
    <key>CFBundleVersion</key><string>4.0</string>
    <key>CFBundleShortVersionString</key><string>4.0</string>
    <key>CFBundlePackageType</key><string>APPL</string>
    <key>CFBundleExecutable</key><string>Mantools</string>
    <key>CFBundleIconFile</key><string>mantools.icns</string>
    <key>NSHighResolutionCapable</key><true/>
</dict></plist>
EOF

echo ""
echo "Mantools installed to $APP"
echo "  Open it from ~/Applications (double-click), or drag it to the Dock."
echo "  Note: 'Office -> PDF' conversions need Microsoft Office + Windows and"
echo "        are automatically disabled on macOS."
