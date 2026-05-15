#!/bin/bash

APP_DIR="$HOME/Pokemon-Sammlung"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

echo "=== Poke Inv – Setup ==="
echo ""

# Find a good Python 3
PYTHON=""
for p in python3.14 python3.13 python3.12 python3.11 python3.10 python3; do
    if command -v "$p" &> /dev/null; then
        PYTHON="$(command -v "$p")"
        break
    fi
done

if [ -z "$PYTHON" ]; then
    echo "ERROR: Python 3 not found. Install from https://www.python.org/downloads/"
    exit 1
fi

PY_VERSION=$("$PYTHON" -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')")
echo "Using Python $PY_VERSION ($PYTHON)"

# Check tkinter is available
if ! "$PYTHON" -c "import tkinter" 2>/dev/null; then
    echo "ERROR: tkinter not available for $PYTHON"
    echo "Install Python from https://www.python.org/downloads/ (includes tkinter)"
    exit 1
fi

# Create base directories
mkdir -p "$APP_DIR/Fotos"
mkdir -p "$APP_DIR/Karten"
echo "✓ Created $APP_DIR"

# Copy main script
cp "$SCRIPT_DIR/pokemon_card_manager.py" "$APP_DIR/pokemon_card_manager.py"
echo "✓ Copied app script"

# Create virtual environment
if [ ! -d "$APP_DIR/.venv" ]; then
    "$PYTHON" -m venv "$APP_DIR/.venv"
    echo "✓ Created virtual environment"
else
    echo "✓ Virtual environment already exists"
fi

# Upgrade pip and install dependencies
"$APP_DIR/.venv/bin/python3" -m pip install --upgrade pip --quiet 2>/dev/null
"$APP_DIR/.venv/bin/pip" install --quiet openpyxl Pillow
echo "✓ Installed dependencies (openpyxl, Pillow)"

# Create Excel workbook if it doesn't exist
if [ ! -f "$APP_DIR/Pokemon-Inventar.xlsx" ]; then
    "$APP_DIR/.venv/bin/python3" "$SCRIPT_DIR/create_excel.py" "$APP_DIR/Pokemon-Inventar.xlsx"
    echo "✓ Created Pokemon-Inventar.xlsx"
else
    echo "✓ Pokemon-Inventar.xlsx already exists"
fi

# Create macOS .app bundle (always recreate to pick up updates)
APP_BUNDLE="$HOME/Desktop/Poke Inv.app"
rm -rf "$APP_BUNDLE"
mkdir -p "$APP_BUNDLE/Contents/MacOS"
mkdir -p "$APP_BUNDLE/Contents/Resources"

# Generate app icon
if "$APP_DIR/.venv/bin/python3" "$SCRIPT_DIR/create_icon.py" "$APP_BUNDLE/Contents/Resources/AppIcon.icns" 2>/dev/null; then
    echo "✓ Generated app icon"
else
    echo "⚠ Icon generation skipped (not critical)"
fi

# Try native launcher (fast), fall back to bash script
if command -v cc &> /dev/null; then
    cat > /tmp/_poke_launcher.c << 'CEOF'
#include <stdlib.h>
#include <stdio.h>
#include <unistd.h>
int main(int argc, char *argv[]) {
    const char *home = getenv("HOME");
    if (!home) { fprintf(stderr, "HOME not set\n"); return 1; }
    char python[512], script[512];
    snprintf(python, sizeof(python), "%s/Pokemon-Sammlung/.venv/bin/python3", home);
    snprintf(script, sizeof(script), "%s/Pokemon-Sammlung/pokemon_card_manager.py", home);
    execl(python, "python3", script, NULL);
    perror("execl failed");
    return 1;
}
CEOF
    if cc -o "$APP_BUNDLE/Contents/MacOS/PokeInv" /tmp/_poke_launcher.c -O2 2>/dev/null; then
        echo "✓ Compiled native launcher"
    else
        echo "⚠ C compile failed, using bash launcher"
        cat > "$APP_BUNDLE/Contents/MacOS/PokeInv" << BASHEOF
#!/bin/bash
exec "$APP_DIR/.venv/bin/python3" "$APP_DIR/pokemon_card_manager.py"
BASHEOF
        chmod +x "$APP_BUNDLE/Contents/MacOS/PokeInv"
    fi
    rm -f /tmp/_poke_launcher.c
else
    cat > "$APP_BUNDLE/Contents/MacOS/PokeInv" << BASHEOF
#!/bin/bash
exec "$APP_DIR/.venv/bin/python3" "$APP_DIR/pokemon_card_manager.py"
BASHEOF
    chmod +x "$APP_BUNDLE/Contents/MacOS/PokeInv"
    echo "✓ Created bash launcher"
fi

cat > "$APP_BUNDLE/Contents/Info.plist" << 'PLISTEOF'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleExecutable</key>
    <string>PokeInv</string>
    <key>CFBundleName</key>
    <string>Poke Inv</string>
    <key>CFBundleIdentifier</key>
    <string>com.pokemon.poke-inv</string>
    <key>CFBundleVersion</key>
    <string>1.0</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>CFBundleIconFile</key>
    <string>AppIcon</string>
    <key>LSUIElement</key>
    <false/>
</dict>
</plist>
PLISTEOF

codesign --force --sign - "$APP_BUNDLE" 2>/dev/null || true
echo "✓ Created desktop app: Poke Inv.app"

echo ""
echo "=== Setup complete! ==="
echo "Double-click 'Poke Inv' on your Desktop to launch."
echo "Data is stored in: $APP_DIR"
