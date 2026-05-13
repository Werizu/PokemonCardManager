#!/bin/bash
set -e

APP_DIR="$HOME/Pokemon-Sammlung"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

echo "=== Pokemon Card Manager – Setup ==="
echo ""

# Create base directories
mkdir -p "$APP_DIR/Fotos"
mkdir -p "$APP_DIR/Karten"
echo "✓ Created $APP_DIR"

# Copy main script
cp "$SCRIPT_DIR/pokemon_card_manager.py" "$APP_DIR/pokemon_card_manager.py"
echo "✓ Copied app script"

# Create virtual environment
if [ ! -d "$APP_DIR/.venv" ]; then
    python3 -m venv "$APP_DIR/.venv"
    echo "✓ Created virtual environment"
else
    echo "✓ Virtual environment already exists"
fi

# Install dependencies
"$APP_DIR/.venv/bin/pip" install --quiet openpyxl Pillow
echo "✓ Installed dependencies (openpyxl, Pillow)"

# Create Excel workbook if it doesn't exist
if [ ! -f "$APP_DIR/Pokemon-Inventar.xlsx" ]; then
    "$APP_DIR/.venv/bin/python3" "$SCRIPT_DIR/create_excel.py" "$APP_DIR/Pokemon-Inventar.xlsx"
    echo "✓ Created Pokemon-Inventar.xlsx"
else
    echo "✓ Pokemon-Inventar.xlsx already exists"
fi

# Create macOS .app bundle
APP_BUNDLE="$HOME/Desktop/Pokemon-Eintragen.app"
if [ ! -d "$APP_BUNDLE" ]; then
    mkdir -p "$APP_BUNDLE/Contents/MacOS"
    cat > "$APP_BUNDLE/Contents/MacOS/run" << 'RUNEOF'
#!/bin/bash
"$HOME/Pokemon-Sammlung/.venv/bin/python3" "$HOME/Pokemon-Sammlung/pokemon_card_manager.py"
RUNEOF
    chmod +x "$APP_BUNDLE/Contents/MacOS/run"

    cat > "$APP_BUNDLE/Contents/Info.plist" << 'PLISTEOF'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleExecutable</key>
    <string>run</string>
    <key>CFBundleName</key>
    <string>Pokemon Eintragen</string>
    <key>CFBundleIdentifier</key>
    <string>com.pokemon.card-manager</string>
    <key>CFBundleVersion</key>
    <string>1.0</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>LSUIElement</key>
    <true/>
</dict>
</plist>
PLISTEOF
    echo "✓ Created desktop app: Pokemon-Eintragen.app"
else
    echo "✓ Desktop app already exists"
fi

echo ""
echo "=== Setup complete! ==="
echo "Double-click 'Pokemon-Eintragen' on your Desktop to launch."
echo "Data is stored in: $APP_DIR"
