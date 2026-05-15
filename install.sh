#!/bin/bash
set -e

APP_DIR="$HOME/Pokemon-Sammlung"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

echo "=== Poke Inv – Setup ==="
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
APP_BUNDLE="$HOME/Desktop/Poke Inv.app"
if [ ! -d "$APP_BUNDLE" ]; then
    mkdir -p "$APP_BUNDLE/Contents/MacOS"
    mkdir -p "$APP_BUNDLE/Contents/Resources"

    # Generate app icon (yellow P on dark background)
    "$APP_DIR/.venv/bin/python3" "$SCRIPT_DIR/create_icon.py" "$APP_BUNDLE/Contents/Resources/AppIcon.icns"

    # Try native launcher (fast), fall back to bash script
    if command -v cc &> /dev/null; then
        cat > /tmp/_poke_launcher.c << 'CEOF'
#include <stdlib.h>
#include <stdio.h>
#include <string.h>
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
        cc -o "$APP_BUNDLE/Contents/MacOS/PokeInv" /tmp/_poke_launcher.c -O2
        rm /tmp/_poke_launcher.c
    else
        cat > "$APP_BUNDLE/Contents/MacOS/PokeInv" << 'BASHEOF'
#!/bin/bash
exec "$HOME/Pokemon-Sammlung/.venv/bin/python3" "$HOME/Pokemon-Sammlung/pokemon_card_manager.py"
BASHEOF
        chmod +x "$APP_BUNDLE/Contents/MacOS/PokeInv"
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
else
    echo "✓ Desktop app already exists"
fi

echo ""
echo "=== Setup complete! ==="
echo "Double-click 'Poke Inv' on your Desktop to launch."
echo "Data is stored in: $APP_DIR"
