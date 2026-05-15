#!/bin/bash

APP_DIR="$HOME/Pokemon-Sammlung"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

echo "=== Poke Inv – Setup ==="
echo ""

# ── 1. Homebrew ──
if ! command -v brew &> /dev/null; then
    echo "Installing Homebrew..."
    /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
    # Add brew to PATH for this session (Apple Silicon vs Intel)
    if [ -f /opt/homebrew/bin/brew ]; then
        eval "$(/opt/homebrew/bin/brew shellenv)"
    elif [ -f /usr/local/bin/brew ]; then
        eval "$(/usr/local/bin/brew shellenv)"
    fi
    echo "✓ Installed Homebrew"
else
    echo "✓ Homebrew found"
fi

# ── 2. Python 3 with tkinter ──
PYTHON=""
for p in python3.14 python3.13 python3.12 python3.11 python3.10 python3; do
    if command -v "$p" &> /dev/null && "$p" -c "import tkinter" 2>/dev/null; then
        PYTHON="$(command -v "$p")"
        break
    fi
done

if [ -z "$PYTHON" ]; then
    echo "Installing Python 3 via Homebrew..."
    brew install python-tk python3
    PYTHON="$(command -v python3)"
    if [ -z "$PYTHON" ]; then
        echo "ERROR: Python installation failed."
        exit 1
    fi
    echo "✓ Installed Python 3"
else
    PY_VERSION=$("$PYTHON" -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')")
    echo "✓ Python $PY_VERSION found ($PYTHON)"
fi

# ── 3. Xcode Command Line Tools (for C compiler) ──
if ! command -v cc &> /dev/null; then
    echo "Installing Xcode Command Line Tools..."
    xcode-select --install 2>/dev/null
    echo "  Waiting for installation to finish (click Install in the popup)..."
    until command -v cc &> /dev/null; do
        sleep 5
    done
    echo "✓ Installed Xcode Command Line Tools"
else
    echo "✓ C compiler found"
fi

# ── 4. Git ──
if ! command -v git &> /dev/null; then
    echo "Installing Git..."
    brew install git
    echo "✓ Installed Git"
else
    echo "✓ Git found"
fi

echo ""

# ── 5. Create directories ──
mkdir -p "$APP_DIR/Fotos"
mkdir -p "$APP_DIR/Karten"
echo "✓ Created $APP_DIR"

# ── 6. Copy app script ──
cp "$SCRIPT_DIR/pokemon_card_manager.py" "$APP_DIR/pokemon_card_manager.py"
echo "✓ Copied app script"

# ── 7. Virtual environment ──
if [ ! -d "$APP_DIR/.venv" ]; then
    "$PYTHON" -m venv "$APP_DIR/.venv"
    echo "✓ Created virtual environment"
else
    echo "✓ Virtual environment exists"
fi

# ── 8. Pip + dependencies ──
"$APP_DIR/.venv/bin/python3" -m pip install --upgrade pip --quiet 2>/dev/null
"$APP_DIR/.venv/bin/pip" install --quiet openpyxl Pillow
echo "✓ Installed dependencies"

# ── 9. Excel workbook ──
if [ ! -f "$APP_DIR/Pokemon-Inventar.xlsx" ]; then
    "$APP_DIR/.venv/bin/python3" "$SCRIPT_DIR/create_excel.py" "$APP_DIR/Pokemon-Inventar.xlsx"
    echo "✓ Created Pokemon-Inventar.xlsx"
else
    echo "✓ Pokemon-Inventar.xlsx already exists"
fi

# ── 10. Desktop app ──
APP_BUNDLE="$HOME/Desktop/Poke Inv.app"
rm -rf "$APP_BUNDLE"
mkdir -p "$APP_BUNDLE/Contents/MacOS"
mkdir -p "$APP_BUNDLE/Contents/Resources"

# Icon
if "$APP_DIR/.venv/bin/python3" "$SCRIPT_DIR/create_icon.py" "$APP_BUNDLE/Contents/Resources/AppIcon.icns" 2>/dev/null; then
    echo "✓ Generated app icon"
else
    echo "⚠ Icon skipped (not critical)"
fi

# Native C launcher
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
    cat > "$APP_BUNDLE/Contents/MacOS/PokeInv" << BASHEOF
#!/bin/bash
exec "$APP_DIR/.venv/bin/python3" "$APP_DIR/pokemon_card_manager.py"
BASHEOF
    chmod +x "$APP_BUNDLE/Contents/MacOS/PokeInv"
    echo "✓ Created bash launcher"
fi
rm -f /tmp/_poke_launcher.c

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
echo "✓ Created Poke Inv.app on Desktop"

echo ""
echo "=== Setup complete! ==="
echo "Double-click 'Poke Inv' on your Desktop to launch."
