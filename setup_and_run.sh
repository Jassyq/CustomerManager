#!/bin/bash
# ============================================================
#  Customer Price Manager – Mac Setup & Launcher
#  Run this script ONCE to set everything up, then use
#  run_app.command for daily use.
# ============================================================

set -e  # Exit on any error

echo ""
echo "======================================"
echo "  Customer Price Manager – Mac Setup"
echo "======================================"
echo ""

# ── 1. Install Homebrew if missing ──────────────────────────
if ! command -v brew &>/dev/null; then
    echo "📦 Homebrew not found. Installing Homebrew..."
    echo "   (You may be prompted for your Mac password)"
    /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

    # Add Homebrew to PATH for Apple Silicon Macs
    if [[ -f "/opt/homebrew/bin/brew" ]]; then
        eval "$(/opt/homebrew/bin/brew shellenv)"
        echo 'eval "$(/opt/homebrew/bin/brew shellenv)"' >> ~/.zprofile
    fi
    echo "✅ Homebrew installed."
else
    echo "✅ Homebrew already installed."
fi

# ── 2. Install Python if missing ────────────────────────────
if ! command -v python3 &>/dev/null; then
    echo "🐍 Python not found. Installing Python via Homebrew..."
    brew install python
    echo "✅ Python installed."
else
    PYTHON_VER=$(python3 --version)
    echo "✅ $PYTHON_VER already installed."
fi

# ── 3. Upgrade pip ──────────────────────────────────────────
echo ""
echo "📦 Upgrading pip..."
python3 -m pip install --upgrade pip -q

# ── 4. Install app dependencies ─────────────────────────────
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

if [[ -f "$SCRIPT_DIR/requirements.txt" ]]; then
    echo "📋 Found requirements.txt – installing packages..."
    python3 -m pip install -r "$SCRIPT_DIR/requirements.txt" -q
else
    echo "📋 No requirements.txt found – installing default packages..."
    python3 -m pip install -q streamlit pandas openpyxl
fi
echo "✅ All packages installed."

# ── 5. Create the daily run_app.command launcher ────────────
LAUNCHER="$SCRIPT_DIR/run_app.command"
cat > "$LAUNCHER" << 'EOF'
#!/bin/bash
# Daily launcher – double-click this to start the app
cd "$(dirname "$0")"
echo ""
echo "🚀 Starting Customer Price Manager..."
echo ""
python3 -m streamlit run app.py
EOF

chmod +x "$LAUNCHER"
echo ""
echo "✅ Created run_app.command (double-click this for daily use)"

# ── 6. Launch the app now ───────────────────────────────────
echo ""
echo "🚀 Launching Customer Price Manager..."
echo "   (Press Ctrl+C to stop the app)"
echo ""
cd "$SCRIPT_DIR"
python3 -m streamlit run app.py
