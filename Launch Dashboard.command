#!/bin/bash
# Double-click this file on your Mac to launch the Surf Dashboard
# (You may need to right-click → Open the first time due to Gatekeeper)

cd "$(dirname "$0")"

echo "================================================"
echo "  🏄  East Coast Surf Dashboard"
echo "================================================"

# Check if Python 3 is available
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 not found. Please install it from https://www.python.org"
    read -p "Press Enter to exit..."
    exit 1
fi

# Install/upgrade dependencies quietly
echo "📦 Checking dependencies..."
python3 -m pip install -r requirements.txt -q --upgrade

echo "🚀 Launching dashboard at http://localhost:8501"
echo "   (Press Ctrl+C in this window to stop)"
echo ""

# Launch Streamlit — opens browser automatically
python3 -m streamlit run surf_dashboard.py \
    --server.headless false \
    --browser.gatherUsageStats false \
    --theme.primaryColor "#1F7A8C" \
    --theme.backgroundColor "#f0f4f8" \
    --theme.secondaryBackgroundColor "#ffffff" \
    --theme.textColor "#1B3A5C"
