#!/bin/bash
# Comment Sheet Aggregator - Mac Setup Script

# Ensure we are running from the script's directory
cd "$(dirname "$0")"

if ! command -v python3 &> /dev/null; then
    echo "Python 3 is not installed. Please install it first."
    echo "Visit python.org or run: brew install python"
    exit 1
fi

if [ ! -f "src/gui_app.py" ]; then
    echo "Error: src/gui_app.py not found in $(pwd)"
    exit 1
fi

# Clean up existing environment to avoid conflicts
echo "Cleaning up previous build files..."
rm -rf venv build dist *.spec CommentAggregator.app

echo "Creating new virtual environment..."
python3 -m venv venv

echo "Activating virtual environment..."
source venv/bin/activate

echo "Installing dependencies..."
# IMPORTANT: 'tk' is NOT installed here. It is part of Python standard library.
# Installing 'tk' via pip causes crashes on macOS.
pip install pandas openpyxl xlrd pyinstaller streamlit

echo "Building Mac Application..."
# Use --onedir for stability
pyinstaller --onedir --windowed --name "CommentAggregator" --clean --noconfirm --hidden-import=xlrd src/gui_app.py

# Move app to root folder and cleanup
mv dist/CommentAggregator.app .
rm -rf build dist *.spec

echo "Build complete! App is located in this folder: CommentAggregator.app"

echo "Setup complete!"
echo ""
echo "To run the app directly:"
echo "open CommentAggregator.app"
