#!/bin/bash
cd "$(dirname "$0")"

# Activate virtual environment if it exists
if [ -d "venv" ]; then
    source venv/bin/activate
else
    echo "Virtual environment not found. Please run setup_mac.sh first!"
    exit 1
fi

echo "Starting Web App..."

# Ensure .streamlit/config.toml exists to disable email prompt (which blocks browser launch)
mkdir -p .streamlit
if [ ! -f .streamlit/config.toml ]; then
    echo "Creating Streamlit config..."
    echo "[browser]" > .streamlit/config.toml
    echo "gatherUsageStats = false" >> .streamlit/config.toml
    echo "[server]" >> .streamlit/config.toml
    echo "headless = false" >> .streamlit/config.toml
fi

echo "----------------------------------------------------------------"
echo "If the browser does not open automatically, please copy and paste"
echo "the URL below (usually http://localhost:8501) into your browser."
echo "----------------------------------------------------------------"
echo ""

streamlit run src/streamlit_app.py
