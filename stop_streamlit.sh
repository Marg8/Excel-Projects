#!/usr/bin/env bash
set -euo pipefail

pkill -f "streamlit run app.py" 2>/dev/null || true
echo "Stopped Streamlit app process (if it was running)."
