#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

if [[ ! -d .venv ]]; then
  echo "Error: .venv not found. Create it first: python -m venv .venv"
  exit 1
fi

# Stop any old app process to prevent port conflicts.
pkill -f "streamlit run app.py" 2>/dev/null || true

source .venv/bin/activate
exec streamlit run app.py --server.port 8501 --server.address 0.0.0.0 --server.headless true
