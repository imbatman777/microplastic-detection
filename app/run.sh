./#!/bin/bash
# PlastiSense AI — Quick Start
# Usage: bash app/run.sh

cd "$(dirname "$0")"

# Use the project's virtual environment Python
VENV_PYTHON="../.venv/bin/python"

echo "═══════════════════════════════════════════════"
echo "  PlastiSense AI — Starting Server"
echo "═══════════════════════════════════════════════"
echo ""
echo "  Open in browser:  http://localhost:8000"
echo ""

"$VENV_PYTHON" -m uvicorn server:app --host 0.0.0.0 --port 8000 --reload
