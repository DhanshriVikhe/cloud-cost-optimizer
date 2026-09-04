#!/bin/bash
# Cloud Instance Optimizer — startup script
set -e

echo ""
echo "============================================"
echo "  Cloud Instance Optimizer — Starting..."
echo "============================================"
echo ""

# Check Python
if ! command -v python3 &> /dev/null; then
  echo "ERROR: python3 not found. Please install Python 3.9+"
  exit 1
fi

echo "Python: $(python3 --version)"

# Install deps if needed
echo ""
echo "Checking dependencies..."
pip install -q -r requirements.txt

echo ""
echo "Starting server on http://localhost:5000"
echo "Press Ctrl+C to stop"
echo ""

python3 app.py
