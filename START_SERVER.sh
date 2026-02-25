#!/bin/bash
echo "============================================"
echo "  TRU Funding Client Analysis Tool"
echo "============================================"
echo ""
echo "Installing dependencies..."
pip install flask pdfplumber openpyxl --break-system-packages -q 2>/dev/null || \
pip install flask pdfplumber openpyxl -q 2>/dev/null

echo "Starting server on http://localhost:5050"
echo "Open your browser to: http://localhost:5050"
echo ""
echo "Press Ctrl+C to stop."
echo ""
python3 app.py
