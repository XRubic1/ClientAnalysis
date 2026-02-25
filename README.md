# ClientAnalysis

TRU Funding Client Analysis Tool — upload a CliAnalysis PDF and get a formatted Excel spreadsheet with sales data, averages, and trend indicators.

## Features

- **Browser-only**: Runs 100% in the browser (no server required). Works on [GitHub Pages](https://pages.github.com/).
- **PDF upload**: Drop or browse to your TRU Funding Client Analysis report (PDF).
- **Auto-extract**: Parses client names, codes, and monthly sales (Oct–Jan).
- **Excel export**: Downloads a styled `.xlsx` with two sheets: Client Sales Analysis and Trend Summary.

## Quick start

1. **GitHub Pages**: Enable Pages for this repo (Settings → Pages → Deploy from branch → main). Open `https://XRubic1.github.io/ClientAnalysis/`.
2. **Local**: Open `index.html` in a browser, or run a static server:
   ```bash
   npx serve .
   # or: python -m http.server 8080
   ```
3. **With Flask (optional)**: For a local server-based run:
   ```bash
   pip install flask pdfplumber openpyxl
   python app.py
   ```
   Then open http://localhost:5050

## Tech

- **Frontend**: HTML, CSS, JavaScript — PDF.js (Mozilla), ExcelJS (CDN).
- **Backend (optional)**: Python, Flask, pdfplumber, openpyxl.

TRU Funding LLC · Client Analysis Tool
