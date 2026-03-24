# 4G KPI Report Generator

This tool automates the generation of a 4G KPI Report Excel file based on raw data exported from your network management system. It automatically applies proper layout, calculations, and color-coding logic to save you time.

## Requirements

1. **Python**: You must have Python installed on your computer. You can download it for free from [python.org](https://www.python.org/downloads/). When installing, ensure you check the box that says **"Add Python to PATH"**.
2. **Raw Data**: You will need your raw exported Excel data files (Site or City level).

## How to Install and Run

1. Simply double-click on the `run.bat` file.
2. The script will automatically configure everything you need (creating a virtual environment and installing required packages).
3. Wait a few moments for the application window to appear.

## How to Use

1. **Select Raw Data File**: Click the "Browse" button and select your raw Excel data file.
2. **Enter Event Name**: Type the name of the event in the text box (e.g., "Omnisport_Match").
3. **Generate Report**: Click the "Generate Report" button.
4. **Result**: Once finished, the tool will notify you and the newly generated Excel report will be automatically saved directly into your computer's **Downloads** folder. The filename will follow the standard format: `KEA_4G_nameOfEvent_YYYYMMDD_HHMMSS.xlsx`.

## Privacy & Confidentiality

This program runs **100% locally** on your computer. No data is sent over the internet or to external APIs. Your raw data and generated reports remain strictly confidential.
