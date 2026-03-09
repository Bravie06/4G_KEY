# Excel Automation Tool

A simple and professional automation tool to map 4G KPI Raw Data to a Reporting Template format. Designed to work fully locally for strict confidentiality.

## Prerequisites

1.  **Python**: Make sure you have Python installed on your machine (Python 3.8+ recommended).
2.  **VSCode**: Open this directory in VSCode to execute it locally.

## Setup Instructions

1.  Open your terminal in this directory.
2.  Install the required dependencies using the following command:
    ```bash
    pip install -r requirements.txt
    ```

## Usage

You can easily run the application using the batch script or directly via Python.

### Option 1: Using the `.bat` file (Windows)
Double click `run.bat` from your file explorer, or run it from the terminal:
```bash
run.bat
```

### Option 2: Using Python directly
Run the application using Python:
```bash
python app.py
```

### How to process data
1.  **Raw Data File**: Browse and select the Raw Data Excel file (e.g., `Performance Management-History Query-4G_KPI...`).
2.  **Template File**: Browse and select the Template Excel file (e.g., `Event_Performance Management-History Query...`).
3.  **Output File**: Browse and choose where to save the generated Report Excel file.
4.  Click the **Run Processing** button.
5.  A success message will appear once the generated file is saved.

## Notes
*   All processing is done **locally** on your machine.
*   The script uses strict mappings for the specified KPIs and applies required conditional formatting exactly as provided.
