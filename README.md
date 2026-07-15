# ROST Productivity Shiny App

This bundle contains the files needed to run the uploadable, multi-year
ROST productivity dashboard. It does not contain project source data.

## Requirements

- Python 3.11 or newer
- Internet access during the initial package installation

## Setup

### macOS and Linux

Open a terminal in this folder and run:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

### Windows PowerShell

Open PowerShell in this folder and run:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

If PowerShell blocks the activation script, allow it for the current session
and try again:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
```

### Windows Command Prompt

Open Command Prompt in this folder and run:

```bat
py -m venv .venv
.venv\Scripts\activate.bat
python -m pip install -r requirements.txt
```

## Launch

From this folder, run:

```bash
shiny run --reload --launch-browser --port 8080 shiny/app.py
```

The same launch command works in Windows PowerShell and Command Prompt after
the virtual environment has been activated:

```powershell
shiny run --reload --launch-browser --port 8080 shiny/app.py
```

If the browser does not open automatically, visit:

```text
http://127.0.0.1:8080
```

Stop the app with `Ctrl+C` in the terminal.

## Using the app

1. Choose **Analyze one species in one year**, **Compare one species across
   years**, or **Compare species within one year**.
2. For single-year analysis (including the current ROST 2025 dataset), upload
   one workbook containing exactly one year and one species.
3. For cross-year analysis, upload at least two single-year productivity `.xlsx`
   workbooks with distinct years and the same species.
4. For cross-species analysis, upload one single-year productivity workbook
   containing at least two species.
5. Optionally upload a resight `.xlsx` workbook.
6. Select the shared worksheet and review the column mappings.
7. Map a species column, or enter one species label for all uploaded rows.
8. Confirm the detected year shown for each workbook, then select
   **Generate analysis**.
9. Use the year and species selectors to explore a single season.
10. Use **Compare** when a comparison mode was selected during upload.
11. Use **Downloads** to export a selected PNG/CSV or a ZIP of all results.

Uploads are temporary and remain isolated to the browser session. The app does
not save uploaded workbooks or generated analysis tables.

## Included files

```text
manager_shiny_bundle/
├── README.md
├── requirements.txt
├── scripts/
│   └── clean_productivity.py
└── shiny/
    ├── app.py
    ├── data.py
    └── www/
        └── styles.css
```
